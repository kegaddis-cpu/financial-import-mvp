const express = require('express');
const path = require('path');
const fs = require('fs');
const Database = require('better-sqlite3');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

const dataDir = path.join(__dirname, 'data');
const uploadsDir = path.join(__dirname, 'uploads');
const publicDir = path.join(__dirname, 'public');
const viewsDir = path.join(__dirname, 'views');

fs.mkdirSync(dataDir, { recursive: true });
fs.mkdirSync(uploadsDir, { recursive: true });
fs.mkdirSync(publicDir, { recursive: true });
fs.mkdirSync(viewsDir, { recursive: true });

const db = new Database(path.join(dataDir, 'financials.db'));
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

app.set('views', viewsDir);
app.set('view engine', 'ejs');

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(publicDir));

const upload = multer({ dest: uploadsDir });

db.exec(`
CREATE TABLE IF NOT EXISTS imports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  filename TEXT NOT NULL,
  imported_at TEXT NOT NULL,
  sheet_count INTEGER DEFAULT 0,
  row_count INTEGER DEFAULT 0,
  notes TEXT
);

CREATE TABLE IF NOT EXISTS account_snapshots (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  import_id INTEGER NOT NULL,
  account_name TEXT,
  balance REAL,
  as_of TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY(import_id) REFERENCES imports(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS transactions (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  import_id INTEGER NOT NULL,
  property_name TEXT,
  txn_date TEXT,
  description TEXT,
  category TEXT,
  amount REAL,
  direction TEXT,
  source_row INTEGER,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY(import_id) REFERENCES imports(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS property_values (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  import_id INTEGER NOT NULL,
  property_name TEXT,
  property_value REAL,
  as_of TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY(import_id) REFERENCES imports(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS import_issues (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  import_id INTEGER NOT NULL,
  issue_type TEXT,
  message TEXT,
  row_number INTEGER,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY(import_id) REFERENCES imports(id) ON DELETE CASCADE
);
`);

const transactionColumns = db.prepare('PRAGMA table_info(transactions)').all().map(col => col.name);

if (!transactionColumns.includes('reason')) {
  db.exec('ALTER TABLE transactions ADD COLUMN reason TEXT');
}

if (!transactionColumns.includes('source_category')) {
  db.exec('ALTER TABLE transactions ADD COLUMN source_category TEXT');
}

if (!transactionColumns.includes('income_amount')) {
  db.exec('ALTER TABLE transactions ADD COLUMN income_amount REAL');
}

if (!transactionColumns.includes('expense_amount')) {
  db.exec('ALTER TABLE transactions ADD COLUMN expense_amount REAL');
}

if (!transactionColumns.includes('year_tag')) {
  db.exec('ALTER TABLE transactions ADD COLUMN year_tag INTEGER');
}

if (!transactionColumns.includes('source_sheet')) {
  db.exec('ALTER TABLE transactions ADD COLUMN source_sheet TEXT');
}

function toNumber(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') return value;

  const raw = String(value).trim();
  if (!raw) return null;

  const negativeByParens = raw.startsWith('(') && raw.endsWith(')');
  const cleaned = raw.replace(/[$,()]/g, '').trim();
  if (!cleaned) return null;

  const num = Number(cleaned);
  if (Number.isNaN(num)) return null;
  return negativeByParens ? -Math.abs(num) : num;
}

function normalizeDate(value) {
  if (!value) return null;

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
  }

  const str = String(value).trim();
  return str || null;
}

function pick(obj, keys) {
  for (const key of keys) {
    if (obj[key] !== undefined && obj[key] !== null && obj[key] !== '') {
      return obj[key];
    }
  }
  return null;
}

function getStats() {
  const importsCount = db.prepare('SELECT COUNT(*) AS count FROM imports').get().count;
  const transactionsCount = db.prepare('SELECT COUNT(*) AS count FROM transactions').get().count;
  const accountCount = db.prepare('SELECT COUNT(*) AS count FROM account_snapshots').get().count;
  const propertyValueCount = db.prepare('SELECT COUNT(*) AS count FROM property_values').get().count;
  const issueCount = db.prepare('SELECT COUNT(*) AS count FROM import_issues').get().count;

  return {
    importCount: importsCount,
    transactionCount: transactionsCount,
    accountCount,
    propertyValueCount,
    issueCount,
    accounts: accountCount,
    transactions: transactionsCount,
    properties: propertyValueCount,
    issues: issueCount
  };
}

function getLatestImport() {
  return db.prepare('SELECT * FROM imports ORDER BY id DESC LIMIT 1').get();
}

function getRecentImports(limit = 10) {
  return db.prepare(`
    SELECT *
    FROM imports
    ORDER BY id DESC
    LIMIT ?
  `).all(limit);
}

function getLatestImportStats() {
  const latest = getLatestImport();

  if (!latest) {
    return {
      accounts: 0,
      transactions: 0,
      properties: 0,
      propertyValues: 0,
      issues: 0
    };
  }

  return {
    accounts: db.prepare('SELECT COUNT(*) AS count FROM account_snapshots WHERE import_id = ?').get(latest.id).count,
    transactions: db.prepare('SELECT COUNT(*) AS count FROM transactions WHERE import_id = ?').get(latest.id).count,
    properties: db.prepare(`
      SELECT COUNT(DISTINCT property_name) AS count
      FROM transactions
      WHERE import_id = ?
        AND property_name IS NOT NULL
        AND TRIM(property_name) <> ''
    `).get(latest.id).count,
    propertyValues: db.prepare('SELECT COUNT(*) AS count FROM property_values WHERE import_id = ?').get(latest.id).count,
    issues: db.prepare('SELECT COUNT(*) AS count FROM import_issues WHERE import_id = ?').get(latest.id).count
  };
}

function getAvailableMonths() {
  return db.prepare(`
    SELECT DISTINCT substr(txn_date, 1, 7) AS month
    FROM transactions
    WHERE txn_date IS NOT NULL
      AND length(txn_date) >= 7
    ORDER BY month DESC
  `).all().map(row => row.month);
}

function getSelectedMonth(requestedMonth) {
  const months = getAvailableMonths();
  if (!months.length) return null;
  if (requestedMonth && months.includes(requestedMonth)) return requestedMonth;
  return months[0];
}

function getMonthlyCashFlow(selectedMonth) {
  if (!selectedMonth) return [];

  return db.prepare(`
    SELECT
      property_name,
      ROUND(SUM(CASE WHEN amount > 0 THEN amount ELSE 0 END), 2) AS income_total,
      ROUND(ABS(SUM(CASE WHEN amount < 0 THEN amount ELSE 0 END)), 2) AS expense_total,
      ROUND(SUM(amount), 2) AS net_total
    FROM transactions
    WHERE txn_date IS NOT NULL
      AND substr(txn_date, 1, 7) = ?
      AND property_name IS NOT NULL
      AND TRIM(property_name) <> ''
    GROUP BY property_name
    ORDER BY property_name
  `).all(selectedMonth);
}

function getRecentTransactions(limit = 20) {
  return db.prepare(`
    SELECT
      txn_date,
      property_name,
      COALESCE(reason, description, '') AS reason,
      COALESCE(source_category, category, '') AS source_category,
      amount
    FROM transactions
    WHERE txn_date IS NOT NULL
    ORDER BY txn_date DESC, id DESC
    LIMIT ?
  `).all(limit);
}

function getLatestPropertyValues(limit = 20) {
  return db.prepare(`
    SELECT
      property_name AS address,
      property_value AS estimate,
      NULL AS profit_loss,
      as_of AS snapshot_date
    FROM property_values
    WHERE property_name IS NOT NULL
      AND TRIM(property_name) <> ''
    ORDER BY as_of DESC, id DESC
    LIMIT ?
  `).all(limit);
}

function getMissingRecurringExpenses(selectedMonth) {
  if (!selectedMonth) return [];

  return db.prepare(`
    WITH months AS (
      SELECT
        ? AS selected_month,
        strftime('%Y-%m', date(? || '-01', '-1 month')) AS prev1,
        strftime('%Y-%m', date(? || '-01', '-2 months')) AS prev2,
        strftime('%Y-%m', date(? || '-01', '-3 months')) AS prev3
    ),
    prior_dedup AS (
      SELECT DISTINCT
        substr(t.txn_date, 1, 7) AS txn_month,
        TRIM(COALESCE(t.property_name, '')) AS property_name,
        LOWER(TRIM(COALESCE(t.reason, t.description, ''))) AS normalized_reason,
        TRIM(COALESCE(t.reason, t.description, '')) AS display_reason
      FROM transactions t
      CROSS JOIN months m
      WHERE t.amount < 0
        AND t.txn_date IS NOT NULL
        AND substr(t.txn_date, 1, 7) IN (m.prev1, m.prev2, m.prev3)
        AND t.property_name IS NOT NULL
        AND TRIM(t.property_name) <> ''
        AND COALESCE(t.reason, t.description, '') <> ''
        AND LOWER(TRIM(t.property_name)) NOT IN ('charity', 'car', 'insurance', 'ethan', 'chloë', 'chloe')
    ),
    recurring_candidates AS (
      SELECT
        property_name,
        normalized_reason,
        MIN(display_reason) AS reason,
        COUNT(DISTINCT txn_month) AS months_present
      FROM prior_dedup
      GROUP BY property_name, normalized_reason
      HAVING COUNT(DISTINCT txn_month) >= 2
    ),
    current_month_dedup AS (
      SELECT DISTINCT
        TRIM(COALESCE(property_name, '')) AS property_name,
        LOWER(TRIM(COALESCE(reason, description, ''))) AS normalized_reason
      FROM transactions
      WHERE amount < 0
        AND txn_date IS NOT NULL
        AND substr(txn_date, 1, 7) = ?
        AND property_name IS NOT NULL
        AND TRIM(property_name) <> ''
        AND COALESCE(reason, description, '') <> ''
    ),
    last_seen AS (
      SELECT
        TRIM(COALESCE(property_name, '')) AS property_name,
        LOWER(TRIM(COALESCE(reason, description, ''))) AS normalized_reason,
        MAX(txn_date) AS last_seen_date,
        ROUND(AVG(ABS(amount)), 2) AS typical_amount
      FROM transactions
      WHERE amount < 0
        AND txn_date IS NOT NULL
        AND property_name IS NOT NULL
        AND TRIM(property_name) <> ''
        AND COALESCE(reason, description, '') <> ''
      GROUP BY TRIM(COALESCE(property_name, '')), LOWER(TRIM(COALESCE(reason, description, '')))
    )
    SELECT
      rc.property_name,
      rc.reason,
      rc.months_present,
      ls.last_seen_date,
      ls.typical_amount
    FROM recurring_candidates rc
    LEFT JOIN current_month_dedup cm
      ON cm.property_name = rc.property_name
     AND cm.normalized_reason = rc.normalized_reason
    LEFT JOIN last_seen ls
      ON ls.property_name = rc.property_name
     AND ls.normalized_reason = rc.normalized_reason
    WHERE cm.property_name IS NULL
    ORDER BY rc.property_name, rc.reason
  `).all(selectedMonth, selectedMonth, selectedMonth, selectedMonth, selectedMonth);
}

app.use((req, res, next) => {
  res.locals.latestImport = null;
  res.locals.recentImports = [];
  res.locals.stats = {
    importCount: 0,
    transactionCount: 0,
    accountCount: 0,
    propertyValueCount: 0,
    issueCount: 0,
    accounts: 0,
    transactions: 0,
    properties: 0,
    issues: 0
  };
  res.locals.latestImportStats = {
    accounts: 0,
    transactions: 0,
    properties: 0,
    propertyValues: 0,
    issues: 0
  };
  res.locals.monthlyCashFlow = [];
  res.locals.accounts = [];
  res.locals.properties = [];
  res.locals.propertyValues = [];
  res.locals.issues = [];
  res.locals.accountSnapshots = [];
  res.locals.propertyRollup = [];
  res.locals.importIssues = [];
  res.locals.latestIssues = [];
  res.locals.latestProperties = [];
  res.locals.portfolioProperties = [];
  res.locals.recentTransactions = [];
  res.locals.transactions = [];
  res.locals.propertiesWithTransactions = [];
  res.locals.propertyCards = [];
  res.locals.dashboardCards = [];
  res.locals.importSummary = {};
  res.locals.summary = {};
  res.locals.alerts = [];
  res.locals.warnings = [];
  res.locals.errors = [];
  res.locals.sales = [];
  res.locals.expenses = [];
  res.locals.latestPropertyValues = [];
  res.locals.availableMonths = [];
  res.locals.selectedMonth = null;
  res.locals.missingRecurringExpenses = [];
  next();
});

app.get('/', (req, res) => {
  try {
    const selectedMonth = getSelectedMonth(req.query.month);
    res.render('index', {
      latestImport: getLatestImport(),
      recentImports: getRecentImports(10),
      stats: getStats(),
      latestImportStats: getLatestImportStats(),
      monthlyCashFlow: getMonthlyCashFlow(selectedMonth),
      recentTransactions: getRecentTransactions(20),
      latestPropertyValues: getLatestPropertyValues(20),
      selectedMonth
    });
  } catch (err) {
    console.error('GET / failed:', err);
    res.status(500).send(`Internal Server Error: ${err.message}`);
  }
});

app.get('/monthly-review', (req, res) => {
  try {
    const availableMonths = getAvailableMonths();
    const selectedMonth = getSelectedMonth(req.query.month);

    res.render('monthly-review', {
      latestImport: getLatestImport(),
      stats: getStats(),
      availableMonths,
      selectedMonth,
      monthlyCashFlow: getMonthlyCashFlow(selectedMonth),
      missingRecurringExpenses: getMissingRecurringExpenses(selectedMonth)
    });
  } catch (err) {
    console.error('GET /monthly-review failed:', err);
    res.status(500).send(`Internal Server Error: ${err.message}`);
  }
});

app.get('/setup', (req, res) => {
  try {
    res.render('setup', {
      latestImport: getLatestImport(),
      recentImports: getRecentImports(10),
      stats: getStats()
    });
  } catch (err) {
    console.error('GET /setup failed:', err);
    res.status(500).send(`Internal Server Error: ${err.message}`);
  }
});

app.post('/setup/import', upload.single('workbook'), (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  try {
    const workbook = XLSX.readFile(req.file.path);
    const importedAt = new Date().toISOString();

    const insertImport = db.prepare(`
      INSERT INTO imports (filename, imported_at, sheet_count, row_count, notes)
      VALUES (?, ?, ?, ?, ?)
    `);

    const currentTxnColumns = db.prepare('PRAGMA table_info(transactions)').all().map(col => col.name);
    const hasLegacyTxnSchema = currentTxnColumns.includes('reason') && currentTxnColumns.includes('source_category');
    const hasNewTxnSchema = currentTxnColumns.includes('description') && currentTxnColumns.includes('category');

    const insertAccount = db.prepare(`
      INSERT INTO account_snapshots (import_id, account_name, balance, as_of)
      VALUES (?, ?, ?, ?)
    `);

    const insertPropertyValue = db.prepare(`
      INSERT INTO property_values (import_id, property_name, property_value, as_of)
      VALUES (?, ?, ?, ?)
    `);

    const insertIssue = db.prepare(`
      INSERT INTO import_issues (import_id, issue_type, message, row_number)
      VALUES (?, ?, ?, ?)
    `);

    const insertTxn = hasLegacyTxnSchema
      ? db.prepare(`
          INSERT INTO transactions (import_id, txn_date, property_name, amount, reason, income_amount, expense_amount, year_tag, source_sheet, source_category)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `)
      : db.prepare(`
          INSERT INTO transactions (import_id, property_name, txn_date, description, category, amount, direction, source_row)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        `);

    const runImport = db.transaction(() => {
      const importResult = insertImport.run(
        req.file.originalname,
        importedAt,
        workbook.SheetNames.length,
        0,
        null
      );

      const importId = Number(importResult.lastInsertRowid);
      let rowCount = 0;

      workbook.SheetNames.forEach((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

        rows.forEach((row, index) => {
          rowCount += 1;

          const lowerRow = {};
          Object.keys(row).forEach((key) => {
            lowerRow[String(key).trim().toLowerCase()] = row[key];
          });

          const accountName = pick(lowerRow, ['account', 'account name', 'name']);
          const balance = toNumber(pick(lowerRow, ['balance', 'ending balance', 'current balance']));
          const asOf = normalizeDate(pick(lowerRow, ['as of', 'date', 'snapshot date']));

          const propertyName = pick(lowerRow, ['property', 'property name']);
          const txnDate = normalizeDate(pick(lowerRow, ['transaction date', 'date', 'txn date']));
          const description = pick(lowerRow, ['description', 'memo', 'notes', 'reason']);
          const category = pick(lowerRow, ['category', 'type', 'source category']);
          const amount = toNumber(pick(lowerRow, ['amount', 'net', 'transaction amount']));
          const direction = pick(lowerRow, ['direction', 'income/expense', 'flow']);
          const propertyValue = toNumber(pick(lowerRow, ['property value', 'value', 'valuation']));

          if (accountName && balance !== null) {
            insertAccount.run(importId, accountName, balance, asOf);
          }

          if (propertyName && amount !== null) {
            if (hasLegacyTxnSchema) {
              const incomeAmount = amount > 0 ? amount : 0;
              const expenseAmount = amount < 0 ? amount : 0;
              const yearTag = txnDate && txnDate.length >= 4 ? Number(txnDate.slice(0, 4)) : null;

              insertTxn.run(
                importId,
                txnDate,
                propertyName,
                amount,
                description,
                incomeAmount,
                expenseAmount,
                yearTag,
                sheetName,
                category || 'payments'
              );
            } else if (hasNewTxnSchema) {
              insertTxn.run(importId, propertyName, txnDate, description, category, amount, direction, index + 2);
            }
          }

          if (propertyName && propertyValue !== null) {
            insertPropertyValue.run(importId, propertyName, propertyValue, asOf);
          }

          if (!accountName && !propertyName && amount === null && propertyValue === null) {
            insertIssue.run(
              importId,
              'unmapped_row',
              `Could not classify row from sheet "${sheetName}"`,
              index + 2
            );
          }
        });
      });

      db.prepare('UPDATE imports SET row_count = ? WHERE id = ?').run(rowCount, importId);
      return importId;
    });

    const importId = runImport();

    try {
      fs.unlinkSync(req.file.path);
    } catch (_) {}

    return res.redirect(`/imports/${importId}`);
  } catch (err) {
    console.error('POST /setup/import failed:', err);

    try {
      fs.unlinkSync(req.file.path);
    } catch (_) {}

    return res.status(500).send(`Import failed: ${err.message}`);
  }
});

app.get('/data/imports', (req, res) => {
  try {
    res.render('imports', {
      latestImport: getLatestImport(),
      recentImports: getRecentImports(25),
      stats: getStats()
    });
  } catch (err) {
    console.error('GET /data/imports failed:', err);
    res.status(500).send(`Internal Server Error: ${err.message}`);
  }
});

app.get('/imports/:id', (req, res) => {
  try {
    const importId = Number(req.params.id);

    if (!Number.isInteger(importId) || importId <= 0) {
      return res.status(400).send('Invalid import ID.');
    }

    const imp = db.prepare('SELECT * FROM imports WHERE id = ?').get(importId);

    if (!imp) {
      return res.status(404).send('Import not found.');
    }

    const accounts = db.prepare(`
      SELECT *
      FROM account_snapshots
      WHERE import_id = ?
      ORDER BY account_name, as_of
    `).all(importId);

    const currentTxnColumns = db.prepare('PRAGMA table_info(transactions)').all().map(col => col.name);
    const hasLegacyTxnSchema = currentTxnColumns.includes('reason') && currentTxnColumns.includes('source_category');

    const properties = hasLegacyTxnSchema
      ? db.prepare(`
          SELECT
            property_name,
            COUNT(*) AS txn_count,
            SUM(CASE WHEN amount > 0 THEN amount ELSE 0 END) AS income_total,
            SUM(CASE WHEN amount < 0 THEN ABS(amount) ELSE 0 END) AS expense_total,
            SUM(amount) AS net_total
          FROM transactions
          WHERE import_id = ?
          GROUP BY property_name
          ORDER BY property_name
        `).all(importId)
      : db.prepare(`
          SELECT
            property_name,
            COUNT(*) AS txn_count,
            SUM(CASE WHEN amount > 0 THEN amount ELSE 0 END) AS income_total,
            SUM(CASE WHEN amount < 0 THEN ABS(amount) ELSE 0 END) AS expense_total,
            SUM(amount) AS net_total
          FROM transactions
          WHERE import_id = ?
          GROUP BY property_name
          ORDER BY property_name
        `).all(importId);

    const propertyValues = db.prepare(`
      SELECT *
      FROM property_values
      WHERE import_id = ?
      ORDER BY property_name, as_of
    `).all(importId);

    const issues = db.prepare(`
      SELECT *
      FROM import_issues
      WHERE import_id = ?
      ORDER BY row_number, id
    `).all(importId);

    res.render('import-detail', {
      imp,
      accounts,
      properties,
      propertyValues,
      issues
    });
  } catch (err) {
    console.error('GET /imports/:id failed:', err);
    res.status(500).send(`Internal Server Error: ${err.message}`);
  }
});

app.post('/imports/:id/rollback', (req, res) => {
  try {
    const importId = Number(req.params.id);

    if (!Number.isInteger(importId) || importId <= 0) {
      return res.status(400).send('Invalid import ID.');
    }

    const imp = db.prepare('SELECT id FROM imports WHERE id = ?').get(importId);

    if (!imp) {
      return res.status(404).send('Import not found.');
    }

    const rollbackImport = db.transaction((id) => {
      db.prepare('DELETE FROM import_issues WHERE import_id = ?').run(id);
      db.prepare('DELETE FROM transactions WHERE import_id = ?').run(id);
      db.prepare('DELETE FROM property_values WHERE import_id = ?').run(id);
      db.prepare('DELETE FROM account_snapshots WHERE import_id = ?').run(id);
      db.prepare('DELETE FROM imports WHERE id = ?').run(id);
    });

    rollbackImport(importId);
    return res.redirect('/data/imports');
  } catch (err) {
    console.error('POST /imports/:id/rollback failed:', err);
    return res.status(500).send(`Rollback failed: ${err.message}`);
  }
});

app.get('/expenses/new', (req, res) => {
  res.send('Add expense page placeholder');
});

app.get('/sales/new', (req, res) => {
  res.send('Record sale page placeholder');
});

app.use((req, res) => {
  res.status(404).send(`Cannot ${req.method} ${req.path}`);
});

app.listen(PORT, () => {
  console.log(`Financial importer running on http://localhost:${PORT}/`);
});