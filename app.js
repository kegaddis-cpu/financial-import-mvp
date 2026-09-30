const express = require('express');
const path = require('path');
const fs = require('fs');
const { Pool } = require('pg');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

const uploadsDir = path.join(__dirname, 'uploads');
const publicDir = path.join(__dirname, 'public');
const viewsDir = path.join(__dirname, 'views');


console.log('DATABASE_URL present:', !!process.env.DATABASE_URL, process.env.DATABASE_URL?.slice(0, 20));

fs.mkdirSync(uploadsDir, { recursive: true });
fs.mkdirSync(publicDir, { recursive: true });
fs.mkdirSync(viewsDir, { recursive: true });

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

app.set('views', viewsDir);
app.set('view engine', 'ejs');

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(publicDir));

const upload = multer({ dest: uploadsDir });

async function query(sql, params = []) {
  return pool.query(sql, params);
}

async function initDb() {
  await query(`
    CREATE TABLE IF NOT EXISTS imports (
      id SERIAL PRIMARY KEY,
      filename TEXT NOT NULL,
      imported_at TIMESTAMP NOT NULL,
      sheet_count INTEGER DEFAULT 0,
      row_count INTEGER DEFAULT 0,
      notes TEXT
    );
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS account_snapshots (
      id SERIAL PRIMARY KEY,
      import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
      account_name TEXT,
      balance NUMERIC,
      as_of TEXT,
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS transactions (
      id SERIAL PRIMARY KEY,
      import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
      property_name TEXT,
      txn_date TEXT,
      description TEXT,
      category TEXT,
      amount NUMERIC,
      direction TEXT,
      source_row INTEGER,
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
      reason TEXT,
      source_category TEXT,
      income_amount NUMERIC,
      expense_amount NUMERIC,
      year_tag INTEGER,
      source_sheet TEXT
    );
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS property_values (
      id SERIAL PRIMARY KEY,
      import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
      property_name TEXT,
      property_value NUMERIC,
      as_of TEXT,
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS import_issues (
      id SERIAL PRIMARY KEY,
      import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
      issue_type TEXT,
      message TEXT,
      row_number INTEGER,
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
  `);
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

async function getStats() {
  const importsCount = (await query('SELECT COUNT(*)::int AS count FROM imports')).rows[0].count;
  const transactionsCount = (await query('SELECT COUNT(*)::int AS count FROM transactions')).rows[0].count;
  const accountCount = (await query('SELECT COUNT(*)::int AS count FROM account_snapshots')).rows[0].count;
  const propertyValueCount = (await query('SELECT COUNT(*)::int AS count FROM property_values')).rows[0].count;
  const issueCount = (await query('SELECT COUNT(*)::int AS count FROM import_issues')).rows[0].count;

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

async function getLatestImport() {
  const result = await query('SELECT * FROM imports ORDER BY id DESC LIMIT 1');
  return result.rows[0] || null;
}

async function getRecentImports(limit = 10) {
  const result = await query(`
    SELECT *
    FROM imports
    ORDER BY id DESC
    LIMIT $1
  `, [limit]);
  return result.rows;
}

async function getLatestImportStats() {
  const latest = await getLatestImport();

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
    accounts: (await query('SELECT COUNT(*)::int AS count FROM account_snapshots WHERE import_id = $1', [latest.id])).rows[0].count,
    transactions: (await query('SELECT COUNT(*)::int AS count FROM transactions WHERE import_id = $1', [latest.id])).rows[0].count,
    properties: (await query(`
      SELECT COUNT(DISTINCT property_name)::int AS count
      FROM transactions
      WHERE import_id = $1
        AND property_name IS NOT NULL
        AND TRIM(property_name) <> ''
    `, [latest.id])).rows[0].count,
    propertyValues: (await query('SELECT COUNT(*)::int AS count FROM property_values WHERE import_id = $1', [latest.id])).rows[0].count,
    issues: (await query('SELECT COUNT(*)::int AS count FROM import_issues WHERE import_id = $1', [latest.id])).rows[0].count
  };
}

async function getAvailableMonths() {
  const result = await query(`
    SELECT DISTINCT substr(txn_date, 1, 7) AS month
    FROM transactions
    WHERE txn_date IS NOT NULL
      AND length(txn_date) >= 7
    ORDER BY month DESC
  `);
  return result.rows.map(row => row.month);
}

async function getSelectedMonth(requestedMonth) {
  const months = await getAvailableMonths();
  if (!months.length) return null;
  if (requestedMonth && months.includes(requestedMonth)) return requestedMonth;
  return months[0];
}

async function getMonthlyCashFlow(selectedMonth) {
  if (!selectedMonth) return [];

  const result = await query(`
    SELECT
      property_name,
      ROUND(SUM(CASE WHEN amount > 0 THEN amount ELSE 0 END)::numeric, 2) AS income_total,
      ROUND(ABS(SUM(CASE WHEN amount < 0 THEN amount ELSE 0 END))::numeric, 2) AS expense_total,
      ROUND(SUM(amount)::numeric, 2) AS net_total
    FROM transactions
    WHERE txn_date IS NOT NULL
      AND substr(txn_date, 1, 7) = $1
      AND property_name IS NOT NULL
      AND TRIM(property_name) <> ''
    GROUP BY property_name
    ORDER BY property_name
  `, [selectedMonth]);

  return result.rows;
}

async function getRecentTransactions(limit = 20) {
  const result = await query(`
    SELECT
      txn_date,
      property_name,
      COALESCE(reason, description, '') AS reason,
      COALESCE(source_category, category, '') AS source_category,
      amount
    FROM transactions
    WHERE txn_date IS NOT NULL
    ORDER BY txn_date DESC, id DESC
    LIMIT $1
  `, [limit]);

  return result.rows;
}

async function getLatestPropertyValues(limit = 20) {
  const result = await query(`
    SELECT
      property_name AS address,
      property_value AS estimate,
      NULL AS profit_loss,
      as_of AS snapshot_date
    FROM property_values
    WHERE property_name IS NOT NULL
      AND TRIM(property_name) <> ''
    ORDER BY as_of DESC, id DESC
    LIMIT $1
  `, [limit]);

  return result.rows;
}

async function getMissingRecurringExpenses(selectedMonth) {
  if (!selectedMonth) return [];

  const result = await query(`
    WITH months AS (
      SELECT
        $1::text AS selected_month,
        to_char((to_date($1 || '-01', 'YYYY-MM-DD') - interval '1 month'), 'YYYY-MM') AS prev1,
        to_char((to_date($1 || '-01', 'YYYY-MM-DD') - interval '2 month'), 'YYYY-MM') AS prev2,
        to_char((to_date($1 || '-01', 'YYYY-MM-DD') - interval '3 month'), 'YYYY-MM') AS prev3
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
        AND LOWER(TRIM(t.property_name)) NOT IN ('charity', 'car', 'insurance', 'ethan', 'chlo�', 'chloe')
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
        AND substr(txn_date, 1, 7) = $1
        AND property_name IS NOT NULL
        AND TRIM(property_name) <> ''
        AND COALESCE(reason, description, '') <> ''
    ),
    last_seen AS (
      SELECT
        TRIM(COALESCE(property_name, '')) AS property_name,
        LOWER(TRIM(COALESCE(reason, description, ''))) AS normalized_reason,
        MAX(txn_date) AS last_seen_date,
        ROUND(AVG(ABS(amount))::numeric, 2) AS typical_amount
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
  `, [selectedMonth]);

  return result.rows;
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

app.get('/', async (req, res) => {
  try {
    const selectedMonth = await getSelectedMonth(req.query.month);
    res.render('index', {
      latestImport: await getLatestImport(),
      recentImports: await getRecentImports(10),
      stats: await getStats(),
      latestImportStats: await getLatestImportStats(),
      monthlyCashFlow: await getMonthlyCashFlow(selectedMonth),
      recentTransactions: await getRecentTransactions(20),
      latestPropertyValues: await getLatestPropertyValues(20),
      selectedMonth
    });
  } catch (err) {
    console.error('GET / failed:', err);
    res.status(500).send(`Internal Server Error: ${err.message}`);
  }
});

app.get('/monthly-review', async (req, res) => {
  try {
    const availableMonths = await getAvailableMonths();
    const selectedMonth = await getSelectedMonth(req.query.month);

    res.render('monthly-review', {
      latestImport: await getLatestImport(),
      stats: await getStats(),
      availableMonths,
      selectedMonth,
      monthlyCashFlow: await getMonthlyCashFlow(selectedMonth),
      missingRecurringExpenses: await getMissingRecurringExpenses(selectedMonth)
    });
  } catch (err) {
    console.error('GET /monthly-review failed:', err);
    res.status(500).send(`Internal Server Error: ${err.message}`);
  }
});

app.get('/setup', async (req, res) => {
  try {
    res.render('setup', {
      latestImport: await getLatestImport(),
      recentImports: await getRecentImports(10),
      stats: await getStats()
    });
  } catch (err) {
    console.error('GET /setup failed:', err);
    res.status(500).send(`Internal Server Error: ${err.message}`);
  }
});

app.post('/setup/import', upload.single('workbook'), async (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  const client = await pool.connect();

  try {
    const workbook = XLSX.readFile(req.file.path);
    const importedAt = new Date();

    await client.query('BEGIN');

    const importResult = await client.query(`
      INSERT INTO imports (filename, imported_at, sheet_count, row_count, notes)
      VALUES ($1, $2, $3, $4, $5)
      RETURNING id
    `, [
      req.file.originalname,
      importedAt,
      workbook.SheetNames.length,
      0,
      null
    ]);

    const importId = importResult.rows[0].id;
    let rowCount = 0;

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

      for (let index = 0; index < rows.length; index += 1) {
        const row = rows[index];
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
          await client.query(`
            INSERT INTO account_snapshots (import_id, account_name, balance, as_of)
            VALUES ($1, $2, $3, $4)
          `, [importId, accountName, balance, asOf]);
        }

        if (propertyName && amount !== null) {
          const incomeAmount = amount > 0 ? amount : 0;
          const expenseAmount = amount < 0 ? amount : 0;
          const yearTag = txnDate && txnDate.length >= 4 ? Number(txnDate.slice(0, 4)) : null;

          await client.query(`
            INSERT INTO transactions (
              import_id, property_name, txn_date, description, category, amount, direction, source_row,
              reason, source_category, income_amount, expense_amount, year_tag, source_sheet
            )
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14)
          `, [
            importId,
            propertyName,
            txnDate,
            description,
            category,
            amount,
            direction,
            index + 2,
            description,
            category || 'payments',
            incomeAmount,
            expenseAmount,
            yearTag,
            sheetName
          ]);
        }

        if (propertyName && propertyValue !== null) {
          await client.query(`
            INSERT INTO property_values (import_id, property_name, property_value, as_of)
            VALUES ($1, $2, $3, $4)
          `, [importId, propertyName, propertyValue, asOf]);
        }

        if (!accountName && !propertyName && amount === null && propertyValue === null) {
          await client.query(`
            INSERT INTO import_issues (import_id, issue_type, message, row_number)
            VALUES ($1, $2, $3, $4)
          `, [
            importId,
            'unmapped_row',
            `Could not classify row from sheet "${sheetName}"`,
            index + 2
          ]);
        }
      }
    }

    await client.query('UPDATE imports SET row_count = $1 WHERE id = $2', [rowCount, importId]);
    await client.query('COMMIT');

    try {
      fs.unlinkSync(req.file.path);
    } catch (_) {}

    return res.redirect(`/imports/${importId}`);
  } catch (err) {
    await client.query('ROLLBACK');
    console.error('POST /setup/import failed:', err);

    try {
      fs.unlinkSync(req.file.path);
    } catch (_) {}

    return res.status(500).send(`Import failed: ${err.message}`);
  } finally {
    client.release();
  }
});

app.get('/data/imports', async (req, res) => {
  try {
    res.render('imports', {
      latestImport: await getLatestImport(),
      recentImports: await getRecentImports(25),
      stats: await getStats()
    });
  } catch (err) {
    console.error('GET /data/imports failed:', err);
    res.status(500).send(`Internal Server Error: ${err.message}`);
  }
});

app.get('/imports/:id', async (req, res) => {
  try {
    const importId = Number(req.params.id);

    if (!Number.isInteger(importId) || importId <= 0) {
      return res.status(400).send('Invalid import ID.');
    }

    const impResult = await query('SELECT * FROM imports WHERE id = $1', [importId]);
    const imp = impResult.rows[0];

    if (!imp) {
      return res.status(404).send('Import not found.');
    }

    const accounts = (await query(`
      SELECT *
      FROM account_snapshots
      WHERE import_id = $1
      ORDER BY account_name, as_of
    `, [importId])).rows;

    const properties = (await query(`
      SELECT
        property_name,
        COUNT(*)::int AS txn_count,
        SUM(CASE WHEN amount > 0 THEN amount ELSE 0 END) AS income_total,
        SUM(CASE WHEN amount < 0 THEN ABS(amount) ELSE 0 END) AS expense_total,
        SUM(amount) AS net_total
      FROM transactions
      WHERE import_id = $1
      GROUP BY property_name
      ORDER BY property_name
    `, [importId])).rows;

    const propertyValues = (await query(`
      SELECT *
      FROM property_values
      WHERE import_id = $1
      ORDER BY property_name, as_of
    `, [importId])).rows;

    const issues = (await query(`
      SELECT *
      FROM import_issues
      WHERE import_id = $1
      ORDER BY row_number, id
    `, [importId])).rows;

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

app.post('/imports/:id/rollback', async (req, res) => {
  const client = await pool.connect();

  try {
    const importId = Number(req.params.id);

    if (!Number.isInteger(importId) || importId <= 0) {
      client.release();
      return res.status(400).send('Invalid import ID.');
    }

    const imp = (await client.query('SELECT id FROM imports WHERE id = $1', [importId])).rows[0];

    if (!imp) {
      client.release();
      return res.status(404).send('Import not found.');
    }

    await client.query('BEGIN');
    await client.query('DELETE FROM import_issues WHERE import_id = $1', [importId]);
    await client.query('DELETE FROM transactions WHERE import_id = $1', [importId]);
    await client.query('DELETE FROM property_values WHERE import_id = $1', [importId]);
    await client.query('DELETE FROM account_snapshots WHERE import_id = $1', [importId]);
    await client.query('DELETE FROM imports WHERE id = $1', [importId]);
    await client.query('COMMIT');

    return res.redirect('/data/imports');
  } catch (err) {
    await client.query('ROLLBACK');
    console.error('POST /imports/:id/rollback failed:', err);
    return res.status(500).send(`Rollback failed: ${err.message}`);
  } finally {
    client.release();
  }
});

app.post('/imports/rollback-last', async (req, res) => {
  const client = await pool.connect();

  try {
    const latest = (await client.query(`
      SELECT id
      FROM imports
      ORDER BY id DESC
      LIMIT 1
    `)).rows[0];

    if (!latest) {
      return res.status(404).send('No imports found to roll back.');
    }

    await client.query('BEGIN');

    await client.query('DELETE FROM import_issues WHERE import_id = $1', [latest.id]);
    await client.query('DELETE FROM transactions WHERE import_id = $1', [latest.id]);
    await client.query('DELETE FROM property_values WHERE import_id = $1', [latest.id]);
    await client.query('DELETE FROM account_snapshots WHERE import_id = $1', [latest.id]);
    await client.query('DELETE FROM imports WHERE id = $1', [latest.id]);

    await client.query('COMMIT');

    return res.redirect('/data/imports');
  } catch (err) {
    await client.query('ROLLBACK');
    console.error('POST /imports/rollback-last failed:', err);
    return res.status(500).send(`Rollback failed: ${err.message}`);
  } finally {
    client.release();
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

initDb()
  .then(() => {
    app.listen(PORT, () => {
      console.log(`Financial importer running on http://localhost:${PORT}/`);
    });
  })
  .catch((err) => {
    console.error('Database initialization failed:', err);
    process.exit(1);
  });
