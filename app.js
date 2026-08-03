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
    return {
        importCount: db.prepare('SELECT COUNT(*) AS count FROM imports').get().count,
        transactionCount: db.prepare('SELECT COUNT(*) AS count FROM transactions').get().count,
        accountCount: db.prepare('SELECT COUNT(*) AS count FROM account_snapshots').get().count,
        propertyValueCount: db.prepare('SELECT COUNT(*) AS count FROM property_values').get().count,
        issueCount: db.prepare('SELECT COUNT(*) AS count FROM import_issues').get().count
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

function getMonthlyCashFlow() {
    try {
        return db.prepare(`
      SELECT
        COALESCE(property_name, 'Unassigned') AS property_name,
        SUM(amount) AS net_amount
      FROM transactions
      GROUP BY COALESCE(property_name, 'Unassigned')
      ORDER BY property_name
    `).all();
    } catch (err) {
        console.error('getMonthlyCashFlow failed:', err);
        return [];
    }
}

app.use((req, res, next) => {
    res.locals.latestImport = null;
    res.locals.recentImports = [];
    res.locals.stats = {
        importCount: 0,
        transactionCount: 0,
        accountCount: 0,
        propertyValueCount: 0,
        issueCount: 0
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
    next();
});

app.get('/', (req, res) => {
    try {
        res.render('index', {
            latestImport: getLatestImport(),
            recentImports: getRecentImports(10),
            stats: getStats(),
            latestImportStats: getLatestImportStats(),
            monthlyCashFlow: getMonthlyCashFlow()
        });
    } catch (err) {
        console.error('GET / failed:', err);
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

        const insertAccount = db.prepare(`
      INSERT INTO account_snapshots (import_id, account_name, balance, as_of)
      VALUES (?, ?, ?, ?)
    `);

        const insertTxn = db.prepare(`
      INSERT INTO transactions (import_id, property_name, txn_date, description, category, amount, direction, source_row)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `);

        const insertPropertyValue = db.prepare(`
      INSERT INTO property_values (import_id, property_name, property_value, as_of)
      VALUES (?, ?, ?, ?)
    `);

        const insertIssue = db.prepare(`
      INSERT INTO import_issues (import_id, issue_type, message, row_number)
      VALUES (?, ?, ?, ?)
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
                    const description = pick(lowerRow, ['description', 'memo', 'notes']);
                    const category = pick(lowerRow, ['category', 'type']);
                    const amount = toNumber(pick(lowerRow, ['amount', 'net', 'transaction amount']));
                    const direction = pick(lowerRow, ['direction', 'income/expense', 'flow']);
                    const propertyValue = toNumber(pick(lowerRow, ['property value', 'value', 'valuation']));

                    if (accountName && balance !== null) {
                        insertAccount.run(importId, accountName, balance, asOf);
                    }

                    if (propertyName && amount !== null) {
                        insertTxn.run(importId, propertyName, txnDate, description, category, amount, direction, index + 2);
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
        } catch (_) { }

        return res.redirect(`/imports/${importId}`);
    } catch (err) {
        console.error('POST /setup/import failed:', err);

        try {
            fs.unlinkSync(req.file.path);
        } catch (_) { }

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

        const properties = db.prepare(`
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