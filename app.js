const express = require('express');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');

const app = express();
const PORT = process.env.PORT || 3000;
const dataDir = path.join(__dirname, 'data');
const uploadsDir = path.join(__dirname, 'uploads');

fs.mkdirSync(dataDir, { recursive: true });
fs.mkdirSync(uploadsDir, { recursive: true });

const db = new Database(path.join(dataDir, 'financials.db'));
db.pragma('journal_mode = WAL');

function initDb() {
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
  snapshot_date TEXT,
  bank TEXT,
  account_type TEXT,
  account_num TEXT,
  routing TEXT,
  balance REAL,
  source_sheet TEXT,
  FOREIGN KEY(import_id) REFERENCES imports(id)
);

CREATE TABLE IF NOT EXISTS transactions (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  import_id INTEGER NOT NULL,
  txn_date TEXT,
  property_name TEXT,
  amount REAL,
  reason TEXT,
  income_amount REAL,
  expense_amount REAL,
  year_tag INTEGER,
  source_sheet TEXT,
  source_category TEXT,
  FOREIGN KEY(import_id) REFERENCES imports(id)
);

CREATE TABLE IF NOT EXISTS property_values (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  import_id INTEGER NOT NULL,
  address TEXT,
  purchased REAL,
  estimate REAL,
  snapshot_date TEXT,
  profit_loss REAL,
  alt_estimate REAL,
  alt_snapshot_date TEXT,
  FOREIGN KEY(import_id) REFERENCES imports(id)
);

CREATE TABLE IF NOT EXISTS import_issues (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  import_id INTEGER NOT NULL,
  sheet_name TEXT,
  issue_type TEXT,
  row_number INTEGER,
  raw_payload TEXT,
  message TEXT,
  FOREIGN KEY(import_id) REFERENCES imports(id)
);
`);
}
initDb();

const storage = multer.diskStorage({
  destination: (_, __, cb) => cb(null, uploadsDir),
  filename: (_, file, cb) => cb(null, `${Date.now()}-${file.originalname}`)
});
const upload = multer({ storage });

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

function excelDateToIso(value) {
  if (value == null || value === '') return null;
  if (value instanceof Date && !isNaN(value)) return value.toISOString().slice(0, 10);

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    const d = new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d));
    return d.toISOString().slice(0, 10);
  }

  const text = String(value).trim();
  const d = new Date(text);
  if (!isNaN(d)) return d.toISOString().slice(0, 10);
  return text;
}

function n(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number' && !Number.isNaN(value)) return value;

  let text = String(value).trim();
  if (!text) return null;

  const neg = text.includes('(') && text.includes(')');
  text = text
    .replace(/\$/g, '')
    .replace(/,/g, '')
    .replace(/\s+/g, '')
    .replace(/[()]/g, '');

  const parsed = Number(text);
  if (Number.isNaN(parsed)) return null;

  return neg ? -Math.abs(parsed) : parsed;
}

function s(value) {
  if (value == null) return null;
  const text = String(value).trim();
  return text === '' ? null : text;
}

function normalizePropertyName(name) {
  if (!name) return null;

  const raw = String(name).trim();
  const map = {
    'The boys': 'The Boys',
    'The boys ': 'The Boys',
    'The Boys ': 'The Boys',
    'Riverview boys': 'The Boys',
    'Riverview Boys': 'The Boys',
    'Riverview': 'Carlton Fields Dr',
    'Riverview ': 'Carlton Fields Dr',
    'Riverview rental': 'Carlton Fields Dr',
    'Fishhawk': 'Bridgecrossing Dr',
    'Fishhawk ': 'Bridgecrossing Dr',
    'Bridgecrossing': 'Bridgecrossing Dr',
    'Bridge Crossing': 'Bridgecrossing Dr',
    'Bloomingdale': 'Brookville Dr',
    'Blue Plume': 'Blue Plume Ct',
    'Blue Plume ': 'Blue Plume Ct',
    'SouthCreek': 'Blue Plume Ct',
    'SouthCreek ': 'Blue Plume Ct',
    'Carlton Fields': 'Carlton Fields Dr'
  };

  return map[raw] || raw;
}

function detectYearTag(sheetName) {
  const match = String(sheetName).match(/(20)?(\d{2})$/);
  if (!match) return null;

  const yy = Number(match[2]);
  return yy >= 70 ? 1900 + yy : 2000 + yy;
}

function addIssue(importId, sheet, type, row, payload, message) {
  db.prepare(`
    INSERT INTO import_issues (
      import_id, sheet_name, issue_type, row_number, raw_payload, message
    ) VALUES (?, ?, ?, ?, ?, ?)
  `).run(importId, sheet, type, row, JSON.stringify(payload ?? {}), message);
}

function getCanonicalPropertySet(workbook) {
  const set = new Set();

  const manual = [
    'Blue Plume Ct',
    'Bridgecrossing Dr',
    'Brookville Dr',
    'Carlton Fields Dr',
    'The Boys',
    'Vistazo',
    'Sedona'
  ];
  manual.forEach(name => set.add(normalizePropertyName(name)));

  const sheetName = workbook.SheetNames.find(name => /^Property Values$/i.test(name));
  if (!sheetName) return set;

  const ws = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

  for (const r of rows) {
    const raw = s(r['Address '] || r['Address']);
    const normalized = normalizePropertyName(raw);
    if (normalized) set.add(normalized);
  }

  return set;
}

function sanitizeImportedPropertyName(name, validProperties) {
  const normalized = normalizePropertyName(name);
  if (!normalized) return null;
  return validProperties.has(normalized) ? normalized : null;
}

function getLatestImportId() {
  const latest = db.prepare('SELECT id FROM imports ORDER BY id DESC LIMIT 1').get();
  return latest ? latest.id : null;
}

function importWorkbook(filePath, originalName) {
  const workbook = XLSX.readFile(filePath, { cellDates: true });
  const validProperties = getCanonicalPropertySet(workbook);

  const importInfo = db.prepare(`
    INSERT INTO imports (filename, imported_at, sheet_count, row_count, notes)
    VALUES (?, datetime('now'), ?, ?, ?)
  `).run(originalName, workbook.SheetNames.length, 0, 'Initial import');

  const importId = importInfo.lastInsertRowid;

  const insertAccount = db.prepare(`
    INSERT INTO account_snapshots (
      import_id, snapshot_date, bank, account_type, account_num, routing, balance, source_sheet
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const insertTxn = db.prepare(`
    INSERT INTO transactions (
      import_id, txn_date, property_name, amount, reason,
      income_amount, expense_amount, year_tag, source_sheet, source_category
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const insertValue = db.prepare(`
    INSERT INTO property_values (
      import_id, address, purchased, estimate, snapshot_date,
      profit_loss, alt_estimate, alt_snapshot_date
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `);

  let rowCount = 0;

  const ignoredSheets = new Set(['Cleanup Notes']);

  for (const sheetName of workbook.SheetNames) {
    if (ignoredSheets.has(sheetName)) {
      continue;
    }

    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

    if (/^Accounts$/i.test(sheetName)) {
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];

        const rawBalance = r['Balance'] ?? r['Balance '] ?? r['BALANCE'] ?? null;

        if (!r['DATE'] && !r['Bank'] && rawBalance == null) continue;
        if (!r['Bank'] || r['Bank'] === 'Bank') continue;

        if (!r['Acct type'] && !r['ACCT NUM'] && !r['ROUTING'] && !r['Bank']) continue;

        const balance = n(rawBalance);
        if (balance == null) {
          addIssue(
            importId,
            sheetName,
            'invalid_account_balance',
            i + 2,
            r,
            `Balance could not be parsed from value: ${rawBalance}`
          );
          continue;
        }

        insertAccount.run(
          importId,
          excelDateToIso(r['DATE']),
          s(r['Bank']),
          s(r['Acct type']),
          s(r['ACCT NUM']),
          s(r['ROUTING']),
          balance,
          sheetName
        );

        rowCount++;
      }
      continue;
    }

    if (/^Payments\s\d+/i.test(sheetName) || /^Taxes\s\d+/i.test(sheetName)) {
      const yearTag = detectYearTag(sheetName);
      const baseCategory = /^Taxes/i.test(sheetName) ? 'taxes' : 'payments';

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const date = r['Date'] || r['DATE'];
        const propertyLabel = r['Property '] || r['Property'] || null;
        let amount = n(r['Amount ']) ?? n(r['Amount']);
        const reason = r['Reason '] || r['Reason'] || null;

        const incomeKey = Object.keys(r).find(k => String(k).startsWith('Income'));
        const expenseKey = Object.keys(r).find(k => String(k).startsWith('Expenses'));

        const income = incomeKey ? n(r[incomeKey]) : null;
        const expense = expenseKey ? n(r[expenseKey]) : null;

        if (!date && !propertyLabel && amount == null && !reason) continue;

        if (amount == null && income != null) amount = income;
        if (amount == null && expense != null) amount = expense;

        if (amount == null) {
          addIssue(importId, sheetName, 'invalid_transaction_amount', i + 2, r, 'Amount could not be parsed');
          continue;
        }

        const safeProperty = sanitizeImportedPropertyName(propertyLabel, validProperties);

        let categoryLabel = baseCategory;
        const nonPropertySet = new Set(['Taxes', 'All', 'Expense', 'Charity', 'Vacation', 'Boca', 'Singer', 'Busch G']);

        if (propertyLabel && !safeProperty) {
          const trimmed = String(propertyLabel).trim();
          if (nonPropertySet.has(trimmed)) {
            categoryLabel = trimmed;
          } else {
            addIssue(
              importId,
              sheetName,
              'unmatched_property_name',
              i + 2,
              r,
              `Ignored non-portfolio property label: ${propertyLabel}`
            );
          }
        }

        insertTxn.run(
          importId,
          excelDateToIso(date),
          safeProperty,
          amount,
          s(reason),
          income,
          expense,
          yearTag,
          sheetName,
          categoryLabel
        );

        rowCount++;
      }
      continue;
    }

    if (/^Property Values$/i.test(sheetName)) {
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        if (!r['Address '] && !r['Purchased '] && !r['Estimate ']) continue;

        insertValue.run(
          importId,
          normalizePropertyName(s(r['Address '] || r['Address'])),
          n(r['Purchased ']),
          n(r['Estimate ']),
          excelDateToIso(r['Unnamed: 3']),
          n(r['Profit/Loss']),
          n(r['Unnamed: 5']),
          excelDateToIso(r['Unnamed: 6'])
        );

        rowCount++;
      }
      continue;
    }

    addIssue(importId, sheetName, 'unmapped_sheet', null, null, `Sheet "${sheetName}" was not imported.`);
  }

  db.prepare('UPDATE imports SET row_count = ? WHERE id = ?').run(rowCount, importId);
  return importId;
}

app.get('/', (req, res) => {
  const latestImport = db.prepare('SELECT * FROM imports ORDER BY id DESC LIMIT 1').get();

  const stats = latestImport
    ? {
        accounts: db.prepare('SELECT COUNT(*) AS c FROM account_snapshots WHERE import_id = ?').get(latestImport.id).c,
        transactions: db.prepare('SELECT COUNT(*) AS c FROM transactions WHERE import_id = ?').get(latestImport.id).c,
        properties: db.prepare('SELECT COUNT(*) AS c FROM property_values WHERE import_id = ?').get(latestImport.id).c,
        values: db.prepare('SELECT COUNT(*) AS c FROM property_values WHERE import_id = ?').get(latestImport.id).c,
        issues: db.prepare('SELECT COUNT(*) AS c FROM import_issues WHERE import_id = ?').get(latestImport.id).c
      }
    : null;

  const recentImports = db.prepare('SELECT * FROM imports ORDER BY id DESC LIMIT 10').all();
  res.render('index', { latestImport, stats, recentImports });
});

app.post('/import', upload.single('financialFile'), (req, res) => {
  if (!req.file) return res.status(400).send('No file uploaded');

  const importId = importWorkbook(req.file.path, req.file.originalname);
  res.redirect(`/imports/${importId}`);
});

app.get('/expenses/new', (req, res) => {
  const latestImportId = getLatestImportId();

  const properties = latestImportId
    ? db.prepare(`
        SELECT DISTINCT address AS property_name
        FROM property_values
        WHERE import_id = ? AND address IS NOT NULL AND TRIM(address) <> ''
        ORDER BY address
      `).all(latestImportId)
    : [];

  res.render('add-expense', {
    latestImportId,
    properties,
    success: req.query.success || '',
    error: req.query.error || ''
  });
});

app.post('/expenses', (req, res) => {
  const latestImportId = getLatestImportId();

  if (!latestImportId) {
    return res.redirect('/expenses/new?error=' + encodeURIComponent('Please import a workbook before adding expenses.'));
  }

  const txnDate = s(req.body.txn_date);
  const propertyName = s(req.body.property_name);
  const reason = s(req.body.reason);
  const amountValue = n(req.body.amount);
  const category = s(req.body.category);

  if (!txnDate || !propertyName || !reason || amountValue == null) {
    return res.redirect('/expenses/new?error=' + encodeURIComponent('Date, property, amount, and reason are required.'));
  }

  const expenseAmount = Math.abs(amountValue);

  db.prepare(`
    INSERT INTO transactions (
      import_id, txn_date, property_name, amount, reason,
      income_amount, expense_amount, year_tag, source_sheet, source_category
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    latestImportId,
    txnDate,
    normalizePropertyName(propertyName),
    -expenseAmount,
    reason,
    null,
    expenseAmount,
    null,
    'Manual Entry',
    category || 'manual-expense'
  );

  res.redirect('/expenses/new?success=' + encodeURIComponent('Expense saved.'));
});

app.get('/sales/new', (req, res) => {
  const latestImportId = getLatestImportId();

  const properties = latestImportId
    ? db.prepare(`
        SELECT DISTINCT address AS property_name
        FROM property_values
        WHERE import_id = ? AND address IS NOT NULL AND TRIM(address) <> ''
        ORDER BY address
      `).all(latestImportId)
    : [];

  res.render('add-sale', {
    properties,
    success: req.query.success || '',
    error: req.query.error || ''
  });
});

app.post('/sales', (req, res) => {
  const latestImportId = getLatestImportId();

  if (!latestImportId) {
    return res.redirect('/sales/new?error=' + encodeURIComponent('Please import a workbook before recording a sale.'));
  }

  const txnDate = s(req.body.txn_date);
  const propertyName = s(req.body.property_name);
  const salePrice = n(req.body.sale_price);
  const reason = s(req.body.reason) || 'Property sale';

  if (!txnDate || !propertyName || salePrice == null) {
    return res.redirect('/sales/new?error=' + encodeURIComponent('Sale date, property, and sale price are required.'));
  }

  db.prepare(`
    INSERT INTO transactions (
      import_id, txn_date, property_name, amount, reason,
      income_amount, expense_amount, year_tag, source_sheet, source_category
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    latestImportId,
    txnDate,
    normalizePropertyName(propertyName),
    salePrice,
    reason,
    salePrice,
    null,
    null,
    'Manual Sale',
    'manual-sale'
  );

  res.redirect('/sales/new?success=' + encodeURIComponent('Sale saved.'));
});

app.get('/imports/:id', (req, res) => {
  const importId = Number(req.params.id);
  const imp = db.prepare('SELECT * FROM imports WHERE id = ?').get(importId);

  if (!imp) return res.status(404).send('Import not found');

  const accounts = db.prepare(`
    SELECT * FROM account_snapshots
    WHERE import_id = ?
    ORDER BY snapshot_date DESC, bank
  `).all(importId);

  const properties = db.prepare(`
    SELECT
      property_name,
      COUNT(*) AS txn_count,
      ROUND(SUM(COALESCE(income_amount, CASE WHEN amount > 0 THEN amount ELSE 0 END)), 2) AS income_total,
      ROUND(ABS(SUM(COALESCE(expense_amount, CASE WHEN amount < 0 THEN amount ELSE 0 END))), 2) AS expense_total,
      ROUND(SUM(amount), 2) AS net_total
    FROM transactions
    WHERE import_id = ? AND property_name IS NOT NULL
    GROUP BY property_name
    ORDER BY net_total DESC
  `).all(importId);

  const propertyValues = db.prepare(`
    SELECT * FROM property_values
    WHERE import_id = ?
    ORDER BY address
  `).all(importId);

  const issues = db.prepare(`
    SELECT * FROM import_issues
    WHERE import_id = ?
    ORDER BY sheet_name, row_number
    LIMIT 100
  `).all(importId);

  res.render('import-detail', { imp, accounts, properties, propertyValues, issues });
});

app.listen(PORT, () => {
  console.log(`Financial importer running on http://localhost:${PORT}`);
});