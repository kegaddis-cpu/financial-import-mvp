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

const currencyFormatter = new Intl.NumberFormat('en-US', {
  style: 'currency',
  currency: 'USD',
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});

function formatCurrency(value) {
  if (value == null || value === '') return '—';
  const num = typeof value === 'number' ? value : Number(value);
  if (Number.isNaN(num)) return '—';
  return currencyFormatter.format(num);
}

app.locals.formatCurrency = formatCurrency;

function columnExists(tableName, columnName) {
  const cols = db.prepare(`PRAGMA table_info(${tableName})`).all();
  return cols.some(col => col.name === columnName);
}

function ensureColumn(tableName, columnSql) {
  const match = columnSql.match(/^\s*([a-zA-Z_][a-zA-Z0-9_]*)\s+/);
  if (!match) return;
  const columnName = match[1];
  if (!columnExists(tableName, columnName)) {
    db.exec(`ALTER TABLE ${tableName} ADD COLUMN ${columnSql}`);
  }
}

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

  ensureColumn('import_issues', `status TEXT DEFAULT 'open'`);
  ensureColumn('import_issues', `resolution_type TEXT`);
  ensureColumn('import_issues', `resolution_payload TEXT`);
  ensureColumn('import_issues', `resolved_at TEXT`);
}
initDb();

/**
 * Multer: disk storage with type/size validation
 */
const storage = multer.diskStorage({
  destination: (_, __, cb) => cb(null, uploadsDir),
  filename: (_, file, cb) => {
    const safeOriginal = String(file.originalname || 'upload.xlsx')
      .replace(/[^a-zA-Z0-9._-]/g, '_');
    cb(null, `${Date.now()}-${safeOriginal}`);
  }
});

const upload = multer({
  storage,
  limits: { fileSize: 15 * 1024 * 1024 }, // 15 MB
  fileFilter: (_, file, cb) => {
    const ext = path.extname(file.originalname || '').toLowerCase();
    const okMime = new Set([
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel',
      'application/octet-stream'
    ]);

    if (ext !== '.xlsx' && ext !== '.xls') {
      return cb(new Error('Only .xlsx or .xls workbook files are allowed'));
    }

    if (file.mimetype && !okMime.has(file.mimetype)) {
      return cb(new Error(`Unsupported file type: ${file.mimetype}`));
    }

    cb(null, true);
  }
});

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

/**
 * Helper functions
 */
function excelDateToIso(value) {
  if (value == null || value === '') return null;
  if (value instanceof Date && !isNaN(value)) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    const d = new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d));
    return d.toISOString().slice(0, 10);
  }

  const text = String(value).trim();
  const d = new Date(text);
  if (!isNaN(d)) return d.toISOString().slice(0, 10);

  if (/^\d{5}$/.test(text)) {
    const parsed = XLSX.SSF.parse_date_code(Number(text));
    if (parsed) {
      const dd = new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d));
      return dd.toISOString().slice(0, 10);
    }
  }

  return text || null;
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
    .replace(/[()]/g, '')
    .replace(/[^0-9.-]/g, '');

  if (!text || text === '-' || text === '.' || text === '-.') return null;
  const parsed = Number(text);
  if (Number.isNaN(parsed)) return null;
  return neg ? -Math.abs(parsed) : parsed;
}

function s(value) {
  if (value == null) return null;
  const text = String(value).trim();
  return text === '' ? null : text;
}

function normalizeSheetKey(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/&/g, 'and')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function findSheet(workbook, target) {
  const wanted = normalizeSheetKey(target);
  return workbook.SheetNames.find(name => normalizeSheetKey(name) === wanted);
}

function normalizePropertyName(name) {
  if (!name) return null;

  const raw = String(name).trim();
  const map = {
    'The boys': 'Blue Plume Ct',
    'The boys ': 'Blue Plume Ct',
    'The Boys': 'Blue Plume Ct',
    'The Boys ': 'Blue Plume Ct',
    'Riverview boys': 'Blue Plume Ct',
    'Riverview Boys': 'Blue Plume Ct',
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
    'Blue Plume Ct ': 'Blue Plume Ct',
    'SouthCreek': 'Blue Plume Ct',
    'SouthCreek ': 'Blue Plume Ct',
    'Carlton Fields': 'Carlton Fields Dr',
    'Carlton Fields ': 'Carlton Fields Dr',
    'Sedona': 'Sedona',
    'Vistazo': 'Vistazo',
    'Singer': 'Singer'
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
      import_id, sheet_name, issue_type, row_number, raw_payload, message, status
    ) VALUES (?, ?, ?, ?, ?, ?, 'open')
  `).run(importId, sheet, type, row, JSON.stringify(payload ?? {}), message);
}

function getCanonicalPropertySet(workbook) {
  const set = new Set();
  const manual = [
    'Blue Plume Ct',
    'Bridgecrossing Dr',
    'Brookville Dr',
    'Carlton Fields Dr',
    'Vistazo',
    'Sedona',
    'Singer'
  ];
  manual.forEach(name => set.add(normalizePropertyName(name)));

  const propertySheetName = findSheet(workbook, 'Property Values');
  if (!propertySheetName) return set;

  const ws = workbook.Sheets[propertySheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

  for (const r of rows) {
    const raw = s(r['Address '] || r['Address']);
    const normalized = normalizePropertyName(raw);
    if (normalized) set.add(normalized);
  }

  return set;
}

function getCanonicalPropertyList(importId = null) {
  const set = new Set([
    'Blue Plume Ct',
    'Bridgecrossing Dr',
    'Brookville Dr',
    'Carlton Fields Dr',
    'Vistazo',
    'Sedona',
    'Singer'
  ]);

  const txnRows = importId
    ? db.prepare(`SELECT DISTINCT property_name FROM transactions WHERE import_id = ? AND property_name IS NOT NULL`).all(importId)
    : db.prepare(`SELECT DISTINCT property_name FROM transactions WHERE property_name IS NOT NULL`).all();

  const valueRows = importId
    ? db.prepare(`SELECT DISTINCT address FROM property_values WHERE import_id = ? AND address IS NOT NULL`).all(importId)
    : db.prepare(`SELECT DISTINCT address FROM property_values WHERE address IS NOT NULL`).all();

  txnRows.forEach(r => {
    const normalized = normalizePropertyName(r.property_name);
    if (normalized) set.add(normalized);
  });

  valueRows.forEach(r => {
    const normalized = normalizePropertyName(r.address);
    if (normalized) set.add(normalized);
  });

  return Array.from(set).sort((a, b) => a.localeCompare(b));
}

function sanitizeImportedPropertyName(name, validProperties) {
  const normalized = normalizePropertyName(name);
  if (!normalized) return null;
  return validProperties.has(normalized) ? normalized : null;
}

function inferPropertyFromReason(reason) {
  const text = s(reason);
  if (!text) return null;

  const candidates = [
    'The Boys',
    'Riverview rental',
    'Riverview',
    'Carlton Fields',
    'Bridgecrossing',
    'Fishhawk',
    'Bloomingdale',
    'Blue Plume',
    'SouthCreek',
    'Sedona',
    'Vistazo',
    'Singer'
  ];

  const lower = text.toLowerCase();
  for (const candidate of candidates) {
    if (lower.includes(candidate.toLowerCase())) return normalizePropertyName(candidate);
  }
  return null;
}

function isNonPortfolioCategory(value) {
  const v = s(value)?.toLowerCase();
  if (!v) return false;

  const nonPortfolio = new Set([
    'charity', 'insurance', 'income', 'taxes', 'tax', 'all', 'expense',
    'expenses', 'vacation', 'music', 'car', 'ethan', 'chlo', 'chloë',
    'busch g', 'boca', 'fun', 'personal', 'health'
  ]);

  return nonPortfolio.has(v);
}

function isSummaryReason(reason) {
  const r = s(reason)?.toLowerCase();
  if (!r) return false;

  return [
    'monthly',
    'loan payment',
    'income',
    'personal expenses',
    'property tax',
    'warranty',
    'riverview rent',
    'without pods',
    'without riverview rent',
    'total',
    'years to pay back down payment'
  ].includes(r);
}

function isProbablySummaryRow(row) {
  const date = s(row['Date'] || row['DATE']);
  const property = s(row['Property '] || row['Property']);
  const reason = s(row['Reason '] || row['Reason']);
  const amount = n(row['Amount '] || row['Amount']);

  if (!date && !property && !reason && amount == null) return true;
  if (property && isSummaryReason(property)) return true;
  if (reason && isSummaryReason(reason)) return true;
  return false;
}

function shouldSkipPayments21Row(row) {
  const date = s(row['Date'] || row['DATE']);
  const reason = s(row['Reason'] || row['Reason ']);
  const amount = n(row['Amount'] || row['Amount ']);
  const combined = [date, reason].filter(Boolean).join(' ').toLowerCase();

  if (combined.includes('nov 30 taxes')) return true;
  if (combined.includes('revenue')) return true;
  if (combined.includes('expenses')) return true;
  if (combined.includes('total monthly')) return true;
  if (combined.includes('weekly rental profit')) return true;
  if (combined.includes('weekly salary')) return true;
  if (combined.includes('save monthly')) return true;
  if (combined.includes('monthly yearly')) return true;
  if (!date && amount == null && !reason) return true;
  return false;
}

function hasExplicitZeroLedger(row) {
  const incomeKey = Object.keys(row).find(k => String(k).trim().startsWith('Income'));
  const expenseKey = Object.keys(row).find(k => String(k).trim().startsWith('Expenses'));
  const income = incomeKey ? n(row[incomeKey]) : null;
  const expense = expenseKey ? n(row[expenseKey]) : null;
  return income === 0 || expense === 0;
}

function isFooterLikeText(row) {
  const vals = Object.values(row)
    .map(v => s(v)?.toLowerCase())
    .filter(Boolean)
    .join(' | ');

  return [
    'monthly',
    'loan payment',
    'without pods',
    'without riverview rent',
    'years to pay back down payment',
    'weekly rental profit',
    'weekly total',
    'total monthly expenses',
    'total monthly income',
    'save monthly for property taxes insurance'
  ].some(x => vals.includes(x));
}

function isGarbageOrSummaryPaymentsRow(row) {
  const vals = Object.values(row)
    .map(v => s(v))
    .filter(Boolean);

  if (!vals.length) return true;

  const joined = vals.join(' ').toLowerCase();
  if (joined.includes('errorref!')) return true;
  if (joined.includes('title payments')) return true;
  if (joined.includes('title taxes')) return true;
  if (joined.includes('revenue amount')) return true;
  if (joined.includes('expenses')) return true;
  if (joined.includes('monthly yearly')) return true;
  if (joined.includes('nov 30 taxes')) return true;
  if (joined.includes('save monthly')) return true;
  if (joined.includes('weekly salary')) return true;
  if (joined.includes('weekly car insurance')) return true;
  if (joined.includes('weekly phone')) return true;
  if (joined.includes('weekly water')) return true;
  if (joined.includes('weekly electric gas')) return true;
  if (joined.includes('weekly internet')) return true;
  if (joined.includes('weekly rental profit')) return true;
  if (joined.includes('weekly total')) return true;
  if (joined.includes('total monthly expenses')) return true;
  if (joined.includes('total monthly income')) return true;
  if (joined.includes('years to pay back down payment')) return true;
  if (joined.includes('without pods')) return true;
  if (joined.includes('without riverview rent')) return true;
  if (/^(income|expenses|total)$/i.test(vals[0] || '')) return true;

  return false;
}

function getLatestImportId() {
  const latest = db.prepare('SELECT id FROM imports ORDER BY id DESC LIMIT 1').get();
  return latest ? latest.id : null;
}

/**
 * Workbook shape validation
 */
function validateWorkbookShape(workbook) {
  const issues = [];

  if (!workbook || !Array.isArray(workbook.SheetNames) || workbook.SheetNames.length === 0) {
    issues.push('Workbook appears to be empty or unreadable');
    return issues;
  }

  const hasAccounts = !!findSheet(workbook, 'Accounts');
  const hasPropertyValues = !!findSheet(workbook, 'Property Values');
  const hasYearSheets = workbook.SheetNames.some(name => /^(Payments|Taxes)\s*\d+/i.test(name));

  if (!hasAccounts) issues.push('Missing required sheet: Accounts');
  if (!hasPropertyValues) issues.push('Missing required sheet: Property Values');
  if (!hasYearSheets) issues.push('Missing year sheets like Payments 26 or Taxes 26');

  return issues;
}

/**
 * Import workbook
 */
function importWorkbook(filePath, originalName) {
  let workbook;
  try {
    workbook = XLSX.readFile(filePath, { cellDates: true });
  } catch (err) {
    throw new Error(`Excel workbook could not be read: ${err.message}`);
  }

  const validationIssues = validateWorkbookShape(workbook);
  if (validationIssues.length) {
    throw new Error(validationIssues.join(' | '));
  }

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
      import_id, txn_date, property_name, amount, reason, income_amount, expense_amount,
      year_tag, source_sheet, source_category
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const insertValue = db.prepare(`
    INSERT INTO property_values (
      import_id, address, purchased, estimate, snapshot_date, profit_loss, alt_estimate, alt_snapshot_date
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `);

  let rowCount = 0;
  let recognizedRevenueSheet = false;

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

    if (/^Accounts$/i.test(sheetName)) {
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const rawDate = r['DATE'];
        const bank = s(r['Bank']);
        const acctType = s(r['Acct type']);
        const balance = n(r['Balance']);

        const hasUsefulAccountData = bank || acctType || balance != null || rawDate;
        if (!hasUsefulAccountData) continue;
        if (!bank || /^bank$/i.test(bank) || /^title$/i.test(bank)) continue;
        if (balance == null) continue;

        insertAccount.run(
          importId,
          excelDateToIso(rawDate),
          bank,
          acctType,
          s(r['ACCT NUM']),
          s(r['ROUTING']),
          balance,
          sheetName
        );
        rowCount++;
      }
      continue;
    }

    if (/^(Payments|Taxes)\s*\d+/i.test(sheetName)) {
      const yearTag = detectYearTag(sheetName);
      const category = /^Taxes/i.test(sheetName) ? 'taxes' : 'payments';

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const rowNum = i + 2;

        if (isProbablySummaryRow(r)) continue;
        if (/^Payments\s*21$/i.test(sheetName) && shouldSkipPayments21Row(r)) continue;
        if (isFooterLikeText(r)) cont
app.get('/api/latest-summary', (req, res) => {
  const importId = getLatestImportId();
  if (!importId) return res.json({ ok: true, summary: null });

  const summary = db.prepare(`
    SELECT
      property_name,
      ROUND(SUM(CASE WHEN amount > 0 THEN amount ELSE 0 END), 2) AS income,
      ROUND(ABS(SUM(CASE WHEN amount < 0 THEN amount ELSE 0 END)), 2) AS expenses,
      ROUND(SUM(amount), 2) AS net
    FROM transactions
    WHERE import_id = ? AND property_name IS NOT NULL
    GROUP BY property_name
    ORDER BY property_name
  `).all(importId);

  res.json({ ok: true, importId, summary });
});

app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);

  if (err instanceof multer.MulterError) {
    return res.status(400).send(`Upload failed: ${err.message}`);
  }

  const message = err?.message || 'Internal Server Error';

  if (
    /missing required sheet|workbook appears to be empty|excel workbook could not be read|only \.xlsx or \.xls|unsupported file type|missing year sheets/i.test(message)
  ) {
    return res.status(400).send(`Import failed: ${message}`);
  }

  return res.status(500).send(`Import failed: ${message}`);
});

app.listen(PORT, () => {
  console.log(`Financial importer running on http://localhost:${PORT}`);
});

        
