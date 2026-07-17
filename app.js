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
    'busch g', 'boca', 'fun', 'income ', 'personal', 'health'
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
        if (!r['DATE'] && !r['Bank'] && !r['Balance']) continue;
        if (!r['Bank'] || String(r['Bank']).trim() === 'Bank') continue;

        const balance = n(r['Balance']);
        if (balance == null) {
          addIssue(importId, sheetName, 'invalid_account_balance', i + 2, r, 'Balance could not be parsed');
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

    if (/^(Payments|Taxes)\s*\d+/i.test(sheetName)) {
      const yearTag = detectYearTag(sheetName);
      const category = /^Taxes/i.test(sheetName) ? 'taxes' : 'payments';

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const rowNum = i + 2;

        if (isProbablySummaryRow(r)) continue;

        const rawDate = r['Date'] ?? r['DATE'];
        const rawProperty = s(r['Property '] || r['Property']);
        const rawReason = s(r['Reason '] || r['Reason']);
        const amount = n(r['Amount '] ?? r['Amount']);

        const incomeKey = Object.keys(r).find(k => String(k).trim().startsWith('Income'));
        const expenseKey = Object.keys(r).find(k => String(k).trim().startsWith('Expenses'));
        const income = incomeKey ? n(r[incomeKey]) : null;
        const expense = expenseKey ? n(r[expenseKey]) : null;

        const looksLikeTransaction =
          excelDateToIso(rawDate) || rawProperty || rawReason || income != null || expense != null || amount != null;

        if (!looksLikeTransaction) continue;

        if (rawProperty && isNonPortfolioCategory(rawProperty)) continue;

        let propertyName = sanitizeImportedPropertyName(rawProperty, validProperties);

        if (!propertyName && /^Payments\s*21$/i.test(sheetName)) {
          const inferred = inferPropertyFromReason(rawReason);
          propertyName = sanitizeImportedPropertyName(inferred, validProperties);
        }

        if (!propertyName && rawProperty && !isNonPortfolioCategory(rawProperty)) {
          const inferred = inferPropertyFromReason(rawReason);
          propertyName = sanitizeImportedPropertyName(inferred, validProperties) || null;
        }

        if (!propertyName && rawProperty && !isNonPortfolioCategory(rawProperty)) {
          addIssue(importId, sheetName, 'unmatched_property_name', rowNum, r, `Ignored non-portfolio property label: ${rawProperty}`);
          continue;
        }

        if (amount == null) {
          addIssue(importId, sheetName, 'invalid_transaction_amount', rowNum, r, 'Amount could not be parsed');
          continue;
        }

        insertTxn.run(
          importId,
          excelDateToIso(rawDate),
          propertyName,
          amount,
          rawReason,
          income,
          expense,
          yearTag,
          sheetName,
          category
        );
        rowCount++;
      }
      continue;
    }

    if (normalizeSheetKey(sheetName) === normalizeSheetKey('Property Values')) {
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const address = s(r['Address '] || r['Address']);
        const purchased = n(r['Purchased ' ] ?? r['Purchased']);
        const estimate = n(r['Estimate ' ] ?? r['Estimate']);

        if (!address && purchased == null && estimate == null) continue;
        if (!address) continue;

        insertValue.run(
          importId,
          normalizePropertyName(address),
          purchased,
          estimate,
          excelDateToIso(r['Unnamed: 3']),
          n(r['Profit/Loss'] ?? r['Profit / Loss']),
          n(r['Unnamed: 5']),
          excelDateToIso(r['Unnamed: 6'])
        );
        rowCount++;
      }
      continue;
    }

    if (normalizeSheetKey(sheetName) === normalizeSheetKey('Revenue and Expenses updated')) {
      recognizedRevenueSheet = true;
      continue;
    }
  }

  if (!recognizedRevenueSheet) {
    const attempted = findSheet(workbook, 'Revenue and Expenses updated');
    if (!attempted) {
      addIssue(importId, 'Revenue and Expenses updated', 'unmapped_sheet', null, {}, 'Sheet "Revenue and Expenses updated" not found');
    }
  }

  db.prepare('UPDATE imports SET row_count = ? WHERE id = ?').run(rowCount, importId);
  return importId;
}

app.get('/', (req, res) => {
  const latestImport = db.prepare('SELECT * FROM imports ORDER BY id DESC LIMIT 1').get();

  const stats = latestImport ? {
    accounts: db.prepare('SELECT COUNT(*) AS c FROM account_snapshots WHERE import_id = ?').get(latestImport.id).c,
    transactions: db.prepare('SELECT COUNT(*) AS c FROM transactions WHERE import_id = ?').get(latestImport.id).c,
    properties: db.prepare('SELECT COUNT(DISTINCT property_name) AS c FROM transactions WHERE import_id = ? AND property_name IS NOT NULL').get(latestImport.id).c,
    values: db.prepare('SELECT COUNT(*) AS c FROM property_values WHERE import_id = ?').get(latestImport.id).c,
    issues: db.prepare('SELECT COUNT(*) AS c FROM import_issues WHERE import_id = ?').get(latestImport.id).c
  } : null;

  const recentImports = db.prepare('SELECT * FROM imports ORDER BY id DESC LIMIT 10').all();
  res.render('index', { latestImport, stats, recentImports });
});

app.get('/import', (req, res) => {
  res.redirect('/');
});

app.post('/import', upload.single('financialFile'), (req, res) => {
  if (!req.file) return res.status(400).send('No file uploaded');

  try {
    const importId = importWorkbook(req.file.path, req.file.originalname);
    res.redirect(`/imports/${importId}`);
  } catch (err) {
    console.error(err);
    res.status(500).send(`Import failed: ${err.message}`);
  }
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
    LIMIT 200
  `).all(importId);

  res.render('import-detail', {
    imp,
    accounts,
    properties,
    propertyValues,
    issues
  });
});

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

app.listen(PORT, () => {
  console.log(`Financial importer running on http://localhost:${PORT}`);
});
