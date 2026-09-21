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
    ORDER BY id
