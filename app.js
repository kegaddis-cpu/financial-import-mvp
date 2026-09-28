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
      id SERIAL
