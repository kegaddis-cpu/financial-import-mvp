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
for (const dir of [uploadsDir, publicDir, viewsDir]) fs.mkdirSync(dir, { recursive: true });
console.log('DATABASE_URL present:', !!process.env.DATABASE_URL);
const pool = new Pool({ connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false });
app.set('views', viewsDir);
app.set('view engine', 'ejs');
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(publicDir));
const upload = multer({ dest: uploadsDir, limits: { fileSize: 20 * 1024 * 1024 } });
const query = (sql, params = []) => pool.query(sql, params);
function removeUploadedFile(filePath) { if (filePath) { try { fs.unlinkSync(filePath); } catch (_) {} } }
async function initDb() {
  await query(`CREATE TABLE IF NOT EXISTS imports (
    id SERIAL PRIMARY KEY, filename TEXT NOT NULL, imported_at TIMESTAMP NOT NULL,
    sheet_count INTEGER DEFAULT 0, row_count INTEGER DEFAULT 0, notes TEXT)`);
  await query(`CREATE TABLE IF NOT EXISTS account_snapshots (
    id SERIAL PRIMARY KEY, import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
    account_name TEXT, balance NUMERIC, as_of TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)`);
  await query(`CREATE TABLE IF NOT EXISTS transactions (
    id SERIAL PRIMARY KEY, import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
    property_name TEXT, txn_date TEXT, description TEXT, category TEXT, amount NUMERIC,
    direction TEXT, source_row INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    reason TEXT, source_category TEXT, income_amount NUMERIC, expense_amount NUMERIC,
    year_tag INTEGER, source_sheet TEXT)`);
  await query(`CREATE TABLE IF NOT EXISTS property_values (
    id SERIAL PRIMARY KEY, import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
    property_name TEXT, property_value NUMERIC, as_of TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)`);
  await query(`CREATE TABLE IF NOT EXISTS import_issues (
    id SERIAL PRIMARY KEY, import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
    issue_type TEXT, message TEXT, row_number INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)`);
}
async function getLatestImport() { return (await query('SELECT * FROM imports ORDER BY id DESC LIMIT 1')).rows[0] || null; }
async function getRecentImports(limit = 10) { return (await query('SELECT * FROM imports ORDER BY id DESC LIMIT $1', [limit])).rows; }
async function getStats() {
  const r = (await query(`SELECT
    (SELECT COUNT(*)::int FROM imports) AS "importCount",
    (SELECT COUNT(*)::int FROM transactions) AS "transactionCount",
    (SELECT COUNT(*)::int FROM account_snapshots) AS "accountCount",
    (SELECT COUNT(*)::int FROM property_values) AS "propertyValueCount",
    (SELECT COUNT(*)::int FROM import_issues) AS "issueCount"`)).rows[0];
  return { ...r, accounts: r.accountCount, transactions: r.transactionCount, properties: r.propertyValueCount, issues: r.issueCount };
}
async function getLatestImportStats() {
  const latest = await getLatestImport();
  if (!latest) return { accounts: 0, transactions: 0, properties: 0, propertyValues: 0, issues: 0 };
  return (await query(`SELECT
    (SELECT COUNT(*)::int FROM account_snapshots WHERE import_id=$1) AS accounts,
    (SELECT COUNT(*)::int FROM transactions WHERE import_id=$1) AS transactions,
    (SELECT COUNT(DISTINCT property_name)::int FROM transactions WHERE import_id=$1 AND TRIM(COALESCE(property_name,''))<>'') AS properties,
    (SELECT COUNT(*)::int FROM property_values WHERE import_id=$1) AS "propertyValues",
    (SELECT COUNT(*)::int FROM import_issues WHERE import_id=$1) AS issues`, [latest.id])).rows[0];
}
async function getAvailableMonths() {
  return (await query(`SELECT DISTINCT substr(txn_date,1,7) AS month FROM transactions
    WHERE txn_date ~ '^\\d{4}-\\d{2}-\\d{2}$' ORDER BY month DESC`)).rows.map(r=>r.month);
}
async function getSelectedMonth(requested) {
  const months = await getAvailableMonths();
  return months.includes(requested) ? requested : months[0] || null;
}
async function getMonthlyCashFlow(month) {
  if (!month) return [];
  return (await query(`SELECT property_name,
    ROUND(SUM(CASE WHEN amount>0 THEN amount ELSE 0 END)::numeric,2) AS income_total,
    ROUND(SUM(CASE WHEN amount<0 THEN ABS(amount) ELSE 0 END)::numeric,2) AS expense_total,
    ROUND(SUM(amount)::numeric,2) AS net_total FROM transactions
    WHERE substr(txn_date,1,7)=$1 AND TRIM(COALESCE(property_name,''))<>''
    GROUP BY property_name ORDER BY property_name`, [month])).rows;
}
async function getRecentTransactions(limit=20) {
  return (await query(`SELECT txn_date,property_name,COALESCE(reason,description,'') AS reason,
    COALESCE(source_category,category,'') AS source_category,amount FROM transactions
    WHERE txn_date IS NOT NULL ORDER BY txn_date DESC,id DESC LIMIT $1`, [limit])).rows;
}
async function getLatestPropertyValues(limit=20) {
  return (await query(`SELECT property_name AS address,property_value AS estimate,NULL AS profit_loss,
    as_of AS snapshot_date FROM property_values WHERE TRIM(COALESCE(property_name,''))<>''
    ORDER BY as_of DESC,id DESC LIMIT $1`, [limit])).rows;
}
async function getMissingRecurringExpenses(month) {
  if (!month) return [];
  return (await query(`WITH bounds AS (
    SELECT to_date($1 || '-01','YYYY-MM-DD') AS start_date
  ), prior AS (
    SELECT substr(t.txn_date,1,7) AS txn_month,TRIM(t.property_name) AS property_name,
      LOWER(TRIM(COALESCE(t.reason,t.description,''))) AS normalized_reason,
      TRIM(COALESCE(t.reason,t.description,'')) AS display_reason,
      t.txn_date,ABS(t.amount) AS expense
    FROM transactions t CROSS JOIN bounds b
    WHERE t.amount<0 AND t.txn_date >= to_char(b.start_date-interval '3 months','YYYY-MM-DD')
      AND t.txn_date < to_char(b.start_date,'YYYY-MM-DD')
      AND TRIM(COALESCE(t.property_name,''))<>'' AND TRIM(COALESCE(t.reason,t.description,''))<>''
      AND LOWER(TRIM(t.property_name)) NOT IN ('charity','car','insurance','ethan','chloe')
  ), candidates AS (
    SELECT property_name,normalized_reason,MIN(display_reason) AS reason,
      COUNT(DISTINCT txn_month)::int AS months_present,MAX(txn_date) AS last_seen_date,
      ROUND(AVG(expense)::numeric,2) AS typical_amount
    FROM prior GROUP BY property_name,normalized_reason HAVING COUNT(DISTINCT txn_month)>=2
  ) SELECT c.* FROM candidates c WHERE NOT EXISTS (
    SELECT 1 FROM transactions t WHERE t.amount<0 AND substr(t.txn_date,1,7)=$1
      AND TRIM(t.property_name)=c.property_name
      AND LOWER(TRIM(COALESCE(t.reason,t.description,'')))=c.normalized_reason
  ) ORDER BY c.property_name,c.reason`, [month])).rows;
}
app.use((req,res,next)=>{
  for (const name of ['recentImports','monthlyCashFlow','accounts','properties','propertyValues','issues',
    'accountSnapshots','propertyRollup','importIssues','latestIssues','latestProperties','portfolioProperties',
    'recentTransactions','transactions','propertiesWithTransactions','propertyCards','dashboardCards',
    'alerts','warnings','errors','sales','expenses','latestPropertyValues','availableMonths','missingRecurringExpenses']) res.locals[name]=[];
  for (const name of ['stats','latestImportStats','importSummary','summary']) res.locals[name]={};
  res.locals.latestImport=null; res.locals.selectedMonth=null; next();
});
const route = handler => (req,res,next) => Promise.resolve(handler(req,res)).catch(next);
app.get('/',route(async(req,res)=>{
  const selectedMonth=await getSelectedMonth(req.query.month);
  res.render('index',{latestImport:await getLatestImport(),recentImports:await getRecentImports(10),
    stats:await getStats(),latestImportStats:await getLatestImportStats(),
    monthlyCashFlow:await getMonthlyCashFlow(selectedMonth),recentTransactions:await getRecentTransactions(20),
    latestPropertyValues:await getLatestPropertyValues(20),selectedMonth});
}));
app.get('/monthly-review',route(async(req,res)=>{
  const availableMonths=await getAvailableMonths(),selectedMonth=await getSelectedMonth(req.query.month);
  res.render('monthly-review',{latestImport:await getLatestImport(),stats:await getStats(),availableMonths,selectedMonth,
    monthlyCashFlow:await getMonthlyCashFlow(selectedMonth),missingRecurringExpenses:await getMissingRecurringExpenses(selectedMonth)});
}));
app.get('/setup',route(async(req,res)=>{
  const latestImport=await getLatestImport();
  const issues=latestImport ? (await query('SELECT issue_type,message,row_number FROM import_issues WHERE import_id=$1 ORDER BY id',[latestImport.id])).rows : [];
  res.render('setup',{latestImport,stats:await getStats(),latestImportStats:await getLatestImportStats(),recentImports:await getRecentImports(10),issues});
}));

function installWorkbookImporter(app, { pool, upload, XLSX, removeUploadedFile }) {
  const text = v => String(v == null ? '' : v).replace(/\s+/g,' ').trim();
  const num = v => {
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    const s=text(v); if (!s) return null;
    const n=Number(s.replace(/[$,()]/g,''));
    return Number.isFinite(n) ? (/^\(.*\)$/.test(s) ? -Math.abs(n) : n) : null;
  };
  function date(v) {
    let y,m,d;
    if (typeof v==='number') { const p=XLSX.SSF.parse_date_code(v); if (!p) return null; ({y,m,d}=p); }
    else {
      let p=text(v).match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (p) [,y,m,d]=p;
      else { p=text(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/); if (!p) return null;
        [,m,d,y]=p; if (y.length===2) y=2000+Number(y); }
    }
    y=Number(y);m=Number(m);d=Number(d);
    const t=new Date(Date.UTC(y,m-1,d));
    if(t.getUTCFullYear()!==y||t.getUTCMonth()!==m-1||t.getUTCDate()!==d) return null;
    return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
  }
  app.post('/setup/import',upload.single('workbook'),async(req,res)=>{
    if(!req.file) return res.status(400).send('No workbook file uploaded');
    let client;
    try {
      const wb=XLSX.readFile(req.file.path,{cellDates:false});
      client=await pool.connect();await client.query('BEGIN');
      // Serialize workbook imports so duplicate checks do not race.
      await client.query('SELECT pg_advisory_xact_lock(7642301)');
      const id=(await client.query(`INSERT INTO imports(filename,imported_at,sheet_count,row_count,notes)
        VALUES($1,CURRENT_TIMESTAMP,$2,0,$3) RETURNING id`,[req.file.originalname,wb.SheetNames.length,
        'Explicit payment columns; ambiguous records held for review; non-payment sheets excluded'])).rows[0].id;
      const issue=(type,sheet,row,msg)=>client.query(`INSERT INTO import_issues(import_id,issue_type,message,row_number)
        VALUES($1,$2,$3,$4)`,[id,type,`${sheet}: ${msg}`,row]);
      let count=0;const seen=new Set();
      for(const sheet of wb.SheetNames) {
        const match=text(sheet).match(/^payments\s+(\d{2})$/i);if(!match)continue;
        const year=2000+Number(match[1]);
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1,raw:true,blankrows:true,range:0,defval:null});
        const h=rows.findIndex(r=>text(r[0]).toLowerCase()==='date'&&r.some(c=>text(c).toLowerCase()==='amount'));
        if(h<0){await issue('missing_header',sheet,null,'Required header not found; sheet skipped');continue;}
        const headers=rows[h].map(v=>text(v).toLowerCase());
        const dc=headers.indexOf('date'),ac=headers.indexOf('amount'),pc=headers.indexOf('property'),rc=headers.indexOf('reason');
        if(dc<0||ac<0||rc<0){await issue('missing_header',sheet,h+1,'Date, Amount or Reason missing');continue;}
        for(let i=h+1;i<rows.length;i++) {
          const row=rows[i]||[];if(!row.some(v=>text(v)))continue;
          const dt=date(row[dc]);
          if(!dt){if(typeof row[dc]==='number'||/^\d/.test(text(row[dc])))await issue('invalid_date',sheet,i+1,`Date unparsed: ${text(row[dc])}; row held`);continue;}
          const amount=num(row[ac]),reason=text(row[rc]),label=pc<0?'':text(row[pc]),lower=label.toLowerCase();
          if(/\b(vistazo|sierra|singer)\b/i.test(label+' '+reason))continue;
          const problems=[];let property=null;
          if(Number(dt.slice(0,4))!==year)problems.push('date year differs from sheet; original not corrected');
          if(amount===null)problems.push('invalid or missing Amount');
          if(!reason)problems.push('missing Reason');
          if(['fishhawk','fishawk','bridgecrossing','bridgecrossing dr'].includes(lower))property='Bridgecrossing Dr';
          else if(['carlton fields','carlton fields dr','riverview','riverview rental'].includes(lower))property='Carlton Fields Dr';
          else if(['blue plume','blue plume ct'].includes(lower))property='Blue Plume Ct';
          else if(['the boys','riverview boys'].includes(lower)) {
            if(amount>0&&/\brent\b/i.test(reason))property=dt<'2025-08-01'?'Blue Plume Ct':'Carlton Fields Dr';
            else problems.push('tenant label on non-rent entry: actual property needs confirmation');
          } else problems.push('outside confirmed property mapping');
          if(/^income$/i.test(reason))problems.push('income source unconfirmed');
          if(/deposit/i.test(reason))problems.push('deposit requires separate classification');
          const key=JSON.stringify([dt,property||lower,amount,reason.toLowerCase()]);
          if(seen.has(key))problems.push('possible duplicate within workbook');seen.add(key);
          if(!problems.length){const existing=await client.query(`SELECT id FROM transactions WHERE txn_date=$1
            AND property_name=$2 AND amount=$3 AND LOWER(TRIM(COALESCE(reason,description,'')))=$4 LIMIT 1`,
            [dt,property,amount,reason.toLowerCase()]);if(existing.rows.length)problems.push('possible duplicate already in database');}
          if(problems.length){await issue('review_required',sheet,i+1,`${problems.join('; ')}. Original: ${JSON.stringify({date:dt,label,amount:row[ac],reason})}`);continue;}
          await client.query(`INSERT INTO transactions(import_id,property_name,txn_date,description,category,amount,
            direction,source_row,reason,source_category,income_amount,expense_amount,year_tag,source_sheet)
            VALUES($1,$2,$3,$4,$5,$6,$7,$8,$4,$5,$9,$10,$11,$12)`,
            [id,property,dt,reason,amount>0&&/\brent\b/i.test(reason)?'rent':'unclassified',amount,
            amount<0?'outflow':'inflow',i+1,amount>0?amount:0,amount<0?Math.abs(amount):0,Number(dt.slice(0,4)),sheet]);count++;
        }
      }
      await issue('import_scope','Workbook',null,'Payments only. Accounts, Property Values, Taxes, Cleanup Notes and summary sheets intentionally skipped. Review-required rows are held as issues, not included in totals.');
      await client.query('UPDATE imports SET row_count=$2 WHERE id=$1',[id,count]);await client.query('COMMIT');
      res.redirect('/setup');
    } catch(err) {
      if(client){try{await client.query('ROLLBACK');}catch(_){}}
      console.error('Workbook import failed:',err);res.status(500).send('Import failed; database changes rolled back. Check server logs.');
    } finally {if(client)client.release();removeUploadedFile(req.file.path);}
  });
}


installWorkbookImporter(app, { pool, upload, XLSX, removeUploadedFile });
app.get('/data/imports',route(async(req,res)=>{
  const latestImport=await getLatestImport(),imports=await getRecentImports(100);
  const issues=latestImport ? (await query('SELECT issue_type,message,row_number,created_at FROM import_issues WHERE import_id=$1 ORDER BY id DESC',[latestImport.id])).rows : [];
  res.render('imports',{latestImport,imports,recentImports:imports,issues,importIssues:issues,
    stats:await getStats(),latestImportStats:await getLatestImportStats()});
}));
app.get('/imports/:id',route(async(req,res)=>{
  const importId=Number(req.params.id);
  if (!Number.isInteger(importId)||importId<=0) return res.status(400).send('Invalid import ID');
  const importRecord=(await query('SELECT * FROM imports WHERE id=$1',[importId])).rows[0];
  if (!importRecord) return res.status(404).send('Import not found');
  const counts=(await query(`SELECT
    (SELECT COUNT(*)::int FROM transactions WHERE import_id=$1) AS transactions,
    (SELECT COUNT(*)::int FROM account_snapshots WHERE import_id=$1) AS accounts,
    (SELECT COUNT(*)::int FROM property_values WHERE import_id=$1) AS property_values,
    (SELECT COUNT(*)::int FROM import_issues WHERE import_id=$1) AS issues`,[importId])).rows[0];
  const issues=(await query('SELECT issue_type,message,row_number,created_at FROM import_issues WHERE import_id=$1 ORDER BY id',[importId])).rows;
  res.render('import-detail',{importRecord,importId,counts,issues});
}));
app.get('/reports/cash-flow',route(async(req,res)=>{
  const latestImport=await getLatestImport();
  if (!latestImport) return res.status(400).send('No imports found');
  const propertyFilter=String(req.query.property_name||'').trim();
  const params=[latestImport.id];
  if (propertyFilter) params.push(propertyFilter);
  const rows=(await query(`SELECT property_name,substr(txn_date,1,7) AS month,COUNT(*)::int AS txn_count,
    ROUND(SUM(CASE WHEN amount>0 THEN amount ELSE 0 END)::numeric,2) AS income_total,
    ROUND(SUM(CASE WHEN amount<0 THEN ABS(amount) ELSE 0 END)::numeric,2) AS expense_total,
    ROUND(SUM(amount)::numeric,2) AS net_total FROM transactions
    WHERE import_id=$1 AND TRIM(COALESCE(property_name,''))<>'' AND txn_date IS NOT NULL
    ${propertyFilter ? 'AND property_name=$2' : ''}
    GROUP BY property_name,substr(txn_date,1,7) ORDER BY property_name,month`,params)).rows;
  const properties=(await query(`SELECT DISTINCT property_name FROM transactions WHERE import_id=$1
    AND TRIM(COALESCE(property_name,''))<>'' ORDER BY property_name`,[latestImport.id])).rows;
  res.render('cash-flow-report',{latestImportId:latestImport.id,propertyFilter,properties,rows});
}));
app.post('/imports/rollback-last',route(async(req,res)=>{
  const latestImport=await getLatestImport();
  if (!latestImport) return res.status(400).send('No imports found to roll back');
  await query('DELETE FROM imports WHERE id=$1',[latestImport.id]);
  res.redirect('/data/imports');
}));
app.get('/health',async(req,res)=>{
  try { await query('SELECT 1'); res.json({ok:true}); }
  catch (_) { res.status(500).json({ok:false,error:'Database unavailable'}); }
});
app.use((req,res)=>res.status(404).send(`Not found: ${req.method} ${req.originalUrl}`));
app.use((err,req,res,next)=>{
  if (req.file) removeUploadedFile(req.file.path);
  console.error(err);
  if (err instanceof multer.MulterError) return res.status(400).send(`Upload error: ${err.message}`);
  res.status(500).send('Internal Server Error. Check server logs.');
});
initDb().then(()=>app.listen(PORT,()=>console.log(`Server listening on port ${PORT}`)))
  .catch(err=>{console.error('Database initialization failed:',err);process.exit(1);});
