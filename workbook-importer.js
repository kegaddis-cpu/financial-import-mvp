module.exports = function installWorkbookImporter(app, { pool, upload, XLSX, removeUploadedFile }) {
  const text = v => String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
  const number = v => {
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    const s = text(v);
    if (!s) return null;
    const n = Number(s.replace(/[$,()]/g, ''));
    return Number.isFinite(n) ? (/^\(.*\)$/.test(s) ? -Math.abs(n) : n) : null;
  };
  const date = v => {
    let y, m, d;
    if (typeof v === 'number') {
      const p = XLSX.SSF.parse_date_code(v);
      if (!p) return null;
      ({ y, m, d } = p);
    } else {
      const s = text(v);
      let p = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (p) [, y, m, d] = p;
      else {
        p = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
        if (!p) return null;
        [, m, d, y] = p;
        if (String(y).length === 2) y = 2000 + Number(y);
      }
    }
    y = Number(y); m = Number(m); d = Number(d);
    const test = new Date(Date.UTC(y, m - 1, d));
    if (test.getUTCFullYear() !== y || test.getUTCMonth() !== m - 1 || test.getUTCDate() !== d) return null;
    return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
  };
  app.post('/setup/import', upload.single('workbook'), async (req, res) => {
    if (!req.file) return res.status(400).send('No workbook file uploaded');
    let client;
    try {
      const workbook = XLSX.readFile(req.file.path, { cellDates: false });
      client = await pool.connect();
      await client.query('BEGIN');
      const result = await client.query(`INSERT INTO imports (filename, imported_at, sheet_count, row_count, notes)
        VALUES ($1, CURRENT_TIMESTAMP, $2, 0, $3) RETURNING id`,
        [req.file.originalname, workbook.SheetNames.length, 'Explicit-column parser; ambiguous rows held as issues; tax and summary sheets excluded']);
      const importId = result.rows[0].id;
      let count = 0;
      const issue = async (type, sheet, row, message) => client.query(
        'INSERT INTO import_issues (import_id, issue_type, message, row_number) VALUES ($1,$2,$3,$4)',
        [importId, type, `${sheet}: ${message}`, row]);
      const seen = new Set();
      for (const sheet of workbook.SheetNames) {
        const normalized = text(sheet).toLowerCase();
        if (!/^payments \d{2}$/.test(normalized)) continue;
        const sheetYear = 2000 + Number(normalized.slice(-2));
        const ws = workbook.Sheets[sheet];
        if (!ws['!ref']) continue;
        const range = XLSX.utils.decode_range(ws['!ref']);
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, blankrows: true, range: 0, defval: null });
        const headerIndex = rows.findIndex(r => text(r[0]).toLowerCase() === 'date' && r.some(c => text(c).toLowerCase() === 'amount'));
        if (headerIndex < 0) { await issue('missing_header', sheet, null, 'Date and Amount headers not found; sheet skipped'); continue; }
        const headers = rows[headerIndex].map(c => text(c).toLowerCase());
        const col = name => headers.indexOf(name);
        const dc = col('date'), ac = col('amount'), pc = col('property'), rc = col('reason');
        if (dc < 0 || ac < 0 || rc < 0) { await issue('missing_header', sheet, headerIndex+1, 'Required Date, Amount, Reason columns missing'); continue; }
        for (let i = headerIndex+1; i <= range.e.r; i++) {
          const row = rows[i] || [];
          if (!row.some(v => text(v))) continue;
          const txnDate = date(row[dc]);
          if (!txnDate) {
            if (typeof row[dc] === 'number' || /^\d/.test(text(row[dc]))) await issue('invalid_date', sheet, i+1, `Unparsed date ${text(row[dc])}; row held`);
            continue;
          }
          const amount = number(row[ac]), reason = text(row[rc]), label = pc < 0 ? '' : text(row[pc]);
          const lower = label.toLowerCase();
          if (/\b(vistazo|sierra|singer)\b/i.test(`${label} ${reason}`)) continue;
          const problems = [];
          if (Number(txnDate.slice(0,4)) !== sheetYear) problems.push('date year differs from sheet year; not corrected');
          if (amount === null) problems.push('missing or invalid Amount');
          if (!reason) problems.push('missing Reason');
          let property = null;
          if (['fishhawk','fishawk','bridgecrossing','bridgecrossing dr'].includes(lower)) property = 'Bridgecrossing Dr';
          else if (['carlton fields','carlton fields dr','riverview','riverview rental'].includes(lower)) property = 'Carlton Fields Dr';
          else if (['blue plume','blue plume ct'].includes(lower)) property = 'Blue Plume Ct';
          else if (['the boys','riverview boys'].includes(lower)) {
            if (amount > 0 && /\brent\b/i.test(reason)) property = txnDate < '2025-08-01' ? 'Blue Plume Ct' : 'Carlton Fields Dr';
            else problems.push('The Boys expense/non-rent label: actual property needs confirmation');
          } else problems.push('outside confirmed property mapping');
          if (/^income$/i.test(reason)) problems.push('income source unconfirmed; not classified as rent');
          if (/deposit/i.test(reason)) problems.push('deposit or combined rent/deposit needs classification');
          const key = JSON.stringify([txnDate, property || lower, amount, reason.toLowerCase()]);
          if (seen.has(key)) problems.push('possible duplicate within workbook');
          seen.add(key);
          if (!problems.length) {
            const existing = await client.query(`SELECT id FROM transactions WHERE txn_date=$1 AND property_name=$2
              AND amount=$3 AND LOWER(TRIM(COALESCE(reason,description,'')))=$4 LIMIT 1`, [txnDate, property, amount, reason.toLowerCase()]);
            if (existing.rows.length) problems.push('possible duplicate of existing database transaction');
          }
          if (problems.length) {
            await issue('review_required', sheet, i+1, `${problems.join('; ')}. Original: ${JSON.stringify({date:txnDate,property:label,amount:row[ac],reason})}`);
            continue;
          }
          await client.query(`INSERT INTO transactions
            (import_id,property_name,txn_date,description,category,amount,direction,source_row,reason,source_category,income_amount,expense_amount,year_tag,source_sheet)
            VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$4,$5,$9,$10,$11,$12)`,
            [importId, property, txnDate, reason, /\brent\b/i.test(reason) && amount > 0 ? 'rent' : 'unclassified', amount,
             amount < 0 ? 'outflow' : 'inflow', i+1, amount > 0 ? amount : 0, amount < 0 ? Math.abs(amount) : 0,
             Number(txnDate.slice(0,4)), sheet]);
          count++;
        }
      }
      await issue('import_scope', 'Workbook', null, 'Only Payments sheets imported. Accounts, Property Values, Taxes, Cleanup Notes and summaries intentionally excluded. Ambiguous records held for review, not discarded.');
      await client.query('UPDATE imports SET row_count=$2 WHERE id=$1', [importId,count]);
      await client.query('COMMIT');
      res.redirect('/setup');
    } catch (err) {
      if (client) { try { await client.query('ROLLBACK'); } catch (_) {} }
      console.error('Workbook import failed:', err);
      res.status(500).send('Workbook import failed; database transaction rolled back. Check server logs.');
    } finally {
      if (client) client.release();
      removeUploadedFile(req.file.path);
    }
  });
};
