module.exports=function installSchedules(app,pool){
 const ready=pool.query(`CREATE TABLE IF NOT EXISTS recurring_schedules (
 id SERIAL PRIMARY KEY, description TEXT NOT NULL, finance_scope TEXT NOT NULL,
 property_name TEXT, category TEXT NOT NULL, amount NUMERIC(14,2) NOT NULL,
 start_date TEXT NOT NULL, end_date TEXT, due_day INTEGER NOT NULL,
 active BOOLEAN NOT NULL DEFAULT TRUE, revision INTEGER NOT NULL DEFAULT 1,
 created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP)`);
 ready.catch(e=>console.error('Schedule initialization failed:',e.message));
 const wrap=fn=>(req,res,next)=>Promise.resolve().then(()=>ready).then(()=>fn(req,res)).catch(next);
 const validDate=s=>/^\d{4}-\d{2}-\d{2}$/.test(s||'')&&Number.isFinite(Date.parse(s+'T00:00:00Z'))&&new Date(s+'T00:00:00Z').toISOString().slice(0,10)===s;
 app.get('/api/schedules',wrap(async(req,res)=>res.json({schedules:(await pool.query('SELECT * FROM recurring_schedules ORDER BY active DESC,description,id')).rows})));
 async function save(req,res){
  const b=req.body,scope=b.scope,description=String(b.description||'').trim(),category=String(b.category||'').trim();
  const amount=Number(b.amount),day=Number(b.day),start=b.start,end=b.end||null;
  const property=scope==='personal'?null:b.property;
  const properties=['Carlton Fields Dr','Bridgecrossing Dr','Blue Plume Ct'];
  if(!['rental','personal'].includes(scope)||!description||description.length>200||!category||category.length>80||!Number.isFinite(amount)||!amount||Math.abs(amount)>1e10||Math.abs(amount*100-Math.round(amount*100))>0.0001||!Number.isInteger(day)||day<1||day>28||!validDate(start)||(end&&(!validDate(end)||end<start))||(scope==='rental'&&!properties.includes(property))||typeof b.active!=='boolean')return res.status(400).json({error:'Check all fields. Monthly due day must be 1–28; amount must be nonzero with at most two decimals.'});
  const id=Number(req.params.id),editing=!!req.params.id;
  if(editing&&(!Number.isSafeInteger(id)||id<1||!Number.isInteger(b.revision)))return res.status(400).json({error:'Invalid schedule or revision'});
  const c=await pool.connect();try{await c.query('BEGIN');await c.query('SELECT pg_advisory_xact_lock(7642301)');
   const duplicate=await c.query(`SELECT id FROM recurring_schedules WHERE active AND $1::boolean AND finance_scope=$2 AND property_name IS NOT DISTINCT FROM $3 AND LOWER(TRIM(description))=LOWER(TRIM($4)) AND id<>$5 AND start_date<=COALESCE($6,'9999-12-31') AND COALESCE(end_date,'9999-12-31')>=$7 LIMIT 1`,[b.active,scope,property,description,editing?id:0,end,start]);
   if(duplicate.rows.length){await c.query('ROLLBACK');return res.status(409).json({error:'An overlapping active schedule with this description already exists.'});}
   if(scope==='rental'&&amount>0&&b.active&&(/^rent$/i.test(description)||/^rent from\s+\S/i.test(description))){
    const other=(await c.query(`SELECT description FROM recurring_schedules WHERE active AND finance_scope='rental' AND property_name=$1 AND amount>0 AND id<>$2 AND start_date<=COALESCE($3,'9999-12-31') AND COALESCE(end_date,'9999-12-31')>=$4`,[property,editing?id:0,end,start])).rows;
    if(other.some(r=>/^rent$/i.test(description)?/^rent from\s+\S/i.test(r.description):/^rent$/i.test(r.description))){await c.query('ROLLBACK');return res.status(409).json({error:'Generic rent overlaps individual tenant schedules for this property.'});}}
   const params=[description,scope,property,category,amount,start,end,day,b.active];let result;
   if(editing){result=await c.query(`UPDATE recurring_schedules SET description=$1,finance_scope=$2,property_name=$3,category=$4,amount=$5,start_date=$6,end_date=$7,due_day=$8,active=$9,revision=revision+1,updated_at=CURRENT_TIMESTAMP WHERE id=$10 AND revision=$11 RETURNING *`,[...params,id,b.revision]);if(!result.rows.length){await c.query('ROLLBACK');return res.status(409).json({error:'Schedule changed or no longer exists. Reload before editing.'});}}
   else result=await c.query(`INSERT INTO recurring_schedules(description,finance_scope,property_name,category,amount,start_date,end_date,due_day,active) VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9) RETURNING *`,params);
   await c.query('COMMIT');res.status(editing?200:201).json({schedule:result.rows[0]});
  }catch(e){await c.query('ROLLBACK');throw e;}finally{c.release();}
 }
 app.post('/api/schedules',wrap(save));app.put('/api/schedules/:id',wrap(save));
};
