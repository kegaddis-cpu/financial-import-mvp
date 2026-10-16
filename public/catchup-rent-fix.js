/* Suggest standalone rent from per-property rent history; never save automatically. */
(()=>{
 'use strict';
 function init(){
  const button=document.getElementById('analyze');
  if(!button||button.dataset.rentFix)return;
  const original=button.onclick;
  if(typeof original!=='function')return;
  button.dataset.rentFix='true';
  const clean=s=>String(s||'').replace(/\s+/g,' ').trim().toLowerCase();
  const escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const dollars=n=>new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'}).format(n);
  button.onclick=function(event){
   original.call(this,event);
   const from=document.getElementById('from').value;
   const to=document.getElementById('through').value;
   if(typeof validDate!=='function'||!validDate(from)||!validDate(to)||from>to)return;
   if(typeof ledger==='undefined'||typeof candidates==='undefined')return;
   if(document.getElementById('scope').value==='personal')return;
   const historical=ledger.filter(t=>t.finance_scope==='rental'&&validDate(t.txn_date)&&t.txn_date<from);
   const hasTenants=(property,month)=>ledger.some(t=>t.finance_scope==='rental'&&t.property_name===property&&t.txn_date?.slice(0,7)===month&&Number(t.amount)>0&&/^rent from\s+\S/i.test(String(t.description||'').trim()));
   const groups=new Map();
   for(const t of historical){
    if(!Number.isFinite(Number(t.amount))||Number(t.amount)<=0||clean(t.description)!=='rent'||!t.property_name||hasTenants(t.property_name,t.txn_date.slice(0,7)))continue;
    if(!groups.has(t.property_name))groups.set(t.property_name,[]);
    groups.get(t.property_name).push(t);
   }
   let added=0;
   for(const [property,allRecords] of groups){
    const recentMonths=[...new Set(allRecords.map(t=>t.txn_date.slice(0,7)))].sort().reverse().slice(0,3);
    if(recentMonths.length<2)continue;
    const records=allRecords.filter(t=>recentMonths.includes(t.txn_date.slice(0,7)));
    records.sort((a,b)=>b.txn_date.localeCompare(a.txn_date)||Number(b.id)-Number(a.id));
    const last=records[0];
    if(ledger.some(t=>t.finance_scope==='rental'&&t.property_name===property&&t.txn_date<from&&t.txn_date.slice(0,7)>=last.txn_date.slice(0,7)&&Number(t.amount)>0&&/^rent from\s+\S/i.test(String(t.description||'').trim())))continue;
    if(candidates.some(c=>c.scope==='rental'&&c.property===property&&clean(c.description)==='rent'))continue;
    const c={scope:'rental',property,description:last.description,category:last.category||'rent',amount:Number(last.amount),last:last.txn_date,day:Math.min(28,Number(last.txn_date.slice(8))),months:recentMonths.length,records:records.map(t=>({date:t.txn_date,amount:Number(t.amount)}))};
    const i=candidates.length;candidates.push(c);added++;
    const tr=document.createElement('tr');
    tr.innerHTML=`<td><input type="checkbox" id="use${i}" aria-label="Select rent for ${escape(property)}"></td><td>rental<br>${escape(property)}</td><td>${escape(c.description)}<br><small>${escape(c.category)}</small></td><td>${c.records.map(r=>escape(r.date)+' '+dollars(r.amount)).join('<br>')}<br><small>Historical evidence only. Confirm rent still applies, verify the amount, and adjust the due day.</small></td><td><input id="amount${i}" type="number" min="0.01" step="0.01" value="${Math.abs(c.amount).toFixed(2)}" aria-label="Expected rent amount"></td><td><input id="day${i}" type="number" min="1" max="28" step="1" value="${c.day}" aria-label="Rent due day"></td>`;
    const body=document.getElementById('candidates');
    body.querySelectorAll('td[colspan]').forEach(td=>td.parentElement.remove());body.appendChild(tr);
   }
   document.getElementById('status').textContent=candidates.length+' possible items found, including '+added+' standalone rent item(s). Historical rent may be old: confirm it still applies before selecting. Nothing has been saved.';
  };
  const note=document.createElement('p');
  note.style.cssText='font-size:14px;color:#60756e';
  note.textContent='Rent fix v2 active: candidates use each property\'s last three recorded generic-rent months before the review start, requiring at least two months. Old history does not prove rent continues. Tenant replacement checks remain active. Confirm the amount and due day, and save actual payments only after verifying them.';
  button.after(note);
 }
 if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
