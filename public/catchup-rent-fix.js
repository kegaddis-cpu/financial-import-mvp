/* Restore standalone generic-rent candidates without bypassing tenant overlap protection. */
(()=>{
 'use strict';
 function init(){
  const button=document.getElementById('analyze');
  if(!button||button.dataset.rentFix)return;
  button.dataset.rentFix='true';
  const original=button.onclick;
  if(typeof original!=='function')return;
  const clean=s=>String(s||'').replace(/\s+/g,' ').trim().toLowerCase();
  const escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const dollars=n=>new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'}).format(n);
  button.onclick=function(event){
   const from=document.getElementById('from').value;
   const to=document.getElementById('through').value;
   if(!from||!to||from>to)return original.call(this,event);
   original.call(this,event);
   if(typeof ledger==='undefined'||typeof candidates==='undefined')return;
   const scope=document.getElementById('scope').value;
   if(scope==='personal')return;
   const historical=ledger.filter(t=>t.finance_scope==='rental'&&/^\d{4}-\d{2}-\d{2}$/.test(t.txn_date||'')&&t.txn_date<from);
   const recentMonths=[...new Set(historical.map(t=>t.txn_date.slice(0,7)))].sort().reverse().slice(0,3);
   const hasTenants=(property,month)=>ledger.some(t=>t.finance_scope==='rental'&&t.property_name===property&&t.txn_date?.slice(0,7)===month&&Number(t.amount)>0&&/^rent from\s+\S/i.test(String(t.description||'').trim()));
   const groups=new Map();
   for(const t of historical){
    if(Number(t.amount)<=0||clean(t.description)!=='rent'||!t.property_name||!recentMonths.includes(t.txn_date.slice(0,7))||hasTenants(t.property_name,t.txn_date.slice(0,7)))continue;
    if(!groups.has(t.property_name))groups.set(t.property_name,[]);
    groups.get(t.property_name).push(t);
   }
   let added=0;
   for(const [property,records] of groups){
    const months=new Set(records.map(t=>t.txn_date.slice(0,7)));
    if(months.size<2)continue;
    records.sort((a,b)=>b.txn_date.localeCompare(a.txn_date)||Number(b.id)-Number(a.id));
    const last=records[0];
    // Do not carry forward an older generic schedule after tenant payments have replaced it.
    if(ledger.some(t=>t.finance_scope==='rental'&&t.property_name===property&&t.txn_date<from&&t.txn_date.slice(0,7)>=last.txn_date.slice(0,7)&&Number(t.amount)>0&&/^rent from\s+\S/i.test(String(t.description||'').trim())))continue;
    if(candidates.some(c=>c.scope==='rental'&&c.property===property&&clean(c.description)==='rent'))continue;
    const c={scope:'rental',property,description:last.description,category:last.category||'rent',amount:Number(last.amount),last:last.txn_date,day:Math.min(28,Number(last.txn_date.slice(8))),months:months.size,records:records.map(t=>({date:t.txn_date,amount:Number(t.amount)}))};
    const i=candidates.length;candidates.push(c);added++;
    const tr=document.createElement('tr');
    tr.innerHTML=`<td><input type="checkbox" id="use${i}" aria-label="Select rent for ${escape(property)}"></td><td>rental<br>${escape(property)}</td><td>${escape(c.description)}<br><small>${escape(c.category)}</small></td><td>${c.records.map(r=>escape(r.date)+' '+dollars(r.amount)).join('<br>')}</td><td><input id="amount${i}" type="number" min="0.01" step="0.01" value="${Math.abs(c.amount).toFixed(2)}" aria-label="Expected rent amount"></td><td><input id="day${i}" type="number" min="1" max="28" step="1" value="${c.day}" aria-label="Rent due day"></td>`;
    const body=document.getElementById('candidates');
    body.querySelectorAll('td[colspan]').forEach(td=>td.parentElement.remove());body.appendChild(tr);
   }
   document.getElementById('status').textContent=candidates.length+' possible items found, including '+added+' standalone rent item(s). Review each before carrying forward.';
  };
  const note=document.createElement('p');
  note.style.cssText='font-size:14px;color:#60756e';
  note.textContent='Rent fix active: standalone generic Rent can be suggested. Generic rent is not carried forward where individual tenant payments have replaced it. Save actual payments only after verifying them.';
  button.after(note);
 }
 if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
