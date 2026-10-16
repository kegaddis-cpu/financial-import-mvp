/* Workspace ledger display sorting. No database writes. */
(()=>{
 'use strict';
 function init(){
  const body=document.getElementById('rows');
  if(!body)return;
  const table=body.closest('table');
  const headers=Array.from(table.querySelectorAll('thead th'));
  // A newer full interface may already provide header sorting.
  if(table.querySelector('thead [data-sort]')||table.dataset.workspaceSorting)return;
  table.dataset.workspaceSorting='true';
  const columns=[{index:0,name:'Date',type:'date'},{index:1,name:'Scope / property',type:'text'},{index:2,name:'Description',type:'text'},{index:3,name:'Amount',type:'amount'},{index:4,name:'Source',type:'text'}];
  let selected=0,direction='desc';
  let rank=new WeakMap(),nextRank=0;
  const collator=new Intl.Collator(undefined,{numeric:true,sensitivity:'base'});
  const normalize=s=>String(s||'').replace(/\s+/g,' ').trim();
  function value(row,column){
   const cell=row.cells[column.index];
   const text=normalize(cell?cell.textContent:'');
   if(column.type==='amount'){
    let n=Number(text.replace(/[^0-9.\-]/g,''));
    if(/^\(.*\)$/.test(text))n=-Math.abs(n);
    return Number.isFinite(n)?n:0;
   }
   return text;
  }
  function apply(){
   observer.disconnect();
   try{
    const rows=Array.from(body.rows).filter(r=>r.cells.length>=5&&!r.cells[0].hasAttribute('colspan'));
    rows.forEach(r=>{if(!rank.has(r))rank.set(r,nextRank++);});
    const column=columns.find(c=>c.index===selected);
    const sorted=[...rows].sort((a,b)=>{
     const x=value(a,column),y=value(b,column);
     const comparison=column.type==='amount'?x-y:collator.compare(x,y);
     return comparison?(direction==='desc'?-comparison:comparison):rank.get(a)-rank.get(b);
    });
    if(sorted.some((row,i)=>row!==rows[i]))sorted.forEach(row=>body.appendChild(row));
    columns.forEach(c=>{
     const th=headers[c.index];if(!th)return;
     th.setAttribute('aria-sort',selected===c.index?(direction==='desc'?'descending':'ascending'):'none');
     const b=th.querySelector('button[data-workspace-sort]');
     b.textContent=c.name+(selected===c.index?(direction==='desc'?' ▼':' ▲'):' ↕');
     b.setAttribute('aria-label','Sort by '+c.name+(selected===c.index?', currently '+(direction==='desc'?'descending':'ascending'):''));
    });
   }finally{observer.observe(body,{childList:true});}
  }
  const observer=new MutationObserver(apply);
  columns.forEach(c=>{
   const th=headers[c.index];if(!th)return;
   const b=document.createElement('button');b.type='button';b.dataset.workspaceSort=String(c.index);
   b.style.cssText='font:inherit;font-weight:600;background:transparent;color:inherit;border:0;padding:4px 0;margin:0;width:auto;cursor:pointer;text-align:left;white-space:nowrap';
   b.addEventListener('click',()=>{
    direction=selected===c.index?(direction==='desc'?'asc':'desc'):(c.type==='date'||c.type==='amount'?'desc':'asc');
    selected=c.index;apply();
   });
   th.replaceChildren(b);
  });
  const note=document.createElement('p');note.style.cssText='font-size:14px;color:#60756e';
  note.textContent='Click a column heading to sort; click again to reverse. Default: newest transaction date first. Amounts sort by signed value.';
  const wrapper=table.closest('.scroll')||table;wrapper.before(note);
  apply();
 }
 if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
  
