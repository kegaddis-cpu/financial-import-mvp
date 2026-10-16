/* Adds selection controls only; never saves or confirms transactions. */
(()=>{
 'use strict';
 function init(){
  const body=document.getElementById('candidates');
  if(!body||document.getElementById('catchup-select-controls'))return;
  const table=body.closest('table');
  const container=table.closest('.scroll')||table;
  const bar=document.createElement('div');
  bar.id='catchup-select-controls';
  bar.style.cssText='display:flex;align-items:center;flex-wrap:wrap;gap:12px;margin:16px 0';
  const label=document.createElement('label');
  label.style.cssText='display:flex;align-items:center;gap:8px;margin:0';
  const master=document.createElement('input');
  master.type='checkbox';master.id='catchup-select-all';
  master.style.cssText='width:auto;margin:0';
  const text=document.createElement('span');text.textContent='Select all for Use';
  label.append(master,text);
  const clear=document.createElement('button');
  clear.type='button';clear.textContent='Clear selection';
  const count=document.createElement('span');count.setAttribute('role','status');
  const note=document.createElement('p');
  note.textContent='Selecting items only includes them in the review queue. It does not save or confirm payments.';
  note.style.cssText='font-size:14px;color:#60756e';
  bar.append(label,clear,count);container.before(bar,note);
  const boxes=()=>Array.from(body.querySelectorAll('input[type="checkbox"]'));
  function update(){
   const items=boxes();const selected=items.filter(b=>b.checked).length;
   master.checked=items.length>0&&selected===items.length;
   master.indeterminate=selected>0&&selected<items.length;
   master.disabled=items.length===0||items.every(b=>b.disabled);
   clear.disabled=selected===0||items.every(b=>b.disabled);
   count.textContent=selected+' of '+items.length+' selected';
  }
  function setAll(value){
   boxes().filter(b=>!b.disabled).forEach(b=>{b.checked=value;b.dispatchEvent(new Event('change',{bubbles:true}));});
   update();
  }
  master.addEventListener('change',()=>setAll(master.checked));
  clear.addEventListener('click',()=>setAll(false));
  body.addEventListener('change',update);
  new MutationObserver(update).observe(body,{childList:true,subtree:true,attributes:true,attributeFilter:['disabled']});
  update();
 }
 if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
