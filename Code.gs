const SPREADSHEET_ID='1eOWiWZh0zTnVLnv53bTDXvlFJJG5ZbvR6Upr7_9uLb8';
const ORDERS_SHEET='ORDENES',EXPENSES_SHEET='GASTOS',PRODUCTS_SHEET='PRODUCTOS',CONFIG_SHEET='CONFIGURACION';
const LOGO_FILE_ID='1KNCCBhFm4vD92Jpi5rdKKgC6xxEWk1ea';
const APP_CACHE_KEY='PRINTA_APP_DATA_V3';

function doGet(){
  const out=HtmlService.createTemplateFromFile('App').evaluate().setTitle('PRINTA').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  out.append(clientEnhancements_());
  return out;
}
function getLogoDataUrl(){try{const b=DriveApp.getFileById(LOGO_FILE_ID).getBlob();return`data:${b.getContentType()};base64,${Utilities.base64Encode(b.getBytes())}`}catch(e){return''}}

function getAppData(){
  const cache=CacheService.getScriptCache();
  const cached=cache.get(APP_CACHE_KEY);
  if(cached){try{return JSON.parse(cached)}catch(e){}}
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const orders=readRows_(ss.getSheetByName(ORDERS_SHEET));
  const expenses=readRows_(ss.getSheetByName(EXPENSES_SHEET));
  const products=readRows_(ss.getSheetByName(PRODUCTS_SHEET));
  const settings=getSettings_(ss);
  const result={orders,expenses,products,settings,summary:buildSummary_(orders,expenses)};
  try{cache.put(APP_CACHE_KEY,JSON.stringify(result),120)}catch(e){}
  return result;
}
function invalidateCache_(){try{CacheService.getScriptCache().remove(APP_CACHE_KEY)}catch(e){}}

function calculateShippingDate(orderDate,origin){const ss=SpreadsheetApp.openById(SPREADSHEET_ID),settings=getSettings_(ss);return addBusinessDays_(parseDate_(orderDate)||new Date(),getHandlingDays_(origin,settings))}
function saveOrder(o){if(!o||!clean_(o.cliente)||!clean_(o.producto))throw new Error('Completa cliente y producto.');const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=ss.getSheetByName(ORDERS_SHEET),now=new Date(),id=o.id&&String(o.id).indexOf('TEMP-')!==0?o.id:nextId_(sh,'PC'),orderDate=o.fecha||formatDate_(now),settings=getSettings_(ss),shippingDate=addBusinessDays_(parseDate_(orderDate)||now,getHandlingDays_(o.origen,settings)),row=[id,orderDate,o.estado||'Pendiente',o.origen||'Otro',clean_(o.cliente),clean_(o.producto),number_(o.cantidad,1),clean_(o.notas),money_(o.pagado),shippingDate,clean_(o.foto),o.creadoEn||now,now];upsertById_(sh,id,row);invalidateCache_();return getAppData()}
function setOrderStatus(id,status){if(!['Pendiente','Listo'].includes(status))throw new Error('Estado inválido.');const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDERS_SHEET),r=findRowById_(sh,id);if(!r)throw new Error('Orden no encontrada.');sh.getRange(r,3).setValue(status);sh.getRange(r,13).setValue(new Date());invalidateCache_();return getAppData()}
function deleteOrder(id){deleteById_(ORDERS_SHEET,id);invalidateCache_();return getAppData()}
function saveExpense(e){if(!e||money_(e.cantidad)<=0)throw new Error('Escribe una cantidad válida.');const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET),now=new Date(),id=e.id||nextId_(sh,'GA');upsertById_(sh,id,[id,e.fecha||formatDate_(now),e.categoria||'Otro',clean_(e.lugar),money_(e.cantidad),clean_(e.nota),clean_(e.recibo),e.creadoEn||now]);invalidateCache_();return getAppData()}
function deleteExpense(id){deleteById_(EXPENSES_SHEET,id);invalidateCache_();return getAppData()}
function saveProduct(p){if(!p||!clean_(p.nombre))throw new Error('Escribe el nombre del producto.');const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PRODUCTS_SHEET),now=new Date(),id=p.id||nextId_(sh,'PR');upsertById_(sh,id,[id,clean_(p.nombre),clean_(p.foto),String(p.disponible)!=='false',p.creadoEn||now,now]);invalidateCache_();return getAppData()}
function deleteProduct(id){deleteById_(PRODUCTS_SHEET,id);invalidateCache_();return getAppData()}

function uploadFile(fileData,fileName,mimeType,ownerId,category){if(!fileData)return'';const folder=getUploadsFolder_(),bytes=Utilities.base64Decode(String(fileData).split(',').pop()),safe=`${ownerId||'GENERAL'}_${category||'archivo'}_${Date.now()}_${fileName||'archivo'}`,file=folder.createFile(Utilities.newBlob(bytes,mimeType||'application/octet-stream',safe));file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return file.getUrl()}
function getUploadsFolder_(){const ss=SpreadsheetApp.openById(SPREADSHEET_ID),c=ss.getSheetByName(CONFIG_SHEET),v=c.getRange(1,1,Math.max(c.getLastRow(),1),2).getValues();let id='',row=0;v.forEach((x,i)=>{if(x[0]==='CARPETA_FOTOS_ID'){id=x[1];row=i+1}});if(id){try{return DriveApp.getFolderById(id)}catch(e){}}const f=DriveApp.createFolder('PRINTA - Fotos y recibos');if(row)c.getRange(row,2).setValue(f.getId());return f}
function getSettings_(ss){const sh=ss.getSheetByName(CONFIG_SHEET),rows=sh.getRange(1,1,Math.max(sh.getLastRow(),1),2).getValues(),out={tiktokDays:2,printaDays:3};rows.forEach(r=>{if(r[0]==='TIKTOK_DIAS_ENVIO')out.tiktokDays=number_(r[1],2);if(r[0]==='PRINTA_DIAS_ENVIO')out.printaDays=number_(r[1],3)});return out}
function getHandlingDays_(origin,settings){return origin==='TikTok'?settings.tiktokDays:settings.printaDays}
function addBusinessDays_(date,days){const d=new Date(date);let added=0;while(added<days){d.setDate(d.getDate()+1);const day=d.getDay();if(day!==0&&day!==6)added++}return formatDate_(d)}
function parseDate_(value){if(!value)return null;const p=String(value).split('-').map(Number);return p.length===3?new Date(p[0],p[1]-1,p[2],12):new Date(value)}
function readRows_(sh){if(!sh||sh.getLastRow()<2)return[];const vals=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getDisplayValues(),h=vals.shift();return vals.filter(r=>r[0]).map(r=>{const o={};h.forEach((x,i)=>o[toKey_(x)]=r[i]);return o})}
function buildSummary_(o,e){const m=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM'),mo=o.filter(x=>String(x.fecha||'').startsWith(m)),me=e.filter(x=>String(x.fecha||'').startsWith(m)),ing=mo.reduce((s,x)=>s+money_(x.pagado),0),gas=me.reduce((s,x)=>s+money_(x.cantidad),0);return{pendientes:o.filter(x=>x.estado==='Pendiente').length,listos:o.filter(x=>x.estado==='Listo').length,ingresos:ing,gastos:gas,ganancia:ing-gas,tiktok:mo.filter(x=>x.origen==='TikTok').reduce((s,x)=>s+money_(x.pagado),0),printaCrea:mo.filter(x=>x.origen==='Printa Crea').reduce((s,x)=>s+money_(x.pagado),0)}}
function upsertById_(sh,id,row){const r=findRowById_(sh,id);r?sh.getRange(r,1,1,row.length).setValues([row]):sh.appendRow(row)}
function deleteById_(name,id){const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name),r=findRowById_(sh,id);if(r)sh.deleteRow(r)}
function findRowById_(sh,id){if(sh.getLastRow()<2)return 0;const m=sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(String(id)).matchEntireCell(true).findNext();return m?m.getRow():0}
function nextId_(sh,p){return`${p}-${String(Math.max(1,sh.getLastRow())).padStart(4,'0')}`}
function formatDate_(d){return Utilities.formatDate(d,'America/New_York','yyyy-MM-dd')}
function money_(v){const n=Number(String(v||0).replace(/[^0-9.-]/g,''));return isNaN(n)?0:n}
function number_(v,f){const n=Number(v);return isNaN(n)?f:n}
function clean_(v){return String(v==null?'':v).trim()}
function toKey_(h){return String(h).toLowerCase().replace(/_([a-z])/g,(_,c)=>c.toUpperCase())}

function clientEnhancements_(){return `
<style>
.order>.gallery{display:none!important}
.btn.saving{position:relative;color:transparent!important;pointer-events:none}.btn.saving:after{content:'';position:absolute;left:50%;top:50%;width:18px;height:18px;margin:-11px 0 0 -11px;border:3px solid #ffffff66;border-top-color:#fff;border-radius:50%;animation:printaSpin .65s linear infinite}.btn.light.saving:after{border-color:#0002;border-top-color:#222}.btn.danger.saving:after{border-color:#b91c1c33;border-top-color:#b91c1c}.btn.good.saving:after{border-color:#16653433;border-top-color:#166534}@keyframes printaSpin{to{transform:rotate(360deg)}}
.printa-toast{position:fixed;left:50%;bottom:22px;transform:translateX(-50%);z-index:9999;background:#20202d;color:#fff;padding:11px 16px;border-radius:999px;font:700 14px Arial;box-shadow:0 10px 30px #0004;opacity:0;transition:.2s;pointer-events:none}.printa-toast.show{opacity:1}
</style>
<script>
(function(){
  const CACHE_KEY='PRINTA_LOCAL_CACHE_V4';
  const DAY_NAMES=['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
  let toast=document.createElement('div');toast.className='printa-toast';document.body.appendChild(toast);
  function notify(t){toast.textContent=t;toast.classList.add('show');setTimeout(()=>toast.classList.remove('show'),1800)}
  function cacheData(d){try{localStorage.setItem(CACHE_KEY,JSON.stringify({time:Date.now(),data:d}))}catch(e){}}
  function readCache(){try{const c=JSON.parse(localStorage.getItem(CACHE_KEY)||'null');return c&&c.data?c.data:null}catch(e){return null}}
  function sleep(ms){return new Promise(r=>setTimeout(r,ms))}
  function setSpin(btn,on){if(!btn)return;if(on){btn.dataset.oldText=btn.textContent;btn.classList.add('saving');btn.disabled=true}else{btn.classList.remove('saving');btn.disabled=false;if(btn.dataset.oldText)btn.textContent=btn.dataset.oldText}}
  function callServer(fn,arg){return new Promise((resolve,reject)=>google.script.run.withSuccessHandler(resolve).withFailureHandler(reject)[fn](arg))}
  async function retryCall(fn,arg){let last;for(let attempt=0;attempt<3;attempt++){try{return await callServer(fn,arg)}catch(e){last=e;if(attempt<2)await sleep(450*(attempt+1))}}throw last}
  function reopenForm(form,error){const modal=form.closest('.modal');if(modal){modal.classList.add('show');document.body.style.overflow='hidden'}const box=form.querySelector('.error');if(box){box.textContent='No se pudo guardar después de 3 intentos. Revisa la conexión y toca Guardar nuevamente.';box.style.display='block'}notify('No se pudo guardar')}
  function weekdayLabel(value){const d=dateObj(value);return d?DAY_NAMES[d.getDay()]:''}

  due=function(o){if(o.estado==='Listo')return'Orden lista';const name=weekdayLabel(o.fechaEnvio);return name?'Entregar antes del '+name.charAt(0).toUpperCase()+name.slice(1):'Fecha de entrega pendiente'};
  card=function(o){return '<div class="card order '+color(o)+'"><div class="top"><span class="badge '+(o.estado==='Listo'?'ready':'')+'">'+esc(o.estado)+'</span><span class="origin">'+esc(o.origen)+'</span></div><h3>'+esc(o.cliente)+'</h3><p><b>'+esc(o.producto)+'</b> × '+esc(o.cantidad)+'</p><p>Pagado: <b>'+money(o.pagado)+'</b></p><div class="deadline">'+due(o)+'</div>'+(o.notas?'<p>'+esc(o.notas)+'</p>':'')+'<div class="actions"><button class="btn good" onclick="status(\''+o.id+'\',\''+(o.estado==='Listo'?'Pendiente':'Listo')+'\',this)">'+(o.estado==='Listo'?'Regresar a pendiente':'LISTO')+'</button><button class="btn light" onclick="editOrderWithSpin(\''+o.id+'\',this)">Editar</button><button class="btn danger" onclick="removeOrder(\''+o.id+'\',this)">Eliminar</button></div></div>'};

  const baseRender=render;render=function(){baseRender();cacheData(data)};
  window.onload=function(){const cached=readCache();if(cached){data=cached;const fm=document.querySelector('#financeMonth');if(fm&&!fm.value)fm.value=currentMonth();render();document.body.classList.remove('loading')}else document.body.classList.add('loading');google.script.run.withSuccessHandler(function(d){data=d;cacheData(d);const fm=document.querySelector('#financeMonth');if(fm&&!fm.value)fm.value=currentMonth();render();document.body.classList.remove('loading')}).withFailureHandler(function(e){document.body.classList.remove('loading');if(!cached)showGlobalError(e)}).getAppData()};

  submitOrder=async function(e){
    e.preventDefault();const f=e.target,btn=f.querySelector('button[type="submit"],button:not([type])'),started=Date.now();
    const o=Object.fromEntries(new FormData(f).entries());o.cliente=String(f.cliente.value||'').trim();o.producto=String(f.producto.value||'').trim();
    if(!o.cliente||!o.producto){const box=document.querySelector('#orderError');box.textContent='Completa cliente y producto.';box.style.display='block';return}
    const oldData=JSON.parse(JSON.stringify(data)),existing=data.orders.findIndex(x=>x.id===o.id),tempId=o.id||('TEMP-'+Date.now());o.id=tempId;o.estado=o.estado||'Pendiente';
    const optimistic={id:tempId,fecha:o.fecha,estado:o.estado,origen:o.origen,cliente:o.cliente,producto:o.producto,cantidad:o.cantidad,notas:o.notas,pagado:o.pagado,fechaEnvio:o.fechaEnvio,foto:o.foto||'',creadoEn:o.creadoEn||new Date().toISOString()};
    setSpin(btn,true);
    setTimeout(function(){if(f.closest('.modal').classList.contains('show')){if(existing>=0)data.orders[existing]=optimistic;else data.orders.push(optimistic);render();closeModal('orderModal');notify('Guardando…')}},500);
    try{
      const old=photoUrls(o.foto),fresh=[];for(const file of pendingFiles)fresh.push(await upload(file,tempId,'orden'));o.foto=[...old,...fresh].join('|||');f.foto.value=o.foto;pendingFiles=[];renderPendingFiles();
      const d=await retryCall('saveOrder',o);await sleep(Math.max(0,500-(Date.now()-started)));data=d;render();setSpin(btn,false);notify('Guardado');
    }catch(err){await sleep(Math.max(0,500-(Date.now()-started)));data=oldData;render();setSpin(btn,false);reopenForm(f,err)}
  };

  submitExpense=async function(e){e.preventDefault();const f=e.target,btn=f.querySelector('button[type="submit"],button:not([type])'),started=Date.now(),o=Object.fromEntries(new FormData(f).entries());setSpin(btn,true);setTimeout(()=>{closeModal('expenseModal');notify('Guardando…')},500);try{if(document.querySelector('#receiptFile').files[0])o.recibo=await upload(document.querySelector('#receiptFile').files[0],'GASTO','recibo');const d=await retryCall('saveExpense',o);await sleep(Math.max(0,500-(Date.now()-started)));data=d;render();setSpin(btn,false);notify('Guardado')}catch(err){setSpin(btn,false);reopenForm(f,err)}};
  submitProduct=async function(e){e.preventDefault();const f=e.target,btn=f.querySelector('button[type="submit"],button:not([type])'),started=Date.now(),o=Object.fromEntries(new FormData(f).entries());setSpin(btn,true);setTimeout(()=>{closeModal('productModal');notify('Guardando…')},500);try{if(document.querySelector('#productImage').files[0])o.foto=await upload(document.querySelector('#productImage').files[0],o.id||'PRODUCTO','producto');const d=await retryCall('saveProduct',o);await sleep(Math.max(0,500-(Date.now()-started)));data=d;render();setSpin(btn,false);notify('Guardado')}catch(err){setSpin(btn,false);reopenForm(f,err)}};

  status=async function(id,s,btn){const index=data.orders.findIndex(x=>x.id===id);if(index<0)return;const old=data.orders[index].estado;setSpin(btn,true);data.orders[index].estado=s;render();try{const d=await Promise.all([retryCall('setOrderStatus',{id:id,status:s}).catch(()=>retryCallStatus(id,s)),sleep(500)]).then(x=>x[0]);data=d;render();setSpin(btn,false)}catch(e){data.orders[index].estado=old;render();setSpin(btn,false);notify('No se pudo cambiar el estado')}};
  async function retryCallStatus(id,s){let last;for(let i=0;i<3;i++){try{return await new Promise((res,rej)=>google.script.run.withSuccessHandler(res).withFailureHandler(rej).setOrderStatus(id,s))}catch(e){last=e;if(i<2)await sleep(450*(i+1))}}throw last}
  editOrderWithSpin=async function(id,btn){setSpin(btn,true);await sleep(500);setSpin(btn,false);editOrder(id)};
  removeOrder=async function(id,btn){if(!confirm('¿Eliminar esta orden?'))return;const index=data.orders.findIndex(x=>x.id===id);if(index<0)return;const removed=data.orders[index];setSpin(btn,true);data.orders.splice(index,1);render();try{let d,last;for(let i=0;i<3;i++){try{d=await new Promise((res,rej)=>google.script.run.withSuccessHandler(res).withFailureHandler(rej).deleteOrder(id));break}catch(e){last=e;if(i<2)await sleep(450*(i+1));else throw last}}await sleep(500);data=d;render();setSpin(btn,false)}catch(e){data.orders.splice(index,0,removed);render();setSpin(btn,false);notify('No se pudo eliminar')}};
})();
</script>`}
