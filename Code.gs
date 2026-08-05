const SPREADSHEET_ID='1eOWiWZh0zTnVLnv53bTDXvlFJJG5ZbvR6Upr7_9uLb8';
const ORDERS_SHEET='ORDENES',EXPENSES_SHEET='GASTOS',PRODUCTS_SHEET='PRODUCTOS',CONFIG_SHEET='CONFIGURACION';
const LOGO_FILE_ID='1KNCCBhFm4vD92Jpi5rdKKgC6xxEWk1ea';

function doGet(){
  const rendered=HtmlService.createTemplateFromFile('App').evaluate().getContent();
  const multiPhotoCss=`<style>.photo-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px;margin-bottom:12px}.photo-grid a{display:block;position:relative;overflow:hidden;border-radius:12px;border:1px solid var(--line);background:#f3f3f7}.photo-grid img{width:100%;height:105px;object-fit:cover;display:block}.photo-count{position:absolute;right:7px;bottom:7px;background:#171724d9;color:#fff;border-radius:999px;padding:4px 7px;font-size:11px;font-weight:900}@media(max-width:600px){.photo-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.photo-grid img{height:125px}}</style>`;
  const multiPhotoScript=`<script>
    function orderPhotoUrls(value){return String(value||'').split(',').map(function(x){return x.trim()}).filter(Boolean)}
    function orderPhotoGallery(value){var urls=orderPhotoUrls(value);if(!urls.length)return '';return '<div class="photo-grid">'+urls.map(function(url,index){var counter=index===0&&urls.length>1?'<span class="photo-count">'+urls.length+' fotos</span>':'';return '<a href="'+url+'" target="_blank" rel="noopener"><img src="'+url+'" alt="Foto de la orden">'+counter+'</a>'}).join('')+'</div>'}
    card=function(o){return '<div class="card order '+color(o)+'">'+orderPhotoGallery(o.foto)+'<div class="top"><span class="badge '+(o.estado==='Listo'?'ready':'')+'">'+esc(o.estado)+'</span><span class="origin">'+esc(o.origen)+'</span></div><h3>'+esc(o.cliente)+'</h3><p><b>'+esc(o.producto)+'</b> × '+esc(o.cantidad)+'</p><p>Pagado: <b>'+money(o.pagado)+'</b></p><div class="deadline">'+due(o)+' · '+esc(o.fechaEnvio||'')+'</div>'+(o.notas?'<p>'+esc(o.notas)+'</p>':'')+'<div class="actions"><button class="btn good" onclick="status(\''+o.id+'\',\''+(o.estado==='Listo'?'Pendiente':'Listo')+'\')">'+(o.estado==='Listo'?'Regresar a pendiente':'Marcar listo')+'</button><button class="btn light" onclick="editOrder(\''+o.id+'\')">Editar</button><button class="btn danger" onclick="removeOrder(\''+o.id+'\')">Eliminar</button></div></div>'}
    async function uploadOrderPhotos(files,ownerId,previous){var urls=orderPhotoUrls(previous);if(!files||!files.length)return urls.join(',');for(var i=0;i<files.length;i++){urls.push(await uploadOne(files[i],ownerId,'orden',''))}return urls.filter(Boolean).join(',')}
    submitOrder=async function(e){e.preventDefault();var f=e.target;busy(f,true);try{var o=Object.fromEntries(new FormData(f).entries());o.foto=await uploadOrderPhotos(document.querySelector('#orderFile').files,o.id||'NUEVA',o.foto);run('saveOrder',o,function(d){data=d;render();closeModal('orderModal');busy(f,false)},function(err){fail('orderError',err,f)})}catch(err){fail('orderError',err,f)}};
    document.addEventListener('DOMContentLoaded',function(){var input=document.querySelector('#orderFile');if(input){input.multiple=true;input.removeAttribute('capture');var label=input.previousElementSibling;if(label)label.textContent='Fotos de la orden o referencias';var hint=document.createElement('div');hint.className='hint';hint.textContent='Puedes seleccionar varias fotos a la vez. Al editar, las fotos nuevas se agregarán a las anteriores.';input.insertAdjacentElement('afterend',hint)}});
  </script>`;
  const html=rendered.replace('</head>',multiPhotoCss+'</head>').replace('</body>',multiPhotoScript+'</body>');
  return HtmlService.createHtmlOutput(html).setTitle('PRINTA').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getLogoDataUrl(){try{const b=DriveApp.getFileById(LOGO_FILE_ID).getBlob();return`data:${b.getContentType()};base64,${Utilities.base64Encode(b.getBytes())}`}catch(e){return''}}

function getAppData(){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const orders=readRows_(ss.getSheetByName(ORDERS_SHEET));
  const expenses=readRows_(ss.getSheetByName(EXPENSES_SHEET));
  const products=readRows_(ss.getSheetByName(PRODUCTS_SHEET));
  const settings=getSettings_(ss);
  return{orders,expenses,products,settings,summary:buildSummary_(orders,expenses)};
}
function calculateShippingDate(orderDate,origin){const ss=SpreadsheetApp.openById(SPREADSHEET_ID);const settings=getSettings_(ss);return addBusinessDays_(parseDate_(orderDate)||new Date(),getHandlingDays_(origin,settings))}
function saveOrder(o){if(!o||!clean_(o.cliente)||!clean_(o.producto))throw new Error('Completa cliente y producto.');const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=ss.getSheetByName(ORDERS_SHEET),now=new Date(),id=o.id||nextId_(sh,'PC'),orderDate=o.fecha||formatDate_(now),settings=getSettings_(ss),shippingDate=addBusinessDays_(parseDate_(orderDate)||now,getHandlingDays_(o.origen,settings));upsertById_(sh,id,[id,orderDate,o.estado||'Pendiente',o.origen||'Otro',clean_(o.cliente),clean_(o.producto),number_(o.cantidad,1),clean_(o.notas),money_(o.pagado),shippingDate,clean_(o.foto),o.creadoEn||now,now]);return getAppData()}
function setOrderStatus(id,status){if(!['Pendiente','Listo'].includes(status))throw new Error('Estado inválido.');const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDERS_SHEET),r=findRowById_(sh,id);if(!r)throw new Error('Orden no encontrada.');sh.getRange(r,3).setValue(status);sh.getRange(r,13).setValue(new Date());return getAppData()}
function deleteOrder(id){deleteById_(ORDERS_SHEET,id);return getAppData()}
function saveExpense(e){if(!e||money_(e.cantidad)<=0)throw new Error('Escribe una cantidad válida.');const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET),now=new Date(),id=e.id||nextId_(sh,'GA');upsertById_(sh,id,[id,e.fecha||formatDate_(now),e.categoria||'Otro',clean_(e.lugar),money_(e.cantidad),clean_(e.nota),clean_(e.recibo),e.creadoEn||now]);return getAppData()}
function deleteExpense(id){deleteById_(EXPENSES_SHEET,id);return getAppData()}
function saveProduct(p){if(!p||!clean_(p.nombre))throw new Error('Escribe el nombre del producto.');const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PRODUCTS_SHEET),now=new Date(),id=p.id||nextId_(sh,'PR');upsertById_(sh,id,[id,clean_(p.nombre),clean_(p.foto),String(p.disponible)!=='false',p.creadoEn||now,now]);return getAppData()}
function deleteProduct(id){deleteById_(PRODUCTS_SHEET,id);return getAppData()}
function uploadFile(fileData,fileName,mimeType,ownerId,category){if(!fileData)return'';const folder=getUploadsFolder_(),bytes=Utilities.base64Decode(String(fileData).split(',').pop()),safe=`${ownerId||'GENERAL'}_${category||'archivo'}_${Date.now()}_${fileName||'archivo'}`,file=folder.createFile(Utilities.newBlob(bytes,mimeType||'application/octet-stream',safe));file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return file.getUrl()}
function getUploadsFolder_(){const ss=SpreadsheetApp.openById(SPREADSHEET_ID),c=ss.getSheetByName(CONFIG_SHEET),v=c.getRange(1,1,Math.max(c.getLastRow(),1),2).getValues();let id='',row=0;v.forEach((x,i)=>{if(x[0]==='CARPETA_FOTOS_ID'){id=x[1];row=i+1}});if(id){try{return DriveApp.getFolderById(id)}catch(e){}}const f=DriveApp.createFolder('PRINTA - Fotos y recibos');if(row)c.getRange(row,2).setValue(f.getId());return f}
function getSettings_(ss){const sh=ss.getSheetByName(CONFIG_SHEET),rows=sh.getRange(1,1,Math.max(sh.getLastRow(),1),2).getValues(),out={tiktokDays:2,printaDays:3};rows.forEach(r=>{if(r[0]==='TIKTOK_DIAS_ENVIO')out.tiktokDays=number_(r[1],2);if(r[0]==='PRINTA_DIAS_ENVIO')out.printaDays=number_(r[1],3)});return out}
function getHandlingDays_(origin,settings){return origin==='TikTok'?settings.tiktokDays:origin==='Printa Crea'?settings.printaDays:settings.printaDays}
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