const SPREADSHEET_ID='1eOWiWZh0zTnVLnv53bTDXvlFJJG5ZbvR6Upr7_9uLb8';
const ORDERS_SHEET='ORDENES';
const EXPENSES_SHEET='GASTOS';
const PRODUCTS_SHEET='PRODUCTOS';
const CONFIG_SHEET='CONFIGURACION';
const LOGO_FILE_ID='1KNCCBhFm4vD92Jpi5rdKKgC6xxEWk1ea';
const APP_CACHE_KEY='PRINTA_APP_DATA_V10';

const ORDER_HEADERS=['ID','FECHA','ESTADO','ORIGEN','CLIENTE','PRODUCTO','CANTIDAD','NOTAS','PAGADO','FECHA_ENVIO','FOTO','CREADO_EN','ACTUALIZADO_EN','COSTO_PRODUCTO','COSTO_TINTA','COSTO_TOTAL','GANANCIA','INVENTARIO_DESCONTADO'];
const EXPENSE_HEADERS=['ID','FECHA','CATEGORIA','LUGAR','CANTIDAD','NOTA','RECIBO','CREADO_EN','PRODUCTO','CANTIDAD_UNIDADES','COSTO_UNITARIO'];
const PRODUCT_HEADERS=['ID','NOMBRE','FOTO','DISPONIBLE','CREADO_EN','ACTUALIZADO_EN','INVENTARIO','COSTO_TOTAL_INVENTARIO','COSTO_UNITARIO','METODO_TINTA','COSTO_TINTA'];

function doGet(){
  let html=HtmlService.createTemplateFromFile('App').evaluate().getContent();
  const icon=getLogoDataUrl();
  html=html.replace('</head>',`<link rel="icon" type="image/png" href="${icon}"><style>
    .order>.gallery{display:none!important}.costline{display:flex;justify-content:space-between;gap:10px;margin-top:7px;padding-top:7px;border-top:1px dashed #ddd;font-size:13px}.costline .gain{color:#15803d;font-weight:900}.costline .loss{color:#dc2626;font-weight:900}
    .stock-pill{display:inline-block;padding:5px 9px;border-radius:999px;background:#eef2ff;color:#5b21b6;font-size:12px;font-weight:900;margin:4px 4px 0 0}.stock-pill.low{background:#fee2e2;color:#b91c1c}
    .btn.saving{position:relative;color:transparent!important;pointer-events:none}.btn.saving:after{content:'';position:absolute;left:50%;top:50%;width:18px;height:18px;margin:-11px 0 0 -11px;border:3px solid #ffffff66;border-top-color:#fff;border-radius:50%;animation:printaSpin .65s linear infinite}
    @keyframes printaSpin{to{transform:rotate(360deg)}}.printa-toast{position:fixed;left:50%;bottom:22px;transform:translateX(-50%);z-index:9999;background:#20202d;color:#fff;padding:11px 16px;border-radius:999px;font:700 14px Arial;opacity:0;transition:.2s}.printa-toast.show{opacity:1}
  </style></head>`);

  html=html.replace(/function due\(o\)\{.*?\}function photoUrls/,`function due(o){if(o.estado==='Listo')return'Orden lista';const d=dateObj(o.fechaEnvio);if(!d)return'Fecha de entrega pendiente';const nombres=['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];return'Entregar antes del '+nombres[d.getDay()]}function photoUrls`);

  html=html.replace(/function card\(o\)\{.*?\}function orders/,`function card(o){const g=num(o.ganancia),c=num(o.costoTotal);return\`<div class="card order \${color(o)}"><div class="top"><span class="badge \${o.estado==='Listo'?'ready':''}">\${esc(o.estado)}</span><span class="origin">\${esc(o.origen)}</span></div><h3>\${esc(o.cliente)}</h3><p><b>\${esc(o.producto)}</b> × \${esc(o.cantidad)}</p><p>Pagado: <b>\${money(o.pagado)}</b></p><div class="costline"><span>Costo: <b>\${money(c)}</b></span><span class="\${g>=0?'gain':'loss'}">Ganancia: \${money(g)}</span></div><div class="deadline">\${due(o)}</div>\${o.notas?\`<p>\${esc(o.notas)}</p>\`:''}<div class="actions"><button class="btn good" onclick="statusFast('\${o.id}','\${o.estado==='Listo'?'Pendiente':'Listo'}',this)">\${o.estado==='Listo'?'Regresar a pendiente':'✅ Listo'}</button><button class="btn light" onclick="editFast('\${o.id}',this)">Editar</button><button class="btn danger" onclick="removeFast('\${o.id}',this)">Eliminar</button></div></div>\`}function orders`);

  html=html.replace(/async function submitOrder\(e\)\{.*?\}async function submitExpense/,clientOrderScript_()+'async function submitExpense');
  html=html.replace(/function status\(id,s\)\{.*?\}function removeOrder\(id\)\{.*?\}function removeExpense/,clientFastButtons_()+'function removeExpense');
  html=html.replace('</body>',inventoryClientScript_()+'</body>');
  return HtmlService.createHtmlOutput(html).setTitle('PRINTA').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function clientOrderScript_(){return `async function submitOrder(e){
 e.preventDefault();const f=e.target,b=f.querySelector('button[type="submit"],button:not([type])'),start=Date.now(),o=Object.fromEntries(new FormData(f).entries());
 o.cliente=String(f.elements.cliente.value||'').trim();o.producto=String(f.elements.producto.value||'').trim();
 if(!o.cliente||!o.producto){const x=$('#orderError');x.textContent='Completa cliente y producto.';x.style.display='block';return}
 const isEdit=String($('#orderTitle').textContent||'').toLowerCase().includes('editar')&&String(o.id||'').trim();o._mode=isEdit?'edit':'create';o.clientKey=isEdit?'':('CLIENT-'+Date.now()+'-'+Math.random().toString(36).slice(2,12));if(!isEdit)o.id='';
 setButtonSpin(b,true);const backup=JSON.parse(JSON.stringify(data)),visualId=isEdit?o.id:('PC-'+o.clientKey);o.estado=o.estado||'Pendiente';
 setTimeout(()=>{const item={...o,id:visualId,costoProducto:0,costoTinta:0,costoTotal:0,ganancia:num(o.pagado)};if(isEdit){const i=data.orders.findIndex(x=>x.id===visualId);if(i>=0)data.orders[i]=item}else if(!data.orders.some(x=>x.id===visualId))data.orders.push(item);render();closeModal('orderModal');toast('Guardando…')},500);
 try{const old=photoUrls(o.foto),fresh=[];for(const file of pendingFiles)fresh.push(await upload(file,visualId,'orden'));o.foto=[...old,...fresh].join('|||');data=await retryServer('saveOrder',[o]);await waitMinimum(start,500);pendingFiles=[];render();setButtonSpin(b,false);toast('Guardado')}
 catch(err){data=backup;render();setButtonSpin(b,false);showModal('orderModal');const x=$('#orderError');x.textContent=err&&err.message?err.message:'No se pudo guardar. Toca Guardar nuevamente.';x.style.display='block'}
}`}

function clientFastButtons_(){return `function setButtonSpin(b,on){if(!b)return;if(on){b.dataset.oldText=b.textContent;b.classList.add('saving');b.disabled=true}else{b.classList.remove('saving');b.disabled=false;if(b.dataset.oldText)b.textContent=b.dataset.oldText}}
function waitMinimum(start,ms){return new Promise(r=>setTimeout(r,Math.max(0,ms-(Date.now()-start))))}function toast(t){let x=document.querySelector('.printa-toast');if(!x){x=document.createElement('div');x.className='printa-toast';document.body.appendChild(x)}x.textContent=t;x.classList.add('show');setTimeout(()=>x.classList.remove('show'),1600)}
function serverCall(fn,args){return new Promise((ok,no)=>{const r=google.script.run.withSuccessHandler(ok).withFailureHandler(no);r[fn].apply(r,args)})}async function retryServer(fn,args){let err;for(let i=0;i<3;i++){try{return await serverCall(fn,args)}catch(e){err=e;if(i<2)await new Promise(r=>setTimeout(r,450*(i+1)))}}throw err}
async function statusFast(id,s,b){const start=Date.now(),backup=JSON.parse(JSON.stringify(data));setButtonSpin(b,true);const i=data.orders.findIndex(x=>x.id===id);if(i>=0)data.orders[i].estado=s;render();try{data=await retryServer('setOrderStatus',[id,s]);await waitMinimum(start,500);render()}catch(e){data=backup;render();toast(e&&e.message?e.message:'No se pudo cambiar')}setButtonSpin(b,false)}
async function editFast(id,b){const start=Date.now();setButtonSpin(b,true);await waitMinimum(start,500);setButtonSpin(b,false);editOrder(id)}
async function removeFast(id,b){if(!confirm('¿Eliminar esta orden?'))return;const start=Date.now(),backup=JSON.parse(JSON.stringify(data));setButtonSpin(b,true);data.orders=data.orders.filter(x=>x.id!==id);render();try{data=await retryServer('deleteOrder',[id]);await waitMinimum(start,500);render()}catch(e){data=backup;render();toast(e&&e.message?e.message:'No se pudo eliminar')}setButtonSpin(b,false)}
`}

function inventoryClientScript_(){return `<script>
(function(){
 const after=(el,html)=>el&&el.insertAdjacentHTML('beforebegin',html);
 function install(){
  const pf=document.querySelector('#productForm .actions');
  if(pf&&!document.querySelector('#inventoryFields'))after(pf,'<div id="inventoryFields" class="full form" style="padding:14px;border:1px solid #e7e7ef;border-radius:14px;background:#fafaff"><div><label>Inventario actual</label><input name="inventario" type="number" min="0" step="1" value="0"></div><div><label>Costo total del inventario</label><input name="costoTotalInventario" type="number" min="0" step="0.01" value="0" oninput="updateUnitPreview()"></div><div><label>Método de tinta</label><select name="metodoTinta"><option>Sawgrass</option><option>Eufy</option><option>Sin tinta</option></select></div><div><label>Costo de tinta por unidad</label><input name="costoTinta" type="number" min="0" step="0.01" value="0"></div><div class="full hint" id="unitPreview">Costo unitario del producto: $0.00</div></div>');
  const ef=document.querySelector('#expenseForm .actions');
  if(ef&&!document.querySelector('#purchaseFields'))after(ef,'<div id="purchaseFields" class="full" style="display:none;padding:14px;border:1px solid #e7e7ef;border-radius:14px;background:#fafaff"><h3 style="margin-bottom:10px">Entrada de inventario</h3><div class="form"><div><label>Producto</label><select name="producto" id="expenseProduct"></select></div><div><label>Cantidad de unidades compradas</label><input name="cantidadUnidades" type="number" min="1" step="1"></div></div><div class="hint">El total pagado se divide automáticamente entre las unidades y actualiza el costo promedio.</div></div>');
  const cat=document.querySelector('#expenseForm [name="categoria"]');if(cat){cat.onchange=togglePurchaseFields}
 }
 window.updateUnitPreview=function(){const f=document.querySelector('#productForm'),q=num(f?.elements.inventario?.value),t=num(f?.elements.costoTotalInventario?.value),u=q>0?t/q:0;const x=document.querySelector('#unitPreview');if(x)x.textContent='Costo unitario del producto: '+money(u)}
 window.togglePurchaseFields=function(){const f=document.querySelector('#expenseForm'),box=document.querySelector('#purchaseFields'),show=f&&f.elements.categoria.value==='Productos';if(box)box.style.display=show?'block':'none';if(show){const s=document.querySelector('#expenseProduct');s.innerHTML=data.products.map(p=>'<option>'+esc(p.nombre)+'</option>').join('')}}
 const oldOpenProduct=window.openProduct;window.openProduct=function(){oldOpenProduct();install();const f=document.querySelector('#productForm');f.elements.inventario.value=0;f.elements.costoTotalInventario.value=0;f.elements.metodoTinta.value='Sawgrass';f.elements.costoTinta.value=0;updateUnitPreview()}
 const oldEditProduct=window.editProduct;window.editProduct=function(id){oldEditProduct(id);install();const p=data.products.find(x=>x.id===id),f=document.querySelector('#productForm');if(!p)return;f.elements.inventario.value=num(p.inventario);f.elements.costoTotalInventario.value=num(p.costoTotalInventario);f.elements.metodoTinta.value=p.metodoTinta||'Sawgrass';f.elements.costoTinta.value=num(p.costoTinta);updateUnitPreview()}
 window.submitProduct=async function(e){e.preventDefault();const f=e.target,o=Object.fromEntries(new FormData(f).entries()),file=document.querySelector('#productImage').files[0];f.classList.add('loading');try{if(file)o.foto=await upload(file,o.id||('PR-'+Date.now()),'producto');google.script.run.withSuccessHandler(d=>{data=d;render();closeModal('productModal');f.classList.remove('loading')}).withFailureHandler(err=>{f.classList.remove('loading');const x=document.querySelector('#productError');x.textContent=err.message||err;x.style.display='block'}).saveProduct(o)}catch(err){f.classList.remove('loading')}}
 const oldOpenExpense=window.openExpense;window.openExpense=function(){oldOpenExpense();install();togglePurchaseFields()}
 window.submitExpense=async function(e){e.preventDefault();const f=e.target,o=Object.fromEntries(new FormData(f).entries()),file=document.querySelector('#receiptFile').files[0];if(o.categoria==='Productos'&&(!o.producto||num(o.cantidadUnidades)<=0)){const x=document.querySelector('#expenseError');x.textContent='Selecciona el producto y la cantidad comprada.';x.style.display='block';return}f.classList.add('loading');try{if(file)o.recibo=await upload(file,'GASTO-'+Date.now(),'recibo');google.script.run.withSuccessHandler(d=>{data=d;render();closeModal('expenseModal');f.classList.remove('loading')}).withFailureHandler(err=>{f.classList.remove('loading');const x=document.querySelector('#expenseError');x.textContent=err.message||err;x.style.display='block'}).saveExpense(o)}catch(err){f.classList.remove('loading')}}
 window.products=function(){const el=document.querySelector('#productsList');if(!el)return;el.innerHTML=data.products.length?data.products.map(p=>{const inv=num(p.inventario),unit=num(p.costoUnitario),ink=num(p.costoTinta);return '<div class="card"><div class="product-row">'+(p.foto?'<img class="product-img" src="'+esc(p.foto)+'">':'')+'<div><h3>'+esc(p.nombre)+'</h3><span class="stock-pill '+(inv<=2?'low':'')+'">Inventario: '+inv+'</span><span class="stock-pill">Unidad: '+money(unit)+'</span><span class="stock-pill">'+esc(p.metodoTinta||'Sin tinta')+': '+money(ink)+'</span><p><b>Costo por orden:</b> '+money(unit+ink)+'</p><div class="actions"><button class="btn light" onclick="editProduct(\''+p.id+'\')">Editar</button><button class="btn danger" onclick="removeProduct(\''+p.id+'\')">Eliminar</button></div></div></div></div>'}).join(''):'<div class="empty">No hay productos.</div>'}
 window.monthData=function(m){const o=data.orders.filter(x=>String(x.fecha||'').startsWith(m)),e=data.expenses.filter(x=>String(x.fecha||'').startsWith(m)),ventas=o.reduce((s,x)=>s+num(x.pagado),0),costos=o.reduce((s,x)=>s+num(x.costoTotal),0),operativos=e.filter(x=>!['Productos','Tinta'].includes(x.categoria)).reduce((s,x)=>s+num(x.cantidad),0);return{o,e,ventas,gastos:operativos,costos,resultado:ventas-costos-operativos,tiktok:o.filter(x=>x.origen==='TikTok').reduce((s,x)=>s+num(x.pagado),0),printa:o.filter(x=>x.origen==='Printa Crea').reduce((s,x)=>s+num(x.pagado),0)}}
 window.metrics=function(){const m=monthData(currentMonth()),s=data.summary||{};document.querySelector('#metrics').innerHTML='<div class="card metric"><small>Pendientes</small><strong>'+s.pendientes+'</strong></div><div class="card metric"><small>Listas</small><strong>'+s.listos+'</strong></div><div class="card metric sales"><small>Ventas</small><strong>'+money(m.ventas)+'</strong></div><div class="card metric expense"><small>Costos + gastos</small><strong>'+money(m.costos+m.gastos)+'</strong></div><div class="card metric '+(m.resultado>=0?'profit':'loss')+'"><small>'+(m.resultado>=0?'Ganancia':'Pérdida')+'</small><strong>'+money(m.resultado)+'</strong></div>'}
 window.renderFinance=function(){const m=monthData(document.querySelector('#financeMonth').value||currentMonth()),r=m.resultado;document.querySelector('#financeMetrics').innerHTML='<div class="card metric sales"><small>Ventas</small><strong>'+money(m.ventas)+'</strong></div><div class="card metric expense"><small>Costo de producción</small><strong>'+money(m.costos)+'</strong></div><div class="card metric expense"><small>Otros gastos</small><strong>'+money(m.gastos)+'</strong></div><div class="card metric '+(r>=0?'profit':'loss')+'"><small>'+(r>=0?'Ganancia real':'Pérdida')+'</small><strong>'+money(r)+'</strong></div><div class="card metric"><small>TikTok / Printa</small><strong>'+money(m.tiktok)+' / '+money(m.printa)+'</strong></div>';const e=m.e;document.querySelector('#expensesList').innerHTML=e.length?e.map(x=>'<div class="card" style="margin-bottom:10px"><b>'+esc(x.categoria)+'</b> · '+money(x.cantidad)+(x.producto?' · '+esc(x.producto)+' × '+esc(x.cantidadUnidades):'')+'<br><small>'+esc(x.fecha)+' '+esc(x.nota||'')+'</small></div>').join(''):'<div class="empty">No hay gastos en este mes.</div>'}
 setTimeout(install,0);
})();
</script>`}

function getLogoDataUrl(){try{const b=DriveApp.getFileById(LOGO_FILE_ID).getBlob();return`data:${b.getContentType()};base64,${Utilities.base64Encode(b.getBytes())}`}catch(e){return''}}

function getAppData(){
 const c=CacheService.getScriptCache(),v=c.get(APP_CACHE_KEY);if(v){try{return JSON.parse(v)}catch(e){}}
 const ss=SpreadsheetApp.openById(SPREADSHEET_ID);ensureSchema_(ss);
 const orders=readRows_(ss.getSheetByName(ORDERS_SHEET)),expenses=readRows_(ss.getSheetByName(EXPENSES_SHEET)),products=readRows_(ss.getSheetByName(PRODUCTS_SHEET)),settings=getSettings_(ss);
 const r={orders,expenses,products,settings,summary:buildSummary_(orders,expenses)};try{c.put(APP_CACHE_KEY,JSON.stringify(r),120)}catch(e){}return r;
}
function invalidateCache_(){try{CacheService.getScriptCache().remove(APP_CACHE_KEY)}catch(e){}}
function calculateShippingDate(orderDate,origin){const ss=SpreadsheetApp.openById(SPREADSHEET_ID);ensureSchema_(ss);return addBusinessDays_(parseDate_(orderDate)||new Date(),getHandlingDays_(origin,getSettings_(ss)))}

function saveOrder(o){
 if(!o||!clean_(o.cliente)||!clean_(o.producto))throw new Error('Completa cliente y producto.');
 const mode=clean_(o._mode).toLowerCase();if(!['create','edit'].includes(mode))throw new Error('Operación inválida.');
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);ensureSchema_(ss);const sh=ss.getSheetByName(ORDERS_SHEET),now=new Date(),fecha=o.fecha||formatDate_(now),envio=addBusinessDays_(parseDate_(fecha)||now,getHandlingDays_(o.origen,getSettings_(ss)));
  let id,rowNum=0,created=now,old=null;
  if(mode==='create'){const key=clean_(o.clientKey).replace(/[^A-Za-z0-9_-]/g,'');if(!key)throw new Error('Toca Guardar nuevamente.');id='PC-'+key;const m=findRowsById_(sh,id);if(m.length===1){invalidateCache_();return getAppData()}if(m.length>1)throw new Error('ID duplicado. No se guardó.')}
  else{id=clean_(o.id);const m=findRowsById_(sh,id);if(m.length!==1)throw new Error(m.length?'ID duplicado. No se editó.':'La orden ya no existe.');rowNum=m[0];old=rowObject_(sh,rowNum);created=sh.getRange(rowNum,12).getValue()||now}
  const cost=orderCost_(ss,o.producto,number_(o.cantidad,1),money_(o.pagado));
  if(old&&String(old.inventarioDescontado).toLowerCase()==='true'){
    restoreInventory_(ss,old.producto,number_(old.cantidad,1));
    consumeInventory_(ss,o.producto,number_(o.cantidad,1));
    cost.inventoryUsed=true;
  }
  const row=[id,fecha,o.estado||'Pendiente',o.origen||'Otro',clean_(o.cliente),clean_(o.producto),number_(o.cantidad,1),clean_(o.notas),money_(o.pagado),envio,clean_(o.foto),created,now,cost.product,cost.ink,cost.total,cost.gain,cost.inventoryUsed];
  rowNum?sh.getRange(rowNum,1,1,row.length).setValues([row]):sh.appendRow(row);invalidateCache_();
 }finally{lock.releaseLock()}return getAppData();
}

function setOrderStatus(id,status){
 if(!['Pendiente','Listo'].includes(status))throw new Error('Estado inválido.');
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);ensureSchema_(ss);const sh=ss.getSheetByName(ORDERS_SHEET),m=findRowsById_(sh,id);if(m.length!==1)throw new Error(m.length?'ID duplicado.':'Orden no encontrada.');
  const r=m[0],o=rowObject_(sh,r),used=String(o.inventarioDescontado).toLowerCase()==='true';
  if(status==='Listo'&&!used){consumeInventory_(ss,o.producto,number_(o.cantidad,1));sh.getRange(r,18).setValue(true)}
  if(status==='Pendiente'&&used){restoreInventory_(ss,o.producto,number_(o.cantidad,1));sh.getRange(r,18).setValue(false)}
  sh.getRange(r,3).setValue(status);sh.getRange(r,13).setValue(new Date());invalidateCache_();
 }finally{lock.releaseLock()}return getAppData();
}

function deleteOrder(id){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{const ss=SpreadsheetApp.openById(SPREADSHEET_ID);ensureSchema_(ss);const sh=ss.getSheetByName(ORDERS_SHEET),m=findRowsById_(sh,id);if(m.length!==1)throw new Error(m.length?'ID duplicado.':'Orden no encontrada.');const o=rowObject_(sh,m[0]);if(String(o.inventarioDescontado).toLowerCase()==='true')restoreInventory_(ss,o.producto,number_(o.cantidad,1));sh.deleteRow(m[0]);invalidateCache_()}finally{lock.releaseLock()}return getAppData();
}

function saveExpense(e){
 if(!e||money_(e.cantidad)<=0)throw new Error('Escribe una cantidad válida.');
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);ensureSchema_(ss);const sh=ss.getSheetByName(EXPENSES_SHEET),now=new Date(),id=e.id||nextId_('GA'),cat=e.categoria||'Otro',total=money_(e.cantidad),units=number_(e.cantidadUnidades,0),product=clean_(e.producto);
  let unit=0;if(cat==='Productos'){if(!product||units<=0)throw new Error('Selecciona producto y cantidad.');unit=addInventoryPurchase_(ss,product,units,total)}
  upsertById_(sh,id,[id,e.fecha||formatDate_(now),cat,clean_(e.lugar),total,clean_(e.nota),clean_(e.recibo),e.creadoEn||now,product,units,unit]);invalidateCache_();
 }finally{lock.releaseLock()}return getAppData();
}
function deleteExpense(id){deleteById_(EXPENSES_SHEET,id);invalidateCache_();return getAppData()}

function saveProduct(p){
 if(!p||!clean_(p.nombre))throw new Error('Escribe el nombre del producto.');
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);ensureSchema_(ss);const sh=ss.getSheetByName(PRODUCTS_SHEET),now=new Date(),id=p.id||nextId_('PR'),m=findRowsById_(sh,id),old=m.length===1?rowObject_(sh,m[0]):{},inv=Math.max(0,number_(p.inventario,number_(old.inventario,0))),total=Math.max(0,money_(p.costoTotalInventario)),unit=inv>0?total/inv:0;
  const row=[id,clean_(p.nombre),clean_(p.foto||old.foto),String(p.disponible)!=='false',p.creadoEn||old.creadoEn||now,now,inv,total,unit,clean_(p.metodoTinta)||'Sawgrass',Math.max(0,money_(p.costoTinta))];upsertById_(sh,id,row);invalidateCache_();
 }finally{lock.releaseLock()}return getAppData();
}
function deleteProduct(id){deleteById_(PRODUCTS_SHEET,id);invalidateCache_();return getAppData()}

function uploadFile(fileData,fileName,mimeType,ownerId,category){if(!fileData)return'';const bytes=Utilities.base64Decode(String(fileData).split(',').pop()),safe=`${ownerId||'GENERAL'}_${category||'archivo'}_${Date.now()}_${fileName||'archivo'}`,f=getUploadsFolder_().createFile(Utilities.newBlob(bytes,mimeType||'application/octet-stream',safe));f.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return f.getUrl()}
function getUploadsFolder_(){const ss=SpreadsheetApp.openById(SPREADSHEET_ID);ensureSchema_(ss);const sh=ss.getSheetByName(CONFIG_SHEET),v=sh.getRange(1,1,Math.max(sh.getLastRow(),1),2).getValues();let id='',row=0;v.forEach((x,i)=>{if(x[0]==='CARPETA_FOTOS_ID'){id=x[1];row=i+1}});if(id){try{return DriveApp.getFolderById(id)}catch(e){}}const f=DriveApp.createFolder('PRINTA - Fotos y recibos');if(row)sh.getRange(row,2).setValue(f.getId());return f}

function ensureSchema_(ss){ensureSheet_(ss,ORDERS_SHEET,ORDER_HEADERS);ensureSheet_(ss,EXPENSES_SHEET,EXPENSE_HEADERS);ensureSheet_(ss,PRODUCTS_SHEET,PRODUCT_HEADERS);ensureSheet_(ss,CONFIG_SHEET,['CLAVE','VALOR'])}
function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);const last=Math.max(sh.getLastColumn(),1),current=sh.getRange(1,1,1,last).getDisplayValues()[0];headers.forEach(h=>{if(!current.includes(h)){sh.getRange(1,sh.getLastColumn()+1).setValue(h);current.push(h)}});return sh}
function rowObject_(sh,row){const h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],v=sh.getRange(row,1,1,sh.getLastColumn()).getDisplayValues()[0],o={};h.forEach((x,i)=>o[toKey_(x)]=v[i]);return o}
function productRow_(ss,name){const sh=ss.getSheetByName(PRODUCTS_SHEET),rows=readRows_(sh),p=rows.find(x=>clean_(x.nombre).toLowerCase()===clean_(name).toLowerCase());if(!p)throw new Error('El producto no está configurado en inventario: '+name);const m=findRowsById_(sh,p.id);return{sh,row:m[0],p}}
function orderCost_(ss,name,qty,paid){const x=productRow_(ss,name).p,product=money_(x.costoUnitario)*qty,ink=(clean_(x.metodoTinta)==='Sin tinta'?0:money_(x.costoTinta))*qty,total=product+ink;return{product,ink,total,gain:paid-total,inventoryUsed:false}}
function consumeInventory_(ss,name,qty){const x=productRow_(ss,name),inv=number_(x.p.inventario,0);if(inv<qty)throw new Error(`Inventario insuficiente de ${name}. Disponible: ${inv}`);const next=inv-qty,unit=money_(x.p.costoUnitario);x.sh.getRange(x.row,7).setValue(next);x.sh.getRange(x.row,8).setValue(next*unit)}
function restoreInventory_(ss,name,qty){const x=productRow_(ss,name),next=number_(x.p.inventario,0)+qty,unit=money_(x.p.costoUnitario);x.sh.getRange(x.row,7).setValue(next);x.sh.getRange(x.row,8).setValue(next*unit)}
function addInventoryPurchase_(ss,name,qty,total){const x=productRow_(ss,name),oldQty=number_(x.p.inventario,0),oldUnit=money_(x.p.costoUnitario),newQty=oldQty+qty,newTotal=oldQty*oldUnit+total,newUnit=newQty>0?newTotal/newQty:0;x.sh.getRange(x.row,7,1,3).setValues([[newQty,newTotal,newUnit]]);return total/qty}

function getSettings_(ss){const sh=ss.getSheetByName(CONFIG_SHEET),rows=sh.getRange(1,1,Math.max(sh.getLastRow(),1),2).getValues(),o={tiktokDays:2,printaDays:3};rows.forEach(r=>{if(r[0]==='TIKTOK_DIAS_ENVIO')o.tiktokDays=number_(r[1],2);if(r[0]==='PRINTA_DIAS_ENVIO')o.printaDays=number_(r[1],3)});return o}
function getHandlingDays_(origin,s){return origin==='TikTok'?s.tiktokDays:s.printaDays}
function addBusinessDays_(date,days){const d=new Date(date);let a=0;while(a<days){d.setDate(d.getDate()+1);if(d.getDay()!==0&&d.getDay()!==6)a++}return formatDate_(d)}
function parseDate_(v){if(!v)return null;const p=String(v).split('-').map(Number);return p.length===3?new Date(p[0],p[1]-1,p[2],12):new Date(v)}
function readRows_(sh){if(!sh||sh.getLastRow()<2)return[];const vals=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getDisplayValues(),h=vals.shift();return vals.filter(r=>r[0]).map(r=>{const o={};h.forEach((x,i)=>o[toKey_(x)]=r[i]);return o})}
function buildSummary_(o,e){const m=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM'),mo=o.filter(x=>String(x.fecha||'').startsWith(m)),me=e.filter(x=>String(x.fecha||'').startsWith(m)),ing=mo.reduce((s,x)=>s+money_(x.pagado),0),cost=mo.reduce((s,x)=>s+money_(x.costoTotal),0),gas=me.filter(x=>!['Productos','Tinta'].includes(x.categoria)).reduce((s,x)=>s+money_(x.cantidad),0);return{pendientes:o.filter(x=>x.estado==='Pendiente').length,listos:o.filter(x=>x.estado==='Listo').length,ingresos:ing,costos:cost,gastos:gas,ganancia:ing-cost-gas}}
function upsertById_(sh,id,row){const m=findRowsById_(sh,id);if(m.length>1)throw new Error('ID duplicado.');m.length===1?sh.getRange(m[0],1,1,row.length).setValues([row]):sh.appendRow(row)}
function deleteById_(name,id){const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name),m=findRowsById_(sh,id);if(m.length>1)throw new Error('ID duplicado.');if(m.length===1)sh.deleteRow(m[0])}
function findRowsById_(sh,id){if(!sh||sh.getLastRow()<2||!clean_(id))return[];const v=sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues(),t=clean_(id),r=[];v.forEach((x,i)=>{if(clean_(x[0])===t)r.push(i+2)});return r}
function nextId_(p){return`${p}-${Utilities.getUuid().replace(/-/g,'').toUpperCase()}`}
function formatDate_(d){return Utilities.formatDate(d,'America/New_York','yyyy-MM-dd')}
function money_(v){const n=Number(String(v||0).replace(/[^0-9.-]/g,''));return isNaN(n)?0:n}
function number_(v,f){const n=Number(v);return isNaN(n)?f:n}
function clean_(v){return String(v==null?'':v).trim()}
function toKey_(h){return String(h).toLowerCase().replace(/_([a-z])/g,(_,c)=>c.toUpperCase())}
