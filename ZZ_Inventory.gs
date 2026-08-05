const PRINTA_V2_CACHE='PRINTA_V2_20260805_1';
const INVENTORY_SHEET='INVENTARIO';
const ORDER_MATERIALS_SHEET='ORDEN_MATERIALES';

function doGet(){
  setupPrintaV2_();
  return HtmlService.createTemplateFromFile('AppV2').evaluate()
    .setTitle('PRINTA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPrintaV2Data(){
  setupPrintaV2_();
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const orders=sheetObjects_(ss.getSheetByName(ORDERS_SHEET));
  const expenses=sheetObjects_(ss.getSheetByName(EXPENSES_SHEET));
  const products=sheetObjects_(ss.getSheetByName(PRODUCTS_SHEET));
  const inventory=sheetObjects_(ss.getSheetByName(INVENTORY_SHEET));
  return {orders,expenses,products,inventory,summary:summaryV2_(orders,expenses)};
}

function setupPrintaV2_(){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheetV2_(ss,INVENTORY_SHEET,['ID','NOMBRE','CATEGORIA','CANTIDAD','COSTO_TOTAL','COSTO_UNITARIO','NOTA','CREADO_EN','ACTUALIZADO_EN']);
  ensureSheetV2_(ss,ORDER_MATERIALS_SHEET,['ORDER_ID','MATERIALES_JSON','OTROS_COSTOS_JSON','COSTO_MATERIALES','COSTO_OTROS','COSTO_TOTAL','GANANCIA','DESCONTADO','ACTUALIZADO_EN']);
  ensureColumnsV2_(ss.getSheetByName(ORDERS_SHEET),['COSTO_TOTAL','GANANCIA','MATERIALES_JSON','OTROS_COSTOS_JSON','INVENTARIO_DESCONTADO']);
}

function saveInventoryItemV2(item){
  setupPrintaV2_();
  if(!item||!cleanV2_(item.nombre))throw new Error('Escribe el nombre del artículo.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(INVENTORY_SHEET);
    const now=new Date(),id=cleanV2_(item.id)||('INV-'+Utilities.getUuid().replace(/-/g,''));
    const old=findObjectByIdV2_(sh,id),oldQty=numV2_(old&&old.cantidad),oldTotal=numV2_(old&&old.costoTotal);
    const addQty=Math.max(0,numV2_(item.cantidadAgregar||item.cantidad));
    const addTotal=Math.max(0,numV2_(item.costoAgregar||item.costoTotal));
    const editing=!!cleanV2_(item.id);
    const qty=editing&&item.modo==='editar'?Math.max(0,numV2_(item.cantidad)):oldQty+addQty;
    const total=editing&&item.modo==='editar'?Math.max(0,numV2_(item.costoTotal)):oldTotal+addTotal;
    const unit=qty>0?total/qty:0;
    const row=[id,cleanV2_(item.nombre),cleanV2_(item.categoria)||'Material',qty,total,unit,cleanV2_(item.nota),old&&old.creadoEn?old.creadoEn:now,now];
    upsertRowV2_(sh,id,row);
  }finally{lock.releaseLock()}
  return getPrintaV2Data();
}

function deleteInventoryItemV2(id){
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=ss.getSheetByName(INVENTORY_SHEET);
    const used=getAllOrderMaterialsV2_().some(x=>(x.materiales||[]).some(m=>m.id===id));
    if(used)throw new Error('Este artículo está usado en órdenes y no se puede borrar. Puedes dejar su cantidad en 0.');
    const r=findRowV2_(sh,id);if(r)sh.deleteRow(r);
  }finally{lock.releaseLock()}
  return getPrintaV2Data();
}

function saveOrderV2(order){
  setupPrintaV2_();
  if(!order||!cleanV2_(order.cliente)||!cleanV2_(order.producto))throw new Error('Completa cliente y producto.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=ss.getSheetByName(ORDERS_SHEET),now=new Date();
    const id=cleanV2_(order.id)||('PC-'+Utilities.getUuid().replace(/-/g,''));
    const existing=findRowV2_(sh,id),headers=getHeadersV2_(sh),obj=existing?rowObjectV2_(sh,existing,headers):{};
    obj.id=id;obj.fecha=cleanV2_(order.fecha)||formatDateV2_(now);obj.estado=cleanV2_(order.estado)||obj.estado||'Pendiente';obj.origen=cleanV2_(order.origen)||'Otro';obj.cliente=cleanV2_(order.cliente);obj.producto=cleanV2_(order.producto);obj.cantidad=Math.max(1,numV2_(order.cantidad));obj.notas=cleanV2_(order.notas);obj.pagado=numV2_(order.pagado);obj.fechaEnvio=cleanV2_(order.fechaEnvio);obj.foto=cleanV2_(order.foto)||obj.foto||'';obj.creadoEn=obj.creadoEn||now;obj.actualizadoEn=now;
    const materiales=parseJsonV2_(order.materialesJson,[]),otros=parseJsonV2_(order.otrosCostosJson,[]);
    validateMaterialsV2_(materiales);
    const costs=calculateOrderCostsV2_(materiales,otros,obj.pagado);
    obj.costoTotal=costs.total;obj.ganancia=costs.ganancia;obj.materialesJson=JSON.stringify(materiales);obj.otrosCostosJson=JSON.stringify(otros);obj.inventarioDescontado=obj.inventarioDescontado||'NO';
    writeObjectRowV2_(sh,existing,obj,headers);
    saveOrderMaterialsV2_(id,materiales,otros,costs,obj.inventarioDescontado==='SI');
  }finally{lock.releaseLock()}
  return getPrintaV2Data();
}

function setOrderStatusV2(id,status){
  if(!['Pendiente','Listo'].includes(status))throw new Error('Estado inválido.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=ss.getSheetByName(ORDERS_SHEET),r=findRowV2_(sh,id);
    if(!r)throw new Error('Orden no encontrada.');
    const headers=getHeadersV2_(sh),o=rowObjectV2_(sh,r,headers),was=String(o.inventarioDescontado).toUpperCase()==='SI';
    const details=getOrderMaterialsV2_(id),materials=details.materiales||[];
    if(status==='Listo'&&!was){adjustInventoryV2_(materials,-1);o.inventarioDescontado='SI'}
    if(status==='Pendiente'&&was){adjustInventoryV2_(materials,1);o.inventarioDescontado='NO'}
    o.estado=status;o.actualizadoEn=new Date();writeObjectRowV2_(sh,r,o,headers);
    saveOrderMaterialsV2_(id,materials,details.otros||[],calculateOrderCostsV2_(materials,details.otros||[],numV2_(o.pagado)),o.inventarioDescontado==='SI');
  }finally{lock.releaseLock()}
  return getPrintaV2Data();
}

function deleteOrderV2(id){
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=ss.getSheetByName(ORDERS_SHEET),r=findRowV2_(sh,id);
    if(!r)throw new Error('Orden no encontrada.');
    const o=rowObjectV2_(sh,r,getHeadersV2_(sh)),d=getOrderMaterialsV2_(id);
    if(String(o.inventarioDescontado).toUpperCase()==='SI')adjustInventoryV2_(d.materiales||[],1);
    sh.deleteRow(r);
    const mh=ss.getSheetByName(ORDER_MATERIALS_SHEET),mr=findRowV2_(mh,id);if(mr)mh.deleteRow(mr);
  }finally{lock.releaseLock()}
  return getPrintaV2Data();
}

function calculateOrderPreviewV2(materials,others,paid){return calculateOrderCostsV2_(materials||[],others||[],paid)}

function calculateOrderCostsV2_(materials,others,paid){
  const inventory=sheetObjects_(SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(INVENTORY_SHEET));
  const map={};inventory.forEach(x=>map[x.id]=x);
  let materialCost=0;
  materials.forEach(m=>{const it=map[m.id];if(it)materialCost+=numV2_(it.costoUnitario)*Math.max(0,numV2_(m.cantidad))});
  const otherCost=others.reduce((s,x)=>s+Math.max(0,numV2_(x.costo)),0),total=materialCost+otherCost;
  return {materiales:materialCost,otros:otherCost,total,ganancia:numV2_(paid)-total};
}

function validateMaterialsV2_(materials){
  const inv=sheetObjects_(SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(INVENTORY_SHEET)),map={};inv.forEach(x=>map[x.id]=x);
  materials.forEach(m=>{if(!map[m.id])throw new Error('Uno de los materiales ya no existe en inventario.');if(numV2_(m.cantidad)<=0)throw new Error('La cantidad de material debe ser mayor que 0.')});
}

function adjustInventoryV2_(materials,direction){
  const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(INVENTORY_SHEET),headers=getHeadersV2_(sh);
  materials.forEach(m=>{const r=findRowV2_(sh,m.id);if(!r)throw new Error('Material no encontrado: '+m.nombre);const o=rowObjectV2_(sh,r,headers),qty=numV2_(m.cantidad),current=numV2_(o.cantidad),next=current+(direction*qty);if(next<0)throw new Error('No hay suficiente inventario de '+o.nombre+'. Disponible: '+current);o.cantidad=next;o.costoTotal=Math.max(0,numV2_(o.costoUnitario)*next);o.actualizadoEn=new Date();writeObjectRowV2_(sh,r,o,headers)});
}

function getOrderMaterialsV2_(id){
  const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDER_MATERIALS_SHEET),r=findRowV2_(sh,id);
  if(!r)return {materiales:[],otros:[]};
  const o=rowObjectV2_(sh,r,getHeadersV2_(sh));return {materiales:parseJsonV2_(o.materialesJson,[]),otros:parseJsonV2_(o.otrosCostosJson,[])};
}
function getAllOrderMaterialsV2_(){return sheetObjects_(SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDER_MATERIALS_SHEET)).map(o=>({orderId:o.orderId,materiales:parseJsonV2_(o.materialesJson,[]),otros:parseJsonV2_(o.otrosCostosJson,[])}))}
function saveOrderMaterialsV2_(id,materials,others,costs,discounted){
  const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDER_MATERIALS_SHEET),row=[id,JSON.stringify(materials),JSON.stringify(others),costs.materiales,costs.otros,costs.total,costs.ganancia,discounted?'SI':'NO',new Date()];upsertRowV2_(sh,id,row);
}

function uploadFileV2(data,name,type,owner){if(!data)return'';const bytes=Utilities.base64Decode(String(data).split(',').pop()),blob=Utilities.newBlob(bytes,type||'image/jpeg',(owner||'PRINTA')+'_'+Date.now()+'_'+(name||'foto'));const f=getUploadsFolder_().createFile(blob);f.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return f.getUrl()}

function summaryV2_(orders,expenses){const month=formatDateV2_(new Date()).slice(0,7),mo=orders.filter(x=>String(x.fecha||'').startsWith(month)),me=expenses.filter(x=>String(x.fecha||'').startsWith(month)),sales=mo.reduce((s,x)=>s+numV2_(x.pagado),0),costs=mo.reduce((s,x)=>s+numV2_(x.costoTotal),0),expensesTotal=me.reduce((s,x)=>s+numV2_(x.cantidad),0);return {pendientes:orders.filter(x=>x.estado==='Pendiente').length,listos:orders.filter(x=>x.estado==='Listo').length,ventas:sales,costos,gastos:expensesTotal,ganancia:sales-costs-expensesTotal}}
function ensureSheetV2_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.appendRow(headers);else ensureColumnsV2_(sh,headers);return sh}
function ensureColumnsV2_(sh,headers){const existing=sh.getLastColumn()?sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0]:[];headers.forEach(h=>{if(!existing.includes(h)){sh.getRange(1,sh.getLastColumn()+1).setValue(h);existing.push(h)}})}
function sheetObjects_(sh){if(!sh||sh.getLastRow()<2)return[];const vals=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getDisplayValues(),headers=vals.shift().map(keyV2_);return vals.filter(r=>r.some(Boolean)).map(r=>{const o={};headers.forEach((h,i)=>o[h]=r[i]);return o})}
function getHeadersV2_(sh){return sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(keyV2_)}
function rowObjectV2_(sh,row,headers){const vals=sh.getRange(row,1,1,headers.length).getValues()[0],o={};headers.forEach((h,i)=>o[h]=vals[i]);return o}
function writeObjectRowV2_(sh,row,obj,headers){const vals=headers.map(h=>obj[h]===undefined?'':obj[h]);if(row)sh.getRange(row,1,1,vals.length).setValues([vals]);else sh.appendRow(vals)}
function upsertRowV2_(sh,id,row){const r=findRowV2_(sh,id);if(r)sh.getRange(r,1,1,row.length).setValues([row]);else sh.appendRow(row)}
function findRowV2_(sh,id){if(!sh||sh.getLastRow()<2)return 0;const target=cleanV2_(id),vals=sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues();for(let i=0;i<vals.length;i++)if(cleanV2_(vals[i][0])===target)return i+2;return 0}
function findObjectByIdV2_(sh,id){const r=findRowV2_(sh,id);return r?rowObjectV2_(sh,r,getHeadersV2_(sh)):null}
function parseJsonV2_(v,f){try{return v?JSON.parse(v):f}catch(e){return f}}
function numV2_(v){const n=Number(String(v==null?'':v).replace(/[^0-9.-]/g,''));return isNaN(n)?0:n}
function cleanV2_(v){return String(v==null?'':v).trim()}
function keyV2_(v){return String(v).trim().toLowerCase().replace(/[^a-z0-9]+(.)/g,(_,c)=>c.toUpperCase())}
function formatDateV2_(d){return Utilities.formatDate(new Date(d),'America/New_York','yyyy-MM-dd')}
