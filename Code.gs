const SPREADSHEET_ID = '1eOWiWZh0zTnVLnv53bTDXvlFJJG5ZbvR6Upr7_9uLb8';
const ORDERS_SHEET = 'ORDENES';
const EXPENSES_SHEET = 'GASTOS';
const CONFIG_SHEET = 'CONFIGURACION';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('PRINTA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const orders = readRows_(ss.getSheetByName(ORDERS_SHEET));
  const expenses = readRows_(ss.getSheetByName(EXPENSES_SHEET));
  return { orders, expenses, summary: buildSummary_(orders, expenses) };
}

function saveOrder(order) {
  validateOrder_(order);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDERS_SHEET);
  const now = new Date();
  const id = order.id || nextId_(sheet, 'PC');
  const row = [
    id,
    order.fecha || formatDate_(now),
    order.estado || 'Pendiente',
    order.origen || 'Otro',
    clean_(order.cliente),
    clean_(order.producto),
    number_(order.cantidad, 1),
    clean_(order.especificaciones),
    money_(order.total),
    money_(order.pagado),
    Math.max(0, money_(order.total) - money_(order.pagado)),
    order.fechaLimite || '',
    clean_(order.notas),
    clean_(order.fotosReferencia),
    clean_(order.fotoProducto),
    clean_(order.fotoEtiqueta),
    order.creadoEn || now,
    now
  ];
  upsertById_(sheet, id, row);
  return getAppData();
}

function setOrderStatus(id, status) {
  if (!['Pendiente', 'Listo'].includes(status)) throw new Error('Estado inválido.');
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDERS_SHEET);
  const row = findRowById_(sheet, id);
  if (!row) throw new Error('Orden no encontrada.');
  sheet.getRange(row, 3).setValue(status);
  sheet.getRange(row, 18).setValue(new Date());
  return getAppData();
}

function deleteOrder(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDERS_SHEET);
  const row = findRowById_(sheet, id);
  if (row) sheet.deleteRow(row);
  return getAppData();
}

function saveExpense(expense) {
  if (!expense || money_(expense.cantidad) <= 0) throw new Error('Escribe una cantidad válida.');
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET);
  const now = new Date();
  const id = expense.id || nextId_(sheet, 'GA');
  const row = [
    id,
    expense.fecha || formatDate_(now),
    expense.categoria || 'Otro',
    clean_(expense.lugar),
    money_(expense.cantidad),
    clean_(expense.nota),
    clean_(expense.recibo),
    expense.creadoEn || now
  ];
  upsertById_(sheet, id, row);
  return getAppData();
}

function deleteExpense(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EXPENSES_SHEET);
  const row = findRowById_(sheet, id);
  if (row) sheet.deleteRow(row);
  return getAppData();
}

function uploadFile(fileData, fileName, mimeType, orderId, category) {
  if (!fileData) return '';
  const folder = getUploadsFolder_();
  const bytes = Utilities.base64Decode(String(fileData).split(',').pop());
  const safeName = `${orderId || 'GENERAL'}_${category || 'archivo'}_${Date.now()}_${fileName || 'archivo'}`;
  const file = folder.createFile(Utilities.newBlob(bytes, mimeType || 'application/octet-stream', safeName));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getUploadsFolder_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const config = ss.getSheetByName(CONFIG_SHEET);
  const values = config.getRange(1, 1, Math.max(config.getLastRow(), 1), 2).getValues();
  let folderId = '';
  let configRow = 0;
  values.forEach((r, i) => {
    if (r[0] === 'CARPETA_FOTOS_ID') { folderId = r[1]; configRow = i + 1; }
  });
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) {}
  }
  const folder = DriveApp.createFolder('PRINTA - Fotos y recibos');
  if (configRow) config.getRange(configRow, 2).setValue(folder.getId());
  return folder;
}

function readRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getDisplayValues();
  const headers = values.shift();
  return values.filter(r => r[0]).map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[toKey_(h)] = r[i]);
    return obj;
  });
}

function buildSummary_(orders, expenses) {
  const month = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const monthOrders = orders.filter(o => String(o.fecha || '').startsWith(month));
  const monthExpenses = expenses.filter(e => String(e.fecha || '').startsWith(month));
  const ingresos = monthOrders.reduce((s, o) => s + money_(o.pagado), 0);
  const gastos = monthExpenses.reduce((s, e) => s + money_(e.cantidad), 0);
  return {
    pendientes: orders.filter(o => o.estado === 'Pendiente').length,
    listos: orders.filter(o => o.estado === 'Listo').length,
    ingresos,
    gastos,
    ganancia: ingresos - gastos,
    tiktok: monthOrders.filter(o => o.origen === 'TikTok').reduce((s, o) => s + money_(o.pagado), 0),
    printaCrea: monthOrders.filter(o => o.origen === 'Printa Crea').reduce((s, o) => s + money_(o.pagado), 0)
  };
}

function upsertById_(sheet, id, row) {
  const existing = findRowById_(sheet, id);
  if (existing) sheet.getRange(existing, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
}

function findRowById_(sheet, id) {
  if (sheet.getLastRow() < 2) return 0;
  const match = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return match ? match.getRow() : 0;
}

function nextId_(sheet, prefix) {
  const n = Math.max(1, sheet.getLastRow());
  return `${prefix}-${String(n).padStart(4, '0')}`;
}

function validateOrder_(o) {
  if (!o) throw new Error('Faltan datos de la orden.');
  if (!clean_(o.cliente)) throw new Error('Escribe el nombre del cliente.');
  if (!clean_(o.producto)) throw new Error('Escribe el producto.');
  if (money_(o.pagado) > money_(o.total)) throw new Error('Lo pagado no puede ser mayor que el total.');
}

function formatDate_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function money_(v) { const n = Number(String(v || 0).replace(/[^0-9.-]/g, '')); return isNaN(n) ? 0 : n; }
function number_(v, fallback) { const n = Number(v); return isNaN(n) ? fallback : n; }
function clean_(v) { return String(v == null ? '' : v).trim(); }
function toKey_(h) { return String(h).toLowerCase().replace(/_([a-z])/g, (_, c) => c.toUpperCase()); }
