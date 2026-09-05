const CONFIG = {
  SPREADSHEET_ID: '1eOWiWZh0zTnVLnv53bTDXvlFJJG5ZbvR6Upr7_9uLb8',
  SHEET_NAME: 'ORDENES',
  PRODUCT_SHEET_NAME: 'PRODUCTOS',
  IMAGE_FOLDER_NAME: 'PRINTA_ORDENES_IMAGES',
  TIME_ZONE: 'America/New_York',
  ORIGIN_DAYS: { TikTok: 2, Shopify: 3, Zelle: 3 },
  HEADERS: ['ID','Nombre','Numero Orden','Origen','Fecha Orden','Dias Habiles','Fecha Limite','Producto','Estado','Indicaciones','Imagenes','Creada','Actualizada','Fecha Envio']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ORDENES · Printacrea')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAppData() {
  const sh = getSheet_();
  const lastRow = sh.getLastRow();
  const products = getActiveProducts_();
  if (lastRow < 2) return { orders: [], products, today: today_(), originDays: CONFIG.ORIGIN_DAYS };

  const values = sh.getRange(2, 1, lastRow - 1, CONFIG.HEADERS.length).getValues();
  const orders = values
    .filter(r => r[0])
    .map(rowToOrder_)
    .sort((a, b) => {
      const da = a.dueDate || '9999-12-31';
      const db = b.dueDate || '9999-12-31';
      return da.localeCompare(db) || (b.createdAt || '').localeCompare(a.createdAt || '');
    });

  return { orders, products, today: today_(), originDays: CONFIG.ORIGIN_DAYS };
}

function getActiveProducts_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.PRODUCT_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  return rows
    .filter(r => String(r[6] || '').toUpperCase() === 'ACTIVE' && String(r[2] || '').trim())
    .map(r => ({
      product: String(r[0] || ''),
      variant: String(r[1] || ''),
      label: String(r[2] || ''),
      productGid: String(r[3] || ''),
      variantGid: String(r[4] || ''),
      sku: String(r[5] || '')
    }))
    .sort((a,b) => a.label.localeCompare(b.label, 'es', {sensitivity:'base'}));
}

function saveOrder(payload) {
  if (!payload || !String(payload.name || '').trim()) throw new Error('El nombre es obligatorio.');
  if (!CONFIG.ORIGIN_DAYS[payload.origin]) throw new Error('Origen no válido.');
  if (!String(payload.product || '').trim()) throw new Error('El producto es obligatorio.');

  const sh = getSheet_();
  const now = new Date();
  const orderDate = normalizeDate_(payload.orderDate) || today_();
  const businessDays = CONFIG.ORIGIN_DAYS[payload.origin];
  const dueDate = addBusinessDays_(orderDate, businessDays);
  const id = Utilities.getUuid();
  const imageUrls = saveImages_(payload.images || [], id, String(payload.name || '').trim());

  const row = [
    id,
    String(payload.name || '').trim(),
    String(payload.orderNumber || '').trim(),
    payload.origin,
    orderDate,
    businessDays,
    dueDate,
    String(payload.product || '').trim(),
    'Pendiente',
    String(payload.instructions || '').trim(),
    JSON.stringify(imageUrls),
    formatDateTime_(now),
    formatDateTime_(now),
    ''
  ];

  sh.appendRow(row);
  formatRowDates_(sh, sh.getLastRow());
  return { ok: true, order: rowToOrder_(row) };
}

function updateOrderStatus(id, status) {
  const allowed = ['Pendiente', 'Lista para envio', 'Enviada'];
  if (!allowed.includes(status)) throw new Error('Estado no válido.');

  const sh = getSheet_();
  const row = findRowById_(sh, id);
  if (!row) throw new Error('No se encontró la orden.');

  sh.getRange(row, 9).setValue(status);
  sh.getRange(row, 13).setValue(formatDateTime_(new Date()));
  sh.getRange(row, 14).setValue(status === 'Enviada' ? formatDateTime_(new Date()) : '');
  return { ok: true };
}

function deleteOrder(id) {
  const sh = getSheet_();
  const row = findRowById_(sh, id);
  if (!row) throw new Error('No se encontró la orden.');
  sh.deleteRow(row);
  return { ok: true };
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME);

  const current = sh.getRange(1, 1, 1, CONFIG.HEADERS.length).getValues()[0];
  if (current.join('|') !== CONFIG.HEADERS.join('|')) {
    sh.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function findRowById_(sh, id) {
  if (!id || sh.getLastRow() < 2) return 0;
  const finder = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return finder ? finder.getRow() : 0;
}

function rowToOrder_(r) {
  const orderDate = dateCellToYmd_(r[4]);
  const dueDate = dateCellToYmd_(r[6]);
  let images = [];
  try { images = Array.isArray(r[10]) ? r[10] : JSON.parse(r[10] || '[]'); } catch (e) { images = String(r[10] || '').split(',').map(s => s.trim()).filter(Boolean); }

  return {
    id: String(r[0] || ''),
    name: String(r[1] || ''),
    orderNumber: String(r[2] || ''),
    origin: String(r[3] || ''),
    orderDate,
    businessDays: Number(r[5] || 0),
    dueDate,
    product: String(r[7] || ''),
    status: String(r[8] || 'Pendiente'),
    instructions: String(r[9] || ''),
    images,
    createdAt: dateTimeCellToString_(r[11]),
    updatedAt: dateTimeCellToString_(r[12]),
    shippedAt: dateTimeCellToString_(r[13])
  };
}

function saveImages_(images, orderId, customerName) {
  if (!images || !images.length) return [];
  const folder = getImageFolder_();
  return images.map((img, index) => {
    if (!img || !img.dataUrl) return null;
    const match = String(img.dataUrl).match(/^data:([^;]+);base64,(.+)$/);
    if (!match) return null;
    const mime = match[1];
    const bytes = Utilities.base64Decode(match[2]);
    const ext = mime.indexOf('png') > -1 ? 'png' : mime.indexOf('webp') > -1 ? 'webp' : 'jpg';
    const safeName = customerName.replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ_-]+/g, '_').slice(0, 50);
    const blob = Utilities.newBlob(bytes, mime, safeName + '_' + orderId.slice(0, 8) + '_' + (index + 1) + '.' + ext);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  }).filter(Boolean);
}

function getImageFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('ORDENES_IMAGE_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) {}
  }
  const it = DriveApp.getFoldersByName(CONFIG.IMAGE_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.IMAGE_FOLDER_NAME);
  props.setProperty('ORDENES_IMAGE_FOLDER_ID', folder.getId());
  return folder;
}

function addBusinessDays_(ymd, days) {
  const parts = ymd.split('-').map(Number);
  let d = new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
  let added = 0;
  while (added < days) {
    d.setDate(d.getDate() + 1);
    const day = d.getDay();
    if (day !== 0 && day !== 6) added++;
  }
  return Utilities.formatDate(d, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
}

function normalizeDate_(value) {
  const s = String(value || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function today_() { return Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd'); }
function formatDateTime_(d) { return Utilities.formatDate(d, CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm:ss'); }

function dateCellToYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  return isNaN(d) ? '' : Utilities.formatDate(d, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
}

function dateTimeCellToString_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
  return String(v);
}

function formatRowDates_(sh, row) {
  sh.getRange(row, 5).setNumberFormat('yyyy-mm-dd');
  sh.getRange(row, 7).setNumberFormat('yyyy-mm-dd');
}
