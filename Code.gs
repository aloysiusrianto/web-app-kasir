/**
 * Web App Kasir — Google Apps Script backend.
 * Spreadsheet: Products, Sales, SaleItems, Settings.
 *
 * Dual transport:
 *   - google.script.run  (jika UI di-host di Apps Script)
 *   - doGet / doPost JSON (jika UI di Vercel; Content-Type text/plain
 *     agar browser tidak preflight CORS)
 */

var SHEET_PRODUCTS = 'Products';
var SHEET_SALES = 'Sales';
var SHEET_ITEMS = 'SaleItems';
var SHEET_SETTINGS = 'Settings';
var PROP_SPREADSHEET_ID = 'KASIR_SPREADSHEET_ID';

var PRODUCT_HEADERS = ['id', 'sku', 'name', 'category', 'price', 'stock', 'active', 'updatedAt'];
var SALES_HEADERS = [
  'id', 'createdAt', 'subtotal', 'taxRate', 'taxAmount', 'total',
  'paymentMethod', 'cashReceived', 'change', 'note', 'itemCount'
];
var ITEM_HEADERS = ['saleId', 'productId', 'sku', 'name', 'qty', 'price', 'lineTotal'];
var SETTING_HEADERS = ['key', 'value'];

var DEFAULT_SETTINGS = {
  storeName: 'Warung Kasir',
  taxRate: '0',
  currency: 'IDR'
};

/* -------------------------------------------------------------------------- */
/*  Web app entry                                                              */
/* -------------------------------------------------------------------------- */

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.api === '1' || params.action) {
    return jsonOutput_(dispatch_(params.action || 'health', parsePayload_(params.payload)));
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Kasir')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  var body = {};
  try {
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    return jsonOutput_({ ok: false, error: 'JSON tidak valid: ' + err.message });
  }
  var params = (e && e.parameter) || {};
  var action = body.action || params.action || 'health';
  var payload = body.payload !== undefined ? body.payload : parsePayload_(params.payload);
  return jsonOutput_(dispatch_(action, payload));
}

function parsePayload_(raw) {
  if (raw === undefined || raw === null || raw === '') return {};
  if (typeof raw === 'object') return raw;
  try {
    return JSON.parse(raw);
  } catch (err) {
    return {};
  }
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function dispatch_(action, payload) {
  try {
    ensureWorkbook_();
    payload = payload || {};
    switch (String(action || '')) {
      case 'health':
        return { ok: true, data: { status: 'ok', spreadsheetId: getSpreadsheetId_() } };
      case 'bootstrap':
        return { ok: true, data: bootstrap_() };
      case 'listProducts':
        return { ok: true, data: listProducts_(payload) };
      case 'saveProduct':
        return { ok: true, data: saveProduct_(payload) };
      case 'deleteProduct':
        return { ok: true, data: deleteProduct_(payload) };
      case 'checkout':
        return { ok: true, data: checkout_(payload) };
      case 'listSales':
        return { ok: true, data: listSales_(payload) };
      case 'getSale':
        return { ok: true, data: getSale_(payload) };
      case 'saveSettings':
        return { ok: true, data: saveSettings_(payload) };
      default:
        return { ok: false, error: 'Aksi tidak dikenal: ' + action };
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

/* google.script.run wrappers (same names as actions) */

function apiHealth() { return dispatch_('health', {}); }
function apiBootstrap() { return dispatch_('bootstrap', {}); }
function apiListProducts(payload) { return dispatch_('listProducts', payload); }
function apiSaveProduct(payload) { return dispatch_('saveProduct', payload); }
function apiDeleteProduct(payload) { return dispatch_('deleteProduct', payload); }
function apiCheckout(payload) { return dispatch_('checkout', payload); }
function apiListSales(payload) { return dispatch_('listSales', payload); }
function apiGetSale(payload) { return dispatch_('getSale', payload); }
function apiSaveSettings(payload) { return dispatch_('saveSettings', payload); }

/* -------------------------------------------------------------------------- */
/*  Spreadsheet bootstrap                                                      */
/* -------------------------------------------------------------------------- */

function getSpreadsheetId_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_SPREADSHEET_ID) || '';
}

function ensureWorkbook_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PROP_SPREADSHEET_ID);
  var ss = null;

  if (id) {
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (err) {
      ss = null;
    }
  }

  if (!ss) {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (err) {
      ss = null;
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create('Kasir — Data Toko');
  }

  props.setProperty(PROP_SPREADSHEET_ID, ss.getId());
  ensureSheet_(ss, SHEET_PRODUCTS, PRODUCT_HEADERS);
  ensureSheet_(ss, SHEET_SALES, SALES_HEADERS);
  ensureSheet_(ss, SHEET_ITEMS, ITEM_HEADERS);
  var settings = ensureSheet_(ss, SHEET_SETTINGS, SETTING_HEADERS);
  seedSettings_(settings);
  seedDemoProducts_(ss);
  return ss;
}

function book_() {
  var id = getSpreadsheetId_();
  if (!id) return ensureWorkbook_();
  return SpreadsheetApp.openById(id);
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  var lastCol = Math.max(sheet.getLastColumn(), headers.length);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var missing = false;
  for (var i = 0; i < headers.length; i++) {
    if (String(existing[i] || '') !== headers[i]) {
      missing = true;
      break;
    }
  }
  if (missing) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function seedSettings_(sheet) {
  var map = readSettingsMap_(sheet);
  var keys = Object.keys(DEFAULT_SETTINGS);
  var writes = [];
  for (var i = 0; i < keys.length; i++) {
    if (map[keys[i]] === undefined) {
      writes.push([keys[i], DEFAULT_SETTINGS[keys[i]]]);
    }
  }
  if (writes.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, writes.length, 2).setValues(writes);
  }
}

function seedDemoProducts_(ss) {
  var sheet = ss.getSheetByName(SHEET_PRODUCTS);
  if (sheet.getLastRow() > 1) return;
  var now = new Date().toISOString();
  var demo = [
    ['P-TEH', 'TEH-001', 'Es Teh Manis', 'Minuman', 5000, 80, 'TRUE', now],
    ['P-KOPI', 'KOPI-001', 'Kopi Tubruk', 'Minuman', 8000, 50, 'TRUE', now],
    ['P-NASI', 'NASI-001', 'Nasi Goreng', 'Makanan', 15000, 30, 'TRUE', now],
    ['P-MIE', 'MIE-001', 'Mie Goreng', 'Makanan', 12000, 30, 'TRUE', now],
    ['P-AIR', 'AIR-001', 'Air Mineral', 'Minuman', 4000, 100, 'TRUE', now],
    ['P-ROTI', 'ROTI-001', 'Roti Bakar', 'Makanan', 10000, 20, 'TRUE', now]
  ];
  sheet.getRange(2, 1, demo.length, PRODUCT_HEADERS.length).setValues(demo);
}

/* -------------------------------------------------------------------------- */
/*  Settings                                                                   */
/* -------------------------------------------------------------------------- */

function readSettingsMap_(sheet) {
  sheet = sheet || book_().getSheetByName(SHEET_SETTINGS);
  var last = sheet.getLastRow();
  var map = {};
  if (last < 2) return map;
  var rows = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    var key = String(rows[i][0] || '').trim();
    if (key) map[key] = String(rows[i][1]);
  }
  return map;
}

function settingsObject_() {
  var map = readSettingsMap_();
  return {
    storeName: map.storeName || DEFAULT_SETTINGS.storeName,
    taxRate: Number(map.taxRate || 0),
    currency: map.currency || 'IDR',
    spreadsheetId: getSpreadsheetId_(),
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + getSpreadsheetId_()
  };
}

function saveSettings_(payload) {
  var sheet = book_().getSheetByName(SHEET_SETTINGS);
  var map = readSettingsMap_(sheet);
  if (payload.storeName !== undefined) map.storeName = String(payload.storeName).trim() || DEFAULT_SETTINGS.storeName;
  if (payload.taxRate !== undefined) {
    var rate = Number(payload.taxRate);
    if (isNaN(rate) || rate < 0 || rate > 100) throw new Error('Pajak harus 0–100.');
    map.taxRate = String(rate);
  }
  if (payload.currency !== undefined) map.currency = String(payload.currency || 'IDR');

  var keys = Object.keys(map);
  var values = keys.map(function (k) { return [k, map[k]]; });
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }
  if (values.length) {
    sheet.getRange(2, 1, values.length, 2).setValues(values);
  }
  return settingsObject_();
}

/* -------------------------------------------------------------------------- */
/*  Products                                                                   */
/* -------------------------------------------------------------------------- */

function rowsToObjects_(sheet, headers) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, 1, last - 1, headers.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = {};
    for (var c = 0; c < headers.length; c++) {
      row[headers[c]] = values[i][c];
    }
    row._row = i + 2;
    out.push(row);
  }
  return out;
}

function toBool_(v) {
  if (v === true || v === 1) return true;
  var s = String(v).toUpperCase();
  return s === 'TRUE' || s === 'YES' || s === '1';
}

function productView_(row) {
  return {
    id: String(row.id),
    sku: String(row.sku || ''),
    name: String(row.name || ''),
    category: String(row.category || ''),
    price: Number(row.price) || 0,
    stock: Number(row.stock) || 0,
    active: toBool_(row.active),
    updatedAt: row.updatedAt ? String(row.updatedAt) : ''
  };
}

function listProducts_(payload) {
  var includeInactive = !!(payload && payload.includeInactive);
  var sheet = book_().getSheetByName(SHEET_PRODUCTS);
  var rows = rowsToObjects_(sheet, PRODUCT_HEADERS);
  var items = [];
  for (var i = 0; i < rows.length; i++) {
    var p = productView_(rows[i]);
    if (!includeInactive && !p.active) continue;
    items.push(p);
  }
  items.sort(function (a, b) {
    return a.name.localeCompare(b.name, 'id');
  });
  return items;
}

function findProductRow_(id) {
  var sheet = book_().getSheetByName(SHEET_PRODUCTS);
  var rows = rowsToObjects_(sheet, PRODUCT_HEADERS);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(id)) return rows[i];
  }
  return null;
}

function newId_(prefix) {
  return prefix + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
}

function saveProduct_(payload) {
  var name = String(payload.name || '').trim();
  if (!name) throw new Error('Nama produk wajib diisi.');
  var price = Number(payload.price);
  var stock = Number(payload.stock);
  if (isNaN(price) || price < 0) throw new Error('Harga tidak valid.');
  if (isNaN(stock) || stock < 0) throw new Error('Stok tidak valid.');

  var sheet = book_().getSheetByName(SHEET_PRODUCTS);
  var now = new Date().toISOString();
  var sku = String(payload.sku || '').trim();
  var category = String(payload.category || '').trim();
  var active = payload.active === false ? 'FALSE' : 'TRUE';

  if (payload.id) {
    var existing = findProductRow_(payload.id);
    if (!existing) throw new Error('Produk tidak ditemukan.');
    sheet.getRange(existing._row, 1, 1, PRODUCT_HEADERS.length).setValues([[
      String(existing.id), sku, name, category, price, stock, active, now
    ]]);
    return productView_(findProductRow_(payload.id));
  }

  var id = newId_('P');
  sheet.appendRow([id, sku, name, category, price, stock, active, now]);
  return productView_({
    id: id, sku: sku, name: name, category: category,
    price: price, stock: stock, active: active, updatedAt: now
  });
}

function deleteProduct_(payload) {
  var existing = findProductRow_(payload && payload.id);
  if (!existing) throw new Error('Produk tidak ditemukan.');
  var sheet = book_().getSheetByName(SHEET_PRODUCTS);
  sheet.getRange(existing._row, 7).setValue('FALSE');
  sheet.getRange(existing._row, 8).setValue(new Date().toISOString());
  return { id: String(existing.id), active: false };
}

/* -------------------------------------------------------------------------- */
/*  Sales                                                                      */
/* -------------------------------------------------------------------------- */

function checkout_(payload) {
  var lines = payload && payload.items;
  if (!lines || !lines.length) throw new Error('Keranjang kosong.');

  var settings = settingsObject_();
  var taxRate = payload.taxRate !== undefined ? Number(payload.taxRate) : settings.taxRate;
  if (isNaN(taxRate) || taxRate < 0) taxRate = 0;

  var paymentMethod = String(payload.paymentMethod || 'cash');
  var cashReceived = Number(payload.cashReceived || 0);
  var note = String(payload.note || '').trim();

  var ss = book_();
  var productSheet = ss.getSheetByName(SHEET_PRODUCTS);
  var products = rowsToObjects_(productSheet, PRODUCT_HEADERS);
  var byId = {};
  for (var i = 0; i < products.length; i++) {
    byId[String(products[i].id)] = products[i];
  }

  var normalized = [];
  var subtotal = 0;
  for (var j = 0; j < lines.length; j++) {
    var line = lines[j];
    var qty = Number(line.qty);
    if (!qty || qty <= 0 || Math.floor(qty) !== qty) {
      throw new Error('Jumlah item harus bilangan bulat positif.');
    }
    var prod = byId[String(line.productId)];
    if (!prod || !toBool_(prod.active)) {
      throw new Error('Produk tidak tersedia: ' + (line.name || line.productId));
    }
    var stock = Number(prod.stock) || 0;
    if (stock < qty) {
      throw new Error('Stok ' + prod.name + ' tidak cukup (tersisa ' + stock + ').');
    }
    var price = Number(prod.price) || 0;
    var lineTotal = price * qty;
    subtotal += lineTotal;
    normalized.push({
      product: prod,
      qty: qty,
      price: price,
      lineTotal: lineTotal
    });
  }

  var taxAmount = Math.round(subtotal * (taxRate / 100));
  var total = subtotal + taxAmount;
  if (paymentMethod === 'cash') {
    if (isNaN(cashReceived) || cashReceived < total) {
      throw new Error('Uang diterima kurang dari total.');
    }
  } else {
    cashReceived = total;
  }
  var change = cashReceived - total;
  var saleId = newId_('S');
  var createdAt = new Date().toISOString();

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // Re-read stock under lock
    products = rowsToObjects_(productSheet, PRODUCT_HEADERS);
    byId = {};
    for (var r = 0; r < products.length; r++) {
      byId[String(products[r].id)] = products[r];
    }
    for (var n = 0; n < normalized.length; n++) {
      var fresh = byId[String(normalized[n].product.id)];
      if (!fresh) throw new Error('Produk hilang saat checkout.');
      var freshStock = Number(fresh.stock) || 0;
      if (freshStock < normalized[n].qty) {
        throw new Error('Stok ' + fresh.name + ' tidak cukup (tersisa ' + freshStock + ').');
      }
      productSheet.getRange(fresh._row, 6).setValue(freshStock - normalized[n].qty);
      productSheet.getRange(fresh._row, 8).setValue(createdAt);
    }

    ss.getSheetByName(SHEET_SALES).appendRow([
      saleId, createdAt, subtotal, taxRate, taxAmount, total,
      paymentMethod, cashReceived, change, note, normalized.length
    ]);

    var itemRows = normalized.map(function (nline) {
      return [
        saleId,
        String(nline.product.id),
        String(nline.product.sku || ''),
        String(nline.product.name),
        nline.qty,
        nline.price,
        nline.lineTotal
      ];
    });
    var itemSheet = ss.getSheetByName(SHEET_ITEMS);
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);
  } finally {
    lock.releaseLock();
  }

  return {
    id: saleId,
    createdAt: createdAt,
    subtotal: subtotal,
    taxRate: taxRate,
    taxAmount: taxAmount,
    total: total,
    paymentMethod: paymentMethod,
    cashReceived: cashReceived,
    change: change,
    note: note,
    items: normalized.map(function (nline) {
      return {
        productId: String(nline.product.id),
        sku: String(nline.product.sku || ''),
        name: String(nline.product.name),
        qty: nline.qty,
        price: nline.price,
        lineTotal: nline.lineTotal
      };
    })
  };
}

function listSales_(payload) {
  var limit = Math.min(Number((payload && payload.limit) || 50), 200);
  var sheet = book_().getSheetByName(SHEET_SALES);
  var rows = rowsToObjects_(sheet, SALES_HEADERS);
  rows.sort(function (a, b) {
    return String(b.createdAt).localeCompare(String(a.createdAt));
  });
  return rows.slice(0, limit).map(saleView_);
}

function saleView_(row) {
  return {
    id: String(row.id),
    createdAt: String(row.createdAt || ''),
    subtotal: Number(row.subtotal) || 0,
    taxRate: Number(row.taxRate) || 0,
    taxAmount: Number(row.taxAmount) || 0,
    total: Number(row.total) || 0,
    paymentMethod: String(row.paymentMethod || ''),
    cashReceived: Number(row.cashReceived) || 0,
    change: Number(row.change) || 0,
    note: String(row.note || ''),
    itemCount: Number(row.itemCount) || 0
  };
}

function getSale_(payload) {
  var id = String((payload && payload.id) || '');
  var sales = rowsToObjects_(book_().getSheetByName(SHEET_SALES), SALES_HEADERS);
  var found = null;
  for (var i = 0; i < sales.length; i++) {
    if (String(sales[i].id) === id) {
      found = saleView_(sales[i]);
      break;
    }
  }
  if (!found) throw new Error('Transaksi tidak ditemukan.');
  var items = rowsToObjects_(book_().getSheetByName(SHEET_ITEMS), ITEM_HEADERS)
    .filter(function (row) { return String(row.saleId) === id; })
    .map(function (row) {
      return {
        productId: String(row.productId),
        sku: String(row.sku || ''),
        name: String(row.name || ''),
        qty: Number(row.qty) || 0,
        price: Number(row.price) || 0,
        lineTotal: Number(row.lineTotal) || 0
      };
    });
  found.items = items;
  return found;
}

function bootstrap_() {
  return {
    settings: settingsObject_(),
    products: listProducts_({ includeInactive: false }),
    sales: listSales_({ limit: 30 })
  };
}
