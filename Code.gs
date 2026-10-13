/**
 * ORDER DESK — Google Apps Script backend
 * ---------------------------------------
 * Receives orders from the order page, writes them to this spreadsheet,
 * gives every customer a permanent number (C-001, C-002 …) keyed on their phone,
 * and tracks Ready / Delivered / Paid in separate columns with automatic timestamps.
 *
 * Setup (once):
 *   1. Paste this file into Extensions → Apps Script of a new Google Sheet.
 *   2. Change ORDER_KEY below (and the same value in index.html).
 *   3. Run  setup  from the toolbar and approve the permissions.
 *   4. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 *   5. Copy the /exec URL into index.html (ENDPOINT).
 */

const CONFIG = {
  SHOP_NAME: 'Order Desk',
  ORDER_KEY: 'change-this-key',  // must match ORDER_KEY in index.html
  MAX_QTY: 500,                  // largest quantity allowed per line
  NOTIFY_EMAIL: '',              // optional: email address to alert on every new order
  FIRST_ORDER_NO: 1001,
  MAX_ORDERS_PER_PHONE_PER_HOUR: 10
};

const TAB = {
  ORDERS: 'Orders',
  LINES: 'Lines',
  CUSTOMERS: 'Customers',
  PRODUCTS: 'Products',
  PICK: 'Pick List'
};

// Orders tab columns (1-based). Ready / Delivered / Paid are separate checkbox columns.
const O = {
  ORDER_NO: 1, PLACED: 2, CUST_NO: 3, NAME: 4, PHONE: 5, ITEMS: 6, TOTAL_LB: 7, TOTAL_PCS: 8,
  FOR_DATE: 9, NOTES: 10, READY: 11, READY_AT: 12, DELIVERED: 13, DELIVERED_AT: 14,
  PAID: 15, STATUS: 16, CLIENT_ID: 17
};
const ORDER_HEADERS = ['Order #', 'Placed At', 'Customer #', 'Customer Name', 'Phone', 'Items',
  'Total lb', 'Total pcs', 'For Date', 'Notes', 'Ready', 'Ready At', 'Delivered', 'Delivered At',
  'Paid', 'Status', 'Client ID'];

const LINE_HEADERS = ['Order #', 'Placed At', 'For Date', 'Customer #', 'Customer Name', 'Product', 'Qty', 'Unit'];
const CUSTOMER_HEADERS = ['Customer #', 'Name', 'Phone', 'Lookup Key', 'First Order', 'Last Order', 'Orders'];
const PRODUCT_HEADERS = ['Product', 'Unit', 'Step', 'Active', 'Sort'];

const SEED_PRODUCTS = [
  ['Drum Sticks', 'lb', 1, true, 10],
  ['Leg Quarters', 'lb', 1, true, 20],
  ['Whole Chicken', 'lb', 1, true, 30],
  ['Chicken Breast', 'lb', 1, true, 40],
  ['Chicken Wings', 'lb', 1, true, 50],
  ['Whole Baby Goat', 'each', 1, true, 60],
  ['Goat Meat (cut)', 'lb', 1, true, 70],
  ['Lamb (cut)', 'lb', 1, true, 80]
];

/* ------------------------------------------------------------------ */
/*  One-time setup: creates and formats every tab. Safe to re-run.     */
/* ------------------------------------------------------------------ */
function setup() {
  const ss = SpreadsheetApp.getActive();

  // Orders
  const orders = ensureSheet_(ss, TAB.ORDERS, ORDER_HEADERS);
  const n = orders.getMaxRows() - 1;
  [O.READY, O.DELIVERED, O.PAID].forEach(function (c) {
    orders.getRange(2, c, n, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireCheckbox().build());
  });
  orders.getRange(2, O.PLACED, n, 1).setNumberFormat('ddd mmm d, h:mm am/pm');
  orders.getRange(2, O.READY_AT, n, 1).setNumberFormat('ddd mmm d, h:mm am/pm');
  orders.getRange(2, O.DELIVERED_AT, n, 1).setNumberFormat('ddd mmm d, h:mm am/pm');
  orders.getRange(2, O.FOR_DATE, n, 1).setNumberFormat('ddd mmm d');
  orders.getRange(2, O.PHONE, n, 1).setNumberFormat('@');
  orders.getRange(2, O.CLIENT_ID, n, 1).setNumberFormat('@');
  orders.getRange(2, O.ITEMS, n, 1).setWrap(true);
  orders.getRange(2, 1, n, ORDER_HEADERS.length).setVerticalAlignment('top');
  setWidths_(orders, { 1: 70, 2: 130, 3: 80, 4: 150, 5: 110, 6: 260, 7: 70, 8: 70, 9: 95, 10: 180,
    11: 60, 12: 130, 13: 75, 14: 130, 15: 50, 16: 85, 17: 90 });
  orders.hideColumns(O.CLIENT_ID);

  const all = orders.getRange(2, 1, n, ORDER_HEADERS.length);
  const paidCell = orders.getRange(2, O.PAID, n, 1);
  orders.setConditionalFormatRules([
    rule_('=AND($M2=TRUE,$O2=FALSE)', paidCell, '#f8d7da'),        // delivered but unpaid
    rule_('=$P2="Delivered"', all, '#d9f0e3'),
    rule_('=$P2="Ready"', all, '#fff1cc'),
    rule_('=$P2="New"', all, '#fde4e1')
  ]);
  if (!orders.getFilter()) orders.getRange(1, 1, orders.getMaxRows(), ORDER_HEADERS.length).createFilter();

  // Lines (one row per item — feeds the pick list)
  const lines = ensureSheet_(ss, TAB.LINES, LINE_HEADERS);
  lines.getRange(2, 2, lines.getMaxRows() - 1, 1).setNumberFormat('ddd mmm d, h:mm am/pm');
  lines.getRange(2, 3, lines.getMaxRows() - 1, 1).setNumberFormat('ddd mmm d');
  setWidths_(lines, { 1: 70, 2: 130, 3: 95, 4: 80, 5: 150, 6: 160, 7: 60, 8: 60 });

  // Customers (permanent customer number per phone)
  const cust = ensureSheet_(ss, TAB.CUSTOMERS, CUSTOMER_HEADERS);
  cust.getRange(2, 3, cust.getMaxRows() - 1, 2).setNumberFormat('@');
  cust.getRange(2, 5, cust.getMaxRows() - 1, 2).setNumberFormat('mmm d, yyyy');
  cust.hideColumns(4);
  setWidths_(cust, { 1: 85, 2: 170, 3: 120, 5: 110, 6: 110, 7: 60 });

  // Products (edit this tab to change what customers can order)
  const prod = ensureSheet_(ss, TAB.PRODUCTS, PRODUCT_HEADERS);
  if (prod.getLastRow() < 2) {
    prod.getRange(2, 1, SEED_PRODUCTS.length, PRODUCT_HEADERS.length).setValues(SEED_PRODUCTS);
  }
  prod.getRange(2, 2, prod.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['lb', 'each'], true).build());
  prod.getRange(2, 4, prod.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build());
  setWidths_(prod, { 1: 180, 2: 60, 3: 50, 4: 60, 5: 50 });

  // Pick List (totals per product for a chosen date + order checklist)
  const pick = ss.getSheetByName(TAB.PICK) || ss.insertSheet(TAB.PICK);
  pick.clear();
  pick.getRange('A1').setValue('Date').setFontWeight('bold');
  pick.getRange('B1').setFormula('=TODAY()').setNumberFormat('ddd mmm d, yyyy')
    .setBackground('#fff1cc').setNote('Type any date here to see that day.');
  pick.getRange('A3').setValue('What to pack').setFontWeight('bold');
  pick.getRange('A4').setFormula(
    '=IFERROR(QUERY(Lines!A2:H,"select F, H, sum(G) where C = date \'"&TEXT(B1,"yyyy-mm-dd")&"\' ' +
    'group by F, H order by F label F \'Product\', H \'Unit\', sum(G) \'Total\'",0),"No orders for this date")');
  pick.getRange('E3').setValue('Orders for this date').setFontWeight('bold');
  pick.getRange('E4').setFormula(
    '=IFERROR(QUERY(Orders!A2:P,"select A, C, D, F, P where I = date \'"&TEXT(B1,"yyyy-mm-dd")&"\' ' +
    'order by A label A \'Order #\', C \'Customer #\', D \'Name\', F \'Items\', P \'Status\'",0),"—")');
  setWidths_(pick, { 1: 170, 2: 140, 3: 70, 4: 30, 5: 70, 6: 80, 7: 150, 8: 260, 9: 85 });
  pick.getRange('H:H').setWrap(true);

  // Counters
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('NEXT_ORDER')) props.setProperty('NEXT_ORDER', String(CONFIG.FIRST_ORDER_NO));
  if (!props.getProperty('NEXT_CUSTOMER')) props.setProperty('NEXT_CUSTOMER', String(Math.max(1, cust.getLastRow())));

  ['Sheet1'].forEach(function (name) {
    const s = ss.getSheetByName(name);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  ss.setActiveSheet(orders);
}

/* ------------------------------------------------------------------ */
/*  Web endpoints                                                       */
/* ------------------------------------------------------------------ */
function doGet() {
  return json_({ ok: true, shop: CONFIG.SHOP_NAME, message: 'Order Desk is running.' });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.key !== CONFIG.ORDER_KEY) return json_({ ok: false, error: 'Invalid shop key.' });

    if (body.action === 'products') {
      return json_({ ok: true, shop: CONFIG.SHOP_NAME, products: getProducts_() });
    }
    if (body.action === 'order') {
      return json_(placeOrder_(body.order || {}));
    }
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return json_({ ok: false, error: 'Server error: ' + (err && err.message ? err.message : err), retry: true });
  }
}

/* ------------------------------------------------------------------ */
/*  Ready / Delivered timestamps (runs automatically when you tick)     */
/* ------------------------------------------------------------------ */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== TAB.ORDERS || e.range.getRow() < 2 || e.range.getNumColumns() !== 1) return;
  const col = e.range.getColumn();
  if (col !== O.READY && col !== O.DELIVERED) return;

  for (let i = 0; i < e.range.getNumRows(); i++) {
    const r = e.range.getRow() + i;
    if (!sh.getRange(r, O.ORDER_NO).getValue()) continue;
    const checked = sh.getRange(r, col).getValue() === true;
    sh.getRange(r, col + 1).setValue(checked ? new Date() : '');
    // Delivering an order that was never marked ready marks it ready too.
    if (col === O.DELIVERED && checked && sh.getRange(r, O.READY).getValue() !== true) {
      sh.getRange(r, O.READY).setValue(true);
      sh.getRange(r, O.READY_AT).setValue(new Date());
    }
  }
}

/* ------------------------------------------------------------------ */
/*  Order handling                                                      */
/* ------------------------------------------------------------------ */
function placeOrder_(o) {
  // Honeypot: real customers never fill the hidden "website" field.
  if (o.website) return { ok: true, order_no: 0 };

  const name = String(o.name || '').trim().replace(/\s+/g, ' ').slice(0, 60);
  const phone = String(o.phone || '').trim().slice(0, 25);
  const phoneKey = phone.replace(/\D/g, '').slice(-10);
  const notes = String(o.notes || '').trim().slice(0, 500);
  const clientId = String(o.client_id || '').slice(0, 64);

  if (!name) return reject_('Please enter your name.');
  if (phoneKey.length < 7) return reject_('Please enter a valid phone number.');
  if (!Array.isArray(o.items) || o.items.length === 0) return reject_('The order has no items.');
  if (o.items.length > 40) return reject_('Too many lines in one order.');

  // Validate items against the Products tab (inactive products still accepted for queued orders).
  const catalog = {};
  getProducts_(true).forEach(function (p) { catalog[p.name.toLowerCase()] = p; });
  const items = [];
  for (let i = 0; i < o.items.length; i++) {
    const it = o.items[i] || {};
    const p = catalog[String(it.product || '').trim().toLowerCase()];
    const qty = Number(it.qty);
    if (!p) return reject_('"' + it.product + '" is no longer on the product list.');
    if (!(qty > 0) || qty > CONFIG.MAX_QTY) return reject_('Check the quantity for ' + p.name + '.');
    if (p.unit === 'each' && Math.round(qty) !== qty) return reject_(p.name + ' must be a whole number.');
    items.push({ product: p.name, qty: Math.round(qty * 100) / 100, unit: p.unit });
  }

  const cache = CacheService.getScriptCache();

  // Idempotency: the page retries until it hears back, so the same order may arrive twice.
  if (clientId) {
    const seen = cache.get('cid:' + clientId);
    if (seen) return JSON.parse(seen);
  }

  // Simple per-phone rate limit.
  const rlKey = 'rl:' + phoneKey;
  const count = Number(cache.get(rlKey) || 0);
  if (count >= CONFIG.MAX_ORDERS_PER_PHONE_PER_HOUR) {
    return reject_('Too many orders from this number in the last hour. Please message the shop.');
  }

  const now = new Date();
  const forDate = parseDate_(o.for_date) || new Date(now.getFullYear(), now.getMonth(), now.getDate());

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const orders = ss.getSheetByName(TAB.ORDERS);

    // Second idempotency check against the sheet (cache can be evicted).
    if (clientId) {
      const existing = findByClientId_(orders, clientId);
      if (existing) return existing;
    }

    const customer = upsertCustomer_(ss, name, phone, phoneKey, now);
    const props = PropertiesService.getScriptProperties();
    const orderNo = Number(props.getProperty('NEXT_ORDER') || CONFIG.FIRST_ORDER_NO);
    props.setProperty('NEXT_ORDER', String(orderNo + 1));

    let totalLb = 0, totalPcs = 0;
    items.forEach(function (it) { if (it.unit === 'lb') totalLb += it.qty; else totalPcs += it.qty; });
    const summary = items.map(function (it) { return it.product + ' — ' + fmtQty_(it); }).join('\n');

    const r = orders.getLastRow() + 1;
    orders.getRange(r, 1, 1, ORDER_HEADERS.length).setValues([[
      orderNo, now, customer.no, name, phone, summary,
      totalLb || '', totalPcs || '', forDate, notes,
      false, '', false, '', false,
      '=IF(M' + r + ',"Delivered",IF(K' + r + ',"Ready","New"))',
      clientId
    ]]);

    const lines = ss.getSheetByName(TAB.LINES);
    const lineRows = items.map(function (it) {
      return [orderNo, now, forDate, customer.no, name, it.product, it.qty, it.unit];
    });
    lines.getRange(lines.getLastRow() + 1, 1, lineRows.length, LINE_HEADERS.length).setValues(lineRows);
    SpreadsheetApp.flush();

    const result = {
      ok: true,
      order_no: orderNo,
      customer_no: customer.no,
      customer_is_new: customer.isNew,
      name: name,
      for_date: Utilities.formatDate(forDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      items: items
    };
    if (clientId) cache.put('cid:' + clientId, JSON.stringify(result), 21600);
    cache.put(rlKey, String(count + 1), 3600);

    if (CONFIG.NOTIFY_EMAIL) {
      try {
        MailApp.sendEmail(CONFIG.NOTIFY_EMAIL,
          'New order #' + orderNo + ' — ' + name + ' (' + customer.no + ')',
          summary + (notes ? '\n\nNotes: ' + notes : '') + '\n\nFor: ' + result.for_date);
      } catch (mailErr) { /* email quota hit — the order is still saved */ }
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}

/** Finds the customer by phone, or creates one with the next number (C-001, C-002 …). */
function upsertCustomer_(ss, name, phone, phoneKey, now) {
  const sh = ss.getSheetByName(TAB.CUSTOMERS);
  const last = sh.getLastRow();
  if (last >= 2) {
    const keys = sh.getRange(2, 4, last - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === phoneKey) {
        const row = i + 2;
        const prev = Number(sh.getRange(row, 7).getValue()) || 0;
        sh.getRange(row, 6, 1, 2).setValues([[now, prev + 1]]);
        return { no: String(sh.getRange(row, 1).getValue()), isNew: false };
      }
    }
  }
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty('NEXT_CUSTOMER') || 1);
  props.setProperty('NEXT_CUSTOMER', String(n + 1));
  const no = 'C-' + ('00' + n).slice(-Math.max(3, String(n).length));
  sh.getRange(last + 1, 1, 1, CUSTOMER_HEADERS.length)
    .setValues([[no, name, phone, phoneKey, now, now, 1]]);
  return { no: no, isNew: true };
}

function findByClientId_(orders, clientId) {
  const last = orders.getLastRow();
  if (last < 2) return null;
  const start = Math.max(2, last - 300);
  const ids = orders.getRange(start, O.CLIENT_ID, last - start + 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === clientId) {
      const row = orders.getRange(start + i, 1, 1, O.FOR_DATE).getValues()[0];
      return {
        ok: true, duplicate: true, order_no: row[0], customer_no: String(row[2]), name: row[3],
        for_date: Utilities.formatDate(new Date(row[8]), Session.getScriptTimeZone(), 'yyyy-MM-dd')
      };
    }
  }
  return null;
}

function getProducts_(includeInactive) {
  const sh = SpreadsheetApp.getActive().getSheetByName(TAB.PRODUCTS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, PRODUCT_HEADERS.length).getValues()
    .filter(function (r) { return String(r[0]).trim() && (includeInactive || r[3] === true); })
    .map(function (r) {
      const unit = String(r[1]).trim().toLowerCase() === 'each' ? 'each' : 'lb';
      return { name: String(r[0]).trim(), unit: unit, step: Number(r[2]) || (unit === 'lb' ? 1 : 1), sort: Number(r[4]) || 999 };
    })
    .sort(function (a, b) { return a.sort - b.sort; })
    .map(function (p) { return { name: p.name, unit: p.unit, step: p.step }; });
}

/* ------------------------------------------------------------------ */
/*  Helpers                                                             */
/* ------------------------------------------------------------------ */
function reject_(msg) { return { ok: false, error: msg }; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fmtQty_(it) {
  return it.unit === 'lb' ? it.qty + ' lb' : it.qty + (it.qty === 1 ? ' pc' : ' pcs');
}

function parseDate_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

function ensureSheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#0f3d5c').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

function setWidths_(sh, map) {
  Object.keys(map).forEach(function (c) { sh.setColumnWidth(Number(c), map[c]); });
}

function rule_(formula, range, color) {
  return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula)
    .setBackground(color).setRanges([range]).build();
}
