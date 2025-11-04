/**
 * Orders Tracking Web App (Google Apps Script)
 * - Reads from an existing "All Orders" sheet (produced by tracking/all orders.gs)
 * - Creates an "OrderAttachments" sheet if absent to track uploaded images/files
 * - Stores uploaded files in Drive under a dedicated folder and returns metadata
 * - Serves a simple bilingual UI from index.html
 */

// =====================
// Config
// =====================
// Prefer "All Orders" (from all orders.gs), fallback to "Order" (sheet organizer)
const ORDERS_PRIMARY_NAME = 'All Orders';
const ORDERS_FALLBACK_NAME = 'Order';
const ATTACH_SHEET_NAME = 'OrderAttachments';
const EVENTS_SHEET_NAME = 'OrderEvents';
const SHEET_OUT_NAME = 'Currently Out';
const SHEET_WAREHOUSE_VIEW = 'Warehouse View';
const DRIVE_FOLDER_NAME = 'Orders Attachments';
const SHARE_FILES_PUBLIC = true; // Anyone with link -> Viewer
const MAX_ORDER_COLS = 80; // read at most this many columns from Orders to avoid timeouts
const MAX_DASHBOARD_ROWS = 1500; // cap rows returned to the UI for responsiveness
const MAX_UPLOAD_SIZE_BYTES = 15 * 1024 * 1024; // 15MB
const DEFAULT_PAGE_SIZE = 250; // rows per page
const MAX_FILTER_SCAN_ROWS = 12000; // when date range is set, scan up to this many latest rows

// Attachment header
const ATTACH_HEADER = [
  'OrderId', // string/number, matches reservation ID in All Orders
  'FileId',
  'FileName',
  'MimeType',
  'Url',
  'UploadedBy',
  'UploadedAt',
  'Notes',
  'Direction' // in | out | other
];

const EVENTS_HEADER = [
  'OrderId',
  'Type', // left_warehouse, returned, etc.
  'Notes',
  'CreatedBy',
  'CreatedAt',
  'AttachmentFileId',
  'AttachmentUrl'
];

// =====================
// Web entry
// =====================
function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Orders Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =====================
// Public API
// =====================

/**
 * Returns dashboard data with optional text query filter.
 * @param {{q?: string}} query
 */
function getDashboardData(query) {
  try {
    ensureSheets_();
    const q = (query && query.q ? String(query.q) : '').toLowerCase().trim();
    const fromStr = query && query.from ? String(query.from) : '';
    const toStr = query && query.to ? String(query.to) : '';
    const dateMode = query && query.dateMode ? String(query.dateMode) : 'both'; // both | tarkeeb | sheel
    let page = Math.max(1, Number(query && query.page ? query.page : 1));
    let pageSize = Math.min(MAX_DASHBOARD_ROWS, Math.max(25, Number(query && query.pageSize ? query.pageSize : DEFAULT_PAGE_SIZE)));
    const fromDate = parseYMD_(fromStr);
    const toDate = parseYMD_(toStr, true); // inclusive end-of-day

    const sheet = getOrdersSheet_();
    // If a date range is specified, expand the scan window to include more rows for accurate filtering
    if (fromDate || toDate) {
      const totalRows = Math.max(sheet.getLastRow() - 1, 0);
      page = 1;
      pageSize = Math.min(totalRows, MAX_FILTER_SCAN_ROWS);
    }
    const slim = getSlimOrders_(sheet, page, pageSize);
    const headers = slim.headers; // limited set
    const orderIdKey = headers[0]; // assume first column is reservation/order id

    const attachSheet = getSheet_(ATTACH_SHEET_NAME);
    const attachValues = attachSheet.getLastRow() > 0 ? getUsedValues_(attachSheet, ATTACH_HEADER.length) : [];
    const attachHeaders = attachValues[0] || ATTACH_HEADER;
    const attachIdx = indexByHeader_(attachHeaders);

  // Count attachments per order
    const attachCountByOrder = new Map();
    if (attachValues.length > 1) {
      attachValues.slice(1).forEach(r => {
        const key = String(r[attachIdx['OrderId']]);
        attachCountByOrder.set(key, (attachCountByOrder.get(key) || 0) + 1);
      });
    }

  // Events: detect status (left / returned / pending)
    const eventsSheet = getSheet_(EVENTS_SHEET_NAME);
    const eventsValues = eventsSheet.getLastRow() > 0 ? getUsedValues_(eventsSheet, EVENTS_HEADER.length) : [];
    const eventsHeaders = eventsValues[0] || EVENTS_HEADER;
    const eIdx = indexByHeader_(eventsHeaders);
    /** @type {Record<string, string>} */
    const lastStatus = {};
    if (eventsValues.length > 1) {
      const rowsEv = eventsValues.slice(1).slice().sort((a,b)=>{
        const da = new Date(a[eIdx['CreatedAt']]).getTime() || 0;
        const db = new Date(b[eIdx['CreatedAt']]).getTime() || 0;
        return da - db;
      });
      rowsEv.forEach(r => {
        const t = String(r[eIdx['Type']] || '');
        const id = String(r[eIdx['OrderId']]);
        if (t === EVENT_LEFT) lastStatus[id] = 'left';
        else if (t === EVENT_RETURNED) lastStatus[id] = 'returned';
      });
    }

    let withTarkeeb = 0;
    let withSheel = 0;

    const resultRows = [];
    const preCount = slim.rows.length;
    let afterSearch = 0;
    let afterDate = 0;
    const dtStat = {
      tarkeeb: { present: 0, valid: 0, invalid: 0, invalidSamples: [] },
      sheel: { present: 0, valid: 0, invalid: 0, invalidSamples: [] },
      sources: { tarkeeb: (slim.dateSources && slim.dateSources.tarkeeb) || [], sheel: (slim.dateSources && slim.dateSources.sheel) || [] }
    };

    slim.rows.forEach(obj => {
      let passQ = true;
      if (q) {
        const combined = Object.values(obj).join(' ').toLowerCase();
        passQ = combined.includes(q);
      }
      if (!passQ) return; else afterSearch++;

      // Date diagnostics: count presence/valid for multiple fields before applying filter
      const tRaw = safeGet_(obj, 'Tarkeeb Date');
      const sRaw = safeGet_(obj, 'Sheel Date');
      const pRaw = safeGet_(obj, 'pickup_date_time');
      const dRaw = safeGet_(obj, 'dropoff_date_time');
      const cRaw = safeGet_(obj, 'created_at');
      if (tRaw) { dtStat.tarkeeb.present++; const td2 = normalizeAnyDate_(tRaw); if (td2) dtStat.tarkeeb.valid++; else { dtStat.tarkeeb.invalid++; if (dtStat.tarkeeb.invalidSamples.length < 5) dtStat.tarkeeb.invalidSamples.push(String(tRaw)); } }
      if (sRaw) { dtStat.sheel.present++; const sd2 = normalizeAnyDate_(sRaw); if (sd2) dtStat.sheel.valid++; else { dtStat.sheel.invalid++; if (dtStat.sheel.invalidSamples.length < 5) dtStat.sheel.invalidSamples.push(String(sRaw)); } }
      if (pRaw) { dtStat.pickup.present++; const pd2 = normalizeAnyDate_(pRaw); if (pd2) dtStat.pickup.valid++; else { dtStat.pickup.invalid++; if (dtStat.pickup.invalidSamples.length < 5) dtStat.pickup.invalidSamples.push(String(pRaw)); } }
      if (dRaw) { dtStat.dropoff.present++; const dd2 = normalizeAnyDate_(dRaw); if (dd2) dtStat.dropoff.valid++; else { dtStat.dropoff.invalid++; if (dtStat.dropoff.invalidSamples.length < 5) dtStat.dropoff.invalidSamples.push(String(dRaw)); } }
      if (cRaw) { dtStat.created.present++; const cd2 = normalizeAnyDate_(cRaw); if (cd2) dtStat.created.valid++; else { dtStat.created.invalid++; if (dtStat.created.invalidSamples.length < 5) dtStat.created.invalidSamples.push(String(cRaw)); } }

      // Date range filter: include if any relevant date matches depending on mode
      let passDate = true;
      if (fromDate || toDate) {
        const tDate = normalizeAnyDate_(tRaw) || (pRaw ? normalizeAnyDate_(pRaw) : null);
        const sDate = normalizeAnyDate_(sRaw) || (dRaw ? normalizeAnyDate_(dRaw) : null);
        const cDate = normalizeAnyDate_(cRaw);
        const okT = tDate ? inRange_(tDate, fromDate, toDate) : false;
        const okS = sDate ? inRange_(sDate, fromDate, toDate) : false;
        const okC = cDate ? inRange_(cDate, fromDate, toDate) : false;
        if (dateMode === 'tarkeeb') passDate = okT;
        else if (dateMode === 'sheel') passDate = okS;
        else passDate = (okT || okS || okC); // both => any
      }
      if (!passDate) return; else afterDate++;

      if (tRaw) withTarkeeb++;
      if (sRaw) withSheel++;

      const orderId = String(obj[orderIdKey]);
      obj['_attachmentCount'] = attachCountByOrder.get(orderId) || 0;
      const st = lastStatus[orderId] || 'pending';
      obj['_status'] = st; // left | returned | pending
      resultRows.push(obj);
    });

    const preferredDisplay = [orderIdKey, 'code', 'name', 'city_id', 'phone_number', 'payment_method', 'grand_total', 'Tarkeeb Date', 'Sheel Date', 'Category Names', 'Warehouse Names'];
    const missingColumns = preferredDisplay.filter(h => headers.indexOf(h) === -1);

    return {
      headers,
      rows: resultRows,
      counts: {
        total: resultRows.length,
        withTarkeeb,
        withSheel
      },
      meta: {
        sheetName: sheet.getName(),
        totalRows: Math.max(sheet.getLastRow() - 1, 0),
        debug: {
          idHeader: headers[0],
          chosenColumns: headers,
          sheetHeaders: slim.sheetHeaders || [],
          foundColumns: slim.foundColumns || [],
          sampleIds: slim.rows.slice(0, 5).map(r => String(r[headers[0]])),
          counts: { pre: preCount, afterSearch: afterSearch, afterDate: afterDate },
          filters: { q, from: fromStr, to: toStr, dateMode },
          missingColumns,
          dateDiagnostics: dtStat,
          window: slim.window || null
        },
        paging: { page, pageSize }
      }
    };
  } catch (e) {
    return { headers: [], rows: [], counts: { total: 0, withTarkeeb: 0, withSheel: 0 }, meta: {}, error: String(e && e.message ? e.message : e) };
  }
}

// Reads minimal columns needed for dashboard to reduce I/O
function getSlimOrders_(sheet, page, pageSize) {
  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return { headers: [], rows: [] };
  // Read header only (cheap even if wide)
  const headerRow = sheet.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h).trim());
  const idHeader = headerRow[0];
  const desired = [
    idHeader,
    // Common business columns from All Orders
    'code','name','address','comments','phone_number','email','payment_method','sub_total','grand_total','created_at','updated_at',
    // City variations
    'City', 'city', 'city_id',
    // Enriched columns
    'Category Names', 'Category Name',
    'Warehouse Names', 'Warehouse Name',
    'Tarkeeb Date', 'Sheel Date',
    'Product Images', 'Product Image',
    // Fallback raw datetime columns
    'pickup_date_time', 'dropoff_date_time'
  ];
  // Map found headers to column indices
  const found = [];
  const colIndexByHeader = {};
  desired.forEach(name => {
    const i = headerRow.indexOf(name);
    if (i !== -1 && colIndexByHeader[name] == null) {
      colIndexByHeader[name] = i;
      found.push(name);
    }
  });
  // Ensure at least ID exists
  if (found.indexOf(idHeader) === -1) { found.unshift(idHeader); colIndexByHeader[idHeader] = 0; }

  const maxRows = Math.min(MAX_DASHBOARD_ROWS, Math.max(1, pageSize || DEFAULT_PAGE_SIZE));
  const endRow = lr - (page - 1) * maxRows;
  if (endRow < 2) return { headers: [idHeader], rows: [] };
  const startRow = Math.max(2, endRow - maxRows + 1);
  const actualCount = (endRow - startRow + 1);

  // Fetch each found column individually
  const colValuesByName = {};
  found.forEach(name => {
    const col = colIndexByHeader[name] + 1;
    colValuesByName[name] = sheet.getRange(startRow, col, actualCount, 1).getValues();
  });

  // Build objects per row
  const rows = [];
  const dateSources = { tarkeeb: [], sheel: [] };
  for (let i = 0; i < actualCount; i++) {
    const obj = {};
    found.forEach(name => { obj[name] = colValuesByName[name][i][0]; });
    // Normalize preferred fields
    if (obj['Category Name'] && !obj['Category Names']) obj['Category Names'] = obj['Category Name'];
    if (obj['Warehouse Name'] && !obj['Warehouse Names']) obj['Warehouse Names'] = obj['Warehouse Name'];
    if (obj['Product Image'] && !obj['Product Images']) obj['Product Images'] = obj['Product Image'];
    if ((obj['city'] || obj['city_id']) && !obj['City']) obj['City'] = obj['city'] || obj['city_id'];
    // Derive Tarkeeb/Sheel dates if missing using pickup/dropoff
    if (!obj['Tarkeeb Date'] && colIndexByHeader['pickup_date_time'] != null) {
      const raw = colValuesByName['pickup_date_time'] ? colValuesByName['pickup_date_time'][i][0] : '';
      const parts = String(raw || '').split(/[ T]/);
      if (parts.length >= 2) { obj['Tarkeeb Date'] = parts[0]; dateSources.tarkeeb.push('pickup_date_time'); }
    } else if (obj['Tarkeeb Date']) {
      dateSources.tarkeeb.push('Tarkeeb Date');
    }
    if (!obj['Sheel Date'] && colIndexByHeader['dropoff_date_time'] != null) {
      const raw = colValuesByName['dropoff_date_time'] ? colValuesByName['dropoff_date_time'][i][0] : '';
      const parts = String(raw || '').split(/[ T]/);
      if (parts.length >= 2) { obj['Sheel Date'] = parts[0]; dateSources.sheel.push('dropoff_date_time'); }
    } else if (obj['Sheel Date']) {
      dateSources.sheel.push('Sheel Date');
    }
    rows.push(obj);
  }

  // Compose headers to return (minimal set for UI)
  const preferredDisplay = [
    idHeader, 'code', 'name', 'city_id', 'phone_number', 'payment_method', 'grand_total',
    'Tarkeeb Date', 'Sheel Date', 'Category Names', 'Warehouse Names'
  ];
  const retHeaders = preferredDisplay.filter(h => (h === idHeader) || found.indexOf(h) !== -1);
  return { headers: retHeaders, rows, sheetHeaders: headerRow, foundColumns: found, dateSources, window: { startRow, count: actualCount } };
}

// Parse YYYY-MM-DD to date; when endOfDay true, set 23:59:59
function parseYMD_(s, endOfDay) {
  if (!s) return null;
  const m = /^\d{4}-\d{2}-\d{2}$/.test(s) ? s.split('-') : null;
  if (!m) return null;
  const d = new Date(Number(m[0]), Number(m[1]) - 1, Number(m[2]), endOfDay ? 23 : 0, endOfDay ? 59 : 0, endOfDay ? 59 : 0, endOfDay ? 999 : 0);
  return isNaN(d.getTime()) ? null : d;
}

// Normalize GS date or string to Date
function normalizeAnyDate_(val) {
  if (!val) return null;
  if (Object.prototype.toString.call(val) === '[object Date]') return val;
  const s = String(val).trim();
  if (!s) return null;
  // Accept:
  // - YYYY-MM-DD
  // - YYYY-MM-DD HH:mm[:ss]
  // - ISO strings
  let t = s;
  const ymd = /^\d{4}-\d{2}-\d{2}$/;
  const ymdTime = /^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(?::\d{2})?$/;
  if (ymd.test(s)) {
    t = s + 'T00:00:00';
  } else if (ymdTime.test(s)) {
    t = s.replace(' ', 'T');
  }
  const d = new Date(t);
  return isNaN(d.getTime()) ? null : d;
}

function inRange_(d, fromD, toD) {
  const t = d.getTime();
  if (fromD && t < fromD.getTime()) return false;
  if (toD && t > toD.getTime()) return false;
  return true;
}

/**
 * Returns attachments for a given order id.
 * @param {string|number} orderId
 */
function getAttachments(orderId) {
  ensureSheets_();
  const sheet = getSheet_(ATTACH_SHEET_NAME);
  if (sheet.getLastRow() === 0) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = indexByHeader_(headers);
  const id = String(orderId);

  return values.slice(1)
    .filter(r => String(r[idx['OrderId']]) === id)
    .map(r => rowToObject_(headers, r));
}

/**
 * Log that order left the warehouse. Optionally accepts a file (image/pdf) and notes.
 * payload = { orderId, notes?, file? }
 * Returns { orderId, eventType, createdAt, attachmentUrl? }
 */
function markLeftWarehouse(payload) {
  ensureSheets_();
  if (!payload) throw new Error('Missing payload');
  const orderId = String(payload.orderId || '').trim();
  if (!orderId) throw new Error('Missing orderId');
  const notes = String(payload.notes || '').trim();

  let attachRes = null;
  if (payload.file) {
    // Reuse uploadAttachment logic (direction = 'out')
    attachRes = uploadAttachment({ orderId, direction: 'out', notes, file: payload.file });
  }

  const events = getSheet_(EVENTS_SHEET_NAME);
  const now = new Date();
  const row = [
    orderId,
    EVENT_LEFT,
    notes,
    Session.getActiveUser().getEmail() || 'unknown',
    now,
    attachRes ? attachRes.fileId : '',
    attachRes ? attachRes.url : ''
  ];
  events.appendRow(row);

  regenerateSummarySheets_();
  return { orderId, eventType: EVENT_LEFT, createdAt: now, attachmentUrl: attachRes ? attachRes.url : '' };
}

/**
 * Mark order returned to warehouse. Optional attachment & notes.
 */
function markReturnedWarehouse(payload) {
  ensureSheets_();
  if (!payload) throw new Error('Missing payload');
  const orderId = String(payload.orderId || '').trim();
  if (!orderId) throw new Error('Missing orderId');
  const notes = String(payload.notes || '').trim();

  let attachRes = null;
  if (payload.file) {
    // direction = 'in'
    attachRes = uploadAttachment({ orderId, direction: 'in', notes, file: payload.file });
  }

  const events = getSheet_(EVENTS_SHEET_NAME);
  const now = new Date();
  const row = [
    orderId,
    EVENT_RETURNED,
    notes,
    Session.getActiveUser().getEmail() || 'unknown',
    now,
    attachRes ? attachRes.fileId : '',
    attachRes ? attachRes.url : ''
  ];
  events.appendRow(row);

  regenerateSummarySheets_();
  return { orderId, eventType: EVENT_RETURNED, createdAt: now, attachmentUrl: attachRes ? attachRes.url : '' };
}

/**
 * Uploads an attachment and logs it in the attachments sheet.
 * payload = { orderId, direction, notes, file: { name, mimeType, data } }
 * data must be base64 string (no data URL prefix) or a data URL.
 */
function uploadAttachment(payload) {
  ensureSheets_();
  const norm = normalizeUploadInput_(payload);

  if (norm.fileBytes.length > MAX_UPLOAD_SIZE_BYTES) {
    throw new Error('File too large');
  }

  const folder = getOrCreateFolder_();
  const orderFolder = getOrCreateChildFolder_(folder, `Order-${norm.orderId}`);
  const blob = Utilities.newBlob(norm.fileBytes, norm.file.mimeType, norm.file.name);
  const file = orderFolder.createFile(blob);

  if (SHARE_FILES_PUBLIC) {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }

  const url = sanitizeUrl_(file.getUrl());

  // Append to attachments sheet
  const sheet = getSheet_(ATTACH_SHEET_NAME);
  const now = new Date();
  const row = [
    String(norm.orderId),
    file.getId(),
    norm.file.name,
    norm.file.mimeType,
    url,
    Session.getActiveUser().getEmail() || 'unknown',
    now,
    norm.notes || '',
    norm.direction || ''
  ];

  sheet.appendRow(row);

  return {
    orderId: String(norm.orderId),
    fileId: file.getId(),
    fileName: norm.file.name,
    mimeType: norm.file.mimeType,
    url,
    uploadedAt: now,
    notes: norm.notes || '',
    direction: norm.direction || ''
  };
}

function getConfig() {
  return {
    maxUploadSizeBytes: MAX_UPLOAD_SIZE_BYTES,
    shareFilesPublic: SHARE_FILES_PUBLIC,
    maxDashboardRows: MAX_DASHBOARD_ROWS,
    defaultPageSize: DEFAULT_PAGE_SIZE
  };
}

// =====================
// Helpers
// =====================

function ensureSheets_() {
  // Ensure orders sheet exists (created by external script); if missing, create header placeholder
  const ss = SpreadsheetApp.getActive();
  let orders = ss.getSheetByName(ORDERS_PRIMARY_NAME) || ss.getSheetByName(ORDERS_FALLBACK_NAME);
  if (!orders) { orders = ss.insertSheet(ORDERS_PRIMARY_NAME); orders.getRange(1, 1, 1, 1).setValues([["id"]]); }

  // Ensure attachments sheet exists
  let attach = ss.getSheetByName(ATTACH_SHEET_NAME);
  if (!attach) {
    attach = ss.insertSheet(ATTACH_SHEET_NAME);
    attach.getRange(1, 1, 1, ATTACH_HEADER.length).setValues([ATTACH_HEADER]);
    attach.getRange(2, 1).setNumberFormat('@'); // OrderId as text-friendly
  } else if (attach.getLastRow() === 0) {
    attach.getRange(1, 1, 1, ATTACH_HEADER.length).setValues([ATTACH_HEADER]);
  }

  // Ensure events sheet exists
  let events = ss.getSheetByName(EVENTS_SHEET_NAME);
  if (!events) {
    events = ss.insertSheet(EVENTS_SHEET_NAME);
    events.getRange(1, 1, 1, EVENTS_HEADER.length).setValues([EVENTS_HEADER]);
    events.getRange(2, 1).setNumberFormat('@');
  } else if (events.getLastRow() === 0) {
    events.getRange(1, 1, 1, EVENTS_HEADER.length).setValues([EVENTS_HEADER]);
  }

  // Ensure summary tabs exist
  if (!ss.getSheetByName(SHEET_OUT_NAME)) ss.insertSheet(SHEET_OUT_NAME);
  if (!ss.getSheetByName(SHEET_WAREHOUSE_VIEW)) ss.insertSheet(SHEET_WAREHOUSE_VIEW);
}

function getOrdersSheet_() {
  const ss = SpreadsheetApp.getActive();
  const primary = ss.getSheetByName(ORDERS_PRIMARY_NAME);
  const fallback = ss.getSheetByName(ORDERS_FALLBACK_NAME);
  // Prefer a sheet that actually has data rows (>=2: header + at least one row)
  if (primary && primary.getLastRow() >= 2) return primary;
  if (fallback && fallback.getLastRow() >= 2) return fallback;
  // If primary exists (even empty), but fallback has no data as well, return whichever exists
  return primary || fallback;
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Missing sheet: ${name}`);
  return sh;
}

function indexByHeader_(headers) {
  const idx = {};
  headers.forEach((h, i) => idx[String(h)] = i);
  return idx;
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => obj[String(h)] = row[i]);
  return obj;
}

function safeGet_(obj, key) {
  return Object.prototype.hasOwnProperty.call(obj, key) ? obj[key] : '';
}

function getOrCreateFolder_() {
  const existing = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (existing.hasNext()) return existing.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function getOrCreateChildFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function sanitizeUrl_(url) {
  try {
    return String(url);
  } catch (e) {
    return String(url);
  }
}

function normalizeUploadInput_(payload) {
  if (!payload) throw new Error('Missing payload');
  const orderId = String(payload.orderId || '').trim();
  if (!orderId) throw new Error('Missing orderId');
  const file = payload.file || {};
  if (!file.name || !file.mimeType || !file.data) {
    throw new Error('Invalid file payload');
  }

  // data may be data URL or base64 string
  let base64 = String(file.data);
  const commaIdx = base64.indexOf(',');
  if (base64.startsWith('data:') && commaIdx > -1) {
    base64 = base64.substring(commaIdx + 1);
  }
  const bytes = Utilities.base64Decode(base64);

  return {
    orderId,
    direction: (payload.direction || '').trim(),
    notes: (payload.notes || '').trim(),
    file: { name: String(file.name), mimeType: String(file.mimeType) },
    fileBytes: bytes
  };
}

// =====================
// Summary sheets generation
// =====================
function regenerateSummarySheets_() {
  const ss = SpreadsheetApp.getActive();
  const ordersSh = getOrdersSheet_();
  if (!ordersSh) return;
  const ordersVals = getUsedValues_(ordersSh, MAX_ORDER_COLS);
  if (!ordersVals.length) return;

  const headers = ordersVals[0].map(String);
  const rows = ordersVals.slice(1);
  const idx = indexByHeader_(headers);
  const idKey = headers[0];

  // Map events to status and get latest out attachment url
  const evSh = getSheet_(EVENTS_SHEET_NAME);
  const evVals = evSh.getLastRow() > 0 ? getUsedValues_(evSh, EVENTS_HEADER.length) : [];
  const evHeaders = evVals[0] || EVENTS_HEADER;
  const eIdx = indexByHeader_(evHeaders);
  /** @type {Record<string, {status: string, lastOutAt: Date|null}>} */
  const statusById = {};
  if (evVals.length > 1) {
    const sorted = evVals.slice(1).slice().sort((a,b)=>{
      const da = new Date(a[eIdx['CreatedAt']]).getTime()||0;
      const db = new Date(b[eIdx['CreatedAt']]).getTime()||0;
      return da - db;
    });
    sorted.forEach(r => {
      const id = String(r[eIdx['OrderId']]);
      const type = String(r[eIdx['Type']]);
      if (!statusById[id]) statusById[id] = { status: 'pending', lastOutAt: null };
      if (type === EVENT_LEFT) { statusById[id].status = 'left'; statusById[id].lastOutAt = new Date(r[eIdx['CreatedAt']]); }
      else if (type === EVENT_RETURNED) { statusById[id].status = 'returned'; }
    });
  }

  // Build OUT sheet data
  const outHeaders = ['OrderId', 'City', 'Category', 'Warehouse', 'Return By (Sheel Date)', 'Image', 'Status'];
  const outRows = [];
  rows.forEach(r => {
    const obj = rowToObject_(headers, r);
    const id = String(obj[idKey]);
    const status = (statusById[id] && statusById[id].status) || 'pending';
    if (status !== 'left') return;
    const city = obj['City'] || obj['city'] || obj['city_id'] || '';
    const category = obj['Category Names'] || obj['Category Name'] || '';
    const warehouse = obj['Warehouse Names'] || obj['Warehouse Name'] || '';
    const returnBy = obj['Sheel Date'] || '';
    const image = obj['Product Images'] || obj['Product Image'] || '';
    outRows.push([id, city, category, warehouse, returnBy, image, 'Left']);
  });

  const outSh = getSheet_(SHEET_OUT_NAME);
  outSh.clearContents();
  outSh.getRange(1,1,1,outHeaders.length).setValues([outHeaders]);
  if (outRows.length) outSh.getRange(2,1,outRows.length,outHeaders.length).setValues(outRows);

  // Build WAREHOUSE VIEW sheet data (returned or pending)
  const whHeaders = ['OrderId', 'City', 'Category', 'Warehouse', 'Tarkeeb Date', 'Sheel Date', 'Image', 'Status'];
  const whRows = [];
  rows.forEach(r => {
    const obj = rowToObject_(headers, r);
    const id = String(obj[idKey]);
    const status = (statusById[id] && statusById[id].status) || 'pending';
    if (status === 'left') return; // skip out
    const city = obj['City'] || obj['city'] || obj['city_id'] || '';
    const category = obj['Category Names'] || obj['Category Name'] || '';
    const warehouse = obj['Warehouse Names'] || obj['Warehouse Name'] || '';
    const tarkeeb = obj['Tarkeeb Date'] || '';
    const sheel = obj['Sheel Date'] || '';
    const image = obj['Product Images'] || obj['Product Image'] || '';
    const st = status === 'returned' ? 'Returned' : 'Pending';
    whRows.push([id, city, category, warehouse, tarkeeb, sheel, image, st]);
  });

  const whSh = getSheet_(SHEET_WAREHOUSE_VIEW);
  whSh.clearContents();
  whSh.getRange(1,1,1,whHeaders.length).setValues([whHeaders]);
  if (whRows.length) whSh.getRange(2,1,whRows.length,whHeaders.length).setValues(whRows);
}

// Read only used range (faster than getDataRange on formatted sheets)
function getUsedValues_(sh, maxColsOptional) {
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr === 0 || lc === 0) return [];
  const width = maxColsOptional ? Math.min(lc, maxColsOptional) : lc;
  return sh.getRange(1, 1, lr, width).getValues();
}
const EVENT_LEFT = 'left_warehouse';
const EVENT_RETURNED = 'returned_warehouse';
