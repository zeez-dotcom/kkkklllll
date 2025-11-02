/**********************
 * CONFIG
 **********************/
const SALES_SPREADSHEET_ID = '';
const SALES_SHEET_NAME = 'SalesRecords';
const SALES_FOLDER_NAME = 'Sales Receipts';
let SALES_FOLDER_ID = '';
const SHARE_RECEIPTS_PUBLIC = false;
const MAX_RECEIPT_SIZE_BYTES = 5 * 1024 * 1024;
const SALES_HEADER = [
  'id',
  'reportDate',
  'totalSales',
  'knetSales',
  'cashSales',
  'expenses',
  'expenseNotes',
  'notes',
  'receiptUrl',
  'receiptName',
  'knetExpectedDate',
  'knetReceivedDate',
  'knetStatus',
  'createdAt',
  'updatedAt'
];
const DEFAULT_KNET_DELAY_DAYS = 10;

/**********************
 * WEB APP ENTRY
 **********************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sales & Expenses Dashboard');
}

/**********************
 * PUBLIC API
 **********************/
function getSalesDashboardData(query) {
  try {
    const filters = sanitizeSalesFilters_(query || {});
    const sheet = getSalesSheet_();
    const lastRow = sheet.getLastRow();
    const headerMap = buildHeaderIndex_(SALES_HEADER);
    if (lastRow < 2) {
      return {
        ok: true,
        summary: emptySummary_(),
        rows: [],
        filters,
        generatedAt: toIsoTimestamp_(new Date())
      };
    }

    const values = sheet.getRange(2, 1, lastRow - 1, SALES_HEADER.length).getValues();
    const summary = emptySummary_();
    const rows = [];
    for (let i = 0; i < values.length; i++) {
      const record = mapSalesRowToObject_(values[i], headerMap);
      const enriched = finalizeSalesRecord_(record);
      if (filterSalesRecord_(enriched, filters)) {
        rows.push(enriched);
      }
      accumulateSummary_(summary, enriched);
    }

    rows.sort((a, b) => compareRecords_(a, b));

    return {
      ok: true,
      summary: summary,
      rows: rows,
      filters,
      generatedAt: toIsoTimestamp_(new Date())
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function recordSalesEntry(payload) {
  try {
    const sheet = getSalesSheet_();
    const headerMap = buildHeaderIndex_(SALES_HEADER);
    const raw = payload || {};
    const sanitized = sanitizeSalesPayload_(raw);
    const receipt = sanitizeReceiptInput_(raw.receipt);
    const now = new Date();
    const nowIso = toIsoTimestamp_(now);
    let record = Object.assign({}, sanitized);
    let rowNumber = -1;
    let existing = null;

    if (record.id) {
      rowNumber = findSalesRowById_(sheet, record.id, headerMap.id);
    }

    if (rowNumber < 0) {
      record.id = generateId_();
      record.createdAt = nowIso;
      record.receiptUrl = record.receiptUrl || '';
      record.receiptName = record.receiptName || '';
    } else {
      const existingRow = sheet.getRange(rowNumber, 1, 1, SALES_HEADER.length).getValues()[0];
      existing = mapSalesRowToObject_(existingRow, headerMap);
      record.createdAt = existing.createdAt || nowIso;
      record.receiptUrl = record.receiptUrl || existing.receiptUrl || '';
      record.receiptName = record.receiptName || existing.receiptName || '';
    }

    if (receipt) {
      const stored = storeReceipt_(receipt);
      record.receiptUrl = stored.url;
      record.receiptName = stored.name;
    }

    record.updatedAt = nowIso;
    const knetInfo = computeKnetStatus_(record);
    record.knetStatus = knetInfo.knetStatus;

    const values = buildRowValues_(record, headerMap);
    if (rowNumber < 0) {
      sheet.appendRow(values);
    } else {
      sheet.getRange(rowNumber, 1, 1, SALES_HEADER.length).setValues([values]);
    }

    return {
      ok: true,
      record: finalizeSalesRecord_(record)
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function markKnetReceived(input) {
  try {
    const data = input || {};
    const id = sanitizeString_(data.id);
    if (!id) {
      throw new Error('Record ID is required.');
    }

    const sheet = getSalesSheet_();
    const headerMap = buildHeaderIndex_(SALES_HEADER);
    const rowNumber = findSalesRowById_(sheet, id, headerMap.id);
    if (rowNumber < 0) {
      throw new Error('Unable to locate the requested record.');
    }

    const values = sheet.getRange(rowNumber, 1, 1, SALES_HEADER.length).getValues()[0];
    const record = mapSalesRowToObject_(values, headerMap);
    const receivedDate = toIsoDate_(data.receivedDate) || todayIso_();
    record.knetReceivedDate = receivedDate;
    record.knetStatus = 'Received';
    record.updatedAt = toIsoTimestamp_(new Date());
    const knetInfo = computeKnetStatus_(record);
    record.knetStatus = knetInfo.knetStatus;

    const rowValues = buildRowValues_(record, headerMap);
    sheet.getRange(rowNumber, 1, 1, SALES_HEADER.length).setValues([rowValues]);

    return {
      ok: true,
      record: finalizeSalesRecord_(record)
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

/**********************
 * HELPERS
 **********************/
function getSalesSheet_() {
  const configuredId = typeof SALES_SPREADSHEET_ID === 'string' ? SALES_SPREADSHEET_ID.trim() : '';
  let ss = null;
  if (configuredId) {
    try {
      ss = SpreadsheetApp.openById(configuredId);
    } catch (err) {
      throw new Error('Unable to open configured sales spreadsheet.');
    }
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) {
    throw new Error('Sales spreadsheet unavailable.');
  }

  let sheet = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SALES_SHEET_NAME);
  }
  ensureHeader_(sheet, SALES_HEADER);
  return sheet;
}

function ensureHeader_(sheet, header) {
  const first = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  const mismatch = first.length !== header.length || header.some(function (name, index) {
    return first[index] !== name;
  });
  if (mismatch) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
}

function getReceiptFolder_() {
  if (SALES_FOLDER_ID) {
    try {
      return DriveApp.getFolderById(SALES_FOLDER_ID);
    } catch (err) {
      SALES_FOLDER_ID = '';
    }
  }
  const targetName = SALES_FOLDER_NAME || 'Sales Receipts';
  const existing = DriveApp.getFoldersByName(targetName);
  const folder = existing.hasNext() ? existing.next() : DriveApp.createFolder(targetName);
  SALES_FOLDER_ID = folder.getId();
  return folder;
}

function sanitizeString_(value) {
  return String(value == null ? '' : value).trim();
}

function sanitizeUrl_(url) {
  const value = sanitizeString_(url);
  if (!value) return '';
  return /^https?:\/\//i.test(value) ? value : '';
}

function sanitizeReceiptInput_(raw) {
  if (!raw) return null;
  const obj = typeof raw === 'object' ? raw : {};
  const name = sanitizeString_(obj.name);
  const b64 = sanitizeString_(obj.b64);
  if (!name || !b64) {
    return null;
  }
  const size = Number(obj.size);
  if (MAX_RECEIPT_SIZE_BYTES && isFinite(size) && size > MAX_RECEIPT_SIZE_BYTES) {
    throw new Error('Receipt image exceeds maximum size limit.');
  }
  const type = sanitizeString_(obj.type) || 'image/jpeg';
  return {
    name: name,
    b64: b64,
    type: type,
    size: isFinite(size) ? size : 0
  };
}

function decodeBase64Safely_(encoded, context, maxBytes) {
  const safeContext = context || 'upload';
  const raw = sanitizeString_(encoded);
  if (!raw) {
    throw new Error('Missing encoded data for ' + safeContext + '.');
  }
  const normalized = raw.indexOf(',') >= 0 ? raw.split(',').pop() : raw;
  const cleaned = normalized.replace(/\s/g, '');
  let bytes;
  try {
    bytes = Utilities.base64Decode(cleaned);
  } catch (err) {
    throw new Error('Unable to decode file for ' + safeContext + '.');
  }
  if (maxBytes && bytes.length > maxBytes) {
    throw new Error('Receipt image exceeds maximum size limit.');
  }
  return bytes;
}

function storeReceipt_(receipt) {
  const folder = getReceiptFolder_();
  const bytes = decodeBase64Safely_(receipt.b64, 'receiptUpload', MAX_RECEIPT_SIZE_BYTES);
  const mimeType = receipt.type || 'image/jpeg';
  const blob = Utilities.newBlob(bytes, mimeType, receipt.name);
  const created = folder.createFile(blob).setName(receipt.name);
  if (SHARE_RECEIPTS_PUBLIC) {
    created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return {
    url: sanitizeUrl_(created.getUrl()),
    name: receipt.name
  };
}

function makeDrivePreviewUrl_(url) {
  const safe = sanitizeUrl_(url);
  if (!safe) return '';
  const idMatch = safe.match(/\/d\/([a-zA-Z0-9_-]+)/);
  const id = idMatch && idMatch[1] ? idMatch[1] : (function () {
    const paramMatch = safe.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    return paramMatch && paramMatch[1] ? paramMatch[1] : '';
  })();
  if (!id) return safe;
  return 'https://drive.google.com/uc?export=view&id=' + id;
}

function toIsoDate_(value, fallback) {
  if (value == null || value === '') {
    return fallback || '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value)) return fallback || '';
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const str = sanitizeString_(value);
  if (!str) return fallback || '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return str;
  }
  const parsed = new Date(str);
  if (isNaN(parsed)) {
    return fallback || '';
  }
  return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function todayIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function toIsoTimestamp_(date) {
  const target = Object.prototype.toString.call(date) === '[object Date]' ? date : new Date(date);
  if (isNaN(target)) {
    return '';
  }
  return Utilities.formatDate(target, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function parseMoney_(value) {
  if (value == null || value === '') return 0;
  if (typeof value === 'number' && isFinite(value)) {
    return round2_(value);
  }
  const cleaned = String(value).replace(/[^0-9.\-]/g, '');
  const num = Number(cleaned);
  return isFinite(num) ? round2_(num) : 0;
}

function round2_(value) {
  return Math.round(Number(value || 0) * 100) / 100;
}

function addDaysIso_(iso, days) {
  const base = dateFromIso_(iso);
  if (!base) return '';
  base.setDate(base.getDate() + Number(days || 0));
  return Utilities.formatDate(base, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function dateFromIso_(iso) {
  if (!iso) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
    const parts = iso.split('-').map(function (p) { return Number(p); });
    const date = new Date(parts[0], parts[1] - 1, parts[2]);
    return isNaN(date) ? null : date;
  }
  const parsed = new Date(iso);
  return isNaN(parsed) ? null : parsed;
}

function generateId_() {
  const random = Math.random().toString(36).slice(2, 8);
  return Date.now().toString(36) + random;
}

function buildHeaderIndex_(header) {
  const map = {};
  for (let i = 0; i < header.length; i++) {
    map[header[i]] = i;
  }
  return map;
}

function mapSalesRowToObject_(row, headerMap) {
  const record = {};
  SALES_HEADER.forEach(function (key, index) {
    record[key] = row[index];
  });
  record.id = sanitizeString_(record.id);
  record.reportDate = toIsoDate_(record.reportDate);
  record.totalSales = parseMoney_(record.totalSales);
  record.knetSales = parseMoney_(record.knetSales);
  record.cashSales = parseMoney_(record.cashSales);
  record.expenses = parseMoney_(record.expenses);
  record.expenseNotes = sanitizeString_(record.expenseNotes);
  record.notes = sanitizeString_(record.notes);
  record.receiptUrl = sanitizeUrl_(record.receiptUrl);
  record.receiptName = sanitizeString_(record.receiptName);
  record.knetExpectedDate = toIsoDate_(record.knetExpectedDate);
  record.knetReceivedDate = toIsoDate_(record.knetReceivedDate);
  record.knetStatus = sanitizeString_(record.knetStatus);
  record.createdAt = record.createdAt ? toIsoTimestamp_(record.createdAt) : '';
  record.updatedAt = record.updatedAt ? toIsoTimestamp_(record.updatedAt) : '';
  return record;
}

function finalizeSalesRecord_(record) {
  const normalized = Object.assign({}, record);
  const knetInfo = computeKnetStatus_(normalized);
  normalized.knetStatus = knetInfo.knetStatus;
  normalized.knetDueDays = knetInfo.knetDueDays;
  normalized.knetDelayDays = knetInfo.knetDelayDays;
  normalized.knetIsOverdue = knetInfo.knetIsOverdue;
  normalized.netSales = round2_(normalized.totalSales - normalized.expenses);
  normalized.knetPendingAmount = normalized.knetStatus === 'Pending' ? normalized.knetSales : 0;
  normalized.knetReceivedDelayDays = knetInfo.knetDelayDays;
  normalized.receiptUrl = sanitizeUrl_(normalized.receiptUrl);
  normalized.receiptName = sanitizeString_(normalized.receiptName || '');
  normalized.receiptPreviewUrl = makeDrivePreviewUrl_(normalized.receiptUrl);
  normalized.hasReceipt = !!normalized.receiptUrl;
  return normalized;
}

function computeKnetStatus_(record) {
  const today = dateFromIso_(todayIso_());
  const expected = dateFromIso_(record.knetExpectedDate);
  const received = dateFromIso_(record.knetReceivedDate);
  const hasKnet = parseMoney_(record.knetSales) > 0;

  let status = sanitizeString_(record.knetStatus);
  if (!hasKnet) {
    status = '';
  } else if (received) {
    status = 'Received';
  } else if (!status) {
    status = 'Pending';
  }

  let dueDays = null;
  let isOverdue = false;
  if (status === 'Pending' && expected && today) {
    const oneDay = 24 * 60 * 60 * 1000;
    dueDays = Math.round((expected.getTime() - today.getTime()) / oneDay);
    isOverdue = dueDays < 0;
  }

  let delayDays = null;
  if (status === 'Received' && expected && received) {
    const oneDay = 24 * 60 * 60 * 1000;
    delayDays = Math.round((received.getTime() - expected.getTime()) / oneDay);
  } else if (status === 'Received' && record.reportDate && received) {
    const base = dateFromIso_(record.reportDate);
    if (base) {
      const oneDay = 24 * 60 * 60 * 1000;
      delayDays = Math.round((received.getTime() - base.getTime()) / oneDay);
    }
  }

  return {
    knetStatus: status,
    knetDueDays: dueDays,
    knetDelayDays: delayDays,
    knetIsOverdue: isOverdue
  };
}

function sanitizeSalesPayload_(raw) {
  const payload = raw || {};
  const reportDate = toIsoDate_(payload.reportDate) || todayIso_();
  const totalSales = parseMoney_(payload.totalSales);
  const knetSales = parseMoney_(payload.knetSales);
  const cashSales = parseMoney_(payload.cashSales);
  const expenses = parseMoney_(payload.expenses);
  let knetExpectedDate = toIsoDate_(payload.knetExpectedDate);
  let knetReceivedDate = toIsoDate_(payload.knetReceivedDate);
  const markReceived = !!payload.markReceived;

  if (!knetExpectedDate && knetSales > 0) {
    knetExpectedDate = addDaysIso_(reportDate, DEFAULT_KNET_DELAY_DAYS);
  }
  if (markReceived && !knetReceivedDate && knetSales > 0) {
    knetReceivedDate = todayIso_();
  }

  let knetStatus = sanitizeString_(payload.knetStatus);
  if (knetSales <= 0) {
    knetStatus = '';
    knetExpectedDate = '';
    knetReceivedDate = '';
  } else if (knetReceivedDate) {
    knetStatus = 'Received';
  } else if (!knetStatus) {
    knetStatus = markReceived ? 'Received' : 'Pending';
  }

  return {
    id: sanitizeString_(payload.id),
    reportDate: reportDate,
    totalSales: totalSales,
    knetSales: knetSales,
    cashSales: cashSales,
    expenses: expenses,
    expenseNotes: sanitizeString_(payload.expenseNotes),
    notes: sanitizeString_(payload.notes),
    receiptUrl: sanitizeUrl_(payload.receiptUrl),
    receiptName: sanitizeString_(payload.receiptName),
    knetExpectedDate: knetExpectedDate,
    knetReceivedDate: knetReceivedDate,
    knetStatus: knetStatus
  };
}

function buildRowValues_(record, headerMap) {
  const values = new Array(SALES_HEADER.length);
  SALES_HEADER.forEach(function (key, index) {
    values[index] = record[key] == null ? '' : record[key];
  });
  return values;
}

function findSalesRowById_(sheet, id, idIndex) {
  if (idIndex == null) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const column = Number(idIndex) + 1;
  const count = lastRow - 1;
  const values = sheet.getRange(2, column, count, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (sanitizeString_(values[i][0]) === id) {
      return i + 2;
    }
  }
  return -1;
}

function sanitizeSalesFilters_(raw) {
  const obj = raw || {};
  return {
    search: sanitizeString_(obj.search).toLowerCase(),
    pendingOnly: !!obj.pendingOnly
  };
}

function filterSalesRecord_(record, filters) {
  const hasSearch = !!filters.search;
  if (hasSearch) {
    const haystack = [
      record.reportDate,
      record.knetExpectedDate,
      record.knetReceivedDate,
      record.notes,
      record.expenseNotes,
      String(record.totalSales),
      String(record.knetSales),
      String(record.cashSales),
      String(record.expenses)
    ].join(' ').toLowerCase();
    if (haystack.indexOf(filters.search) === -1) {
      return false;
    }
  }
  if (filters.pendingOnly && record.knetStatus !== 'Pending') {
    return false;
  }
  return true;
}

function compareRecords_(a, b) {
  const aDate = a.reportDate || '';
  const bDate = b.reportDate || '';
  if (aDate < bDate) return 1;
  if (aDate > bDate) return -1;
  const aUpdated = a.updatedAt || '';
  const bUpdated = b.updatedAt || '';
  if (aUpdated < bUpdated) return 1;
  if (aUpdated > bUpdated) return -1;
  return 0;
}

function emptySummary_() {
  return {
    totalSales: 0,
    totalKnetSales: 0,
    totalCashSales: 0,
    totalExpenses: 0,
    netIncome: 0,
    pendingKnetAmount: 0,
    pendingKnetCount: 0,
    overdueKnetCount: 0,
    receivedKnetAmount: 0
  };
}

function accumulateSummary_(summary, record) {
  summary.totalSales = round2_(summary.totalSales + record.totalSales);
  summary.totalKnetSales = round2_(summary.totalKnetSales + record.knetSales);
  summary.totalCashSales = round2_(summary.totalCashSales + record.cashSales);
  summary.totalExpenses = round2_(summary.totalExpenses + record.expenses);
  summary.netIncome = round2_(summary.netIncome + record.netSales);
  if (record.knetStatus === 'Pending') {
    summary.pendingKnetAmount = round2_(summary.pendingKnetAmount + record.knetSales);
    summary.pendingKnetCount += 1;
    if (record.knetIsOverdue) {
      summary.overdueKnetCount += 1;
    }
  } else if (record.knetStatus === 'Received') {
    summary.receivedKnetAmount = round2_(summary.receivedKnetAmount + record.knetSales);
  }
}
