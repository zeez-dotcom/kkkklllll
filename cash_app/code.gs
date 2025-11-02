/**********************
 * CONFIG
 **********************/
const CASH_SPREADSHEET_ID = '';
const CASH_SHEET_NAME = 'CashLedger';
const CASH_HEADER = [
  'id',
  'entryDate',
  'direction',
  'category',
  'amount',
  'description',
  'notes',
  'knetExpectedDate',
  'knetReceivedDate',
  'knetStatus',
  'profitTransferDate',
  'linkedEntryId',
  'status',
  'createdAt',
  'updatedAt'
];
const CASH_DEFAULT_KNET_DELAY_DAYS = 10;

/**********************
 * WEB APP ENTRY
 **********************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Cash In Hand Dashboard');
}

/**********************
 * PUBLIC API
 **********************/
function getCashDashboardData(query) {
  try {
    const sheet = getCashSheet_();
    const filters = sanitizeCashFilters_(query || {});
    const last = sheet.getLastRow();
    if (last < 2) {
      return {
        ok: true,
        summary: emptyCashSummary_(),
        rows: [],
        filters,
        generatedAt: toIsoTimestamp_(new Date())
      };
    }
    const values = sheet.getRange(2, 1, last - 1, CASH_HEADER.length).getValues();
    const headerMap = buildHeaderIndex_(CASH_HEADER);
    const rows = [];
    const summary = emptyCashSummary_();

    for (let i = 0; i < values.length; i++) {
      const record = mapCashRowToObject_(values[i], headerMap);
      const enriched = finalizeCashRecord_(record);
      if (!filterCashRecord_(enriched, filters)) {
        continue;
      }
      accumulateCashSummary_(summary, enriched);
      rows.push(enriched);
    }

    rows.sort(compareCashRecords_);

    return {
      ok: true,
      summary,
      rows,
      filters,
      generatedAt: toIsoTimestamp_(new Date())
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function recordCashEntry(payload) {
  try {
    const sheet = getCashSheet_();
    const headerMap = buildHeaderIndex_(CASH_HEADER);
  const raw = payload || {};
  const sanitized = sanitizeCashPayload_(raw);
    const now = new Date();
    const nowIso = toIsoTimestamp_(now);
    let record = Object.assign({}, sanitized);
    let rowNumber = -1;
    let existing = null;

    if (record.id) {
      rowNumber = findCashRowById_(sheet, record.id, headerMap.id);
    }

    if (rowNumber < 0) {
      record.id = generateId_();
      record.createdAt = nowIso;
    } else {
      const row = sheet.getRange(rowNumber, 1, 1, CASH_HEADER.length).getValues()[0];
      existing = mapCashRowToObject_(row, headerMap);
      record.createdAt = existing.createdAt || nowIso;
      if (!record.profitTransferDate && existing.profitTransferDate) {
        record.profitTransferDate = existing.profitTransferDate;
      }
      if (!record.linkedEntryId && existing.linkedEntryId) {
        record.linkedEntryId = existing.linkedEntryId;
      }
    }

  record.updatedAt = nowIso;
  const knetInfo = computeCashKnetStatus_(record);
  record.knetStatus = knetInfo.knetStatus;

  const values = buildCashRowValues_(record);
    if (rowNumber < 0) {
      sheet.appendRow(values);
    } else {
      sheet.getRange(rowNumber, 1, 1, CASH_HEADER.length).setValues([values]);
    }

    return {
      ok: true,
      record: finalizeCashRecord_(record)
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function markCashKnetReceived(input) {
  try {
    const data = input || {};
    const id = sanitizeString_(data.id);
    if (!id) {
      throw new Error('Record ID is required.');
    }
    const sheet = getCashSheet_();
    const headerMap = buildHeaderIndex_(CASH_HEADER);
    const rowNumber = findCashRowById_(sheet, id, headerMap.id);
    if (rowNumber < 0) {
      throw new Error('Unable to locate the requested record.');
    }
    const row = sheet.getRange(rowNumber, 1, 1, CASH_HEADER.length).getValues()[0];
    const record = mapCashRowToObject_(row, headerMap);
    if (record.direction !== 'IN' || record.category !== 'KnetSale') {
      throw new Error('Only KNET income entries can be marked as received.');
    }
    const receivedDate = toIsoDate_(data.receivedDate) || todayIso_();
    record.knetReceivedDate = receivedDate;
    record.knetStatus = 'Received';
    record.status = 'Received';
    record.updatedAt = toIsoTimestamp_(new Date());
    const values = buildCashRowValues_(record);
    sheet.getRange(rowNumber, 1, 1, CASH_HEADER.length).setValues([values]);
    return {
      ok: true,
      record: finalizeCashRecord_(record)
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function transferCashToProfit(input) {
  try {
    const data = input || {};
    const id = sanitizeString_(data.id);
    if (!id) {
      throw new Error('Record ID is required.');
    }
    const sheet = getCashSheet_();
    const headerMap = buildHeaderIndex_(CASH_HEADER);
    const sourceRowNumber = findCashRowById_(sheet, id, headerMap.id);
    if (sourceRowNumber < 0) {
      throw new Error('Unable to locate the requested record.');
    }
    const sourceRow = sheet.getRange(sourceRowNumber, 1, 1, CASH_HEADER.length).getValues()[0];
    const source = mapCashRowToObject_(sourceRow, headerMap);
    if (source.direction !== 'IN') {
      throw new Error('Only income entries can be transferred to profits.');
    }
    if (source.status !== 'Received') {
      throw new Error('The income must be marked as received before transferring to profits.');
    }
    if (source.profitTransferDate) {
      throw new Error('This entry has already been transferred to profits.');
    }

    const transferDate = toIsoDate_(data.transferDate) || todayIso_();
    const outRecord = {
      id: generateId_(),
      entryDate: transferDate,
      direction: 'OUT',
      category: 'ProfitTransfer',
      amount: source.amount,
      description: 'Profit transfer for ' + (source.description || source.category || source.id),
      notes: sanitizeString_(data.notes || ''),
      knetExpectedDate: '',
      knetReceivedDate: '',
      knetStatus: '',
      profitTransferDate: '',
      linkedEntryId: source.id,
      status: 'Posted',
      createdAt: toIsoTimestamp_(new Date()),
      updatedAt: toIsoTimestamp_(new Date())
    };

    const appendValues = buildCashRowValues_(outRecord);
    sheet.appendRow(appendValues);

    source.profitTransferDate = transferDate;
    source.status = 'Transferred';
    source.updatedAt = toIsoTimestamp_(new Date());

    const updateValues = buildCashRowValues_(source);
    sheet.getRange(sourceRowNumber, 1, 1, CASH_HEADER.length).setValues([updateValues]);

    return {
      ok: true,
      record: finalizeCashRecord_(source)
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

/**********************
 * HELPERS
 **********************/
function getCashSheet_() {
  const configuredId = typeof CASH_SPREADSHEET_ID === 'string' ? CASH_SPREADSHEET_ID.trim() : '';
  let ss = null;
  if (configuredId) {
    try {
      ss = SpreadsheetApp.openById(configuredId);
    } catch (err) {
      throw new Error('Unable to open configured cash spreadsheet.');
    }
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) {
    throw new Error('Cash spreadsheet unavailable.');
  }
  let sheet = ss.getSheetByName(CASH_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CASH_SHEET_NAME);
  }
  ensureHeader_(sheet, CASH_HEADER);
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

function sanitizeCashFilters_(raw) {
  const obj = raw || {};
  return {
    search: sanitizeString_(obj.search).toLowerCase(),
    showPendingOnly: !!obj.showPendingOnly
  };
}

function filterCashRecord_(record, filters) {
  if (!record) return false;
  if (filters.showPendingOnly && record.status !== 'Pending') {
    return false;
  }
  const search = filters.search;
  if (search) {
    const haystack = [
      record.entryDate,
      record.category,
      record.direction,
      record.status,
      record.description,
      record.notes,
      record.knetStatus,
      record.knetExpectedDate,
      record.knetReceivedDate,
      record.profitTransferDate,
      record.amount
    ].join(' ').toLowerCase();
    if (haystack.indexOf(search) === -1) {
      return false;
    }
  }
  return true;
}

function emptyCashSummary_() {
  return {
    totalIn: 0,
    totalOut: 0,
    cashInHand: 0,
    pendingKnetAmount: 0,
    pendingKnetCount: 0,
    overdueKnetCount: 0,
    profitTransferredAmount: 0,
    profitTransferredCount: 0
  };
}

function accumulateCashSummary_(summary, record) {
  const amount = record.amountNumber;
  if (record.direction === 'IN') {
    summary.totalIn = round2_(summary.totalIn + amount);
    if (record.status === 'Pending' && record.category === 'KnetSale') {
      summary.pendingKnetAmount = round2_(summary.pendingKnetAmount + amount);
      summary.pendingKnetCount += 1;
      if (record.knetIsOverdue) {
        summary.overdueKnetCount += 1;
      }
    }
    if (record.status === 'Received' && !record.profitTransferDate) {
      summary.cashInHand = round2_(summary.cashInHand + amount);
    }
  } else {
    summary.totalOut = round2_(summary.totalOut + amount);
    summary.cashInHand = round2_(summary.cashInHand - amount);
    if (record.category === 'ProfitTransfer') {
      summary.profitTransferredAmount = round2_(summary.profitTransferredAmount + amount);
      summary.profitTransferredCount += 1;
    }
  }
}

function compareCashRecords_(a, b) {
  const dateA = a.entryDate || '';
  const dateB = b.entryDate || '';
  if (dateA < dateB) return 1;
  if (dateA > dateB) return -1;
  const updatedA = a.updatedAt || '';
  const updatedB = b.updatedAt || '';
  if (updatedA < updatedB) return 1;
  if (updatedA > updatedB) return -1;
  return 0;
}

function sanitizeCashPayload_(raw) {
  const payload = raw || {};
  const entryDate = toIsoDate_(payload.entryDate) || todayIso_();
  const direction = sanitizeDirection_(payload.direction, payload.category);
  const category = sanitizeCategory_(payload.category, direction);
  const amount = parseMoney_(payload.amount);
  if (!(amount > 0)) {
    throw new Error('Amount must be greater than zero.');
  }
  let knetExpectedDate = toIsoDate_(payload.knetExpectedDate);
  let knetReceivedDate = toIsoDate_(payload.knetReceivedDate);
  let status = sanitizeString_(payload.status);
  const markReceived = !!payload.markReceived;

  if (direction === 'IN' && category === 'KnetSale') {
    if (!knetExpectedDate) {
      knetExpectedDate = addDaysIso_(entryDate, CASH_DEFAULT_KNET_DELAY_DAYS);
    }
    if (markReceived && !knetReceivedDate) {
      knetReceivedDate = todayIso_();
    }
    if (knetReceivedDate) {
      status = 'Received';
    } else if (!status) {
      status = 'Pending';
    }
  } else {
    knetExpectedDate = '';
    knetReceivedDate = '';
    if (direction === 'IN') {
      status = 'Received';
    } else {
      status = 'Posted';
    }
  }

  return {
    id: sanitizeString_(payload.id),
    entryDate: entryDate,
    direction: direction,
    category: category,
    amount: amount,
    description: sanitizeString_(payload.description),
    notes: sanitizeString_(payload.notes),
    knetExpectedDate: knetExpectedDate,
    knetReceivedDate: knetReceivedDate,
    knetStatus: computeKnetStatusLabel_(direction, category, status, knetReceivedDate),
    profitTransferDate: toIsoDate_(payload.profitTransferDate),
    linkedEntryId: sanitizeString_(payload.linkedEntryId),
    status: status,
    createdAt: sanitizeString_(payload.createdAt),
    updatedAt: sanitizeString_(payload.updatedAt)
  };
}

function computeKnetStatusLabel_(direction, category, status, knetReceivedDate) {
  if (direction !== 'IN' || category !== 'KnetSale') {
    return '';
  }
  if (status === 'Received' || knetReceivedDate) {
    return 'Received';
  }
  return 'Pending';
}

function sanitizeDirection_(direction, category) {
  const value = sanitizeString_(direction).toUpperCase();
  if (value === 'IN' || value === 'OUT') {
    return value;
  }
  const cat = sanitizeString_(category);
  if (cat === 'Expense' || cat === 'ProfitTransfer' || cat === 'CashOut') {
    return 'OUT';
  }
  return 'IN';
}

function sanitizeCategory_(category, direction) {
  const value = sanitizeString_(category);
  if (value) return value;
  return direction === 'IN' ? 'CashSale' : 'Expense';
}

function buildCashRowValues_(record) {
  return CASH_HEADER.map(function (key) {
    return record[key] == null ? '' : record[key];
  });
}

function findCashRowById_(sheet, id, idIndex) {
  if (idIndex == null) return -1;
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const column = Number(idIndex) + 1;
  const values = sheet.getRange(2, column, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (sanitizeString_(values[i][0]) === id) {
      return i + 2;
    }
  }
  return -1;
}

function mapCashRowToObject_(row, headerMap) {
  const record = {};
  CASH_HEADER.forEach(function (key, index) {
    record[key] = row[index];
  });
  record.id = sanitizeString_(record.id);
  record.entryDate = toIsoDate_(record.entryDate);
  record.direction = sanitizeDirection_(record.direction, record.category);
  record.category = sanitizeCategory_(record.category, record.direction);
  record.amount = parseMoney_(record.amount);
  record.description = sanitizeString_(record.description);
  record.notes = sanitizeString_(record.notes);
  record.knetExpectedDate = toIsoDate_(record.knetExpectedDate);
  record.knetReceivedDate = toIsoDate_(record.knetReceivedDate);
  record.knetStatus = sanitizeString_(record.knetStatus);
  record.profitTransferDate = toIsoDate_(record.profitTransferDate);
  record.linkedEntryId = sanitizeString_(record.linkedEntryId);
  record.status = sanitizeString_(record.status);
  record.createdAt = record.createdAt ? toIsoTimestamp_(record.createdAt) : '';
  record.updatedAt = record.updatedAt ? toIsoTimestamp_(record.updatedAt) : '';
  return record;
}

function finalizeCashRecord_(record) {
  const normalized = Object.assign({}, record);
  normalized.amountNumber = parseMoney_(normalized.amount);
  normalized.amountFormatted = round2_(normalized.amountNumber);
  const knetInfo = computeCashKnetStatus_(normalized);
  normalized.knetStatus = knetInfo.knetStatus;
  normalized.knetDueDays = knetInfo.knetDueDays;
  normalized.knetIsOverdue = knetInfo.knetIsOverdue;
  normalized.knetDelayDays = knetInfo.knetDelayDays;
  normalized.isIncome = normalized.direction === 'IN';
  normalized.isExpense = normalized.direction === 'OUT' && normalized.category !== 'ProfitTransfer';
  normalized.isProfitTransfer = normalized.category === 'ProfitTransfer';
  normalized.profitTransferred = !!normalized.profitTransferDate;
  normalized.knetExpectedDisplay = normalized.knetExpectedDate;
  normalized.knetReceivedDisplay = normalized.knetReceivedDate;
  return normalized;
}

function computeCashKnetStatus_(record) {
  if (record.direction !== 'IN' || record.category !== 'KnetSale') {
    return {
      knetStatus: '',
      knetDueDays: null,
      knetIsOverdue: false,
      knetDelayDays: null
    };
  }
  const today = dateFromIso_(todayIso_());
  const expected = dateFromIso_(record.knetExpectedDate);
  const received = dateFromIso_(record.knetReceivedDate);
  let status = sanitizeString_(record.knetStatus);
  if (received) {
    status = 'Received';
  } else {
    status = status || 'Pending';
  }
  let dueDays = null;
  let overdue = false;
  if (status === 'Pending' && expected && today) {
    const oneDay = 24 * 60 * 60 * 1000;
    dueDays = Math.round((expected.getTime() - today.getTime()) / oneDay);
    overdue = dueDays < 0;
  }
  let delayDays = null;
  if (status === 'Received' && expected && received) {
    const oneDay = 24 * 60 * 60 * 1000;
    delayDays = Math.round((received.getTime() - expected.getTime()) / oneDay);
  }
  return {
    knetStatus: status,
    knetDueDays: dueDays,
    knetIsOverdue: overdue,
    knetDelayDays: delayDays
  };
}

function sanitizeString_(value) {
  return String(value == null ? '' : value).trim();
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

function toIsoTimestamp_(value) {
  const date = Object.prototype.toString.call(value) === '[object Date]' ? value : new Date(value);
  if (isNaN(date)) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function todayIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
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
