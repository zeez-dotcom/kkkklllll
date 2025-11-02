/**********************
 * WEB APP ENTRY
 **********************/
/**********************
 * CONFIG
 **********************/
const SALES_SPREADSHEET_ID = '';
const SALES_SHEET_NAME = 'SalesRecords';
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
const SALES_RECEIPT_FOLDER_NAME = 'Sales Receipts';
let SALES_RECEIPT_FOLDER_ID = '';
const SHARE_RECEIPTS_PUBLIC = false;
const MAX_RECEIPT_SIZE_BYTES = 5 * 1024 * 1024;

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
    .setTitle('Sales & Cash Dashboard');
}

/**********************
 * SALES API
 **********************/
function getSalesDashboardData(query) {
  try {
    const filters = sanitizeSalesFilters_(query || {});
    const sheet = getSalesSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return {
        ok: true,
        summary: salesEmptySummary_(),
        rows: [],
        filters,
        generatedAt: toIsoTimestamp_(new Date())
      };
    }
    const values = sheet.getRange(2, 1, lastRow - 1, SALES_HEADER.length).getValues();
    const headerMap = buildHeaderIndex_(SALES_HEADER);
    const rows = [];
    const summary = salesEmptySummary_();
    for (let i = 0; i < values.length; i++) {
      const record = mapSalesRowToObject_(values[i], headerMap);
      const enriched = finalizeSalesRecord_(record);
      if (filterSalesRecord_(enriched, filters)) {
        rows.push(enriched);
      }
      accumulateSalesSummary_(summary, enriched);
    }
    rows.sort(compareSalesRecords_);
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

    if (record.id) {
      rowNumber = findRowById_(sheet, record.id, headerMap.id);
    }

    if (rowNumber < 0) {
      record.id = generateId_();
      record.createdAt = nowIso;
      record.receiptUrl = record.receiptUrl || '';
      record.receiptName = record.receiptName || '';
    } else {
      const existingRow = sheet.getRange(rowNumber, 1, 1, SALES_HEADER.length).getValues()[0];
      const existing = mapSalesRowToObject_(existingRow, headerMap);
      record.createdAt = existing.createdAt || nowIso;
      record.receiptUrl = record.receiptUrl || existing.receiptUrl || '';
      record.receiptName = record.receiptName || existing.receiptName || '';
    }

    if (receipt) {
      const stored = storeSalesReceipt_(receipt);
      record.receiptUrl = stored.url;
      record.receiptName = stored.name;
    }

    record.updatedAt = nowIso;
    record = Object.assign(record, computeSalesKnetStatus_(record));

    const values = buildRowValues_(record, SALES_HEADER);
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
    if (!id) throw new Error('Record ID is required.');

    const sheet = getSalesSheet_();
    const headerMap = buildHeaderIndex_(SALES_HEADER);
    const rowNumber = findRowById_(sheet, id, headerMap.id);
    if (rowNumber < 0) {
      throw new Error('Unable to locate the requested record.');
    }
    const values = sheet.getRange(rowNumber, 1, 1, SALES_HEADER.length).getValues()[0];
    const record = mapSalesRowToObject_(values, headerMap);
    record.knetReceivedDate = toIsoDate_(data.receivedDate) || todayIso_();
    record.knetStatus = 'Received';
    record.updatedAt = toIsoTimestamp_(new Date());

    const computed = Object.assign(record, computeSalesKnetStatus_(record));
    sheet.getRange(rowNumber, 1, 1, SALES_HEADER.length).setValues([
      buildRowValues_(computed, SALES_HEADER)
    ]);

    return {
      ok: true,
      record: finalizeSalesRecord_(computed)
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

/**********************
 * CASH API
 **********************/
function getCashDashboardData(query) {
  try {
    const sheet = getCashSheet_();
    const filters = sanitizeCashFilters_(query || {});
    const last = sheet.getLastRow();
    if (last < 2) {
      return {
        ok: true,
        summary: cashEmptySummary_(),
        rows: [],
        filters,
        generatedAt: toIsoTimestamp_(new Date())
      };
    }
    const values = sheet.getRange(2, 1, last - 1, CASH_HEADER.length).getValues();
    const headerMap = buildHeaderIndex_(CASH_HEADER);
    const rows = [];
    const summary = cashEmptySummary_();
    for (let i = 0; i < values.length; i++) {
      const record = mapCashRowToObject_(values[i], headerMap);
      const enriched = finalizeCashRecord_(record);
      if (!filterCashRecord_(enriched, filters)) continue;
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
    const nowIso = toIsoTimestamp_(new Date());
    let record = Object.assign({}, sanitized);
    let rowNumber = -1;

    if (record.id) {
      rowNumber = findRowById_(sheet, record.id, headerMap.id);
    }

    if (rowNumber < 0) {
      record.id = generateId_();
      record.createdAt = nowIso;
    } else {
      const existingRow = sheet.getRange(rowNumber, 1, 1, CASH_HEADER.length).getValues()[0];
      const existing = mapCashRowToObject_(existingRow, headerMap);
      record.createdAt = existing.createdAt || nowIso;
      if (!record.profitTransferDate && existing.profitTransferDate) {
        record.profitTransferDate = existing.profitTransferDate;
      }
      if (!record.linkedEntryId && existing.linkedEntryId) {
        record.linkedEntryId = existing.linkedEntryId;
      }
    }

    record.updatedAt = nowIso;
    record = Object.assign(record, computeCashKnetStatus_(record));

    const values = buildRowValues_(record, CASH_HEADER);
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
    if (!id) throw new Error('Record ID is required.');

    const sheet = getCashSheet_();
    const headerMap = buildHeaderIndex_(CASH_HEADER);
    const rowNumber = findRowById_(sheet, id, headerMap.id);
    if (rowNumber < 0) {
      throw new Error('Unable to locate the requested record.');
    }

    const values = sheet.getRange(rowNumber, 1, 1, CASH_HEADER.length).getValues()[0];
    const record = mapCashRowToObject_(values, headerMap);
    if (record.direction !== 'IN' || record.category !== 'KnetSale') {
      throw new Error('Only KNET income entries can be marked as received.');
    }

    record.knetReceivedDate = toIsoDate_(data.receivedDate) || todayIso_();
    record.knetStatus = 'Received';
    record.status = 'Received';
    record.updatedAt = toIsoTimestamp_(new Date());

    const computed = Object.assign(record, computeCashKnetStatus_(record));
    sheet.getRange(rowNumber, 1, 1, CASH_HEADER.length).setValues([
      buildRowValues_(computed, CASH_HEADER)
    ]);

    return {
      ok: true,
      record: finalizeCashRecord_(computed)
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function transferCashToProfit(input) {
  try {
    const data = input || {};
    const id = sanitizeString_(data.id);
    if (!id) throw new Error('Record ID is required.');

    const sheet = getCashSheet_();
    const headerMap = buildHeaderIndex_(CASH_HEADER);
    const sourceRowNumber = findRowById_(sheet, id, headerMap.id);
    if (sourceRowNumber < 0) {
      throw new Error('Unable to locate the requested record.');
    }

    const sourceRow = sheet.getRange(sourceRowNumber, 1, 1, CASH_HEADER.length).getValues()[0];
    const source = mapCashRowToObject_(sourceRow, headerMap);
    if (source.direction !== 'IN') {
      throw new Error('Only income entries can be transferred to profits.');
    }
    if (source.status !== 'Received') {
      throw new Error('Mark the entry as received before transferring to profits.');
    }
    if (source.profitTransferDate) {
      throw new Error('This entry has already been transferred to profits.');
    }

    const transferDate = toIsoDate_(data.transferDate) || todayIso_();
    const nowIso = toIsoTimestamp_(new Date());
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
      createdAt: nowIso,
      updatedAt: nowIso
    };

    sheet.appendRow(buildRowValues_(outRecord, CASH_HEADER));

    source.profitTransferDate = transferDate;
    source.status = 'Transferred';
    source.updatedAt = nowIso;
    const updatedSource = Object.assign(source, computeCashKnetStatus_(source));
    sheet.getRange(sourceRowNumber, 1, 1, CASH_HEADER.length).setValues([
      buildRowValues_(updatedSource, CASH_HEADER)
    ]);

    return {
      ok: true,
      record: finalizeCashRecord_(updatedSource)
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

/**********************
 * SALES HELPERS
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
  if (!ss) throw new Error('Sales spreadsheet unavailable.');
  let sheet = ss.getSheetByName(SALES_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SALES_SHEET_NAME);
  ensureHeader_(sheet, SALES_HEADER);
  return sheet;
}

function sanitizeSalesFilters_(raw) {
  const obj = raw || {};
  return {
    search: sanitizeString_(obj.search).toLowerCase(),
    pendingOnly: !!obj.pendingOnly
  };
}

function filterSalesRecord_(record, filters) {
  if (!record) return false;
  if (filters.pendingOnly && record.knetStatus !== 'Pending') {
    return false;
  }
  if (filters.search) {
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
  return true;
}

function salesEmptySummary_() {
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

function accumulateSalesSummary_(summary, record) {
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

function compareSalesRecords_(a, b) {
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

function sanitizeSalesPayload_(raw) {
  const payload = raw || {};
  const reportDate = toIsoDate_(payload.reportDate) || todayIso_();
  const totalSales = parseMoney_(payload.totalSales);
  const knetSales = parseMoney_(payload.knetSales);
  const cashSales = parseMoney_(payload.cashSales);
  const expenses = parseMoney_(payload.expenses);
  let knetExpectedDate = toIsoDate_(payload.knetExpectedDate);
  let knetReceivedDate = toIsoDate_(payload.knetReceivedDate);
  let status = sanitizeString_(payload.knetStatus);
  const markReceived = !!payload.markReceived;

  if (knetSales > 0) {
    if (!knetExpectedDate) {
      knetExpectedDate = addDaysIso_(reportDate, DEFAULT_KNET_DELAY_DAYS);
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
    status = '';
  }

  return {
    id: sanitizeString_(payload.id),
    reportDate,
    totalSales,
    knetSales,
    cashSales,
    expenses,
    expenseNotes: sanitizeString_(payload.expenseNotes),
    notes: sanitizeString_(payload.notes),
    receiptUrl: sanitizeUrl_(payload.receiptUrl),
    receiptName: sanitizeString_(payload.receiptName),
    knetExpectedDate,
    knetReceivedDate,
    knetStatus: status,
    createdAt: sanitizeString_(payload.createdAt),
    updatedAt: sanitizeString_(payload.updatedAt)
  };
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
  const knetInfo = computeSalesKnetStatus_(normalized);
  normalized.knetStatus = knetInfo.knetStatus;
  normalized.knetDueDays = knetInfo.knetDueDays;
  normalized.knetDelayDays = knetInfo.knetDelayDays;
  normalized.knetIsOverdue = knetInfo.knetIsOverdue;
  normalized.netSales = round2_(normalized.totalSales - normalized.expenses);
  normalized.knetPendingAmount = normalized.knetStatus === 'Pending' ? normalized.knetSales : 0;
  normalized.receiptUrl = sanitizeUrl_(normalized.receiptUrl);
  normalized.receiptName = sanitizeString_(normalized.receiptName || '');
  normalized.receiptPreviewUrl = makeDrivePreviewUrl_(normalized.receiptUrl);
  normalized.hasReceipt = !!normalized.receiptUrl;
  return normalized;
}

function computeSalesKnetStatus_(record) {
  const hasKnet = parseMoney_(record.knetSales) > 0;
  if (!hasKnet) {
    return {
      knetStatus: '',
      knetDueDays: null,
      knetDelayDays: null,
      knetIsOverdue: false
    };
  }
  const today = dateFromIso_(todayIso_());
  const expected = dateFromIso_(record.knetExpectedDate);
  const received = dateFromIso_(record.knetReceivedDate);
  let status = sanitizeString_(record.knetStatus);
  if (received) {
    status = 'Received';
  } else if (!status) {
    status = 'Pending';
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
    knetDelayDays: delayDays,
    knetIsOverdue: overdue
  };
}

function getSalesReceiptFolder_() {
  if (SALES_RECEIPT_FOLDER_ID) {
    try {
      return DriveApp.getFolderById(SALES_RECEIPT_FOLDER_ID);
    } catch (err) {
      SALES_RECEIPT_FOLDER_ID = '';
    }
  }
  const name = SALES_RECEIPT_FOLDER_NAME || 'Sales Receipts';
  const existing = DriveApp.getFoldersByName(name);
  const folder = existing.hasNext() ? existing.next() : DriveApp.createFolder(name);
  SALES_RECEIPT_FOLDER_ID = folder.getId();
  return folder;
}

function storeSalesReceipt_(receipt) {
  const folder = getSalesReceiptFolder_();
  const bytes = decodeBase64Safely_(receipt.b64, 'receiptUpload', MAX_RECEIPT_SIZE_BYTES);
  const mime = receipt.type || 'image/jpeg';
  const blob = Utilities.newBlob(bytes, mime, receipt.name);
  const created = folder.createFile(blob).setName(receipt.name);
  if (SHARE_RECEIPTS_PUBLIC) {
    created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return {
    url: sanitizeUrl_(created.getUrl()),
    name: receipt.name
  };
}

/**********************
 * CASH HELPERS
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
  if (!ss) throw new Error('Cash spreadsheet unavailable.');
  let sheet = ss.getSheetByName(CASH_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CASH_SHEET_NAME);
  ensureHeader_(sheet, CASH_HEADER);
  return sheet;
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
  if (filters.search) {
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
    if (haystack.indexOf(filters.search) === -1) {
      return false;
    }
  }
  return true;
}

function cashEmptySummary_() {
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
    if (record.category === 'KnetSale' && record.status === 'Pending') {
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
  const aDate = a.entryDate || '';
  const bDate = b.entryDate || '';
  if (aDate < bDate) return 1;
  if (aDate > bDate) return -1;
  const aUpdated = a.updatedAt || '';
  const bUpdated = b.updatedAt || '';
  if (aUpdated < bUpdated) return 1;
  if (aUpdated > bUpdated) return -1;
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
    status = direction === 'IN' ? 'Received' : 'Posted';
  }

  return {
    id: sanitizeString_(payload.id),
    entryDate,
    direction,
    category,
    amount,
    description: sanitizeString_(payload.description),
    notes: sanitizeString_(payload.notes),
    knetExpectedDate,
    knetReceivedDate,
    knetStatus: sanitizeString_(payload.knetStatus) || computeCashKnetStatusLabel_(direction, category, status, knetReceivedDate),
    profitTransferDate: toIsoDate_(payload.profitTransferDate),
    linkedEntryId: sanitizeString_(payload.linkedEntryId),
    status,
    createdAt: sanitizeString_(payload.createdAt),
    updatedAt: sanitizeString_(payload.updatedAt)
  };
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
  const knetInfo = computeCashKnetStatus_(normalized);
  normalized.knetStatus = knetInfo.knetStatus;
  normalized.knetDueDays = knetInfo.knetDueDays;
  normalized.knetIsOverdue = knetInfo.knetIsOverdue;
  normalized.knetDelayDays = knetInfo.knetDelayDays;
  normalized.isIncome = normalized.direction === 'IN';
  normalized.isExpense = normalized.direction === 'OUT' && normalized.category !== 'ProfitTransfer';
  normalized.isProfitTransfer = normalized.category === 'ProfitTransfer';
  normalized.profitTransferred = !!normalized.profitTransferDate;
  return normalized;
}

function computeCashKnetStatusLabel_(direction, category, status, knetReceivedDate) {
  if (direction !== 'IN' || category !== 'KnetSale') {
    return '';
  }
  if (status === 'Received' || knetReceivedDate) {
    return 'Received';
  }
  return 'Pending';
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

function sanitizeDirection_(direction, category) {
  const value = sanitizeString_(direction).toUpperCase();
  if (value === 'IN' || value === 'OUT') return value;
  const cat = sanitizeString_(category);
  if (cat === 'Expense' || cat === 'ProfitTransfer' || cat === 'CashOut') return 'OUT';
  return 'IN';
}

function sanitizeCategory_(category, direction) {
  const value = sanitizeString_(category);
  if (value) return value;
  return direction === 'IN' ? 'CashSale' : 'Expense';
}

/**********************
 * SHARED HELPERS
 **********************/
function ensureHeader_(sheet, header) {
  const first = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  const mismatch = first.length !== header.length || header.some(function (name, index) {
    return first[index] !== name;
  });
  if (mismatch) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
}

function sanitizeString_(value) {
  return String(value == null ? '' : value).trim();
}

function sanitizeUrl_(url) {
  const value = sanitizeString_(url);
  if (!value) return '';
  return /^https?:\/\//i.test(value) ? value : '';
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
  if (isNaN(parsed)) return fallback || '';
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

function findRowById_(sheet, id, index) {
  if (index == null) return -1;
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const column = Number(index) + 1;
  const values = sheet.getRange(2, column, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (sanitizeString_(values[i][0]) === id) {
      return i + 2;
    }
  }
  return -1;
}

function buildRowValues_(record, header) {
  return header.map(function (key) {
    return record[key] == null ? '' : record[key];
  });
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
    name,
    b64,
    type,
    size: isFinite(size) ? size : 0
  };
}

function decodeBase64Safely_(encoded, context, maxBytes) {
  const raw = sanitizeString_(encoded);
  if (!raw) throw new Error('Missing encoded data for ' + (context || 'upload') + '.');
  const normalized = raw.indexOf(',') >= 0 ? raw.split(',').pop() : raw;
  const cleaned = normalized.replace(/\s/g, '');
  let bytes;
  try {
    bytes = Utilities.base64Decode(cleaned);
  } catch (err) {
    throw new Error('Unable to decode file for ' + (context || 'upload') + '.');
  }
  if (maxBytes && bytes.length > maxBytes) {
    throw new Error('Receipt image exceeds maximum size limit.');
  }
  return bytes;
}

function makeDrivePreviewUrl_(url) {
  const safe = sanitizeUrl_(url);
  if (!safe) return '';
  const idMatch = safe.match(/\/d\/([a-zA-Z0-9_-]+)/);
  const id = idMatch && idMatch[1]
    ? idMatch[1]
    : (function () {
        const paramMatch = safe.match(/[?&]id=([a-zA-Z0-9_-]+)/);
        return paramMatch && paramMatch[1] ? paramMatch[1] : '';
      })();
  if (!id) return safe;
  return 'https://drive.google.com/uc?export=view&id=' + id;
}
