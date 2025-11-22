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
  'knetReceivedAmount',
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
const SHARE_EXPORTS_PUBLIC = false;

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
    const fromDate = toIsoDate_(query && query.fromDate);
    const toDate = toIsoDate_(query && query.toDate) || todayIso_();
    const sheet = getSalesSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return {
        ok: true,
        summary: salesEmptySummary_(),
        rows: [],
        filters,
        range: { fromDate: fromDate || '', toDate: toDate || '' },
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
      // Date range filter (default: today)
      const d = enriched.reportDate;
      if (d && (fromDate && d < fromDate)) continue;
      if (d && (toDate && d > toDate)) continue;
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
      range: { fromDate: fromDate || '', toDate: toDate || '' },
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
      if (!(parseMoney_(record.knetReceivedAmount) > 0) && parseMoney_(existing.knetReceivedAmount) > 0) {
        record.knetReceivedAmount = existing.knetReceivedAmount;
      }
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
    const amount = parseMoney_(data.amount);

    const sheet = getSalesSheet_();
    const headerMap = buildHeaderIndex_(SALES_HEADER);
    const rowNumber = findRowById_(sheet, id, headerMap.id);
    if (rowNumber < 0) {
      throw new Error('Unable to locate the requested record.');
    }
    const values = sheet.getRange(rowNumber, 1, 1, SALES_HEADER.length).getValues()[0];
    const record = mapSalesRowToObject_(values, headerMap);
    const total = parseMoney_(record.knetSales);
    const current = parseMoney_(record.knetReceivedAmount);
    const remaining = Math.max(total - current, 0);
    let receiveNow = amount > 0 ? amount : remaining;
    if (!(receiveNow > 0)) {
      throw new Error('Amount to receive must be greater than zero.');
    }
    if (receiveNow > remaining) {
      throw new Error('Amount exceeds pending KNET amount.');
    }
    record.knetReceivedAmount = round2_(current + receiveNow);
    // Only set received date when fully received
    if (record.knetReceivedAmount >= total && total > 0) {
      record.knetReceivedDate = toIsoDate_(data.receivedDate) || todayIso_();
      record.knetStatus = 'Received';
    } else {
      record.knetReceivedDate = '';
      record.knetStatus = 'Pending';
    }
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
 * SALES BULK KNET RECEIVE
 **********************/
function receiveKnetBulk(input) {
  try {
    const data = input || {};
    const amount = parseMoney_(data.amount);
    if (!(amount > 0)) {
      throw new Error('Amount to receive must be greater than zero.');
    }
    const receiveDate = toIsoDate_(data.receivedDate) || todayIso_();
    const sheet = getSalesSheet_();
    const headerMap = buildHeaderIndex_(SALES_HEADER);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { ok: true, applied: 0, remaining: amount, updated: 0 };
    }
    const values = sheet.getRange(2, 1, lastRow - 1, SALES_HEADER.length).getValues();
    // Build objects with pending amounts
    const rows = values.map(function (row, idx) {
      const rec = mapSalesRowToObject_(row, headerMap);
      const total = parseMoney_(rec.knetSales);
      const received = parseMoney_(rec.knetReceivedAmount);
      const pending = Math.max(total - received, 0);
      return { index: idx, rec: rec, pending: pending };
    }).filter(function (x) { return x.pending > 0; });

    // Sort FIFO by reportDate asc, then createdAt asc
    rows.sort(function (a, b) {
      const aDate = a.rec.reportDate || '';
      const bDate = b.rec.reportDate || '';
      if (aDate < bDate) return -1;
      if (aDate > bDate) return 1;
      const aCreated = a.rec.createdAt || '';
      const bCreated = b.rec.createdAt || '';
      if (aCreated < bCreated) return -1;
      if (aCreated > bCreated) return 1;
      return 0;
    });

    let remaining = amount;
    let updatedCount = 0;
    for (let i = 0; i < rows.length && remaining > 0; i++) {
      const entry = rows[i];
      const rec = entry.rec;
      const take = Math.min(entry.pending, remaining);
      if (!(take > 0)) continue;
      const newReceived = round2_(parseMoney_(rec.knetReceivedAmount) + take);
      rec.knetReceivedAmount = newReceived;
      if (newReceived >= parseMoney_(rec.knetSales) && parseMoney_(rec.knetSales) > 0) {
        rec.knetReceivedDate = receiveDate;
        rec.knetStatus = 'Received';
      } else {
        rec.knetReceivedDate = '';
        rec.knetStatus = 'Pending';
      }
      rec.updatedAt = toIsoTimestamp_(new Date());
      const valuesRow = buildRowValues_(rec, SALES_HEADER);
      sheet.getRange(2 + entry.index, 1, 1, SALES_HEADER.length).setValues([valuesRow]);
      updatedCount += 1;
      remaining = round2_(remaining - take);
    }

    return { ok: true, applied: round2_(amount - remaining), remaining: remaining, updated: updatedCount };
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
 * REPORT EXPORTS
 **********************/
function exportSalesReport(input) {
  try {
    const data = input || {};
    const from = toIsoDate_(data.fromDate);
    const to = toIsoDate_(data.toDate) || todayIso_();
    const format = (String(data.format || 'xlsx').toLowerCase() === 'pdf') ? 'pdf' : 'xlsx';

    const sheet = getSalesSheet_();
    const lastRow = sheet.getLastRow();
    const headerMap = buildHeaderIndex_(SALES_HEADER);
    const rows = [];
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, SALES_HEADER.length).getValues();
      for (var i = 0; i < values.length; i++) {
        const rec = mapSalesRowToObject_(values[i], headerMap);
        const d = rec.reportDate;
        if (d && (!from || d >= from) && (!to || d <= to)) {
          rows.push(finalizeSalesRecord_(rec));
        }
      }
    }
    const report = buildSalesReportFile_(rows, from, to);
    const exportUrl = buildSpreadsheetExportUrl_(report.id, format === 'pdf' ? report.summarySheetId : report.sheetId, format);
    return { ok: true, id: report.id, name: report.name, url: report.url, exportUrl: exportUrl, format: format };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function exportCashReport(input) {
  try {
    const data = input || {};
    const from = toIsoDate_(data.fromDate);
    const to = toIsoDate_(data.toDate) || todayIso_();
    const format = (String(data.format || 'xlsx').toLowerCase() === 'pdf') ? 'pdf' : 'xlsx';

    const sheet = getCashSheet_();
    const lastRow = sheet.getLastRow();
    const headerMap = buildHeaderIndex_(CASH_HEADER);
    const rows = [];
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, CASH_HEADER.length).getValues();
      for (var i = 0; i < values.length; i++) {
        const rec = mapCashRowToObject_(values[i], headerMap);
        const d = rec.entryDate;
        if (d && (!from || d >= from) && (!to || d <= to)) {
          rows.push(finalizeCashRecord_(rec));
        }
      }
    }
    const report = buildCashReportFile_(rows, from, to);
    const exportUrl = buildSpreadsheetExportUrl_(report.id, format === 'pdf' ? report.summarySheetId : report.sheetId, format);
    return { ok: true, id: report.id, name: report.name, url: report.url, exportUrl: exportUrl, format: format };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function exportBothReport(input) {
  try {
    const data = input || {};
    const from = toIsoDate_(data.fromDate);
    const to = toIsoDate_(data.toDate) || todayIso_();
    const format = (String(data.format || 'xlsx').toLowerCase() === 'pdf') ? 'pdf' : 'xlsx';

    // Build Sales rows
    const salesSheet = getSalesSheet_();
    const salesLast = salesSheet.getLastRow();
    const salesHeaderMap = buildHeaderIndex_(SALES_HEADER);
    const sales = [];
    if (salesLast >= 2) {
      const vals = salesSheet.getRange(2, 1, salesLast - 1, SALES_HEADER.length).getValues();
      for (var i = 0; i < vals.length; i++) {
        const rec = mapSalesRowToObject_(vals[i], salesHeaderMap);
        const d = rec.reportDate;
        if (d && (!from || d >= from) && (!to || d <= to)) {
          sales.push(finalizeSalesRecord_(rec));
        }
      }
    }

    // Build Cash rows
    const cashSheet = getCashSheet_();
    const cashLast = cashSheet.getLastRow();
    const cashHeaderMap = buildHeaderIndex_(CASH_HEADER);
    const cash = [];
    if (cashLast >= 2) {
      const vals = cashSheet.getRange(2, 1, cashLast - 1, CASH_HEADER.length).getValues();
      for (var j = 0; j < vals.length; j++) {
        const rec = mapCashRowToObject_(vals[j], cashHeaderMap);
        const d = rec.entryDate;
        if (d && (!from || d >= from) && (!to || d <= to)) {
          cash.push(finalizeCashRecord_(rec));
        }
      }
    }

    const report = buildBothReportFile_(sales, cash, from, to);
    const exportUrl = buildSpreadsheetExportUrl_(report.id, format === 'pdf' ? report.summarySheetId : report.firstSheetId, format);
    return { ok: true, id: report.id, name: report.name, url: report.url, exportUrl: exportUrl, format: format };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function buildBothReportFile_(salesRows, cashRows, from, to) {
  const name = 'Sales+Cash Report ' + (from || '') + (to ? ' to ' + to : '');
  const ss = SpreadsheetApp.create(name || 'Sales+Cash Report');
  // Sales sheet
  const salesSh = ss.getActiveSheet();
  salesSh.setName('Sales');
  const salesHeader = ['Date', 'Total', 'KNET', 'Cash', 'Expenses', 'Net', 'KNET Status', 'Expected', 'Received', 'Notes'];
  const salesValues = [salesHeader];
  let s_total = 0, s_k = 0, s_c = 0, s_e = 0, s_net = 0, s_pend = 0, s_pcount = 0;
  salesRows.forEach(function (r) {
    salesValues.push([r.reportDate, r.totalSales, r.knetSales, r.cashSales, r.expenses, r.netSales, r.knetStatus || '', r.knetExpectedDate || '', r.knetReceivedDate || '', r.expenseNotes || r.notes || '']);
    s_total += r.totalSales; s_k += r.knetSales; s_c += r.cashSales; s_e += r.expenses; s_net += r.netSales; if (r.knetStatus === 'Pending') { s_pend += r.knetSales; s_pcount++; }
  });
  salesSh.getRange(1, 1, salesValues.length, salesHeader.length).setValues(salesValues);
  salesSh.getRange(1, 1, 1, salesHeader.length).setFontWeight('bold');
  salesSh.autoResizeColumns(1, salesHeader.length);

  // Cash sheet
  const cashSh = ss.insertSheet('Cash');
  const cashHeader = ['Date', 'Direction', 'Category', 'Amount', 'Status', 'KNET Expected', 'KNET Received', 'Profit Transfer', 'Notes'];
  const cashValues = [cashHeader];
  let c_in = 0, c_out = 0, c_pend = 0, c_pcount = 0, c_inhand = 0;
  cashRows.forEach(function (r) {
    cashValues.push([r.entryDate, r.direction, r.category, r.amount, r.status, r.knetExpectedDate || '', r.knetReceivedDate || '', r.profitTransferDate || '', r.notes || '']);
    if (r.direction === 'IN') { c_in += r.amountNumber; if (r.category === 'KnetSale' && r.status === 'Pending') { c_pend += r.amountNumber; c_pcount++; } if (r.status === 'Received' && !r.profitTransferDate) { c_inhand += r.amountNumber; } }
    else { c_out += r.amountNumber; c_inhand -= r.amountNumber; }
  });
  cashSh.getRange(1, 1, cashValues.length, cashHeader.length).setValues(cashValues);
  cashSh.getRange(1, 1, 1, cashHeader.length).setFontWeight('bold');
  cashSh.autoResizeColumns(1, cashHeader.length);

  // Aggregated Cash for chart
  const aggMap = {};
  cashRows.forEach(function (r) {
    const d = r.entryDate || '';
    if (!d) return;
    if (!aggMap[d]) aggMap[d] = { In: 0, Out: 0 };
    if (r.direction === 'IN') aggMap[d].In += r.amountNumber; else aggMap[d].Out += r.amountNumber;
  });
  const dates = Object.keys(aggMap).sort();
  const aggSh = ss.insertSheet('CashAgg');
  const aggValues = [['Date', 'In', 'Out']].concat(dates.map(function (d) { return [d, aggMap[d].In, aggMap[d].Out]; }));
  aggSh.getRange(1, 1, aggValues.length, 3).setValues(aggValues);
  aggSh.getRange(1, 1, 1, 3).setFontWeight('bold');

  // Summary sheet with KPI and charts
  const sumSh = ss.insertSheet('Summary');
  const summary = [
    ['From', from || ''],
    ['To', to || ''],
    ['Total Sales', s_total],
    ['KNET Sales', s_k],
    ['Cash Sales', s_c],
    ['Expenses', s_e],
    ['Net Income', s_net],
    ['Pending KNET (count)', s_pcount],
    ['Pending KNET (amount)', s_pend],
    ['Total In', c_in],
    ['Total Out', c_out],
    ['Cash In Hand', c_inhand]
  ];
  sumSh.getRange(1, 1, summary.length, 2).setValues(summary);
  sumSh.getRange(1, 1, 2, 2).setFontWeight('bold');
  sumSh.getRange(3, 1, summary.length - 2, 2).setFontSize(12);
  sumSh.getRange(3, 2, summary.length - 2, 1).setFontWeight('bold');
  try {
    // Sales chart
    const salesChart = sumSh.newChart().asColumnChart()
      .addRange(salesSh.getRange(1, 1, Math.max(2, salesValues.length), 5))
      .setOption('title', 'Sales by Date')
      .setPosition(14, 1, 0, 0)
      .build();
    sumSh.insertChart(salesChart);
  } catch (e) {}
  try {
    // Cash In/Out chart
    const cashChart = sumSh.newChart().asColumnChart()
      .addRange(aggSh.getRange(1, 1, Math.max(2, aggValues.length), 3))
      .setOption('title', 'Cash In/Out by Date')
      .setPosition(28, 1, 0, 0)
      .build();
    sumSh.insertChart(cashChart);
  } catch (e) {}

  if (SHARE_EXPORTS_PUBLIC) {
    DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return { id: ss.getId(), firstSheetId: salesSh.getSheetId(), summarySheetId: sumSh.getSheetId(), name: name, url: ss.getUrl() };
}
function buildSalesReportFile_(rows, from, to) {
  const name = 'Sales Report ' + (from || '') + (to ? ' to ' + to : '');
  const ss = SpreadsheetApp.create(name || 'Sales Report');
  const sh = ss.getActiveSheet();
  sh.setName('Sales');
  const header = ['Date', 'Total', 'KNET', 'Cash', 'Expenses', 'Net', 'KNET Status', 'Expected', 'Received', 'Notes'];
  const values = [header];
  let total = 0, totalK = 0, totalC = 0, totalE = 0, net = 0, pendingK = 0, pendingCount = 0;
  rows.forEach(function (r) {
    values.push([
      r.reportDate,
      r.totalSales,
      r.knetSales,
      r.cashSales,
      r.expenses,
      r.netSales,
      r.knetStatus || '',
      r.knetExpectedDate || '',
      r.knetReceivedDate || '',
      r.expenseNotes || r.notes || ''
    ]);
    total += r.totalSales; totalK += r.knetSales; totalC += r.cashSales; totalE += r.expenses; net += r.netSales; if (r.knetStatus === 'Pending') { pendingK += r.knetSales; pendingCount++; }
  });
  sh.getRange(1, 1, values.length, header.length).setValues(values);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sh.autoResizeColumns(1, header.length);
  const sumSh = ss.insertSheet('Summary');
  const summary = [
    ['From', from || ''],
    ['To', to || ''],
    ['Total Sales', total],
    ['KNET Sales', totalK],
    ['Cash Sales', totalC],
    ['Expenses', totalE],
    ['Net Income', net],
    ['Pending KNET (count)', pendingCount],
    ['Pending KNET (amount)', pendingK]
  ];
  sumSh.getRange(1, 1, summary.length, 2).setValues(summary);
  sumSh.getRange(1, 1, 2, 2).setFontWeight('bold');
  // KPI formatting
  sumSh.getRange(3, 1, 7, 2).setFontSize(12);
  sumSh.getRange(3, 2, 7, 1).setFontWeight('bold');
  // Chart on Summary using Sales sheet data: Date vs Total/KNET/Cash/Expenses
  try {
    const chartRange = sh.getRange(1, 1, Math.max(2, values.length), 5);
    const chart = sumSh.newChart()
      .asColumnChart()
      .addRange(chartRange)
      .setOption('title', 'Sales by Date')
      .setPosition(10, 1, 0, 0)
      .build();
    sumSh.insertChart(chart);
  } catch (e) {}
  if (SHARE_EXPORTS_PUBLIC) {
    DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return { id: ss.getId(), sheetId: sh.getSheetId(), summarySheetId: sumSh.getSheetId(), name: name, url: ss.getUrl() };
}

function buildCashReportFile_(rows, from, to) {
  const name = 'Cash Report ' + (from || '') + (to ? ' to ' + to : '');
  const ss = SpreadsheetApp.create(name || 'Cash Report');
  const sh = ss.getActiveSheet();
  sh.setName('Cash');
  const header = ['Date', 'Direction', 'Category', 'Amount', 'Status', 'KNET Expected', 'KNET Received', 'Profit Transfer', 'Notes'];
  const values = [header];
  let totalIn = 0, totalOut = 0, pendingK = 0, pendingCount = 0, inHand = 0;
  rows.forEach(function (r) {
    values.push([
      r.entryDate,
      r.direction,
      r.category,
      r.amount,
      r.status,
      r.knetExpectedDate || '',
      r.knetReceivedDate || '',
      r.profitTransferDate || '',
      r.notes || ''
    ]);
    if (r.direction === 'IN') {
      totalIn += r.amountNumber;
      if (r.category === 'KnetSale' && r.status === 'Pending') { pendingK += r.amountNumber; pendingCount++; }
      if (r.status === 'Received' && !r.profitTransferDate) { inHand += r.amountNumber; }
    } else {
      totalOut += r.amountNumber;
      inHand -= r.amountNumber;
    }
  });
  sh.getRange(1, 1, values.length, header.length).setValues(values);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sh.autoResizeColumns(1, header.length);
  const sumSh = ss.insertSheet('Summary');
  const summary = [
    ['From', from || ''],
    ['To', to || ''],
    ['Total In', totalIn],
    ['Total Out', totalOut],
    ['Pending KNET (count)', pendingCount],
    ['Pending KNET (amount)', pendingK],
    ['Cash In Hand', inHand]
  ];
  sumSh.getRange(1, 1, summary.length, 2).setValues(summary);
  sumSh.getRange(1, 1, 2, 2).setFontWeight('bold');
  // Build aggregated IN/OUT per date for chart
  const aggMap = {};
  rows.forEach(function (r) {
    const d = r.entryDate || '';
    if (!d) return;
    if (!aggMap[d]) aggMap[d] = { In: 0, Out: 0 };
    if (r.direction === 'IN') aggMap[d].In += r.amountNumber;
    else aggMap[d].Out += r.amountNumber;
  });
  const dates = Object.keys(aggMap).sort();
  const aggSh = ss.insertSheet('CashAgg');
  const aggValues = [['Date', 'In', 'Out']].concat(dates.map(function (d) { return [d, aggMap[d].In, aggMap[d].Out]; }));
  aggSh.getRange(1, 1, aggValues.length, 3).setValues(aggValues);
  aggSh.getRange(1, 1, 1, 3).setFontWeight('bold');
  try {
    const chart = sumSh.newChart()
      .asColumnChart()
      .addRange(aggSh.getRange(1, 1, Math.max(2, aggValues.length), 3))
      .setOption('title', 'Cash In/Out by Date')
      .setPosition(10, 1, 0, 0)
      .build();
    sumSh.insertChart(chart);
  } catch (e) {}
  if (SHARE_EXPORTS_PUBLIC) {
    DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return { id: ss.getId(), sheetId: sh.getSheetId(), summarySheetId: sumSh.getSheetId(), name: name, url: ss.getUrl() };
}

function buildSpreadsheetExportUrl_(id, sheetId, format) {
  // format: 'xlsx' or 'pdf'
  const base = 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(id) + '/export';
  if (format === 'pdf') {
    const params = [
      'format=pdf',
      'size=A4',
      'portrait=false',
      'fitw=true',
      'sheetnames=true',
      'printtitle=true',
      'pagenumbers=RIGHT',
      'gridlines=false',
      'fzr=true',
      'gid=' + encodeURIComponent(sheetId)
    ].join('&');
    return base + '?' + params;
  }
  return base + '?format=xlsx';
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
  let knetReceivedAmount = parseMoney_(payload.knetReceivedAmount);
  const cashSales = parseMoney_(payload.cashSales);
  const expenses = parseMoney_(payload.expenses);
  let knetExpectedDate = toIsoDate_(payload.knetExpectedDate);
  let knetReceivedDate = toIsoDate_(payload.knetReceivedDate);
  let status = sanitizeString_(payload.knetStatus);
  const markReceived = !!payload.markReceived;
  const reconciled = reconcileSalesAmounts_(totalSales, knetSales, cashSales);

  if (reconciled.knetSales > 0) {
    if (!knetExpectedDate) {
      knetExpectedDate = addDaysIso_(reportDate, DEFAULT_KNET_DELAY_DAYS);
    }
    if (markReceived && !knetReceivedDate) {
      knetReceivedDate = todayIso_();
    }
    if (markReceived && !(knetReceivedAmount > 0)) {
      knetReceivedAmount = reconciled.knetSales;
    }
    if (knetReceivedAmount > reconciled.knetSales) {
      knetReceivedAmount = reconciled.knetSales;
    }
    if (knetReceivedDate) {
      status = 'Received';
    } else if (!status) {
      status = 'Pending';
    }
  } else {
    knetExpectedDate = '';
    knetReceivedDate = '';
    knetReceivedAmount = 0;
    status = '';
  }

  return {
    id: sanitizeString_(payload.id),
    reportDate,
    totalSales: reconciled.totalSales,
    knetSales: reconciled.knetSales,
    knetReceivedAmount,
    cashSales: reconciled.cashSales,
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

/**
 * Ensures totalSales always matches the tender mix (KNET + cash) so downstream
 * summaries and validations operate on consistent numbers.
 */
function reconcileSalesAmounts_(totalSales, knetSales, cashSales) {
  let total = round2_(Number(totalSales) || 0);
  let knet = round2_(Math.max(Number(knetSales) || 0, 0));
  let cash = round2_(Math.max(Number(cashSales) || 0, 0));

  if (cash === 0 && total > knet) {
    cash = round2_(Math.max(total - knet, 0));
  }
  if (!(total > 0)) {
    total = round2_(knet + cash);
  }
  if (total < knet) {
    total = round2_(Math.max(knet, knet + cash));
  }

  const tenderSum = round2_(knet + cash);
  if (tenderSum !== total) {
    total = tenderSum;
  }

  return {
    totalSales: total,
    knetSales: knet,
    cashSales: cash
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
  record.knetReceivedAmount = parseMoney_(record.knetReceivedAmount);
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
  const received = parseMoney_(normalized.knetReceivedAmount);
  const pending = Math.max(parseMoney_(normalized.knetSales) - received, 0);
  normalized.knetPendingAmount = pending;
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
  const receivedAmount = parseMoney_(record.knetReceivedAmount);
  const total = parseMoney_(record.knetSales);
  if (received || receivedAmount >= total) {
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
  const b64 = sanitizeString_(obj.b64 || obj.base64);
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
