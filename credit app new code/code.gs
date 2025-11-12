/*********************************************************
 * Credit Ledger (New Build)
 * Backend: Google Apps Script
 * Author: Codex CLI Agent
 *********************************************************/
const CL = {
  PEOPLE_SHEET: 'People',
  LEDGER_SHEET: 'Ledger',
  RECEIPTS_FOLDER_PROP: 'CL_RECEIPTS_FOLDER_ID',
  RECEIPTS_FOLDER_NAME: 'Credit Ledger Receipts',
  TIMEZONE: 'Asia/Kuwait',
  CURRENCY: 'KWD',
  MAX_UPLOAD_BYTES: 5 * 1024 * 1024
};

const CL_HEADERS = {
  PEOPLE: ['id','name','phone','location','notes','profileFileId','createdAt','updatedAt','active'],
  LEDGER: ['id','personId','type','amountKWD','tranDate','note','receiptFileId','createdAt']
};

/*********************************************************
 * Entry points
 *********************************************************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Credit Ledger');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Credit Ledger')
    .addItem('Open Dashboard', 'openSidebar')
    .addSeparator()
    .addItem('Initial Setup', 'clSetup')
    .addItem('Send Outstanding Alerts', 'sendOutstandingAlerts')
    .addItem('Install Monthly Alert Trigger', 'installMonthlyAlertTrigger')
    .addToUi();
}

function openSidebar() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('index').setTitle('Credit Ledger'));
}

function clSetup() {
  return withLock_(() => {
    ensureSheet_(CL.PEOPLE_SHEET, CL_HEADERS.PEOPLE);
    ensureSheet_(CL.LEDGER_SHEET, CL_HEADERS.LEDGER);
    ensureReceiptsFolder_();
    return { ok: true };
  });
}

/*********************************************************
 * Public API (called from frontend)
 *********************************************************/
function apiGetDashboard(filters) {
  filters = sanitizeFilters_(filters || {});
  ensureSheet_(CL.PEOPLE_SHEET, CL_HEADERS.PEOPLE);
  ensureSheet_(CL.LEDGER_SHEET, CL_HEADERS.LEDGER);
  const people = listPeople_();
  const ledger = listLedger_(filters);
  const summary = buildSummary_(people, ledger);
  return {
    summary,
    currency: CL.CURRENCY,
    generatedAt: nowIso_()
  };
}

function apiListPeople() {
  ensureSheet_(CL.PEOPLE_SHEET, CL_HEADERS.PEOPLE);
  return listPeople_();
}

function apiUpsertPerson(payload) {
  if (!payload || !payload.name) throw new Error('Name is required');
  ensureSheet_(CL.PEOPLE_SHEET, CL_HEADERS.PEOPLE);
  return withLock_(() => {
    const sheet = getSheet_(CL.PEOPLE_SHEET);
    const table = getTable_(sheet);
    const map = indexBy_(table.rows, 'id');
    const now = nowIso_();
    if (payload.id && map[payload.id] != null) {
      const idx = map[payload.id];
      const row = table.rows[idx];
      row.name = payload.name.trim();
      row.phone = (payload.phone || '').trim();
      row.location = (payload.location || '').trim();
      row.notes = (payload.notes || '').trim();
      row.updatedAt = now;
      writeRows_(sheet, CL_HEADERS.PEOPLE, [row], idx + 2);
      return row;
    }
    const newRow = {
      id: Utilities.getUuid(),
      name: payload.name.trim(),
      phone: (payload.phone || '').trim(),
      location: (payload.location || '').trim(),
      notes: (payload.notes || '').trim(),
      profileFileId: payload.profileFileId || '',
      createdAt: now,
      updatedAt: now,
      active: 'TRUE'
    };
    appendRow_(sheet, CL_HEADERS.PEOPLE, newRow);
    return newRow;
  });
}

function apiSetPersonProfilePic(personId, base64, fileName) {
  if (!personId) throw new Error('Missing person id');
  if (!base64 || !fileName) throw new Error('Missing file data');
  ensureSheet_(CL.PEOPLE_SHEET, CL_HEADERS.PEOPLE);
  return withLock_(() => {
    const sheet = getSheet_(CL.PEOPLE_SHEET);
    const table = getTable_(sheet);
    const map = indexBy_(table.rows, 'id');
    const idx = map[personId];
    if (idx == null) throw new Error('Person not found');
    const fileId = saveBase64File_(base64, fileName);
    table.rows[idx].profileFileId = fileId;
    table.rows[idx].updatedAt = nowIso_();
    writeRows_(sheet, CL_HEADERS.PEOPLE, [table.rows[idx]], idx + 2);
    return { id: personId, profileFileId: fileId };
  });
}

function apiAddEntry(payload) {
  const required = ['personId','type','amountKWD','tranDate'];
  required.forEach(key => {
    if (!payload || payload[key] == null || payload[key] === '') {
      throw new Error('Missing field: ' + key);
    }
  });
  ensureSheet_(CL.PEOPLE_SHEET, CL_HEADERS.PEOPLE);
  ensureSheet_(CL.LEDGER_SHEET, CL_HEADERS.LEDGER);
  return withLock_(() => {
    const people = listPeople_();
    if (!people.some(p => p.id === payload.personId)) {
      throw new Error('Person not found');
    }
    const entry = {
      id: Utilities.getUuid(),
      personId: payload.personId,
      type: String(payload.type).toUpperCase() === 'PAYMENT' ? 'PAYMENT' : 'CREDIT',
      amountKWD: round3_(payload.amountKWD),
      tranDate: toDateString_(payload.tranDate),
      note: (payload.note || '').trim(),
      receiptFileId: '',
      createdAt: nowIso_()
    };
    if (!(entry.amountKWD > 0)) throw new Error('Amount must be greater than 0');
    if (!entry.tranDate) throw new Error('Invalid date');
    if (payload.receiptBase64 && payload.receiptName) {
      entry.receiptFileId = saveBase64File_(payload.receiptBase64, payload.receiptName);
    }
    appendRow_(getSheet_(CL.LEDGER_SHEET), CL_HEADERS.LEDGER, entry);
    return entry;
  });
}

function apiListTransactions(filters) {
  filters = sanitizeFilters_(filters || {});
  ensureSheet_(CL.LEDGER_SHEET, CL_HEADERS.LEDGER);
  return listLedger_(filters);
}

function sendOutstandingAlerts() {
  return sendOutstandingAlerts_();
}

function installMonthlyAlertTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendOutstandingAlerts') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('sendOutstandingAlerts')
    .timeBased()
    .onMonthDay(28)
    .atHour(9)
    .inTimezone(CL.TIMEZONE)
    .create();
  return { ok: true };
}

/*********************************************************
 * Dashboard helpers
 *********************************************************/
function buildSummary_(people, ledger) {
  const totals = ledger.reduce((acc, row) => {
    if (!acc[row.personId]) {
      acc[row.personId] = { credit: 0, payment: 0, lastTranDate: '' };
    }
    const bucket = acc[row.personId];
    if (row.type === 'CREDIT') bucket.credit += row.amountKWD;
    if (row.type === 'PAYMENT') bucket.payment += row.amountKWD;
    if (!bucket.lastTranDate || (row.tranDate && row.tranDate > bucket.lastTranDate)) {
      bucket.lastTranDate = row.tranDate;
    }
    return acc;
  }, {});

  return people.map(person => {
    const bucket = totals[person.id] || { credit: 0, payment: 0, lastTranDate: '' };
    const balance = round3_(bucket.credit - bucket.payment);
    return {
      personId: person.id,
      name: person.name,
      phone: person.phone,
      location: person.location,
      notes: person.notes,
      profileFileId: person.profileFileId,
      totalCredit: round3_(bucket.credit),
      totalPayment: round3_(bucket.payment),
      balance,
      lastTranDate: bucket.lastTranDate
    };
  });
}

function sanitizeFilters_(filters) {
  const clean = Object.assign({}, filters);
  if (!clean.personId) delete clean.personId;
  if (!clean.type) delete clean.type;
  if (!clean.dateFrom) delete clean.dateFrom;
  if (!clean.dateTo) delete clean.dateTo;
  if (clean.hasReceipt === true || clean.hasReceipt === false) return clean;
  if (clean.hasReceipt === 'has') clean.hasReceipt = true;
  else if (clean.hasReceipt === 'none') clean.hasReceipt = false;
  else delete clean.hasReceipt;
  return clean;
}

/*********************************************************
 * Sheet helpers
 *********************************************************/
function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const diff = headers.some((header, idx) => existingHeaders[idx] !== header);
  if (diff) {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const rows = values.slice(1);
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) {
      const remapped = rows.map(row => headers.map((header, idx) => row[idx] ?? ''));
      sheet.getRange(2, 1, remapped.length, headers.length).setValues(remapped);
    }
  }
  if (sheet.getFrozenRows() !== 1) sheet.setFrozenRows(1);
  return sheet;
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name + '. Run setup first.');
  return sheet;
}

function getTable_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 1) return { headers: [], rows: [] };
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = {};
    headers.forEach((header, idx) => row[header] = values[i][idx]);
    if (row.amountKWD !== undefined && row.amountKWD !== '') {
      row.amountKWD = Number(row.amountKWD);
    }
    if (row.tranDate) {
      row.tranDate = toDateString_(row.tranDate);
    }
    rows.push(row);
  }
  return { headers, rows };
}

function appendRow_(sheet, headers, row) {
  sheet.appendRow(headers.map(header => row[header] ?? ''));
}

function writeRows_(sheet, headers, rows, startRow) {
  const data = rows.map(row => headers.map(header => row[header] ?? ''));
  sheet.getRange(startRow, 1, data.length, headers.length).setValues(data);
}

function indexBy_(rows, key) {
  return rows.reduce((map, row, idx) => {
    map[row[key]] = idx;
    return map;
  }, {});
}

function listPeople_() {
  const sheet = getSheet_(CL.PEOPLE_SHEET);
  const table = getTable_(sheet);
  const updates = [];
  table.rows.forEach((row, idx) => {
    let dirty = false;
    if (!row.id) { row.id = Utilities.getUuid(); dirty = true; }
    if (!row.createdAt) { row.createdAt = nowIso_(); dirty = true; }
    if (!row.updatedAt) { row.updatedAt = row.createdAt; dirty = true; }
    if (!row.active) { row.active = 'TRUE'; dirty = true; }
    if (dirty) updates.push({ row, idx });
  });
  if (updates.length) {
    updates.forEach(update => writeRows_(sheet, CL_HEADERS.PEOPLE, [update.row], update.idx + 2));
  }
  return table.rows.filter(row => String(row.active || 'TRUE').toUpperCase() !== 'FALSE');
}

function listLedger_(filters) {
  const sheet = getSheet_(CL.LEDGER_SHEET);
  const table = getTable_(sheet);
  return table.rows.filter(row => {
    if (filters.personId && row.personId !== filters.personId) return false;
    if (filters.type && row.type !== filters.type.toUpperCase()) return false;
    if (filters.dateFrom && row.tranDate && row.tranDate < filters.dateFrom) return false;
    if (filters.dateTo && row.tranDate && row.tranDate > filters.dateTo) return false;
    if (filters.hasReceipt === true && !row.receiptFileId) return false;
    if (filters.hasReceipt === false && row.receiptFileId) return false;
    return true;
  }).sort((a,b) => {
    if (a.tranDate === b.tranDate) return (a.createdAt || '').localeCompare(b.createdAt || '');
    if (!a.tranDate) return 1;
    if (!b.tranDate) return -1;
    return a.tranDate < b.tranDate ? -1 : 1;
  });
}

/*********************************************************
 * Drive helpers
 *********************************************************/
function ensureReceiptsFolder_() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty(CL.RECEIPTS_FOLDER_PROP);
  if (folderId) {
    try {
      DriveApp.getFolderById(folderId);
      return folderId;
    } catch (err) {
      folderId = null;
    }
  }
  const folder = DriveApp.createFolder(CL.RECEIPTS_FOLDER_NAME);
  props.setProperty(CL.RECEIPTS_FOLDER_PROP, folder.getId());
  return folder.getId();
}

function getReceiptsFolder_() {
  return DriveApp.getFolderById(ensureReceiptsFolder_());
}

function saveBase64File_(dataUrl, fileName) {
  const parts = dataUrl.split(',');
  const meta = parts[0];
  const payload = parts[1] || parts[0];
  const mime = (meta.match(/data:(.*);base64/) || [])[1] || 'application/octet-stream';
  const bytes = Utilities.base64Decode(payload);
  if (bytes.length > CL.MAX_UPLOAD_BYTES) throw new Error('File must be ≤ 5MB');
  const blob = Utilities.newBlob(bytes, mime, fileName || 'receipt');
  const file = getReceiptsFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
  return file.getId();
}

/*********************************************************
 * Alerts
 *********************************************************/
function sendOutstandingAlerts_() {
  const dashboard = apiGetDashboard({});
  const outstanding = dashboard.summary.filter(row => row.balance > 0);
  const email = Session.getEffectiveUser().getEmail();
  if (!email) throw new Error('No email address available');
  if (!outstanding.length) {
    MailApp.sendEmail(email, 'Credit Ledger — All clear', 'No outstanding balances as of ' + nowIso_());
    return { sent: true, count: 0 };
  }
  const rows = outstanding.map(row => `
    <tr>
      <td>${escapeHtml_(row.name)}</td>
      <td>${row.balance.toFixed(3)}</td>
      <td>${row.lastTranDate || '—'}</td>
    </tr>
  `).join('');
  const total = outstanding.reduce((sum, row) => sum + row.balance, 0);
  const htmlBody = `
    <h2>Outstanding Balances (${CL.CURRENCY})</h2>
    <table border="1" cellpadding="6" style="border-collapse:collapse;">
      <tr><th>Name</th><th>Balance</th><th>Last Transaction</th></tr>
      ${rows}
    </table>
    <p><strong>Total: ${total.toFixed(3)} ${CL.CURRENCY}</strong></p>
  `;
  MailApp.sendEmail({ to: email, subject: 'Credit Ledger — Outstanding Balances', htmlBody });
  return { sent: true, count: outstanding.length, total };
}

/*********************************************************
 * Utilities
 *********************************************************/
function nowIso_() {
  return Utilities.formatDate(new Date(), CL.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function toDateString_(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(date, CL.TIMEZONE, 'yyyy-MM-dd');
}

function round3_(value) {
  return Math.round(Number(value) * 1000) / 1000;
}

function escapeHtml_(value) {
  const map = { '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' };
  return String(value || '').replace(/[&<>"']/g, ch => map[ch]);
}

function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
