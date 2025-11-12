/***************************************
 * Credit Ledger App — Backend (Code.gs)
 * Timezone: Asia/Kuwait
 ***************************************/
const CONFIG = {
  SHEET_PEOPLE: 'People',
  SHEET_LEDGER: 'Ledger',
  PROP_FOLDER_ID: 'RECEIPTS_FOLDER_ID',
  TIMEZONE: 'Asia/Kuwait',
  CURRENCY: 'KWD'
};

const HEADERS = {
  PEOPLE: [
    'id',
    'name',
    'phone',
    'location',
    'notes',
    'profileFileId',
    'createdAt',
    'updatedAt',
    'active'
  ],
  LEDGER: [
    'id',
    'personId',
    'type',
    'amountKWD',
    'tranDate',
    'note',
    'receiptFileId',
    'createdAt'
  ]
};

/** ====== Entry points ====== **/
function onOpen() {
  try {
    if (SpreadsheetApp.getActive().getSpreadsheetTimeZone() !== CONFIG.TIMEZONE) {
      SpreadsheetApp.getActive().setSpreadsheetTimeZone(CONFIG.TIMEZONE);
    }

    SpreadsheetApp.getUi()
      .createMenu('Credit Ledger')
      .addItem('Open Dashboard', 'openSidebar')
      .addSeparator()
      .addItem('Run Initial Setup', 'setup')
      .addItem('Send Monthly Alerts Now', 'sendMonthlyAlerts')
      .addSeparator()
      .addItem('Install Monthly Alert Trigger', 'installMonthlyAlertTrigger')
      .addToUi();
  } catch (error) {
    console.error('Error in onOpen:', error);
  }
}

function openSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Credit Ledger');
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    console.error('Error opening sidebar:', error);
    SpreadsheetApp.getUi().alert('Error opening sidebar: ' + error.message);
  }
}

function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('Index').setTitle('Credit Ledger');
  } catch (error) {
    console.error('Error in doGet:', error);
    return HtmlService.createHtmlOutput('<h1>Error loading application</h1>');
  }
}

/** ====== One-time setup ====== **/
function setup() {
  return withLock_(() => {
    try {
      ensureBaseSheets_();

      const props = PropertiesService.getScriptProperties();
      if (!props.getProperty(CONFIG.PROP_FOLDER_ID)) {
        const folder = DriveApp.createFolder('CreditLedger_Receipts');
        props.setProperty(CONFIG.PROP_FOLDER_ID, folder.getId());
      }
      return { ok: true, message: 'Setup completed successfully' };
    } catch (error) {
      console.error('Setup error:', error);
      throw new Error('Setup failed: ' + error.message);
    }
  });
}

function ensureSheet_(name, headers) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      initializeSheet_(sh, headers);
      return sh;
    }

    const range = sh.getDataRange();
    const values = range.getValues();
    const headerIndex = findHeaderRowIndex_(values, headers);
    const currentHeaders = headerIndex >= 0 ? values[headerIndex] : (values.length ? values[0] : []);
    const normalizedCurrent = currentHeaders.map(normalizeHeaderKey_);
    const normalizedExpected = headers.map(normalizeHeaderKey_);

    const headerLookup = {};
    normalizedCurrent.forEach((key, idx) => {
      if (key) headerLookup[key] = idx;
    });

    const headerMismatch = normalizedExpected.some(
      (expected, idx) => expected !== (normalizedCurrent[idx] || '')
    );
    const missingRequired = normalizedExpected.some(
      key => key && headerLookup[key] == null
    );
    const hasExtraNamed = normalizedCurrent.some(
      key => key && !normalizedExpected.includes(key)
    );

    if (headerMismatch || missingRequired || hasExtraNamed) {
      const sourceRows =
        headerIndex >= 0
          ? values.slice(headerIndex + 1)
          : values.length > 1
            ? values.slice(1)
            : [];
      const remappedRows = sourceRows
        .filter(hasRowData_)
        .map(row =>
          headers.map(h => {
            const idx = headerLookup[normalizeHeaderKey_(h)];
            return idx == null ? '' : row[idx];
          })
        );
      initializeSheet_(sh, headers, remappedRows);
      return sh;
    }

    if (sh.getMaxColumns() > headers.length) {
      sh.deleteColumns(headers.length + 1, sh.getMaxColumns() - headers.length);
    } else if (sh.getMaxColumns() < headers.length) {
      sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
    }

    if (sh.getFrozenRows() !== 1) {
      sh.setFrozenRows(1);
    }

    return sh;
  } catch (error) {
    console.error('Error ensuring sheet:', name, error);
    throw error;
  }
}

function initializeSheet_(sh, headers, rows) {
  sh.clear();
  const maxCols = sh.getMaxColumns();
  if (maxCols > headers.length) {
    sh.deleteColumns(headers.length + 1, maxCols - headers.length);
  } else if (maxCols < headers.length) {
    sh.insertColumnsAfter(maxCols, headers.length - maxCols);
  }
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sh.setFrozenRows(1);
}

function normalizeHeaderKey_(value) {
  if (value == null) return '';
  return String(value).trim().toLowerCase();
}

function hasRowData_(row) {
  if (!row || !row.length) return false;
  return row.some(cell => cell !== '' && cell !== null && String(cell).trim() !== '');
}

function findHeaderRowIndex_(values, expectedHeaders) {
  if (!values || !values.length) return -1;
  const normalizedExpected = expectedHeaders.map(normalizeHeaderKey_);
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (!hasRowData_(row)) continue;
    const normalizedRow = row.map(normalizeHeaderKey_);
    const matches = normalizedRow.reduce((count, key) => {
      if (!key) return count;
      return normalizedExpected.includes(key) ? count + 1 : count;
    }, 0);
    if (matches >= Math.min(normalizedExpected.length, 4)) {
      return i;
    }
  }
  return -1;
}

function ensureBaseSheets_() {
  ensureSheet_(CONFIG.SHEET_PEOPLE, HEADERS.PEOPLE);
  ensureSheet_(CONFIG.SHEET_LEDGER, HEADERS.LEDGER);
}

/** ====== Data helpers ====== **/
function nowStr_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function toDateString_(d) {
  if (!d) return '';
  try {
    const date = new Date(d);
    return Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  } catch (error) {
    console.error('Date conversion error:', d, error);
    return '';
  }
}

function uuid_() {
  return Utilities.getUuid();
}

function getReceiptsFolder_() {
  try {
    const id = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_FOLDER_ID);
    if (!id) throw new Error('Receipts folder not configured. Run setup first.');
    return DriveApp.getFolderById(id);
  } catch (error) {
    console.error('Error getting receipts folder:', error);
    throw new Error('Could not access receipts folder: ' + error.message);
  }
}

function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  const hasLock = lock.tryLock(10000); // 10 seconds
  if (!hasLock) throw new Error('Could not acquire lock. Please try again.');
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/** ====== Public API (called from client) ====== **/

// Return dashboard summary + people + recent ledger (optional filters)
function apiGetDashboard(params) {
  try {
    params = params || {};
    ensureBaseSheets_();
    const people = listPeople_();
    const ledger = listLedger_(params);

    // Build balances and last transaction date per person
    const byPerson = {};
    ledger.forEach(row => {
      const pid = row.personId;
      if (!byPerson[pid]) byPerson[pid] = { credit:0, payment:0, lastTranDate:null, count:0 };
      if (row.type === 'CREDIT') byPerson[pid].credit += row.amountKWD;
      if (row.type === 'PAYMENT') byPerson[pid].payment += row.amountKWD;
      byPerson[pid].count++;
      if (!byPerson[pid].lastTranDate || (row.tranDate && row.tranDate > byPerson[pid].lastTranDate)) {
        byPerson[pid].lastTranDate = row.tranDate;
      }
    });

    const summary = people.map(p => {
      const agg = byPerson[p.id] || { credit:0, payment:0, lastTranDate:null, count:0 };
      const balance = +(agg.credit - agg.payment).toFixed(3);
      return {
        personId: p.id,
        name: p.name,
        phone: p.phone,
        location: p.location,
        notes: p.notes,
        profileFileId: p.profileFileId,
        balance,
        totalCredit: +agg.credit.toFixed(3),
        totalPayment: +agg.payment.toFixed(3),
        lastTranDate: agg.lastTranDate,
        transactions: agg.count
      };
    });

    // Optional sorting
    if (params.sortBy) {
      const dir = params.sortDir === 'asc' ? 1 : -1;
      summary.sort((a,b) => {
        const va = a[params.sortBy], vb = b[params.sortBy];
        if (va == null && vb == null) return 0;
        if (va == null) return 1;
        if (vb == null) return -1;
        if (va < vb) return -1 * dir;
        if (va > vb) return 1 * dir;
        return 0;
      });
    }

    return { 
      people, 
      ledger: ledger.slice(0, 1000), // Limit for performance
      summary, 
      currency: CONFIG.CURRENCY, 
      timezone: CONFIG.TIMEZONE 
    };
  } catch (error) {
    console.error('Error in apiGetDashboard:', error);
    throw new Error('Failed to load dashboard: ' + error.message);
  }
}

// Create or update a person
function apiUpsertPerson(person) {
  if (!person) throw new Error('Missing person data');
  if (!person.name || person.name.trim() === '') throw new Error('Name is required');

  return withLock_(() => {
    try {
      const sh = ensureSheet_(CONFIG.SHEET_PEOPLE, HEADERS.PEOPLE);
      const data = getTable_(sh);
      const idxById = indexBy_(data.rows, 'id');

      // Check for duplicate names (case insensitive, excluding current person and inactive)
      if (!person.id || data.rows[idxById[person.id]]?.name !== person.name) {
        const existingName = data.rows.find(r => 
          r.name && r.name.toLowerCase() === person.name.toLowerCase() && 
          r.id !== person.id && String(r.active || 'TRUE') !== 'FALSE'
        );
        if (existingName) throw new Error('Person with this name already exists');
      }

      if (!person.id) {
        // create
        const id = uuid_();
        const row = {
          id,
          name: (person.name || '').trim(),
          phone: (person.phone || '').trim(),
          location: (person.location || '').trim(),
          notes: (person.notes || '').trim(),
          profileFileId: person.profileFileId || '',
          createdAt: nowStr_(),
          updatedAt: nowStr_(),
          active: 'TRUE'
        };
        appendRow_(sh, HEADERS.PEOPLE, row);
        return row;
      } else {
        // update
        const idx = idxById[person.id];
        if (idx == null) throw new Error('Person not found');
        const row = data.rows[idx];
        row.name = (person.name ?? row.name).trim();
        row.phone = (person.phone ?? row.phone).trim();
        row.location = (person.location ?? row.location).trim();
        row.notes = (person.notes ?? row.notes).trim();
        row.profileFileId = person.profileFileId ?? row.profileFileId;
        row.updatedAt = nowStr_();
        writeRows_(sh, HEADERS.PEOPLE, [row], idx+2); // +2 = header row + 1-based
        return row;
      }
    } catch (error) {
      console.error('Error in apiUpsertPerson:', error);
      throw error;
    }
  });
}

// Add a ledger entry: CREDIT or PAYMENT (with optional base64 receipt/profile)
function apiAddEntry(entry) {
  if (!entry) throw new Error('Missing entry data');

  return withLock_(() => {
    try {
      const required = ['personId','type','amountKWD','tranDate'];
      required.forEach(k => { 
        if (entry[k] == null || entry[k] === '') 
          throw new Error('Missing required field: ' + k); 
      });

      const type = String(entry.type).toUpperCase();
      if (!['CREDIT','PAYMENT'].includes(type)) throw new Error('Type must be CREDIT or PAYMENT');

      const amount = Number(entry.amountKWD);
      if (!isFinite(amount) || amount <= 0) throw new Error('Amount must be a positive number');
      if (amount > 1000000) throw new Error('Amount too large');

      // Validate person exists
      const people = listPeople_();
      const personExists = people.some(p => p.id === entry.personId);
      if (!personExists) throw new Error('Person not found');

      let receiptFileId = '';
      if (entry.receiptBase64 && entry.receiptName) {
        receiptFileId = saveBase64ToDrive_(entry.receiptBase64, entry.receiptName);
      }

      const sh = ensureSheet_(CONFIG.SHEET_LEDGER, HEADERS.LEDGER);
      const id = uuid_();
      const row = {
        id,
        personId: entry.personId,
        type,
        amountKWD: +Number(Math.round(amount * 1000) / 1000).toFixed(3),
        tranDate: toDateString_(entry.tranDate),
        note: (entry.note || '').trim(),
        receiptFileId,
        createdAt: nowStr_()
      };
      appendRow_(sh, HEADERS.LEDGER, row);
      return row;
    } catch (error) {
      console.error('Error in apiAddEntry:', error);
      throw error;
    }
  });
}

// List transactions (optional filters: personId, dateFrom, dateTo, type, hasReceipt)
function apiListTransactions(filters) {
  try {
    return listLedger_(filters || {});
  } catch (error) {
    console.error('Error in apiListTransactions:', error);
    throw new Error('Failed to load transactions: ' + error.message);
  }
}

// List people
function apiListPeople() {
  try {
    return listPeople_();
  } catch (error) {
    console.error('Error in apiListPeople:', error);
    throw new Error('Failed to load people: ' + error.message);
  }
}

// Attach/update profile picture for a person
function apiSetPersonProfilePic(personId, base64, filename) {
  if (!personId) throw new Error('Missing person ID');
  if (!base64 || !filename) throw new Error('Missing file data');

  return withLock_(() => {
    try {
      const sh = ensureSheet_(CONFIG.SHEET_PEOPLE, HEADERS.PEOPLE);
      const data = getTable_(sh);
      const idxById = indexBy_(data.rows, 'id');
      const idx = idxById[personId];
      if (idx == null) throw new Error('Person not found');
      
      const fileId = saveBase64ToDrive_(base64, filename);
      data.rows[idx].profileFileId = fileId;
      data.rows[idx].updatedAt = nowStr_();
      writeRows_(sh, HEADERS.PEOPLE, [data.rows[idx]], idx+2);
      return { personId, fileId };
    } catch (error) {
      console.error('Error in apiSetPersonProfilePic:', error);
      throw error;
    }
  });
}

// Build a Drive URL (thumbnail link)
function apiFileUrl(fileId) {
  if (!fileId) return '';
  return 'https://drive.google.com/uc?id=' + encodeURIComponent(fileId);
}

/** ====== Alerts & Triggers ====== **/

// Email outstanding balances nearing month-end
function sendMonthlyAlerts() {
  try {
    const dashboard = apiGetDashboard({});
    const outstanding = dashboard.summary
      .filter(p => p.balance > 0)
      .sort((a,b) => b.balance - a.balance);

    const owner = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail();
    if (!owner) throw new Error('No email address found for user');

    if (!outstanding.length) {
      MailApp.sendEmail(owner, 'Credit Ledger — No outstanding balances',
        'Great news! All balances are cleared as of ' + nowStr_());
      return { sent: true, count: 0, message: 'No outstanding balances' };
    }

    const subject = 'Credit Ledger — Outstanding Balances Report';
    const htmlBody = `
      <h2>Credit Ledger — Outstanding Balances</h2>
      <p>As of ${nowStr_()}, here are the current outstanding balances:</p>
      <table border="1" cellpadding="8" style="border-collapse: collapse; width: 100%;">
        <thead>
          <tr style="background-color: #f0f0f0;">
            <th>Name</th>
            <th>Balance (${CONFIG.CURRENCY})</th>
            <th>Last Transaction</th>
            <th>Total Credit</th>
            <th>Total Payment</th>
          </tr>
        </thead>
        <tbody>
          ${outstanding.map(p => `
            <tr>
              <td><strong>${escapeHtml_(p.name)}</strong></td>
              <td style="color: #e74c3c; font-weight: bold;">${p.balance.toFixed(3)}</td>
              <td>${p.lastTranDate || '—'}</td>
              <td>${p.totalCredit.toFixed(3)}</td>
              <td>${p.totalPayment.toFixed(3)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
      <br>
      <p><em>Total Outstanding: <strong>${outstanding.reduce((sum, p) => sum + p.balance, 0).toFixed(3)} ${CONFIG.CURRENCY}</strong></em></p>
      <p><i>Tip:</i> Open the "Credit Ledger → Open Dashboard" menu in the Sheet to collect and record payments.</p>
    `;

    MailApp.sendEmail({
      to: owner,
      subject: subject,
      htmlBody: htmlBody
    });

    return { sent: true, count: outstanding.length, total: outstanding.reduce((sum, p) => sum + p.balance, 0) };
  } catch (error) {
    console.error('Error in sendMonthlyAlerts:', error);
    throw new Error('Failed to send alerts: ' + error.message);
  }
}

// Install a monthly trigger (runs on day 28 at 09:00 Kuwait time)
function installMonthlyAlertTrigger() {
  try {
    // Clear existing triggers for this function
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === 'sendMonthlyAlerts') ScriptApp.deleteTrigger(t);
    });
    
    ScriptApp.newTrigger('sendMonthlyAlerts')
      .timeBased()
      .onMonthDay(28)         // a couple of days before month-end
      .atHour(9)
      .inTimezone(CONFIG.TIMEZONE)
      .create();
    return { ok: true, message: 'Monthly alert trigger installed (runs on 28th at 09:00)' };
  } catch (error) {
    console.error('Error installing trigger:', error);
    throw new Error('Failed to install trigger: ' + error.message);
  }
}

/** ====== Lower-level data utils ====== **/
function listPeople_() {
  try {
    const sh = ensureSheet_(CONFIG.SHEET_PEOPLE, HEADERS.PEOPLE);
    const { rows } = getTable_(sh);
    const updates = [];
    rows.forEach((row, idx) => {
      let changed = false;
      if (!row.id) {
        row.id = uuid_();
        changed = true;
      }
      if (!row.createdAt) {
        row.createdAt = nowStr_();
        changed = true;
      }
      if (!row.updatedAt) {
        row.updatedAt = row.createdAt;
        changed = true;
      }
      if (!row.active) {
        row.active = 'TRUE';
        changed = true;
      }
      if (changed) {
        updates.push({ row, index: idx });
      }
    });
    if (updates.length) {
      updates.forEach(({ row, index }) => {
        writeRows_(sh, HEADERS.PEOPLE, [row], index + 2);
      });
    }
    return rows.filter(r => String(r.active || 'TRUE').toUpperCase() !== 'FALSE');
  } catch (error) {
    console.error('Error in listPeople_:', error);
    return [];
  }
}

function listLedger_(filters) {
  try {
    const sh = ensureSheet_(CONFIG.SHEET_LEDGER, HEADERS.LEDGER);
    const { rows } = getTable_(sh);
    const fromDate = filters.dateFrom ? toDateString_(filters.dateFrom) : '';
    const toDate = filters.dateTo ? toDateString_(filters.dateTo) : '';
    const typeFilter = filters.type ? String(filters.type).toUpperCase() : '';

    return rows.filter(r => {
      if (filters.personId && r.personId !== filters.personId) return false;
      if (typeFilter && r.type !== typeFilter) return false;
      if (fromDate && (!r.tranDate || r.tranDate < fromDate)) return false;
      if (toDate && (!r.tranDate || r.tranDate > toDate)) return false;
      if (filters.hasReceipt === true  && !r.receiptFileId) return false;
      if (filters.hasReceipt === false &&  r.receiptFileId) return false;
      return true;
    }).sort((a,b) => (a.tranDate < b.tranDate ? -1 : a.tranDate > b.tranDate ? 1 : 0));
  } catch (error) {
    console.error('Error in listLedger_:', error);
    return [];
  }
}

function getTable_(sh) {
  try {
    const rng = sh.getDataRange();
    const values = rng.getValues();
    if (values.length < 1) return { headers: [], rows: [] };
    
    const headers = values[0];
    const rows = [];
    for (let i=1; i<values.length; i++) {
      const obj = {};
      headers.forEach((h,idx) => obj[h] = values[i][idx]);
      // Coerce numeric/date fields
      if (obj.amountKWD !== undefined && obj.amountKWD !== '') obj.amountKWD = Number(obj.amountKWD);
      if (obj.tranDate) obj.tranDate = toDateString_(obj.tranDate);
      rows.push(obj);
    }
    return { headers, rows };
  } catch (error) {
    console.error('Error in getTable_:', error);
    return { headers: [], rows: [] };
  }
}

function appendRow_(sh, headers, rowObj) {
  try {
    const row = headers.map(h => rowObj[h] ?? '');
    sh.appendRow(row);
  } catch (error) {
    console.error('Error in appendRow_:', error);
    throw error;
  }
}

function writeRows_(sh, headers, rowObjs, startRow) {
  try {
    const values = rowObjs.map(obj => headers.map(h => obj[h] ?? ''));
    sh.getRange(startRow, 1, values.length, headers.length).setValues(values);
  } catch (error) {
    console.error('Error in writeRows_:', error);
    throw error;
  }
}

function indexBy_(rows, key) {
  const m = {};
  if (rows && Array.isArray(rows)) {
    rows.forEach((r,i) => {
      if (r && r[key] !== undefined) m[r[key]] = i;
    });
  }
  return m;
}

function saveBase64ToDrive_(base64, filename) {
  try {
    if (!base64 || !filename) throw new Error('Missing file data');
    
    // Validate file size (5MB limit)
    const base64Data = base64.split(',')[1] || base64;
    const fileSize = (base64Data.length * 3) / 4; // Approximate size in bytes
    if (fileSize > 5 * 1024 * 1024) throw new Error('File size exceeds 5MB limit');
    
    const parts = base64.split(',');
    const data = Utilities.base64Decode(parts[1] || parts[0]);
    const contentType = (parts[0].match(/data:(.*);base64/) || [])[1] || 'application/octet-stream';
    
    // Validate file types
    const allowedTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'application/pdf'];
    if (!allowedTypes.includes(contentType)) throw new Error('Only JPEG, PNG, GIF, WebP images and PDF files are allowed');
    
    const blob = Utilities.newBlob(data, contentType, filename);
    const file = getReceiptsFolder_().createFile(blob).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
    return file.getId();
  } catch (error) {
    console.error('Error in saveBase64ToDrive_:', error);
    throw new Error('Failed to save file: ' + error.message);
  }
}

function escapeHtml_(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}
