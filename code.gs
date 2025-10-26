/**********************
 * CONFIG
 **********************/
const SPREADSHEET_ID = '';       // optional: set to force a specific spreadsheet
const SHEET_NAME = 'Licenses';
const HEADER = [
  'id',
  'name','nameAr',
  'description','descriptionAr',
  'expiryLabel','expiryLabelAr','expiryDate','expiryStatus',
  'status','fileUrl','fileName','createdAt'
];
let FOLDER_ID = '';                 // optional: preset Drive folder ID
const SHARE_FILES_PUBLIC = true;    // auto "Anyone with link → Viewer"
const MAX_UPLOAD_SIZE_BYTES = 5 * 1024 * 1024; // 5 MB safety limit
const LICENSE_HISTORY_SHEET_NAME = 'LicenseHistory';
const LICENSE_HISTORY_HEADER = [
  'id',
  'timestamp',
  'action',
  'prevExpiryLabel',
  'prevExpiryLabelAr',
  'prevExpiryDate',
  'prevExpiryStatus',
  'prevStatus',
  'prevStatusType',
  'prevFileUrl',
  'prevFileName'
];

/**********************
 * WEB APP ENTRY
 **********************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Admin Office Licenses');
}

/**********************
 * HELPERS
 **********************/
function getSheet_() {
  const configuredId = typeof SPREADSHEET_ID === 'string' ? SPREADSHEET_ID.trim() : '';
  let ss = null;

  if (configuredId) {
    try {
      ss = SpreadsheetApp.openById(configuredId);
    } catch (err) {
      throw new Error(`Unable to open spreadsheet with configured ID (${configuredId}). ${err && err.message ? err.message : err}`);
    }
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  if (!ss) {
    throw new Error('Unable to locate target spreadsheet. Set SPREADSHEET_ID or bind the script to a spreadsheet.');
  }

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const first = sh.getRange(1,1,1,HEADER.length).getValues()[0];
  if (first.length !== HEADER.length || HEADER.some((h,i)=>first[i] !== h)) {
    sh.getRange(1,1,1,HEADER.length).setValues([HEADER]);
  }
  migrateLegacySheet_(sh);
  return sh;
}
function getFolder_() {
  if (FOLDER_ID) { try { return DriveApp.getFolderById(FOLDER_ID); } catch(e) {} }
  const name = 'Admin Office Licenses';
  const it = DriveApp.getFoldersByName(name);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(name);
  FOLDER_ID = folder.getId();
  return folder;
}
function getHistorySheet_() {
  const sh = getSheet_(); // ensure spreadsheet exists
  const ss = sh.getParent();
  let history = ss.getSheetByName(LICENSE_HISTORY_SHEET_NAME);
  if (!history) {
    history = ss.insertSheet(LICENSE_HISTORY_SHEET_NAME);
  }
  const first = history.getRange(1, 1, 1, LICENSE_HISTORY_HEADER.length).getValues()[0];
  if (first.length !== LICENSE_HISTORY_HEADER.length || LICENSE_HISTORY_HEADER.some((h, i) => first[i] !== h)) {
    history.getRange(1, 1, 1, LICENSE_HISTORY_HEADER.length).setValues([LICENSE_HISTORY_HEADER]);
  }
  migrateLegacyHistory_(history);
  return history;
}
function sanitizeString_(value) {
  return String(value == null ? '' : value).trim();
}
function sanitizeUrl_(url) {
  const s = sanitizeString_(url);
  return /^https?:\/\//i.test(s) ? s : '';
}
function toIso_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(d);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const parsed = new Date(s);
  return isNaN(parsed) ? '' : Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function formatTimestamp_(value) {
  if (!value) return '';
  const date = Object.prototype.toString.call(value) === '[object Date]' ? value : new Date(value);
  if (isNaN(date)) return '';
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(date, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
}
function daysUntil_(iso) {
  if (!iso) return null;
  const tz = Session.getScriptTimeZone();
  const today = new Date(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'));
  const d = new Date(iso + 'T00:00:00');
  return Math.round((d - today) / 86400000);
}
function computeExpiryStatus_(iso) {
  if (!iso) return null;

  const dd = daysUntil_(iso);
  if (dd === null) return null;

  if (dd < 0) {
    return {
      type: 'Expired',
      label: 'Expired',
      daysUntil: dd,
      withinThreshold: false,
      mode: 'past'
    };
  }

  if (dd <= 30) {
    let label;
    if (dd === 0) {
      label = 'Upcoming (today)';
    } else if (dd === 1) {
      label = 'Upcoming (in 1 day)';
    } else {
      label = `Upcoming (in ${dd} days)`;
    }
    return {
      type: 'Upcoming',
      label,
      daysUntil: dd,
      withinThreshold: true,
      mode: dd === 0 ? 'today' : 'exact'
    };
  }

  return {
    type: 'Active',
    label: 'Active',
    daysUntil: dd,
    withinThreshold: false,
    mode: 'future'
  };
}

function computeStatus_(expiryIso) {
  const expiryStatus = computeExpiryStatus_(expiryIso);

  const defaultOverall = {
    type: 'Active',
    label: 'Active',
    daysUntil: null,
    withinThreshold: false,
    mode: 'none'
  };

  const overall = expiryStatus
    ? Object.assign({}, expiryStatus)
    : Object.assign({}, defaultOverall);

  return { overall, expiry: expiryStatus };
}

function inferStatusType_(label) {
  const normalized = sanitizeString_(label).toLowerCase();
  if (!normalized) return '';
  if (normalized === 'expired') return 'Expired';
  if (normalized === 'active') return 'Active';
  if (normalized.startsWith('upcoming')) return 'Upcoming';
  return '';
}

function mergeStoredStatus_(computed, storedLabel) {
  const label = sanitizeString_(storedLabel);
  if (computed) {
    const merged = Object.assign({}, computed);
    if (label) {
      merged.label = label;
      if (!merged.type) {
        const inferred = inferStatusType_(label);
        if (inferred) merged.type = inferred;
      }
    }
    return merged;
  }
  if (label) {
    const type = inferStatusType_(label);
    return {
      type: type || 'Active',
      label,
      daysUntil: null,
      withinThreshold: type === 'Upcoming',
      mode: 'legacy'
    };
  }
  return null;
}

function formatStatusInfo_(status, defaults) {
  const result = {
    type: '',
    label: '',
    daysUntil: null,
    withinThreshold: false,
    mode: ''
  };

  if (status) {
    if (status.type) result.type = String(status.type);
    if (status.label != null) result.label = sanitizeString_(status.label);
    if (status.daysUntil != null && isFinite(status.daysUntil)) {
      result.daysUntil = Number(status.daysUntil);
    }
    if (status.mode != null) result.mode = String(status.mode);
    if (status.withinThreshold != null) {
      result.withinThreshold = !!status.withinThreshold;
    }
  }

  const fallback = defaults || {};
  if (!result.type && fallback.type) {
    result.type = String(fallback.type);
  }
  if (!result.label && fallback.label != null) {
    result.label = sanitizeString_(fallback.label);
  }
  if (result.daysUntil == null && fallback.daysUntil != null && isFinite(fallback.daysUntil)) {
    result.daysUntil = Number(fallback.daysUntil);
  }
  if (!result.mode && fallback.mode != null) {
    result.mode = String(fallback.mode);
  }
  if (!result.withinThreshold && fallback.withinThreshold) {
    result.withinThreshold = !!fallback.withinThreshold;
  }

  if (!result.type && result.label) {
    const inferred = inferStatusType_(result.label);
    if (inferred) result.type = inferred;
  }

  if (!result.type && !result.label) {
    return null;
  }

  return result;
}
function parseDriveFileId_(url) {
  if (!url) return '';
  let m = String(url).match(/\/file\/d\/([a-zA-Z0-9_-]+)/); if (m) return m[1];
  m = String(url).match(/[?&]id=([a-zA-Z0-9_-]+)/); if (m) return m[1];
  m = String(url).match(/\/uc\?id=([a-zA-Z0-9_-]+)/); if (m) return m[1];
  return '';
}
function makePreviewUrl_(fileUrl) {
  const safeUrl = sanitizeUrl_(fileUrl);
  const id = parseDriveFileId_(safeUrl);
  if (!id) return safeUrl;

  const allowedQueryKeys = ['resourcekey', 'usp', 'export'];
  const allowedQueryLookup = allowedQueryKeys.reduce((acc, key) => {
    acc[key] = true;
    return acc;
  }, {});
  let preservedQuery = '';

  if (safeUrl) {
    let queryString = '';
    const questionIdx = safeUrl.indexOf('?');
    if (questionIdx !== -1) {
      queryString = safeUrl.substring(questionIdx + 1);
      const hashIdx = queryString.indexOf('#');
      if (hashIdx !== -1) {
        queryString = queryString.substring(0, hashIdx);
      }
    }

    if (queryString) {
      const preservedPairs = [];
      queryString.split('&').forEach(part => {
        if (!part) return;
        const eqIdx = part.indexOf('=');
        const rawKey = eqIdx === -1 ? part : part.substring(0, eqIdx);
        const rawValue = eqIdx === -1 ? '' : part.substring(eqIdx + 1);
        let key;
        try {
          key = decodeURIComponent(rawKey || '').trim();
        } catch (err) {
          key = String(rawKey || '').trim();
        }
        if (!key) return;
        const normalizedKey = key.toLowerCase();
        if (!allowedQueryLookup[normalizedKey]) return;
        let value = '';
        if (rawValue) {
          try {
            value = decodeURIComponent(rawValue);
          } catch (err) {
            value = String(rawValue);
          }
        }
        preservedPairs.push(`${encodeURIComponent(key)}=${encodeURIComponent(value)}`);
      });
      if (preservedPairs.length) {
        preservedQuery = `?${preservedPairs.join('&')}`;
      }
    }
  }

  const previewUrl = `https://drive.google.com/file/d/${id}/preview${preservedQuery}`;

  if (
    typeof console !== 'undefined' &&
    safeUrl &&
    safeUrl.indexOf('resourcekey=') !== -1 &&
    previewUrl.indexOf('resourcekey=') === -1 &&
    typeof console.warn === 'function'
  ) {
    console.warn('makePreviewUrl_ dropped resourcekey parameter while rebuilding preview URL.', safeUrl, previewUrl);
  }

  return previewUrl;
}

(function previewUrlRegressionCheck_() {
  const sampleId = 'SAMPLE_ID';
  const resourceKeySamples = [
    { query: 'resourcekey=1-testKey', label: 'resourcekey' },
    { query: 'resourceKey=1-testKeyCamel', label: 'resourceKey' }
  ];

  resourceKeySamples.forEach(sample => {
    const sampleUrl = `https://drive.google.com/file/d/${sampleId}/view?usp=drivesdk&${sample.query}`;
    const preview = makePreviewUrl_(sampleUrl);
    if (typeof console !== 'undefined') {
      const containsResourceKey = preview.indexOf(sample.query) !== -1;
      if (typeof console.assert === 'function') {
        console.assert(
          containsResourceKey,
          `makePreviewUrl_ should retain ${sample.label} query parameter for preview URLs.`
        );
      } else if (!containsResourceKey && typeof console.warn === 'function') {
        console.warn(
          `makePreviewUrl_ failed to retain ${sample.label} during regression check.`,
          preview
        );
      }
    }
  });
})();
function getAllRows_() {
  const sh = getSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const lastColumn = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0];
  const indexMap = buildHeaderIndex_(headers);
  const values = sh.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const records = [];
  values.forEach((row, idx) => {
    const record = buildRecordFromRow_(row, indexMap, idx);
    if (record && (record.id || record.name)) {
      records.push(record);
    }
  });
  return records;
}
function findRowById_(id) {
  const target = String(id || '').trim();
  if (!target) return null;
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const lastColumn = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0];
  const indexMap = buildHeaderIndex_(headers);
  const range = sh.getRange(2, 1, last - 1, lastColumn);
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const record = buildRecordFromRow_(row, indexMap, i);
    if (record && String(record.id) === target) {
      const normalizedRow = HEADER.map((_, idx) => row[idx] != null ? row[idx] : '');
      return {
        rowNumber: i + 2,
        values: normalizedRow,
        object: record
      };
    }
  }
  return null;
}
function nextId_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return '1';
  const values = sh.getRange(2, 1, last - 1, 1).getValues();
  const max = values.reduce((m, r) => {
    const v = Number(r[0]);
    return isFinite(v) && v > m ? v : m;
  }, 0);
  return String(max + 1);
}
function estimateBase64Bytes_(b64) {
  const cleaned = sanitizeString_(b64).replace(/=+$/, '');
  return Math.ceil(cleaned.length * 3 / 4);
}
function decodeBase64Safely_(b64, context) {
  try {
    return Utilities.base64Decode(b64);
  } catch (err) {
    const label = context ? ` (${context})` : '';
    Logger.log('base64 decode failed%s: %s', label, err && err.stack ? err.stack : err);
    throw new Error('Uploaded file data is invalid or corrupted. Please reselect the file and try again.');
  }
}
function ensureDate_(value, fallback) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }
  if (value != null) {
    const parsed = new Date(value);
    if (!isNaN(parsed)) return parsed;
  }
  return fallback instanceof Date && !isNaN(fallback) ? fallback : new Date();
}
function buildHeaderIndex_(headers) {
  const index = {};
  headers.forEach((raw, i) => {
    const key = sanitizeString_(raw);
    if (key && !(key in index)) {
      index[key] = i;
    }
  });
  return index;
}
function valueFromRow_(row, indexMap, candidates, fallback) {
  for (let i = 0; i < candidates.length; i++) {
    const key = sanitizeString_(candidates[i]);
    if (!key) continue;
    const idx = indexMap.hasOwnProperty(key) ? indexMap[key] : -1;
    if (idx != null && idx >= 0 && idx < row.length) {
      const value = row[idx];
      if (value !== '' && value != null) {
        return value;
      }
    }
  }
  return fallback;
}
function buildRecordFromRow_(row, indexMap, rowIndex) {
  const idRaw = valueFromRow_(row, indexMap, ['id'], '');
  const id = sanitizeString_(idRaw) || String(rowIndex + 1);
  const name = sanitizeString_(valueFromRow_(row, indexMap, ['name'], ''));
  const nameAr = sanitizeString_(valueFromRow_(row, indexMap, ['nameAr'], ''));
  const description = sanitizeString_(valueFromRow_(row, indexMap, ['description'], ''));
  const descriptionAr = sanitizeString_(valueFromRow_(row, indexMap, ['descriptionAr'], ''));
  const expiryLabel = sanitizeString_(valueFromRow_(row, indexMap, ['expiryLabel', 'exp1Label', 'expiry'], ''));
  const expiryLabelAr = sanitizeString_(valueFromRow_(row, indexMap, ['expiryLabelAr', 'exp1LabelAr'], ''));
  const expiryDate = toIso_(valueFromRow_(row, indexMap, ['expiryDate', 'exp1Date', 'expiry'], ''));
  const statusSnapshot = computeStatus_(expiryDate);
  const storedExpiryStatus = sanitizeString_(valueFromRow_(row, indexMap, ['expiryStatus', 'exp1Status'], ''));
  const storedOverallStatus = sanitizeString_(valueFromRow_(row, indexMap, ['status', 'overallStatus'], ''));
  const mergedExpiry = mergeStoredStatus_(statusSnapshot.expiry, storedExpiryStatus);
  const mergedOverall = mergeStoredStatus_(statusSnapshot.overall, storedOverallStatus);
  const expiryStatusInfo = formatStatusInfo_(mergedExpiry);
  const overallInfo = formatStatusInfo_(mergedOverall, statusSnapshot.overall) || formatStatusInfo_(statusSnapshot.overall);
  const fileUrl = sanitizeUrl_(valueFromRow_(row, indexMap, ['fileUrl', 'file', 'documentUrl', 'document'], ''));
  const fileName = sanitizeString_(valueFromRow_(row, indexMap, ['fileName', 'documentName', 'filename'], '')) || name;
  const createdAt = ensureDate_(valueFromRow_(row, indexMap, ['createdAt', 'created', 'timestamp'], new Date()), new Date());

  return {
    id,
    name,
    nameAr,
    description,
    descriptionAr,
    expiryLabel,
    expiryLabelAr,
    expiryDate,
    expiryStatus: expiryStatusInfo && expiryStatusInfo.label ? expiryStatusInfo.label : '',
    expiryStatusInfo,
    status: overallInfo && overallInfo.label ? overallInfo.label : '',
    statusInfo: overallInfo,
    statusType: overallInfo && overallInfo.type ? overallInfo.type : '',
    statusDaysUntil: overallInfo && overallInfo.daysUntil != null ? overallInfo.daysUntil : null,
    fileUrl,
    filePreviewUrl: makePreviewUrl_(fileUrl),
    fileName,
    createdAt
  };
}
function buildHistoryRecordFromRow_(row, indexMap) {
  const id = sanitizeString_(valueFromRow_(row, indexMap, ['id'], ''));
  if (!id) return null;
  const timestamp = ensureDate_(valueFromRow_(row, indexMap, ['timestamp', 'date'], new Date()), new Date());
  const action = sanitizeString_(valueFromRow_(row, indexMap, ['action'], ''));
  const prevExpiryLabel = sanitizeString_(valueFromRow_(row, indexMap, ['prevExpiryLabel', 'expiryLabel', 'exp1Label'], ''));
  const prevExpiryLabelAr = sanitizeString_(valueFromRow_(row, indexMap, ['prevExpiryLabelAr', 'expiryLabelAr', 'exp1LabelAr'], ''));
  const prevExpiryDate = toIso_(valueFromRow_(row, indexMap, ['prevExpiryDate', 'expiryDate', 'exp1Date'], ''));
  const prevExpiryStatus = sanitizeString_(valueFromRow_(row, indexMap, ['prevExpiryStatus', 'expiryStatus', 'exp1Status'], ''));
  const prevStatus = sanitizeString_(valueFromRow_(row, indexMap, ['prevStatus', 'status'], ''));
  const prevStatusType = sanitizeString_(valueFromRow_(row, indexMap, ['prevStatusType', 'statusType'], ''));
  const prevFileUrl = sanitizeUrl_(valueFromRow_(row, indexMap, ['prevFileUrl', 'fileUrl', 'documentUrl'], ''));
  const prevFileName = sanitizeString_(valueFromRow_(row, indexMap, ['prevFileName', 'fileName', 'documentName'], ''));

  return {
    id,
    timestamp: formatTimestamp_(timestamp),
    action,
    prevExpiryLabel,
    prevExpiryLabelAr,
    prevExpiryDate,
    prevExpiryStatus,
    prevStatus,
    prevStatusType,
    prevFileUrl,
    prevFilePreviewUrl: makePreviewUrl_(prevFileUrl),
    prevFileName
  };
}
function migrateLegacySheet_(sh) {
  const totalColumns = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const lastColumn = sh.getLastColumn();
  const width = lastColumn;
  const rows = sh.getRange(2, 1, lastRow - 1, width).getValues();
  const migrated = rows.map(row => legacyRowToCurrent_(row));
  sh.getRange(2, 1, migrated.length, HEADER.length).setValues(migrated);
  if (lastColumn > HEADER.length) {
    sh.deleteColumns(HEADER.length + 1, lastColumn - HEADER.length);
  }
}
function legacyRowToCurrent_(row) {
  if (!Array.isArray(row)) row = [];
  const valueAt = (idx, fallback) => {
    if (row[idx] == null || row[idx] === '') return fallback;
    return row[idx];
  };
  const id = sanitizeString_(valueAt(0, ''));
  const name = sanitizeString_(valueAt(1, ''));
  const nameAr = sanitizeString_(valueAt(2, ''));
  const description = sanitizeString_(valueAt(3, ''));
  const descriptionAr = sanitizeString_(valueAt(4, ''));
  const rawLabels = [
    valueAt(5, ''),
    valueAt(9, ''),
    valueAt(10, ''),
    valueAt(11, '')
  ].map(sanitizeString_).filter(Boolean);
  const expiryLabel = rawLabels.length ? rawLabels[0] : '';
  const expiryLabelAr = sanitizeString_(valueAt(6, ''));
  const expiryDate = toIso_(valueAt(7, ''));
  const statusInfo = computeStatus_(expiryDate);
  const legacyExpiryStatus = sanitizeString_(valueAt(8, ''));
  const expiryStatus = legacyExpiryStatus || (statusInfo.expiry && statusInfo.expiry.label ? statusInfo.expiry.label : '');
  const legacyOverallStatus = sanitizeString_(valueAt(13, ''));
  const overallStatus = legacyOverallStatus || (statusInfo.overall && statusInfo.overall.label ? statusInfo.overall.label : '');
  const fileUrl = sanitizeUrl_(valueAt(14, ''));
  const fileName = sanitizeString_(valueAt(15, '') || name);
  const createdAt = ensureDate_(valueAt(16, new Date()), new Date());
  return [
    id,
    name,
    nameAr,
    description,
    descriptionAr,
    expiryLabel,
    expiryLabelAr,
    expiryDate,
    expiryStatus,
    overallStatus,
    fileUrl,
    fileName,
    createdAt
  ];
}
function migrateLegacyHistory_(history) {
  const totalColumns = history.getLastColumn();
  const last = history.getLastRow();
  if (last < 2) return;
  const lastColumn = history.getLastColumn();
  const width = lastColumn;
  const values = history.getRange(2, 1, last - 1, width).getValues();
  const migrated = values.map(legacyHistoryRowToCurrent_);
  history.getRange(2, 1, migrated.length, LICENSE_HISTORY_HEADER.length).setValues(migrated);
  if (lastColumn > LICENSE_HISTORY_HEADER.length) {
    history.deleteColumns(LICENSE_HISTORY_HEADER.length + 1, lastColumn - LICENSE_HISTORY_HEADER.length);
  }
}
function legacyHistoryRowToCurrent_(row) {
  if (!Array.isArray(row)) row = [];
  const valueAt = (idx, fallback) => {
    if (row[idx] == null || row[idx] === '') return fallback;
    return row[idx];
  };
  const timestamp = ensureDate_(row[1], new Date());
  const expiryLabel = sanitizeString_(valueAt(3, valueAt(7, '')));
  const expiryLabelAr = sanitizeString_(valueAt(4, valueAt(8, '')));
  const expiryDate = toIso_(valueAt(5, valueAt(9, '')));
  const statusInfo = computeStatus_(expiryDate);
  const legacyExpiryStatus = sanitizeString_(valueAt(6, valueAt(10, '')));
  const expiryStatus = legacyExpiryStatus || (statusInfo.expiry && statusInfo.expiry.label ? statusInfo.expiry.label : '');
  const legacyOverallStatus = sanitizeString_(valueAt(11, ''));
  const legacyOverallType = sanitizeString_(valueAt(12, ''));
  const overallStatus = legacyOverallStatus || (statusInfo.overall && statusInfo.overall.label ? statusInfo.overall.label : '');
  const overallType = legacyOverallType || (statusInfo.overall && statusInfo.overall.type ? statusInfo.overall.type : '');
  const fileUrl = sanitizeUrl_(valueAt(13, ''));
  const fileName = sanitizeString_(valueAt(14, ''));
  return [
    sanitizeString_(row[0]),
    timestamp,
    sanitizeString_(row[2]),
    expiryLabel,
    expiryLabelAr,
    expiryDate,
    expiryStatus,
    overallStatus,
    overallType,
    fileUrl,
    fileName
  ];
}
function normalizeUploadInput_(obj) {
  const normalized = {
    name: sanitizeString_(obj && obj.name),
    nameAr: sanitizeString_(obj && obj.nameAr),
    description: sanitizeString_(obj && obj.description),
    descriptionAr: sanitizeString_(obj && obj.descriptionAr),
    expiryLabel: sanitizeString_(obj && obj.expiryLabel),
    expiryLabelAr: sanitizeString_(obj && obj.expiryLabelAr),
    expiryDate: toIso_(obj && obj.expiryDate),
    file: null
  };

  if (!normalized.name) {
    throw new Error('Name is required.');
  }

  const fileB64 = sanitizeString_(obj && obj.fileB64);
  if (fileB64) {
    const approxBytes = estimateBase64Bytes_(fileB64);
    if (approxBytes > MAX_UPLOAD_SIZE_BYTES) {
      throw new Error('File is too large. Maximum size is 5 MB.');
    }
    const fileName = sanitizeString_(obj && obj.fileName) || normalized.name;
    normalized.file = {
      b64: fileB64,
      name: fileName,
      type: sanitizeString_(obj && obj.fileType) || MimeType.PDF
    };
  }

  return normalized;
}
function recordHistoryEntry_(id, existingObj, action) {
  if (!existingObj) return;
  const history = getHistorySheet_();
  const timestamp = new Date();
  history.appendRow([
    sanitizeString_(id),
    timestamp,
    sanitizeString_(action) || 'update',
    sanitizeString_(existingObj.expiryLabel),
    sanitizeString_(existingObj.expiryLabelAr),
    toIso_(existingObj.expiryDate),
    sanitizeString_(existingObj.expiryStatusInfo && existingObj.expiryStatusInfo.label ? existingObj.expiryStatusInfo.label : existingObj.expiryStatus),
    sanitizeString_(existingObj.statusInfo && existingObj.statusInfo.label ? existingObj.statusInfo.label : existingObj.status),
    sanitizeString_(existingObj.statusInfo && existingObj.statusInfo.type ? existingObj.statusInfo.type : existingObj.statusType),
    sanitizeUrl_(existingObj.fileUrl),
    sanitizeString_(existingObj.fileName || existingObj.name)
  ]);
}

function getLicenseHistoryIndex_() {
  const sheet = getHistorySheet_();
  const last = sheet.getLastRow();
  if (last < 2) {
    return new Set();
  }
  const values = sheet.getRange(2, 1, last - 1, 1).getValues();
  const ids = new Set();
  values.forEach(row => {
    const id = sanitizeString_(row[0]);
    if (id) ids.add(id);
  });
  return ids;
}

/**********************
 * READ API
 **********************/
function getDashboardData(q) {
  try {
    const safeLogValue = value => {
      try {
        return JSON.stringify(value);
      } catch (jsonErr) {
        return String(value);
      }
    };

    const all = getAllRows_();
    if (!Array.isArray(all)) {
      Logger.log('Unexpected getAllRows_() result: %s', safeLogValue(all));
      throw new Error('Unable to serialise dashboard data: rows is not an array');
    }

    let historyIds;
    try {
      historyIds = getLicenseHistoryIndex_();
    } catch (err) {
      historyIds = new Set();
    }
    all.forEach(row => {
      const key = sanitizeString_(row.id);
      row.hasHistory = key ? historyIds.has(key) : false;
    });

    const query = (q || '').trim().toLowerCase();
    const filtered = !query ? all : all.filter(r => {
      return [
        r.id,
        r.name, r.nameAr,
        r.description, r.descriptionAr,
        r.expiryLabel, r.expiryLabelAr,
        r.fileName
      ]
        .map(x => String(x||'').toLowerCase())
        .some(s => s.includes(query));
    });

    if (!Array.isArray(filtered)) {
      Logger.log('Unexpected filtered rows result: %s', safeLogValue(filtered));
      throw new Error('Unable to serialise dashboard data: filtered rows is not an array');
    }

    const stats = arr => {
      const counts = {
        total: Array.isArray(arr) ? arr.length : 0,
        expired: { total: 0 },
        upcoming: { total: 0 },
        active: { total: 0 }
      };

      const bucketKeyForType = type => {
        if (type === 'Expired') return 'expired';
        if (type === 'Upcoming') return 'upcoming';
        if (type === 'Active') return 'active';
        return '';
      };

      const normalizeType = value => {
        if (!value) return '';
        const inferred = inferStatusType_(value);
        if (inferred) return inferred;
        const sanitized = sanitizeString_(value);
        if (!sanitized) return '';
        const lowered = sanitized.toLowerCase();
        if (lowered === 'expired') return 'Expired';
        if (lowered === 'upcoming') return 'Upcoming';
        if (lowered === 'active') return 'Active';
        return '';
      };

      const resolveStatusType = row => {
        const candidates = [
          row && row.expiryStatusInfo && row.expiryStatusInfo.type,
          row && row.expiryStatusInfo && row.expiryStatusInfo.label,
          row && row.expiryStatus,
          row && row.statusInfo && row.statusInfo.type,
          row && row.statusInfo && row.statusInfo.label,
          row && row.status
        ];
        for (let i = 0; i < candidates.length; i++) {
          const type = normalizeType(candidates[i]);
          if (type) return type;
        }
        return 'Active';
      };

      (Array.isArray(arr) ? arr : []).forEach(r => {
        const type = resolveStatusType(r);
        const bucketKey = bucketKeyForType(type) || 'active';
        if (!counts[bucketKey]) counts[bucketKey] = { total: 0 };
        counts[bucketKey].total += 1;
      });

      return counts;
    };

    const countsAll = stats(all);
    const countsFiltered = stats(filtered);

    const isPlainObject = value => {
      if (!value || typeof value !== 'object') return false;
      const proto = Object.getPrototypeOf(value);
      return proto === Object.prototype || proto === null;
    };

    if (!isPlainObject(countsAll)) {
      Logger.log('Unexpected countsAll result: %s', safeLogValue(countsAll));
      throw new Error('Unable to serialise dashboard data: countsAll is not a plain object');
    }

    if (!isPlainObject(countsFiltered)) {
      Logger.log('Unexpected countsFiltered result: %s', safeLogValue(countsFiltered));
      throw new Error('Unable to serialise dashboard data: countsFiltered is not a plain object');
    }

    const key = r => {
      const date = r && r.expiryDate ? String(r.expiryDate) : '';
      return (date || '9999-12-31') + '|' + (r.name||'');
    };
    filtered.sort((a,b)=> key(a) < key(b) ? -1 : key(a) > key(b) ? 1 : 0);

    return { rows: filtered, countsAll, countsFiltered };
  } catch (err) {
    Logger.log('getDashboardData error: %s', err && err.stack ? err.stack : err);
    throw err;
  }
}

/**********************
 * WRITE API (base64)
 **********************/
function uploadDocument(obj) {
  try {
    const sh = getSheet_();
    const id = nextId_();
    const now = new Date();
    const data = normalizeUploadInput_(obj);

    let fileUrl = '';
    let fileName = '';

    if (data.file) {
      const folder = getFolder_();
      const bytes = decodeBase64Safely_(data.file.b64, 'uploadDocument');
      const mime = data.file.type || MimeType.PDF;
      const blob = Utilities.newBlob(bytes, mime, data.file.name);
      const created = folder.createFile(blob).setName(data.file.name);
      if (SHARE_FILES_PUBLIC) created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = sanitizeUrl_(created.getUrl());
      fileName = data.file.name;
    }

    const status = computeStatus_(data.expiryDate);
    const overall = status.overall || { label: 'Active' };

    sh.appendRow([
      id,
      data.name,
      data.nameAr,
      data.description,
      data.descriptionAr,
      data.expiryLabel,
      data.expiryLabelAr,
      data.expiryDate,
      status.expiry && status.expiry.label ? status.expiry.label : '',
      overall.label,
      fileUrl,
      fileName || data.name,
      now
    ]);

    const safeFileUrl = sanitizeUrl_(fileUrl);
    return {
      ok: true,
      id,
      fileUrl: safeFileUrl,
      filePreviewUrl: makePreviewUrl_(safeFileUrl),
      name: data.name,
      nameAr: data.nameAr,
      description: data.description,
      descriptionAr: data.descriptionAr,
      expiryLabel: data.expiryLabel,
      expiryLabelAr: data.expiryLabelAr,
      expiryDate: data.expiryDate,
      expiryStatus: status.expiry && status.expiry.label ? status.expiry.label : ''
    };
  } catch (err) {
    return { ok:false, error:String(err && err.message || err) };
  }
}
/**
 * Updates an existing license record and records a snapshot of the previous values.
 * @param {Object} obj normalized license payload including the record `id`.
 * @returns {Object} response consumed by the web UI.
 */
function updateLicense(obj) {
  try {
    const id = sanitizeString_(obj && obj.id);
    if (!id) {
      throw new Error('Record ID is required.');
    }

    const found = findRowById_(id);
    if (!found) {
      throw new Error('Record not found.');
    }

    const sh = getSheet_();
    const data = normalizeUploadInput_(obj);
    const mode = sanitizeString_(obj && obj.mode).toLowerCase();
    const action = mode === 'renew' ? 'renew' : 'update';

    let fileUrl = sanitizeUrl_(found.object.fileUrl);
    let fileName = sanitizeString_(found.object.fileName || found.object.name);

    if (data.file) {
      const folder = getFolder_();
      const bytes = decodeBase64Safely_(data.file.b64, 'updateLicense');
      const mime = data.file.type || MimeType.PDF;
      const blob = Utilities.newBlob(bytes, mime, data.file.name);
      const created = folder.createFile(blob).setName(data.file.name);
      if (SHARE_FILES_PUBLIC) {
        created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
      fileUrl = sanitizeUrl_(created.getUrl());
      fileName = data.file.name;
    }

    const status = computeStatus_(data.expiryDate);
    const overall = status.overall || { label: 'Active', type: 'Active' };
    const expiryStatusLabel = status.expiry && status.expiry.label ? status.expiry.label : '';

    const createdAtIndex = HEADER.indexOf('createdAt');
    const existingCreated = createdAtIndex >= 0 ? found.values[createdAtIndex] : new Date();

    // Capture the state before it is overwritten so history always reflects prior values.
    recordHistoryEntry_(id, found.object, action);

    const rowValues = [
      id,
      data.name,
      data.nameAr,
      data.description,
      data.descriptionAr,
      data.expiryLabel,
      data.expiryLabelAr,
      data.expiryDate,
      expiryStatusLabel,
      overall.label,
      fileUrl,
      fileName || data.name,
      existingCreated
    ];

    sh.getRange(found.rowNumber, 1, 1, HEADER.length).setValues([rowValues]);

    const safeFileUrl = sanitizeUrl_(fileUrl);
    return {
      ok: true,
      id,
      action,
      fileUrl: safeFileUrl,
      filePreviewUrl: makePreviewUrl_(safeFileUrl),
      name: data.name,
      nameAr: data.nameAr,
      description: data.description,
      descriptionAr: data.descriptionAr,
      expiryLabel: data.expiryLabel,
      expiryLabelAr: data.expiryLabelAr,
      expiryDate: data.expiryDate,
      hasHistory: true
    };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

/**
 * Convenience alias that enforces `mode:"renew"` semantics when renewing a license.
 * @param {Object} obj payload accepted by {@link updateLicense}.
 */
function renewLicense(obj) {
  const payload = obj ? Object.assign({}, obj, { mode: 'renew' }) : { mode: 'renew' };
  return updateLicense(payload);
}

/**
 * Returns prior revisions of a license stored in the history sheet, newest-first.
 * @param {string|number} id license identifier.
 * @returns {Array<Object>} chronological entries for the UI timeline.
 */
function getLicenseHistory(id) {
  const target = sanitizeString_(id);
  if (!target) {
    throw new Error('Record ID is required.');
  }
  const sheet = getHistorySheet_();
  const last = sheet.getLastRow();
  if (last < 2) {
    return [];
  }
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const indexMap = buildHeaderIndex_(headers);
  const values = sheet.getRange(2, 1, last - 1, lastColumn).getValues();
  const filtered = values
    .map(row => buildHistoryRecordFromRow_(row, indexMap))
    .filter(entry => entry && sanitizeString_(entry.id) === target);
  filtered.sort((a, b) => {
    const da = new Date(a.timestamp);
    const db = new Date(b.timestamp);
    if (isNaN(db) && isNaN(da)) return 0;
    if (isNaN(db)) return -1;
    if (isNaN(da)) return 1;
    return db.getTime() - da.getTime();
  });
  return filtered;
}

// Backward compatible aliases for existing client integrations.
function updateDocument(obj) {
  return updateLicense(obj);
}

function getDocumentHistory(id) {
  return getLicenseHistory(id);
}
