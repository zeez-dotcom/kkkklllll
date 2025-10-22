/**********************
 * CONFIG
 **********************/
const SPREADSHEET_ID = '';       // optional: set to force a specific spreadsheet
const SHEET_NAME = 'Licenses';
const HEADER = [
  'id',
  'name','nameAr',
  'description','descriptionAr',
  'exp1Label','exp1LabelAr','exp1Date','exp1Status',
  'exp2Label','exp2LabelAr','exp2Date','exp2Status',
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
  'prevExp1Label',
  'prevExp1LabelAr',
  'prevExp1Date',
  'prevExp1Status',
  'prevExp2Label',
  'prevExp2LabelAr',
  'prevExp2Date',
  'prevExp2Status',
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
  migrateLegacyHistoryRows_(history);
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

function isDateLike_(value) {
  if (!value) return false;
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return !isNaN(value);
  }
  const str = String(value);
  return /^\d{4}-\d{2}-\d{2}/.test(str);
}

function normalizeHistoryRow_(row) {
  const width = LICENSE_HISTORY_HEADER.length;
  const values = new Array(width).fill('');
  for (let i = 0; i < width && i < row.length; i++) {
    values[i] = row[i];
  }

  const newFileUrl = sanitizeUrl_(values[13]);
  const legacyFileUrl = sanitizeUrl_(row[5]) || sanitizeUrl_(row[6]);
  const labelArLooksDate = isDateLike_(row[4]);
  const exp1DateLooksWrong = !isDateLike_(row[5]) && !!row[5];
  const legacyDetected = ((!newFileUrl && !!legacyFileUrl) || (labelArLooksDate && exp1DateLooksWrong));

  if (legacyDetected) {
    values[3] = sanitizeString_(row[3]);
    values[4] = '';
    values[5] = row[4] || '';
    values[6] = sanitizeString_(row[5]);
    values[7] = '';
    values[8] = '';
    values[9] = '';
    values[10] = '';
    values[11] = sanitizeString_(row[5]);
    values[12] = '';
    values[13] = legacyFileUrl;
    values[14] = sanitizeString_(row[7]);
    return { values, migrated: true };
  }

  return { values, migrated: false };
}

function migrateLegacyHistoryRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) {
    return;
  }
  const width = LICENSE_HISTORY_HEADER.length;
  const range = sheet.getRange(2, 1, last - 1, width);
  const existing = range.getValues();
  let mutated = false;
  const upgraded = existing.map(row => {
    const normalized = normalizeHistoryRow_(row);
    if (normalized.migrated) {
      mutated = true;
    }
    return normalized.values;
  });
  if (mutated) {
    range.setValues(upgraded);
  }
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

function computeStatus_(exp1Iso, exp2Iso) {
  const exp1Status = computeExpiryStatus_(exp1Iso);
  const exp2Status = computeExpiryStatus_(exp2Iso);

  const statuses = [exp1Status, exp2Status].filter(Boolean);
  const defaultOverall = {
    type: 'Active',
    label: 'Active',
    daysUntil: null,
    withinThreshold: false,
    mode: 'none'
  };

  if (!statuses.length) {
    return { overall: defaultOverall, exp1: exp1Status, exp2: exp2Status };
  }

  const severity = status => {
    if (!status || !status.type) return 3;
    if (status.type === 'Expired') return 0;
    if (status.type === 'Upcoming') return 1;
    if (status.type === 'Active') return 2;
    return 3;
  };

  statuses.sort((a, b) => {
    const sa = severity(a);
    const sb = severity(b);
    if (sa !== sb) return sa - sb;
    const da = a && a.daysUntil != null ? a.daysUntil : Number.POSITIVE_INFINITY;
    const db = b && b.daysUntil != null ? b.daysUntil : Number.POSITIVE_INFINITY;
    return da - db;
  });

  const overall = Object.assign({}, statuses[0]);

  return { overall, exp1: exp1Status, exp2: exp2Status };
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
  return id ? `https://drive.google.com/file/d/${id}/preview` : safeUrl;
}
function getAllRows_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last-1, HEADER.length).getValues();
  return values
    .filter(r => String(r[0]||'').trim() !== '' || String(r[1]||'').trim() !== '')
    .map(normalizeRow_);
}
function findRowById_(id) {
  const target = String(id || '').trim();
  if (!target) return null;
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const range = sh.getRange(2, 1, last - 1, HEADER.length);
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === target) {
      return {
        rowNumber: i + 2,
        values: values[i],
        object: normalizeRow_(values[i])
      };
    }
  }
  return null;
}
function normalizeRow_(row) {
  const obj = {};
  HEADER.forEach((h, i) => obj[h] = row[i] ?? '');
  obj.exp1Date = toIso_(obj.exp1Date);
  obj.exp2Date = toIso_(obj.exp2Date);
  const exp1Stored = obj.exp1Status;
  const exp2Stored = obj.exp2Status;
  const overallStored = obj.status;
  const status = computeStatus_(obj.exp1Date, obj.exp2Date);
  const mergedExp1 = mergeStoredStatus_(status.exp1, exp1Stored);
  const mergedExp2 = mergeStoredStatus_(status.exp2, exp2Stored);
  const mergedOverall = mergeStoredStatus_(status.overall, overallStored);
  const defaultOverall = {
    type: 'Active',
    label: 'Active',
    daysUntil: null,
    withinThreshold: false,
    mode: 'none'
  };
  const exp1Info = formatStatusInfo_(mergedExp1);
  const exp2Info = formatStatusInfo_(mergedExp2);
  const overallInfo = formatStatusInfo_(mergedOverall, defaultOverall) || Object.assign({}, defaultOverall);
  obj.exp1StatusInfo = exp1Info;
  obj.exp2StatusInfo = exp2Info;
  obj.exp1Status = sanitizeString_(exp1Stored);
  obj.exp2Status = sanitizeString_(exp2Stored);
  obj.status = overallInfo.label;
  obj.statusInfo = overallInfo;
  obj.statusType = overallInfo.type;
  obj.statusDaysUntil = overallInfo.daysUntil;
  obj.name = sanitizeString_(obj.name);
  obj.nameAr = sanitizeString_(obj.nameAr);
  obj.description = sanitizeString_(obj.description);
  obj.descriptionAr = sanitizeString_(obj.descriptionAr);
  obj.exp1Label = sanitizeString_(obj.exp1Label);
  obj.exp1LabelAr = sanitizeString_(obj.exp1LabelAr);
  obj.exp2Label = sanitizeString_(obj.exp2Label);
  obj.exp2LabelAr = sanitizeString_(obj.exp2LabelAr);
  obj.fileUrl = sanitizeUrl_(obj.fileUrl);
  obj.filePreviewUrl = makePreviewUrl_(obj.fileUrl);
  obj.fileName = sanitizeString_(obj.fileName || obj.name);
  return obj;
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
function normalizeUploadInput_(obj) {
  const normalized = {
    name: sanitizeString_(obj && obj.name),
    nameAr: sanitizeString_(obj && obj.nameAr),
    description: sanitizeString_(obj && obj.description),
    descriptionAr: sanitizeString_(obj && obj.descriptionAr),
    exp1Label: sanitizeString_(obj && obj.exp1Label),
    exp1LabelAr: sanitizeString_(obj && obj.exp1LabelAr),
    exp2Label: sanitizeString_(obj && obj.exp2Label),
    exp2LabelAr: sanitizeString_(obj && obj.exp2LabelAr),
    exp1Date: toIso_(obj && obj.exp1Date),
    exp2Date: toIso_(obj && obj.exp2Date),
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
    sanitizeString_(existingObj.exp1Label),
    sanitizeString_(existingObj.exp1LabelAr),
    toIso_(existingObj.exp1Date),
    sanitizeString_(existingObj.exp1StatusInfo && existingObj.exp1StatusInfo.label ? existingObj.exp1StatusInfo.label : existingObj.exp1Status),
    sanitizeString_(existingObj.exp2Label),
    sanitizeString_(existingObj.exp2LabelAr),
    toIso_(existingObj.exp2Date),
    sanitizeString_(existingObj.exp2StatusInfo && existingObj.exp2StatusInfo.label ? existingObj.exp2StatusInfo.label : existingObj.exp2Status),
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
  const all = getAllRows_();

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
      r.exp1Label, r.exp1LabelAr,
      r.exp2Label, r.exp2LabelAr,
      r.fileName
    ]
      .map(x => String(x||'').toLowerCase())
      .some(s => s.includes(query));
  });

  const stats = arr => {
    const makeBucket = () => ({ total: 0, exp1: 0, exp2: 0, overall: 0 });
    const counts = {
      total: arr.length,
      expired: makeBucket(),
      upcoming: makeBucket(),
      active: makeBucket()
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

    const consider = (slot, statusObj, label, fallbackType) => {
      let type = normalizeType(statusObj && statusObj.type);
      if (!type && fallbackType) type = normalizeType(fallbackType);
      if (!type && statusObj && statusObj.label) type = normalizeType(statusObj.label);
      if (!type && label) type = normalizeType(label);
      if (!type) return;

      const bucketKey = bucketKeyForType(type);
      if (!bucketKey || !counts[bucketKey]) return;

      const bucket = counts[bucketKey];
      bucket.total += 1;
      if (!Object.prototype.hasOwnProperty.call(bucket, slot)) {
        bucket[slot] = 0;
      }
      bucket[slot] += 1;
    };

    arr.forEach(r => {
      consider('exp1', r.exp1StatusInfo, r.exp1Status, null);
      consider('exp2', r.exp2StatusInfo, r.exp2Status, null);
      consider('overall', r.statusInfo, r.status, r.statusType);
    });

    return counts;
  };

  const countsAll = stats(all);
  const countsFiltered = stats(filtered);

  const key = r => {
    const ds = [r.exp1Date, r.exp2Date].filter(Boolean).sort();
    return (ds[0] || '9999-12-31') + '|' + (r.name||'');
  };
  filtered.sort((a,b)=> key(a) < key(b) ? -1 : key(a) > key(b) ? 1 : 0);

  return { rows: filtered, countsAll, countsFiltered };
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
      const bytes = Utilities.base64Decode(data.file.b64);
      const mime = data.file.type || MimeType.PDF;
      const blob = Utilities.newBlob(bytes, mime, data.file.name);
      const created = folder.createFile(blob).setName(data.file.name);
      if (SHARE_FILES_PUBLIC) created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = sanitizeUrl_(created.getUrl());
      fileName = data.file.name;
    }

    const status = computeStatus_(data.exp1Date, data.exp2Date);
    const overall = status.overall || { label: 'Active' };

    sh.appendRow([
      id,
      data.name,
      data.nameAr,
      data.description,
      data.descriptionAr,
      data.exp1Label,
      data.exp1LabelAr,
      data.exp1Date,
      status.exp1 && status.exp1.label ? status.exp1.label : '',
      data.exp2Label,
      data.exp2LabelAr,
      data.exp2Date,
      status.exp2 && status.exp2.label ? status.exp2.label : '',
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
      exp1Label: data.exp1Label,
      exp1LabelAr: data.exp1LabelAr,
      exp2Label: data.exp2Label,
      exp2LabelAr: data.exp2LabelAr,
      exp1Date: data.exp1Date,
      exp2Date: data.exp2Date
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
      const bytes = Utilities.base64Decode(data.file.b64);
      const mime = data.file.type || MimeType.PDF;
      const blob = Utilities.newBlob(bytes, mime, data.file.name);
      const created = folder.createFile(blob).setName(data.file.name);
      if (SHARE_FILES_PUBLIC) {
        created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
      fileUrl = sanitizeUrl_(created.getUrl());
      fileName = data.file.name;
    }

    const status = computeStatus_(data.exp1Date, data.exp2Date);
    const overall = status.overall || { label: 'Active', type: 'Active' };
    const exp1StatusLabel = status.exp1 && status.exp1.label ? status.exp1.label : '';
    const exp2StatusLabel = status.exp2 && status.exp2.label ? status.exp2.label : '';

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
      data.exp1Label,
      data.exp1LabelAr,
      data.exp1Date,
      exp1StatusLabel,
      data.exp2Label,
      data.exp2LabelAr,
      data.exp2Date,
      exp2StatusLabel,
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
      exp1Label: data.exp1Label,
      exp1LabelAr: data.exp1LabelAr,
      exp2Label: data.exp2Label,
      exp2LabelAr: data.exp2LabelAr,
      exp1Date: data.exp1Date,
      exp2Date: data.exp2Date,
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
  const values = sheet.getRange(2, 1, last - 1, LICENSE_HISTORY_HEADER.length).getValues();
  const filtered = values.filter(row => String(row[0]) === target);
  filtered.sort((a, b) => {
    const da = new Date(a[1]);
    const db = new Date(b[1]);
    if (isNaN(db) && isNaN(da)) return 0;
    if (isNaN(db)) return -1;
    if (isNaN(da)) return 1;
    return db.getTime() - da.getTime();
  });
  return filtered.map(row => {
    const normalized = normalizeHistoryRow_(row).values;
    const fileUrl = sanitizeUrl_(normalized[13]);
    return {
      id: target,
      timestamp: formatTimestamp_(normalized[1]),
      action: sanitizeString_(normalized[2] || ''),
      prevExp1Label: sanitizeString_(normalized[3]),
      prevExp1LabelAr: sanitizeString_(normalized[4]),
      prevExp1Date: toIso_(normalized[5]),
      prevExp1Status: sanitizeString_(normalized[6]),
      prevExp2Label: sanitizeString_(normalized[7]),
      prevExp2LabelAr: sanitizeString_(normalized[8]),
      prevExp2Date: toIso_(normalized[9]),
      prevExp2Status: sanitizeString_(normalized[10]),
      prevStatus: sanitizeString_(normalized[11]),
      prevStatusType: sanitizeString_(normalized[12]),
      prevFileUrl: fileUrl,
      prevFilePreviewUrl: makePreviewUrl_(fileUrl),
      prevFileName: sanitizeString_(normalized[14])
    };
  });
}

// Backward compatible aliases for existing client integrations.
function updateDocument(obj) {
  return updateLicense(obj);
}

function getDocumentHistory(id) {
  return getLicenseHistory(id);
}
