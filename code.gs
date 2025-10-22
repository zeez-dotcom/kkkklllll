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

/**********************
 * READ API
 **********************/
function getDashboardData(q) {
  const all = getAllRows_();

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
