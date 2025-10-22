/**********************
 * CONFIG
 **********************/
const SHEET_NAME = 'Licenses';
const HEADER = [
  'id',
  'name','nameAr',
  'description','descriptionAr',
  'exp1Label','exp1LabelAr','exp1Date',
  'exp2Label','exp2LabelAr','exp2Date',
  'status','fileUrl','fileName','createdAt'
];
const LEGACY_HEADER = [
  'id',
  'name',
  'description',
  'exp1Label',
  'exp1Date',
  'exp2Label',
  'exp2Date',
  'status',
  'fileUrl',
  'fileName',
  'createdAt'
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
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const width = Math.max(sh.getLastColumn(), HEADER.length, LEGACY_HEADER.length);
  const first = sh.getRange(1,1,1,width).getValues()[0];
  const isCurrentHeader = HEADER.every((h,i)=>first[i] === h);
  if (!isCurrentHeader) {
    sh.getRange(1,1,1,HEADER.length).setValues([HEADER]);
  }
  upgradeLegacyRowsIfNeeded_(sh);
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
function computeStatus_(exp1Iso, exp2Iso) {
  const c = [exp1Iso, exp2Iso].filter(Boolean).sort();
  if (!c.length) return 'Active';
  const dd = daysUntil_(c[0]);
  if (dd === null) return 'Active';
  if (dd < 0) return 'Expired';
  if (dd <= 30) return `Upcoming (in ${dd} day${dd===1?'':'s'})`;
  return 'Active';
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
function looksLegacyRow_(row) {
  const newStatus = sanitizeString_(row[11]);
  if (newStatus) return false;
  const legacyStatus = sanitizeString_(row[7]);
  if (!legacyStatus) return false;
  const normalized = legacyStatus.toLowerCase();
  const isLegacyStatus = normalized === 'active' || normalized === 'expired' || normalized.startsWith('upcoming');
  if (!isLegacyStatus) return false;
  return true;
}
function convertLegacyRow_(row) {
  const exp1Date = toIso_(row[4]);
  const exp2Date = toIso_(row[6]);
  const legacyStatus = sanitizeString_(row[7]);
  const status = legacyStatus || computeStatus_(exp1Date, exp2Date);
  return [
    row[0] ?? '',
    sanitizeString_(row[1]),
    '',
    sanitizeString_(row[2]),
    '',
    sanitizeString_(row[3]),
    '',
    exp1Date,
    sanitizeString_(row[5]),
    '',
    exp2Date,
    status,
    sanitizeUrl_(row[8]),
    sanitizeString_(row[9]) || sanitizeString_(row[1]),
    row[10] ?? ''
  ];
}
function upgradeLegacyRowsIfNeeded_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const width = Math.max(sh.getLastColumn(), HEADER.length);
  const values = sh.getRange(2, 1, last - 1, width).getValues();
  let needsUpdate = false;
  const upgraded = values.map(row => {
    if (looksLegacyRow_(row)) {
      needsUpdate = true;
      return convertLegacyRow_(row);
    }
    return HEADER.map((_, i) => row[i] ?? '');
  });
  if (needsUpdate) {
    sh.getRange(2, 1, upgraded.length, HEADER.length).setValues(upgraded);
  }
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
  obj.status = computeStatus_(obj.exp1Date, obj.exp2Date);
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

  const stats = arr => ({
    total: arr.length,
    expired: arr.filter(r=>r.status==='Expired').length,
    upcoming: arr.filter(r=>String(r.status).startsWith('Upcoming')).length,
    active: arr.filter(r=>r.status==='Active').length
  });

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

    sh.appendRow([
      id,
      data.name,
      data.nameAr,
      data.description,
      data.descriptionAr,
      data.exp1Label,
      data.exp1LabelAr,
      data.exp1Date,
      data.exp2Label,
      data.exp2LabelAr,
      data.exp2Date,
      status,
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
