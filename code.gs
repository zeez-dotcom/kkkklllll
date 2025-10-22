/**********************
 * CONFIG
 **********************/
const SHEET_NAME = 'Licenses';
const HEADER = [
  'id','name','nameAr','description','descriptionAr',
  'exp1Label','exp1LabelAr','exp1Date',
  'exp2Label','exp2LabelAr','exp2Date',
  'status','fileUrl','fileName','createdAt'
];
let FOLDER_ID = '';                 // optional: preset Drive folder ID
const SHARE_FILES_PUBLIC = true;    // auto "Anyone with link → Viewer"

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
  const id = parseDriveFileId_(fileUrl);
  return id ? `https://drive.google.com/file/d/${id}/preview` : (fileUrl || '');
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
  obj.filePreviewUrl = makePreviewUrl_(obj.fileUrl);
  return obj;
}
function nextId_() {
  const rows = getAllRows_();
  const maxId = rows.reduce((m,r)=> {
    const n = Number(r.id);
    return isFinite(n) && n > m ? n : m;
  }, 0);
  return String(maxId + 1);
}

/**********************
 * READ API
 **********************/
function getDashboardData(q) {
  const all = getAllRows_();

  const query = (q || '').trim().toLowerCase();
  const filtered = !query ? all : all.filter(r => {
    return [r.id, r.name, r.nameAr, r.description, r.descriptionAr, r.exp1Label, r.exp1LabelAr, r.exp2Label, r.exp2LabelAr, r.fileName]
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
    const sortName = r.name || r.nameAr || '';
    return (ds[0] || '9999-12-31') + '|' + sortName;
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

    let fileUrl = '';
    let fileName = '';

    if (obj && obj.fileB64 && obj.fileName) {
      const folder = getFolder_();
      const bytes = Utilities.base64Decode(obj.fileB64);
      const mime = obj.fileType || MimeType.PDF;
      const blob = Utilities.newBlob(bytes, mime, obj.fileName);
      const created = folder.createFile(blob).setName(obj.fileName);
      if (SHARE_FILES_PUBLIC) created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = created.getUrl();
      fileName = obj.fileName;
    }

    const name        = String(obj && obj.name || '');
    const nameAr      = String(obj && obj.nameAr || '');
    const description = String(obj && obj.description || '');
    const descriptionAr = String(obj && obj.descriptionAr || '');
    const exp1Label   = String(obj && obj.exp1Label || '');
    const exp1LabelAr = String(obj && obj.exp1LabelAr || '');
    const exp1Date    = toIso_(obj && obj.exp1Date);
    const exp2Label   = String(obj && obj.exp2Label || '');
    const exp2LabelAr = String(obj && obj.exp2LabelAr || '');
    const exp2Date    = toIso_(obj && obj.exp2Date);
    const status      = computeStatus_(exp1Date, exp2Date);

    sh.appendRow([
      id,
      name,
      nameAr,
      description,
      descriptionAr,
      exp1Label,
      exp1LabelAr,
      exp1Date,
      exp2Label,
      exp2LabelAr,
      exp2Date,
      status,
      fileUrl,
      fileName || name || nameAr,
      now
    ]);

    return { ok:true, id, fileUrl, filePreviewUrl: makePreviewUrl_(fileUrl) };
  } catch (err) {
    return { ok:false, error:String(err && err.message || err) };
  }
}
