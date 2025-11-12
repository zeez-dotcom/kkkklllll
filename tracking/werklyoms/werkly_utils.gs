function stringifyForHtml_(value) {
  return JSON.stringify(value)
    .replace(/</g, '\\u003C')
    .replace(/>/g, '\\u003E')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

function parseAmount_(value) {
  if (typeof value === 'number') {
    return isNaN(value) ? 0 : value;
  }
  if (typeof value === 'string') {
    const normalized = value.replace(/[^\d.-]/g, '');
    const parsed = Number(normalized);
    return isNaN(parsed) ? 0 : parsed;
  }
  return 0;
}

function padTwo_(value) {
  return String(value).padStart(2, '0');
}

function normalizeDate(date) {
  const timeZone = Session.getScriptTimeZone();
  if (date instanceof Date && !isNaN(date.getTime())) {
    return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
  }
  if (typeof date === 'string') {
    const trimmed = date.trim();
    if (!trimmed) {
      return '';
    }
    const isoMatch = trimmed.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (isoMatch) {
      return [isoMatch[1], padTwo_(isoMatch[2]), padTwo_(isoMatch[3])].join('-');
    }
    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, timeZone, 'yyyy-MM-dd');
    }
    return trimmed;
  }
  return '';
}

function normalizeTime(time) {
  return time instanceof Date
    ? Utilities.formatDate(time, Session.getScriptTimeZone(), 'HH:mm:ss')
    : String(time).trim();
}

function getDatesInRange(startDate, endDate) {
  const dateArray = [];
  let currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    dateArray.push(new Date(currentDate));
    currentDate.setDate(currentDate.getDate() + 1);
  }
  return dateArray;
}

function findOrCreateSubfolder_(parentFolder, name) {
  const iterator = parentFolder.getFoldersByName(name);
  if (iterator.hasNext()) {
    return iterator.next();
  }
  return parentFolder.createFolder(name);
}
