var WERKLY_WAREHOUSE_SHEET_NAME = typeof WERKLY_WAREHOUSE_SHEET_NAME !== 'undefined' ? WERKLY_WAREHOUSE_SHEET_NAME : 'WarehouseLog';
var WERKLY_WAREHOUSE_ROOT_NAME = typeof WERKLY_WAREHOUSE_ROOT_NAME !== 'undefined' ? WERKLY_WAREHOUSE_ROOT_NAME : 'OMS Warehouse Photos';
var WERKLY_WAREHOUSE_ROOT_ID = typeof WERKLY_WAREHOUSE_ROOT_ID !== 'undefined' ? WERKLY_WAREHOUSE_ROOT_ID : '';
var WERKLY_WAREHOUSE_SHARE_PUBLIC = typeof WERKLY_WAREHOUSE_SHARE_PUBLIC !== 'undefined' ? WERKLY_WAREHOUSE_SHARE_PUBLIC : false;

function getWarehouseLogs(orderNumbers) {
  if (!Array.isArray(orderNumbers) || !orderNumbers.length) {
    return {};
  }

  const targets = orderNumbers
    .map(function(num) { return String(num).trim(); })
    .filter(function(val) { return val; });
  if (!targets.length) {
    return {};
  }

  const sheet = ensureWarehouseSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return {};
  }

  const header = values.shift();
  const indexMap = {};
  header.forEach(function(key, idx) {
    indexMap[String(key)] = idx;
  });

  const lookup = {};
  values.forEach(function(row) {
    const orderNumber = String(row[indexMap.orderNumber] || '').trim();
    if (!targets.includes(orderNumber)) {
      return;
    }

    const entry = {
      timestamp: row[indexMap.timestamp],
      operator: row[indexMap.operatorEmail],
      orderNumber: orderNumber,
      productIndex: row[indexMap.productIndex],
      productName: row[indexMap.productName],
      direction: row[indexMap.direction],
      photoUrl: row[indexMap.photoUrl],
      notes: row[indexMap.notes] || ''
    };

    if (!lookup[orderNumber]) {
      lookup[orderNumber] = [];
    }
    lookup[orderNumber].push(entry);
  });

  return lookup;
}

function uploadWarehousePhoto(payload) {
  if (!payload || !payload.orderNumber || typeof payload.productIndex === 'undefined' || !payload.direction || !payload.base64) {
    throw new Error('Missing required upload payload fields');
  }

  const now = new Date();
  const operatorEmail = Session.getActiveUser().getEmail() || '';

  const folder = getOrCreateWarehouseOrderFolder_(payload.orderNumber, payload.direction);
  const blob = Utilities.newBlob(
    Utilities.base64Decode(payload.base64),
    payload.mimeType || 'image/jpeg',
    payload.fileName || 'photo_' + Number(now) + '.jpg'
  );
  const file = folder.createFile(blob);

  if (WERKLY_WAREHOUSE_SHARE_PUBLIC) {
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err) {
      // Ignore sharing errors (e.g., insufficient permissions)
    }
  }

  const sheet = ensureWarehouseSheet_();
  const row = [
    now,
    operatorEmail,
    String(payload.orderNumber),
    Number(payload.productIndex),
    String(payload.productName || ''),
    String(payload.direction),
    file.getUrl(),
    String(payload.notes || '')
  ];
  sheet.appendRow(row);

  return {
    timestamp: now,
    operator: operatorEmail,
    orderNumber: String(payload.orderNumber),
    productIndex: Number(payload.productIndex),
    productName: String(payload.productName || ''),
    direction: String(payload.direction),
    photoUrl: file.getUrl(),
    notes: String(payload.notes || '')
  };
}

function ensureWarehouseSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(WERKLY_WAREHOUSE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(WERKLY_WAREHOUSE_SHEET_NAME);
    sheet.appendRow([
      'timestamp',
      'operatorEmail',
      'orderNumber',
      'productIndex',
      'productName',
      'direction',
      'photoUrl',
      'notes'
    ]);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
  }
  return sheet;
}

function getOrCreateWarehouseRootFolder_() {
  if (WERKLY_WAREHOUSE_ROOT_ID) {
    try {
      return DriveApp.getFolderById(WERKLY_WAREHOUSE_ROOT_ID);
    } catch (err) {
      // Fall through to name lookup
    }
  }

  const iterator = DriveApp.getFoldersByName(WERKLY_WAREHOUSE_ROOT_NAME);
  if (iterator.hasNext()) {
    return iterator.next();
  }
  return DriveApp.createFolder(WERKLY_WAREHOUSE_ROOT_NAME);
}

function getOrCreateWarehouseOrderFolder_(orderNumber, direction) {
  const root = getOrCreateWarehouseRootFolder_();
  const orderFolder = findOrCreateSubfolder_(root, String(orderNumber));
  const dirName = direction === 'in' ? 'inbound' : 'outbound';
  return findOrCreateSubfolder_(orderFolder, dirName);
}
