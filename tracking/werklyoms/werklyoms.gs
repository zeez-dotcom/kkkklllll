function doGet(e) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = spreadsheet.getSheetByName("Order");

  if (!orderSheet) {
    return ContentService.createTextOutput("Error: 'Order' sheet not found.");
  }

  const orderData = orderSheet.getDataRange().getValues();
  const allOrders = [];
  const orderMap = new Map(); // To track orders and their multiple products
  const availableDates = new Set();
  const locationSummary = {};
  const maintenanceOrders = []; // Array to store maintenance-specific orders

  for (let i = 1; i < orderData.length; i++) {
    const row = orderData[i];
    if (!row[0]) continue; // Skip rows without Order Number

    const orderNumber = String(row[0]).trim();
    if (!orderNumber) continue;

    const tarkeebDate = row[24];
    const sheelDate = row[26];
    const normalizedTarkeebDate = tarkeebDate ? normalizeDate(tarkeebDate) : "";
    const normalizedSheelDate = sheelDate ? normalizeDate(sheelDate) : "";
    const productName = (row[20] && String(row[20]).trim()) || "Unnamed Product";
    const productImage = (row[21] && String(row[21]).trim()) || "https://via.placeholder.com/80";
    const productCategory = (row[22] && String(row[22]).trim()) || "Uncategorized";
    const isCancelledProduct = productName.indexOf("CANCEL - كنسل") !== -1;
    const isMaintenanceProduct = productName.includes("Maintenance") || productName.includes("صيانة");

    // Build or update the order object
    if (!orderMap.has(orderNumber)) {
      orderMap.set(orderNumber, {
        orderNumber,
        customerName: (row[6] && String(row[6]).trim()) || "N/A",
        phone: (row[9] && String(row[9]).trim()) || "N/A",
        location: (row[4] && String(row[4]).trim()) || "Unknown",
        comments: (row[8] && String(row[8]).trim()) || "No comments",
        tarkeebDate: normalizedTarkeebDate,
        sheelDate: normalizedSheelDate,
        tarkeebTime: row[25] ? normalizeTime(row[25]) : "",
        sheelTime: row[27] ? normalizeTime(row[27]) : "",
        paymentMethod: (row[11] && String(row[11]).trim()) || "N/A", // Payment Method from column L
        products: [],
        grandTotal: parseAmount_(row[14]),
        isCancelled: isCancelledProduct,
        isMaintenanceOrder: false, // Default to false for non-maintenance orders
        maintenanceDates: [], // Holds dates for maintenance orders
      });
    }

    // Add product details to the existing order
    const order = orderMap.get(orderNumber);
    order.products.push({ productName, productCategory, productImage });
    order.isCancelled = order.isCancelled || isCancelledProduct;
    order.grandTotal = parseAmount_(row[14]) || order.grandTotal || 0;
    if (!order.tarkeebDate && normalizedTarkeebDate) {
      order.tarkeebDate = normalizedTarkeebDate;
    }
    if (!order.sheelDate && normalizedSheelDate) {
      order.sheelDate = normalizedSheelDate;
    }
    orderMap.set(orderNumber, order);

    // Check if the product is a maintenance product and mark the order
    if (isMaintenanceProduct) {
      order.isMaintenanceOrder = true;

      // Populate dates if both tarkeeb and sheel dates exist
      if (tarkeebDate && sheelDate) {
        const allDates = getDatesInRange(new Date(tarkeebDate), new Date(sheelDate));
        const merged = new Set(order.maintenanceDates);
        allDates.map(date => normalizeDate(date)).forEach(dateStr => {
          if (dateStr) merged.add(dateStr);
        });
        order.maintenanceDates = Array.from(merged);
      }
    }

    // Add dates to availableDates set
    if (normalizedTarkeebDate) availableDates.add(normalizedTarkeebDate);
    if (normalizedSheelDate) availableDates.add(normalizedSheelDate);

    // Update location summary
    const locationKey = (row[4] && String(row[4]).trim()) || "Unknown";
    if (!locationSummary[locationKey]) {
      locationSummary[locationKey] = { tarkeeb: 0, sheel: 0, cancelled: 0, totalOrders: 0 };
    }
    if (normalizedTarkeebDate) locationSummary[locationKey].tarkeeb++;
    if (normalizedSheelDate) locationSummary[locationKey].sheel++;
    if (order.isCancelled) locationSummary[locationKey].cancelled++;
    locationSummary[locationKey].totalOrders++;
  }

  // Extract all maintenance orders
  maintenanceOrders.push(...Array.from(orderMap.values()).filter(order => order.isMaintenanceOrder));

  // Convert orderMap to array
  const allOrdersArray = Array.from(orderMap.values());

  const template = HtmlService.createTemplateFromFile("werklytemp");
  template.allOrders = JSON.stringify(allOrdersArray);
  template.availableDates = JSON.stringify(Array.from(availableDates));
  template.locationSummary = JSON.stringify(locationSummary);
  template.maintenanceOrders = JSON.stringify(maintenanceOrders); // Add maintenance orders to template
  template.allOrdersJson = stringifyForHtml_(allOrdersArray);
  template.maintenanceOrdersJson = stringifyForHtml_(maintenanceOrders);

  return template.evaluate();
}

function stringifyForHtml_(value) {
  return JSON.stringify(value)
    .replace(/</g, "\\u003C")
    .replace(/>/g, "\\u003E")
    .replace(/&/g, "\\u0026")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
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
    return Utilities.formatDate(date, timeZone, "yyyy-MM-dd");
  }
  if (typeof date === 'string') {
    const trimmed = date.trim();
    if (!trimmed) {
      return "";
    }
    const isoMatch = trimmed.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (isoMatch) {
      return [isoMatch[1], padTwo_(isoMatch[2]), padTwo_(isoMatch[3])].join("-");
    }
    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, timeZone, "yyyy-MM-dd");
    }
    return trimmed;
  }
  return "";
}

function normalizeTime(time) {
  return time instanceof Date
    ? Utilities.formatDate(time, Session.getScriptTimeZone(), "HH:mm:ss")
    : String(time).trim();
}

function getDatesInRange(startDate, endDate) {
  const dateArray = [];
  let currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    dateArray.push(new Date(currentDate));
    currentDate.setDate(currentDate.getDate() + 1); // Increment by one day
  }
  return dateArray;
}

/**
 * Warehouse tracking configuration
 */
const WAREHOUSE_SHEET_NAME = 'WarehouseLog';
const WAREHOUSE_ROOT_FOLDER_NAME = 'OMS Warehouse Photos';
const WAREHOUSE_ROOT_FOLDER_ID = '';
const WAREHOUSE_SHARE_PUBLIC = false;

/**
 * Return warehouse logs for the provided order numbers.
 * @param {string[]} orderNumbers
 * @return {Object<string, Object[]>}
 */
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
  header.forEach((key, idx) => {
    indexMap[String(key)] = idx;
  });

  const lookup = {};
  values.forEach((row) => {
    const orderNumber = String(row[indexMap.orderNumber] || '').trim();
    if (!targets.includes(orderNumber)) {
      return;
    }

    const entry = {
      timestamp: row[indexMap.timestamp],
      operator: row[indexMap.operatorEmail],
      orderNumber,
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

/**
 * Receive an image for a product direction and store in Drive & sheet.
 * @param {Object} payload
 */
function uploadWarehousePhoto(payload) {
  if (!payload || !payload.orderNumber || typeof payload.productIndex === 'undefined' || !payload.direction || !payload.base64) {
    throw new Error('Missing required upload payload fields');
  }

  const now = new Date();
  const operatorEmail = Session.getActiveUser().getEmail() || '';

  const folder = getOrCreateWarehouseOrderFolder_(payload.orderNumber, payload.direction);
  const blob = Utilities.newBlob(Utilities.base64Decode(payload.base64), payload.mimeType || 'image/jpeg', payload.fileName || `photo_${+now}.jpg`);
  const file = folder.createFile(blob);

  if (WAREHOUSE_SHARE_PUBLIC) {
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
  let sheet = ss.getSheetByName(WAREHOUSE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(WAREHOUSE_SHEET_NAME);
    sheet.appendRow(['timestamp', 'operatorEmail', 'orderNumber', 'productIndex', 'productName', 'direction', 'photoUrl', 'notes']);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
  }
  return sheet;
}

function getOrCreateWarehouseRootFolder_() {
  if (WAREHOUSE_ROOT_FOLDER_ID) {
    try {
      return DriveApp.getFolderById(WAREHOUSE_ROOT_FOLDER_ID);
    } catch (err) {
      // Fall through to name lookup
    }
  }

  const iterator = DriveApp.getFoldersByName(WAREHOUSE_ROOT_FOLDER_NAME);
  if (iterator.hasNext()) {
    return iterator.next();
  }
  return DriveApp.createFolder(WAREHOUSE_ROOT_FOLDER_NAME);
}

function getOrCreateWarehouseOrderFolder_(orderNumber, direction) {
  const root = getOrCreateWarehouseRootFolder_();
  const orderFolder = findOrCreateSubfolder_(root, String(orderNumber));
  const dirName = direction === 'in' ? 'inbound' : 'outbound';
  return findOrCreateSubfolder_(orderFolder, dirName);
}

function findOrCreateSubfolder_(parentFolder, name) {
  const iterator = parentFolder.getFoldersByName(name);
  if (iterator.hasNext()) {
    return iterator.next();
  }
  return parentFolder.createFolder(name);
}
