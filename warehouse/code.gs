// ===== Tracking sheet metadata =================================================
const TRACKING_SHEETS = {
  ORDERS: "TrackingOrders",
  ORDER_ITEMS: "TrackingOrderItems",
  ORDER_EVENTS: "TrackingOrderEvents",
  ORDER_MEDIA: "TrackingOrderMedia",
};

const TRACKING_HEADERS = {
  ORDERS: [
    "order_uuid",
    "order_number",
    "created_at",
    "customer_name",
    "phone",
    "location",
    "status",
    "last_event_type",
    "last_event_at",
    "last_event_uuid",
    "notes",
  ],
  ORDER_ITEMS: [
    "item_uuid",
    "order_uuid",
    "order_number",
    "item_label",
    "item_category",
    "quantity",
    "unit",
    "last_event_uuid",
    "last_event_type",
    "last_event_at",
  ],
  ORDER_EVENTS: [
    "event_uuid",
    "order_uuid",
    "order_number",
    "timeline_group",
    "item_uuid",
    "item_label",
    "event_type",
    "status_after",
    "driver_name",
    "notes",
    "location_note",
    "media_urls",
    "media_drive_ids",
    "recorded_at",
    "recorded_by",
  ],
  ORDER_MEDIA: [
    "media_uuid",
    "event_uuid",
    "order_uuid",
    "order_number",
    "item_uuid",
    "file_name",
    "mime_type",
    "drive_id",
    "file_url",
    "uploaded_at",
  ],
};

const TRACKING_HEADER_INDEX = {
  ORDERS: buildHeaderIndex_(TRACKING_HEADERS.ORDERS),
  ORDER_ITEMS: buildHeaderIndex_(TRACKING_HEADERS.ORDER_ITEMS),
  ORDER_EVENTS: buildHeaderIndex_(TRACKING_HEADERS.ORDER_EVENTS),
  ORDER_MEDIA: buildHeaderIndex_(TRACKING_HEADERS.ORDER_MEDIA),
};

const ORDER_STATUS_AFTER_EVENT = {
  READY: "READY",
  OUT_FOR_DELIVERY: "OUT_FOR_DELIVERY",
  DELIVERED: "DELIVERED",
  RETURNED_TO_WAREHOUSE: "RETURNED_TO_WAREHOUSE",
  CANCELLED: "CANCELLED",
};

const TRACKING_FOLDER_ID = "";
const TRACKING_FOLDER_NAME = "Order Tracking Media";
const MAX_TRACKING_MEDIA_BYTES = 5 * 1024 * 1024; // 5MB per file
const MAX_TRACKING_MEDIA_FILES = 5;

const ORDER_COLUMN_INDEX = {
  ORDER_NUMBER: 0,
  LOCATION: 4,
  COMMENTS: 8,
  CUSTOMER_NAME: 6,
  PHONE: 9,
  PAYMENT_METHOD: 11,
  QUANTITY: 12,
  UNIT: 13,
  GRAND_TOTAL: 14,
  TARKEEB_DATE: 24,
  TARKEEB_TIME: 25,
  SHEEL_DATE: 26,
  SHEEL_TIME: 27,
  PRODUCT_NAME: 20,
  PRODUCT_IMAGE: 21,
  PRODUCT_CATEGORY: 22,
};

const PRODUCT_COLUMN_DEFAULTS = {
  NAME: 20,
  IMAGE: 21,
  CATEGORY: 22,
};

const PRODUCT_HEADER_KEYWORDS = {
  NAME: ["productname", "itemname", "product", "item", "الصنف", "المنتج"],
  IMAGE: ["productimage", "image", "photo", "picture", "img", "صورة"],
  CATEGORY: ["productcategory", "category", "type", "group", "القسم", "النوع"],
};

const ORDER_SHEET_NAME = "All Orders";

// ===== Entry points ============================================================
function doGet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = spreadsheet.getSheetByName(ORDER_SHEET_NAME);

  if (!orderSheet) {
    return ContentService.createTextOutput("Error: '" + ORDER_SHEET_NAME + "' sheet not found.");
  }

  const orderData = orderSheet.getDataRange().getValues();
  const snapshot = buildOrdersSnapshot_(orderData);

  const template = HtmlService.createTemplateFromFile("index");
  template.allOrders = JSON.stringify(snapshot.orders);
  template.availableDates = JSON.stringify(snapshot.availableDates);
  template.locationSummary = JSON.stringify(snapshot.locationSummary);
  template.maintenanceOrders = JSON.stringify(snapshot.maintenanceOrders);

  return template.evaluate();
}

// ===== Data shaping helpers ===================================================
function buildOrdersSnapshot_(rows) {
  if (!Array.isArray(rows) || rows.length <= 1) {
    return {
      orders: [],
      maintenanceOrders: [],
      availableDates: [],
      locationSummary: {},
    };
  }

  const headerRow = rows[0] || [];
  const columns = buildOrderColumnLookup_(headerRow);
  const orderMap = new Map();
  const availableDates = new Set();
  const locationSummary = {};
  const placeholderImage = "https://via.placeholder.com/80";

  const toDate = value => {
    if (value instanceof Date && !isNaN(value.getTime())) return value;
    if (value) {
      const parsed = new Date(value);
      return isNaN(parsed.getTime()) ? null : parsed;
    }
    return null;
  };

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    const orderNumber = String(getCellValue_(row, columns.ORDER_NUMBER, "")).trim();
    if (!orderNumber) continue;

    const order = upsertOrderFromRow_(orderMap, row, columns);
    if (!order) continue;
    const productName = getCellValue_(row, columns.PRODUCT_NAME, "Unnamed Product");
    const productImage = getCellValue_(row, columns.PRODUCT_IMAGE, placeholderImage) || placeholderImage;
    const productCategory = getCellValue_(row, columns.PRODUCT_CATEGORY, "Uncategorized");

    order.products.push({ productName, productCategory, productImage });
    order.isCancelled = order.isCancelled || isCancelledMarker_(productName);

    const tarkeebDate = getCellValue_(row, columns.TARKEEB_DATE, "");
    const sheelDate = getCellValue_(row, columns.SHEEL_DATE, "");

    if (isMaintenanceProduct_(productName)) {
      order.isMaintenanceOrder = true;
      const startDate = toDate(tarkeebDate);
      const endDate = toDate(sheelDate);
      if (startDate && endDate) {
        const allDates = getDatesInRange(startDate, endDate);
        order.maintenanceDates.push(...allDates.map(date => normalizeDate(date)));
      }
    }

    if (tarkeebDate) availableDates.add(normalizeDate(tarkeebDate));
    if (sheelDate) availableDates.add(normalizeDate(sheelDate));

    updateLocationSummary_(locationSummary, order.location, {
      hasTarkeeb: Boolean(tarkeebDate),
      hasSheel: Boolean(sheelDate),
      isCancelled: order.isCancelled,
    });
  }

  const orders = Array.from(orderMap.values());
  return {
    orders,
    maintenanceOrders: orders.filter(order => order.isMaintenanceOrder),
    availableDates: Array.from(availableDates),
    locationSummary,
  };
}

function upsertOrderFromRow_(orderMap, row, columns) {
  const orderNumber = String(getCellValue_(row, columns.ORDER_NUMBER, "")).trim();
  if (!orderNumber) return null;

  if (!orderMap.has(orderNumber)) {
    orderMap.set(orderNumber, {
      orderNumber,
      customerName: getCellValue_(row, columns.CUSTOMER_NAME, "N/A"),
      phone: getCellValue_(row, columns.PHONE, "N/A"),
      location: getCellValue_(row, columns.LOCATION, "Unknown"),
      comments: getCellValue_(row, columns.COMMENTS, "No comments"),
      tarkeebDate: normalizeDate(getCellValue_(row, columns.TARKEEB_DATE, "")),
      sheelDate: normalizeDate(getCellValue_(row, columns.SHEEL_DATE, "")),
      tarkeebTime: normalizeTime(getCellValue_(row, columns.TARKEEB_TIME, "")),
      sheelTime: normalizeTime(getCellValue_(row, columns.SHEEL_TIME, "")),
      paymentMethod: getCellValue_(row, columns.PAYMENT_METHOD, "N/A"),
      products: [],
      grandTotal: parseFloat(getCellValue_(row, columns.GRAND_TOTAL, 0)) || 0,
      isCancelled: false,
      isMaintenanceOrder: false,
      maintenanceDates: [],
    });
  }
  return orderMap.get(orderNumber);
}

function isMaintenanceProduct_(productName) {
  if (!productName) return false;
  return productName.includes("Maintenance") || productName.includes("صيانة");
}

function updateLocationSummary_(summary, location, flags) {
  const key = location || "Unknown";
  if (!summary[key]) {
    summary[key] = { tarkeeb: 0, sheel: 0, cancelled: 0, totalOrders: 0 };
  }
  if (flags.hasTarkeeb) summary[key].tarkeeb++;
  if (flags.hasSheel) summary[key].sheel++;
  if (flags.isCancelled) summary[key].cancelled++;
  summary[key].totalOrders++;
}

function normalizeDate(date) {
  return date instanceof Date
    ? Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd")
    : String(date).trim();
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

function getTrackingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ensureTrackingSheets_(ss);

  return {
    orders: sheetToObjects_(sheets.orders, TRACKING_HEADERS.ORDERS),
    items: sheetToObjects_(sheets.items, TRACKING_HEADERS.ORDER_ITEMS),
    events: sheetToObjects_(sheets.events, TRACKING_HEADERS.ORDER_EVENTS),
    media: sheetToObjects_(sheets.media, TRACKING_HEADERS.ORDER_MEDIA),
    eventTypes: Object.keys(ORDER_STATUS_AFTER_EVENT).filter(type => type !== "READY"),
    statusMap: ORDER_STATUS_AFTER_EVENT,
  };
}

function recordOrderEvent(payload) {
  const normalized = normalizeEventPayload_(payload);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const { snapshot, products } = findOrderSnapshot_(normalized.orderNumber) || {};
    if (!snapshot && !normalized.fallbackMetadata) {
      throw new Error("Order could not be resolved in the source sheet. Provide fallbackMetadata to continue.");
    }

    const sheets = ensureTrackingSheets_(ss);
    const resolvedOrder = ensureOrderRecord_(sheets.orders, snapshot || normalized.fallbackMetadata || normalized);
    const itemRecords = ensureItemRecords_(sheets.items, resolvedOrder, normalized, products);
    const mediaMeta = normalized.files.length
      ? persistMediaFiles_(normalized.files)
      : [];

    const eventRows = buildEventRows_(
      sheets.events,
      resolvedOrder,
      itemRecords,
      normalized,
      mediaMeta
    );
    appendRows_(sheets.events, eventRows.rowsToWrite);
    if (mediaMeta.length && eventRows.rowsToWrite.length) {
      linkMediaToEvents_(sheets.media, mediaMeta, resolvedOrder, eventRows.rowsToWrite[0][0]);
    }
    updateItemsAfterEvent_(sheets.items, eventRows.itemEventPairs, normalized);
    updateOrderStatusAfterEvent_(sheets.orders, resolvedOrder, eventRows.statusToPersist);

    return {
      success: true,
      orderUuid: resolvedOrder.order_uuid,
      recordedEvents: eventRows.rowsToWrite.length,
      eventGroupId: eventRows.statusToPersist.timelineGroup,
    };
  } finally {
    lock.releaseLock();
  }
}

function ensureTrackingSheets_(spreadsheet) {
  return {
    orders: ensureSheet_(spreadsheet, TRACKING_SHEETS.ORDERS, TRACKING_HEADERS.ORDERS),
    items: ensureSheet_(spreadsheet, TRACKING_SHEETS.ORDER_ITEMS, TRACKING_HEADERS.ORDER_ITEMS),
    events: ensureSheet_(spreadsheet, TRACKING_SHEETS.ORDER_EVENTS, TRACKING_HEADERS.ORDER_EVENTS),
    media: ensureSheet_(spreadsheet, TRACKING_SHEETS.ORDER_MEDIA, TRACKING_HEADERS.ORDER_MEDIA),
  };
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const existing = headerRange.getValues()[0];
  const needsHeader = headers.some((header, idx) => existing[idx] !== header);
  if (needsHeader) {
    headerRange.setValues([headers]);
  }
  return sheet;
}

function sheetToObjects_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map(row => {
    const obj = {};
    headers.forEach((header, idx) => {
      obj[header] = row[idx];
    });
    return obj;
  });
}

function normalizeEventPayload_(payload) {
  if (!payload) throw new Error("Payload required.");
  const orderNumber = String(payload.orderNumber || "").trim();
  if (!orderNumber) throw new Error("orderNumber is required.");

  const rawEventType = String(payload.eventType || "").toUpperCase();
  if (!ORDER_STATUS_AFTER_EVENT[rawEventType]) {
    throw new Error("Unsupported event type: " + rawEventType);
  }

  const files = Array.isArray(payload.files) ? payload.files.slice(0, MAX_TRACKING_MEDIA_FILES) : [];
  files.forEach(file => {
    if (!file || !file.dataUrl) {
      throw new Error("Every file entry must include a dataUrl.");
    }
    const sizeEstimate = Math.ceil((file.dataUrl.length / 4) * 3);
    if (sizeEstimate > MAX_TRACKING_MEDIA_BYTES) {
      throw new Error("File " + (file.name || "") + " exceeds size limit.");
    }
  });

  return {
    orderNumber,
    eventType: rawEventType,
    driverName: String(payload.driverName || "").trim(),
    notes: String(payload.notes || "").trim(),
    locationNote: String(payload.locationNote || "").trim(),
    recordedAt: payload.recordedAt ? new Date(payload.recordedAt) : new Date(),
    recordedBy: (Session.getActiveUser() && Session.getActiveUser().getEmail()) || "anonymous",
    itemLabels: Array.isArray(payload.itemLabels) ? payload.itemLabels : [],
    files,
    fallbackMetadata: payload.orderMetadata || null,
  };
}

function findOrderSnapshot_(orderNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ORDER_SHEET_NAME);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  const targetOrderNumber = String(orderNumber || "").trim();
  if (!targetOrderNumber) return null;

  const columns = buildOrderColumnLookup_(data[0] || []);
  const orderMap = new Map();
  const productMap = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentOrderNumber = String(getCellValue_(row, columns.ORDER_NUMBER, "")).trim();
    if (currentOrderNumber !== targetOrderNumber) continue;
    if (!orderMap.has(targetOrderNumber)) {
      orderMap.set(targetOrderNumber, {
        orderNumber: targetOrderNumber,
        customerName: getCellValue_(row, columns.CUSTOMER_NAME, "N/A"),
        phone: getCellValue_(row, columns.PHONE, "N/A"),
        location: getCellValue_(row, columns.LOCATION, "Unknown"),
        notes: getCellValue_(row, columns.COMMENTS, ""),
        created_at:
          getCellValue_(row, columns.TARKEEB_DATE) ||
          getCellValue_(row, columns.SHEEL_DATE) ||
          new Date(),
      });
    }
    productMap.push({
      item_label: getCellValue_(row, columns.PRODUCT_NAME, "Unnamed Product"),
      category: getCellValue_(row, columns.PRODUCT_CATEGORY, "Uncategorized"),
      quantity: getCellValue_(row, columns.QUANTITY, 1),
      unit: getCellValue_(row, columns.UNIT, ""),
    });
  }

  if (!orderMap.has(targetOrderNumber)) return null;
  return { snapshot: orderMap.get(targetOrderNumber), products: productMap };
}

function ensureOrderRecord_(sheet, snapshot) {
  const headers = TRACKING_HEADERS.ORDERS;
  const idx = TRACKING_HEADER_INDEX.ORDERS;
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const range = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < range.length; i++) {
      if (String(range[i][idx.order_number]) === snapshot.orderNumber) {
        return rowToObject_(range[i], headers, i + 2);
      }
    }
  }

  const orderUuid = Utilities.getUuid();
  const row = [
    orderUuid,
    snapshot.orderNumber,
    snapshot.created_at || new Date(),
    snapshot.customerName || "",
    snapshot.phone || "",
    snapshot.location || "",
    ORDER_STATUS_AFTER_EVENT.READY,
    "",
    "",
    "",
    snapshot.notes || "",
  ];
  sheet.appendRow(row);
  return rowToObject_(row, headers, sheet.getLastRow());
}

function ensureItemRecords_(sheet, orderRecord, payload, products) {
  const headers = TRACKING_HEADERS.ORDER_ITEMS;
  const idx = TRACKING_HEADER_INDEX.ORDER_ITEMS;
  const desiredLabels = payload.itemLabels.length
    ? payload.itemLabels
    : (products || []).map(product => product.item_label);

  if (!desiredLabels.length) {
    return [
      {
        item_uuid: "",
        order_uuid: orderRecord.order_uuid,
        order_number: orderRecord.order_number,
        item_label: "",
        rowNumber: null,
      },
    ];
  }

  const existing = sheet.getLastRow() >= 2
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues().map((row, idxRow) => ({
        rowNumber: idxRow + 2,
        data: row,
      }))
    : [];

  const matches = [];
  desiredLabels.forEach(label => {
    const match = existing.find(entry =>
      entry.data[idx.order_number] === orderRecord.order_number &&
      entry.data[idx.item_label] === label
    );
    if (match) {
      matches.push(rowToObject_(match.data, headers, match.rowNumber));
    } else {
      const snapshotItem = (products || []).find(product => product.item_label === label) || {};
      const newRow = [
        Utilities.getUuid(),
        orderRecord.order_uuid,
        orderRecord.order_number,
        label,
        snapshotItem.category || "",
        snapshotItem.quantity || "",
        snapshotItem.unit || "",
        "",
        "",
        "",
      ];
      sheet.appendRow(newRow);
      matches.push(rowToObject_(newRow, headers, sheet.getLastRow()));
    }
  });

  return matches;
}

function persistMediaFiles_(files) {
  const folder = getTrackingFolder_();
  const metadata = [];
  files.forEach(file => {
    const parsed = parseDataUrl_(file.dataUrl);
    const blob = Utilities.newBlob(parsed.bytes, parsed.mimeType, file.name || ("photo-" + new Date().getTime()));
    const driveFile = folder.createFile(blob);
    try {
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err) {
      // Ignored if sharing is restricted by domain policy
    }
    metadata.push({
      media_uuid: Utilities.getUuid(),
      drive_id: driveFile.getId(),
      file_url: driveFile.getUrl(),
      file_name: driveFile.getName(),
      mime_type: parsed.mimeType,
    });
  });
  return metadata;
}

function linkMediaToEvents_(mediaSheet, mediaMeta, orderRecord, targetEventUuid) {
  if (!mediaMeta.length) return;
  const rows = mediaMeta.map(meta => [
    meta.media_uuid,
    targetEventUuid,
    orderRecord.order_uuid,
    orderRecord.order_number,
    "",
    meta.file_name,
    meta.mime_type,
    meta.drive_id,
    meta.file_url,
    new Date(),
  ]);
  appendRows_(mediaSheet, rows);
}

function buildEventRows_(sheet, orderRecord, itemRecords, payload, mediaMeta) {
  const headers = TRACKING_HEADERS.ORDER_EVENTS;
  const idx = TRACKING_HEADER_INDEX.ORDER_EVENTS;
  const timelineGroup = Utilities.getUuid();
  const statusAfter = ORDER_STATUS_AFTER_EVENT[payload.eventType];
  const recordedRows = [];
  const itemEventPairs = [];

  const targets = itemRecords.length ? itemRecords : [{
    item_uuid: "",
    item_label: "",
  }];

  targets.forEach(item => {
    const eventUuid = Utilities.getUuid();
    const mediaUrls = JSON.stringify(mediaMeta.map(meta => meta.file_url));
    const mediaIds = JSON.stringify(mediaMeta.map(meta => meta.drive_id));
    const row = [
      eventUuid,
      orderRecord.order_uuid,
      orderRecord.order_number,
      timelineGroup,
      item.item_uuid || "",
      item.item_label || (payload.itemLabels.length ? payload.itemLabels.join(", ") : ""),
      payload.eventType,
      statusAfter,
      payload.driverName,
      payload.notes,
      payload.locationNote,
      mediaUrls,
      mediaIds,
      payload.recordedAt,
      payload.recordedBy,
    ];
    recordedRows.push(row);
    if (item.rowNumber) {
      itemEventPairs.push({
        rowNumber: item.rowNumber,
        eventUuid,
      });
    }
  });

  return {
    rowsToWrite: recordedRows,
    itemEventPairs,
    statusToPersist: {
      orderUuid: orderRecord.order_uuid,
      eventType: payload.eventType,
      statusAfter,
      lastEventAt: payload.recordedAt,
      timelineGroup,
      lastEventUuid: recordedRows[recordedRows.length - 1][idx.event_uuid],
    },
  };
}

function appendRows_(sheet, rows) {
  if (!rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function updateItemsAfterEvent_(sheet, itemEventPairs, payload) {
  if (!itemEventPairs.length) return;
  const headers = TRACKING_HEADERS.ORDER_ITEMS;
  const idx = TRACKING_HEADER_INDEX.ORDER_ITEMS;
  itemEventPairs.forEach(pair => {
    if (!pair.rowNumber) return;
    const range = sheet.getRange(pair.rowNumber, 1, 1, headers.length);
    const row = range.getValues()[0];
    row[idx.last_event_uuid] = pair.eventUuid;
    row[idx.last_event_type] = payload.eventType;
    row[idx.last_event_at] = payload.recordedAt;
    range.setValues([row]);
  });
}

function updateOrderStatusAfterEvent_(sheet, orderRecord, statusPayload) {
  if (!statusPayload) return;
  const headers = TRACKING_HEADERS.ORDERS;
  const idx = TRACKING_HEADER_INDEX.ORDERS;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, headers.length);
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][idx.order_uuid] === orderRecord.order_uuid) {
      values[i][idx.status] = statusPayload.statusAfter;
      values[i][idx.last_event_type] = statusPayload.eventType;
      values[i][idx.last_event_at] = statusPayload.lastEventAt;
      values[i][idx.last_event_uuid] = statusPayload.lastEventUuid;
      range.setValues(values);
      return;
    }
  }
}

function parseDataUrl_(dataUrl) {
  const matches = String(dataUrl).match(/^data:(.*?);base64,(.*)$/);
  if (!matches) {
    throw new Error("Invalid dataUrl format.");
  }
  return {
    mimeType: matches[1],
    bytes: Utilities.base64Decode(matches[2]),
  };
}

function getTrackingFolder_() {
  if (TRACKING_FOLDER_ID) {
    return DriveApp.getFolderById(TRACKING_FOLDER_ID);
  }
  const folders = DriveApp.getFoldersByName(TRACKING_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(TRACKING_FOLDER_NAME);
}

function buildOrderColumnLookup_(headerRow) {
  const lookup = Object.assign({}, ORDER_COLUMN_INDEX);
  const productColumns = resolveProductColumns_(headerRow);
  lookup.PRODUCT_NAME = productColumns.name;
  lookup.PRODUCT_IMAGE = productColumns.image;
  lookup.PRODUCT_CATEGORY = productColumns.category;
  return lookup;
}

function resolveProductColumns_(headerRow) {
  const normalized = (headerRow || []).map(normalizeHeaderValue_);
  return {
    name: findHeaderIndex_(normalized, PRODUCT_HEADER_KEYWORDS.NAME, PRODUCT_COLUMN_DEFAULTS.NAME),
    image: findHeaderIndex_(normalized, PRODUCT_HEADER_KEYWORDS.IMAGE, PRODUCT_COLUMN_DEFAULTS.IMAGE),
    category: findHeaderIndex_(normalized, PRODUCT_HEADER_KEYWORDS.CATEGORY, PRODUCT_COLUMN_DEFAULTS.CATEGORY),
  };
}

function normalizeHeaderValue_(value) {
  return String(value == null ? "" : value)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\u0600-\u06FF]+/g, "");
}

function findHeaderIndex_(normalizedRow, keywords, fallback) {
  if (!Array.isArray(normalizedRow)) return fallback;
  if (Array.isArray(keywords) && keywords.length) {
    for (let i = 0; i < normalizedRow.length; i++) {
      const header = normalizedRow[i];
      if (!header) continue;
      for (let j = 0; j < keywords.length; j++) {
        if (header.indexOf(keywords[j]) !== -1) {
          return i;
        }
      }
    }
  }
  return fallback;
}

function getCellValue_(row, index, fallback) {
  if (!Array.isArray(row)) return fallback;
  if (typeof index !== "number" || index < 0 || index >= row.length) return fallback;
  const value = row[index];
  return value === undefined || value === null || value === "" ? fallback : value;
}

function isCancelledMarker_(value) {
  if (!value) return false;
  const normalized = String(value).toLowerCase();
  return normalized.includes("cancel") || normalized.includes("كنسل");
}

function rowToObject_(row, headers, rowNumber) {
  const obj = { rowNumber };
  headers.forEach((header, idx) => {
    obj[header] = row[idx];
  });
  return obj;
}

function buildHeaderIndex_(headers) {
  return headers.reduce((acc, header, idx) => {
    acc[header] = idx;
    return acc;
  }, {});
}
