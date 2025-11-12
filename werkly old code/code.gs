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

    const orderNumber = row[0];
    const tarkeebDate = row[24];
    const sheelDate = row[26];
    const productName = row[20] || "Unnamed Product";
    const productImage = row[21] || "https://via.placeholder.com/80";
    const productCategory = row[22] || "Uncategorized";
    const isMaintenanceProduct = productName.includes("Maintenance") || productName.includes("صيانة");

    // Build or update the order object
    if (!orderMap.has(orderNumber)) {
      orderMap.set(orderNumber, {
        orderNumber,
        customerName: row[6] || "N/A",
        phone: row[9] || "N/A",
        location: row[4] || "Unknown",
        comments: row[8] || "No comments",
        tarkeebDate: tarkeebDate ? normalizeDate(tarkeebDate) : "",
        sheelDate: sheelDate ? normalizeDate(sheelDate) : "",
        tarkeebTime: row[25] ? normalizeTime(row[25]) : "",
        sheelTime: row[27] ? normalizeTime(row[27]) : "",
        paymentMethod: row[11] || "N/A", // Payment Method from column L
        products: [],
        grandTotal: parseFloat(row[14]) || 0,
        isCancelled: (row[20] || "").includes("CANCEL - كنسل"),
        isMaintenanceOrder: false, // Default to false for non-maintenance orders
        maintenanceDates: [], // Holds dates for maintenance orders
      });
    }

    // Add product details to the existing order
    const order = orderMap.get(orderNumber);
    order.products.push({ productName, productCategory, productImage });
    orderMap.set(orderNumber, order);

    // Check if the product is a maintenance product and mark the order
    if (isMaintenanceProduct) {
      order.isMaintenanceOrder = true;

      // Populate dates if both tarkeeb and sheel dates exist
      if (tarkeebDate && sheelDate) {
        const allDates = getDatesInRange(new Date(tarkeebDate), new Date(sheelDate));
        order.maintenanceDates.push(...allDates.map(date => normalizeDate(date)));
      }
    }

    // Add dates to availableDates set
    if (tarkeebDate) availableDates.add(normalizeDate(tarkeebDate));
    if (sheelDate) availableDates.add(normalizeDate(sheelDate));

    // Update location summary
    if (!locationSummary[row[4]]) {
      locationSummary[row[4]] = { tarkeeb: 0, sheel: 0, cancelled: 0, totalOrders: 0 };
    }
    if (tarkeebDate) locationSummary[row[4]].tarkeeb++;
    if (sheelDate) locationSummary[row[4]].sheel++;
    if (order.isCancelled) locationSummary[row[4]].cancelled++;
    locationSummary[row[4]].totalOrders++;
  }

  // Extract all maintenance orders
  maintenanceOrders.push(...Array.from(orderMap.values()).filter(order => order.isMaintenanceOrder));

  // Convert orderMap to array
  const allOrdersArray = Array.from(orderMap.values());

  const template = HtmlService.createTemplateFromFile("index");
  template.allOrders = JSON.stringify(allOrdersArray);
  template.availableDates = JSON.stringify(Array.from(availableDates));
  template.locationSummary = JSON.stringify(locationSummary);
  template.maintenanceOrders = JSON.stringify(maintenanceOrders); // Add maintenance orders to template

  return template.evaluate();
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
  const sheet = ss.getSheetByName("Order");
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  const orderMap = new Map();
  const productMap = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]) !== orderNumber) continue;
    if (!orderMap.has(orderNumber)) {
      orderMap.set(orderNumber, {
        orderNumber,
        customerName: row[6] || "N/A",
        phone: row[9] || "N/A",
        location: row[4] || "Unknown",
        notes: row[8] || "",
        created_at: row[24] || row[26] || new Date(),
      });
    }
    productMap.push({
      item_label: row[20] || "Unnamed Product",
      category: row[22] || "Uncategorized",
      quantity: row[12] || 1,
      unit: row[13] || "",
    });
  }

  if (!orderMap.has(orderNumber)) return null;
  return { snapshot: orderMap.get(orderNumber), products: productMap };
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
