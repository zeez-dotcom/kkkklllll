/**
 * Laundry POS backend (Google Apps Script)
 * - Orders + line items live in LaundryOrders / LaundryOrderItems.
 * - Catalog drives all products/services/prices/images from a sheet.
 * - UI talks to getOrders/getCatalog/saveCatalog/recordOrder.
 */

const LAUNDRY_SHEETS = {
  ORDERS: "LaundryOrders",
  ITEMS: "LaundryOrderItems",
};

const SETTINGS_SHEET = "Settings";
const SETTINGS_HEADERS = ["key", "value"];
const DEFAULT_LOGO = "https://placehold.co/180x60?text=Laundry+POS";

const CUSTOMER_SHEET = "Customers";
const CUSTOMER_HEADERS = ["customer_id", "name", "nickname", "phone", "notes", "created_at"];

const SUBSCRIPTION_SHEET = "Subscriptions";
const SUBSCRIPTION_HEADERS = [
  "subscription_id",
  "customer_id",
  "customer_name",
  "package_name",
  "product_key",
  "service_code",
  "free_allowance",
  "remaining_free",
  "start_date",
  "end_date",
  "active",
  "created_at",
];

const LAUNDRY_HEADERS = {
  ORDERS: [
    "order_id",
    "order_number",
    "walk_in_date",
    "exit_date",
    "customer_name",
    "phone",
    "customer_nickname",
    "customer_id",
    "subscription_id",
    "status",
    "subtotal",
    "discount",
    "total",
    "item_count",
    "notes",
    "created_at",
    "created_by",
  ],
  ITEMS: [
    "item_id",
    "order_id",
    "order_number",
    "product_type",
    "product_label",
    "service_code",
    "service_label",
    "quantity",
    "unit_price",
    "line_total",
  ],
};

const LAUNDRY_HEADER_INDEX = {
  ORDERS: buildHeaderIndex_(LAUNDRY_HEADERS.ORDERS),
  ITEMS: buildHeaderIndex_(LAUNDRY_HEADERS.ITEMS),
};

const ORDER_STATUS = {
  PENDING: "PENDING",
  IN_PROGRESS: "IN_PROGRESS",
  READY: "READY",
  DELIVERED: "DELIVERED",
};

const CATALOG_SHEET = "Catalog";
// Catalog sheet is the single source of truth for the price book.
const CATALOG_HEADERS = [
  "product_key",
  "service_code",
  "product_label_en",
  "product_label_ar",
  "service_label_en",
  "service_label_ar",
  "price",
  "image_url",
  "active",
];

// Seed rows to avoid a blank experience on first deploy.
const DEFAULT_CATALOG = [
  {
    product_key: "dishdasha",
    service_code: "WASH",
    product_label_en: "Dishdasha",
    product_label_ar: "دشداشة",
    service_label_en: "Wash only",
    service_label_ar: "غسيل فقط",
    price: 0.35,
    image_url: "https://placehold.co/200x200?text=Dishdasha",
    active: true,
  },
  {
    product_key: "dishdasha",
    service_code: "WASH_IRON",
    product_label_en: "Dishdasha",
    product_label_ar: "دشداشة",
    service_label_en: "Wash & iron",
    service_label_ar: "غسيل وكي",
    price: 0.5,
    image_url: "https://placehold.co/200x200?text=Dishdasha",
    active: true,
  },
  {
    product_key: "dishdasha",
    service_code: "IRON",
    product_label_en: "Dishdasha",
    product_label_ar: "دشداشة",
    service_label_en: "Iron only",
    service_label_ar: "كي فقط",
    price: 0.25,
    image_url: "https://placehold.co/200x200?text=Dishdasha",
    active: true,
  },
  {
    product_key: "abaya",
    service_code: "WASH",
    product_label_en: "Abaya",
    product_label_ar: "عباية",
    service_label_en: "Wash only",
    service_label_ar: "غسيل فقط",
    price: 0.4,
    image_url: "https://placehold.co/200x200?text=Abaya",
    active: true,
  },
  {
    product_key: "abaya",
    service_code: "WASH_IRON",
    product_label_en: "Abaya",
    product_label_ar: "عباية",
    service_label_en: "Wash & steam",
    service_label_ar: "غسيل وتبخير",
    price: 0.6,
    image_url: "https://placehold.co/200x200?text=Abaya",
    active: true,
  },
  {
    product_key: "abaya",
    service_code: "SPOT_CLEAN",
    product_label_en: "Abaya",
    product_label_ar: "عباية",
    service_label_en: "Spot clean",
    service_label_ar: "تنظيف موضعي",
    price: 0.3,
    image_url: "https://placehold.co/200x200?text=Abaya",
    active: true,
  },
  {
    product_key: "bedding",
    service_code: "WASH",
    product_label_en: "Bedding set",
    product_label_ar: "مفروشات",
    service_label_en: "Wash only",
    service_label_ar: "غسيل فقط",
    price: 0.75,
    image_url: "https://placehold.co/200x200?text=Bedding",
    active: true,
  },
  {
    product_key: "bedding",
    service_code: "WASH_IRON",
    product_label_en: "Bedding set",
    product_label_ar: "مفروشات",
    service_label_en: "Wash & iron",
    service_label_ar: "غسيل وكي",
    price: 1.0,
    image_url: "https://placehold.co/200x200?text=Bedding",
    active: true,
  },
  {
    product_key: "bedding",
    service_code: "DEEP_CLEAN",
    product_label_en: "Bedding set",
    product_label_ar: "مفروشات",
    service_label_en: "Deep clean",
    service_label_ar: "تنظيف عميق",
    price: 1.25,
    image_url: "https://placehold.co/200x200?text=Bedding",
    active: true,
  },
  {
    product_key: "custom",
    service_code: "CUSTOM",
    product_label_en: "Custom item",
    product_label_ar: "عنصر مخصص",
    service_label_en: "Custom service",
    service_label_ar: "خدمة مخصصة",
    price: 0,
    image_url: "https://placehold.co/200x200?text=Custom",
    active: true,
  },
];

const CURRENCY = "KWD";
const TIME_ZONE = Session.getScriptTimeZone();

function doGet() {
  const catalog = loadCatalog_();
  const settings = loadSettings_();
  const template = HtmlService.createTemplateFromFile("index");
  template.priceBook = JSON.stringify(buildPublicPriceBook_(catalog.priceBook));
  template.currency = CURRENCY;
  template.statuses = JSON.stringify(ORDER_STATUS);
  template.logoUrl = settings.logoUrl || DEFAULT_LOGO;
  return template.evaluate().setTitle("Laundry POS");
}

function getOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ensureLaundrySheets_(ss);
  const catalog = loadCatalog_(ss);
  const settings = loadSettings_(ss);
  const customers = loadCustomers_(ss);
  const subscriptions = loadSubscriptions_(ss);

  const orders = sheetToObjectsFlexible_(sheets.orders, LAUNDRY_HEADERS.ORDERS).map(serializeOrder_);
  const items = sheetToObjectsFlexible_(sheets.items, LAUNDRY_HEADERS.ITEMS);

  const itemsByOrder = items.reduce((map, item) => {
    if (!map[item.order_id]) map[item.order_id] = [];
    map[item.order_id].push(item);
    return map;
  }, {});

  let ordersWithItems = orders.map(order => Object.assign({}, order, { items: itemsByOrder[order.order_id] || [] }));

  // Fallback: if orders sheet is empty but items exist, group items into lightweight orders so UI still renders.
  if (!ordersWithItems.length && items.length) {
    const grouped = items.reduce((map, item) => {
      const key = item.order_number || item.order_id || "UNKNOWN";
      if (!map[key]) {
        map[key] = {
          order_id: item.order_id || Utilities.getUuid(),
          order_number: item.order_number || key,
          walk_in_date: "",
          exit_date: "",
          customer_name: item.customer_name || "",
          customer_nickname: "",
          phone: "",
          status: ORDER_STATUS.PENDING,
          subtotal: 0,
          discount: 0,
          total: 0,
          item_count: 0,
          notes: "",
          created_at: new Date(),
          created_by: "",
          subscription_id: item.subscription_id || "",
        };
      }
      map[key].items = map[key].items || [];
      map[key].items.push(item);
      map[key].item_count += 1;
      map[key].subtotal += Number(item.line_total) || 0;
      map[key].total = map[key].subtotal;
      return map;
    }, {});
    ordersWithItems = Object.values(grouped);
  }

  // Show newest orders first to keep freshly created records visible after refresh.
  ordersWithItems.sort((a, b) => {
    const aTime = a.created_at instanceof Date && !isNaN(a.created_at) ? a.created_at.getTime() : 0;
    const bTime = b.created_at instanceof Date && !isNaN(b.created_at) ? b.created_at.getTime() : 0;
    if (aTime !== bTime) return bTime - aTime;
    return String(b.order_number || "").localeCompare(String(a.order_number || ""));
  });

  const statusCounts = Object.keys(ORDER_STATUS).reduce((acc, key) => {
    acc[ORDER_STATUS[key]] = 0;
    return acc;
  }, {});
  let totalRevenue = 0;
  ordersWithItems.forEach(order => {
    const status = order.status || ORDER_STATUS.PENDING;
    if (statusCounts[status] !== undefined) {
      statusCounts[status] += 1;
    }
    totalRevenue += Number(order.total) || 0;
  });

  return {
    orders: ordersWithItems,
    stats: {
      totalOrders: ordersWithItems.length,
      byStatus: statusCounts,
      totalRevenue: roundAmount_(totalRevenue),
    },
    priceBook: buildPublicPriceBook_(catalog.priceBook),
    logoUrl: settings.logoUrl || DEFAULT_LOGO,
    customers,
    subscriptions,
  };
}

function serializeOrder_(order) {
  const copy = Object.assign({}, order);
  copy.walk_in_date = order.walk_in_date ? new Date(order.walk_in_date) : "";
  copy.exit_date = order.exit_date ? new Date(order.exit_date) : "";
  copy.created_at = order.created_at ? new Date(order.created_at) : "";
  copy.total = Number(order.total) || 0;
  copy.subtotal = Number(order.subtotal) || 0;
  copy.discount = Number(order.discount) || 0;
  copy.status = copy.status || ORDER_STATUS.PENDING;
  return copy;
}

function getCatalog() {
  const catalog = loadCatalog_();
  const settings = loadSettings_();
  return {
    catalog: catalog.rows,
    priceBook: buildPublicPriceBook_(catalog.priceBook),
    logoUrl: settings.logoUrl || DEFAULT_LOGO,
  };
}

function saveCatalog(payload) {
  if (!payload || !Array.isArray(payload.entries)) {
    throw new Error("entries array is required.");
  }
  const normalized = normalizeCatalogPayload_(payload.entries);
  persistCatalogRows_(normalized);
  if (payload.logoUrl != null) {
    updateSettings_({ logoUrl: sanitizeUrl_(payload.logoUrl) });
  }
  return getCatalog();
}

function recordOrder(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catalog = loadCatalog_(ss);
  const normalized = normalizeOrderPayload_(payload, catalog.priceBook);
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);

  try {
    const sheets = ensureLaundrySheets_(ss);
    const orderId = Utilities.getUuid();
    const orderNumber = generateOrderNumber_(sheets.orders);
    const now = new Date();
    const customer = ensureCustomerForOrder_(ss, normalized);
    const subscriptionResult = normalized.subscriptionId
      ? applySubscriptionAllowance_(ss, normalized.subscriptionId, normalized.items)
      : { items: normalized.items, remainingFree: null };
    const finalItems = subscriptionResult.items;
    const subtotal = roundAmount_(finalItems.reduce((sum, item) => sum + item.lineTotal, 0));
    const total = roundAmount_(Math.max(0, subtotal - normalized.discount));

    const orderRow = [
      orderId,
      orderNumber,
      normalized.walkInDate,
      normalized.exitDate,
      customer.name,
      normalized.phone,
      customer.nickname,
      customer.customerId,
      normalized.subscriptionId || "",
      ORDER_STATUS.PENDING,
      subtotal,
      normalized.discount,
      total,
      finalItems.length,
      normalized.notes,
      now,
      normalized.createdBy,
    ];

    sheets.orders.appendRow(orderRow);

    if (finalItems.length) {
      const startRow = sheets.items.getLastRow() + 1;
      const rows = finalItems.map(item => [
        Utilities.getUuid(),
        orderId,
        orderNumber,
        item.productType,
        item.productLabel,
        item.serviceCode,
        item.serviceLabel,
        item.quantity,
        item.unitPrice,
        item.lineTotal,
      ]);
      sheets.items.getRange(startRow, 1, rows.length, LAUNDRY_HEADERS.ITEMS.length).setValues(rows);
    }

    return { success: true, orderNumber, orderId, subscriptionId: normalized.subscriptionId || "" };
  } finally {
    lock.releaseLock();
  }
}

function ensureLaundrySheets_(spreadsheet) {
  return {
    orders: ensureSheet_(spreadsheet, LAUNDRY_SHEETS.ORDERS, LAUNDRY_HEADERS.ORDERS),
    items: ensureSheet_(spreadsheet, LAUNDRY_SHEETS.ITEMS, LAUNDRY_HEADERS.ITEMS),
  };
}

function ensureCustomerForOrder_(spreadsheet, normalized) {
  const name = sanitizeText_(normalized.customerName);
  const phone = sanitizeText_(normalized.phone);
  const nickname = sanitizeText_(normalized.customerNickname);
  if (normalized.customerId) {
    return {
      customerId: normalized.customerId,
      name: name || normalized.customerName || "Walk-in",
      nickname,
    };
  }
  if (!name && !phone) {
    return { customerId: "", name: "Walk-in", nickname: "" };
  }
  const sheet = ensureCustomerSheet_(spreadsheet);
  const row = [
    Utilities.getUuid(),
    name || "Walk-in",
    nickname,
    phone,
    "",
    new Date(),
  ];
  sheet.appendRow(row);
  const customerObj = rowToObject_(row, CUSTOMER_HEADERS);
  return {
    customerId: customerObj.customer_id,
    name: customerObj.name,
    nickname: customerObj.nickname,
  };
}

function applySubscriptionAllowance_(spreadsheet, subscriptionId, items) {
  const sheet = ensureSubscriptionSheet_(spreadsheet);
  const subs = sheetToObjects_(sheet, SUBSCRIPTION_HEADERS);
  const foundIndex = subs.findIndex(sub => String(sub.subscription_id) === String(subscriptionId));
  if (foundIndex === -1) {
    return { items, remainingFree: null };
  }
  const sub = subs[foundIndex];
  if (!parseBoolean_(sub.active)) return { items, remainingFree: sub.remaining_free };
  const now = new Date();
  if (sub.start_date && sub.start_date > now) return { items, remainingFree: sub.remaining_free };
  if (sub.end_date && sub.end_date < now) return { items, remainingFree: sub.remaining_free };

  let remaining = Number(sub.remaining_free || sub.free_allowance || 0);
  const adjusted = items.map(item => {
    if (
      remaining > 0 &&
      sanitizeKey_(item.productType) === sanitizeKey_(sub.product_key) &&
      sanitizeKey_(item.serviceCode) === sanitizeKey_(sub.service_code)
    ) {
      const freeQty = Math.min(remaining, item.quantity);
      remaining -= freeQty;
      const chargeableQty = Math.max(0, item.quantity - freeQty);
      return Object.assign({}, item, {
        lineTotal: roundAmount_(item.unitPrice * chargeableQty),
        freeApplied: freeQty,
      });
    }
    return item;
  });

  subs[foundIndex].remaining_free = remaining;
  const values = subs.map(subRow => SUBSCRIPTION_HEADERS.map(key => subRow[key]));
  if (values.length) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, SUBSCRIPTION_HEADERS.length).setValues([SUBSCRIPTION_HEADERS]);
    sheet.getRange(2, 1, values.length, SUBSCRIPTION_HEADERS.length).setValues(values);
  }

  return { items: adjusted, remainingFree: remaining };
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

function ensureCatalogSheet_(spreadsheet) {
  const sheet = ensureSheet_(spreadsheet, CATALOG_SHEET, CATALOG_HEADERS);
  if (sheet.getLastRow() < 2) {
    seedCatalog_(sheet);
  }
  return sheet;
}

function ensureSettingsSheet_(spreadsheet) {
  return ensureSheet_(spreadsheet, SETTINGS_SHEET, SETTINGS_HEADERS);
}

function ensureCustomerSheet_(spreadsheet) {
  return ensureSheet_(spreadsheet, CUSTOMER_SHEET, CUSTOMER_HEADERS);
}

function ensureSubscriptionSheet_(spreadsheet) {
  return ensureSheet_(spreadsheet, SUBSCRIPTION_SHEET, SUBSCRIPTION_HEADERS);
}

function seedCatalog_(sheet) {
  const rows = DEFAULT_CATALOG.map(obj => CATALOG_HEADERS.map(header => obj[header]));
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, CATALOG_HEADERS.length).setValues(rows);
  }
}

function loadCatalog_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureCatalogSheet_(ss);
  const rows = sheetToObjects_(sheet, CATALOG_HEADERS);
  return {
    rows,
    priceBook: buildPriceBookFromRows_(rows),
  };
}

function loadSettings_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSettingsSheet_(ss);
  const rows = sheetToObjects_(sheet, SETTINGS_HEADERS);
  const map = rows.reduce((acc, row) => {
    acc[String(row.key || "").trim()] = row.value;
    return acc;
  }, {});
  return {
    logoUrl: map.logoUrl || DEFAULT_LOGO,
  };
}

function updateSettings_(updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSettingsSheet_(ss);
  const rows = sheetToObjects_(sheet, SETTINGS_HEADERS);
  const map = rows.reduce((acc, row) => {
    acc[String(row.key || "").trim()] = row.value;
    return acc;
  }, {});
  Object.keys(updates || {}).forEach(key => {
    map[key] = updates[key];
  });
  const entries = Object.keys(map).map(key => [key, map[key]]);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, SETTINGS_HEADERS.length).setValues([SETTINGS_HEADERS]);
  if (entries.length) {
    sheet.getRange(2, 1, entries.length, SETTINGS_HEADERS.length).setValues(entries);
  }
}

function loadCustomers_(spreadsheet, query) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureCustomerSheet_(ss);
  const rows = sheetToObjects_(sheet, CUSTOMER_HEADERS);
  if (!query) return rows;
  const q = String(query || "").toLowerCase();
  return rows.filter(row => {
    return (
      String(row.name || "").toLowerCase().includes(q) ||
      String(row.nickname || "").toLowerCase().includes(q) ||
      String(row.phone || "").toLowerCase().includes(q)
    );
  });
}

function saveCustomer(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureCustomerSheet_(ss);
  const name = sanitizeText_(payload && payload.name);
  const phone = sanitizeText_(payload && payload.phone);
  if (!name && !phone) throw new Error("Name or phone is required.");
  const row = [
    Utilities.getUuid(),
    name || "",
    sanitizeText_(payload.nickname),
    phone || "",
    sanitizeText_(payload.notes),
    new Date(),
  ];
  sheet.appendRow(row);
  return rowToObject_(row, CUSTOMER_HEADERS);
}

function getCustomers(query) {
  return loadCustomers_(null, query);
}

function loadSubscriptions_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSubscriptionSheet_(ss);
  return sheetToObjects_(sheet, SUBSCRIPTION_HEADERS);
}

function getSubscriptions() {
  return loadSubscriptions_();
}

function saveSubscription(payload) {
  if (!payload) throw new Error("Payload required.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSubscriptionSheet_(ss);
  const customerId = sanitizeKey_(payload.customer_id || payload.customerId);
  if (!customerId) throw new Error("customer_id is required.");
  const row = [
    Utilities.getUuid(),
    customerId,
    sanitizeText_(payload.customer_name || payload.customerName),
    sanitizeText_(payload.package_name || payload.packageName),
    sanitizeKey_(payload.product_key || payload.productKey),
    sanitizeKey_(payload.service_code || payload.serviceCode),
    Math.max(0, Number(payload.free_allowance || payload.freeAllowance) || 0),
    Math.max(0, Number(payload.remaining_free || payload.free_allowance || payload.freeAllowance) || 0),
    payload.start_date ? new Date(payload.start_date) : "",
    payload.end_date ? new Date(payload.end_date) : "",
    parseBoolean_(payload.active),
    new Date(),
  ];
  sheet.appendRow(row);
  return loadSubscriptions_(ss);
}

// Convert flat rows to a nested map that the UI and order logic can consume.
function buildPriceBookFromRows_(rows) {
  const priceBook = {};
  rows
    .filter(row => parseBoolean_(row.active) !== false)
    .forEach(row => {
      const productKey = sanitizeKey_(row.product_key);
      const serviceCode = sanitizeKey_(row.service_code);
      if (!productKey || !serviceCode) return;

      const productLabel = sanitizeText_(row.product_label_en) || productKey;
      const productLabelAr = sanitizeText_(row.product_label_ar) || productLabel;
      if (!priceBook[productKey]) {
        priceBook[productKey] = {
          label: productLabel,
          labelAr: productLabelAr,
          imageUrl: sanitizeUrl_(row.image_url),
          services: {},
        };
      }
      const serviceLabel = sanitizeText_(row.service_label_en) || serviceCode;
      const serviceLabelAr = sanitizeText_(row.service_label_ar) || serviceLabel;
      priceBook[productKey].services[serviceCode] = {
        label: serviceLabel,
        labelAr: serviceLabelAr,
        price: roundAmount_(row.price),
      };
    });
  return priceBook;
}

function normalizeCatalogPayload_(entries) {
  const normalized = entries.map(entry => {
    const productKey = sanitizeKey_(entry.product_key || entry.productKey);
    const serviceCode = sanitizeKey_(entry.service_code || entry.serviceCode);
    if (!productKey || !serviceCode) {
      throw new Error("product_key and service_code are required.");
    }
    return {
      product_key: productKey,
      service_code: serviceCode,
      product_label_en: sanitizeText_(entry.product_label_en || entry.productLabelEn || entry.product_label || entry.productLabel),
      product_label_ar: sanitizeText_(entry.product_label_ar || entry.productLabelAr),
      service_label_en: sanitizeText_(entry.service_label_en || entry.serviceLabelEn || entry.service_label || entry.serviceLabel),
      service_label_ar: sanitizeText_(entry.service_label_ar || entry.serviceLabelAr),
      price: roundAmount_(entry.price),
      image_url: sanitizeUrl_(entry.image_url || entry.imageUrl),
      active: parseBoolean_(entry.active) !== false,
    };
  });
  const hasCustom = normalized.some(entry => entry.product_key === "custom");
  if (!hasCustom) {
    const fallback = DEFAULT_CATALOG.find(entry => entry.product_key === "custom") || {
      product_key: "custom",
      service_code: "CUSTOM",
      product_label_en: "Custom item",
      product_label_ar: "عنصر مخصص",
      service_label_en: "Custom service",
      service_label_ar: "خدمة مخصصة",
      price: 0,
      image_url: "",
      active: true,
    };
    normalized.push(fallback);
  }
  return normalized;
}

function persistCatalogRows_(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureCatalogSheet_(ss);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, CATALOG_HEADERS.length).setValues([CATALOG_HEADERS]);
  if (rows.length) {
    const values = rows.map(row => CATALOG_HEADERS.map(key => row[key]));
    sheet.getRange(2, 1, values.length, CATALOG_HEADERS.length).setValues(values);
  }
}

function sheetToObjects_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const range = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return range.map(row => rowToObject_(row, headers));
}

function sheetToObjectsFlexible_(sheet, headers) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return [];
  return values.slice(1).filter(row => row.some(cell => cell !== "" && cell !== null)).map(row => {
    const obj = {};
    headers.forEach((header, idx) => {
      obj[header] = row[idx];
    });
    return obj;
  });
}

function rowToObject_(row, headers) {
  return headers.reduce((obj, header, idx) => {
    obj[header] = row[idx];
    return obj;
  }, {});
}

function buildHeaderIndex_(headers) {
  return headers.reduce((map, header, idx) => {
    map[header] = idx;
    return map;
  }, {});
}

function normalizeOrderPayload_(payload, priceBook) {
  if (!payload) throw new Error("Payload is required.");

  const walkInDate = payload.walkInDate ? new Date(payload.walkInDate) : new Date();
  const exitDate = payload.exitDate ? new Date(payload.exitDate) : walkInDate;
  if (!payload.items || !payload.items.length) {
    throw new Error("At least one item is required.");
  }

  const items = payload.items.map(item => normalizeItemPayload_(item, priceBook));
  const subtotal = roundAmount_(items.reduce((sum, item) => sum + item.lineTotal, 0));
  const discount = roundAmount_(Number(payload.discount) || 0);
  const total = roundAmount_(Math.max(0, subtotal - discount));

  return {
    customerId: sanitizeKey_(payload.customerId),
    customerName: sanitizeText_(payload.customerName) || "Walk-in",
    customerNickname: sanitizeText_(payload.customerNickname || payload.nickname),
    phone: sanitizeText_(payload.phone),
    subscriptionId: sanitizeKey_(payload.subscriptionId),
    walkInDate,
    exitDate,
    notes: sanitizeText_(payload.notes),
    items,
    subtotal,
    discount,
    total,
    createdBy: (Session.getActiveUser() && Session.getActiveUser().getEmail()) || "anonymous",
  };
}

function normalizeItemPayload_(item, priceBook) {
  const book = priceBook || {};
  const productType = sanitizeKey_(item.productType) || "custom";
  const serviceCode = sanitizeKey_(item.serviceCode) || "CUSTOM";
  const quantity = Math.max(1, Number(item.quantity) || 1);
  const unitPrice = resolvePrice_(productType, serviceCode, item.unitPrice, book);
  const lineTotal = roundAmount_(unitPrice * quantity);

  const product = book[productType] || {
    label: sanitizeText_(item.productLabel) || productType,
    labelAr: sanitizeText_(item.productLabelAr) || productType,
    services: {},
  };
  const service = (product.services && product.services[serviceCode]) || {
    label: sanitizeText_(item.serviceLabel) || "Custom",
    labelAr: sanitizeText_(item.serviceLabelAr) || "مخصص",
  };

  return {
    productType,
    productLabel: product.label,
    serviceCode,
    serviceLabel: service.label,
    quantity,
    unitPrice,
    lineTotal,
  };
}

function resolvePrice_(productType, serviceCode, overridePrice, priceBook) {
  const numericOverride = Number(overridePrice);
  if (!isNaN(numericOverride) && numericOverride >= 0) {
    return roundAmount_(numericOverride);
  }

  const product = priceBook && priceBook[productType];
  const service = product && product.services ? product.services[serviceCode] : null;
  if (service && typeof service.price === "number") {
    return roundAmount_(service.price);
  }
  return 0;
}

function generateOrderNumber_(sheet) {
  const count = Math.max(0, sheet.getLastRow() - 1) + 1;
  const datePart = Utilities.formatDate(new Date(), TIME_ZONE, "yyyyMMdd");
  return "LD-" + datePart + "-" + String(count).padStart(4, "0");
}

function buildPublicPriceBook_(priceBook) {
  const source = priceBook || loadCatalog_().priceBook || {};
  return Object.keys(source).reduce((acc, key) => {
    const product = source[key];
    acc[key] = {
      label: product.label,
      labelAr: product.labelAr,
      imageUrl: product.imageUrl,
      services: Object.keys(product.services || {}).reduce((svcAcc, svcKey) => {
        const svc = product.services[svcKey];
        svcAcc[svcKey] = {
          label: svc.label,
          labelAr: svc.labelAr,
          price: svc.price,
        };
        return svcAcc;
      }, {}),
    };
    return acc;
  }, {});
}

function sanitizeText_(value) {
  return String(value || "").trim();
}

function sanitizeKey_(value) {
  return String(value || "").trim();
}

function sanitizeUrl_(value) {
  return String(value || "").trim();
}

function parseBoolean_(value) {
  if (typeof value === "boolean") return value;
  const str = String(value || "").trim().toLowerCase();
  if (str === "") return true;
  return !(str === "false" || str === "0" || str === "no");
}

function roundAmount_(value) {
  return Math.round((Number(value) || 0) * 1000) / 1000;
}
