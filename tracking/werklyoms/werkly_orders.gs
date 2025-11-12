function loadOrderDataset_() {
  const sheet = getOrderSheet_();
  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return {
      allOrders: [],
      maintenanceOrders: [],
      availableDates: [],
      locationSummary: {}
    };
  }

  const orderMap = new Map();
  const availableDates = new Set();
  const locationSummary = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) {
      continue;
    }

    const orderNumber = String(row[0]).trim();
    if (!orderNumber) {
      continue;
    }

    const tarkeebDate = row[24];
    const sheelDate = row[26];
    const normalizedTarkeebDate = tarkeebDate ? normalizeDate(tarkeebDate) : '';
    const normalizedSheelDate = sheelDate ? normalizeDate(sheelDate) : '';

    const productName = (row[20] && String(row[20]).trim()) || 'Unnamed Product';
    const productImage = (row[21] && String(row[21]).trim()) || 'https://via.placeholder.com/80';
    const productCategory = (row[22] && String(row[22]).trim()) || 'Uncategorized';
    const isCancelledProduct = productName.indexOf('CANCEL - كنسل') !== -1;
    const isMaintenanceProduct = productName.indexOf('Maintenance') !== -1 || productName.indexOf('صيانة') !== -1;

    if (!orderMap.has(orderNumber)) {
      orderMap.set(orderNumber, {
        orderNumber,
        customerName: (row[6] && String(row[6]).trim()) || 'N/A',
        phone: (row[9] && String(row[9]).trim()) || 'N/A',
        location: (row[4] && String(row[4]).trim()) || 'Unknown',
        comments: (row[8] && String(row[8]).trim()) || 'No comments',
        tarkeebDate: normalizedTarkeebDate,
        sheelDate: normalizedSheelDate,
        tarkeebTime: row[25] ? normalizeTime(row[25]) : '',
        sheelTime: row[27] ? normalizeTime(row[27]) : '',
        paymentMethod: (row[11] && String(row[11]).trim()) || 'N/A',
        products: [],
        grandTotal: parseAmount_(row[14]),
        isCancelled: isCancelledProduct,
        isMaintenanceOrder: false,
        maintenanceDates: []
      });
    }

    const order = orderMap.get(orderNumber);
    order.products.push({
      productName,
      productCategory,
      productImage
    });
    order.isCancelled = order.isCancelled || isCancelledProduct;
    order.grandTotal = parseAmount_(row[14]) || order.grandTotal || 0;
    if (!order.tarkeebDate && normalizedTarkeebDate) {
      order.tarkeebDate = normalizedTarkeebDate;
    }
    if (!order.sheelDate && normalizedSheelDate) {
      order.sheelDate = normalizedSheelDate;
    }

    if (isMaintenanceProduct) {
      order.isMaintenanceOrder = true;
      if (tarkeebDate && sheelDate) {
        const mergedDates = new Set(order.maintenanceDates);
        const rangeDates = getDatesInRange(new Date(tarkeebDate), new Date(sheelDate));
        rangeDates
          .map(function(date) { return normalizeDate(date); })
          .forEach(function(dateStr) {
            if (dateStr) {
              mergedDates.add(dateStr);
            }
          });
        order.maintenanceDates = Array.from(mergedDates);
      }
    }

    if (normalizedTarkeebDate) {
      availableDates.add(normalizedTarkeebDate);
    }
    if (normalizedSheelDate) {
      availableDates.add(normalizedSheelDate);
    }

    const locationKey = (row[4] && String(row[4]).trim()) || 'Unknown';
    if (!locationSummary[locationKey]) {
      locationSummary[locationKey] = {
        tarkeeb: 0,
        sheel: 0,
        cancelled: 0,
        totalOrders: 0
      };
    }
    if (normalizedTarkeebDate) {
      locationSummary[locationKey].tarkeeb++;
    }
    if (normalizedSheelDate) {
      locationSummary[locationKey].sheel++;
    }
    if (order.isCancelled) {
      locationSummary[locationKey].cancelled++;
    }
    locationSummary[locationKey].totalOrders++;
  }

  const allOrders = Array.from(orderMap.values());
  const maintenanceOrders = allOrders.filter(function(order) { return order.isMaintenanceOrder; });

  return {
    allOrders,
    maintenanceOrders,
    availableDates: Array.from(availableDates),
    locationSummary
  };
}

function getOrderSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Order');
  if (!sheet) {
    throw new Error("'Order' sheet not found.");
  }
  return sheet;
}
