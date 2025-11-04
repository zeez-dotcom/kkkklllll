function processAllOrderSummaryWithCityMapping() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Sheet names
    const reservationsSheetName = "Reservations";
    const accountsSheetName = "Accounts";
    const citiesSheetName = "Cities";
    const detailsSheetName = "reservation_details";
    const productsSheetName = "Products";
    const mediaSheetName = "Media";
    const categoriesSheetName = "Categories";
    const warehousesSheetName = "Warehouses";
    const outputSheetName = "All Orders";
    const reportSheetName = "Report";

    // Get sheets
    const reservationsSheet = spreadsheet.getSheetByName(reservationsSheetName);
    const accountsSheet = spreadsheet.getSheetByName(accountsSheetName);
    const citiesSheet = spreadsheet.getSheetByName(citiesSheetName);
    const detailsSheet = spreadsheet.getSheetByName(detailsSheetName);
    const productsSheet = spreadsheet.getSheetByName(productsSheetName);
    const mediaSheet = spreadsheet.getSheetByName(mediaSheetName);
    const categoriesSheet = spreadsheet.getSheetByName(categoriesSheetName);
    const warehousesSheet = spreadsheet.getSheetByName(warehousesSheetName);
    let outputSheet = spreadsheet.getSheetByName(outputSheetName);

    if (!outputSheet) {
      outputSheet = spreadsheet.insertSheet(outputSheetName);
    } else {
      outputSheet.clear();
    }

    let reportSheet = spreadsheet.getSheetByName(reportSheetName);
    if (!reportSheet) {
      reportSheet = spreadsheet.insertSheet(reportSheetName);
    }

    // Validate required sheets
    if (!reservationsSheet || !accountsSheet || !citiesSheet || !detailsSheet || !productsSheet) {
      throw new Error("One or more required sheets are missing.");
    }

    // Read data
    const reservationsData = reservationsSheet.getDataRange().getValues();
    const allReservations = reservationsData.slice(1); // All reservations except header
    const reservationsHeaders = reservationsData[0];
    const accountsData = accountsSheet.getDataRange().getValues();
    const citiesData = citiesSheet.getDataRange().getValues();
    const detailsData = detailsSheet.getDataRange().getValues();
    const productsData = productsSheet.getDataRange().getValues();
    const mediaData = mediaSheet.getDataRange().getValues();
    const categoriesData = categoriesSheet.getDataRange().getValues();
    const warehousesData = warehousesSheet.getDataRange().getValues();

    // Maps
    const accountsMap = new Map(accountsData.map(row => [row[0], row[4]]));
    const citiesMap = new Map(citiesData.map(row => {
      try {
        const cityNames = JSON.parse(row[2]); // Parse JSON from column C in Cities
        const cityDisplayName = `${cityNames.ar}-${cityNames.en}`;
        return [Math.round(row[0]), cityDisplayName];
      } catch (error) {
        Logger.log(`Error parsing city name for ID ${row[0]}: ${error.message}`);
        return [Math.round(row[0]), "Unknown"];
      }
    }));
    const detailsGrouped = detailsData.reduce((map, row) => {
      const reservationId = row[1];
      if (!map[reservationId]) map[reservationId] = [];
      map[reservationId].push(row);
      return map;
    }, {});
    const productsMap = new Map(productsData.map(row => [row[0], row]));
    const mediaMap = new Map(mediaData.map(row => {
      const folder = row[0];
      const fileName = row[6];
      return row[2] ? [row[2], `https://natatiti.com/uploads/${folder}/${fileName}`] : null;
    }).filter(Boolean));
    const categoriesMap = new Map(categoriesData.map(row => {
      try {
        const categoryNames = JSON.parse(row[3]);
        return [row[0], `${categoryNames.en}-${categoryNames.ar}`];
      } catch {
        return [row[0], "Unknown"];
      }
    }));
    const warehousesMap = new Map(warehousesData.map(row => {
      try {
        const warehouseNames = JSON.parse(row[3]);
        return [row[0], `${warehouseNames.en}-${warehouseNames.ar}`];
      } catch {
        return [row[0], "Unknown"];
      }
    }));

    // Prepare output headers
    const outputHeaders = [
      ...reservationsHeaders,
      "ReservationDetails IDs",
      "Product Names",
      "Product Images",
      "Category Names",
      "Warehouse Names",
      "Tarkeeb Date",
      "Tarkeeb Time",
      "Sheel Date",
      "Sheel Time"
    ];

    // To store processed data before writing
    const finalData = [outputHeaders];

    const parseDateTime = (dateTime) => {
      if (!dateTime) return ["", ""];
      const [date, time] = String(dateTime).split("T");
      return [date || "", time || ""];
    };

    // Process all reservations by first building a map: reservationId -> { reservationRow, product info arrays }
    const reservationsMap = new Map();

    allReservations.forEach(reservationRow => {
      const reservationId = reservationRow[0];

      // Map account IDs (B, C, D)
      const mappedRow = [...reservationRow];
      [1, 2, 3].forEach(colIndex => {
        const accountId = mappedRow[colIndex];
        mappedRow[colIndex] = accountsMap.get(accountId) || accountId;
      });

      // Map city ID (Column E -> Index 4)
      const cityIdColumnIndex = 4;
      const cityId = Math.round(mappedRow[cityIdColumnIndex]);
      const cityName = citiesMap.get(cityId) || "Unknown";
      mappedRow[cityIdColumnIndex] = cityName;

      // Parse Tarkeeb / Sheel date time from reservation (Columns P,Q -> Index 15,16)
      const pickupDateTime = mappedRow[15];
      const dropoffDateTime = mappedRow[16];
      const [tarkeebDate, tarkeebTime] = parseDateTime(pickupDateTime);
      const [sheelDate, sheelTime] = parseDateTime(dropoffDateTime);

      reservationsMap.set(reservationId, {
        reservationRow: mappedRow,
        tarkeebDate,
        tarkeebTime,
        sheelDate,
        sheelTime,
        detailIds: [],
        productNames: [],
        productImages: [],
        categoryNames: [],
        warehouseNames: []
      });
    });

    // Fill product arrays for each reservation
    for (let reservationId in detailsGrouped) {
      const detailRows = detailsGrouped[reservationId];
      const reservationEntry = reservationsMap.get(Number(reservationId));
      if (!reservationEntry) continue;

      detailRows.forEach(detailRow => {
        const productId = detailRow[2];
        const product = productsMap.get(productId) || [];
        const productName = product[7] || "";
        const productImage = mediaMap.get(productId) || "";
        const categoryName = categoriesMap.get(product[5]) || "Unknown";
        const warehouseName = warehousesMap.get(product[4]) || "Unknown";

        // Push all info; will remove duplicates later
        reservationEntry.detailIds.push(detailRow[0]);
        reservationEntry.productNames.push(productName);
        reservationEntry.productImages.push(productImage);
        reservationEntry.categoryNames.push(categoryName);
        reservationEntry.warehouseNames.push(warehouseName);
      });
    }

    // Build final rows: one per reservation
    for (const [reservationId, entry] of reservationsMap.entries()) {
      // Remove duplicates for ReservationDetails IDs and Warehouse Names
      const uniqueDetailIds = Array.from(new Set(entry.detailIds));
      const uniqueWarehouseNames = Array.from(new Set(entry.warehouseNames));

      const combinedDetailIds = uniqueDetailIds.join(", ");
      const combinedProductNames = entry.productNames.join(", ");
      const combinedProductImages = entry.productImages.join(", ");
      const combinedCategoryNames = entry.categoryNames.join(", ");
      const combinedWarehouseNames = uniqueWarehouseNames.join(", ");

      const outputRow = [
        ...entry.reservationRow,
        combinedDetailIds,
        combinedProductNames,
        combinedProductImages,
        combinedCategoryNames,
        combinedWarehouseNames,
        entry.tarkeebDate,
        entry.tarkeebTime,
        entry.sheelDate,
        entry.sheelTime
      ];
      finalData.push(outputRow);
    }

    // Write to the All Orders sheet
    outputSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);

    // Clear and update the Report sheet
    reportSheet.clear();
    reportSheet.getRange(2, 2).setValue("Please wait");

    Logger.log("All reservations processed successfully with combined products and no duplicates for IDs and warehouse names!");
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
  }
}
