function processLatest1000OrderSummaryWithCityMapping() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // -- SHEET NAMES --
    const reservationsSheetName = "Reservations";
    const accountsSheetName = "Accounts";
    const citiesSheetName = "Cities";
    const detailsSheetName = "reservation_details";
    const productsSheetName = "products";
    const mediaSheetName = "Media";
    const categoriesSheetName = "Categories";
    const warehousesSheetName = "Warehouses";
    const outputSheetName = "Order";
    const reportSheetName = "Report";

    // -- SHEET REFERENCES --
    const reservationsSheet = spreadsheet.getSheetByName(reservationsSheetName);
    const accountsSheet = spreadsheet.getSheetByName(accountsSheetName);
    const citiesSheet = spreadsheet.getSheetByName(citiesSheetName);
    const detailsSheet = spreadsheet.getSheetByName(detailsSheetName);
    const productsSheet = spreadsheet.getSheetByName(productsSheetName);
    const mediaSheet = spreadsheet.getSheetByName(mediaSheetName);
    const categoriesSheet = spreadsheet.getSheetByName(categoriesSheetName);
    const warehousesSheet = spreadsheet.getSheetByName(warehousesSheetName);

    // -- CREATE OR CLEAR OUTPUT SHEETS --
    let outputSheet = spreadsheet.getSheetByName(outputSheetName);
    if (!outputSheet) {
      outputSheet = spreadsheet.insertSheet(outputSheetName);
    } else {
      outputSheet.clear();
    }

    let reportSheet = spreadsheet.getSheetByName(reportSheetName);
    if (!reportSheet) {
      reportSheet = spreadsheet.insertSheet(reportSheetName);
    } else {
      reportSheet.clear();
    }

    // -- VALIDATE REQUIRED SHEETS --
    if (
      !reservationsSheet ||
      !accountsSheet ||
      !citiesSheet ||
      !detailsSheet ||
      !productsSheet ||
      !mediaSheet ||
      !categoriesSheet ||
      !warehousesSheet
    ) {
      throw new Error("One or more required sheets are missing.");
    }

    // -- READ DATA FROM SHEETS --
    const reservationsData = reservationsSheet.getDataRange().getValues();
    const totalRows = reservationsData.length;
    // Start from row 1 if fewer than 1000 data rows exist
    const dataStartIndex = Math.max(1, totalRows - 1000);
    // Grab the last 1000 rows of data (excluding the header if present)
    const latestReservations = reservationsData.slice(dataStartIndex);

    const accountsData = accountsSheet.getDataRange().getValues();
    const citiesData = citiesSheet.getDataRange().getValues();
    const detailsData = detailsSheet.getDataRange().getValues();
    const productsData = productsSheet.getDataRange().getValues();
    const mediaData = mediaSheet.getDataRange().getValues();
    const categoriesData = categoriesSheet.getDataRange().getValues();
    const warehousesData = warehousesSheet.getDataRange().getValues();

    // -- CREATE MAPS (Accounts, Cities, Categories, Warehouses) --
    const accountsMap = new Map(accountsData.map(row => [row[0], row[4]]));
    const citiesMap = new Map(
      citiesData.map(row => {
        try {
          // JSON with city names in column C => row[2]
          const cityNames = JSON.parse(row[2]); // e.g., {ar: "...", en: "..."}
          const cityDisplayName = `${cityNames.ar}-${cityNames.en}`;
          return [Math.round(row[0]), cityDisplayName];
        } catch (error) {
          Logger.log(`Error parsing city name for ID ${row[0]}: ${error.message}`);
          return [Math.round(row[0]), "Unknown"];
        }
      })
    );

    const categoriesMap = new Map(
      categoriesData.map(row => {
        try {
          // JSON with category names in column D => row[3]
          const categoryNames = JSON.parse(row[3]);
          return [row[0], `${categoryNames.en}-${categoryNames.ar}`];
        } catch {
          return [row[0], "Unknown"];
        }
      })
    );

    const warehousesMap = new Map(
      warehousesData.map(row => {
        try {
          // JSON with warehouse names in column D => row[3]
          const warehouseNames = JSON.parse(row[3]);
          return [row[0], `${warehouseNames.en}-${warehouseNames.ar}`];
        } catch {
          return [row[0], "Unknown"];
        }
      })
    );

    // -- GROUP RESERVATION DETAILS BY RESERVATION ID (Column B => index 1) --
    const detailsGrouped = detailsData.reduce((map, row) => {
      const reservationId = row[1];
      if (!map[reservationId]) map[reservationId] = [];
      map[reservationId].push(row);
      return map;
    }, {});

    // ----------------------------------------------------------
    //  IMAGE MAPPING LOGIC (adjust column indexes if needed)
    // ----------------------------------------------------------
    // 1) Build a map from the Media sheet: modelCode => final image link
    //    We'll overwrite as we go so the last row for a given code wins
    const mediaCodeMap = new Map();

    for (let i = 1; i < mediaData.length; i++) {
      const row = mediaData[i];
      // For example:
      //   col A => folder ID   => row[0]
      //   col C => model code  => row[2]
      //   col G => file name   => row[6]  (IMPORTANT: filename is in Column G)

      const folder = row[0];
      const code = row[2];
      const fileName = row[6]; // Column G

      // Debugging: Log the fetched values
      Logger.log(`Media Row ${i + 1}: Folder=${folder}, Code=${code}, FileName=${fileName}`);

      if (folder && code && fileName) {
        // Build final link
        const finalLink = `https://natatiti.com/uploads/${folder}/${fileName}`;
        mediaCodeMap.set(String(code).trim(), finalLink);
        Logger.log(`Mapped ModelCode=${code} to Link=${finalLink}`);
      } else {
        Logger.log(`Incomplete data in Media row ${i + 1}: Skipping this row.`);
      }
    }

    // 2) Function to extract the "model code" from the product code, e.g. "NP-003-340" => "340"
    function extractModelCode(productCode) {
      if (typeof productCode !== "string") return "";
      const parts = productCode.split("-");
      return parts.length > 2 ? parts[2].trim() : "";
    }

    // 3) Initialize productsMap before updating image links
    let productsMap = new Map(productsData.map(row => [row[0], row]));

    // 4) Update product image links in the Products sheet (column N => index 13) for ALL rows
    //    Retain existing links if no new link is found
    const updatedImageLinks = productsData.map((row, index) => {
      if (index === 0) {
        // Header row, retain as-is
        return row[13];
      }

      const productCode = row[6]; // Column G
      const modelCode = extractModelCode(productCode);
      const newLink = mediaCodeMap.get(modelCode);

      if (newLink) {
        Logger.log(`Product Row ${index + 1}: ProductCode=${productCode}, ModelCode=${modelCode}, New ImageLink=${newLink}`);
        return newLink; // Update with new link
      } else {
        Logger.log(`Product Row ${index + 1}: ProductCode=${productCode}, ModelCode=${modelCode}, Retaining Existing ImageLink=${row[13]}`);
        return row[13]; // Retain existing link
      }
    });

    // 5) Write updated Image Links back to Column N (Index 13)
    //    Starting from row 2 to skip the header
    const imageLinksToWrite = updatedImageLinks.slice(1).map(link => [link]);
    productsSheet.getRange(2, 14, imageLinksToWrite.length, 1).setValues(imageLinksToWrite);
    SpreadsheetApp.flush(); // Ensure all pending changes are applied

    // 6) Rebuild the productsMap from updated productsData with new image links
    for (let i = 1; i < productsData.length; i++) {
      const productId = productsData[i][0]; // Column A
      const updatedLink = updatedImageLinks[i];
      // Create a new product row with the updated image link in column N (index 13)
      const updatedProductRow = [
        ...productsData[i].slice(0, 13),
        updatedLink,
        ...productsData[i].slice(14)
      ];
      productsMap.set(productId, updatedProductRow);
    }

    // -- PREPARE OUTPUT HEADERS --
    // reservationsData[0] is presumably the header row
    const reservationsHeaders = reservationsData[0];
    const outputHeaders = [
      ...reservationsHeaders,
      "ReservationDetails ID",
      "Product Name",
      "Product Image",
      "Category Name",
      "Warehouse Name",
      "Tarkeeb Date",
      "Tarkeeb Time",
      "Sheel Date",
      "Sheel Time"
    ];
    const outputData = [outputHeaders];

    // -- HELPER: PARSE DATE/TIME --
    function parseDateTime(dateTime) {
      if (!dateTime) return ["", ""];
      // e.g., "2025-01-30T14:07:00" => ["2025-01-30", "14:07:00"]
      const [date, time] = String(dateTime).split("T");
      return [date || "", time || ""];
    }

    // -- PROCESS LATEST 1000 RESERVATIONS --
    latestReservations.forEach((reservationRow, index) => {
      // Skip the case where we might have the header row included in our slice
      if (index === 0 && reservationRow[0] === reservationsHeaders[0]) {
        return;
      }
      const reservationId = reservationRow[0];

      // Pickup/Dropoff columns (adjust as needed)
      const pickupDateTime = reservationRow[15];  // col P => index 15
      const dropoffDateTime = reservationRow[16]; // col Q => index 16

      const [tarkeebDate, tarkeebTime] = parseDateTime(pickupDateTime);
      const [sheelDate, sheelTime] = parseDateTime(dropoffDateTime);

      // Map accounts in columns B, C, D => indices 1, 2, 3
      const mappedRow = [...reservationRow];
      [1, 2, 3].forEach(colIndex => {
        const accountId = mappedRow[colIndex];
        mappedRow[colIndex] = accountsMap.get(accountId) || accountId;
      });

      // Map city ID in column E => index 4
      const cityIdColumnIndex = 4;
      const cityId = Math.round(mappedRow[cityIdColumnIndex]);
      const cityName = citiesMap.get(cityId) || "Unknown";
      mappedRow[cityIdColumnIndex] = cityName;

      // Retrieve reservation details from grouped data
      const relatedDetails = detailsGrouped[reservationId] || [];
      // For each detail, produce an output row
      relatedDetails.forEach(detailRow => {
        const detailId = detailRow[0];   // "ReservationDetails ID"
        const productId = detailRow[2];  // product ID
        const product = productsMap.get(productId) || [];

        // Indices in productsData:
        //   product[7] => product name
        //   product[13] => product image link (the one we just updated)
        //   product[4] => warehouse ID
        //   product[5] => category ID
        const productName = product[7] || "";
        const productImage = product[13] || "";
        const warehouseName = warehousesMap.get(product[4]) || "Unknown";
        const categoryName = categoriesMap.get(product[5]) || "Unknown";

        // Build the final combined row
        const outputRow = [
          ...mappedRow,
          detailId,
          productName,
          productImage,
          categoryName,
          warehouseName,
          tarkeebDate,
          tarkeebTime,
          sheelDate,
          sheelTime
        ];
        outputData.push(outputRow);
      });
    });

    // -- WRITE TO "Order" SHEET --
    if (outputData.length > 1) { // Check if there are data rows beyond headers
      outputSheet
        .getRange(1, 1, outputData.length, outputHeaders.length)
        .setValues(outputData);
    } else {
      Logger.log("No data rows to write to the Order sheet.");
    }

    // -- FINISH UP --
    reportSheet.getRange(1, 1).setValue("Latest 1,000 reservations processed successfully!");
    Logger.log("Process completed with updated product images and reservation data.");

  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    // Optionally write the error to the Report sheet as well
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let reportSheet = spreadsheet.getSheetByName("Report");
    if (reportSheet) {
      reportSheet.getRange(1, 1).setValue(`Error: ${error.message}`);
    }
  }
}
