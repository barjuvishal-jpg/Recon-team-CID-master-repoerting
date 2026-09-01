function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Master Import")
    .addItem("Refresh Source Tabs", "refreshSourceTabs")
    .addSeparator()
    .addItem("Sync Data", "syncData")
    .addToUi();
}

/**
 * Refresh source tab dropdowns
 */
function refreshSourceTabs() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName("Config");

  const sourceUrl = config
    .getRange("B2")
    .getDisplayValue()
    .trim();

  if (!sourceUrl) {
    SpreadsheetApp.getUi().alert(
      "Please enter Source Spreadsheet URL in Config!B2"
    );
    return;
  }

  try {

    const sourceSS = SpreadsheetApp.openByUrl(sourceUrl);

    const sheetNames = sourceSS
      .getSheets()
      .map(sheet => sheet.getName());

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(sheetNames, true)
      .setAllowInvalid(false)
      .build();

    config.getRange("B9:B10").setDataValidation(rule);

    SpreadsheetApp.getUi().alert(
      `Success!\n\n${sheetNames.length} tabs loaded into dropdowns.`
    );

  } catch (error) {

    SpreadsheetApp.getUi().alert(
      "Unable to access source spreadsheet.\n\n" +
      error.message
    );

  }
}

/**
 * Import selected tabs
 */
function syncData() {

  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const config = masterSS.getSheetByName("Config");

  const sourceUrl = config
    .getRange("B2")
    .getDisplayValue()
    .trim();

  if (!sourceUrl) {
    SpreadsheetApp.getUi().alert(
      "Please enter Source Spreadsheet URL in Config!B2"
    );
    return;
  }

  try {

    const sourceSS = SpreadsheetApp.openByUrl(sourceUrl);

    const mappings = config
      .getRange("A9:D10")
      .getValues();

    let totalImported = 0;
    let closeRows = 0;
    let forwardedRows = 0;

    mappings.forEach(mapping => {

      const importType = mapping[0];
      const sourceTab = mapping[1];
      const destinationTab = mapping[2];
      const clearBeforeImport = mapping[3];

      if (!sourceTab || !destinationTab) return;

      const sourceSheet =
        sourceSS.getSheetByName(sourceTab);

      const destinationSheet =
        masterSS.getSheetByName(destinationTab);

      if (!sourceSheet) {
        throw new Error(
          `Source tab not found: ${sourceTab}`
        );
      }

      if (!destinationSheet) {
        throw new Error(
          `Destination tab not found: ${destinationTab}`
        );
      }

      const data = sourceSheet
        .getDataRange()
        .getValues();

      if (data.length <= 1) return;

      // Clear existing data but keep headers
      if (
        clearBeforeImport === true ||
        String(clearBeforeImport).toUpperCase() === "TRUE"
      ) {

        const lastRow =
          destinationSheet.getLastRow();

        if (lastRow > 1) {

          destinationSheet
            .getRange(
              2,
              1,
              lastRow - 1,
              destinationSheet.getMaxColumns()
            )
            .clearContent();
        }
      }

      // Skip header row
      const rows = data.slice(1);

      destinationSheet
        .getRange(
          destinationSheet.getLastRow() + 1,
          1,
          rows.length,
          rows[0].length
        )
        .setValues(rows);

      totalImported += rows.length;

      if (importType === "Close Cases") {
        closeRows += rows.length;
      }

      if (importType === "Forwarded Cases") {
        forwardedRows += rows.length;
      }

    });

    // Update status
    config.getRange("B3").setValue(new Date());
    config.getRange("B4").setValue("Success");
    config.getRange("B5").setValue(totalImported);

    // Sync Log
    const logRow = Math.max(
      config.getLastRow() + 1,
      16
    );

    config.getRange(logRow, 1, 1, 4).setValues([
      [
        new Date(),
        closeRows,
        forwardedRows,
        "Success"
      ]
    ]);

    SpreadsheetApp.getUi().alert(
      "Import Completed\n\n" +
      "Close Cases: " + closeRows + " rows\n" +
      "Forwarded Cases: " + forwardedRows + " rows\n" +
      "Total Imported: " + totalImported + " rows"
    );

  } catch (error) {

    config.getRange("B3").setValue(new Date());
    config.getRange("B4").setValue("Failed");

    SpreadsheetApp.getUi().alert(
      "Import Failed\n\n" +
      error.message
    );
  }
}