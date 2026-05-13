/**
 * Ensures the sheet always has exactly `blankRows` blank rows below the last non-empty row,
 * also considering as empty the cells whose value is the empty string ("").
 *
 * @param {object} options
 * @param {string} options.sheetName  Name of the sheet (e.g., 'Main')
 * @param {number} options.blankRows  Number of blank rows to maintain below the last data row
 *
 * @example
 * utils.maintainBlankRows({ sheetName: 'Main', blankRows: 10 });
 */
function maintainBlankRows(options) {
  const { sheetName, blankRows } = options;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  // Get all values from the used range
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Find the last row that contains at least one "real" value
  // (i.e., different from "" or null)
  let lastDataRow = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    // .some() returns true if any value is different from "" and different from null
    if (values[i].some(cell => cell !== "" && cell !== null)) {
      lastDataRow = i + 1;
      break;
    }
  }

  // If nothing was found (everything is empty), consider lastDataRow = 0
  // and then we would want only blankRows blank lines
  const desiredTotalRows = lastDataRow + blankRows;

  const currentRows = sheet.getMaxRows();

  if (currentRows < desiredTotalRows) {
    const rowsToAdd = desiredTotalRows - currentRows;
    sheet.insertRowsAfter(currentRows, rowsToAdd);
    Logger.log("Added " + rowsToAdd + " rows to reach " + desiredTotalRows + " total rows.");
  }
  else if (currentRows > desiredTotalRows) {
    const rowsToDelete = currentRows - desiredTotalRows;
    sheet.deleteRows(desiredTotalRows + 1, rowsToDelete);
    Logger.log("Deleted " + rowsToDelete + " rows to reduce to " + desiredTotalRows + " total rows.");
  }
  else {
    Logger.log("No change needed. Already at " + currentRows + " total rows.");
  }
}
