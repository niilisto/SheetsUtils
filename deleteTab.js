/**
 * Deletes the rows in a sheet starting from a specific row, without affecting rows above that range.
 * Rows are removed from the sheet, so the remaining data below shifts up.
 *
 * @param {string} spreadsheetID - Spreadsheet ID.
 * @param {string} spreadsheetName - Sheet name.
 * @param {number} startRow - First row to delete, inclusive.
 * @returns {number} Number of rows deleted.
 */
function deleteTab(spreadsheetID, spreadsheetName, startRow) {
  if (!spreadsheetID) {
    throw new Error('spreadsheetID is required.');
  }

  if (!spreadsheetName) {
    throw new Error('spreadsheetName is required.');
  }

  const firstRow = Number(startRow);
  if (!Number.isInteger(firstRow) || firstRow < 1) {
    throw new Error('startRow must be a positive integer.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetID);
  const sheet = ss.getSheetByName(spreadsheetName);

  if (!sheet) {
    throw new Error('Sheet not found: ' + spreadsheetName);
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < firstRow) {
    console.log(`No rows to delete in sheet "${spreadsheetName}" from row ${firstRow}.`);
    return 0;
  }

  const rowsToDelete = lastRow - firstRow + 1;
  sheet.deleteRows(firstRow, rowsToDelete);

  console.log(
    `Deleted ${rowsToDelete} row(s) in sheet "${spreadsheetName}" starting from row ${firstRow}.`
  );

  return rowsToDelete;
}