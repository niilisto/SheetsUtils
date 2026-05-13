/**
 * Clears the data in a sheet starting from a specific row, without affecting rows above that range.
 * Only cell contents are removed; rows are not deleted and formatting is preserved.
 *
 * @param {string} spreadsheetID - Spreadsheet ID.
 * @param {string} spreadsheetName - Sheet name.
 * @param {number} startRow - First row to clear, inclusive.
 * @returns {number} Number of rows cleared.
 */
function clearTab(spreadsheetID, spreadsheetName, startRow) {
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
  const lastCol = sheet.getLastColumn();

  if (lastRow < firstRow || lastCol < 1) {
    console.log(`No data to clear in sheet "${spreadsheetName}" from row ${firstRow}.`);
    return 0;
  }

  const rowsToClear = lastRow - firstRow + 1;
  sheet.getRange(firstRow, 1, rowsToClear, lastCol).clearContent();

  console.log(
    `Cleared ${rowsToClear} row(s) in sheet "${spreadsheetName}" starting from row ${firstRow}.`
  );

  return rowsToClear;
}