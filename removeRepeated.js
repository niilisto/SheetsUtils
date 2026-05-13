/**
 * Removes duplicate rows based on a referencer column (ticket ID) for specified sheets.
 * If more than one instance of the same ticket is found, duplicate rows are deleted,
 * leaving only the first occurrence.
 *
 * @param {string} spreadsheetId - ID of the target spreadsheet.
 * @param {Array<string>} sheetNames - List of sheet names to process.
 * @param {string} [refColumn='A'] - Column letter where ticket IDs are located.
 * @param {number} [startRow=2] - Row number to start checking for duplicates.
 */
function removeDuplicateTickets(spreadsheetId, sheetNames, refColumn = 'A', startRow = 2) {
  // Open spreadsheet by ID
  var ss = SpreadsheetApp.openById(spreadsheetId);

  // Process each specified sheet
  sheetNames.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;  // Skip if sheet not found

    var lastRow = sheet.getLastRow();
    if (lastRow < startRow) return;  // Nothing to process

    // Get all values in refColumn from startRow to lastRow
    var range = sheet.getRange(refColumn + startRow + ':' + refColumn + lastRow);
    var values = range.getValues();
    var seen = {};  // Track seen ticket IDs

    // Iterate from bottom to top to safely delete rows
    for (var i = values.length - 1; i >= 0; i--) {
      var ticket = values[i][0];
      var rowIndex = i + startRow;

      // If ticket is non-empty and already seen, delete the row
      if (ticket && seen[ticket]) {
        sheet.deleteRow(rowIndex);
      } else {
        seen[ticket] = true;  // Mark ticket as seen
      }
    }
  });
}

/**
 * Example wrapper with defaults for this specific spreadsheet:
 */
function runRemoveDuplicates() {
  removeDuplicateTickets(
    '19l8eX_MuHNdBOzBZXyoKDdefI3K_oElpUvPpTbeFxaY',
    ['DENTRO DE GARANTIA', 'FORA DE GARANTIA']
  );
}
