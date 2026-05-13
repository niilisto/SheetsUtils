/**
 * Removes rows from the output sheet based on tickets present in the input sheet, with optional status filtering.
 *
 * @param {Object} config - Configuration object.
 * @param {Object} config.input - Input sheet config: {spreadsheetId, sheetName, startRow, ticketColumn, statusColumn, statusFilter}
 * @param {Object} config.output - Output sheet config: {spreadsheetId, sheetName, startRow, ticketColumn, statusColumn, statusFilter}
 *
 * @example
 * removeMatchingTickets({
 *   input: {
 *     spreadsheetId: 'inputSheetId',
 *     sheetName: 'Input',
 *     startRow: 2,
 *     ticketColumn: 'A',
 *     statusColumn: 'B',
 *     statusFilter: 'Closed'
 *   },
 *   output: {
 *     spreadsheetId: 'outputSheetId',
 *     sheetName: 'Output',
 *     startRow: 2,
 *     ticketColumn: 'A',
 *     statusColumn: 'B',
 *     statusFilter: 'Closed'
 *   }
 * });
 */
function removeMatchingTickets(config) {
  const { input, output } = config;

  // Get ticket values from input sheet, optionally filtered by status
  const inputData = getSheetData(input.spreadsheetId, input.sheetName, input.startRow);
  let inputTickets = inputData.map(row => columnConverter(input.ticketColumn) > 0 ? row[columnConverter(input.ticketColumn) - 1] : undefined);
  if (input.statusColumn && input.statusFilter) {
    inputTickets = inputData
      .filter(row => columnConverter(input.statusColumn) > 0 && row[columnConverter(input.statusColumn) - 1] === input.statusFilter)
      .map(row => columnConverter(input.ticketColumn) > 0 ? row[columnConverter(input.ticketColumn) - 1] : undefined);
  }

  // Get output sheet and its data
  const outputSheet = SpreadsheetApp.openById(output.spreadsheetId).getSheetByName(output.sheetName);
  const outputData = getSheetData(output.spreadsheetId, output.sheetName, output.startRow);

  // Find rows to remove
  const rowsToRemove = [];
  outputData.forEach((row, index) => {
    const ticket = columnConverter(output.ticketColumn) > 0 ? row[columnConverter(output.ticketColumn) - 1] : undefined;
    const status = output.statusColumn ? (columnConverter(output.statusColumn) > 0 ? row[columnConverter(output.statusColumn) - 1] : undefined) : null;
    const matchesStatus = !output.statusColumn || status === output.statusFilter;
    if (inputTickets.includes(ticket) && matchesStatus) {
      rowsToRemove.push(index + output.startRow); // Adjust for real row number
    }
  });

  // Remove rows in reverse order to preserve indices
  for (let i = rowsToRemove.length - 1; i >= 0; i--) {
    outputSheet.deleteRow(rowsToRemove[i]);
  }
}

/**
 * Reads data from a Google Sheet starting from a specific row.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {string} sheetName - The name of the sheet.
 * @param {number} startRow - The starting row (1-based).
 * @returns {Array[]} The values from the sheet.
 * @example
 * const data = getSheetData('sheetId', 'Sheet1', 2);
 */
function getSheetData(spreadsheetId, sheetName, startRow) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet '${sheetName}' not found.`);
    const range = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, sheet.getLastColumn());
    return range.getValues();
  } catch (error) {
    Logger.log(`Error getting sheet data: ${spreadsheetId}, sheet: ${sheetName}. Error: ${error}`);
    SpreadsheetApp.getUi().alert(
      `Error getting sheet data: ${spreadsheetId}, sheet: ${sheetName}. Check IDs, sheet names, and if the sheet exists.\n${error}`
    );
    return [];
  }
}

/**
 * Writes data to the end of a Google Sheet.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {string} sheetName - The name of the sheet.
 * @param {number} startRow - The starting row (unused, kept for compatibility).
 * @param {Array[]} data - The data to write.
 * @example
 * writeToSheet('sheetId', 'Sheet1', 2, [[1,2,3],[4,5,6]]);
 */
function writeToSheet(spreadsheetId, sheetName, startRow, data) {
  if (!data.length) return;
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  const range = sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length);
  range.setValues(data);
}