/**
 * Formats text in one or more sheets of the active spreadsheet from a specified starting row.
 *
 * @param {Object} [options]                                      - Configuration object.
 * @param {boolean} [options.applyToAll=true]                     - If true, applies formatting to all tabs.
 * @param {string[]} [options.sheetNames]                         - If applyToAll is false, an array of specific sheet names to format.
 * @param {string} [options.fontFamily="Calibri"]                - Font family to apply.
 * @param {number} [options.fontSize=10]                          - Font size to apply.
 * @param {boolean} [options.bold=false]                          - Whether text should be bold.
 * @param {string} [options.horizontalAlignment="CENTER"]        - Horizontal alignment: "LEFT", "CENTER", "RIGHT".
 * @param {string} [options.fontColor="#000000"]                 - Font color in hex (e.g. "#000000").
 * @param {number} [options.startRow=2]                            - 1-based row number from which to begin formatting.
 */
function formatText(options = {}) {
  const {
    applyToAll = true,
    sheetNames = [],
    fontFamily = "Calibri",
    fontSize = 10,
    bold = false,
    horizontalAlignment = "CENTER",
    fontColor = "#000000",
    startRow = 2
  } = options;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = [];

  if (applyToAll) {
    sheets = ss.getSheets();
  } else if (Array.isArray(sheetNames) && sheetNames.length) {
    sheetNames.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet) sheets.push(sheet);
    });
  } else {
    // No sheets to format if applyToAll is false and no valid names provided
    return;
  }

  sheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < startRow) return; // nothing to format
    const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol);
    range
      .setFontFamily(fontFamily)
      .setFontSize(fontSize)
      .setFontWeight(bold ? "bold" : "normal")
      .setHorizontalAlignment(horizontalAlignment)
      .setFontColor(fontColor);
  });
}

// Examples:
// formatText(); // formats from row 2 on all sheets with defaults
// formatText({ applyToAll: false, sheetNames: ['Sheet1','Output'], startRow: 5 });
// formatText({ fontFamily: 'Arial', fontSize: 12, bold: true, startRow: 3 });
// Full example with all parameters:
// formatText({
//   applyToAll: false,
//   sheetNames: ['Sheet1', 'Output'],
//   fontFamily: 'Times New Roman',
//   fontSize: 14,
//   bold: true,
//   horizontalAlignment: 'LEFT',
//   fontColor: '#FF0000',
//   startRow: 4
// });