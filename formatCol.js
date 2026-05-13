/**
 * Applies a specific format to one or more columns in a Google Sheet.
 *
 * @param {string} sourceSpreadsheetId - The ID of the spreadsheet containing the sheet.
 * @param {string} sheetName - The name of the sheet where formatting will be applied.
 * @param {string[]} columnLetters - An array of column letters to format (e.g., ['A', 'B', 'C']).
 * @param {string} formatType - The format type to apply. Valid options are:
 *   'TEXT', 'NUMBER', 'DECIMAL', 'DD/MM/YYYY', 'MM/YYYY', 'CURRENCY', and 'JOIN'.
 */
function formatColumns(sourceSpreadsheetId, sheetName, columnLetters, formatType) {
  const sheet = SpreadsheetApp.openById(sourceSpreadsheetId).getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet '${sheetName}' not found in spreadsheet '${sourceSpreadsheetId}'.`);
    return;
  }

  const validFormats = ['TEXT', 'NUMBER', 'DECIMAL', 'DD/MM/YYYY', 'MM/YYYY', 'CURRENCY', 'JOIN'];
  if (!validFormats.includes(formatType)) {
    Logger.log(`Invalid format: '${formatType}'. Valid formats are: ${validFormats.join(', ')}.`);
    return;
  }

  columnLetters.forEach(columnLetter => {
    let columnIndex;
    try {
      columnIndex = letterToCol(columnLetter);
    } catch (e) {
      Logger.log(e.message);
      return;
    }

    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(1, columnIndex, lastRow, 1);

    // JOIN: remove dots and concatenate digits, leave empty cells unchanged
    if (formatType === 'JOIN') {
      const newVals = range.getDisplayValues().map(([cell]) => {
        if (!cell) return [''];
        const joined = cell.replace(/\./g, '');
        return [isNaN(joined) ? joined : Number(joined)];
      });
      range.setValues(newVals);
      Logger.log(`JOIN applied: digits concatenated in column '${columnLetter}'.`);
      return;
    }

    // For other formats, apply number format patterns
    let formatPattern = '@';
    switch (formatType) {
      case 'NUMBER':
        formatPattern = '0';
        break;
      case 'DECIMAL':
        formatPattern = '0.00';
        break;
      case 'DD/MM/YYYY':
        formatPattern = 'dd/mm/yyyy';
        break;
      case 'MM/YYYY':
        formatPattern = 'mm/yyyy';
        break;
      case 'CURRENCY':
        formatPattern = 'R$ #,##0.00';
        break;
      // TEXT remains '@'
    }

    range.setNumberFormat(formatPattern);
    Logger.log(`Format '${formatType}' applied to column '${columnLetter}'.`);
  });

  Logger.log(`Operation complete: columns [${columnLetters.join(', ')}] on sheet '${sheetName}', format '${formatType}'.`);
}