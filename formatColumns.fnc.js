/**
 * Applies a specific format to one or more columns in a Google Sheets spreadsheet.
 *
 * @param {string} sheetName - Name of the sheet where the formatting will be applied.
 * @param {string[]} columnLetters - Array of column letters to be formatted (e.g., ['A', 'B', 'C']).
 * @param {string} formatType - Format type: 'TEXT', 'NUMBER', 'DECIMAL', 'DD/MM/YYYY', 'MM/YYYY', 'CURRENCY', 'JOIN'.
 */
function formatCol(sheetName, columnLetters, formatType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet '${sheetName}' not found.`);
    return;
  }

  const validFormats = ['TEXT', 'NUMBER', 'DECIMAL', 'DD/MM/YYYY', 'MM/YYYY', 'CURRENCY', 'JOIN'];
  if (!validFormats.includes(formatType)) {
    Logger.log(`Invalid format: '${formatType}'. Valid formats: ${validFormats.join(', ')}.`);
    return;
  }

  // Letter to index conversion (A=1, B=2, ..., AA=27)
  function letterToCol(letter) {
    if (!letter || typeof letter !== 'string') return 0;
    let col = 0;
    const up = letter.toUpperCase();
    for (let i = 0; i < up.length; i++) {
      const code = up.charCodeAt(i);
      if (code < 65 || code > 90) return 0;
      col = col * 26 + (code - 64);
    }
    return col;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return;

  // Optimize: process all columns in a single batch when possible
  if (formatType === 'JOIN') {
    columnLetters.forEach(columnLetter => {
      const columnIndex = letterToCol(columnLetter);
      if (columnIndex === 0) {
        Logger.log(`Invalid column letter: '${columnLetter}'.`);
        return;
      }
      const range = sheet.getRange(1, columnIndex, lastRow, 1);
      const newVals = range.getDisplayValues().map(([cell]) => {
        if (!cell) return [''];
        const joined = cell.replace(/\./g, '');
        return [isNaN(joined) ? joined : Number(joined)];
      });
      range.setValues(newVals);
      Logger.log(`JOIN applied: digits concatenated in column '${columnLetter}'.`);
    });
  } else {
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

    // Batch apply format to all columns at once for performance
    const colIndexes = columnLetters.map(letterToCol).filter(idx => idx > 0);
    if (colIndexes.length === 0) {
      Logger.log('No valid columns to format.');
      return;
    }
    colIndexes.forEach((colIdx, i) => {
      const colLetter = columnLetters[i];
      if (colIdx === 0) {
        Logger.log(`Invalid column letter: '${colLetter}'.`);
        return;
      }
      const range = sheet.getRange(1, colIdx, lastRow, 1);
      range.setNumberFormat(formatPattern);
      Logger.log(`Format '${formatType}' applied to column '${colLetter}' in sheet '${sheetName}'.`);
    });
  }

  Logger.log(`Operation completed: columns [${columnLetters.join(', ')}] in sheet '${sheetName}', format '${formatType}'.`);
}