/**
 * Compares two date columns row by row and clears only the configured column
 * when the comparison column contains a newer date.
 *
 * The column configured for clearing is never swapped automatically.
 * If the clear column date is newer or equal, nothing happens.
 * If the comparison column is empty and the clear column has a value, the clear column is cleared.
 *
 * @param {Object} config
 * @param {string} config.sourceSpreadsheetId - Spreadsheet ID for the sheet to clear.
 * @param {string} config.sourceSheetName - Sheet name for the sheet to clear.
 * @param {number} config.sourceStartRow - 1-based row number to start processing.
 * @param {string} config.clearColumn - Column letter that may be cleared.
 * @param {string} config.compareColumn - Column letter used for date comparison.
 * @param {string} [config.compareSpreadsheetId] - Spreadsheet ID for the comparison sheet.
 * @param {string} [config.compareSheetName] - Sheet name for the comparison sheet.
 * @param {number} [config.compareStartRow] - Optional starting row for the comparison sheet.
 * @param {string} [config.spreadsheetId] - Legacy alias for sourceSpreadsheetId.
 * @param {string} [config.sheetName] - Legacy alias for sourceSheetName.
 * @param {number} [config.startRow] - Legacy alias for sourceStartRow.
 * @returns {number} Number of cells cleared.
 */
function clearCellIfComparisonDateIsNewer(config) {
  const {
    sourceSpreadsheetId,
    sourceSheetName,
    sourceStartRow,
    clearColumn,
    compareColumn,
    compareSpreadsheetId = sourceSpreadsheetId || config.spreadsheetId,
    compareSheetName = sourceSheetName || config.sheetName,
    compareStartRow = sourceStartRow || config.startRow,
    spreadsheetId,
    sheetName,
    startRow
  } = config;

  const letterToCol = letter => {
    const normalized = String(letter).toUpperCase();
    if (!/^[A-Z]+$/.test(normalized)) {
      throw new Error(`Invalid column letter: "${letter}".`);
    }
    return normalized.split('').reduce((accumulator, character) => accumulator * 26 + character.charCodeAt(0) - 64, 0);
  };

  const toComparableStamp = value => {
    const toDateOnlyStamp = dateObj => Date.UTC(
      dateObj.getFullYear(),
      dateObj.getMonth(),
      dateObj.getDate()
    );

    if (value instanceof Date && !isNaN(value.getTime())) {
      return toDateOnlyStamp(value);
    }

    if (value == null) {
      return null;
    }

    if (typeof value === 'number' && isFinite(value)) {
      const numericDate = new Date(value);
      return isNaN(numericDate.getTime()) ? null : toDateOnlyStamp(numericDate);
    }

    const text = String(value).trim();
    if (!text) {
      return null;
    }

    const parseDateParts = (year, month, day, hour = '0', minute = '0', second = '0') => {
      const parsedDate = new Date(
        Number(year),
        Number(month) - 1,
        Number(day),
        Number(hour),
        Number(minute),
        Number(second),
        0
      );

      return isNaN(parsedDate.getTime()) ? null : toDateOnlyStamp(parsedDate);
    };

    const patterns = [
      /^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/,
      /^(\d{2})-(\d{2})-(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/,
      /^(\d{4})\/(\d{2})\/(\d{2})(?:\s+(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/,
      /^(\d{4})-(\d{2})-(\d{2})(?:[T\s](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/,
      /^(\d{2})\/(\d{2})\/(\d{4})(?:[T\s](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/,
      /^(\d{2})-(\d{2})-(\d{4})(?:[T\s](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/
    ];

    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (!match) continue;

      if (pattern.source.startsWith('^(\\d{4})')) {
        const [, year, month, day, hour, minute, second] = match;
        const stamp = parseDateParts(year, month, day, hour, minute, second);
        if (stamp != null) return stamp;
      } else {
        const [, first, second, year, hour, minute, secondPart] = match;
        const dayFirstStamp = parseDateParts(year, second, first, hour, minute, secondPart);
        const monthFirstStamp = parseDateParts(year, first, second, hour, minute, secondPart);

        if (dayFirstStamp != null && monthFirstStamp != null) {
          if (Number(first) > 12) return dayFirstStamp;
          if (Number(second) > 12) return monthFirstStamp;
          return dayFirstStamp;
        }

        if (dayFirstStamp != null) return dayFirstStamp;
        if (monthFirstStamp != null) return monthFirstStamp;
      }
    }

    const parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
      return toDateOnlyStamp(parsed);
    }

    return null;
  };

  const targetSpreadsheetId = sourceSpreadsheetId || spreadsheetId;
  const targetSheetName = sourceSheetName || sheetName;
  const targetStartRow = sourceStartRow || startRow;

  if (!targetSpreadsheetId) {
    throw new Error('sourceSpreadsheetId is required.');
  }

  if (!targetSheetName) {
    throw new Error('sourceSheetName is required.');
  }

  if (!clearColumn) {
    throw new Error('clearColumn is required.');
  }

  if (!compareColumn) {
    throw new Error('compareColumn is required.');
  }

  const firstRow = Number(targetStartRow);
  const compareFirstRow = Number(compareStartRow);
  if (!Number.isInteger(firstRow) || firstRow < 1) {
    throw new Error('startRow must be a positive integer.');
  }
  if (!Number.isInteger(compareFirstRow) || compareFirstRow < 1) {
    throw new Error('compareStartRow must be a positive integer.');
  }

  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  const targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
  if (!targetSheet) {
    throw new Error(`Sheet not found: ${targetSheetName}`);
  }

  const comparisonSpreadsheet = compareSpreadsheetId === targetSpreadsheetId
    ? targetSpreadsheet
    : SpreadsheetApp.openById(compareSpreadsheetId);
  const comparisonSheet = comparisonSpreadsheet.getSheetByName(compareSheetName);
  if (!comparisonSheet) {
    throw new Error(`Comparison sheet not found: ${compareSheetName}`);
  }

  const clearColIdx = letterToCol(clearColumn);
  const compareColIdx = letterToCol(compareColumn);

  const targetLastRow = targetSheet.getLastRow();
  const comparisonLastRow = comparisonSheet.getLastRow();

  const targetRowsAvailable = Math.max(0, targetLastRow - firstRow + 1);
  const comparisonRowsAvailable = Math.max(0, comparisonLastRow - compareFirstRow + 1);
  const rowsToProcess = Math.min(targetRowsAvailable, comparisonRowsAvailable);

  if (rowsToProcess <= 0) {
    Logger.log(`No rows to compare for sheet "${targetSheetName}" starting at row ${firstRow}.`);
    return 0;
  }

  const clearValues = targetSheet.getRange(firstRow, clearColIdx, rowsToProcess, 1).getValues().flat();
  const compareValues = comparisonSheet.getRange(compareFirstRow, compareColIdx, rowsToProcess, 1).getValues().flat();

  let clearedCells = 0;

  for (let index = 0; index < rowsToProcess; index++) {
    const clearStamp = toComparableStamp(clearValues[index]);
    const compareStamp = toComparableStamp(compareValues[index]);

    if (clearStamp == null && compareStamp == null) {
      continue;
    }

    if (compareStamp == null && clearStamp != null) {
      targetSheet.getRange(firstRow + index, clearColIdx).clearContent();
      clearedCells++;
      continue;
    }

    if (clearStamp == null || compareStamp == null) {
      continue;
    }

    if (compareStamp > clearStamp) {
      targetSheet.getRange(firstRow + index, clearColIdx).clearContent();
      clearedCells++;
    }
  }

  Logger.log(
    `Cleared ${clearedCells} cell(s) in column ${clearColumn} on sheet "${targetSheetName}" because the comparison column was newer.`
  );

  return clearedCells;
}

// Example:
// function executeClearCellIfComparisonDateIsNewer() {
//   clearCellIfComparisonDateIsNewer({
//     sourceSpreadsheetId: 'TARGET_SPREADSHEET_ID',
//     sourceSheetName: 'Main',
//     sourceStartRow: 2,
//     clearColumn: 'O',
//     compareColumn: 'N',
//     compareSpreadsheetId: 'COMPARE_SPREADSHEET_ID',
//     compareSheetName: 'OutraAba',
//     compareStartRow: 2
//   });
// }