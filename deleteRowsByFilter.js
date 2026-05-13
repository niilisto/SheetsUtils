/**
 * @name deleteRowsByFilter
 * @version 2 (last updated: YYYY-MM-DD - Batch Deletion Implemented)
 * Deletes rows from a specified sheet if they match the given filter criteria,
 * using batch deletion for improved performance.
 *
 * @param {Object}   config                       – Configuration object
 * @param {string}   config.spreadsheetId         – ID of the spreadsheet.
 * @param {string}   config.sheetName             – Name of the sheet to process.
 * @param {number}   [config.startRow=2]          – Row number (1-based) to start scanning for deletions.
 * @param {Object}   config.filter                – Filter criteria {columnLetter: "value" or "!value" or ["A","B"] or "!" (not blank)}.
 *                                                  Rows matching ALL criteria will be deleted.
 *                                                  If empty or undefined, no rows will be deleted.
 */
function deleteRowsByFilter(config) {
  const {
    spreadsheetId,
    sheetName,
    startRow = 1, // Default to row 1 if not specified
    filter = {}
  } = config;

  if (Object.keys(filter).length === 0) {
    const msg = `No filter criteria provided for sheet "${sheetName}". No rows will be deleted.`;
    console.log(msg);
    SpreadsheetApp.getActiveSpreadsheet()?.toast(msg, 'ℹ️ Info', 5);
    return;
  }

  // Helper function to convert column letter to 0-based index
  const letterToCol = l => {
    const L = String(l).toUpperCase();
    if (!/^[A-Z]+$/.test(L)) {
      throw new Error(`Invalid column letter format: "${l}". Must consist only of A-Z characters.`);
    }
    if (L.length > 3) {
      throw new Error(`Column string "${l}" is too long (${L.length} characters) to be a standard column letter (max 3 characters).`);
    }
    return L.split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);
  };

  let ss;
  try {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    throw new Error(`Could not open spreadsheet with ID "${spreadsheetId}". Error: ${e.message}`);
  }
  
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found in spreadsheet ID "${spreadsheetId}".`);
  }

  // Process filter keys (column letters) into 0-based column indices
  const processedFilter = {};
  for (const colLetter in filter) {
    if (Object.prototype.hasOwnProperty.call(filter, colLetter)) {
      try {
        const colIdx = letterToCol(colLetter) - 1;
        processedFilter[colIdx] = filter[colLetter];
      } catch (e) {
        throw new Error(`Invalid column letter "${colLetter}" in filter: ${e.message}`);
      }
    }
  }

  // Helper function to check if a row matches the filter criteria
  const checkFilter = (rowValues, filterMap) => {
    if (Object.keys(filterMap).length === 0) return false;
    return Object.entries(filterMap).every(([colIdxStr, condition]) => {
      const colIdx = parseInt(colIdxStr);
      const val = colIdx < rowValues.length ? rowValues[colIdx] : "";

      if (Array.isArray(condition)) {
        return condition.includes(val);
      } else if (typeof condition === 'string' && condition.startsWith('!')) {
        if (condition === '!') return val !== "" && val !== null && val !== undefined;
        return String(val) != condition.slice(1);
      }
      return String(val) == String(condition);
    });
  };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < startRow) {
    const msg = `No data to process in sheet "${sheetName}" starting from row ${startRow}. Last row is ${lastRow}.`;
    console.log(msg);
    SpreadsheetApp.getActiveSpreadsheet()?.toast(msg, 'ℹ️ Info', 5);
    return;
  }
  
  if (lastCol === 0 && lastRow >= startRow) {
    const msg = `Sheet "${sheetName}" has rows but no data in columns to filter by. No rows deleted.`;
    console.log(msg);
    SpreadsheetApp.getActiveSpreadsheet()?.toast(msg, 'ℹ️ Info', 5);
    return;
  }

  const numRowsToRead = lastRow - startRow + 1;
  const dataRange = sheet.getRange(startRow, 1, numRowsToRead, Math.max(1, lastCol));
  const dataValues = dataRange.getValues();

  const rowIndicesToDelete = [];

  dataValues.forEach((row, index) => {
    if (checkFilter(row, processedFilter)) {
      const originalSheetRowIndex = startRow + index;
      rowIndicesToDelete.push(originalSheetRowIndex);
    }
  });

  // --- DELETE ROWS FROM SOURCE (Batch approach) ---
  let deletedRowsCount = 0;
  if (rowIndicesToDelete.length > 0) {
    // Sort in descending order. This is crucial for the batching logic below,
    // as we'll iterate through the sorted indices to find contiguous blocks
    // and delete them from the highest row number downwards.
    rowIndicesToDelete.sort((a, b) => b - a); // e.g., [10, 8, 7, 5, 2]

    let i = 0;
    while (i < rowIndicesToDelete.length) {
      // The current row index (rowIndicesToDelete[i]) is the highest row number
      // in the contiguous block we are about to identify.
      const lastRowInPotentialBatch = rowIndicesToDelete[i]; 
      let countInBatch = 1;

      // Look ahead in the sorted list to find how many subsequent rows are contiguous.
      // Since the list is sorted descending, a contiguous row will be `lastRowInPotentialBatch - countInBatch`.
      let k = i + 1;
      while (k < rowIndicesToDelete.length && rowIndicesToDelete[k] === lastRowInPotentialBatch - countInBatch) {
        countInBatch++;
        k++;
      }
      
      // The first row (lowest row number) in this contiguous block to be deleted.
      const firstRowInBatch = lastRowInPotentialBatch - countInBatch + 1;

      try {
        // console.log(`Batch Deleting: ${countInBatch} rows starting from sheet row ${firstRowInBatch} (original last in batch: ${lastRowInPotentialBatch})`);
        sheet.deleteRows(firstRowInBatch, countInBatch);
        deletedRowsCount += countInBatch;
      } catch (e) {
        console.error(`Failed to delete batch of ${countInBatch} rows starting at sheet row ${firstRowInBatch} from sheet "${sheetName}": ${e.message}`);
        SpreadsheetApp.getActiveSpreadsheet()?.toast(`Error deleting batch at row ${firstRowInBatch}. See logs.`, '⚠️ Deletion Error', 5);
        // Optional: Decide on error strategy. For now, log and attempt next batch.
        // If one batch fails, subsequent row indices are still correct relative to the *current* state of the sheet
        // because we process from highest row numbers downwards.
      }
      
      // Move index 'i' past the rows that were just processed (or attempted) in this batch.
      i = k; 
    }
  }

  let resultMsg;
  if (deletedRowsCount > 0) {
    resultMsg = `${deletedRowsCount} row(s) deleted from sheet "${sheetName}".`;
  } else {
    if (rowIndicesToDelete.length > 0 && deletedRowsCount === 0) {
      // This case implies rows were identified for deletion but errors occurred during all batch deletions.
      resultMsg = `Identified ${rowIndicesToDelete.length} row(s) for deletion, but errors occurred during the process. Check logs.`;
    } else {
      resultMsg = `No rows matched the filter criteria in sheet "${sheetName}". No rows deleted.`;
    }
  }
  console.log(resultMsg);
  SpreadsheetApp.getActiveSpreadsheet()?.toast(resultMsg, '✅ Deletion Result', 10);
}