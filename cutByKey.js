/**
 * @name cutByKey
 * @version 3 (last updated: 2025-05-03)
 * Transfers rows from a source sheet to a target sheet if the source key does NOT exist in the target,
 * and then deletes the transferred rows from the source sheet.
 * If sourceKeyCol or targetKeyCol are empty (noKeyMode), rows are copied sequentially and then deleted.
 * If a key from a source row ALREADY EXISTS in the target sheet, the function exits immediately without processing further.
 *
 * @param {Object}   config                       – Configuration object
 * @param {string}   config.sourceSpreadsheetId   – ID of the source spreadsheet.
 * @param {string}   config.sourceSheetName       – Name of the source sheet.
 * @param {number}   config.sourceStartRow        – Row number (1-based) to start reading from in the source.
 * @param {string[]} config.sourceKeyCol          – Letter(s) of the key column(s) in the source (e.g., "B" or ["B","C"]).
 * @param {string}   config.targetSpreadsheetId   – ID of the target spreadsheet.
 * @param {string}   config.targetSheetName       – Name of the target sheet.
 * @param {number}   config.targetStartRow        – Row number (1-based) to start writing/checking in the target.
 * @param {string[]} config.targetKeyCol          – Letter(s) of the key column(s) in the target.
 * @param {Object}   config.colsMap               – Column mapping {sourceColOrFixedValue: targetCol}. 
 *                                               – Se srcDef for uma string que não corresponda a qualquer coluna existente no sheet de origem,
 *                                                 ela será tratada como valor fixo.
 * @param {Object}   [config.inputFil]            – (optional) Filters for source data {sourceCol: "value" or "!value" or ["A","B"]}.
 * @param {string[]} [config.timestampCols]       – (optional) Array of target column letters to receive a timestamp for NEW rows.
 * @param {string}   [config.timestampLocale]     – (optional) Timezone for timestamp. Required if timestampCols is used.
 * @param {string}   [config.timestampFormat]     – (optional) Format for timestamp. Required if timestampCols is used.
 */
function cutByKey(config) {
  const {
    sourceSpreadsheetId, sourceSheetName, sourceStartRow, sourceKeyCol = [],
    targetSpreadsheetId, targetSheetName, targetStartRow, targetKeyCol = [],
    colsMap, inputFil = {},
    timestampCols = [], timestampLocale, timestampFormat
  } = config;

  const letterToCol = l => {
    const L = String(l).toUpperCase();
    if (!/^[A-Z]+$/.test(L)) throw new Error(`Invalid column letter: "${l}".`);
    return L.split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);
  };

  // Helpers
  const srcKeyIdxs = Array.isArray(sourceKeyCol)
    ? sourceKeyCol.map(l => letterToCol(l) - 1)
    : (sourceKeyCol ? [letterToCol(sourceKeyCol) - 1] : []);
  const tgtKeyIdxs = Array.isArray(targetKeyCol)
    ? targetKeyCol.map(l => letterToCol(l) - 1)
    : (targetKeyCol ? [letterToCol(targetKeyCol) - 1] : []);
  const noKeyMode = srcKeyIdxs.length === 0 || tgtKeyIdxs.length === 0;

  const srcSheet = SpreadsheetApp.openById(sourceSpreadsheetId).getSheetByName(sourceSheetName);
  const tgtSheet = SpreadsheetApp.openById(targetSpreadsheetId).getSheetByName(targetSheetName);
  if (!srcSheet || !tgtSheet) throw new Error('Sheet not found');

  // Read source data
  const srcLastRow = srcSheet.getLastRow();
  const numSrc = Math.max(0, srcLastRow - sourceStartRow + 1);
  const srcData = numSrc
    ? srcSheet.getRange(sourceStartRow, 1, numSrc, srcSheet.getLastColumn()).getValues()
    : [];

  // Prepare mappings
  const processedMappings = Object.entries(colsMap).map(([srcDef, tgtLet]) => {
    const targetColIdx = letterToCol(tgtLet) - 1;
    let map = { targetColIdx, isFixedValue: false, sourceColIdx: -1, fixedValue: null };
    try {
      const idx = letterToCol(srcDef) - 1;
      if (idx < srcSheet.getLastColumn()) {
        map.sourceColIdx = idx;
      } else {
        map.isFixedValue = true;
        map.fixedValue = srcDef;
      }
    } catch {
      map.isFixedValue = true;
      map.fixedValue = srcDef;
    }
    return map;
  });

  // Filters
  const inputMap = Object.fromEntries(
    Object.entries(inputFil).map(([l, f]) => [letterToCol(l) - 1, f])
  );
  const checkFilter = row =>
    Object.entries(inputMap).every(([i, cond]) => {
      const val = row[i];
      if (Array.isArray(cond)) return cond.includes(val);
      if (typeof cond === 'string' && cond.startsWith('!')) {
        return cond === '!' ? val != null && val !== '' : val != cond.slice(1);
      }
      return val == cond;
    });

  // Target key map
  let tgtKeyMap = new Map();
  if (!noKeyMode) {
    const tgtLastRow = tgtSheet.getLastRow();
    const numTgt = Math.max(0, tgtLastRow - targetStartRow + 1);
    const tgtData = numTgt
      ? tgtSheet.getRange(targetStartRow, 1, numTgt, tgtSheet.getLastColumn()).getValues()
      : [];
    tgtData.forEach(r => {
      const key = tgtKeyIdxs.map(j => r[j] || '').join('|');
      if (key.split('|').some(k => k !== '')) tgtKeyMap.set(key, true);
    });
  }

  // Timestamp setup
  let formattedTimestamp;
  if (timestampCols.length) {
    if (!timestampLocale || !timestampFormat) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Missing locale/format for timestamp', '⚠️', 5);
    } else {
      formattedTimestamp = Utilities.formatDate(new Date(), timestampLocale, timestampFormat);
    }
  }
  const timestampIdxs = formattedTimestamp
    ? timestampCols.map(l => letterToCol(l) - 1)
    : [];

  // Batch arrays
  const rowsToAppend = [];
  const rowsToDelete = [];

  // Collect rows
  srcData.forEach((srcRow, idx) => {
    const sheetRow = sourceStartRow + idx;
    if (!checkFilter(srcRow)) return;
    const key = srcKeyIdxs.map(j => srcRow[j] || '').join('|');
    if (!noKeyMode) {
      if (!key || key.split('|').every(k => k === '')) return;
      if (tgtKeyMap.has(key)) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Key "${key}" exists -> abort`, 'ℹ️', 10);
        return;
      }
      tgtKeyMap.set(key, true);
    }
    const rowData = Array(
      Math.max(...processedMappings.map(m => m.targetColIdx)) + 1
    ).fill('');
    processedMappings.forEach(m => {
      rowData[m.targetColIdx] = m.isFixedValue
        ? m.fixedValue
        : (m.sourceColIdx < srcRow.length ? srcRow[m.sourceColIdx] : '');
    });
    rowsToAppend.push(rowData);
    rowsToDelete.push(sheetRow);
  });

  // Append in batch
  if (rowsToAppend.length) {
    const firstRow = Math.max(targetStartRow, tgtSheet.getLastRow() + 1);
    if (firstRow + rowsToAppend.length - 1 > tgtSheet.getMaxRows()) {
      tgtSheet.insertRowsAfter(
        tgtSheet.getMaxRows(),
        firstRow + rowsToAppend.length - 1 - tgtSheet.getMaxRows()
      );
    }
    const maxCols = rowsToAppend.reduce((mx, r) => Math.max(mx, r.length), 0);
    if (maxCols > tgtSheet.getMaxColumns()) {
      tgtSheet.insertColumnsAfter(
        tgtSheet.getMaxColumns(),
        maxCols - tgtSheet.getMaxColumns()
      );
    }
    tgtSheet
      .getRange(firstRow, 1, rowsToAppend.length, maxCols)
      .setValues(rowsToAppend);

    if (formattedTimestamp) {
      timestampIdxs.forEach(colIdx => {
        const tsVals = rowsToAppend.map(() => [formattedTimestamp]);
        if (colIdx + 1 > tgtSheet.getMaxColumns()) {
          tgtSheet.insertColumnsAfter(
            tgtSheet.getMaxColumns(),
            colIdx + 1 - tgtSheet.getMaxColumns()
          );
        }
        tgtSheet
          .getRange(firstRow, colIdx + 1, tsVals.length, 1)
          .setValues(tsVals);
      });
    }
  }

  // Delete rows in contiguous batches
  if (rowsToDelete.length) {
    rowsToDelete.sort((a, b) => a - b);
    const deleteRanges = [];
    let start = rowsToDelete[0], count = 1;
    for (let i = 1; i < rowsToDelete.length; i++) {
      if (rowsToDelete[i] === rowsToDelete[i - 1] + 1) {
        count++;
      } else {
        deleteRanges.push({ start, count });
        start = rowsToDelete[i]; count = 1;
      }
    }
    deleteRanges.push({ start, count });
    // Delete from bottom to top
    deleteRanges.reverse().forEach(({ start, count }) => {
      try {
        srcSheet.deleteRows(start, count);
      } catch (e) {
        console.error(`Failed to delete rows at ${start} count ${count}: ${e}`);
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Error deleting rows ${start}-${start + count - 1}`,
          '⚠️ Deletion Error',
          5
        );
      }
    });
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${rowsToAppend.length} rows moved, ${rowsToDelete.length} deleted.`,
    '✅',
    10
  );
}