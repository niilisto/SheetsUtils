/**
 * Copies rows from a source sheet to a target sheet based on key matching.
 * If sourceKeyCol or targetKeyCol are empty (noKeyMode), rows are copied sequentially from source to target.
 * If a key from a source row is NOT found in the target sheet, the source row is copied to a new row in the target.
 * If a key from a source row IS found in the target sheet:
 *   - If 'overwrite' is true, the existing target row is updated with data from the source row.
 *   - If 'overwrite' is false, the existing target row is NOT modified.
 *
 * @param {Object}   config                   – Configuration object
 * @param {string}   config.sourceSpreadsheetId   – ID of the source spreadsheet.
 * @param {string}   config.sourceSheetName       – Name of the source sheet/tab.
 * @param {number}   config.sourceStartRow        – 1-based row number to start reading in source.
 * @param {string}   config.sourceKeyCol          – Column letter(s) of the key in source (e.g., "B" or "B,C").
 * @param {string}   config.targetSpreadsheetId   – ID of the target spreadsheet.
 * @param {string}   config.targetSheetName       – Name of the target sheet/tab.
 * @param {number}   config.targetStartRow        – 1-based row number to start writing/checking in target.
 * @param {string}   config.targetKeyCol          – Column letter(s) of the key in target.
 * @param {Object}   config.colsMap              – Column mapping {sourceColLetterOrFixedValue: targetColLetter}.
 * @param {Object}   [config.inputFil]           – (optional) Filters for source data {sourceColLetter: "value" or "!value" or ["A","B"]}.
 * @param {Object}   [config.outputFil]          – (optional) Filters for target data if overwriting {targetColLetter: "value" or "!value" or ["A","B"]}.
 * @param {boolean}  config.overwrite            – If true and key matches, overwrite existing target row; otherwise, skip.
 * @param {string[]} [config.timestampCols]      – (optional) Array of target column letters to receive a timestamp.
 * @param {string}   [config.timestampLocale]    – (optional) Timezone for timestamp. Required if timestampCols is used.
 * @param {string}   [config.timestampFormat]    – (optional) Format for timestamp. Required if timestampCols is used.
 */
function copyByKey(config) {
  const {
    sourceSpreadsheetId, sourceSheetName, sourceStartRow, sourceKeyCol = [],
    targetSpreadsheetId, targetSheetName, targetStartRow, targetKeyCol = [],
    colsMap, overwrite = false, inputFil = {}, outputFil = {},
    timestampCols = [], timestampLocale, timestampFormat
  } = config;

  const letterToCol = l => {
    const L = String(l).toUpperCase();
    if (!/^[A-Z]+$/.test(L)) {
      throw new Error(`Invalid column letter format: "${l}". Must consist of A-Z characters only.`);
    }
    if (L.length > 3) {
      throw new Error(`Column string "${l}" is too long (${L.length} chars) to be a standard column letter (max 3 chars). Interpreting as fixed value.`);
    }
    return L.split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);
  };

  const sourceKeyColsStr = String(sourceKeyCol);
  const targetKeyColsStr = String(targetKeyCol);

  const srcKeyIdxs = sourceKeyColsStr && sourceKeyColsStr.trim() !== '' ? sourceKeyColsStr.split(/\s*,\s*/).map(l => letterToCol(l) - 1) : [];
  const tgtKeyIdxs = targetKeyColsStr && targetKeyColsStr.trim() !== '' ? targetKeyColsStr.split(/\s*,\s*/).map(l => letterToCol(l) - 1) : [];

  const noKeyMode = srcKeyIdxs.length === 0 || tgtKeyIdxs.length === 0;

  const srcSheet = SpreadsheetApp.openById(sourceSpreadsheetId).getSheetByName(sourceSheetName);
  const tgtSheet = SpreadsheetApp.openById(targetSpreadsheetId).getSheetByName(targetSheetName);

  if (!srcSheet) throw new Error(`Source sheet "${sourceSheetName}" not found in spreadsheet ID "${sourceSpreadsheetId}".`);
  if (!tgtSheet) throw new Error(`Target sheet "${targetSheetName}" not found in spreadsheet ID "${targetSpreadsheetId}".`);

  const processedMappings = [];
  const sourceDefs = Object.keys(colsMap);
  const targetLetters = Object.values(colsMap);
  let maxTargetColMappedIdx = 0;

  for (let i = 0; i < sourceDefs.length; i++) {
    const srcDef = sourceDefs[i];
    const tgtLet = targetLetters[i];
    let tgtIdx;
    try {
      tgtIdx = letterToCol(tgtLet) - 1;
    } catch (e) {
      throw new Error(`Invalid target column letter in colsMap for source "${srcDef}": "${tgtLet}". ${e.message}`);
    }
    if (tgtIdx > maxTargetColMappedIdx) maxTargetColMappedIdx = tgtIdx;

    const map = { targetColIdx: tgtIdx, isFixedValue: false, sourceColIdx: -1, fixedValue: null };
    try {
      map.sourceColIdx = letterToCol(srcDef) - 1;
    } catch (e) {
      map.isFixedValue = true;
      map.fixedValue = srcDef;
    }
    processedMappings.push(map);
  }

  const inputMap = Object.fromEntries(
    Object.entries(inputFil).map(([l, f]) => [letterToCol(l) - 1, f])
  );
  const outputMap = Object.fromEntries(
    Object.entries(outputFil).map(([l, f]) => [letterToCol(l) - 1, f])
  );

  const checkFilter = (rowValues, filterMap) => {
    return Object.entries(filterMap).every(([colIdx, condition]) => {
      let val = rowValues[parseInt(colIdx)]; // Ensure colIdx is number
      // Normalize value for comparison
      if (val instanceof Date) {
        val = val.toISOString();
      } else if (val !== null && val !== undefined) {
        val = String(val).trim();
      }
      
      if (Array.isArray(condition)) {
        return condition.some(c => {
          const normalizedCondition = c instanceof Date ? c.toISOString() : String(c).trim();
          return val === normalizedCondition;
        });
      } else if (typeof condition === 'string' && condition.startsWith('!')) {
        const normalizedCondition = condition.slice(1).trim();
        return val !== normalizedCondition;
      } else {
        const normalizedCondition = condition instanceof Date ? condition.toISOString() : String(condition).trim();
        return val === normalizedCondition;
      }
    });
  };

  const srcLastRow = srcSheet.getLastRow();
  const srcLastCol = srcSheet.getLastColumn();
  const numSrcRowsToRead = srcLastRow < sourceStartRow ? 0 : srcLastRow - sourceStartRow + 1;
  const srcData = numSrcRowsToRead > 0
    ? srcSheet.getRange(sourceStartRow, 1, numSrcRowsToRead, srcLastCol).getValues()
    : [];

  const tgtLastRow = tgtSheet.getLastRow();
  const tgtLastCol = tgtSheet.getMaxColumns(); // Use MaxColumns for sparse data potential
  const numTgtRowsToRead = tgtLastRow < targetStartRow ? 0 : tgtLastRow - targetStartRow + 1;
  const tgtData = (numTgtRowsToRead > 0 && !noKeyMode) // Only read target data if in key mode
    ? tgtSheet.getRange(targetStartRow, 1, numTgtRowsToRead, tgtLastCol).getValues()
    : [];

  const normalizeKeyValue = (val) => {
    if (val instanceof Date) {
      // Convert Date objects to a consistent string format
      // Truncate milliseconds to avoid precision issues
      const d = new Date(val);
      d.setMilliseconds(0);
      return d.toISOString();
    }
    if (val === null || val === undefined) {
      return '';
    }
    
    // Convert to string and trim whitespace for consistency
    const str = String(val).trim();
    
    // Try to parse date strings in various formats (dd/MM/yyyy HH:mm:ss, etc.)
    // Common Brazilian format: dd/MM/yyyy HH:mm:ss
    const brDatePattern = /^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/;
    const brMatch = str.match(brDatePattern);
    if (brMatch) {
      const [, day, month, year, hour, minute, second] = brMatch;
      // Create Date object (month is 0-indexed in JavaScript)
      const d = new Date(parseInt(year), parseInt(month) - 1, parseInt(day), 
                         parseInt(hour), parseInt(minute), parseInt(second), 0);
      if (!isNaN(d.getTime())) {
        return d.toISOString();
      }
    }
    
    // Try to parse ISO-like strings or other date formats
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime()) && str.match(/\d{4}[-\/]\d{2}[-\/]\d{2}|\d{2}[-\/]\d{2}[-\/]\d{4}/)) {
      parsed.setMilliseconds(0);
      return parsed.toISOString();
    }
    
    // Remove any non-breaking spaces or other whitespace characters
    return str.replace(/\s+/g, ' ').trim();
  };

  const makeKey = (row, idxs) => idxs.map(i => {
    const val = (row && i < row.length) ? row[i] : '';
    return normalizeKeyValue(val);
  }).join('|');

  const rowsToAppend = [];
  const cellUpdatesForOverwrite = []; // {row: sheetRowNum, col: sheetColNum, value: any}
  const affectedTargetRowsForTimestamp = new Set(); // 1-based sheet row indices

  // Filter duplicates in source data before processing -- Added in: 07/30/2025
  let filteredSrcData = srcData;
  if (!noKeyMode) {
    const seenKeys = new Set();
    filteredSrcData = srcData.filter(srcRow => {
      const srcKey = makeKey(srcRow, srcKeyIdxs);
      if (!srcKey || srcKey.split('|').every(kPart => kPart === "")) return false;
      if (seenKeys.has(srcKey)) {
        console.warn(`[copyByKey] Chave duplicada encontrada no source, ignorando: "${srcKey}"`);
        return false; // já vimos essa chave, pula
      }
      seenKeys.add(srcKey);
      return true;
    });
    console.log(`[copyByKey] ${filteredSrcData.length} linhas únicas no source após filtrar duplicatas (de ${srcData.length} total)`);
    if (filteredSrcData.length > 0) {
      // Log first key as example
      const firstKey = makeKey(filteredSrcData[0], srcKeyIdxs);
      console.log(`[copyByKey] Exemplo de chave source: "${firstKey}"`);
    }
  }

  if (noKeyMode) {
    let currentTargetRowForAppending = Math.max(targetStartRow, tgtSheet.getLastRow() + 1);
    if (tgtSheet.getLastRow() === 0 && targetStartRow > 1) { // Target is empty, start at targetStartRow
      currentTargetRowForAppending = targetStartRow;
    }

    srcData.forEach(srcRow => {
      if (!checkFilter(srcRow, inputMap)) return;

      const newTargetRowData = new Array(maxTargetColMappedIdx + 1).fill("");
      processedMappings.forEach(m => {
        newTargetRowData[m.targetColIdx] = m.isFixedValue ? m.fixedValue : (m.sourceColIdx < srcRow.length ? srcRow[m.sourceColIdx] : "");
      });
      rowsToAppend.push(newTargetRowData);
      affectedTargetRowsForTimestamp.add(currentTargetRowForAppending + rowsToAppend.length - 1);
    });

  } else { // Key-matching mode
    const tgtKeyMap = new Map(); // Map<string, {rowIndexInTgtDataArray: number, originalSheetRowIndex: number, rowData: any[]}>
    tgtData.forEach((tgtRow, idx) => {
      const key = makeKey(tgtRow, tgtKeyIdxs);
      if (key && !key.split('|').every(kPart => kPart === "")) { // Ensure key is not empty and not just separators
        if (tgtKeyMap.has(key)) {
          console.warn(`[copyByKey] Chave duplicada encontrada no target na linha ${targetStartRow + idx}: "${key}"`);
        }
        tgtKeyMap.set(key, {
          rowIndexInTgtDataArray: idx,
          originalSheetRowIndex: targetStartRow + idx,
          rowData: tgtRow
        });
      }
    });
    
    console.log(`[copyByKey] Target key map criado com ${tgtKeyMap.size} chaves únicas`);
    if (tgtKeyMap.size > 0) {
      // Log first key as example
      const firstTargetKey = Array.from(tgtKeyMap.keys())[0];
      console.log(`[copyByKey] Exemplo de chave target: "${firstTargetKey}"`);
    }

    let nextAppendRowIndex = Math.max(targetStartRow, tgtSheet.getLastRow() + 1);
    if (tgtSheet.getLastRow() === 0 && targetStartRow > 1) {
      nextAppendRowIndex = targetStartRow;
    }


    filteredSrcData.forEach(srcRow => {
      if (!checkFilter(srcRow, inputMap)) return;

      const srcKey = makeKey(srcRow, srcKeyIdxs);
      if (!srcKey || srcKey.split('|').every(kPart => kPart === "")) return; // Skip if source key is blank

      const existingTargetMatch = tgtKeyMap.get(srcKey);

      if (existingTargetMatch) { // Key found in target
        console.log(`[copyByKey] Chave "${srcKey}" encontrada no target - verificando overwrite`);
        if (overwrite) {
          if (!checkFilter(existingTargetMatch.rowData, outputMap)) return;

          let rowWasModified = false;
          processedMappings.forEach(m => {
            const srcVal = m.isFixedValue ? m.fixedValue : (m.sourceColIdx < srcRow.length ? srcRow[m.sourceColIdx] : "");
            const existingVal = (m.targetColIdx < existingTargetMatch.rowData.length) ? existingTargetMatch.rowData[m.targetColIdx] : "";

            if (String(existingVal) !== String(srcVal)) {
              cellUpdatesForOverwrite.push({
                row: existingTargetMatch.originalSheetRowIndex,
                col: m.targetColIdx + 1,
                value: srcVal
              });
              rowWasModified = true;
            }
          });
          if (rowWasModified) {
            affectedTargetRowsForTimestamp.add(existingTargetMatch.originalSheetRowIndex);
          }
        }
        // else: key found, overwrite is false, so do nothing
      } else { // Key NOT found in target - prepare to append
        console.log(`[copyByKey] Chave "${srcKey}" NÃO encontrada no target - preparando para adicionar`);
        const newTargetRowData = new Array(maxTargetColMappedIdx + 1).fill("");
        processedMappings.forEach(m => {
          newTargetRowData[m.targetColIdx] = m.isFixedValue ? m.fixedValue : (m.sourceColIdx < srcRow.length ? srcRow[m.sourceColIdx] : "");
        });
        rowsToAppend.push(newTargetRowData);
        const newRowIndex = nextAppendRowIndex + rowsToAppend.length - 1;
        affectedTargetRowsForTimestamp.add(newRowIndex);
        
        // CRITICAL FIX: Add this key to tgtKeyMap so it's not added again in the same execution
        tgtKeyMap.set(srcKey, {
          rowIndexInTgtDataArray: -1, // Not in existing data
          originalSheetRowIndex: newRowIndex,
          rowData: newTargetRowData
        });
      }
    });
  }

  // --- Apply changes ---
  let appendedRowsCount = 0;
  if (rowsToAppend.length > 0) {
    let firstAppendRowOnSheet = Math.max(targetStartRow, tgtSheet.getLastRow() + 1);
    if (tgtSheet.getLastRow() === 0 && targetStartRow > 1) { // Handle empty sheet with specific start row
      firstAppendRowOnSheet = targetStartRow;
    }

    const numColsToAppend = rowsToAppend.reduce((max, row) => Math.max(max, row.length), 0);
    if (numColsToAppend > 0) {
      if (firstAppendRowOnSheet + rowsToAppend.length - 1 > tgtSheet.getMaxRows()) {
        tgtSheet.insertRowsAfter(tgtSheet.getMaxRows(), (firstAppendRowOnSheet + rowsToAppend.length - 1) - tgtSheet.getMaxRows());
      }
      if (numColsToAppend > tgtSheet.getMaxColumns()) {
        tgtSheet.insertColumnsAfter(tgtSheet.getMaxColumns(), numColsToAppend - tgtSheet.getMaxColumns());
      }
      tgtSheet.getRange(firstAppendRowOnSheet, 1, rowsToAppend.length, numColsToAppend).setValues(rowsToAppend);
      appendedRowsCount = rowsToAppend.length;
    }
  }

  // Consolidate all cell updates (overwrites + timestamps)
  const finalCellUpdates = [...cellUpdatesForOverwrite];

  if (affectedTargetRowsForTimestamp.size > 0 && timestampCols && timestampCols.length > 0) {
    if (!timestampLocale || !timestampFormat) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Carimbo de data/hora solicitado, mas locale ou formato ausentes. Ignorando carimbo.', '⚠️ Atenção', 5);
      console.warn('Timestamping requested but locale or format missing. Skipping.');
    } else {
      let formattedTimestamp;
      try {
        formattedTimestamp = Utilities.formatDate(new Date(), timestampLocale, timestampFormat);
      } catch (e) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Erro ao formatar carimbo: ${e.message}. Ignorando.`, '❌ Erro no Carimbo', 5);
        console.error(`Error formatting timestamp: ${e.message}. Skipping.`);
      }

      if (formattedTimestamp !== undefined) {
        const timestampTargetColIndices = [];
        for (const colLetter of timestampCols) {
          try {
            timestampTargetColIndices.push(letterToCol(colLetter) - 1);
          } catch (e) {
            console.warn(`Invalid column letter "${colLetter}" in timestampCols. Skipping this column. Error: ${e.message}`);
          }
        }

        affectedTargetRowsForTimestamp.forEach(rowNum => {
          timestampTargetColIndices.forEach(targetColIdx => {
            // Ensure timestamp column is within sheet bounds if writing to it
            if (targetColIdx + 1 > tgtSheet.getMaxColumns()) {
              tgtSheet.insertColumnsAfter(tgtSheet.getMaxColumns(), (targetColIdx + 1) - tgtSheet.getMaxColumns());
            }
            finalCellUpdates.push({ row: rowNum, col: targetColIdx + 1, value: formattedTimestamp });
          });
        });
      }
    }
  }

  let updatedCellsCount = 0;
  if (finalCellUpdates.length > 0) {
    // Sort updates by row then column to potentially improve efficiency for `setValues`, though here we do cell by cell.
    // For cell-by-cell, order doesn't matter much for performance, but can be good for logical consistency if debugging.
    finalCellUpdates.sort((a, b) => a.row - b.row || a.col - b.col);

    const uniqueUpdates = [];
    const seenCells = new Set();
    // Prioritize non-timestamp data if a cell is targeted by both data and timestamp
    // Iterate in reverse if fixed values/data should overwrite timestamps.
    // Or, ensure timestamp columns are distinct from data columns in colsMap.
    // Current logic: last update to a cell wins if multiple updates target it.
    // To give data precedence over timestamps for the same cell:
    // One way is to apply data updates first, then only apply timestamps to cells not already written.
    // For simplicity, assuming timestampCols are usually dedicated or it's acceptable for timestamp to overwrite if mapped.

    finalCellUpdates.forEach(u => {
      const cellKey = `${u.row},${u.col}`;
      if (!seenCells.has(cellKey)) { // Process only the last intended update for a cell if duplicates exist
        // (This simple Set approach takes the first one encountered if sorted, or last if iterated in reverse)
        // A more robust way if specific precedence is needed: group by cell, then pick.
        // For now, let's assume the order or uniqueness is handled by logic generating finalCellUpdates.
      }
      // Ensure row for update is within sheet bounds
      if (u.row > tgtSheet.getMaxRows()) {
        tgtSheet.insertRowsAfter(tgtSheet.getMaxRows(), u.row - tgtSheet.getMaxRows());
      }
      if (u.col > tgtSheet.getMaxColumns()) { // Ensure column exists
        tgtSheet.insertColumnsAfter(tgtSheet.getMaxColumns(), u.col - tgtSheet.getMaxColumns());
      }

      tgtSheet.getRange(u.row, u.col).setValue(u.value);
      updatedCellsCount++; // Counts each setValue call
    });
  }

  let msg = "";
  if (appendedRowsCount > 0) {
    msg += `${appendedRowsCount} linha(s) copiada(s) para "${targetSheetName}". `;
  }
  const overwriteModificationsCount = cellUpdatesForOverwrite.length;
  if (overwriteModificationsCount > 0) {
    msg += `${overwriteModificationsCount} célula(s) de dados atualizada(s) em linhas existentes em "${targetSheetName}". `;
  }
  const timestampsAppliedCount = finalCellUpdates.length - overwriteModificationsCount;
  if (timestampsAppliedCount > 0 && affectedTargetRowsForTimestamp.size > 0) {
    msg += `Carimbos de data/hora aplicados ${timestampsAppliedCount > 1 ? timestampsAppliedCount + ' vezes' : ''} em ${affectedTargetRowsForTimestamp.size} linha(s) afetada(s).`;
  }

  if (!msg.trim()) {
    msg = 'Nenhuma linha copiada ou célula modificada.';
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(msg.trim(), '✅ Resultado', 10);
  console.log(msg.trim());
}