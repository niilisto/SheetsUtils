/**
 * Transfers data from one sheet to another by matching one or more key-comparison groups.
 * When multiple target rows match, the row with the highest sheet row number wins,
 * but only among rows that also pass outputFil.
 *
 * @param {Object} config                        - Transfer configuration.
 * @param {string} config.sourceSpreadsheetId     - ID of the source spreadsheet.
 * @param {string} config.sourceSheetName         - Name of the source sheet/tab.
 * @param {number} config.sourceStartRow          - 1-based row number to start reading in source.
 * @param {string|string[]|Array<string[]|string>} config.sourceKeyCol - Single key group or array of key groups.
 * @param {string} config.targetSpreadsheetId     - ID of the target spreadsheet.
 * @param {string} config.targetSheetName         - Name of the target sheet/tab.
 * @param {number} config.targetStartRow          - 1-based row number to start writing in target.
 * @param {string|string[]|Array<string[]|string>} config.targetKeyCol - Single key group or array of key groups.
 * @param {Object} config.colsMap                 - Column mapping {sourceColOrFixedValue: targetCol}.
 * @param {Object} [config.inputFil]              - Optional filters for source data.
 * @param {Object} [config.outputFil]             - Optional filters for target data before overwrite.
 * @param {boolean} config.overwrite              - If true, always overwrite; if false, only fill empty cells.
 * @param {string[]} [config.timestampCols]       - Optional array of target column letters to receive a timestamp.
 * @param {string} [config.timestampLocale]       - Timezone for timestamp.
 * @param {string} [config.timestampFormat]       - Format for timestamp.
 */
function updateByKeyMulti(config) {
  const {
    sourceSpreadsheetId, sourceSheetName, sourceStartRow,
    sourceKeyCol = [], sourceKeyCols,
    targetSpreadsheetId, targetSheetName, targetStartRow,
    targetKeyCol = [], targetKeyCols,
    colsMap, overwrite, inputFil = {}, outputFil = {},
    timestampCols = [], timestampLocale, timestampFormat
  } = config;

  const letterToCol = l => {
    const L = String(l).toUpperCase();
    if (!/^[A-Z]+$/.test(L)) {
      throw new Error(`Invalid column letter format: "${l}". Must consist of A-Z characters only.`);
    }
    if (L.length > 3) {
      throw new Error(`Column string "${l}" is too long (${L.length} chars) to be a standard column letter (max 3 chars).`);
    }
    return L.split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);
  };

  const parseKeyDefinition = definition => {
    if (definition == null || definition === '') return [];
    if (Array.isArray(definition)) {
      return definition.flatMap(part => parseKeyDefinition(part));
    }
    return String(definition)
      .split(/\s*,\s*/)
      .map(part => part.trim())
      .filter(Boolean);
  };

  const normalizeKeyGroups = input => {
    if (input == null || input === '') return [];
    if (!Array.isArray(input)) {
      return [parseKeyDefinition(input)];
    }

    if (input.length === 0) return [];

    if (input.some(Array.isArray)) {
      return input.map(parseKeyDefinition);
    }

    return [parseKeyDefinition(input)];
  };

  const sourceKeyGroups = normalizeKeyGroups(sourceKeyCols ?? sourceKeyCol);
  const targetKeyGroups = normalizeKeyGroups(targetKeyCols ?? targetKeyCol);

  const noKeyMode = sourceKeyGroups.length === 0 && targetKeyGroups.length === 0;
  if (!noKeyMode && (sourceKeyGroups.length === 0 || targetKeyGroups.length === 0)) {
    throw new Error('sourceKeyCol and targetKeyCol must both be provided for multi comparison mode, or both left empty for sequential mode.');
  }
  if (!noKeyMode && sourceKeyGroups.length !== targetKeyGroups.length) {
    throw new Error('sourceKeyCol and targetKeyCol must contain the same number of key definitions.');
  }
  if (!noKeyMode && sourceKeyGroups.some(g => g.length === 0)) {
    throw new Error('sourceKeyCol contains an empty key definition.');
  }
  if (!noKeyMode && targetKeyGroups.some(g => g.length === 0)) {
    throw new Error('targetKeyCol contains an empty key definition.');
  }

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

  const outputFilterEntries = Object.entries(outputMap);

  const checkFilter = (rowValues, filterMap) => {
    return Object.entries(filterMap).every(([colIdx, condition]) => {
      const val = rowValues[parseInt(colIdx)];
      if (Array.isArray(condition)) {
        return condition.includes(val);
      }
      if (typeof condition === 'string' && condition.startsWith('!')) {
        return val !== condition.slice(1);
      }
      return val === condition;
    });
  };

  const srcLastRow = srcSheet.getLastRow();
  const srcLastCol = srcSheet.getLastColumn();
  const numSrcRows = srcLastRow < sourceStartRow ? 0 : srcLastRow - sourceStartRow + 1;
  const srcData = numSrcRows > 0
    ? srcSheet.getRange(sourceStartRow, 1, numSrcRows, srcLastCol).getValues()
    : [];

  const tgtLastRow = tgtSheet.getLastRow();
  const tgtLastCol = tgtSheet.getLastColumn();
  const numTgtRows = noKeyMode ? 0 : (tgtLastRow < targetStartRow ? 0 : tgtLastRow - targetStartRow + 1);
  const tgtData = numTgtRows > 0
    ? tgtSheet.getRange(targetStartRow, 1, numTgtRows, tgtLastCol).getValues()
    : [];

  const makeKey = (row, idxs) => idxs.map(i => (row && i < row.length) ? row[i] : '').join('|');
  const isBlankKey = key => !key || key.split('|').every(part => part == null || part === '');
  const passesOutputFilter = rowValues => {
    return outputFilterEntries.every(([colIdx, condition]) => {
      const val = rowValues[parseInt(colIdx)];
      if (Array.isArray(condition)) {
        return condition.includes(val);
      }
      if (typeof condition === 'string' && condition.startsWith('!')) {
        return val !== condition.slice(1);
      }
      return val === condition;
    });
  };

  const updates = [];

  if (noKeyMode) {
    srcData.forEach((srcRow, srcIdx) => {
      if (!Object.entries(inputMap).every(([ci, f]) => {
        const v = srcRow[ci];
        return Array.isArray(f) ? f.includes(v)
          : (typeof f === 'string' && f.startsWith('!') ? v !== f.slice(1)
            : v === f);
      })) return;

      processedMappings.forEach(m => {
        const tc = m.targetColIdx;
        const baseValue = m.isFixedValue ? m.fixedValue : srcRow[m.sourceColIdx];
        const tgtRowNum = targetStartRow + srcIdx;
        let existing = '';
        try {
          if (tgtRowNum <= tgtSheet.getMaxRows() && (tc + 1) <= tgtSheet.getMaxColumns()) {
            existing = tgtSheet.getRange(tgtRowNum, tc + 1).getValue();
          } else {
            console.warn(`Target cell R${tgtRowNum}C${tc + 1} is out of bounds for sheet "${targetSheetName}". Defaulting to empty for comparison.`);
          }
        } catch (e) {
          console.warn(`Error reading target cell at row ${tgtRowNum}, col ${tc + 1}. Defaulting to empty. Error: ${e.message}`);
        }
        if ((overwrite || existing === '' || existing == null) && existing !== baseValue) {
          updates.push({ row: tgtRowNum, col: tc + 1, value: baseValue });
        }
      });
    });
  } else {
    const keyPairs = sourceKeyGroups.map((sourceGroup, pairIdx) => ({
      sourceIdxs: sourceGroup.map(l => letterToCol(l) - 1),
      targetIdxs: targetKeyGroups[pairIdx].map(l => letterToCol(l) - 1)
    }));

    const targetMaps = keyPairs.map(({ targetIdxs }) => {
      const targetMap = new Map();
      tgtData.forEach((tgtRow, idx) => {
        const key = makeKey(tgtRow, targetIdxs);
        if (isBlankKey(key)) return;

        const candidate = {
          rowIndexInTgtDataArray: idx,
          originalSheetRowIndex: targetStartRow + idx,
          rowData: tgtRow
        };

        if (!targetMap.has(key)) {
          targetMap.set(key, []);
        }

        targetMap.get(key).push(candidate);
      });
      return targetMap;
    });

    srcData.forEach(srcRow => {
      if (!checkFilter(srcRow, inputMap)) return;

      let bestMatch = null;

      for (let pairIdx = 0; pairIdx < keyPairs.length; pairIdx++) {
        const sourceKey = makeKey(srcRow, keyPairs[pairIdx].sourceIdxs);
        if (isBlankKey(sourceKey)) continue;

        const candidates = targetMaps[pairIdx].get(sourceKey);
        if (!candidates || candidates.length === 0) continue;

        for (let i = candidates.length - 1; i >= 0; i--) {
          const candidate = candidates[i];
          if (!passesOutputFilter(candidate.rowData)) continue;

          if (!bestMatch || candidate.originalSheetRowIndex > bestMatch.originalSheetRowIndex) {
            bestMatch = candidate;
          }
          break;
        }
      }

      if (!bestMatch) return;

      processedMappings.forEach(m => {
        const srcVal = m.isFixedValue ? m.fixedValue : (m.sourceColIdx < srcRow.length ? srcRow[m.sourceColIdx] : '');
        const existingVal = (m.targetColIdx < bestMatch.rowData.length) ? bestMatch.rowData[m.targetColIdx] : '';

        if (String(existingVal) !== String(srcVal)) {
          updates.push({
            row: bestMatch.originalSheetRowIndex,
            col: m.targetColIdx + 1,
            value: srcVal
          });
        }
      });
    });
  }

  if (updates.length > 0 && timestampCols && timestampCols.length > 0) {
    if (!timestampLocale || !timestampFormat) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Carimbo de data/hora solicitado, mas locale ou formato ausentes. Ignorando carimbo.', '⚠️ Atenção', 5);
      console.warn('Timestamping requested (timestampCols provided), but timestampLocale or timestampFormat is missing. Skipping timestamping.');
    } else {
      let formattedTimestamp;
      try {
        const now = new Date();
        formattedTimestamp = Utilities.formatDate(now, timestampLocale, timestampFormat);
      } catch (e) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Erro ao formatar carimbo: ${e.message}. Ignorando.`, '❌ Erro no Carimbo', 5);
        console.error(`Error formatting timestamp with locale "${timestampLocale}" and format "${timestampFormat}": ${e.message}. Skipping timestamping.`);
      }

      if (formattedTimestamp !== undefined) {
        const affectedRowNumbers = new Set();
        updates.forEach(u => affectedRowNumbers.add(u.row));

        const timestampTargetColIndices = [];
        for (const colLetter of timestampCols) {
          try {
            timestampTargetColIndices.push(letterToCol(colLetter) - 1);
          } catch (e) {
            console.warn(`Invalid column letter "${colLetter}" in timestampCols. Skipping this column for timestamping. Error: ${e.message}`);
          }
        }

        affectedRowNumbers.forEach(rowNum => {
          timestampTargetColIndices.forEach(targetColIdx => {
            updates.push({ row: rowNum, col: targetColIdx + 1, value: formattedTimestamp });
          });
        });
      }
    }
  }

  if (updates.length) {
    updates.forEach(u => {
      if (u.row > 0 && u.col > 0 && u.row <= tgtSheet.getMaxRows() && u.col <= tgtSheet.getMaxColumns()) {
        tgtSheet.getRange(u.row, u.col).setValue(u.value);
      } else {
        console.warn(`Skipping update for out-of-bounds cell: Row ${u.row}, Col ${u.col} in sheet "${targetSheetName}"`);
      }
    });
  }

  const msg = updates.length
    ? `${updates.length} células atualizadas com sucesso.`
    : 'Nenhuma célula modificada.';
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, '✅ Resultado', 5);
}

// Example:
// function executeUpdateByKeyMulti() {
//   updateByKeyMulti({
//     sourceSpreadsheetId: 'SOURCE_ID',
//     sourceSheetName: 'Main',
//     sourceStartRow: 2,
//     sourceKeyCol: [['C'], ['B'], ['B', 'C']],
//     targetSpreadsheetId: 'TARGET_ID',
//     targetSheetName: 'Main',
//     targetStartRow: 2,
//     targetKeyCol: [['C'], ['A'], ['B', 'C']],
//     colsMap: {
//       'A': 'A',
//       'B': 'B'
//     },
//     overwrite: false
//   });
// }