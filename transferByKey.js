// function executeupdateByKey() {
//   updateByKey({
//     sourceSpreadsheetId: '1CkBD-VK_v1bRbSzA2H7Eq55NtEDCjANcRPZBOg2_uPc',
//     sourceSheetName: 'Main',
//     sourceStartRow: 3,
//     sourceKeyCol: ['B, C'],

//     targetSpreadsheetId: '1sbNHOQ0-m2RaZmOy4w1nhwKeYXjH-if0zPTFKyOcaR4',
//     targetSheetName: 'REINCIDENCE',
//     targetStartRow: 2,
//     targetKeyCol: ['B, H'],

//     colsMap: {
//       'A': 'I',
//       'V': 'M',
//     },

//     overwrite: false,
//     inputFil: { 'T': 'FINISH' },
//     // outputFil: { 'AA': '!CANCELLED' },
//     // timestampCols: ['F'], // Exemplo de coluna para carimbo
//     // timestampLocale: 'America/Sao_Paulo',
//     // timestampFormat: "dd/MM/yyyy"
//   });

//     updateByKey({
//     sourceSpreadsheetId: '1CkBD-VK_v1bRbSzA2H7Eq55NtEDCjANcRPZBOg2_uPc',
//     sourceSheetName: 'Main',
//     sourceStartRow: 3,
//     sourceKeyCol: ['C'],

//     targetSpreadsheetId: '1sbNHOQ0-m2RaZmOy4w1nhwKeYXjH-if0zPTFKyOcaR4',
//     targetSheetName: 'Reincidence',
//     targetStartRow: 2,
//     targetKeyCol: ['H'],

//     colsMap: {
//       'A': 'I',
//       'V': 'M',
//     },

//     overwrite: false,
//     inputFil: { 'T': 'FINISH' },
//     // outputFil: { 'AA': '!CANCELLED' },
//     // timestampCols: ['F'], // Exemplo de coluna para carimbo
//     // timestampLocale: 'America/Sao_Paulo',
//     // timestampFormat: "dd/MM/yyyy"
//   });
// }

/**
 * Transfers data from one sheet to another by matching a key column or multiple key columns.
 * If sourceKeyCol or targetKeyCol are empty, rows are copied sequentially without matching keys.
 * You can optionally provide:
 *  – inputFil: filters on source sheet columns, specified by column letter.
 *  – outputFil: filters on target sheet columns, specified by column letter.
 *
 * Filters support:
 *  – "value" = exact match
 *  – "!value" = not equal
 *  – ["A", "B"] = OR match
 *
 * @param {Object} config                   – Transfer configuration
 * @param {string}   config.sourceSpreadsheetId   – ID of the source spreadsheet.
 * @param {string}   config.sourceSheetName       – Name of the source sheet/tab.
 * @param {number}   config.sourceStartRow        – 1-based row number to start reading in source.
 * @param {string}   config.sourceKeyCol          – Column letter(s) of the key in source (e.g., "B" or "B,C").
 * @param {string}   config.targetSpreadsheetId   – ID of the target spreadsheet.
 * @param {string}   config.targetSheetName       – Name of the target sheet/tab.
 * @param {number}   config.targetStartRow        – 1-based row number to start writing in target.
 * @param {string}   config.targetKeyCol          – Column letter(s) of the key in target.
 * @param {Object}   config.colsMap              – Column mapping {sourceColOrFixedValue: targetCol}.
 * @param {Object}   [config.inputFil]           – (optional) Filters for source data, by column letter.
 * @param {Object}   [config.outputFil]          – (optional) Filters for target data, by column letter.
 * @param {boolean}  config.overwrite            – If true, always overwrite; if false, only fill empty cells.
 * @param {string[]} [config.timestampCols]      – (optional) Array of target column letters to receive a timestamp.
 * @param {string}   [config.timestampLocale]    – (optional) Timezone for timestamp (e.g., "America/Sao_Paulo"). Required if timestampCols is used.
 * @param {string}   [config.timestampFormat]    – (optional) Format for timestamp (e.g., "dd/MM/yyyy HH:mm:ss"). Required if timestampCols is used.
 */

function updateByKey(config) {
  const {
    sourceSpreadsheetId, sourceSheetName, sourceStartRow, sourceKeyCol = [],
    targetSpreadsheetId, targetSheetName, targetStartRow, targetKeyCol = [],
    colsMap, overwrite, inputFil = {}, outputFil = {},
    timestampCols = [], timestampLocale, timestampFormat
  } = config;

  // Converte letra para índice 0-based
  const letterToCol = l => {
    const L = String(l).toUpperCase();
    if (!/^[A-Z]+$/.test(L)) {
      throw new Error(`Invalid column letter format: "${l}". Must consist of A-Z characters only.`);
    }
    if (L.length > 3) { // XFD is max
      // This specific error part about fixed value is more relevant for colsMap source keys
      // For other uses like timestampCols, it's simply an invalid column.
      // However, to keep the function identical except for the new feature, we retain it.
      // The caller using it for timestampCols should handle the error if it expects only column letters.
      throw new Error(`Column string "${l}" is too long (${L.length} chars) to be a standard column letter (max 3 chars). Interpreting as fixed value.`);
    }
    return L.split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);
  };

  // Suporta múltiplas colunas de chave, separadas por vírgula
  const sourceKeyCols = sourceKeyCol ? String(sourceKeyCol).split(/\s*,\s*/) : [];
  const targetKeyCols = targetKeyCol ? String(targetKeyCol).split(/\s*,\s*/) : [];

  const srcKeyIdxs = sourceKeyCols.map(l => letterToCol(l) - 1);
  const tgtKeyIdxs = targetKeyCols.map(l => letterToCol(l) - 1);

  const noKeyMode = srcKeyIdxs.length === 0 || tgtKeyIdxs.length === 0;

  const srcSheet = SpreadsheetApp.openById(sourceSpreadsheetId).getSheetByName(sourceSheetName);
  const tgtSheet = SpreadsheetApp.openById(targetSpreadsheetId).getSheetByName(targetSheetName);

  // Processa mapeamento de colunas
  const processedMappings = [];
  const sourceDefs = Object.keys(colsMap);
  const targetLetters = Object.values(colsMap);

  for (let i = 0; i < sourceDefs.length; i++) {
    const srcDef = sourceDefs[i];
    const tgtLet = targetLetters[i];
    let tgtIdx;
    try {
      tgtIdx = letterToCol(tgtLet) - 1;
    } catch (e) {
      throw new Error(`Invalid target column letter in colsMap for source "${srcDef}": "${tgtLet}". ${e.message}`);
    }

    const map = { targetColIdx: tgtIdx, isFixedValue: false, sourceColIdx: -1, fixedValue: null };
    try {
      map.sourceColIdx = letterToCol(srcDef) - 1;
    } catch (e) {
      map.isFixedValue = true;
      map.fixedValue = srcDef;
    }
    processedMappings.push(map);
  }

  // Converte filtros para índices
  const inputMap = Object.fromEntries(
    Object.entries(inputFil).map(([l, f]) => [letterToCol(l) - 1, f])
  );
  const outputMap = Object.fromEntries(
    Object.entries(outputFil).map(([l, f]) => [letterToCol(l) - 1, f])
  );

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

  const updates = [];

  // Função para gerar chave composta
  const makeKey = (row, idxs) => idxs.map(i => row[i]).join('|');

  if (noKeyMode) {
    // Modo sequencial
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
          // Ensure target row/col is valid before attempting to read
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
    // Modo por chave composta
    const srcMap = {};
    srcData.forEach(srcRow => {
      const keyParts = srcKeyIdxs.map(i => srcRow[i]); // Get key parts before joining
      if (keyParts.some(k => k == null || k === '')) return; // Check for blank key parts

      if (!Object.entries(inputMap).every(([ci, f]) => {
        const v = srcRow[ci];
        return Array.isArray(f) ? f.includes(v)
          : (typeof f === 'string' && f.startsWith('!') ? v !== f.slice(1)
            : v === f);
      })) return;
      srcMap[keyParts.join('|')] = processedMappings.map(m =>
        m.isFixedValue ? m.fixedValue : srcRow[m.sourceColIdx]
      );
    });

    tgtData.forEach((tgtRow, tgtIdx) => {
      const key = makeKey(tgtRow, tgtKeyIdxs);
      if (!(key in srcMap)) return;
      if (!Object.entries(outputMap).every(([ci, f]) => {
        const v = tgtRow[ci];
        return Array.isArray(f) ? f.includes(v)
          : (typeof f === 'string' && f.startsWith('!') ? v !== f.slice(1)
            : v === f);
      })) return;

      const values = srcMap[key];
      values.forEach((val, mi) => {
        const m = processedMappings[mi];
        if (val === undefined && !m.isFixedValue) {
          console.warn(`[KeyMode Update] Valor indefinido para chave composta ${key}, índice ${mi}.`);
          return;
        }
        const tc = m.targetColIdx;
        const existing = tc < tgtRow.length ? tgtRow[tc] : '';
        if ((overwrite || existing === '' || existing == null) && existing !== val) {
          updates.push({ row: targetStartRow + tgtIdx, col: tc + 1, value: val });
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

  // Aplica atualizações
  if (updates.length) {
    updates.forEach(u => {
      // Basic boundary check before writing
      if (u.row > 0 && u.col > 0 && u.row <= tgtSheet.getMaxRows() && u.col <= tgtSheet.getMaxColumns()) {
        tgtSheet.getRange(u.row, u.col).setValue(u.value);
      } else {
        console.warn(`Skipping update for out-of-bounds cell: Row ${u.row}, Col ${u.col} in sheet "${targetSheetName}"`);
      }
    });
  }

  const msg = updates.length
    ? `${updates.length} células atualizadas com sucesso.` // This count now includes timestamps
    : 'Nenhuma célula modificada.';
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, '✅ Resultado', 5);
}