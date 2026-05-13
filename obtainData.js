// This grab the data from the spreadsheet

function obtainDataAll(update = false) { // Added optional update parameter, defaults to false

  // TROCA EM GARANTIA

  obtainData({
    input: {
      spreadsheetId: '1SkGLGBJlUyStxUimCEcczrKN8jx0bnj_SNJ1Lbg_nc8',
      sheetName: 'Under Warranty',
      startRow: 3,
      ticketColumn: 'A'
    },
    output: {
      spreadsheetId: '19l8eX_MuHNdBOzBZXyoKDdefI3K_oElpUvPpTbeFxaY',
      sheetName: 'DENTRO DE GARANTIA',
      startRow: 2,
      ticketColumn: 'A'
    },
    columnMap: {
      'A': 'A',
      'B': 'B',
      'C': 'C',
      'D': 'D',
      'E': 'E',
      'F': 'F',
      'G': 'G',
      'H': 'H',
      'I': 'I',
      'J': 'J',
      'K': 'K',
      'L': 'L',
      'M': 'M',
      'ADVANCED EXCHANGE': 'N'
    },
    filter: {
      filterCol: 'T',
      filterVal: ['TROCA EM GARANTIA (S1.0)']
    },
    // ignore: {
    //   spreadsheetId: '1TOea7J1Hwpl_4jNFADKQ1wMGzFqMt7wWcJjefUv3EM8',
    //   sheetName: 'IGNORE',
    //   startRow: 2,
    //   ticketColumn: 'A',
    //   compareColumn: 'A'
    // }
    update: update // Pass the update flag
  });

  // REPARO NO DEFEITUOSO

  obtainData({
    input: {
      spreadsheetId: '1SkGLGBJlUyStxUimCEcczrKN8jx0bnj_SNJ1Lbg_nc8',
      sheetName: 'Under Warranty',
      startRow: 3,
      ticketColumn: 'A'
    },
    output: {
      spreadsheetId: '19l8eX_MuHNdBOzBZXyoKDdefI3K_oElpUvPpTbeFxaY',
      sheetName: 'DENTRO DE GARANTIA',
      startRow: 2,
      ticketColumn: 'A'
    },
    columnMap: {
      'A': 'A',
      'B': 'B',
      'C': 'C',
      'D': 'D',
      'E': 'E',
      'F': 'F',
      'G': 'G',
      'H': 'H',
      'I': 'I',
      'J': 'J',
      'K': 'K',
      'L': 'L',
      'M': 'M',
      'REPAIR THE DEFECTIVE': 'N'
    },
    filter: {
      filterCol: 'T',
      filterVal: ['REPARO NO DEFEITUOSO (S1.RD)']
    },
    // ignore: {
    //   spreadsheetId: '1TOea7J1Hwpl_4jNFADKQ1wMGzFqMt7wWcJjefUv3EM8',
    //   sheetName: 'IGNORE',
    //   startRow: 2,
    //   ticketColumn: 'A',
    //   compareColumn: 'A'
    // }
    update: true
  });

  obtainData({
    input: {
      spreadsheetId: '19l8eX_MuHNdBOzBZXyoKDdefI3K_oElpUvPpTbeFxaY',
      sheetName: 'DENTRO DE GARANTIA',
      startRow: 2,
      ticketColumn: 'Q'
    },
    output: {
      spreadsheetId: '1TOea7J1Hwpl_4jNFADKQ1wMGzFqMt7wWcJjefUv3EM8',
      sheetName: 'Main',
      startRow: 2,
      ticketColumn: 'P'
    },
    columnMap: {
      'A': 'A',
      'B': 'B',
      'C': 'C',
      'E': 'D',
      'REPAIR THE DEFECTIVE': 'E',
      'L': 'F',
      'Q': 'P'
    },
    filter: {
      filterCol: 'N',
      filterVal: ['REPAIR THE DEFECTIVE']
    },
    ignore: {
      spreadsheetId: '1TOea7J1Hwpl_4jNFADKQ1wMGzFqMt7wWcJjefUv3EM8',
      sheetName: 'IGNORE',
      startRow: 2,
      ticketColumn: 'A',
      compareColumn: 'Q'
    },
    update: true
  });
}

function obtainData(config) {
  const { input, output, columnMap, filter, ignore, update = false } = config; // Added update, default to false

  // Obter dados da entrada
  const inputData = getSheetData(input.spreadsheetId, input.sheetName, input.startRow);
  if (!inputData || inputData.length === 0) return;

  // Obter dados da saída
  const outputData = getSheetData(output.spreadsheetId, output.sheetName, output.startRow);
  const outputTicketColIdx = letterToCol(output.ticketColumn);
  const outputTickets = new Set();
  if (outputTicketColIdx >= 0) {
    outputData.forEach(row => {
      if (outputTicketColIdx < row.length && row[outputTicketColIdx] !== undefined && row[outputTicketColIdx] !== "") {
        outputTickets.add(row[outputTicketColIdx]);
      }
    });
  }


  // Se configurado, obter tickets para ignorar
  let ignoreTickets = new Set();
  if (ignore && ignore.spreadsheetId && ignore.sheetName && ignore.ticketColumn) {
    const ignoreData = getSheetData(ignore.spreadsheetId, ignore.sheetName, ignore.startRow);
    const ignoreTicketColIdx = letterToCol(ignore.ticketColumn);
    if (ignoreTicketColIdx >= 0) {
      ignoreData.forEach(row => {
        if (ignoreTicketColIdx < row.length && row[ignoreTicketColIdx] !== undefined && row[ignoreTicketColIdx] !== "") {
          ignoreTickets.add(row[ignoreTicketColIdx]);
        }
      });
    }
  }

  // Filtrar dados da entrada
  const filteredData = inputData.filter(row => {
    const inputTicketColIdx = letterToCol(input.ticketColumn);
    let ticketValue;
    if (inputTicketColIdx >= 0 && inputTicketColIdx < row.length) {
      ticketValue = row[inputTicketColIdx];
    }

    if (!ticketValue) return false; // make sure to cancel if row's empty or ticketValue is empty/undefined

    let matchesFilter = true;
    if (filter && filter.filterCol && filter.filterVal) { // Ensure filter object and its properties exist
      const filterColIdx = letterToCol(filter.filterCol);
      if (filterColIdx >= 0 && filterColIdx < row.length) {
        matchesFilter = filter.filterVal.includes(row[filterColIdx]);
      } else {
        matchesFilter = false; // Column to filter on does not exist in row or is invalid
      }
    }


    const notInOutputOrUpdateMode = true || !outputTickets.has(ticketValue);


    let notIgnored = true;
    if (ignore && ignore.compareColumn) {
      const compareColIdx = letterToCol(ignore.compareColumn);
      if (compareColIdx >= 0 && compareColIdx < row.length) {
        const compareValue = row[compareColIdx];
        notIgnored = !ignoreTickets.has(compareValue);
      } else {
        notIgnored = true; // If compare column doesn't exist, consider it not ignored. Or false, depending on desired logic.
      }
    }
    return matchesFilter && notInOutputOrUpdateMode && notIgnored;
  });

  // Mapear colunas e valores constantes
  const mappedData = filteredData.map(row => {
    const mappedRow = [];
    for (const [origin, destiny] of Object.entries(columnMap)) {
      const originIndex = letterToCol(origin);
      const destinyIndex = letterToCol(destiny);
      if (destinyIndex < 0) continue;
      if (originIndex >= 0 && originIndex < row.length) { // Check if origin is a valid column index within the row
        mappedRow[destinyIndex] = row[originIndex];
      } else {
        // If origin is not a valid column letter (e.g. "ADVANCED EXCHANGE"), treat as constant value
        // And also ensure originIndex isn't just out of bounds for a valid letter
        // A more robust check: if letterToCol(origin) is -1 (not a col letter) OR it's out of bounds
        if (originIndex === -1 || originIndex >= row.length) { // Treat as constant value
          mappedRow[destinyIndex] = origin;
        } else {
          // This case should ideally not be hit if originIndex was >=0 but < row.length was handled above
          // For safety, assign undefined or handle as error
          mappedRow[destinyIndex] = undefined;
        }
      }
    }
    return mappedRow;
  });

  // Escrever dados na saída
  writeToSheet(output.spreadsheetId, output.sheetName, output.startRow, mappedData, update, output.ticketColumn); // Pass update flag and outputTicketColumn
}

function getSheetData(spreadsheetId, sheetName, startRow) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // Logger.log(`Sheet ${sheetName} not found in spreadsheet ${spreadsheetId}`);
    return []; // Return empty array if sheet not found
  }
  const lastRow = sheet.getLastRow();
  // If startRow is beyond any data, or sheet is empty and startRow > 0
  if (startRow > lastRow && lastRow > 0) {
    return [];
  }
  const numRowsToGet = Math.max(0, lastRow - startRow + 1);
  if (numRowsToGet === 0) {
    return [];
  }
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(startRow, 1, numRowsToGet, lastCol > 0 ? lastCol : 1);
  return range.getValues();
}

function writeToSheet(spreadsheetId, sheetName, startRow, data, update = false, outputTicketColumnLetter) { // Added update and outputTicketColumnLetter
  if (data.length === 0) return;
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // Logger.log(`Sheet ${sheetName} not found in spreadsheet ${spreadsheetId} for writing.`);
    return;
  }
  const outputTicketColIdx = letterToCol(outputTicketColumnLetter);

  if (update && outputTicketColIdx >= 0) { // Update logic only makes sense if there's a valid ticket column
    const lastSheetRow = sheet.getLastRow();
    const numExistingRowsToRead = Math.max(0, lastSheetRow - startRow + 1);
    const sheetMaxCols = sheet.getMaxColumns(); // Get max columns of the sheet once

    const actualExistingData = numExistingRowsToRead > 0 ?
      sheet.getRange(startRow, 1, numExistingRowsToRead, sheetMaxCols > 0 ? sheetMaxCols : 1).getValues() :
      [];

    const existingTicketToRowIndexMap = new Map(); // Maps ticket value to its 0-based index in actualExistingData
    actualExistingData.forEach((row, index) => {
      if (outputTicketColIdx < row.length && row[outputTicketColIdx] !== undefined && row[outputTicketColIdx] !== '') {
        existingTicketToRowIndexMap.set(row[outputTicketColIdx], index);
      }
    });

    const rowsToAppend = [];

    data.forEach(newRowData => {
      const ticketValue = (outputTicketColIdx < newRowData.length) ? newRowData[outputTicketColIdx] : undefined;
      if (ticketValue !== undefined && ticketValue !== '' && existingTicketToRowIndexMap.has(ticketValue)) {
        const existingDataRowIndex = existingTicketToRowIndexMap.get(ticketValue); // 0-based index in actualExistingData
        const sheetRowToUpdate = startRow + existingDataRowIndex; // 1-based sheet row number

        // Start with a copy of the existing row from the sheet
        const targetSheetRowData = [...actualExistingData[existingDataRowIndex]];

        let changed = false;
        newRowData.forEach((val, idx) => {
          // Only consider mapped values from newRowData (where val is not undefined)
          if (val !== undefined) {
            // If this index is beyond the current length of targetSheetRowData, or the value is different
            if (idx >= targetSheetRowData.length || targetSheetRowData[idx] !== val) {
              targetSheetRowData[idx] = val; // Apply update, this might extend the row if idx is new
              changed = true;
            }
          }
        });

        if (changed) {
          // Determine the number of columns to write for this specific update.
          // It should be the length of the (potentially extended) targetSheetRowData.
          const numColsToWriteThisRow = targetSheetRowData.length;
          if (numColsToWriteThisRow > 0) {
            sheet.getRange(sheetRowToUpdate, 1, 1, numColsToWriteThisRow).setValues([targetSheetRowData]);
          }
        }
      } else {
        rowsToAppend.push(newRowData);
      }
    });

    if (rowsToAppend.length > 0) {
      let maxColsNew = 0;
      rowsToAppend.forEach(row => { if (row.length > maxColsNew) maxColsNew = row.length; });

      if (maxColsNew === 0 && rowsToAppend.length > 0) maxColsNew = 1; // Ensure at least one column if appending essentially "empty" rows

      if (maxColsNew > 0) {
        const finalRowsToAppend = rowsToAppend.map(row => {
          const fullRow = Array(maxColsNew).fill(undefined); // Use undefined for empty cells
          row.forEach((val, idx) => fullRow[idx] = val);
          return fullRow;
        });
        const appendTargetRow = sheet.getLastRow() + 1;
        sheet.getRange(appendTargetRow, 1, finalRowsToAppend.length, maxColsNew).setValues(finalRowsToAppend);
      }
    }

  } else {
    // Original append-only logic (or if update is true but ticket column is invalid/not provided)
    const numExistingRows = Math.max(0, sheet.getLastRow() - startRow + 1);
    const sheetMaxCols = sheet.getMaxColumns();
    const existingData = numExistingRows > 0 ? sheet.getRange(startRow, 1, numExistingRows, sheetMaxCols > 0 ? sheetMaxCols : 1).getValues() : [];

    let numColsInData = 0; // Max number of columns present in the incoming 'data'
    if (data.length > 0) {
      data.forEach(r => { if (r.length > numColsInData) numColsInData = r.length; });
    }
    // If data has rows but no columns (e.g. [[], []]), there's nothing to compare or write meaningfully per row.
    if (numColsInData === 0 && data.length > 0) return;

    const existingRows = new Set();
    if (numColsInData > 0) { // Only build the set if there's a basis for comparison
      existingData.forEach(row => {
        const normalizedPortion = Array(numColsInData).fill(undefined);
        for (let i = 0; i < numColsInData; i++) {
          // Populate normalizedPortion from existing row, up to numColsInData
          if (i < row.length) {
            // Treat null and empty string from sheet as 'undefined' for comparison consistency
            normalizedPortion[i] = (row[i] === null || row[i] === "") ? undefined : row[i];
          }
        }
        existingRows.add(normalizedPortion.join(','));
      });
    }

    const uniqueData = [];
    if (numColsInData > 0) {
      data.forEach(row => {
        const normalizedPortion = Array(numColsInData).fill(undefined);
        // Populate normalizedPortion from incoming 'row' (from 'data')
        for (let i = 0; i < numColsInData; i++) {
          if (i < row.length) {
            normalizedPortion[i] = (row[i] === null || row[i] === "") ? undefined : row[i];
          }
        }
        if (!existingRows.has(normalizedPortion.join(','))) {
          uniqueData.push(row); // Add the original sparse row; padding happens before setValues
        }
      });
    } else if (data.length > 0 && numColsInData === 0) {
      // This case means data might be like [[],[]], no content to make unique.
      // Depending on desired behavior, could append all or none. Current logic results in no appends.
    }

    if (uniqueData.length === 0) return;

    let maxColsForAppend = 0; // Max columns in the actual uniqueData to be appended
    uniqueData.forEach(row => { if (row.length > maxColsForAppend) maxColsForAppend = row.length; });

    // If after filtering, uniqueData has rows but they are all empty (maxColsForAppend is 0)
    if (maxColsForAppend === 0 && uniqueData.length > 0) maxColsForAppend = 1; // Default to 1 col to avoid error
    if (maxColsForAppend === 0) return; // Still nothing to write

    const finalUniqueData = uniqueData.map(row => {
      const fullRow = Array(maxColsForAppend).fill(undefined);
      row.forEach((val, idx) => fullRow[idx] = val);
      return fullRow;
    });

    const appendStartSheetRow = sheet.getLastRow() + 1;
    sheet.getRange(appendStartSheetRow, 1, finalUniqueData.length, maxColsForAppend).setValues(finalUniqueData);
  }
}

// Converte letras de coluna em índice baseado em 0 (A -> 0, B -> 1, ..., Z -> 25, AA -> 26, etc.)
const letterToCol = l => {
  if (!l || typeof l !== 'string' || l.trim() === "") return -1; // Added type and empty string check
  return l.toUpperCase().split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0) - 1;
};