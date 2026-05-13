/**
 * Filters tickets from input spreadsheets based on configuration and transfers unique tickets to the output spreadsheet.
 *
 * This function validates the configuration, applies filters to input data, checks for duplicates in other sheets,
 * maps columns as specified, and writes the filtered data to the output sheet. It is robust to errors in inputB sheets
 * and continues processing if one fails. All filtering and mapping logic is driven by the config object.
 *
 * @param {Object} config - Configuration object for the operation.
 * @param {Object} config.inputA - Main input sheet configuration (spreadsheetId, sheetName, startRow, ticketColumn, filter).
 * @param {Array<Object>} config.inputB - Array of additional input sheet configurations.
 * @param {Object} config.output - Output sheet configuration (spreadsheetId, sheetName, startRow, ticketColumn, filter).
 * @param {Object} config.columnMap - Mapping from source columns to destination columns.
 * @throws {Error} If configuration is invalid or if reading/writing sheets fails.
 *
 * @example
 * filterAndTransferTickets({
 *   inputA: {
 *     spreadsheetId: 'spreadsheet_id_1',
 *     sheetName: 'Sheet1',
 *     startRow: 2,
 *     ticketColumn: 'A',
 *     filter: { column: 'N', values: ['ADVANCED EXCHANGE', 'REPAIR THE DEFECTIVE'] },
 *   },
 *   inputB: [
 *     {
 *       spreadsheetId: 'spreadsheet_id_2',
 *       sheetName: 'Removed',
 *       startRow: 2,
 *       ticketColumn: 'A',
 *       filter: null,
 *     },
 *     {
 *       spreadsheetId: 'spreadsheet_id_3',
 *       sheetName: 'Envios',
 *       startRow: 2,
 *       ticketColumn: 'C',
 *       filter: null,
 *     },
 *   ],
 *   output: {
 *     spreadsheetId: 'spreadsheet_id_2',
 *     sheetName: 'Main',
 *     startRow: 2,
 *     ticketColumn: 'A',
 *     filter: null,
 *   },
 *   columnMap: {
 *     A: 'A',
 *     B: 'B',
 *     C: 'C',
 *     L: 'D',
 *     E: 'E',
 *     F: 'F',
 *     N: 'G',
 *     O: 'H',
 *   },
 * });
 */
function filterAndTransferTickets(config) {
  // Validate main configuration
  if (!config || typeof config !== 'object') {
    throw new Error('Missing or invalid main configuration. Expected an object.');
  }
  if (!config.inputA || typeof config.inputA !== 'object') {
    throw new Error('Missing or invalid "inputA" configuration. Expected an object.');
  }
  if (!config.inputB || !Array.isArray(config.inputB)) {
    throw new Error('Missing or invalid "inputB" configuration. Expected an array of configurations.');
  }
  if (!config.output || typeof config.output !== 'object') {
    throw new Error('Missing or invalid "output" configuration. Expected an object.');
  }
  if (!config.columnMap || typeof config.columnMap !== 'object') {
    throw new Error('Missing or invalid "columnMap" configuration. Expected an object.');
  }

  const {
    inputA,
    inputB,
    output,
    columnMap,
  } = config;

  // Validate inputA configuration
  ['spreadsheetId', 'sheetName', 'startRow', 'ticketColumn', 'filter'].forEach(prop => {
    if (inputA[prop] === undefined) {
      throw new Error(`Property "${prop}" missing in inputA configuration.`);
    }
  });
  if (typeof inputA.spreadsheetId !== 'string') throw new Error('inputA.spreadsheetId must be a string.');
  if (typeof inputA.sheetName !== 'string') throw new Error('inputA.sheetName must be a string.');
  if (!Number.isInteger(inputA.startRow) || inputA.startRow < 1) throw new Error('inputA.startRow must be a positive integer.');
  if (typeof inputA.ticketColumn !== 'string') throw new Error('inputA.ticketColumn must be a string.');
  if (inputA.filter !== null && typeof inputA.filter !== 'object') throw new Error('inputA.filter must be an object or null.');
  if (inputA.filter && typeof inputA.filter.column !== 'string') throw new Error('inputA.filter.column must be a string.');
  if (inputA.filter && !Array.isArray(inputA.filter.values)) throw new Error('inputA.filter.values must be an array.');

  // Validate output configuration
  ['spreadsheetId', 'sheetName', 'startRow', 'ticketColumn'].forEach(prop => {
    if (output[prop] === undefined) {
      throw new Error(`Property "${prop}" missing in output configuration.`);
    }
  });
  if (typeof output.spreadsheetId !== 'string') throw new Error('output.spreadsheetId must be a string.');
  if (typeof output.sheetName !== 'string') throw new Error('output.sheetName must be a string.');
  if (!Number.isInteger(output.startRow) || output.startRow < 1) throw new Error('output.startRow must be a positive integer.');
  if (typeof output.ticketColumn !== 'string') throw new Error('output.ticketColumn must be a string.');
  if (output.filter !== null && typeof output.filter !== 'object') throw new Error('output.filter must be an object or null.');

  // Validate columnMap
  for (const source in columnMap) {
    if (typeof source !== 'string') throw new Error('Keys of columnMap (source columns) must be strings.');
    if (typeof columnMap[source] !== 'string') throw new Error('Values of columnMap (destination columns) must be strings.');
  }

  /**
   * Applies a filter to the data based on filterConfig.
   * Returns the filtered data or the original data if filterConfig is invalid.
   * @param {Array<Array>} data - The data to filter.
   * @param {Object|null} filterConfig - The filter configuration. Should use 'column' and 'values' as keys.
   * @returns {Array<Array>} Filtered data.
   */
  function applyFilter(data, filterConfig) {
    if (!filterConfig || !filterConfig.column || !filterConfig.values || !Array.isArray(filterConfig.values)) {
      return data;
    }
    try {
      // Try-catch to handle unexpected filter errors
      return data.filter((row) => {
        return filterConfig.values.includes(row[columnConverter(filterConfig.column) - 1]);
      });
    } catch (filterError) {
      Logger.log('Error applying filter: ' + filterError);
      SpreadsheetApp.getUi().alert('Error applying filter. Check filter configuration. ' + filterError);
      return data; // On filter error, return unfiltered data to avoid total failure.
    }
  }

  let inputAData;
  try {
    inputAData = getSheetData(inputA.spreadsheetId, inputA.sheetName, inputA.startRow);
  } catch (e) {
    throw new Error(`Error getting inputA data: ${e.message || e}`); // Rethrow to stop if inputA fails
  }
  inputAData = applyFilter(inputAData, inputA.filter);

  let allInputBTickets = new Set();
  if (inputB && Array.isArray(inputB)) {
    for (const input of inputB) {
      // Validate each inputB configuration
      ['spreadsheetId', 'sheetName', 'startRow', 'ticketColumn', 'filter'].forEach(prop => {
        if (input[prop] === undefined) {
          throw new Error(`Property "${prop}" missing in an inputB configuration.`);
        }
      });
      if (typeof input.spreadsheetId !== 'string') throw new Error('inputB item spreadsheetId must be a string.');
      if (typeof input.sheetName !== 'string') throw new Error('inputB item sheetName must be a string.');
      if (!Number.isInteger(input.startRow) || input.startRow < 1) throw new Error('inputB item startRow must be a positive integer.');
      if (typeof input.ticketColumn !== 'string') throw new Error('inputB item ticketColumn must be a string.');
      if (input.filter !== null && typeof input.filter !== 'object') throw new Error('inputB item filter must be an object or null.');
      if (input.filter && typeof input.filter.column !== 'string') throw new Error('inputB item filter.column must be a string.');
      if (input.filter && !Array.isArray(input.filter.values)) throw new Error('inputB item filter.values must be an array.');

      try {
        let inputBData = getSheetData(input.spreadsheetId, input.sheetName, input.startRow);
        inputBData = applyFilter(inputBData, input.filter);
        const inputBTickets = inputBData.map((row) => row[columnConverter(input.ticketColumn) - 1]);
        inputBTickets.forEach((ticket) => allInputBTickets.add(ticket));
      } catch (error) {
        Logger.log(`Error processing inputB sheet (ID: ${input.spreadsheetId}, Tab: ${input.sheetName}): ${error}`);
        SpreadsheetApp.getUi().alert(
          `Error processing inputB sheet (Tab: ${input.sheetName}). Check ID and tab name. Details in log. ${error}`
        );
        continue; // Continue with other inputB sheets instead of stopping everything.
      }
    }
  }

  let outputData;
  try {
    outputData = getSheetData(output.spreadsheetId, output.sheetName, output.startRow);
  } catch (e) {
    throw new Error(`Error getting output data: ${e.message || e}`); // Rethrow if output fails
  }
  outputData = applyFilter(outputData, output.filter);
  const outputTickets = new Set(outputData.map((row) => row[columnConverter(output.ticketColumn) - 1]));

  // --- Final filtering: only tickets not present in inputB or output ---
  const filteredData = inputAData.filter((row) => {
    const ticket = row[columnConverter(inputA.ticketColumn) - 1];
    const notInInputB = !allInputBTickets.has(ticket);
    const notInOutput = !outputTickets.has(ticket);
    return notInInputB && notInOutput; // Only check for duplicates!
  });

  // Map columns from inputA to output as specified in columnMap
  const mappedData = filteredData.map((row) => {
    const mappedRow = [];
    for (const [source, target] of Object.entries(columnMap)) {
      const sourceIndex = columnConverter(source) - 1;
      const targetIndex = columnConverter(target) - 1;
      mappedRow[targetIndex] = row[sourceIndex];
    }
    return mappedRow;
  });

  try {
    writeToSheet(output.spreadsheetId, output.sheetName, output.startRow, mappedData);
  } catch (e) {
    throw new Error(`Error writing data to output sheet: ${e.message || e}`); // Rethrow if write fails
  }
}

/**
 * Example function to filter and transfer tickets using filterAndTransferTickets.
 * Demonstrates usage with multiple input sheets and a service type filter on inputA.
 */
function obtainDataAll() {
  // Example with multiple verification sheets (inputB)
  // Now, service type classification is done *exclusively* by the 'inputA' filter.
  // There is no longer a 'global' filter.

  utils.filterAndTransferTickets({
    inputA: {
      spreadsheetId: '19l8eX_MuHNdBOzBZXyoKDdefI3K_oElpUvPpTbeFxaY',
      sheetName: 'DENTRO DE GARANTIA',
      startRow: 2,
      ticketColumn: 'A',
      filter: { column: 'N', values: ['ADVANCED EXCHANGE', 'REPAIR THE DEFECTIVE'] },
    },
    inputB: [
      {
        spreadsheetId: '1XUWI5WV7_EIrSSDM4yLsKlZWC7cfQGkQbMrLwn--iJE',
        sheetName: 'Removed',
        startRow: 2,
        ticketColumn: 'A',
        filter: null,
      },
      {
        spreadsheetId: '1UU2zwxcZK0OieuH0RNjq2wFtzSCPQC9PwBUmolU5xHI',
        sheetName: 'Envios',
        startRow: 2,
        ticketColumn: 'C',
        filter: null,
        // filter: { column: 'H', values: ['SENT'] },
      },
    ],
    output: {
      spreadsheetId: '1XUWI5WV7_EIrSSDM4yLsKlZWC7cfQGkQbMrLwn--iJE',
      sheetName: 'Main',
      startRow: 2,
      ticketColumn: 'A',
      filter: null, // No sense to use this here, but kept for compatibility.
    },
    columnMap: {
      A: 'A',
      B: 'B',
      C: 'C',
      L: 'D',
      E: 'E',
      F: 'F',
      N: 'G',
      O: 'H',
    },
  });
}