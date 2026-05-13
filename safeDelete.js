/**
 * Safely deletes rows from a source sheet that meet specified filters,
 * ONLY IF the row's key is also found in a destination (comparison) sheet.
 * This ensures deletion from the source only occurs if there's a corresponding entry
 * in the destination sheet.
 * Empty key values in the source sheet are not deleted.
 * Deletions are performed in contiguous batches for efficiency.
 *
 * @param {Object} config
 * @param {string} config.spreadsheetId       - ID of the spreadsheet.
 * @param {string} config.sheetName           - Name of the primary (source) sheet from which rows might be deleted.
 * @param {number} config.startRow            - 1-based index of the starting row for evaluation in the source sheet.
 * @param {Object} [config.filter]            - Optional filter conditions in the form { columnLetter: condition }.
 * @param {string} config.compareSheetName  - Name of the destination/comparison sheet.
 * @param {string} config.keyColumn         - Column letter of the key in the source sheet (e.g., 'W').
 * @param {string} config.compareKeyColumn  - Column letter of the key in the destination/comparison sheet (e.g., 'W').
 */
function safeDelete(config) { // Renamed for clarity during debugging
  console.log("Starting safeDeleteStrict_Debug with config:", JSON.stringify(config));
  const {
    spreadsheetId,
    sheetName,
    startRow,
    filter,
    compareSheetName,
    keyColumn,
    compareKeyColumn
  } = config;

  if (!compareSheetName || !keyColumn || !compareKeyColumn) {
    const errorMessage = "Parameters 'compareSheetName', 'keyColumn', and 'compareKeyColumn' are mandatory.";
    SpreadsheetApp.getActive()?.toast(errorMessage, "⚠️ Incomplete Configuration", 10);
    console.error(errorMessage, config);
    return;
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sourceSheet = ss.getSheetByName(sheetName);
  if (!sourceSheet) {
    console.error(`Source sheet "${sheetName}" not found.`);
    SpreadsheetApp.getActive()?.toast(`Source sheet "${sheetName}" not found.`, "❌ Error", 5);
    return;
  }

  const destinationSheet = ss.getSheetByName(compareSheetName);
  if (!destinationSheet) {
    console.error(`Destination (comparison) sheet "${compareSheetName}" not found.`);
    SpreadsheetApp.getActive()?.toast(`Destination sheet "${compareSheetName}" not found.`, "❌ Error", 5);
    return;
  }

  const letterToIndex = l => String(l).toUpperCase().split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0) - 1;

  // Build a set of keys from the destination sheet
  const destKeyIndex = letterToIndex(compareKeyColumn);
  const lastDestRow = destinationSheet.getLastRow();
  let destinationKeySet = new Set();
  console.log(`Reading destination keys from sheet: "${compareSheetName}", column: ${compareKeyColumn} (index ${destKeyIndex}), last row: ${lastDestRow}`);

  if (lastDestRow >= 1) {
    const destKeyValuesRaw = destinationSheet
      .getRange(1, destKeyIndex + 1, lastDestRow, 1)
      .getValues()
      .flat();
    
    console.log("Raw destination keys (first 20):", JSON.stringify(destKeyValuesRaw.slice(0,20)));

    const destKeyValuesProcessed = destKeyValuesRaw
      .map(v => (v === null || v === undefined) ? '' : String(v).trim()) // Convert to string, trim, handle null/undefined
      .filter(v => v !== ''); // Filter out empty strings AFTER trimming
    
    destinationKeySet = new Set(destKeyValuesProcessed);
    console.log(`Processed ${destinationKeySet.size} unique, non-empty, trimmed keys from destination sheet.`);
    if (destinationKeySet.size < 20 && destinationKeySet.size > 0) { // Log small sets for inspection
        console.log("Sample of destination keys:", JSON.stringify(Array.from(destinationKeySet)));
    }
  }

  if (destinationKeySet.size === 0) {
    const message = `No valid keys found in column '${compareKeyColumn}' of destination sheet '${compareSheetName}'. No rows will be deleted from '${sheetName}'.`;
    SpreadsheetApp.getActive()?.toast(message, "ℹ️ Notice", 7);
    console.warn(message);
    return;
  }

  const filterMap = {};
  for (const [col, cond] of Object.entries(filter || {})) {
    filterMap[letterToIndex(col)] = cond;
  }

  const lastSourceRow = sourceSheet.getLastRow();
  if (lastSourceRow < startRow) {
    console.log(`No rows to process in source sheet "${sheetName}" (lastSourceRow < startRow).`);
    return;
  }

  const numRowsToEvaluate = lastSourceRow - startRow + 1;
  const numSourceCols = sourceSheet.getLastColumn();
  if (numRowsToEvaluate <= 0 || numSourceCols <= 0) {
    console.log(`No data rows or columns to process in source sheet "${sheetName}".`);
    return;
  }
  const sourceData = sourceSheet.getRange(startRow, 1, numRowsToEvaluate, numSourceCols).getValues();
  console.log(`Fetched ${sourceData.length} rows from source sheet "${sheetName}".`);

  const rowsToDelete = [];
  const sourceKeyIndex = letterToIndex(keyColumn);
  console.log(`Using source key from column: ${keyColumn} (index ${sourceKeyIndex})`);


  sourceData.forEach((row, i) => {
    const currentRowNumberInSheet = startRow + i;

    const rawSourceKeyValue = row[sourceKeyIndex];
    // ALWAYS convert to string and trim for comparison, same as destination keys
    const sourceKeyValue = (rawSourceKeyValue === null || rawSourceKeyValue === undefined) ? '' : String(rawSourceKeyValue).trim();

    if (i < 5 || (rowsToDelete.length > 0 && i < rowsToDelete[rowsToDelete.length-1] + 5) ) { // Log first few and around potential deletes
         console.log(`--- Processing Row ${currentRowNumberInSheet} (index ${i}) ---`);
         console.log(`Raw source key value: [${rawSourceKeyValue}] (Type: ${typeof rawSourceKeyValue})`);
         console.log(`Processed source key value: [${sourceKeyValue}] (Type: ${typeof sourceKeyValue})`);
    }


    if (sourceKeyValue === '') {
      if (i < 5) console.log(`Row ${currentRowNumberInSheet}: Empty processed source key. Skipping.`);
      return;
    }

    const keyFoundInDestination = destinationKeySet.has(sourceKeyValue);
    if (i < 5 || (rowsToDelete.length > 0 && i < rowsToDelete[rowsToDelete.length-1] + 5) || !keyFoundInDestination ) {
        console.log(`Row ${currentRowNumberInSheet}: Key [${sourceKeyValue}] found in destination set? ${keyFoundInDestination}`);
    }


    if (!keyFoundInDestination) {
      // This is the crucial check. If this is false, it should NOT be added to rowsToDelete.
      if (i < 5) console.log(`Row ${currentRowNumberInSheet}: Key NOT in destination. SKIPPING deletion processing.`);
      return;
    }

    // Key WAS found in destination, now check filters
    console.log(`Row ${currentRowNumberInSheet}: Key [${sourceKeyValue}] MATCHED in destination. Proceeding to check filters.`);
    let passesFilter = true;
    if (Object.keys(filterMap).length > 0) {
      passesFilter = Object.entries(filterMap).every(([colIdx, cond]) => {
        const cellValue = row[colIdx]; // Use raw cell value for filter comparison unless filter implies type change
        const filterResult = (() => {
            if (Array.isArray(cond)) return cond.includes(cellValue);
            if (typeof cond === 'string' && cond.startsWith('!')) {
                const target = cond.slice(1);
                return target === '' ? (cellValue !== '' && cellValue !== null) : cellValue !== target;
            }
            return cellValue === cond;
        })();
        if (i < 5 || !filterResult) { // Log if first few or if filter fails
            console.log(`Row ${currentRowNumberInSheet}: Filter: ColIdx ${colIdx}, Condition [${cond}], CellValue [${cellValue}], Passes? ${filterResult}`);
        }
        return filterResult;
      });
    }

    if (passesFilter) {
      console.log(`Row ${currentRowNumberInSheet}: Key [${sourceKeyValue}] PASSED filters (or no filters). Marking for deletion.`);
      rowsToDelete.push(currentRowNumberInSheet);
    } else {
      console.log(`Row ${currentRowNumberInSheet}: Key [${sourceKeyValue}] FAILED filters. Not marking for deletion.`);
    }
  });

  console.log(`Found ${rowsToDelete.length} rows to delete. Rows: ${JSON.stringify(rowsToDelete.slice(0,20))}`);

  if (rowsToDelete.length === 0) {
    SpreadsheetApp.getActive()?.toast(`No rows in "${sheetName}" matched keys and filters.`, 'ℹ️ Nothing deleted', 7);
    console.log(`No rows in "${sheetName}" matched keys and met filter criteria for deletion.`);
    return;
  }

  const batches = [];
  rowsToDelete.sort((a, b) => a - b);
  let batchStart = rowsToDelete[0];
  let batchCount = 1;

  for (let i = 1; i < rowsToDelete.length; i++) {
    if (rowsToDelete[i] === rowsToDelete[i - 1] + 1) {
      batchCount++;
    } else {
      batches.push({ start: batchStart, count: batchCount });
      batchStart = rowsToDelete[i];
      batchCount = 1;
    }
  }
  batches.push({ start: batchStart, count: batchCount });
  console.log(`Deleting in ${batches.length} batches.`);

  batches.reverse().forEach(batch => {
    console.log(`Deleting ${batch.count} rows starting at row ${batch.start}`);
    sourceSheet.deleteRows(batch.start, batch.count);
  });

  SpreadsheetApp.getActive()?.toast(
    `${rowsToDelete.length} row(s) deleted from sheet "${sheetName}".`,
    '✅ Secure Deletion Completed',
    7
  );
  console.log(`${rowsToDelete.length} row(s) successfully deleted from sheet "${sheetName}".`);
}