/**
 * Fills the inverter family column in the main sheet based on a model-to-family mapping from the database sheet.
 *
 * @param {Object} options - Configuration options.
 * @param {string} [options.mainSheetName="Envios"] - Name of the main sheet to update.
 * @param {string} [options.dbSheetName="db"] - Name of the database sheet containing model-family mapping.
 * @param {string} [options.modelColLetter="D"] - Column letter in the main sheet where the inverter model is located.
 * @param {string} [options.familyColLetter="M"] - Column letter in the main sheet where the inverter family should be written.
 * @param {string} [options.dbModelColLetter="A"] - Column letter in the database sheet for the inverter model.
 * @param {string} [options.dbFamilyColLetter="B"] - Column letter in the database sheet for the inverter family.
 * @param {number} [options.startRow=2] - Row number in the main sheet to start processing (1-based).
 *
 * @example
 * getInverterFamily({
 *   mainSheetName: "Envios2",
 *   dbSheetName: "db2",
 *   modelColLetter: 'E',
 *   familyColLetter: 'N',
 *   dbModelColLetter: 'C',
 *   dbFamilyColLetter: 'D',
 *   startRow: 3
 * });
 */
// Helper to convert column letter to index (A=1, B=2, ...)
function letterCol(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col *= 26;
    col += letter.charCodeAt(i) - 64;
  }
  return col;
}

function getInverterFamily({
  mainSheetName,
  dbSheetName,
  modelColLetter,
  familyColLetter,
  dbModelColLetter,
  dbFamilyColLetter,
  startRow
} = {}) {
  const ss = SpreadsheetApp.getActive();
  const mainSheet = ss.getSheetByName(mainSheetName);
  const dbSheet = ss.getSheetByName(dbSheetName);

  // Convert letters to indices (only once)
  const MODEL_COL_INDEX = letterCol(modelColLetter);
  const FAMILY_COL_INDEX = letterCol(familyColLetter);
  const DB_MODEL_COL_INDEX = letterCol(dbModelColLetter);
  const DB_FAMILY_COL_INDEX = letterCol(dbFamilyColLetter);

  // Read all data
  const mainValues = mainSheet.getDataRange().getValues();
  const dbValues = dbSheet.getDataRange().getValues();

  // Create lookup object model → family
  const inverterMap = {};
  dbValues.slice(1).forEach(row => {
    const model = row[DB_MODEL_COL_INDEX - 1];
    const family = row[DB_FAMILY_COL_INDEX - 1];
    if (model) inverterMap[model] = family;
  });

  // Prepare result array (family) for batch write
  const output = mainValues.slice(startRow - 1).map(row => {
    const model = row[MODEL_COL_INDEX - 1];
    return [model && inverterMap[model] ? inverterMap[model] : ''];
  });

  // Write all at once to the family column, starting from startRow
  if (output.length) {
    mainSheet
      .getRange(startRow, FAMILY_COL_INDEX, output.length, 1)
      .setValues(output);
  }

  Logger.log("Inverter family retrieval completed!");
}