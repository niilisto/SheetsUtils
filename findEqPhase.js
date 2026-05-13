/**
 * Classifies the inverter phase as 'THREE-PHASE' or 'SINGLE-PHASE' based on whether
 * the equipment name contains any of the defined three-phase prefixes, regardless of order or position.
 * Writes the result to the specified column in the given sheet.
 *
 * @param {string} sheetName - Name of the sheet.
 * @param {number} startRow - Starting row (1-based).
 * @param {string} equipmentColLetter - Column letter of the equipment.
 * @param {string} resultColLetter - Column letter for the result.
 */
function findEqPhase(sheetName, startRow, equipmentColLetter, resultColLetter) {
  function columnLetterToIndex(letter) {
    if (!/^[A-Za-z]+$/.test(letter)) return -1;
    return letter.toUpperCase().split('')
      .reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
  }

  function classifyPhase(equipmentName) {
    if (!equipmentName) return '';
    const name = equipmentName.toUpperCase();
    const threePhasePrefixes = ['MAC', 'MAX', 'MID', 'MOD', 'TL3'];
    return threePhasePrefixes.some(prefix => name.includes(prefix))
      ? 'THREE-PHASE'
      : 'SINGLE-PHASE';
  }

  const eqCol = columnLetterToIndex(equipmentColLetter);
  const resCol = columnLetterToIndex(resultColLetter);
  if (eqCol < 1 || resCol < 1) {
    throw new Error(`Invalid column: equipment=${equipmentColLetter}, result=${resultColLetter}`);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

  const lastRow = sheet.getLastRow();
  const numRows = lastRow - startRow + 1;
  if (numRows <= 0) return;

  const eqValues = sheet.getRange(startRow, eqCol, numRows, 1).getValues();

  const output = eqValues.map(row => [classifyPhase((row[0] || '').trim())]);

  sheet.getRange(startRow, resCol, numRows, 1).setValues(output);
}