/**
 * Function to classify the inverter phase based on its prefix.
 * @param {string} sheetName Tab name where the function will run.
 * @param {number} startRow First row to process (1-based).
 * @param {string} equipmentColLetter Letter of column with equipment names.
 * @param {string} resultColLetter Letter from the column where the result will be written.
 * @returns "ON-GRID" or "OFF-GRID" based on the equipment name prefix.
 * 
 * Example usage of findEqPhase.
 *
 * This will classify the inverter phase for the sheet "DENTRO DE GARANTIA",
 * starting from row 2, reading equipment names from column B,
 * and writing the result to column O.
 *
 * @example
 * findEqGrid('SHEET NAME', 2, 'B', 'N');
 */

function findEqGrid(sheetName, startRow, equipmentColLetter, resultColLetter) {
  // Converts column letter(s) to a 1-based index (A=1, B=2, ...).
  const letterToCol = l => {
    if (!/^[A-Za-z]+$/.test(l)) return -1;
    return l.toUpperCase().split('')
      .reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);
  };

  const eqCol = letterToCol(equipmentColLetter);
  const resCol = letterToCol(resultColLetter);
  if (eqCol < 1 || resCol < 1) {
    throw new Error(`Invalid column: equipment=${equipmentColLetter}, result=${resultColLetter}`);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Tab "${sheetName}" not found.`);

  const lastRow = sheet.getLastRow();
  const numRows = lastRow - startRow + 1;
  if (numRows <= 0) return;

  const eqRange = sheet.getRange(startRow, eqCol, numRows, 1);
  const resRange = sheet.getRange(startRow, resCol, numRows, 1);
  const eqValues = eqRange.getValues();   // 2D array [ [val], [val], ... ]

  // Build output array
  const output = [];
  for (let i = 0; i < numRows; i++) {
    const equip = ('' + eqValues[i][0]).trim();
    let writeVal = '';

    if (equip !== '') {
      const prefix = equip.substring(0, 3).toUpperCase();
      if (prefix === 'SPF') {
        writeVal = 'OFF-GRID';
      } else {
        writeVal = 'ON-GRID';
      }
    }
    output.push([writeVal]);
  }

  // Write all at once
  resRange.setValues(output);
}