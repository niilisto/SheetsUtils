// function timestampTriggerOnChange() {
//   convertColumnDateFormat(2, "Main", [{ trigger: 'H', target: 'I' }], "dd/MM/yyyy", "MM/yyyy", 'America/Sao_Paulo');
// }

/**
 * Reads dates/times from a 'trigger' column in a specific format,
 * converts them to a new format, and writes them to the 'target' column.
 * It iterates through rows starting from 'fromRow'. If a cell in the 'trigger' column
 * contains a valid date/time in the 'sourceTimestampFormat', it's converted to
 * the 'targetTimestampFormat' and written to the corresponding cell in the 'target' column.
 * This function is standalone and does NOT depend on an 'e' edit event object.
 *
 * @param {number} fromRow - The starting row for processing.
 * @param {string} sheetName - The name of the sheet.
 * @param {Array<{ trigger: string, target: string }>} columnPairs - Pairs of trigger/target columns.
 * @param {string} sourceTimestampFormat - The expected date/time format in the 'trigger' column if the value is a string.
 * @param {string} targetTimestampFormat - The desired date/time format for the 'target' column.
 * @param {string} timeZone - The timezone for parsing and formatting dates/times.
 */
function convertColumnDateFormat(fromRow, sheetName, columnPairs, sourceTimestampFormat, targetTimestampFormat, timeZone) {
  if (!fromRow || !sheetName || !columnPairs || !sourceTimestampFormat || !targetTimestampFormat || !timeZone) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return;
  }

  const letterToCol = l =>
    l.toUpperCase().split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);

  const lastDataRowInSheet = sheet.getLastRow();
  const numRowsToProcess = lastDataRowInSheet - fromRow + 1;

  if (numRowsToProcess <= 0) {
    return;
  }

  columnPairs.forEach(pair => {
    const triggerColIndex = letterToCol(pair.trigger);
    const targetColIndex = letterToCol(pair.target);

    const triggerColumnAllValues = sheet.getRange(fromRow, triggerColIndex, numRowsToProcess, 1).getValues();
    const targetColumnNewValues = sheet.getRange(fromRow, targetColIndex, numRowsToProcess, 1).getValues();

    let valuesWereChanged = false;

    for (let i = 0; i < numRowsToProcess; i++) {
      const valueInTriggerCellOfCurrentRow = triggerColumnAllValues[i][0];
      let dateObject = null;

      if (valueInTriggerCellOfCurrentRow instanceof Date) {
        dateObject = valueInTriggerCellOfCurrentRow;
      } else if (typeof valueInTriggerCellOfCurrentRow === 'object' && valueInTriggerCellOfCurrentRow !== null && valueInTriggerCellOfCurrentRow.toString().includes('GMT')) {
        try {
          dateObject = new Date(valueInTriggerCellOfCurrentRow.toString());
          if (isNaN(dateObject.getTime())) {
            dateObject = null;
          }
        } catch (e) {
          dateObject = null;
        }
      } else if (valueInTriggerCellOfCurrentRow && typeof valueInTriggerCellOfCurrentRow === 'string' && valueInTriggerCellOfCurrentRow.trim() !== "") {
        try {
          dateObject = Utilities.parseDate(valueInTriggerCellOfCurrentRow, timeZone, sourceTimestampFormat);
        } catch (parseError) {
          dateObject = null;
        }
      }

      if (dateObject) {
        try {
          const formattedDateForTarget = Utilities.formatDate(dateObject, timeZone, targetTimestampFormat);
          if (targetColumnNewValues[i][0] !== formattedDateForTarget) {
            targetColumnNewValues[i][0] = formattedDateForTarget;
            valuesWereChanged = true;
          }
        } catch (formatError) {

        }
      }
    }

    if (valuesWereChanged) {
      sheet.getRange(fromRow, targetColIndex, numRowsToProcess, 1).setValues(targetColumnNewValues);
    }
  });
}