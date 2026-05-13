/*
 * Writes unique tokens only in rows with more than 4 filled cells,
 * and only in empty cells of the target column, using the format: <prefix>!<timestamp>i<random>.
 * random: 18 characters (digits, uppercase letters, lowercase letters).
 * timestamp: date/time stamp in format ddMMyyyyHHmmssSSS (America/Sao_Paulo timezone).
 * The token is only written if at least one of the specified "checkColumns" has data in the row.
 *
 * Parameters:
 *   sheetName:      Name of the worksheet/tab in the spreadsheet.
 *   startRow:       First row to process (1-based index).
 *   targetColumn:   Column to write tokens into (letter or index).
 *   prefix:         Initial prefix for the token (trimmed).
 *   checkColumns:   Array of columns (letter or index) to test for existence of data.
 */
function writeToken(sheetName, startRow, targetColumn, prefix, checkColumns) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" not found.');

  // Convert column letter to index (1-based)
  function letterToCol(letter) {
    return letter.toUpperCase().split('').reduce(function (result, char) {
      return result * 26 + char.charCodeAt(0) - 64;
    }, 0);
  }

  // Normalize targetColumn to index
  var colIndex = typeof targetColumn === 'string'
    ? letterToCol(targetColumn)
    : targetColumn;

  // Normalize checkColumns to array of indices
  var checkIndices = (checkColumns || []).map(function(col) {
    return (typeof col === 'string') ? letterToCol(col) : col;
  });

  prefix = (prefix || '').trim();

  var lastRow = sheet.getLastRow();
  var numRows = lastRow - startRow + 1;
  if (numRows < 1) return;

  var lastCol = sheet.getLastColumn();
  var allRows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  var targetRange = sheet.getRange(startRow, colIndex, numRows, 1);
  var targetValues = targetRange.getValues();

  var seen = {};
  function genRandom() {
    var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    var result = '';
    for (var i = 0; i < 18; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return result;
  }

  var output = [];
  for (var i = 0; i < numRows; i++) {
    var rowData = allRows[i];
    // count filled cells only in checkColumns
    var filledCount = checkIndices.reduce(function (count, idx) {
      var val = rowData[idx - 1];
      return (val !== '' && val !== null) ? count + 1 : count;
    }, 0);
    var currentValue = targetValues[i][0];
    // Só escreve token se célula estiver vazia e todos os checkColumns estiverem preenchidos
    if (filledCount === checkIndices.length && (currentValue === '' || currentValue === null)) {
      var randomPart;
      do {
        randomPart = genRandom();
      } while (seen[randomPart]);
      seen[randomPart] = true;

      var timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'ddMMyyyyHHmmssSSS');
      var token = prefix + '!' + timestamp + 'i' + randomPart;
      output.push([token]);
    } else {
      output.push([currentValue]);
    }
  }

  // Batch write tokens to the target column
  targetRange.setValues(output);
}