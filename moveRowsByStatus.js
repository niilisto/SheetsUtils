/**
 * WORK BLOCK: Moves rows from one sheet to another based on status criteria.
 * Accepts only the new parameter format:
 * {
 *   sourceSheetInfo: { ssId: string, sName: string, headerRowEnd: number },
 *   destSheetInfo:   { ssId: string, sName: string },
 *   columnConfig:    { ticketCol: string, statusCol: string },
 *   moveCriteria:    { checkMode: 'CHECK_IN'|'CHECK_AGAINST', targetStatuses: string[] },
 *   sourceText:      string
 * }
 *
 * @param {object} options Configuration object in the new format.
 * @param {{ ssId: string, sName: string, headerRowEnd: number }} options.sourceSheetInfo
 *        Information about the source sheet: spreadsheet ID, sheet name, and header row index.
 * @param {{ ssId: string, sName: string }} options.destSheetInfo
 *        Information about the destination sheet: spreadsheet ID and sheet name.
 * @param {{ ticketCol: string, statusCol: string }} options.columnConfig
 *        Column letters for ticket ID and status (e.g., 'A', 'L').
 * @param {{ checkMode: 'CHECK_IN'|'CHECK_AGAINST', targetStatuses: string[] }} options.moveCriteria
 *        Criteria controlling which statuses trigger row moves.
 * @param {string} options.sourceText
 *        Text identifier added to each moved row.
 * @throws {Error} If column letters are invalid or specified sheets are not found.
 */
function moveRowsByStatus(options) {
  const letterToCol = (letter) => /^[A-Za-z]+$/.test(letter)
    ? [...letter.toUpperCase()].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0)
    : -1;

  // Destructure new-format parameters
  const { sourceSheetInfo, destSheetInfo, columnConfig, moveCriteria, sourceText } = options;

  // Convert column letters to indices
  const ticketColNum = letterToCol(columnConfig.ticketCol);
  const statusColNum = letterToCol(columnConfig.statusCol);
  if (ticketColNum === -1 || statusColNum === -1) {
    throw new Error(
      `Invalid column letters: ticketCol='${columnConfig.ticketCol}', statusCol='${columnConfig.statusCol}'.`
    );
  }

  // Open spreadsheet and retrieve sheets
  const ss = SpreadsheetApp.openById(sourceSheetInfo.ssId);
  const sourceSheet = ss.getSheetByName(sourceSheetInfo.sName);
  const destSheet   = ss.getSheetByName(destSheetInfo.sName);
  if (!sourceSheet || !destSheet) {
    throw new Error(
      `Source sheet "${sourceSheetInfo.sName}" or destination sheet "${destSheetInfo.sName}" not found.`
    );
  }

  // Determine data range
  const startRow = sourceSheetInfo.headerRowEnd + 1;
  const lastRow  = sourceSheet.getLastRow();
  if (lastRow < startRow) {
    Logger.log('No data rows to process.');
    return;
  }
  const numRows = lastRow - startRow + 1;
  const data    = sourceSheet.getRange(startRow, 1, numRows, sourceSheet.getMaxColumns()).getValues();

  const toMove       = [];
  const rowsToDelete = [];
  const timestamp    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  // Process each row
  data.forEach((row, index) => {
    const status   = String(row[statusColNum - 1] || '').trim();
    const ticketId = row[ticketColNum - 1];
    const matches  = moveCriteria.targetStatuses.includes(status);
    const shouldMove = (moveCriteria.checkMode === 'CHECK_IN' && matches)
                     || (moveCriteria.checkMode === 'CHECK_AGAINST' && !matches);
    if (shouldMove) {
      toMove.push([ticketId, status, timestamp, sourceText]);
      rowsToDelete.push(startRow + index);
    }
  });

  // Write moved rows to destination
  if (toMove.length > 0) {
    Logger.log(`Moving ${toMove.length} rows to "${destSheetInfo.sName}".`);
    destSheet.getRange(destSheet.getLastRow() + 1, 1, toMove.length, toMove[0].length).setValues(toMove);
  } else {
    Logger.log('No rows met move criteria.');
  }

  // Delete rows from source in reverse order
  if (rowsToDelete.length > 0) {
    Logger.log(`Deleting ${rowsToDelete.length} rows from "${sourceSheetInfo.sName}".`);
    rowsToDelete.reverse().forEach(rowNum => sourceSheet.deleteRow(rowNum));
  }

  Logger.log('Row transfer process completed.');
}

/**
 * Example trigger function
 */
function runMoveRowsByStatus() {
  moveRowsByStatus({
    sourceSheetInfo: {
      ssId: '1XUWI5WV7_EIrSSDM4yLsKlZWC7cfQGkQbMrLwn--iJE',
      sName: 'Main',
      headerRowEnd: 1
    },
    destSheetInfo: {
      ssId: '1XUWI5WV7_EIrSSDM4yLsKlZWC7cfQGkQbMrLwn--iJE',
      sName: 'Removed'
    },
    columnConfig: {
      ticketCol: 'A',
      statusCol: 'L'
    },
    moveCriteria: {
      checkMode: 'CHECK_IN',
      targetStatuses: [
        'S4 - COLETA REVERSA',
        'S5.0 - ENTRADA ESTOQUE',
        'S6 - REPARO LABORATORIO',
        'Resolvido',
        'Cancelado',
        'S1.ND.0 - ND',
        'S1.ND.1 - NF ENTRADA ND',
        'Fechado',
        'S1.ND.2 - AGUARDANDO ENTRADA - (Outros Distribuidores)',
        'Em atendimento',
        'Aguardando PAC reverso',
        'S4.erro- PEDIDO DE COLETA MANUAL'
      ]
    },
    sourceText: 'API MOVIDESK'
  });
}
