// function findWorkflowTrigger() {
//   findWorkflow('DENTRO DE GARANTIA', 2, 'E', 'O', 'N');
// }

/**
 * Função principal para classificar o fluxo do inversor.
 *
 * @param {string} sheetName Nome da aba onde vai rodar.
 * @param {number} startRow Linha inicial (base 1) para aplicar.
 * @param {string} distributorColLetter Letra da coluna do DISTRIBUIDOR.
 * @param {string} phaseColLetter Letra da coluna onde está a FASE do equipamento.
 * @param {string} resultColLetter Letra da coluna que será escrita com o fluxo.
 */
function findWorkflow(sheetName, startRow, distributorColLetter, phaseColLetter, resultColLetter) {
  // Converte letra da coluna em índice (0-based)
  const letterToCol = l => {
    if (!l) return -1;
    return l.toUpperCase().split('')
      .reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0) - 1;
  };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Aba "${sheetName}" não encontrada`);

  const distCol = letterToCol(distributorColLetter) + 1; // 1-based
  const phaseCol = letterToCol(phaseColLetter) + 1;
  const resCol = letterToCol(resultColLetter) + 1;
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - startRow + 1;
  if (numRows <= 0) return;

  // lê em batch
  const distRange = sheet.getRange(startRow, distCol, numRows, 1);
  const phaseRange = sheet.getRange(startRow, phaseCol, numRows, 1);
  const resRange = sheet.getRange(startRow, resCol, numRows, 1);
  const distValues = distRange.getValues();
  const phaseValues = phaseRange.getValues();
  const resValues = resRange.getValues();

  // lista de distribuidores para REPAIR THE DEFECTIVE
  const repairList = [
    'ALDO', 'FORTLEV', 'BLUSUN', 'BLUESUN', 'SOLAR INOVE/SOLFÁCIL', 'SOLAR INOVE', 'SOLFÁCIL', 'VALMONT', 'BELENERGY'
  ].map(s => s.toUpperCase());

  // prepara output
  const output = [];
  for (let i = 0; i < numRows; i++) {
    const dist = ('' + distValues[i][0]).trim().toUpperCase();
    const phase = ('' + phaseValues[i][0]).trim().toUpperCase();
    const existing = ('' + resValues[i][0]).trim();
    let writeVal = existing;

    // só processa se há fase preenchida e célula de resultado vazia
    if (phase !== '' && existing === '') {
      if (phase === 'THREE-PHASE' && repairList.includes(dist)) {
        writeVal = 'REPAIR THE DEFECTIVE';
      } else {
        writeVal = 'ADVANCED EXCHANGE';
      }
    }
    output.push([writeVal]);
  }

  // grava em batch
  resRange.setValues(output);
}
