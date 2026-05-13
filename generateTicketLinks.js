/**
 * Converte os números de ticket na coluna em hyperlinks 
 * (rich text) para **todas** as linhas da coluna A.
 *
 * @param {number} fromRow        Linha inicial (inclusive).
 * @param {string} sheetName      Nome da planilha alvo.
 * @param {string} inputColLetter Letra da coluna (padrão "A").
 */
function generateTicketLinks(fromRow, sheetName, inputColLetter = 'A') {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Planilha "${sheetName}" não encontrada.`);
  
  // Converte letra para índice de coluna
  const letterToCol = l => {
    const L = String(l).toUpperCase();
    if (!/^[A-Z]+$/.test(L)) throw new Error(`Coluna inválida: "${l}"`);
    return L.split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);
  };
  const inputCol = letterToCol(inputColLetter);
  
  // Descobre última linha com conteúdo
  const lastRow = sheet.getLastRow();
  if (lastRow < fromRow) return;  // nada a fazer

  // Lê todos os valores da coluna de fromRow até lastRow
  const numRows = lastRow - fromRow + 1;
  const range   = sheet.getRange(fromRow, inputCol, numRows, 1);
  const values  = range.getValues();

  // Para cada célula não vazia e numérica, cria RichText e aplica
  for (let i = 0; i < values.length; i++) {
    const raw = (values[i][0] || '').toString().trim();
    if (!raw || isNaN(raw)) continue;
    
    const ticket = raw;
    const url    = `https://growatt.movidesk.com/Ticket/edit/${ticket}`;
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(ticket)
      .setLinkUrl(url)
      .build();
    
    sheet.getRange(fromRow + i, inputCol)
         .setRichTextValue(richText);
  }
}
