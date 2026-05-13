/**
 * Divide linhas em valores separados por vírgula em uma ou mais colunas-chave,
 * e escreve cada valor individual como sua própria linha na planilha de exportação.
 *
 * @param {Object}   config
 * @param {string}   config.sourceSheetName   – Nome da planilha de origem.
 * @param {string}   config.exportSheetName   – Nome da planilha de destino.
 * @param {number}   config.startRow          – Primeira linha de dados na origem.
 * @param {string[]} config.keyColumns        – Letras das colunas a dividir.
 * @param {string[]} config.copyColumns       – Letras das colunas a copiar (na ordem desejada).
 */
function splitAndExport(config) {
  const {
    sourceSheetName = "@Main",
    exportSheetName = "Export",
    startRow = 2,
    keyColumns = ["H"],
    copyColumns = ["A","B","C", "D", "E", "F", "G", "H", "I", "J"]
  } = config;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(sourceSheetName);
  const dest   = ss.getSheetByName(exportSheetName);
  if (!source || !dest) {
    throw new Error("Cannot find one of the sheets. Check config names.");
  }

  // helper: letter → 1-based index (A=1, B=2, …)
  const letterToCol = l => {
    if (!/^[A-Za-z]+$/.test(l)) return -1;
    return l.toUpperCase().split('')
      .reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);
  };

  // Build arrays of source-col indices
  const keyColIdxs  = keyColumns.map(letterToCol);
  const copyColIdxs = copyColumns.map(letterToCol);

  // Read all data rows
  const lastRow = source.getLastRow();
  const lastCol = source.getLastColumn();
  const numRows = lastRow - startRow + 1;
  if (numRows < 1) return;  // nothing to do
  const data = source
    .getRange(startRow, 1, numRows, lastCol)
    .getValues();

  // Clear old data in Export (below header row)
  const maxRows = dest.getMaxRows();
  const maxCols = dest.getMaxColumns();
  if (maxRows > 1) {
    dest
      .getRange(2, 1, maxRows - 1, maxCols)
      .clearContent();
  }

  // Build newRows array
  const newRows = [];
  data.forEach(row => {
    keyColumns.forEach((colLetter, ki) => {
      const colIdx = keyColIdxs[ki] - 1;
      const cellVal = row[colIdx];
      const raw = (cellVal == null ? "" : cellVal.toString());
      const parts = raw.includes(",")
        ? raw.split(",").map(s => s.trim()).filter(s => s !== "")
        : [ raw.trim() ];

      parts.forEach(part => {
        const out = copyColIdxs.map(ci => {
          if (ci === keyColIdxs[ki]) {
            return part;
          } else {
            return row[ci - 1];
          }
        });
        newRows.push(out);
      });
    });
  });

  // Write to Export, starting at row 2, col 1
  if (newRows.length > 0) {
    dest
      .getRange(2, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }
}
