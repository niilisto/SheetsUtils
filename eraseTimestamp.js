/**
 * @name eraseTimestamp
 * @version 1.2 (2025-06-06)
 *
 * Clears target columns when trigger columns are cleared (i.e., previously filled and now empty).
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object.
 * @param {number} fromRow - Starting row (inclusive).
 * @param {string} sheetName - Target sheet name.
 * @param {Array<{
 *   trigger: string|string[],
 *   target: string|string[]
 * }>} columnGroups - Pairs of trigger and target columns.
 */
function eraseTimestamp(e, fromRow, sheetName, columnGroups) {
  if (!e || !e.range) return;
  const sheet = e.source.getSheetByName(sheetName);
  if (!sheet) return;

  const range = e.range;
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  const letterToCol = l =>
    l.toUpperCase().split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);

  const lastCol = sheet.getLastColumn();
  const allRows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();

  const updates = [];

  for (let i = 0; i < numRows; i++) {
    const rowNum = startRow + i;
    if (rowNum < fromRow) continue;

    const rowVals = allRows[i];

    for (let j = 0; j < numCols; j++) {
      const colNum = startCol + j;

      columnGroups.forEach(group => {
        const triggers = Array.isArray(group.trigger) ? group.trigger : [group.trigger];
        const triggerCols = triggers.map(letterToCol);
        if (!triggerCols.includes(colNum)) return;

        const newValue = rowVals[colNum - 1];
        if (newValue === '' || newValue === null) {
          // Limpa os alvos se a trigger estiver vazia AGORA
          const targets = Array.isArray(group.target) ? group.target : [group.target];
          targets.forEach(tgt => {
            const targetCol = letterToCol(tgt);
            updates.push({ row: rowNum, col: targetCol, value: '' });
          });
        }
      });
    }
  }

  updates.forEach(u => {
    sheet.getRange(u.row, u.col).setValue(u.value);
  });
}
