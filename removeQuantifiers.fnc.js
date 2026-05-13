/**
 * Cleans up values in a specific column of a given sheet by removing:
 *  - Leading numeric prefixes (e.g., "1 ITEM" → "ITEM")
 *  - Trailing (Qn)-style suffixes (e.g., "ITEM (Q1)" → "ITEM")
 *
 * The original sheet is modified in-place.
 *
 * @param {string} sheetName - Name of the sheet to modify.
 * @param {string} colLetter - Letter of the column to clean (e.g., "K").
 *
 * @example
 * // Sheet: "Inventory", column "K":
 * // Before: "1 MODEL", "MODEL (Q2)", "TOOL (G1)"
 * // After:  "MODEL",    "MODEL",      "TOOL (G1)"
 *
 * removeQuantifiers("Inventory", "K");
 */
function removeQuantifiers(sheetName, colLetter) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

  const colIndex = colLetter.toUpperCase()
    .split('')
    .reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, colIndex, lastRow - 1, 1); // Skip header row
  const values = range.getValues();

  const cleaned = values.map(([cell]) => {
    let text = (cell || '').toString().trim();

    // Split itens by comma
    const items = text.split(',')
      .map(item => {
        let it = item.trim();
        // Remove numeric prefix
        it = it.replace(/^\d+\s+/, '');
        // Remove (Qn) suffix
        it = it.replace(/\s*\(Q\d+\)$/, '');
        return it;
      });

    return [items.join(', ')];
  });
  
  range.setValues(cleaned);
}