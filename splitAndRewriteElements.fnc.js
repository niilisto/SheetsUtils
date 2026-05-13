/**
 * Copies data from an input sheet, expands comma-separated items in target columns,
 * removes numeric prefixes (e.g., "1 MODEL") or suffixes like "(Q1)", "(Q2)", etc.,
 * and writes the cleaned, expanded data into an output sheet.
 *
 * Assumes each item has either a numeric prefix or a (Qn) suffix — never both.
 *
 * @param {string} inputTab - Name of the input sheet (source data).
 * @param {string} outputTab - Name of the output sheet (destination for cleaned data).
 * @param {Array<string>} header - Header row to write as the first row in the output sheet.
 * @param {Array<string>} writeCols - Column letters to copy from the input sheet.
 * @param {Array<string>} targetCols - Subset of `writeCols` where values may contain comma-separated items to expand.
 *
 * @example
 * // Input sheet "Input":
 * // | Code | Equipment           |
 * // |------|---------------------|
 * // | 001  | 1 MODEL, 2 MODEL    |
 *
 * splitAndRewriteElements(
 *   "Input",
 *   "Output",
 *   ["Code", "Equipment"],
 *   ["A", "B"],       // Columns A (Code) and B (Equipment)
 *   ["B"]             // Only "Equipment" column will be expanded
 * );
 *
 * // Output sheet "Output":
 * // | Code | Equipment |
 * // |------|-----------|
 * // | 001  | MODEL     |
 * // | 001  | MODEL     |
 *
 * @example
 * // With (Qn) suffixes:
 * // Input:
 * // | Code | Equipment                           |
 * // |------|-------------------------------------|
 * // | 002  | MODEL (Q1), MODEL (Q2), ITEM        |
 *
 * // Output:
 * // | Code | Equipment |
 * // |------|-----------|
 * // | 002  | MODEL     |
 * // | 002  | MODEL     |
 * // | 002  | ITEM      |
 *
 * @note
 * This function removes:
 *  - Leading numeric prefixes like "1 ", "2 ", etc. (e.g., "1 MODEL" → "MODEL")
 *  - Trailing suffixes in the format "(Qn)" (e.g., "MODEL (Q1)" → "MODEL")
 *
 * It does **not** remove other suffixes such as "(G1)", "(Extra)", etc.
 */

function splitAndRewriteElements(inputTab, outputTab, header, writeCols, targetCols) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = sheet.getSheetByName(inputTab);
  const outputSheet = sheet.getSheetByName(outputTab);

  const letterToCol = l =>
    l.toUpperCase().split('').reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0) - 1;

  const data = inputSheet.getDataRange().getValues();
  const newRows = [header];

  const writeIndexes = writeCols.map(letterToCol);
  const targetIndexes = new Set(targetCols.map(letterToCol));

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const values = writeIndexes.map(index => row[index]);

    // Find the indexes of columns that need to be expanded
    const expandedIndexes = writeIndexes // Skip empty
      .map((index, i) => (targetIndexes.has(index) ? i : -1))
      .filter(i => i !== -1);
    // If there's no colmuns to expand, just add the rows as is
    if (expandedIndexes.length === 0) {
      newRows.push(values);
      continue;
    }

    const targets = expandedIndexes
      .map(i => (values[i] || '').split(',').map(s => {
        let cleaned = s.trim();
        if (/^\d+\s+/.test(cleaned)) {
          cleaned = cleaned.replace(/^\d+\s*/, ''); // remove prefix (e.g. "1 ", "2 ", etc.)
        } else if (/\(Q\d+\)$/.test(cleaned)) {
          cleaned = cleaned.replace(/\s*\(Q\d+\)\s*$/, ''); // remove suffix (e.g. "(Q1)", "(Q2)", etc.)
        }
        return cleaned;
      }).filter(Boolean))
      .flat();

    for (let item of targets) {
      const newRow = [...values];
      for (let i of expandedIndexes) newRow[i] = item;
      newRows.push(newRow);
    }
  }

  outputSheet.clearContents();
  outputSheet.getRange(1, 1, newRows.length, newRows[0].length).setValues(newRows);
}
