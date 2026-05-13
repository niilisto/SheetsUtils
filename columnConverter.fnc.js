/**
 * Converts a column letter (e.g., A, B, AA) to its corresponding 1-based index.
 * @param {string} columnLetter - The column letter to convert.
 * @return {number} The column index (1-based). Returns 0 for invalid input.
 */
function columnConverter(columnLetter) {
  if (!columnLetter || typeof columnLetter !== 'string') return 0;

  let index = 0;
  const upperCaseLetter = columnLetter.toUpperCase();

  for (let i = 0; i < upperCaseLetter.length; i++) {
    const charCode = upperCaseLetter.charCodeAt(i);
    if (charCode < 65 || charCode > 90) return 0; // Invalid character
    index = index * 26 + (charCode - 64);
  }
  return index;
}