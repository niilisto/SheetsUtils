/**
 * Return the column index of the edited cell
 * @param {Object} e – event object from onEdit trigger
 * @returns {number} – column index (1-based)
 */
function editedCol(e) {
    return e.range.getColumn();
}

/**
 * Convert a column letter (e.g., "A", "B", ..., "Z", "AA", "AB", ...) to a 1-based column index.
 * @param {string} l - The column letter(s) to convert.
 * @returns {number} The corresponding column index (1-based).
 * @throws {Error} If the input is invalid or too long.
 * @example
 * // letterCol("A") returns 1
 * // letterCol("Z") returns 26
 * // letterCol("AA") returns 27        
 */
function letterCol(l) {
    if (typeof l !== 'string' || !/^[A-Z]+$/i.test(l) || l.length > 3) {
        throw new Error('Invalid column letter input');
    }
    return l
        .toUpperCase()
        .split('')
        .reduce((r, c) => r * 26 + c.charCodeAt(0) - 64, 0);
}
