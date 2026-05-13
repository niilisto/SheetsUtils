/**
 * @name getTimestamp
 * @version 3.1 (last updated: 2025-06-03, added overwrite param)
 * 
 * Applies a timestamp to target columns when trigger columns are edited in a Google Sheet.
 * Supports multiple triggers, multiple targets, different time zones and timestamp formats, and optional filters.
 * Works for both manual edits and multi-row pastes.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object from Google Sheets.
 * @param {number} fromRow - The starting row (inclusive) to consider for timestamping.
 * @param {string} sheetName - The name of the target sheet.
 * @param {Array<{
 *   trigger: string|string[],              // Trigger column(s) (e.g. 'A' or ['A','B'])
 *   target: string|string[],               // Target column(s) for the timestamp (e.g. 'G' or ['G','H'])
 *   timeZone?: string|string[],            // Time zone(s) for the timestamp (optional, default: 'America/Sao_Paulo')
 *   timestampFormat?: string|string[],     // Timestamp format(s) (optional, default: 'dd/MM/yyyy HH:mm:ss')
 *   filterType?: 'equals'|'notEquals',     // Filter type (optional, case-insensitive)
 *   filterContent?: string[]               // Filter content(s) (optional, case-insensitive)
 * }>} columnGroups - Definitions of trigger/target groups and optional filter/formatting options.
 * @param {boolean} [overwrite=false] - If true, always writes the timestamp (even if cell is not empty).
 *
 * @example
 * // Simple case: when editing column A, writes timestamp to G
 * getTimestamp(e, 2, 'Envios', [
 *   { trigger: 'A', target: 'G' }
 * ]);
 *
 * @example
 * // Multiple triggers and targets with different time zones and formats
 * getTimestamp(e, 2, 'Envios', [
 *   {
 *     trigger: ['O', 'P'],
 *     target: ['J', 'K'],
 *     timeZone: ['America/Sao_Paulo', 'UTC'],
 *     timestampFormat: ['MM/yyyy', 'yyyy'],
 *     filterType: 'notEquals',
 *     filterContent: ['']
 *   }
 * ]);
 *
 * @example
 * // Single trigger writing to multiple targets with different formats
 * getTimestamp(e, 2, 'Envios', [
 *   {
 *     trigger: 'A',
 *     target: ['B', 'J'],
 *     timeZone: ['America/Sao_Paulo'],
 *     timestampFormat: ['dd/MM/yyyy', 'MM/yyyy'],
 *     filterType: 'notEquals',
 *     filterContent: ['']
 *   }
 * ]);
 * 
 * @example
 * // Forçar sobrescrita do carimbo de data/hora
 * getTimestamp(e, 2, 'Envios', [
 *   { trigger: 'A', target: 'G' }
 * ], true);
 */
function getTimestamp(e, fromRow, sheetName, columnGroups, overwrite) {
  if (!e || !e.range) return;
  const sheet = e.source.getSheetByName(sheetName);
  if (!sheet) return;

  const range = e.range;
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  const now = new Date();
  const DEFAULT_TZ = 'America/Sao_Paulo';
  const DEFAULT_FMT = 'dd/MM/yyyy HH:mm:ss';

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
      const newValue = (numRows === 1 && numCols === 1 && e.value !== undefined)
        ? e.value
        : (rowVals[colNum - 1] || '');

      columnGroups.forEach(group => {
        const triggers = Array.isArray(group.trigger) ? group.trigger : [group.trigger];
        const triggerCols = triggers.map(letterToCol);
        if (!triggerCols.includes(colNum)) return;

        if (group.filterContent && group.filterContent.length) {
          const filterType = (group.filterType || '').toLowerCase();
          const newValueNorm = (newValue ?? '').toString().toLowerCase();
          const filterContentNorm = group.filterContent.map(v => (v ?? '').toString().toLowerCase());
          const match = filterContentNorm.includes(newValueNorm);
          if ((filterType === 'equals' && !match) ||
              (filterType === 'notequals' && match)) {
            return;
          }
        }

        const allFilled = triggerCols.every(c => rowVals[c - 1] !== '');
        if (!allFilled) return;

        const targets = Array.isArray(group.target) ? group.target : [group.target];
        const tzArray = Array.isArray(group.timeZone)
          ? group.timeZone
          : group.timeZone
            ? [group.timeZone]
            : [DEFAULT_TZ];
        const fmtArray = Array.isArray(group.timestampFormat)
          ? group.timestampFormat
          : group.timestampFormat
            ? [group.timestampFormat]
            : [DEFAULT_FMT];

        targets.forEach((tgt, idx) => {
          const targetCol = letterToCol(tgt);
          const groupOverwrite = typeof group.overwrite === 'boolean' ? group.overwrite : overwrite;
          if (groupOverwrite === true || rowVals[targetCol - 1] === '') {
            const tz = tzArray[idx] || DEFAULT_TZ;
            const fmt = fmtArray[idx] || DEFAULT_FMT;
            let timestamp = '';
            if (typeof Utilities !== 'undefined' && Utilities.formatDate) {
              timestamp = Utilities.formatDate(now, tz, fmt);
            } else {
              timestamp = now.toLocaleString('en-GB', { timeZone: tz });
            }
            updates.push({ row: rowNum, col: targetCol, value: timestamp });
          }
        });
      });
    }
  }

  updates.forEach(u => {
    sheet.getRange(u.row, u.col).setValue(u.value);
  });
}