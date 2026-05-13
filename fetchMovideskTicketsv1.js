// /**
//  * Fetches Movidesk tickets and appends new rows to a Google Sheet,
//  * using filterGroups with internal joiners (`joiner`) and external joiners (`extJoiner`).
//  *
//  * @param {Object} options
//  * @param {string} options.token           -- Movidesk API token
//  * @param {string} options.baseendpoint         -- Base API endpoint URL
//  * @param {Array<{col: number, fv: string|number, isDate?: boolean}>} options.colMapping
//  * @param {Array<{
//  *   type: 'status'|'cfv',
//  *   rules: Array<{
//  *     field?: string,
//  *     operator?: string,
//  *     value?: string,
//  *     customFieldId?: number,
//  *     itemValue?: string,
//  *     joiner?: 'and'|'or'
//  *   }>,
//  *   extJoiner?: 'and'|'or'
//  * }>} options.filterGroups        -- Array of filter groups
//  * @param {{ ssUrl: string, sName: string, headerRowEnd: number }} options.sheetInfo
//  */
// function fetchMovideskTicketsBeta(options) {
//   const { token, baseendpoint, colMapping, filterGroups, sheetInfo } = options;

//   // 1) Build OData filter string
//   const filterString = buildODataFilterString(filterGroups);
//   const filterParam  = encodeURIComponent(filterString);

//   // 2) Build $select param
//   const selectFields = colMapping
//     .filter(m => typeof m.fv === 'string')
//     .map(m => m.fv)
//     .join(',');
//   const selectParam = selectFields ? `&$select=${encodeURIComponent(selectFields)}` : '';

//   // 3) Build $expand for numeric custom fields
//   const cfIds = [...new Set(
//     colMapping
//       .filter(m => typeof m.fv === 'number')
//       .map(m => m.fv)
//   )];
//   const expandParam = cfIds.length
//     ? `&$expand=${encodeURIComponent(
//         `customFieldValues($filter=${cfIds.map(id => `customFieldId eq ${id}`).join(' or ')};`
//         + `$select=value,customFieldId,items;`
//         + `$expand=items($select=customFieldItem,storageFileGuid))`
//       )}`
//     : '';

//   const endpoint = `${baseendpoint}?token=${token}` +
//                  `&$filter=${filterParam}` +
//                  selectParam +
//                  expandParam;

//   Logger.log('API URL: ' + endpoint);

//   try {
//     // 4) Fetch with exponential backoff
//     let resp, attempts = 0;
//     const maxAttempts = 5, baseDelay = 500;
//     while (attempts < maxAttempts) {
//       try {
//         resp = UrlFetchApp.fetch(endpoint, {
//           method: 'GET',
//           headers: { Accept: 'application/json' },
//           muteHttpExceptions: true
//         });
//         if (resp.getResponseCode() === 200) break;
//         throw new Error(`API returned ${resp.getResponseCode()}`);
//       } catch (e) {
//         attempts++;
//         if (attempts >= maxAttempts) throw e;
//         Utilities.sleep(baseDelay * Math.pow(2, attempts));
//       }
//     }

//     const tickets = JSON.parse(resp.getContentText());
//     if (!Array.isArray(tickets) || tickets.length === 0) {
//       Logger.log('No tickets found for the given criteria.');
//       return;
//     }

//     // 5) Open sheet and dedupe
//     const ss = SpreadsheetApp.openByUrl(sheetInfo.ssUrl);
//     const sheet = ss.getSheetByName(sheetInfo.sName);
//     const startRow = sheetInfo.headerRowEnd + 1;
//     const lastRow  = sheet.getLastRow();
//     const existingIds = new Set(
//       sheet.getRange(startRow, 1, Math.max(0, lastRow - startRow + 1), 1)
//            .getValues().flat().map(String)
//     );
//     const newTickets = tickets.filter(t => {
//       const id = t.id || t.Id;
//       return id != null && !existingIds.has(id.toString());
//     });
//     if (newTickets.length === 0) {
//       Logger.log('No new tickets to add.');
//       return;
//     }

//     // 6) Map and append rows
//     const numCols = colMapping.length ? Math.max(...colMapping.map(m => m.col)) + 1 : 0;
//     const rows = newTickets.map(ticket => {
//       const row = Array(numCols).fill('');
//       const cfMap = new Map();
//       (ticket.customFieldValues || []).forEach(cf => {
//         const val = cf.items?.length
//           ? cf.items.map(i => i.customFieldItem || '').join(', ')
//           : cf.value;
//         cfMap.set(cf.customFieldId, val);
//       });
//       colMapping.forEach(m => {
//         let v = (typeof m.fv === 'number')
//           ? cfMap.get(m.fv) || ''
//           : ticket[m.fv] || '';
//         if (m.isDate && v) {
//           const d = new Date(v);
//           if (!isNaN(d)) v = d;
//         }
//         row[m.col] = v;
//       });
//       return row;
//     });

//     sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, numCols)
//          .setValues(rows);

//     Logger.log(`${rows.length} new tickets added successfully.`);
//   }
//   catch (e) {
//     Logger.log('Error: ' + e);
//     throw e;
//   }
// }


// /**
//  * Builds an OData filter string using internal and external joiners.
//  *
//  * @param {Array<{
//  *   type: 'status'|'cfv',
//  *   rules: Array<{
//  *     field?: string,
//  *     operator?: string,
//  *     value?: string,
//  *     customFieldId?: number,
//  *     itemValue?: string,
//  *     joiner?: 'and'|'or'
//  *   }>,
//  *   extJoiner?: 'and'|'or'
//  * }>} filterGroups
//  * @returns {string}
//  */
// function buildODataFilterString(filterGroups) {
//   const sanitize = v =>
//     (typeof v === 'string' ? v.replace(/'/g, "''") : String(v));

//   const validOps = ['eq','ne','lt','gt','ge','le','startsWith'];

//   // Build each group: ( A joiner B joiner C )
//   function buildGroup(group) {
//     const parts = group.rules.map((r, i, arr) => {
//       let clause;
//       if (group.type === 'status') {
//         if (!validOps.includes(r.operator))
//           throw new Error(`Invalid operator: ${r.operator}`);
//         clause = `${r.operator}(${r.field},'${sanitize(r.value)}')`;
//       } else if (group.type === 'cfv') {
//         clause = `customFieldValues/any(c: c/customFieldId eq ${parseInt(r.customFieldId)} and c/items/any(i: i/customFieldItem eq '${sanitize(r.itemValue)}'))`;
//       } else {
//         throw new Error(`Unknown type: ${group.type}`);
//       }
//       if (r.joiner && i < arr.length - 1) {
//         clause += ` ${r.joiner.toLowerCase() === 'or' ? 'or' : 'and'} `;
//       }
//       return clause;
//     });
//     return `(${parts.join('')})`;
//   }

//   // Combine groups with extJoiner from the previous group
//   let result = '';
//   filterGroups.forEach((grp, idx) => {
//     const clause = buildGroup(grp);
//     if (idx === 0) {
//       result = clause;
//     } else {
//       const prev = filterGroups[idx - 1];
//       const joinOp = (prev.extJoiner && prev.extJoiner.toLowerCase() === 'or') ? 'or' : 'and';
//       result = `${result} ${joinOp} ${clause}`;
//     }
//   });

//   return result;
// }