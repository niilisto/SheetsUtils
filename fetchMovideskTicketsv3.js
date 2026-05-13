// // =================================================================================
// // 1. CONFIGURATION & GENERIC FUNCTIONS
// // =================================================================================

// // -- Global Settings --
// const apiBaseToken = "80c1fb64-3e4a-48c9-b105-160958e7f5c5";
// const apiBaseUrl = "https://api.movidesk.com/public/v1/tickets";
// const TOP = 1000;


// // -- Date Utility Functions --
// function getODataDateTime(days) {
//   const daysToSubtract = (typeof days === 'number' && !isNaN(days)) ? days : 0;
//   if (daysToSubtract !== days) {
//     Logger.log(`Warning in getODataDateTime: Invalid input '${days}'. Defaulting to ${daysToSubtract} days.`);
//   }
//   const date = new Date();
//   date.setDate(date.getDate() - daysToSubtract);
//   const datePart = Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
//   const timePart = 'T00:00:00Z';
//   return datePart + timePart;
// }

// function formatDateToDDMMYYYY(date) {
//   if (!(date instanceof Date) || isNaN(date)) return '';
//   const day = String(date.getDate()).padStart(2, '0');
//   const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are 0-based
//   const year = date.getFullYear();
//   return `${day}/${month}/${year}`;
// }


// // -- Core Engine --

// /**
//  * Fetches Movidesk tickets. This version is MODIFIED to support an optional
//  * 'postProcessFilter' function to perform client-side filtering when the API
//  * lacks the required filtering capabilities (e.g., date comparison on custom fields).
//  *
//  * @param {object} options The configuration object.
//  * @param {string} options.token The Movidesk API token.
//  * @param {string} options.baseendpoint The base API endpoint.
//  * @param {Array<{col: number, fv: string|number, isDate?: boolean}>} options.colMapping Defines field-to-column mapping.
//  * @param {Array<object>} [options.filterGroups] An array of filter groups for the OData `$filter` clause.
//  * @param {function(object):boolean} [options.postProcessFilter] A function that receives a ticket object and returns true to keep it or false to discard it.
//  * @param {{ssUrl: string, sName: string, headerRowEnd: number}} options.sheetInfo Information about the target Google Sheet.
//  * @throws {Error} If required parameters are missing or if the API call fails.
//  */
// function fetchMovideskTickets(options) {
//   // MODIFIED: Added postProcessFilter to the destructured options
//   const { token, baseendpoint, colMapping, filterGroups = [], sheetInfo, top, orderBy, postProcessFilter } = options;

//   if (!token || !baseendpoint || !sheetInfo || !colMapping) {
//     throw new Error("Missing required options: token, baseendpoint, sheetInfo, and colMapping are mandatory.");
//   }

//   const filterString = buildODataFilterString(filterGroups);
//   const filterParam = filterString ? `$filter=${encodeURIComponent(filterString)}` : '';

//   const selectFields = colMapping.filter(m => typeof m.fv === 'string' && m.fv !== 'todayDate').map(m => m.fv);
//   const selectParam = selectFields.length ? `$select=${encodeURIComponent(selectFields.join(','))}` : '';
  
//   const cfIds = [...new Set(colMapping.filter(m => typeof m.fv === 'number').map(m => m.fv))];
//   const expandParam = cfIds.length
//     ? `$expand=${encodeURIComponent(
//       `customFieldValues($filter=${cfIds.map(id => `customFieldId eq ${id}`).join(' or ')};` +
//       `$select=value,customFieldId,items;$expand=items($select=customFieldItem))`
//     )}` : '';

//   const topParam = top ? `$top=${top}` : '';
//   const orderByParam = orderBy ? `$orderby=${encodeURIComponent(orderBy)}` : '';
  
//   const parts = [`token=${token}`, filterParam, selectParam, expandParam, orderByParam, topParam].filter(Boolean);

//   const todayDateString = formatDateToDDMMYYYY(new Date());
//   let page = 0;
//   const allRows = [];

//   while (true) {
//     const skipParam = top ? `&$skip=${page * top}` : '';
//     const url = `${baseendpoint}?${parts.join('&')}${skipParam}`;
//     Logger.log('Fetching URL: ' + url);

//     const resp = UrlFetchApp.fetch(url, { headers: { Accept: 'application/json' }, muteHttpExceptions: true });
//     if (resp.getResponseCode() !== 200) {
//       const errorMsg = `API Error: ${resp.getResponseCode()} - ${resp.getContentText()}`;
//       Logger.log(errorMsg);
//       throw new Error(errorMsg);
//     }

//     const items = JSON.parse(resp.getContentText());
//     if (!Array.isArray(items) || items.length === 0) break;

//     items.forEach(ticket => {
//       // ----------------------------------------------------------------------------------
//       // NEW CHANGE: Apply the client-side filter here, before processing the row.
//       // If the filter function is provided and returns false, skip this ticket.
//       if (postProcessFilter && !postProcessFilter(ticket)) {
//         return; // This is the core of the workaround.
//       }
//       // ----------------------------------------------------------------------------------

//       const cfMap = new Map();
//       (ticket.customFieldValues || []).forEach(cf => {
//         const value = cf.items?.length ? cf.items.map(i => i.customFieldItem).join(', ') : cf.value;
//         cfMap.set(cf.customFieldId, value);
//       });

//       const row = Array(Math.max(...colMapping.map(m => m.col)) + 1).fill('');
//       colMapping.forEach(m => {
//         let value;
//         if (m.fv === 'todayDate') value = todayDateString;
//         else if (typeof m.fv === 'number') value = cfMap.get(m.fv) || '';
//         else value = ticket[m.fv] || '';
        
//         if (m.isDate && value && !(value instanceof Date)) {
//           const d = new Date(value);
//           if (!isNaN(d)) value = d;
//         }
//         row[m.col] = value;
//       });
//       allRows.push(row);
//     });

//     if (!top) break;
//     page++;
//   }

//   if (allRows.length === 0) {
//     Logger.log('No tickets found matching the final criteria.');
//     return;
//   }

//   const ss = SpreadsheetApp.openByUrl(sheetInfo.ssUrl);
//   const sh = ss.getSheetByName(sheetInfo.sName);
//   const idColConfig = colMapping.find(m => m.fv === 'id');
//   if (!idColConfig) {
//     throw new Error("The 'colMapping' must include an entry for the ticket 'id' to prevent duplicates.");
//   }
//   const idColIndex = idColConfig.col;

//   const startRow = sheetInfo.headerRowEnd + 1;
//   const lastRow = sh.getLastRow();
//   let existingIds = new Set();
  
//   if (lastRow >= startRow) {
//     const numExisting = lastRow - startRow + 1;
//     const vals = sh.getRange(startRow, idColIndex + 1, numExisting, 1).getValues().flat();
//     existingIds = new Set(vals.map(String));
//   }

//   const newRows = allRows.filter(r => !existingIds.has(String(r[idColIndex])));

//   if (newRows.length > 0) {
//     sh.getRange(sh.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
//     Logger.log(`Successfully appended ${newRows.length} new rows.`);
//   } else {
//     Logger.log('No new tickets to add.');
//   }
// }

// /**
//  * Builds a valid OData $filter string. This function is correct but
//  * the API does not support all its features (like casting inside custom fields).
//  */
// function buildODataFilterString(filterGroups = []) {
//   if (!filterGroups || filterGroups.length === 0) return '';
//   const sanitize = v => typeof v === 'string' ? v.replace(/'/g, "''") : String(v);
//   const isDateOrDateTime = (val) => /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}Z)?$/.test(String(val));

//   function buildGroup(group) {
//     const clauses = group.rules.map((rule, idx) => {
//       let expr = '';
//       if (group.type === 'status' || group.type === 'date') {
//           expr = `${rule.operator}(${rule.field},'${sanitize(rule.value)}')`;
//       } else if (group.type === 'cfv') {
//         let subFilter = `c/customFieldId eq ${rule.customFieldId}`;
//         // This part is syntactically correct, but the API doesn't support it.
//         // It's left here for other potential uses where the API might support it.
//         if (isDateOrDateTime(rule.valueContent)) {
//           subFilter += ` and cast(c/value, 'Edm.DateTimeOffset') ${rule.valueOperator} ${rule.valueContent}`;
//         } else if (rule.valueOperator && rule.valueContent !== undefined) {
//            const valueContent = (typeof rule.valueContent === 'number') ? rule.valueContent : `'${sanitize(rule.valueContent)}'`;
//            subFilter += ` and c/value ${rule.valueOperator} ${valueContent}`;
//         }
//         if (rule.itemValue) {
//           subFilter += ` and c/items/any(i: i/customFieldItem eq '${sanitize(rule.itemValue)}')`;
//         }
//         expr = `customFieldValues/any(c: ${subFilter})`;
//       }
//       const joiner = (rule.joiner || 'and').toLowerCase();
//       return (idx === group.rules.length - 1) ? expr : `${expr} ${joiner} `;
//     });
//     return `(${clauses.join('')})`;
//   }

//   return filterGroups.map((group, idx) => {
//     const gStr = buildGroup(group);
//     if (idx === 0) return gStr;
//     const prevJoiner = (filterGroups[idx - 1].extJoiner || 'and').toLowerCase();
//     return ` ${prevJoiner} ${gStr}`;
//   }).join('');
// }


// // =================================================================================
// // 2. BUSINESS LOGIC FUNCTIONS
// // =================================================================================

// /**
//  * REWRITTEN: This function now fetches tickets by status and then filters by date
//  * within the script, working around the API limitation.
//  * @param {number} daysPast The number of days in the past to filter by.
//  */
// function getS1TicketsByDaysPast(daysPast) {
//   // --- WORKAROUND STEP 1: Define the date to compare against ---
//   const cutoffDate = new Date();
//   cutoffDate.setDate(cutoffDate.getDate() - daysPast);
//   cutoffDate.setHours(0, 0, 0, 0); // Normalize to the beginning of the day

//   fetchMovideskTickets({
//     token: apiBaseToken,
//     baseendpoint: apiBaseUrl,
//     top: 1000,
//     sheetInfo: {
//       ssUrl: "https://docs.google.com/spreadsheets/d/1SkGLGBJlUyStxUimCEcczrKN8jx0bnj_SNJ1Lbg_nc8/",
//       sName: "Under Warranty",
//       headerRowEnd: 1,
//     },
//     colMapping: [
//       { col: 0, fv: "id" }, 
//       { col: 1, fv: 152179 }, 
//       { col: 2, fv: 92408 },
//       { col: 3, fv: 102776 }, 
//       { col: 3, fv: 102777 }, 
//       { col: 4, fv: 92834 },
//       { col: 5, fv: 102763 }, 
//       { col: 6, fv: 102753 }, 
//       { col: 7, fv: 102774 },
//       { col: 8, fv: 144016 }, 
//       { col: 9, fv: 107733, isDate: true },
//       { col: 10, fv: 117072, isDate: true }, 
//       { col: 11, fv: 152180, isDate: true },
//       { col: 12, fv: 95991 }, 
//       { col: 17, fv: "todayDate" }, 
//       { col: 18, fv: "status" }
//     ],
//     // --- WORKAROUND STEP 2: Remove the date filter from the API call ---
//     // Only filter by status at the API level.
//     filterGroups: [
//       {
//         type: 'status',
//         rules: [
//           { field: 'status', operator: 'startsWith', value: 'S1', joiner: 'or' },
//           { field: 'status', operator: 'startsWith', value: 'S2', joiner: 'or' },
//           { field: 'status', operator: 'startsWith', value: 'S3', joiner: 'or' },
//           { field: 'status', operator: 'startsWith', value: 'S4' }
//         ]
//       }
//     ],
//     // --- WORKAROUND STEP 3: Provide a function to filter results in the script ---
//     postProcessFilter: (ticket) => {
//       // Find the custom field for the date (ID: 152180)
//       const dateField = (ticket.customFieldValues || []).find(cf => cf.customFieldId === 152180);

//       // If the field doesn't exist on this ticket, or has no value, discard the ticket.
//       if (!dateField || !dateField.value) {
//         return false;
//       }
      
//       // Convert the field's string value to a real JavaScript Date object.
//       const ticketDate = new Date(dateField.value);

//       // Return true if the ticket's date is greater than or equal to our cutoff date.
//       return ticketDate >= cutoffDate;
//     }
//   });
// }

// // Example of how to run it
// function runMyReport() {
//   getS1TicketsByDaysPast(7); // e.g., get tickets from the last 7 days.
// }

// /**
//  * Fetches all tickets with a status starting with S1, S2, S3, or S4.
//  */
// function getS1() {
//   fetchMovideskTickets({
//     token: apiBaseToken,
//     baseendpoint: apiBaseUrl,
//     top: 1000,
//     sheetInfo: {
//       ssUrl: "https://docs.google.com/spreadsheets/d/1SkGLGBJlUyStxUimCEcczrKN8jx0bnj_SNJ1Lbg_nc8/",
//       sName: "Under Warranty",
//       headerRowEnd: 1,
//     },
//     colMapping: [
//       { col: 0, fv: "id" }, 
//       { col: 1, fv: 152179 }, 
//       { col: 2, fv: 92408 },
//       { col: 3, fv: 102776 }, 
//       { col: 3, fv: 102777 }, 
//       { col: 4, fv: 92834 },
//       { col: 5, fv: 102763 }, 
//       { col: 6, fv: 102753 }, 
//       { col: 7, fv: 102774 },
//       { col: 8, fv: 144016 }, 
//       { col: 9, fv: 107733, isDate: true },
//       { col: 10, fv: 117072, isDate: true }, 
//       { col: 11, fv: 152180, isDate: true },
//       { col: 12, fv: 95991 }, 
//       { col: 17, fv: "todayDate" }, 
//       { col: 18, fv: "status" }, 
//       { col: 19, fv: 215335 }
//     ],
//     filterGroups: [{
//       type: 'status',
//       rules: [
//         { field: 'status', operator: 'startsWith', value: 'S1', joiner: 'or' },
//         { field: 'status', operator: 'startsWith', value: 'S2', joiner: 'or' },
//         { field: 'status', operator: 'startsWith', value: 'S3', joiner: 'or' },
//         { field: 'status', operator: 'startsWith', value: 'S4' }
//       ]
//     }]
//   });
// }


// /**
//  * Fetches "Out of Warranty" tickets with statuses starting with F1, F2, F3, or F4.
//  */
// function getF1() {
//   fetchMovideskTickets({
//     token: apiBaseToken,
//     baseendpoint: apiBaseUrl,
//     top: 1000,
//     sheetInfo: {
//       ssUrl: "https://docs.google.com/spreadsheets/d/1SkGLGBJlUyStxUimCEcczrKN8jx0bnj_SNJ1Lbg_nc8/",
//       sName: "Out of Warranty",
//       headerRowEnd: 1,
//     },
//     colMapping: [
//       { col: 0, fv: "id" }, 
//       { col: 1, fv: 152179 }, 
//       { col: 2, fv: 92408 },
//       { col: 3, fv: 102776 }, 
//       { col: 3, fv: 102777 }, 
//       { col: 4, fv: 92834 },
//       { col: 5, fv: 92993 }, 
//       { col: 6, fv: 102753 }, 
//       { col: 7, fv: 102774 },
//       { col: 8, fv: 144016 }, 
//       { col: 10, fv: 95991 }, 
//       { col: 15, fv: "status" }
//     ],
//     filterGroups: [{
//       type: 'status',
//       rules: [
//         { field: 'status', operator: 'startsWith', value: 'F1', joiner: 'or' },
//         { field: 'status', operator: 'startsWith', value: 'F2', joiner: 'or' },
//         { field: 'status', operator: 'startsWith', value: 'F3', joiner: 'or' },
//         { field: 'status', operator: 'startsWith', value: 'F4' }
//       ]
//     }]
//   });
// }

// /**
//  * Fetches tickets with status starting with "F4".
//  */
// function getF4() {
//   fetchMovideskTickets({
//     token: apiBaseToken,
//     baseendpoint: apiBaseUrl,
//     top: 1000,
//     sheetInfo: {
//       ssUrl: "https://docs.google.com/spreadsheets/d/1ZGYY-_Ik_mrt3XRzFbRAiq1jC8HKOOp6E4I-_mpzI8U/",
//       sName: "Main",
//       headerRowEnd: 1,
//     },
//     colMapping: [
//       { col: 0, fv: "id" }, 
//       { col: 1, fv: 152179 }, 
//       { col: 2, fv: 92408 },
//       { col: 3, fv: 102776 }, 
//       { col: 3, fv: 102777 }, 
//       { col: 4, fv: 92834 },
//       { col: 5, fv: 144016 }, 
//       { col: 6, fv: 95991 }
//     ],
//     filterGroups: [{
//       type: 'status',
//       rules: [
//         { field: 'status', operator: 'startsWith', value: 'F4' }
//       ]
//     }]
//   });
// }

// /**
//  * REWRITTEN: Fetches S5 Braspress tickets. This version applies the client-side
//  * filtering workaround for the date comparison, as the API does not support it.
//  */
// function getS5Braspress() {
//   // WORKAROUND STEP 1: Define the date to compare against (last 7 days).
//   const cutoffDate = new Date();
//   cutoffDate.setDate(cutoffDate.getDate() - 7);
//   cutoffDate.setHours(0, 0, 0, 0); // Normalize to the beginning of the day

//   fetchMovideskTickets({
//     token: apiBaseToken,
//     baseendpoint: apiBaseUrl,
//     sheetInfo: {
//       ssUrl: "https://docs.google.com/spreadsheets/d/10NRz6zsDpKq-X7fLIO31X0iL7iaHuoCkX78m7_h81Ro/",
//       sName: "22025 - GARANTIA",
//       headerRowEnd: 1,
//     },
//     colMapping: [
//       { col: 0, fv: 116372, isDate: true }, 
//       { col: 1, fv: "id" }, 
//       { col: 2, fv: 113146 },
//       { col: 3, fv: 152179 }, 
//       { col: 4, fv: 106871 }, 
//       { col: 5, fv: 105791 }, 
//       { col: 6, fv: 105789 }
//     ],
//     // WORKAROUND STEP 2: Simplify the API filter.
//     // We remove the date comparison from here and let the API handle the rest.
//     filterGroups: [
//       {
//         type: 'status',
//         rules: [{ field: 'status', operator: 'startsWith', value: 'S5' }],
//         extJoiner: 'and'
//       },
//       {
//         type: 'cfv',
//         rules: [
//           // The API CAN handle these simple string comparisons, so we keep them for efficiency.
//           { customFieldId: 187792, valueOperator: 'eq', valueContent: 'Braspress', joiner: 'and' },
//           { customFieldId: 106871, valueOperator: 'ne', valueContent: 'null' }
//         ]
//       }
//     ],
//     // WORKAROUND STEP 3: Provide a function to filter the date in the script.
//     postProcessFilter: (ticket) => {
//       // Find the custom field for the date (ID: 116372)
//       const dateField = (ticket.customFieldValues || []).find(cf => cf.customFieldId === 116372);

//       // If the field doesn't exist or has no value, discard the ticket.
//       if (!dateField || !dateField.value) {
//         return false;
//       }
      
//       // Convert the field's string value to a real JavaScript Date object.
//       const ticketDate = new Date(dateField.value);

//       // Return true ONLY if the ticket's date is greater than or equal to our cutoff date.
//       return ticketDate >= cutoffDate;
//     }
//   });
// }

// function runMyReport() {
//   getS1TicketsByDaysPast(7); // Busca tickets da última semana
// }


// /**
//  * Fetches tickets across all "S" statuses that belong to the "SOL MAIS" distributor.
//  */
// function getSolMais() {
//   fetchMovideskTickets({
//     token: apiBaseToken,
//     baseendpoint: apiBaseUrl,
//     top: 1000,
//     sheetInfo: {
//       ssUrl: "https://docs.google.com/spreadsheets/d/1mAD-xuqdKmRBhxPxwy6vKUgey33VUvmfoeweDze0imI/",
//       sName: "Main",
//       headerRowEnd: 1,
//     },
//     colMapping: [
//       { col: 0, fv: "id" }, 
//       { col: 1, fv: 107733, isDate: true },
//       { col: 2, fv: 113146 }, 
//       { col: 3, fv: 152179 }, 
//       { col: 5, fv: "status" },
//     ],
//     filterGroups: [
//       {
//         type: 'status',
//         rules: [{ field: 'status', operator: 'startsWith', value: 'S' }],
//         extJoiner: 'and'
//       },
//       {
//         type: 'cfv',
//         rules: [{
//           customFieldId: 92834,
//           itemValue: 'SOL MAIS'
//         }]
//       }
//     ]
//   });
// }




// /**
//  * Fetches tickets with a status starting with S1-S4 and THEN filters them
//  * in the script because the API does not support date comparison on custom fields.
//  * @param {number} daysPast The number of days in the past to filter by.
//  */
// function getS1TicketsByDaysPast_Workaround(daysPast) {
//   // 1. Defina a data de corte AQUI no script
//   const cutoffDate = new Date();
//   cutoffDate.setDate(cutoffDate.getDate() - daysPast);
//   cutoffDate.setHours(0, 0, 0, 0); // Zera o horário para comparar apenas o dia

//   fetchMovideskTickets({
//     token: apiBaseToken,
//     baseendpoint: apiBaseUrl,
//     top: 1000,
//     sheetInfo: {
//       ssUrl: "https://docs.google.com/spreadsheets/d/1SkGLGBJlUyStxUimCEcczrKN8jx0bnj_SNJ1Lbg_nc8/",
//       sName: "Under Warranty",
//       headerRowEnd: 1,
//     },
//     colMapping: [
//       { col: 0, fv: "id" }, 
//       { col: 1, fv: 152179 }, 
//       { col: 2, fv: 92408 },
//       { col: 3, fv: 102776 }, 
//       { col: 3, fv: 102777 }, 
//       { col: 4, fv: 92834 },
//       { col: 5, fv: 102763 }, 
//       { col: 6, fv: 102753 }, 
//       { col: 7, fv: 102774 },
//       { col: 8, fv: 144016 }, 
//       { col: 9, fv: 107733, isDate: true },
//       { col: 10, fv: 117072, isDate: true }, 
//       { col: 11, fv: 152180, isDate: true },
//       { col: 12, fv: 95991 }, 
//       { col: 17, fv: "todayDate" }, 
//       { col: 18, fv: "status" }
//     ],
//     // 2. Simplifique o filtro da API: PEÇA APENAS PELO STATUS
//     filterGroups: [
//       {
//         type: 'status',
//         rules: [
//           { field: 'status', operator: 'startsWith', value: 'S1', joiner: 'or' },
//           { field: 'status', operator: 'startsWith', value: 'S2', joiner: 'or' },
//           { field: 'status', operator: 'startsWith', value: 'S3', joiner: 'or' },
//           { field: 'status', operator: 'startsWith', value: 'S4' }
//         ]
//       }
//     ],
//     // 3. Adicione uma nova função de pós-processamento para filtrar os resultados
//     postProcessFilter: (ticket) => {
//       const customFieldMap = new Map();
//       (ticket.customFieldValues || []).forEach(cf => {
//         customFieldMap.set(cf.customFieldId, cf.value);
//       });

//       const dateFieldValue = customFieldMap.get(152180); // Pega o valor do campo de data
//       if (!dateFieldValue) {
//         return false; // Se não tiver o campo, descarte o ticket
//       }

//       const ticketDate = new Date(dateFieldValue);
//       // Retorna true se a data do ticket for maior ou igual à data de corte
//       return ticketDate >= cutoffDate;
//     }
//   });
// }
