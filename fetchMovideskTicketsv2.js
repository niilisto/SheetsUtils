// /**
//  * Fetches Movidesk tickets (with optional paging), dynamically builds
//  * $select, $filter, $expand, $orderby, $top, and appends new rows to a Google Sheet.
//  *
//  * @param {Object} options
//  * @param {string} options.token                                  Movidesk API token
//  * @param {string} options.baseendpoint                                Base endpoint (e.g. https://api.movidesk.com/public/v1/tickets)
//  * @param {Array<{col:number,fv:string|number,isDate?:boolean}>} options.colMapping
//  *                                                                  Column mapping for sheet output
//  * @param {Array<{
//  *   type: 'status'|'cfv'|'date',
//  *   rules: Array<{
//  *     field?: string, operator?: string, value?: string,
//  *     customFieldId?: number, itemValue?: string,
//  *     joiner?: 'and'|'or'
//  *   }>,
//  *   extJoiner?: 'and'|'or'
//  * }>} [options.filterGroups]                                      Filter groups with internal (`rules[].joiner`) and external (`extJoiner`) joiners
//  * @param {string[]} [options.selectFields]                          Fields for $select
//  * @param {Object} [options.expandCustomFields]
//  * @param {number[]} [options.expandCustomFields.filterIds]          customFieldIds to expand
//  * @param {string[]} [options.expandCustomFields.select]             Fields inside customFieldValues to select
//  * @param {boolean}  [options.expandCustomFields.expandItems]        Whether to expand items
//  * @param {Object} [options.expandClients]
//  * @param {string[]} [options.expandClients.select]                  Fields inside clients to select
//  * @param {string[]} [options.expandClients.expandOrganization]      Fields inside clients.organization to select
//  * @param {string} [options.orderBy]                                 $orderby clause
//  * @param {number} [options.top]                                     $top (page size)
//  * @param {{ ssUrl:string, sName:string, headerRowEnd:number }} options.sheetInfo
//  *
//  * @example
//  * // 1) Status filter + single customField expand + sort + select
//  * fetchMovideskTickets({
//  *   token: '…',
//  *   baseendpoint: 'https://api.movidesk.com/public/v1/tickets',
//  *   selectFields: ['id','category','createdDate'],
//  *   filterGroups: [{
//  *     type: 'status',
//  *     rules: [{ field:'status', operator:'eq', value:'S4 - COLETA REVERSA' }]
//  *   }],
//  *   expandCustomFields: {
//  *     filterIds: [92408],
//  *     select: ['value'],
//  *     expandItems: false
//  *   },
//  *   orderBy: 'id desc',
//  *   colMapping: [ { col:0, fv:'id' } ],
//  *   sheetInfo: { ssUrl:'…', sName:'Main', headerRowEnd:1 }
//  * });
//  *
//  * @example
//  * // 2) Category & status & CFV.any filter + select id only
//  * fetchMovideskTickets({
//  *   token: '…',
//  *   baseendpoint: 'https://api.movidesk.com/public/v1/tickets',
//  *   selectFields: ['id'],
//  *   filterGroups: [
//  *     {
//  *       type: 'status',
//  *       rules: [
//  *         { field:'category', operator:'eq', value:'Garantia', joiner:'and' },
//  *         { field:'status',   operator:'eq', value:'S6 - REPARO LABORATORIO' }
//  *       ]
//  *     },
//  *     {
//  *       type: 'cfv',
//  *       rules: [
//  *         { customFieldId: 0, itemValue:'DXH7CA318F' }
//  *       ]
//  *     }
//  *   ],
//  *   colMapping: [ { col:0, fv:'id' } ],
//  *   sheetInfo: { ssUrl:'…', sName:'Main', headerRowEnd:1 }
//  * });
//  *
//  * @example
//  * // 3) Heavy CFV expansion for a single ticket
//  * fetchMovideskTickets({
//  *   token: '…',
//  *   baseendpoint: 'https://api.movidesk.com/public/v1/tickets',
//  *   selectFields: ['id'],
//  *   filterGroups: [{ type:'status', rules:[{ field:'id',operator:'eq',value:'143043' }] }],
//  *   expandCustomFields: {
//  *     filterIds: [152179,92408,95991,92834,92993,144662,103454,187792,105792,106440],
//  *     select: ['value','customFieldId','customFieldRuleId','items'],
//  *     expandItems: true
//  *   },
//  *   colMapping: [ { col:0, fv:'id' } ],
//  *   sheetInfo: { ssUrl:'…', sName:'Main', headerRowEnd:1 }
//  * });
//  *
//  * @example
//  * // 4) Date range OR status prefix combos + paging + orderBy
//  * fetchMovideskTickets({
//  *   token: '…',
//  *   baseendpoint: 'https://api.movidesk.com/public/v1/tickets',
//  *   selectFields: ['id','status','createdDate'],
//  *   filterGroups: [
//  *     {
//  *       type:'date',
//  *       rules:[
//  *         { field:'createdDate',operator:'ge',value:'2022-10-24',joiner:'and' },
//  *         { field:'createdDate',operator:'le',value:'2022-10-25',joiner:'and' },
//  *         { field:'status',operator:'startsWith',value:'S' }
//  *       ],
//  *       extJoiner:'or'
//  *     },
//  *     {
//  *       type:'date',
//  *       rules:[
//  *         { field:'createdDate',operator:'ge',value:'2022-10-24',joiner:'and' },
//  *         { field:'createdDate',operator:'le',value:'2022-10-25',joiner:'and' },
//  *         { field:'status',operator:'startsWith',value:'F' }
//  *       ]
//  *     }
//  *   ],
//  *   orderBy: 'id desc',
//  *   top: 100,
//  *   colMapping: [
//  *     { col:0, fv:'id' },
//  *     { col:1, fv:'status' },
//  *     { col:2, fv:'createdDate', isDate:true }
//  *   ],
//  *   sheetInfo: { ssUrl:'…', sName:'Main', headerRowEnd:1 }
//  * });
//  */

// function fetchMovideskTickets(options) {
//   const { token, baseendpoint, colMapping, filterGroups, sheetInfo } = options;

//   // 1) Build $filter
//   const filterString = buildODataFilterString(filterGroups);
//   const filterParam = filterString ? `$filter=${encodeURIComponent(filterString)}` : '';

//   // 2) Build $select from colMapping strings
//   const selectFields = colMapping
//     .filter(m => typeof m.fv === 'string')
//     .map(m => m.fv);
//   const selectParam = selectFields.length
//     ? `$select=${encodeURIComponent(selectFields.join(','))}`
//     : '';

//   // 3) Build $expand for numeric CF IDs
//   const cfIds = [...new Set(
//     colMapping
//       .filter(m => typeof m.fv === 'number')
//       .map(m => m.fv)
//   )];
//   const expandParam = cfIds.length
//     ? `$expand=${encodeURIComponent(
//       `customFieldValues(` +
//       `$filter=${cfIds.map(id => `customFieldId eq ${id}`).join(' or ')};` +
//       `$select=value,customFieldId,items;` +
//       `$expand=items($select=customFieldItem,storageFileGuid)` +
//       `)`
//     )}`
//     : '';

//   // 4) Always include token first, then any non-empty clauses
//   const parts = [
//     `token=${token}`,
//     filterParam,
//     selectParam,
//     expandParam
//   ].filter(Boolean);

//   // 5) Paging loop
//   let page = 0;
//   const allRows = [];

//   while (true) {
//     const skipParam = parts.includes('$top=') ? `&$skip=${page * top}` : '';
//     const url = `${baseendpoint}?${parts.join('&')}${skipParam}`;
//     Logger.log('URL: ' + url);

//     const resp = UrlFetchApp.fetch(url, { headers: { Accept: 'application/json' } });
//     if (resp.getResponseCode() !== 200) {
//       throw new Error(`API returned ${resp.getResponseCode()}`);
//     }

//     const items = JSON.parse(resp.getContentText());
//     if (!Array.isArray(items) || items.length === 0) {
//       break;
//     }

//     // Map tickets → rows
//     items.forEach(ticket => {
//       const cfMap = new Map();
//       (ticket.customFieldValues || []).forEach(cf => {
//         const v = cf.items?.length
//           ? cf.items.map(i => i.customFieldItem).join(', ')
//           : cf.value;
//         cfMap.set(cf.customFieldId, v);
//       });

//       const row = Array(Math.max(...colMapping.map(m => m.col)) + 1).fill('');
//       colMapping.forEach(m => {
//         let v = typeof m.fv === 'number'
//           ? (cfMap.get(m.fv) || '')
//           : (ticket[m.fv] || '');
//         if (m.isDate && v) {
//           const d = new Date(v);
//           if (!isNaN(d)) v = d;
//         }
//         row[m.col] = v;
//       });

//       allRows.push(row);
//     });

//     // If no paging desired, break now
//     if (!parts.some(p => p.startsWith('$top='))) break;
//     page++;
//   }

//   // 6) Append only new rows, even if sheet is empty
//   const ss = SpreadsheetApp.openByUrl(sheetInfo.ssUrl);
//   const sh = ss.getSheetByName(sheetInfo.sName);
//   const startRow = sheetInfo.headerRowEnd + 1;
//   const lastRow = sh.getLastRow();

//   // If there are no rows yet (lastRow < startRow), treat existingIds as empty set
//   let existingIds = new Set();
//   if (lastRow >= startRow) {
//     // safely read only if there's at least one data row
//     const numExisting = lastRow - startRow + 1;
//     const vals = sh
//       .getRange(startRow, 1, numExisting, 1)
//       .getValues()
//       .flat();
//     existingIds = new Set(vals.map(String));
//   }

//   // Filter out rows whose ID already exists
//   const newRows = allRows.filter(r => !existingIds.has(String(r[0])));

//   if (newRows.length) {
//     sh
//       .getRange(lastRow + 1, 1, newRows.length, newRows[0].length)
//       .setValues(newRows);
//     Logger.log(`Appended ${newRows.length} new rows.`);
//   } else {
//     Logger.log('No new tickets to add.');
//   }
// }


// /**
//  * Builds an OData filter string with internal rule joiners and group-level external joiners.
//  *
//  * @param {Array<{
//  *   type: 'status'|'cfv'|'date',
//  *   rules: Array<{
//  *     field?: string, operator?: string, value?: string,
//  *     customFieldId?: number, itemValue?: string,
//  *     joiner?: 'and'|'or'
//  *   }>,
//  *   extJoiner?: 'and'|'or'
//  * }>} filterGroups
//  * @returns {string}
//  */
// function buildODataFilterString(filterGroups) {
//   const sanitize = v =>
//     typeof v === 'string' ? v.replace(/'/g, "''") : String(v);

//   function buildGroup(group) {
//     const clauses = group.rules.map((rule, idx) => {
//       let expr = '';

//       if (group.type === 'status' || group.type === 'date') {
//         if (!rule.field || !rule.operator) {
//           throw new Error(`Missing 'field' or 'operator' in rule: ${JSON.stringify(rule)}`);
//         }
//         if (rule.operator === 'startsWith' || rule.operator === 'endsWith' || rule.operator === 'substringof') {
//           expr = `${rule.operator}(${rule.field},'${sanitize(rule.value)}')`;
//         } else {
//           expr = `${rule.field} ${rule.operator} '${sanitize(rule.value)}'`;
//         }
//       } else if (group.type === 'cfv') {
//         if (!rule.customFieldId) {
//           throw new Error(`Missing 'customFieldId' in cfv rule: ${JSON.stringify(rule)}`);
//         }
//         const itemClause = rule.itemValue
//           ? ` and c/items/any(i: i/customFieldItem eq '${sanitize(rule.itemValue)}')`
//           : '';
//         expr = `customFieldValues/any(c: c/customFieldId eq ${rule.customFieldId}${itemClause})`;
//       } else {
//         throw new Error(`Unknown group type: ${group.type}`);
//       }

//       const joiner = (rule.joiner || 'and').toLowerCase();
//       const isLast = idx === group.rules.length - 1;
//       return isLast ? expr : `${expr} ${joiner} `;
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