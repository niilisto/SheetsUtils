# Sheets Utils

Utility collection for Google Sheets.
The project uses the V8 runtime, the `America/Sao_Paulo` timezone, and standalone `.js` files that can be synced with `clasp`.

## What is included

- Move, copy, and clean rows across sheets using single keys, multi-keys, filters, or status rules.
- Format columns, text, dates, currency, and numeric values.
- Generate automatic timestamps from `onEdit` triggers.
- Create Movidesk ticket links directly in the spreadsheet.
- Classify inverters by grid, phase, workflow, and family.
- Split, expand, and consolidate data for export or synchronization.

## Install in Apps Script

### Option 1: Google Apps Script editor

1. Open the Google Sheet or create a blank Apps Script project.
2. Copy the `.js`, `.clasp.json`, and `appsscript.json` files into the project.
3. Make sure the runtime is set to V8 and the timezone remains `America/Sao_Paulo`.
4. Save the project and run a simple function once to grant spreadsheet access.
5. If you will use edit automation, create simple or installable triggers for the `onEdit` functions.

### Option 2: clasp

1. Install the CLI: `npm install -g @google/clasp`.
2. Log in to your Google account: `clasp login`.
3. Open this repository folder and push the code with `clasp push`.
4. If you want a new project, run `clasp create --type sheets --title "sheetsUtils"` and update the `scriptId` in `.clasp.json`.
5. Use `clasp pull` and `clasp push` to keep the local copy and the Apps Script copy in sync.

## How to run

- Functions such as `getTimestamp`, `eraseTimestamp`, `moveRowsByStatus`, and `runRemoveDuplicates` work best inside wrappers or triggers.
- Functions that use `SpreadsheetApp.getActiveSpreadsheet()` should run in a spreadsheet-bound project or with the target sheet open in the correct context.
- Functions that receive a `spreadsheetId` can work on external files without depending on the active spreadsheet.

### `onEdit` trigger example

```javascript
function onEdit(e) {
  getTimestamp(e, 2, 'Orders', [
    { trigger: 'C', target: 'G' }
  ]);

  eraseTimestamp(e, 2, 'Orders', [
    { trigger: 'C', target: 'G' }
  ]);
}
```

## Spreadsheet tips

- Keep the header on row 1 and the data starting on row 2 whenever possible.
- Use column letters consistently, such as `A`, `B`, `AA`, and `AB`.
- For key-based functions, make sure the key column is actually unique in the source and destination.
- For dates, prefer values that Google Sheets can recognize or standardized text formats.
- For large datasets, prefer filters and batch operations instead of reading and writing cell by cell.
- For `onEdit` functions, validate that the trigger is attached only to the expected columns to avoid unwanted overwrites.

## Ready-to-use examples

### 1. Generate an automatic timestamp

```javascript
function onEdit(e) {
  getTimestamp(e, 2, 'Orders', [
    { trigger: 'C', target: 'G' }
  ]);
}
```

### 2. Copy or update data by key

```javascript
function syncTickets() {
  copyByKey({
    sourceSpreadsheetId: 'SOURCE_SPREADSHEET_ID',
    sourceSheetName: 'Raw',
    sourceStartRow: 2,
    sourceKeyCol: 'A',
    targetSpreadsheetId: 'TARGET_SPREADSHEET_ID',
    targetSheetName: 'Tickets',
    targetStartRow: 2,
    targetKeyCol: 'A',
    colsMap: {
      A: 'A',
      B: 'B',
      C: 'C',
      D: 'D'
    },
    overwrite: true,
    timestampCols: ['H'],
    timestampLocale: 'America/Sao_Paulo',
    timestampFormat: 'dd/MM/yyyy HH:mm:ss'
  });
}
```

### 3. Generate Movidesk ticket links in the ticket column

```javascript
function linkTickets() {
  generateTicketLinks(2, 'Tickets', 'A');
}
```

### 4. Remove duplicate tickets

```javascript
function cleanDuplicates() {
  removeDuplicateTickets('SPREADSHEET_ID', ['Sheet1', 'Sheet2'], 'A', 2);
}
```

## Quick function reference

### Cleanup, copy, and transfer

| File | Status | Summary |
| --- | --- | --- |
| `clearTab.js` | active | Clears cell contents starting from a row without deleting rows. |
| `deleteTab.js` | active | Deletes all rows starting at `startRow`. |
| `deleteRowsByFilter.js` | active | Deletes rows in batches when filters match. |
| `cutByKey.js` | active | Moves rows between spreadsheets by key and removes them from the source. |
| `copyByKey.js` | active | Copies or updates rows by key with filters and optional timestamps. |
| `updateByKeyMulti.js` | active | Updates values using single or multiple key groups. |
| `transferByKey.js` | active/variant | Variant for transferring and updating by multi-key rules. |
| `safeDelete.js` | active | Deletes rows only when the key exists in the comparison sheet. |
| `removeMatchingTickets.js` | active | Removes output rows whose ticket already exists in the input. |
| `removeRepeated.js` | active | Removes duplicate tickets using a reference column. |
| `moveRowsByStatus.js` | active | Moves rows between sheets based on status rules. |
| `maintainBlankRows.js` | active | Keeps a fixed number of blank rows at the bottom of a sheet. |
| `filterAndTransferTickets.js` | active | Filters tickets across multiple sheets and transfers only unique rows. |
| `splitAndExport.js` | active | Splits comma-separated values into individual rows during export. |
| `splitAndRewriteElements.fnc.js` | active | Expands comma-separated elements and rewrites the output. |
| `obtainData.js` | active | Collects data from multiple sheets using column mapping. |

### Columns, formatting, and dates

| File | Status | Summary |
| --- | --- | --- |
| `convertColumnDateFormat.fnc.js` | active | Converts dates between formats with timezone support. |
| `clearCellIfComparisonDateIsNewer.js` | active | Clears a cell when the comparison date is newer. |
| `formatCol.js` | active | Applies TEXT, NUMBER, DECIMAL, DATE, CURRENCY, and JOIN formats. |
| `formatColumns.fnc.js` | active | Formats columns in an external spreadsheet by ID. |
| `formatText.js` | active | Adjusts font, size, color, and alignment. |
| `convertColToStar.js` | active | Converts numeric scores into star representations. |
| `columnConverter.fnc.js` | active | Converts column letters to 1-based numeric indexes. |
| `letterCol.js` | active | Converts column letters and helps identify the edited column. |
| `removeQuantifiers.fnc.js` | active | Removes numeric prefixes and quantity suffixes from a value. |
| `writeToken.js` | active | Generates unique tokens with prefix, timestamp, and random suffix. |
| `writeToken copy.js` | backup | Backup copy of the token generator. |

### Events and timestamps

| File | Status | Summary |
| --- | --- | --- |
| `getTimestamp.js` | active | Writes automatic timestamps from `onEdit`. |
| `eraseTimestamp.js` | active | Clears timestamps when the trigger cell becomes empty. |
| `onlyRunOnColumns.js` | active | Runs a callback only for allowed columns. |

### Tickets, links, and Movidesk

| File | Status | Summary |
| --- | --- | --- |
| `generateTicketLinks.js` | active | Creates RichText links that point to Movidesk tickets. |
| `fetchMovideskTicketsv1.js` | commented | Draft for importing tickets from the Movidesk API. |
| `fetchMovideskTicketsv2.js` | commented | Version with paging and dynamic `$select` and `$expand`. |
| `fetchMovideskTicketsv3.js` | commented | Version with an additional client-side filter. |

### Inverters and classification

| File | Status | Summary |
| --- | --- | --- |
| `findEqGrid.js` | active | Classifies the equipment as on-grid or off-grid. |
| `findEqPhase.js` | active | Classifies the equipment as single-phase or three-phase. |
| `findWorkflow.js` | active | Determines the workflow by combining distributor and phase. |
| `getInverterFamily.js` | active | Fills the inverter family based on a model-to-family mapping table. |

### Legacy and references

| File | Status | Summary |
| --- | --- | --- |
| `updateByKey.js` | legacy/commented | Older version of the key-based update helper. |

## Notes

- The remaining `.fnc.js` suffix is just a local convention used to separate helper variants.
- Some functions still exist as older versions, comments, or backups; the table above reflects their status.
- When adapting the project for another spreadsheet, the main changes are usually the `spreadsheetId`, sheet names, and column letters.
