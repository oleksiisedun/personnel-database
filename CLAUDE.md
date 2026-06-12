# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Deployment

```bash
# Push to the primary bound script project
clasp push

# Push to all target spreadsheets listed in clasp-targets.json
./clasp-push.sh
```

There is no build step, linter, or test suite.

## Code conventions

JS style and JSDoc rules are in the global `~/.claude/CLAUDE.md`. Project-specific rule: all tuneable constants belong in `Config.js` (see below).

## Architecture

This is a **Google Apps Script** project (V8 runtime) bound to a Google Spreadsheet. The web editor runs inside an `HtmlService` modal dialog opened from the Sheets menu.

### Sheet layout

The `Database` sheet has this fixed structure:
- Row 1 — column names (header)
- Row 2 — column types (`text`, `image`, `*-table`, `unit`, `origin`, `marital-status`, `sex`, `tin`)
- Row 3+ — one person record per row

The `Handbook` sheet holds schema metadata, dropdown option lists, the Master Mode toggle, the data folder cell, and export correspondence tables. Constants in `Config.js` define each range address.

Every Handbook range read inside `getSchemaAndData()` (`getMasterMode()`, `getMasterSources()`, `HANDBOOK_TYPES_RANGE`, and the dropdown-option ranges) is wrapped in try/catch with a safe fallback (`false`, `[]`, `{}`, or `[]`). This keeps the editor usable even if a particular Handbook range is unreadable for some user. New Handbook reads added to `getSchemaAndData()` should follow the same pattern.

Key Handbook cells:
- `M2` (`MASTER_MODE_CELL`) — Master Mode checkbox
- `M4` (`DATA_FOLDER`) — Google Drive folder ID containing person images and PDFs
- `M6` (`ACTUAL_PERSONNEL_SPREADSHEET_CELL`) — link/ID of the spreadsheet containing the actual personnel list
- `M7` (`ACTUAL_PERSONNEL_RANGE_CELL`) — range address within that spreadsheet (e.g. `Sheet1!A:A`) holding full names
- `M9` (`EXPORT_F1_TEMPLATE_CELL`) — Google Drive file ID of the F-1 Docs template
- `M11` (`EXPORT_WC_TEMPLATE_CELL`) — Google Drive file ID of the Wanted Card Docs template
- `M13` (`EXPORT_FOLDER_CELL`) — Google Drive folder ID where exported documents are saved
- `N2:N` (`MASTER_MODE_SOURCES_RANGE`) — source spreadsheet IDs for Master Mode

### How the HTML template works

`WebEditor.html` is served via `HtmlService.createTemplateFromFile()`. It includes CSS and JS using scriptlet tags:

```html
<?!= HtmlService.createHtmlOutputFromFile('WebEditor.css').getContent(); ?>
<?!= HtmlService.createHtmlOutputFromFile('WebEditor.js').getContent(); ?>
```

The `!` in `<?!=` means the content is included **without** HTML sanitization — the raw `<style>` and `<script>` tags are injected as-is. Column width CSS variables are also set via scriptlets in `WebEditor.html` using `getColumnMinWidths()` / `getColumnMaxWidths()` from `Code.js`.

### Client ↔ server communication

`google.script.run` is the only way the client talks to the server. Every call is asynchronous and must chain `.withSuccessHandler()` and `.withFailureHandler()` before the function name:

```js
google.script.run
  .withSuccessHandler(result => { ... })
  .withFailureHandler(err => { ... })
  .serverFunctionName(arg1, arg2);
```

The client-side code in `WebEditor.js.html` is not a module — all functions are globals within the `HtmlService` sandbox.

### Opening remote spreadsheets safely (`openSpreadsheetSafely`)

If a `SpreadsheetApp.openById(id)` call throws because the current user can't access `id`, that permission error — even when caught with try/catch — still causes the *overall* enclosing execution to fail (e.g. the whole `getSchemaAndData()` RPC fails with "У вас немає дозволу на доступ до потрібного документа" / "You do not have permission to access the required document"). This is an Apps Script runtime quirk where a caught `SpreadsheetApp.openById()` permission error poisons the execution's final status regardless of being handled in JS.

`openSpreadsheetSafely(id)` (in `Code.js`) avoids this: it first probes accessibility with `DriveApp.getFileById(id)` (caught, returns `null` on failure — the same safe pattern `getImagesDataUrls()` uses for "no-access" images) and only calls `SpreadsheetApp.openById(id)` if that probe succeeds, returning `null` on any failure. **Every** call site that opens a spreadsheet by an ID that isn't guaranteed accessible to the current user — `Handbook!N2:N` master sources, `Handbook!M6` actual-personnel link, and client-supplied `spreadsheetId`s in `movePersonnel`/`updateRow`/`deleteRow`/`Export.js` — uses this helper instead of calling `SpreadsheetApp.openById()` directly. New code that opens a spreadsheet by such an ID should do the same.

### Master Mode

When `Handbook!M2` is `true`, `getSchemaAndData()` reads spreadsheet IDs from `Handbook!N2:N` and appends rows from each remote spreadsheet's `Database` sheet to the local rows via `openSpreadsheetSafely()`, skipping any source that's inaccessible. Remote rows carry a `spreadsheetId` property; local rows do not.

`updateRow(rowIndex, values, spreadsheetId)` and `deleteRow(rowIndex, spreadsheetId)` route to the correct spreadsheet based on `spreadsheetId` — `openSpreadsheetSafely()` for remote (throwing a clean `Error` if inaccessible), `getActiveSpreadsheet()` for local. New rows (`addRowWithData(values)`) always go to the local `Database` sheet. Editing is never restricted by Master Mode state.

When masterMode is true, `getSchemaAndData()` also calls `getSourceSpreadsheetInfos()` and includes the result as `masterSources: Array<{id, name}>` in the return value (current spreadsheet has `id: null`; inaccessible sources fall back to showing the raw ID as the name). The client uses this to populate the Move destination dropdown.

### Actual personnel filter

`getActualPersonnelNames()` reads the spreadsheet link from `Handbook!M6` and the range address from `Handbook!M7`, opens the spreadsheet via `openSpreadsheetSafely()`, and returns the flat list of non-empty name strings. Returns `null` if either cell is empty or the spreadsheet is inaccessible.

`getSchemaAndData()` includes the result as `actualPersonnelNames: string[]|null` in its return value. The client enables the **"Actual personnel"** toolbar checkbox only when the array is non-null and non-empty; otherwise the checkbox stays disabled. When the checkbox is checked, `applyFilters()` additionally requires that `row.values[0]` (first column = full name) is present in `actualPersonnelNames`. The filter composes with all existing column text filters and the regex toggle.

### Move personnel (`movePersonnel`)

`movePersonnel(rowEntries, destinationSpreadsheetId)` is available only in Master Mode. It moves rows between spreadsheets:

1. Groups `rowEntries` by source `spreadsheetId` and sorts each group in descending `rowIndex` order (so deleting lower rows doesn't shift higher ones).
2. For each row: reads the row data, appends it to the destination `Database` sheet, then **hard-deletes** it from the source (`sheet.deleteRow()` — not a soft delete like `deleteRow()`). Returns the new `rowIndex` and `values` so the client can update `schema.rows` in place without reloading.
3. Rows where source === destination spreadsheet are skipped (logged but not moved or deleted).
4. After each row move, tries to move the person's Drive folder (named after the first column value) from the source `DATA_FOLDER` to the destination `DATA_FOLDER` using `DriveApp.getFolderById()` / `getFoldersByName()` / `moveTo()`. If either `DATA_FOLDER` is not configured, or the named folder is not found, the row move still completes and a note is logged. `parseDriveId(value)` strips any Drive URL wrapper from the cell value to get a bare folder ID.

### Soft delete (Trash sheet)

`deleteRow()` never permanently removes data. It copies the row to the `Trash` sheet and then deletes it from `Database`. If `Trash` doesn't exist yet, it is created with the same two header rows as `Database`. There is no restore UI — recovery requires manually moving rows back in Sheets.

### Edit view layout

`#edit-form` is a two-column CSS grid (`grid-template-columns: 1fr 1fr`). Most field blocks occupy one column. Table-type fields get the additional class `field-block--full` (→ `grid-column: 1 / -1`) so they span the full width, since their multi-row editors are too wide for a half-width cell.

### `*-table` column widths

Both `buildTableEditor()` (edit view) and `buildMiniTable()` (list view) prepend a `<colgroup>` to their `<table>` with each `<col>` width set as a percentage proportional to the **square root** of the maximum string length found in that column (header text or any data value, whichever is longer, minimum 8 chars). The sqrt scaling prevents very long columns (e.g. address) from dominating and squeezing short but important ones (e.g. phone). `mini-table` uses `table-layout: fixed` so the browser honours the colgroup widths instead of overriding them with content-based sizing.

### Dropdown column types

`unit`, `origin`, `marital-status`, and `sex` columns each read their allowed values from a dedicated Handbook range at load time. The pattern for adding a new dropdown type is:
1. Add a `HANDBOOK_XYZ_RANGE` constant in `Config.js`
2. Add an entry to the `DROPDOWN_TYPES` table in `getSchemaAndData()` — `{ type, range, key }` — one line
3. No change needed in `openEditView()` — `buildDropdownField()` handles all dropdown types automatically via the `col.xyzOptions` key
4. No change needed in `buildListView()` — filter inputs are rendered unconditionally for all column types

### Image loading

Images are fetched server-side via `DriveApp` (using the script owner's OAuth token) so users without personal Drive access can still view images. The client caches results in two layers:

1. **IndexedDB** (`pdb_images` / `images` store) — persistent across dialog sessions. `openImageDB()` is called eagerly in `init()` so the DB is ready before `onDataLoaded` fires. `loadCacheFromIndexedDB()` (async) pre-populates `imageCache` before `renderList()` runs; `saveToIndexedDB()` writes each fetched result back. Entries expire after `IMAGE_CACHE_TTL_DAYS` days (default 7). Use IndexedDB (not localStorage) because the full base64 dataset easily exceeds the ~5 MB localStorage quota.
2. **In-memory `imageCache`** — session-only map from fileId to result, used for O(1) lookups during rendering and lightbox opens.

Loading uses a **concurrency pool** controlled by `IMAGE_FETCH_CONCURRENCY` and `IMAGE_FETCH_BATCH_SIZE` in `Config.js` — raising concurrency too high exceeds Apps Script's ~30 concurrent-execution limit and causes dropped images. See the comments in `Config.js` for tuning guidance.

`DriveApp.getFileById()` throws for inaccessible files; the server catches it and returns `{ type: 'no-access' }`. The client renders a gray "No access" badge.

### Document export (`Export.js`)

Export is triggered from the toolbar. The Export F-1 and Export WC buttons are disabled until at least one row is checked via the selection checkbox column. `runExport()` collects `{ rowIndex, spreadsheetId }` entries from `selectedKeys`-checked rows and passes them to the server (Master Mode rows carry a non-null `spreadsheetId`).

`exportF1(rowEntries)` and `exportWC(rowEntries)` copy Google Docs templates into the export folder and fill placeholders with row data. Both delegate to `_exportDoc()`, which caches sheet data per `spreadsheetId` (so each remote spreadsheet is read at most once per export call) and runs four passes in strict order:

1. **Image columns** — `{Column Name}` placeholder replaced with the actual image blob.
2. **Service history table** — `{Проходження служби}` placeholder row expanded into one table row per service entry. Must run before pass 3 or the placeholder text would be consumed before the table handler can locate it.
3. **Direct text columns** — remaining `{Column Name}` placeholders replaced with cell text.
4. **Correspondence table** — Handbook-defined aliases and computed values (e.g. `totalServiceLength`, `motherFullName`) replace their own placeholders.

Placeholders with no match are left untouched. `_exportDoc()` checks elapsed time against `EXPORT_TIME_LIMIT_MS` before each row (5 min — 1 min safety margin before the GAS 6-min hard kill) and returns `{ results, remaining }` so the client can surface any unprocessed rows.

## Constants — always in `Config.js`

All tuneable values belong in `Config.js`. Constants needed by the client must also be threaded through the `getSchemaAndData()` return value in `Code.js` (see `filterDebounceMs`, `imageFetchBatchSize`, `imageFetchConcurrency`, `masterMode` as examples). Never hardcode magic numbers in `WebEditor.js.html`.
