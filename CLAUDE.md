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

### CSS design system (`WebEditor.css.html`)

Concrete tokens and shared classes already defined here (see global CSS design-system conventions for the reuse-first principle):

- **Variables**: `--color-primary`, `--color-primary-hover`, `--color-focus`, `--color-focus-shadow`, `--color-danger`, `--color-danger-hover`, `--color-border`, `--radius`.
- **Shared classes**: `.btn-primary` (filled accent — Save/Move/Close), `.btn-secondary` (outline — Back/Cancel/toolbar), `.overlay`/`.overlay-dialog` (modal scaffolding — export progress, move dialog).
- **Global states**: `button:disabled` (`opacity: 0.5; cursor: default;`), text input/select focus (`border-color: var(--color-focus); box-shadow: 0 0 0 2px var(--color-focus-shadow);`).

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

All write functions that need the `Database` sheet (`addRowWithData`, `updateRow`, `deleteRow`, `deleteRows`, `movePersonnel` destination) call `getDatabaseSheet(spreadsheetId)` from `Utils.js`, which wraps `resolveSpreadsheet` + `getSheetByName(SHEET_DATABASE)` and throws on failure. It returns `{ ss, sheet }` — destructure only what the caller needs. New write functions should use this helper instead of repeating the four-line pattern inline.

If a user reports this exact `PERMISSION_DENIED` error despite having real access to everything involved, check for multiple signed-in Google accounts in their browser before chasing a code fix — the dialog's `google.script.run` calls can authenticate against the wrong signed-in account, producing the same error as a genuine sharing gap. Confirmed fix in one case: log out of all Google accounts except the one with access.

### Master Mode

When `Handbook!M2` is `true`, `getSchemaAndData()` reads source spreadsheet IDs from `Handbook!N2:N` via `getMasterSources()` and returns them as `masterSourceIds: string[]` — it does **not** open any remote spreadsheet itself. Local rows have no `spreadsheetId` property; `masterSources` is seeded with just the local entry (`{ id: null, name: <current spreadsheet name> }`).

After the initial render, the client (`onDataLoaded()` in `WebEditor.js.html`) calls `queueMasterSourceFetch()`, which fetches each remote source's rows one at a time through a worker pool (size `masterModeFetchConcurrency`/`MASTER_MODE_FETCH_CONCURRENCY`, mirroring the image-fetch concurrency pool below) via the server function `getMasterSourceRows(spreadsheetId)`. Each result's rows (carrying a `spreadsheetId` property) are pushed onto `schema.rows`, its `{id, name}` is pushed onto `schema.masterSources`, and the list is silently re-filtered/re-rendered (debounced via `schema.filterDebounceMs`) — so remote rows stream in after the local-only data is already visible instead of blocking the initial load. Inaccessible sources are skipped silently, same as before. New code that needs "all master-mode rows" should be aware that, on the client, they may still be arriving for a short time after `onDataLoaded()` fires.

`updateRow(rowIndex, values, spreadsheetId)` and `deleteRow(rowIndex, spreadsheetId)` route to the correct spreadsheet based on `spreadsheetId` — `openSpreadsheetSafely()` for remote (throwing a clean `Error` if inaccessible), `getActiveSpreadsheet()` for local. New rows go to the local `Database` sheet by default. In Master Mode, a spreadsheet selector appears left of the Save button in the Add person view, letting the user pick any loaded master source as the target; `addRowWithData(values, spreadsheetId)` routes accordingly via `openSpreadsheetSafely()`. Editing is never restricted by Master Mode state.

The client uses `masterSources: Array<{id, name}>` to populate the Move destination dropdown (current spreadsheet has `id: null`; inaccessible sources fall back to showing the raw ID as the name) — it grows as remote sources resolve.

### Actual personnel filter

`getActualPersonnelNames()` reads the spreadsheet link from `Handbook!M6` and the range address from `Handbook!M7`, opens the spreadsheet via `openSpreadsheetSafely()`, and returns the flat list of non-empty name strings. Returns `null` if either cell is empty or the spreadsheet is inaccessible.

`getSchemaAndData()` includes the result as `actualPersonnelNames: string[]|null` in its return value. The client enables the **"Actual personnel"** toolbar checkbox only when the array is non-null and non-empty; otherwise the checkbox stays disabled. When the checkbox is checked, `applyFilters()` additionally requires that `row.values[0]` (first column = full name) is present in `actualPersonnelNames`. The filter composes with all existing column text filters and the regex toggle.

### Move personnel (`movePersonnel`)

`movePersonnel(rowEntries, destinationSpreadsheetId)` is available only in Master Mode. It moves rows between spreadsheets:

1. Groups `rowEntries` by source `spreadsheetId` and sorts each group in descending `rowIndex` order (so deleting lower rows doesn't shift higher ones).
2. For each row: reads the row data, appends it to the destination `Database` sheet, then **hard-deletes** it from the source (`sheet.deleteRow()` — not a soft delete like `deleteRow()`). Returns the new `rowIndex` and `values` so the client can update `schema.rows` in place without reloading.
3. Rows where source === destination spreadsheet, or whose source spreadsheet/sheet is inaccessible, are skipped (logged but not moved or deleted) and recorded in the returned `skippedEntries: Array<{rowIndex, spreadsheetId}>`.
4. After each row move, tries to move the person's Drive folder (named after the first column value) from the source `DATA_FOLDER` to the destination `DATA_FOLDER` using `DriveApp.getFolderById()` / `getFoldersByName()` / `moveTo()`. If either `DATA_FOLDER` is not configured, or the named folder is not found, the row move still completes and a note is logged. `parseDriveId(value)` strips any Drive URL wrapper from the cell value to get a bare folder ID.

`movePersonnel` returns `{ log, movedRows, skippedEntries }`. The client (`runMove()`/`onMoveSuccess()` in `WebEditor.js.html`) stashes the exact `rowEntries` it sent as `pendingMoveEntries`, then on success subtracts `skippedEntries` from it to get the entries that were *actually* hard-deleted. This matters for two reasons: (1) skipped rows must stay in `schema.rows` since they're still physically in the source sheet, and (2) every other remaining row sharing a skipped/moved row's source `spreadsheetId` needs its cached `rowIndex` decremented by the count of actually-deleted rows above it — otherwise a later action (move/edit/delete) on a row below a previous move's deletion point will address the wrong physical sheet row, since `sheet.deleteRow()` shifts all subsequent row numbers up by one. New code that hard-deletes rows from a source sheet on the client's behalf must apply the same `rowIndex` adjustment to the rest of `schema.rows` (see `onDeleteSuccess()` for the single-row equivalent).

`WebEditor.js.html` has a parallel client-side implementation, `extractDriveId()` (used for image-field preview/lightbox). Its URL-matching patterns are kept aligned with `parseDriveId()`'s — update both together when changing supported Drive URL formats. They still differ in fallback behavior: `parseDriveId()` returns trusted Handbook config values unchanged if no pattern matches, while `extractDriveId()` validates untrusted user-typed text against a bare-ID length check and returns `null` on failure.

### Soft delete (Trash sheet)

`deleteRow()` never permanently removes data. It copies the row to the `Trash` sheet and then deletes it from `Database`. If `Trash` doesn't exist yet, it is created with the same two header rows as `Database`. There is no restore UI — recovery requires manually moving rows back in Sheets.

### Phone number normalization (`fixPhoneNumbers`)

Triggered from the "More... ⭐️" custom menu (added by `onOpen()` alongside "Open Web Editor"). Scans the `COL_PHONE_NUMBER` (`Номер телефону`) column of `Database` and rewrites two malformed shapes in place: bare 9-digit numbers missing the leading `0`, and 12-digit numbers carrying a `38` country-code prefix. Numbers already in canonical 10-digit form are left untouched. Runs synchronously over the whole sheet; reports a count via `ui.alert()`. No undo beyond manual edit or Trash recovery.

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

Export is triggered from the toolbar. The Export F-1 and Export WC buttons are disabled until at least one row is checked via the selection checkbox column. `runExport()` collects `{ rowIndex, spreadsheetId }` entries from `selectedKeys`-checked rows. When the count exceeds `EXPORT_CONFIRM_THRESHOLD`, it shows a confirmation dialog (with a time estimate derived from `EXPORT_SECONDS_PER_DOC`) before starting; otherwise it proceeds immediately. The actual progress/batching logic lives in `startExportProgress()`, which `runExport()` calls either directly (small batch) or from the confirm dialog's "Export" button handler (large batch).

`exportF1(rowEntries)` and `exportWC(rowEntries)` copy Google Docs templates into the export folder and fill placeholders with row data. Both delegate to `_exportDoc()`, which caches sheet data per `spreadsheetId` (so each remote spreadsheet is read at most once per export call) and runs four passes in strict order:

1. **Image columns** — `{Column Name}` placeholder replaced with the actual image blob.
2. **Service history table** — `{Проходження служби}` placeholder row expanded into one table row per service entry. Must run before pass 3 or the placeholder text would be consumed before the table handler can locate it.
3. **Direct text columns** — remaining `{Column Name}` placeholders replaced with cell text.
4. **Correspondence table** — Handbook-defined aliases and computed values (e.g. `totalServiceLength`, `motherFullName`) replace their own placeholders.

Placeholders with no match are left untouched. `_exportDoc()` checks elapsed time against `EXPORT_TIME_LIMIT_MS` before each row (5 min — 1 min safety margin before the GAS 6-min hard kill) and returns `{ results, remaining }` so the client can surface any unprocessed rows.

## Constants — always in `Config.js`

All tuneable values belong in `Config.js`. Constants needed by the client must also be threaded through the `getSchemaAndData()` return value in `Code.js` (see `filterDebounceMs`, `imageFetchBatchSize`, `imageFetchConcurrency`, `masterMode` as examples). Never hardcode magic numbers in `WebEditor.js.html`.
