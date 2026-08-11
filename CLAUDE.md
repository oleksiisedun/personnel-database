## Deployment

```bash
# Push to the primary bound script project
clasp push

# Push to all target spreadsheets listed in clasp-targets.json
./clasp-push.sh
```

There is no build step, linter, or test suite.

**Never run `clasp push`/`clasp-push.sh` or otherwise deploy/test changes yourself.** These scripts push live to real bound spreadsheets (including production personnel data across all targets in `clasp-targets.json`). Leave deployment and live testing to the user.

## Code conventions

JS style and JSDoc rules are in the global `~/.claude/CLAUDE.md`. Project-specific rule: all tuneable constants belong in `Config.js` (see below).

### CSS design system (`WebEditor.css.html`)

Concrete tokens and shared classes already defined here (see global CSS design-system conventions for the reuse-first principle):

- **Variables**: `--color-primary`, `--color-primary-hover`, `--color-focus`, `--color-focus-shadow`, `--color-danger`, `--color-danger-hover`, `--color-success`, `--color-warning`, `--color-warning-hover`, `--color-border`, `--radius`.
- **Shared classes**: `.btn-primary` (filled accent — Save/Move/Close), `.btn-secondary` (outline — Back/Cancel/toolbar), `.btn-danger` (outline, danger color — Delete, Discard Changes), `.btn-success` (outline, success color — Refresh), `.btn-warning` (outline, warning color — schema-mismatch ⚠ button, Reset), `.btn-dialog-action` (sizing only — `padding`/`font-size` for overlay action buttons: export Cancel/Proceed/Close, move Cancel/Confirm, schema-warning Close, unsaved-changes Keep Editing/Discard Changes — combined with a color class like `.btn-primary`/`.btn-secondary`/`.btn-danger`), `.btn-toolbar` (sizing only — `padding: 5px 12px` for every button inside `#toolbar` — Add Person, Move, Delete Personnel, Export F-1/WC/XLSX, ⚠ schema-warning, Reset, Refresh, Columns ▾ — combined with a color class the same way `.btn-dialog-action` is), `.overlay`/`.overlay-dialog` (modal scaffolding — export progress, move dialog, unsaved-changes dialog), `.dialog-actions` (sizing-agnostic `display:flex; gap:10px; justify-content:flex-end` action row — used by the unsaved-changes dialog; new overlay action rows should use this instead of a one-off `#foo-actions` ID rule).
- **Global states**: `button:disabled` (`opacity: 0.5; cursor: default;`), text input/select focus (`border-color: var(--color-focus); box-shadow: 0 0 0 2px var(--color-focus-shadow);`).
- **Every new toolbar button must add `btn-toolbar` to its `class` attribute**, alongside a color class (`class="btn-secondary btn-toolbar"` etc.) — this is not optional cosmetic polish. `.btn-toolbar` used to be a single `#btn-foo, #btn-bar, ...` ID-selector list that every new toolbar button had to be manually appended to; it was forgotten at least once (`#btn-export-xlsx` shipped without it, rendering with the browser's default button padding instead of matching its neighbors) before being converted to this class. Do not reintroduce a per-ID padding list for toolbar buttons — it is a duplication trap for exactly this reason.

## Architecture

This is a **Google Apps Script** project (V8 runtime) bound to a Google Spreadsheet. The web editor runs inside an `HtmlService` modal dialog opened from the Sheets menu.

### Sheet layout

The `Database` sheet has this fixed structure:
- Row 1 — column names (header)
- Row 2 — column types (`text`, `image`, `*-table`, `unit`, `origin`, `marital-status`, `sex`, `tin`, `number`)
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

Handbook cells that hold a Drive ID or sharing URL (`DATA_FOLDER`, `EXPORT_F1_TEMPLATE_CELL`, `EXPORT_WC_TEMPLATE_CELL`, `EXPORT_FOLDER_CELL`) are read via `getDriveIdFromHandbook(handbookSheet, cellAddress)` in `Utils.js` — it trims the cell value and pipes it through `parseDriveId()`, returning `''` if the sheet is missing. Used by `movePersonnel()` (source/destination `DATA_FOLDER`) and `_exportDoc()` (template and export-folder IDs). New code reading a Drive ID from a Handbook cell should use this helper instead of repeating the `handbookSheet ? parseDriveId(...) : ''` pattern inline.

The `COL_DRAFT_DATE`, `COL_SERVICE_HISTORY`, `COL_CLOSE_RELATIVES`, `COL_MARITAL_STATUS`, `COL_CONTRACT_UNTIL`, and `COL_PHONE_NUMBER` constants in `Config.js` are case-insensitive regexes (e.g. `/дата призову/i`), not exact strings, so that header-name matching tolerates casing drift across Master Mode source spreadsheets. Because of this, code can no longer use them as a direct object key (`data[COL_X]`) or compare with `===`. `findKeyByPattern(obj, pattern)` and `getFieldByPattern(data, pattern)` in `Utils.js` resolve the actual matching key from a row-data map — `getFieldByPattern()` for the cell value (used throughout `Export.js`'s correspondence/computed-value functions), `findKeyByPattern()` when the literal matched header text itself is needed (used by `_fillServiceHistoryTable()` to build the `{header}` placeholder it searches for in the exported Doc). `Code.js`'s `fixPhoneNumbers()` uses `COL_PHONE_NUMBER.test(...)` for the same reason. New code matching a column by one of these constants should use these helpers instead of exact-match comparisons.

If a user reports this exact `PERMISSION_DENIED` error despite having real access to everything involved, check for multiple signed-in Google accounts in their browser before chasing a code fix — the dialog's `google.script.run` calls can authenticate against the wrong signed-in account, producing the same error as a genuine sharing gap. Confirmed fix in one case: log out of all Google accounts except the one with access.

### Master Mode

When `Handbook!M2` is `true`, `getSchemaAndData()` reads source spreadsheet IDs from `Handbook!N2:N` via `getMasterSources()` and returns them as `masterSourceIds: string[]` — it does **not** open any remote spreadsheet itself. Local rows have no `spreadsheetId` property; `masterSources` is seeded with just the local entry (`{ id: null, name: <current spreadsheet name> }`).

After the initial render, the client (`onDataLoaded()` in `WebEditor.js.html`) calls `queueMasterSourceFetch()`, which fetches each remote source's rows one at a time through a worker pool (size `masterModeFetchConcurrency`/`MASTER_MODE_FETCH_CONCURRENCY`, mirroring the image-fetch concurrency pool below) via the server function `getMasterSourceRows(spreadsheetId)`. Before returning rows, `getMasterSourceRows()` validates the remote sheet's column schema against the local `Database` sheet using `compareColumnSchemas()` (in `Utils.js`), which compares column names and types at every index and skips trailing blank positions (from `getDataRange()` overreach). If mismatches are found, `getMasterSourceRows()` returns `rows: []` and a non-null `columnMismatches: Array<{colIndex, localName, localType, remoteName, remoteType}>` instead of the actual rows. The client records these in `schema.sourceColumnMismatches` (a plain object keyed by `spreadsheetId`) and reveals the amber `#btn-schema-warning` button (⚠, left of the "All Units" checkbox); clicking it opens an overlay listing each mismatched source by name and each differing column by letter, local header, and remote header. Sources with matching schemas have their rows pushed onto `schema.rows` normally. Inaccessible sources are still skipped silently with `columnMismatches: null` — they are never flagged as mismatched. The `{id, name}` entry is always pushed onto `schema.masterSources` regardless of mismatch, so the warning overlay can resolve IDs to human-readable names. New code that needs "all master-mode rows" should be aware that, on the client, they may still be arriving for a short time after `onDataLoaded()` fires.

The green `#btn-refresh` toolbar button (always visible, positioned between the `#btn-reset` button and the "All Units" checkbox — unlike Move it is not gated behind Master Mode, since re-pulling local rows is useful whenever multiple users share one spreadsheet) re-runs this whole load without closing the dialog, via `refreshAll()` → `onRefreshSuccess()` in `WebEditor.js.html`. Unlike `onDataLoaded()`, `onRefreshSuccess()` does not call `buildListView()` (would wipe typed filter text) or re-attach the `#chk-all-units`/`#chk-actual` change listeners (would stack duplicates) — it only replaces the data fields on `schema` (`rows`, `masterSources`, `masterSourceIds`, `sourceColumnMismatches`, `actualPersonnelNames`/`actualPersonnelNamesSet`), clears `selectedKeys` (row positions may have shifted upstream), and re-renders via `applyFilters()`. Both `onDataLoaded()` and `onRefreshSuccess()` share `buildActualPersonnelNamesSet(names)` and `maybeQueueMasterSourceFetch(data)` for this — new code touching either code path should reuse these helpers rather than re-inlining the logic. Because a refresh can be clicked again while a previous refresh's `queueMasterSourceFetch()` is still streaming remote rows in, `queueMasterSourceFetch()` captures a `refreshGeneration` counter (bumped by every `refreshAll()` call) at start and checks it before every mutation of `schema`; a stale run whose generation no longer matches silently stops instead of pushing duplicate/outdated rows onto the new `schema.rows`. While the refresh's server round-trip is in flight, `#btn-refresh` is disabled and `#data-table` gets the `.filtering` dimming class (same class `resetListState()` uses), cleared in both `onRefreshSuccess()` and `onRefreshError()`.

The orange `#btn-reset` toolbar button (always visible, positioned between `#btn-schema-warning` ⚠ and `#btn-refresh` — unlike Move it is not gated behind Master Mode, since filters/columns/selections exist regardless) calls `resetListState()` in `WebEditor.js.html`. It clears the three pieces of client-only view state — column filter text, the `useRegex`/`showAllUnits`/`filterActualPersonnel` toggles (and their checkboxes), `hiddenColumns`, and `selectedKeys` — back to their defaults and re-renders via `buildColumnPanel()` + `applyFilters()`. It never touches `schema.rows` and makes no server call, unlike `refreshAll()`. While the reset runs, `#btn-reset` is disabled and `#data-table` gets the `.filtering` dimming class via the shared `dimTableWhile(fn)` helper in `WebEditor.js.html` (also used by `onDataLoaded()`'s `#chk-all-units`/`#chk-actual` listeners) — it wraps `fn` in a double-`requestAnimationFrame` so the disabled/dimmed state paints before the synchronous rebuild work runs. New code that needs to dim the table around synchronous work should reuse this helper instead of re-inlining the class-toggle/rAF pattern.

`applyFilters()` sorts `filteredRows` by source order (local first, then each remote source in its `masterSourceIds` position) and by `rowIndex` within each source before passing them to `renderList()`. This guarantees a stable display order even though `schema.rows` is appended in non-deterministic arrival order. `schema.rows` itself is never sorted — edit/delete/move logic always uses `rowIndex` + `spreadsheetId` directly and is unaffected.

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
4. After each row move, tries to move the person's Drive folder (named after the first column value) from the source `DATA_FOLDER` to the destination `DATA_FOLDER` using `DriveApp.getFolderById()` / `getFoldersByName()` / `moveTo()`. If either `DATA_FOLDER` is not configured, or the named folder is not found, the row move still completes and a note is logged. `parseDriveId(value)` strips any Drive URL wrapper from the cell value to get a bare folder ID; the source/destination `DATA_FOLDER` reads themselves go through `getDriveIdFromHandbook()` (see above).

`movePersonnel` returns `{ log, movedRows, skippedEntries }`. The client (`runMove()`/`onMoveSuccess()` in `WebEditor.js.html`) stashes the exact `rowEntries` it sent as `pendingMoveEntries`, then on success subtracts `skippedEntries` from it to get the entries that were *actually* hard-deleted. This matters for two reasons: (1) skipped rows must stay in `schema.rows` since they're still physically in the source sheet, and (2) every other remaining row sharing a skipped/moved row's source `spreadsheetId` needs its cached `rowIndex` decremented by the count of actually-deleted rows above it — otherwise a later action (move/edit/delete) on a row below a previous move's deletion point will address the wrong physical sheet row, since `sheet.deleteRow()` shifts all subsequent row numbers up by one. New code that hard-deletes rows from a source sheet on the client's behalf must apply the same `rowIndex` adjustment to the rest of `schema.rows` (see `onDeleteSuccess()` for the single-row equivalent).

`WebEditor.js.html` has a parallel client-side implementation, `extractDriveId()` (used for image-field preview/lightbox). Its URL-matching patterns are kept aligned with `parseDriveId()`'s — update both together when changing supported Drive URL formats. They still differ in fallback behavior: `parseDriveId()` returns trusted Handbook config values unchanged if no pattern matches, while `extractDriveId()` validates untrusted user-typed text against a bare-ID length check and returns `null` on failure.

### Soft delete (Trash sheet)

`deleteRow()` never permanently removes data. It copies the row to the `Trash` sheet and then deletes it from `Database`. If `Trash` doesn't exist yet, it is created with the same two header rows as `Database`. There is no restore UI — recovery requires manually moving rows back in Sheets.

### Phone number normalization (`fixPhoneNumbers`)

Triggered from the "More... ⭐️" custom menu (added by `onOpen()` alongside "Open Web Editor"). Scans the `COL_PHONE_NUMBER` (`Номер телефону`) column of `Database` and rewrites two malformed shapes in place: bare 9-digit numbers missing the leading `0`, and 12-digit numbers carrying a `38` country-code prefix. Numbers already in canonical 10-digit form are left untouched. Runs synchronously over the whole sheet; reports a count via `ui.alert()`. No undo beyond manual edit or Trash recovery.

### Full name normalization (`fixFullNames`)

Also triggered from the "More... ⭐️" custom menu, following the same header-lookup pattern as `fixPhoneNumbers`. Scans the `COL_FULL_NAME` (`ПІБ`) column of `Database` and rewrites each value via `normalizeFullName()` (`Utils.js`): trims surrounding whitespace, collapses internal whitespace runs — including newlines — to a single space via `WHITESPACE_RUN_REGEX`, and uppercases the first word (the surname). Already-normalized values are left untouched. Runs synchronously over the whole sheet; reports a count via `ui.alert()`. No undo beyond manual edit or Trash recovery.

### Edit view layout

`#edit-form` is a two-column CSS grid (`grid-template-columns: 1fr 1fr`). Most field blocks occupy one column. Table-type fields get the additional class `field-block--full` (→ `grid-column: 1 / -1`) so they span the full width, since their multi-row editors are too wide for a half-width cell.

### Unsaved changes confirmation

`openEditView(row)` captures `originalFormValues = collectFormValues()` right after the form DOM is built (before `showView('edit')`), so the snapshot goes through the same serialization as any later read — this matters because `*-table` fields round-trip through `encodeTableEditor()`/`mergeData()`, which can reformat a cell's string slightly even with zero edits, so comparing against the raw `row.values` instead would risk false positives. `collectFormValues()` (also used by `saveRow()` to build its save payload) and `formValuesChanged()` (`JSON.stringify` comparison against the snapshot) live in `WebEditor.js.html` next to `saveRow()`. Clicking `#btn-back` calls `formValuesChanged()`: if unchanged, it goes straight to `leaveEditView()` (`isNewRow = false; showView('list');`); if dirty, it shows the `#unsaved-changes-overlay` (Keep Editing / Discard Changes), and Discard calls `leaveEditView()`. New code that needs "did the user change anything in the edit form" should reuse `formValuesChanged()` rather than re-deriving it.

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

Export is triggered from the toolbar. The Export F-1 and Export WC buttons are disabled until at least one row is checked via the selection checkbox column. `runExport()` collects `{ rowIndex, spreadsheetId }` entries from `selectedKeys`-checked rows via `getSelectedRowEntries()` (also used by `runMove()` and `deletePersonnelBulk()` — any client action that operates on the current checkbox selection should use this same helper). When the count exceeds `EXPORT_CONFIRM_THRESHOLD`, it shows a confirmation dialog (with a time estimate derived from `EXPORT_SECONDS_PER_DOC`) before starting; otherwise it proceeds immediately. The actual progress/batching logic lives in `startExportProgress()`, which `runExport()` calls either directly (small batch) or from the confirm dialog's "Export" button handler (large batch).

`exportF1(rowEntries)` and `exportWC(rowEntries)` copy Google Docs templates into the export folder and fill placeholders with row data. Both delegate to `_exportDoc()`, which caches sheet data per `spreadsheetId` (so each remote spreadsheet is read at most once per export call) and runs four passes in strict order:

1. **Image columns** — `{Column Name}` placeholder replaced with a compressed image blob (see below).
2. **Service history table** — `{Проходження служби}` placeholder row expanded into one table row per service entry. Must run before pass 3 or the placeholder text would be consumed before the table handler can locate it.
3. **Direct text columns** — remaining `{Column Name}` placeholders replaced with cell text.
4. **Correspondence table** — Handbook-defined aliases and computed values (e.g. `totalServiceLength`, `motherFullName`) replace their own placeholders.

Placeholders with no match are left untouched. `_exportDoc()` checks elapsed time against `EXPORT_TIME_LIMIT_MS` before each row (5 min — 1 min safety margin before the GAS 6-min hard kill) and returns `{ results, remaining }` so the client can surface any unprocessed rows.

Pass 1's image blob comes from `_getExportImageBlob(fileId)`, not a raw `DriveApp.getFileById(fileId).getBlob()` — Apps Script embeds a blob's actual bytes into the Doc regardless of the display size set afterward, so inserting the full-resolution original (often several MB for a phone photo) bloated exported files even though `_replacePlaceholderWithImage()` immediately shrinks the *displayed* size down to `IMAGE_MAX_HEIGHT`. `_getExportImageBlob()` instead calls the Advanced Drive Service (`Drive.Files.get(fileId, {fields: 'thumbnailLink'})`, enabled via `enabledAdvancedServices` in `appsscript.json`) to get a `thumbnailLink`, rewrites its size parameter to request a `EXPORT_IMAGE_THUMBNAIL_SIZE`-px-wide thumbnail, and fetches those bytes with `UrlFetchApp` (authorized via `ScriptApp.getOAuthToken()`). If the Drive advanced service call, the thumbnail link, or the fetch fails for any reason, it falls back to the original `DriveApp.getFileById(fileId).getBlob()` behavior — export never regresses to a blanked placeholder because of this optimization. The `EXPORT_IMAGE_THUMBNAIL_SIZE` constant is kept above `IMAGE_MAX_HEIGHT` on purpose: the thumbnail governs stored byte size, `IMAGE_MAX_HEIGHT` still governs the displayed height, and a thumbnail smaller than the display size would look soft.

### XLSX export (`Export.js`)

The "Export XLSX" toolbar button (next to Export WC) exports the current checkbox selection, restricted to currently-visible columns (i.e. not hidden via the column-visibility panel), as a single `.xlsx` file — one file for all selected rows, unlike the one-Doc-per-row model above. Row order matches on-screen order: `runExportXlsx()` sorts `getSelectedRowEntries()` through `sortRowsBySourceOrder()` (also used by `applyFilters()` for the list display) before sending them to the server. Image/PDF/folder-type cells become clickable `HYPERLINK()` formulas rather than embedded images or plain text — this was a deliberate scope cut (no image embedding needed) that keeps the export purely value/formula-based, no blob fetching required. `*-table` cells are written verbatim in their native pipe/newline-encoded storage format.

Apps Script cannot author `.xlsx` bytes directly and there is no build step to bundle a library like ExcelJS, so `exportXLSX(rowEntries, visibleColumnIndices)` builds the grid in a temporary `SpreadsheetApp.create()`d sheet and converts it via `_fetchXlsxExportBlob()`, which hits the Sheets export URL (`https://docs.google.com/spreadsheets/d/{id}/export?format=xlsx`) with `UrlFetchApp` authorized via `ScriptApp.getOAuthToken()` — the same authenticated-fetch pattern `_getExportImageBlob()` already uses for Drive thumbnails. **`Blob.getAs()`/`File.getAs()` do NOT work here** — they only support a narrow set of conversions (mainly to PDF or image formats) and throw `"Конвертування ... не підтримується"` for a Sheet→xlsx conversion; don't reintroduce that call. The temp spreadsheet is always trashed in a `finally` block (with its own inner try/catch so a cleanup failure can't mask the real error), even if the export throws.

The grid is written in two passes, deliberately never mixed into a single `Range.setFormulas()` call: (1) the whole grid as plain values via `Range.setValues()`, and (2) individual `Range.setFormula()` calls only for the specific cells that resolved to a `HYPERLINK()` formula, overwriting their pass-1 placeholder. **`setFormulas()` is not a safe way to write a mix of plain text and formulas** — despite the Apps Script docs implying non-`=`-prefixed strings are stored as literal values, it actually parses *every* cell as though typed into the formula bar, even without `=`: a bare column-header word gets looked up as a named range and fails with `#NAME?`, ordinary multi-word text fails to parse at all and shows `#ERROR!`. `setValues()` has no such behavior — it never reinterprets a string as a formula regardless of content. (This was the cause of a real bug: an early version used one `setFormulas()` call for the whole grid, corrupting every plain-text cell, most visibly the header row.)

A brand-new spreadsheet also defaults to Automatic number formatting, which would silently mangle text-shaped values on write (e.g. a phone number or `tin`/`number` column value loses its leading zero, a date-like string gets reinterpreted as a real date serial) — a problem the source `Database` sheet doesn't have, since its cells already hold literal strings via `getValues()`. `exportXLSX()` works around this by force-formatting the whole destination data range as plain text (`'@'`) before the `setValues()` pass, then resetting only the specific cells that will hold a `HYPERLINK()` formula back to `'General'` format (via `sheet.getRangeList(a1Addresses).setNumberFormat('General')`, a single bulk call) so those formulas evaluate instead of displaying as literal text.

Link resolution reuses `getImagesDataUrls()`'s mimetype-based classification logic, refactored out into `_classifyDriveFile(file, fileId)` (Code.js) so it can be shared without forcing `getImagesDataUrls()`'s blob/base64 fetch onto callers that only need the view URL. `resolveDriveFileForExport(fileId)` is the lightweight, blob-free wrapper `exportXLSX()` actually calls (cached per fileId within one export call via `_buildXlsxLinkCell()`'s `resolveDriveInfo` closure in `Export.js`). Per-cell linkification rule: `image`-type columns always attempt resolution (their stored value is a bare Drive ID or URL, same convention as `_exportDoc()`'s pass 1); any other non-table column only attempts it when the raw value itself looks like a Drive URL (`looksLikeDriveUrl()`, sharing `DRIVE_URL_REGEX` with `parseDriveId()`); `*-table` columns are never linkified. An unresolvable file (no access / lookup error) falls back to the raw value as plain text rather than dropping it.

Unlike `_exportDoc()`, there is no partial/continuation result: `exportXLSX()` still checks elapsed time against `EXPORT_TIME_LIMIT_MS` while building the grid, but throws instead of returning `remaining` rows to retry, since a half-built spreadsheet has no standalone value the way a partial batch of independent Docs does. This path should rarely trigger anyway — there's no per-row blob fetch (the biggest cost in `_exportDoc()`), just a cached sheet read and an optional cached Drive metadata lookup per row.

`_makeSheetDataLoader(localSs)` (Export.js) is the per-spreadsheet cache/loader factory shared by `_exportDoc()` and `exportXLSX()`, so a Master Mode export spanning multiple sources still reads each distinct spreadsheet at most once per call. New export-style functions that need "read each selected row's source spreadsheet, cached" should reuse this instead of re-inlining the cache.

## Constants — always in `Config.js`

All tuneable values belong in `Config.js`. Constants needed by the client must also be threaded through the `getSchemaAndData()` return value in `Code.js` (see `filterDebounceMs`, `imageFetchBatchSize`, `imageFetchConcurrency`, `masterMode` as examples). Never hardcode magic numbers in `WebEditor.js.html`.
