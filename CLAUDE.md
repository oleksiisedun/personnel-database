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

The `Handbook` sheet holds schema metadata, dropdown option lists, the Master Mode toggle, and export correspondence tables. Constants in `Config.js` define each range address.

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

### Master Mode

When `Handbook!M2` is `true`, `getSchemaAndData()` reads spreadsheet IDs from `Handbook!N2:N` and appends rows from each remote spreadsheet's `Database` sheet to the local rows. Remote rows carry a `spreadsheetId` property; local rows do not.

`updateRow(rowIndex, values, spreadsheetId)` and `deleteRow(rowIndex, spreadsheetId)` route to the correct spreadsheet based on `spreadsheetId` — `SpreadsheetApp.openById()` for remote, `getActiveSpreadsheet()` for local. New rows (`addRow()`) always go to the local `Database` sheet. Editing is never restricted by Master Mode state.

### Soft delete (Trash sheet)

`deleteRow()` never permanently removes data. It copies the row to the `Trash` sheet and then deletes it from `Database`. If `Trash` doesn't exist yet, it is created with the same two header rows as `Database`. There is no restore UI — recovery requires manually moving rows back in Sheets.

### Dropdown column types

`unit`, `origin`, `marital-status`, and `sex` columns each read their allowed values from a dedicated Handbook range at load time. The pattern for adding a new dropdown type is:
1. Add a `HANDBOOK_XYZ_RANGE` constant in `Config.js`
2. Read and attach `col.xyzOptions` in `getSchemaAndData()` (same pattern as `unitOptions`)
3. Render a `<select>` in `openEditView()` in `WebEditor.js.html`
4. Add the type to the filter input condition in `buildListView()`

### Image loading

Images are fetched server-side via `DriveApp` (using the script owner's OAuth token) so users without personal Drive access can still view images. The client caches results in two layers:

1. **IndexedDB** (`pdb_images` / `images` store) — persistent across dialog sessions. `openImageDB()` is called eagerly in `init()` so the DB is ready before `onDataLoaded` fires. `loadCacheFromIndexedDB()` (async) pre-populates `imageCache` before `renderList()` runs; `saveToIndexedDB()` writes each fetched result back. Entries expire after `IMAGE_CACHE_TTL_DAYS` days (default 7). Use IndexedDB (not localStorage) because the full base64 dataset easily exceeds the ~5 MB localStorage quota.
2. **In-memory `imageCache`** — session-only map from fileId to result, used for O(1) lookups during rendering and lightbox opens.

Loading uses a **concurrency pool** controlled by `IMAGE_FETCH_CONCURRENCY` and `IMAGE_FETCH_BATCH_SIZE` in `Config.js` — raising concurrency too high exceeds Apps Script's ~30 concurrent-execution limit and causes dropped images. See the comments in `Config.js` for tuning guidance.

`DriveApp.getFileById()` throws for inaccessible files; the server catches it and returns `{ type: 'no-access' }`. The client renders a gray "No access" badge.

### Document export (`Export.js`)

`exportF1(rowIndices)` and `exportWC(rowIndices)` copy Google Docs templates into the export folder and fill placeholders with row data. Both delegate to `_exportDoc()`, which runs four passes in strict order:

1. **Image columns** — `{Column Name}` placeholder replaced with the actual image blob.
2. **Service history table** — `{Проходження служби}` placeholder row expanded into one table row per service entry. Must run before pass 3 or the placeholder text would be consumed before the table handler can locate it.
3. **Direct text columns** — remaining `{Column Name}` placeholders replaced with cell text.
4. **Correspondence table** — Handbook-defined aliases and computed values (e.g. `totalServiceLength`, `motherFullName`) replace their own placeholders.

Placeholders with no match are left untouched. `_exportDoc()` checks elapsed time against `EXPORT_TIME_LIMIT_MS` before each row (5 min — 1 min safety margin before the GAS 6-min hard kill) and returns `{ results, remaining }` so the client can surface any unprocessed rows.

## Constants — always in `Config.js`

All tuneable values belong in `Config.js`. Constants needed by the client must also be threaded through the `getSchemaAndData()` return value in `Code.js` (see `filterDebounceMs`, `imageFetchBatchSize`, `imageFetchConcurrency`, `masterMode` as examples). Never hardcode magic numbers in `WebEditor.js.html`.
