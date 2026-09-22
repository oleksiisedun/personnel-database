# Web editor internals and Database-sheet utilities

Read this before touching the edit view, save/validation flow, `*-table` rendering, dropdown types, image loading, Trash, or the "More... ⭐️" menu fixers. Paths are relative to `src/`; client code is in the `WebEditor*.js.html` fragments (core, list, images, tables, edit, export, move — see the File map in CLAUDE.md).

## Edit view layout

`openEditView(row)` builds `#edit-form` as sibling `.edit-panel` divs, not one grid:

- one `.edit-panel--grid` panel (`grid-template-columns: 1fr 1fr` when active) with every non-table, non-image field;
- one more `.edit-panel--grid` panel (`dataset.tabKey = 'documentPhotos'`) with every `image`-type field — created only if `schema.columns` has at least one `image` column, and built before the `*-table` panels so it's the first dynamic tab;
- one `.edit-panel--table` panel per `*-table` column, holding that column's `buildTableEditor()` output.

Only the panel matching `activeEditTab` gets `.active` (`display:grid`/`display:block`); the others are `display:none` **but stay mounted for the whole edit session**. `collectFormValues()`/`encodeTableEditor()` query by ancestry/`data-col-index`, not visibility, so destroying a hidden panel would lose its in-progress edits.

The tab strip `#edit-tabs` sits in the sticky `#edit-header`, right of the person's name. `#edit-title` is itself the first tab (`class="edit-tab"`, click → `switchEditTab('main')`; its listener is attached once in `init()` since the element persists). One `.edit-tab` button is built for the `documentPhotos` panel (label `schema.documentPhotoTabName`, from `Config.js`'s `DOCUMENT_PHOTO_TAB_NAME` via `getSchemaAndData()`) and one per `*-table` column, freshly created on every `openEditView()`. `switchEditTab(tabKey)` (`'main'`, `'documentPhotos'`, or a table column index as a string) toggles `.active` on the panel, tab button and `#edit-title`; `openEditView()` calls `switchEditTab('main')` before showing the view so every row opens on the main tab.

## Unsaved changes confirmation

`openEditView(row)` captures `originalFormValues = collectFormValues()` right after building the form DOM (before `showView('edit')`), so the snapshot goes through the same serialization as any later read. This matters because `*-table` fields round-trip through `encodeTableEditor()`/`mergeData()`, which can reformat a cell string even with zero edits — comparing against raw `row.values` would give false positives.

`collectFormValues()` (also used by `saveRow()`) and `formValuesChanged()` (`JSON.stringify` compare against the snapshot) live next to `saveRow()`. `#btn-back` calls `formValuesChanged()`: unchanged → `leaveEditView()` (`isNewRow = false; showView('list');`); dirty → `#unsaved-changes-overlay` (Keep Editing / Discard Changes), where Discard calls `leaveEditView()`. Reuse `formValuesChanged()` for any "did the user change anything" check.

## Save-time validation and errors

`saveRow()` validates every `VALIDATORS`-typed field (`date`, `tin`, `number` — plain columns and `*-table` sub-columns) across **every** edit tab, not just `activeEditTab`, since a field can hold a pre-existing invalid value the user never touched (e.g. a malformed date typed directly into the sheet). It collects every failure as `{ tabKey, fieldName, message }` (`tabKey` via `inp.closest('.edit-panel').dataset.tabKey` for plain/photo fields, or the table column index as a string for sub-columns) instead of stopping at the first. On failure it `switchEditTab()`s to the *first* invalid field's tab and shows a single `alert()` naming every invalid field, its tab (`getEditTabLabel()`) and format hint.

Server failures from `updateRow()`/`addRowWithData()`/`deleteRow()` (`onSaveError`, `onDeleteError`) also use `alert()`. **Do not add a persistent error element inside `#edit-header`**: any text there reflows `#edit-header-actions` and shifts the Save/Delete buttons and tab strip (an earlier `#save-error` span did exactly this).

## `*-table` column widths

`buildTableEditor()` (edit view) and `buildMiniTable()` (list view) prepend a `<colgroup>` with each `<col>` width as a percentage proportional to the **square root** of the longest string in that column (header or any value, minimum 8 chars). Square-root scaling stops very long columns (address) from squeezing short important ones (phone). `mini-table` uses `table-layout: fixed` so the browser honours the colgroup.

A single-word mini-table header that's still too wide doesn't wrap mid-word or truncate: `shrinkOverflowingSingleWordHeaders()` (run after each `buildMiniTable()` render) scales its font-size down to fit on one line, floor 8px. Multi-word headers wrap normally at spaces (`word-break: normal` on `.mini-table thead th`).

## Dropdown column types and typed `*-table` sub-columns

Dropdown types are entirely data-driven — adding one is a Handbook edit, not a code change. `getDataTypeOptionsMap()` (`Utils.js`) reads `HANDBOOK_DATA_TYPES_RANGE` (`L2:AL`) into `{ typeName: values[] }`, skipping rows with no values (so `text`, `date`, `image`, `tin`, `number` are documentation only, while `unit`, `origin`, `marital-status`, `sex`, `blood-type` become dropdowns). `getSchemaAndData()` attaches the list as `col.dropdownOptions` to any top-level column whose `type` matches. The client never branches on a type-name string: `openEditView()` checks `col.dropdownOptions`, and `buildDropdownField()`/`populateSelectOptions()` render the `<select>`.

`*-table` sub-columns are typed the same way. `getTableColumnsMap()` (`Utils.js`) reads `HANDBOOK_TABLE_COLUMNS_RANGE` (`H2:J`, `Table Type | Column Name | Column Type`) into `{ tableType: [{name, type}, ...] }`; `getSchemaAndData()` builds `col.tableHeaders` as `Array<{name, type, dropdownOptions?}>`, attaching `dropdownOptions` whenever the sub-column's `type` is in the Data Types table. `addTableEditorRow()` renders a `<select>` for a sub-column with `dropdownOptions`, else a text input (this is how `medical-table`'s `Група крові` becomes a blood-type dropdown). `tableHeaderNames(col)` returns bare names where `parseData()`/`mergeData()` need strings. `buildListView()` needs no change — filter inputs render for all column types.

## Image loading

Images are fetched server-side via `DriveApp` (script owner's OAuth token) so users without personal Drive access can still see them. The client caches in two layers:

1. **IndexedDB** (`pdb_images` / `images` store) — persists across dialog sessions. `openImageDB()` is called eagerly in `init()`; `loadCacheFromIndexedDB()` pre-populates `imageCache` before `renderList()`; `saveToIndexedDB()` writes each fetched result. Entries expire after `IMAGE_CACHE_TTL_DAYS` (default 7). IndexedDB, not localStorage, because the base64 dataset exceeds localStorage's ~5 MB quota.
2. **In-memory `imageCache`** — session-only fileId → result map for O(1) lookups during rendering and lightbox opens.

Loading uses a **concurrency pool** (`IMAGE_FETCH_CONCURRENCY`, `IMAGE_FETCH_BATCH_SIZE` in `Config.js`); too much concurrency exceeds Apps Script's ~30 concurrent-execution limit and drops images (tuning notes are in `Config.js`). `DriveApp.getFileById()` throws for inaccessible files; the server catches it and returns `{ type: 'no-access' }`, which the client renders as a gray "No access" badge.

`extractDriveId()` in `WebEditor.images.js.html` (image preview/lightbox) mirrors `parseDriveId()`'s URL patterns — update both together when supporting a new Drive URL format. They differ in fallback: `parseDriveId()` returns trusted Handbook values unchanged when nothing matches, while `extractDriveId()` validates untrusted typed text against a bare-ID length check and returns `null`.

## Soft delete (Trash sheet)

`deleteRow()` never permanently removes data: it copies the row to `Trash` (created on first use with the same two header rows as `Database`), then deletes it from `Database`. There is no restore UI — recovery means manually moving rows back in Sheets.

## "More... ⭐️" menu fixers

Both live in `DataFixes.js` and share `_normalizeDatabaseColumn()`. They run synchronously over the whole `Database` sheet, locate their column via `findColumnIndex()` (`SchemaHelpers.js`), report a count via `ui.alert()`, and have no undo beyond manual edit or Trash recovery. Already-normalized values are left untouched.

- **`fixPhoneNumbers()`** — in the `COL_PHONE_NUMBER` (`Номер телефону`) column, rewrites bare 9-digit numbers missing the leading `0` and 12-digit numbers with a `38` prefix.
- **`fixFullNames()`** — in the `COL_FULL_NAME` (`ПІБ`) column, applies `normalizeFullName()` (`Formatting.js`): trims, collapses whitespace runs (including newlines, via `WHITESPACE_RUN_REGEX`) to one space, uppercases the first word (surname).
