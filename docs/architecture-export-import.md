# Export and import (`Export*.js`, `Import.js`)

Read this before touching Doc export, XLSX export, S-КАДР photo export, or award import. Paths are relative to `src/`.

File split: `Export.js` (F-1/WC document export, shared `_makeSheetDataLoader`, template placeholders, export images), `ExportValues.js` (the pure `_compute*` computed-value functions, `_pluralizeUk`, `_calendarDuration` — no Apps Script services, unit-tested in `tests/export-values.test.js`; **keep new computed values pure and testable here**), `ExportXlsx.js` (`exportXLSX` and its helpers), `ExportPhotos.js` (`startPhotoExport`/`copyPhotosBatch`). The `*-table` cell codec `_parseSubTable`/`_encodeSubTable` lives in `Utils.js` because `Import.js` uses it too.

## Document export (F-1 / Wanted Card)

Triggered from the toolbar. Export F-1 and Export WC stay disabled until at least one row is checked. `runExport()` collects `{ rowIndex, spreadsheetId }` entries via `getSelectedRowEntries()` (also used by `runMove()`, `deletePersonnelBulk()`, `openMoveOverlay()` — any client action on the checkbox selection should use it, or `getSelectedRows()` when the full row object is needed). Above `EXPORT_CONFIRM_THRESHOLD` it shows a confirmation dialog with a time estimate from `EXPORT_SECONDS_PER_DOC`; otherwise it proceeds immediately. Progress/batching lives in `startExportProgress()`.

`exportF1(rowEntries)` / `exportWC(rowEntries)` copy Docs templates into the export folder and fill placeholders. Both delegate to `_exportDoc()`, which caches sheet data per `spreadsheetId` and runs four passes in **strict order**:

1. **Image columns** — `{Column Name}` replaced with a compressed image blob (see below).
2. **Service history table** — the `{Проходження служби}` placeholder row is expanded into one table row per entry. Must run before pass 3, or the placeholder text would be consumed before the table handler can find it.
3. **Direct text columns** — remaining `{Column Name}` placeholders replaced with cell text. `*-table` columns are skipped here (their raw pipe/newline-encoded storage string is never a meaningful replacement); passes 2 or 4 must resolve them.
4. **Correspondence table** — Handbook-defined aliases and computed values (e.g. `totalServiceLength`, `motherFullName`, `awardsList` — a semicolon-joined sentence from the "Нагороди" sub-table, each entry `{name} №{order number} від {order date}` with empty number/date parts omitted) replace their own placeholders.

Placeholders with no match are left untouched. `_exportDoc()` checks elapsed time against `EXPORT_TIME_LIMIT_MS` (5 min — 1 min margin before the GAS 6-min kill) before each row and returns `{ results, remaining }` so the client can surface unprocessed rows.

**Image blobs** come from `_getExportImageBlob(fileId)`, not `DriveApp.getFileById(fileId).getBlob()`: Apps Script embeds a blob's actual bytes regardless of the display size set afterward, so full-resolution originals bloated the Docs. It calls the Advanced Drive Service (`Drive.Files.get(fileId, {fields: 'thumbnailLink'})`, enabled via `enabledAdvancedServices` in `appsscript.json`), rewrites the size parameter to `EXPORT_IMAGE_THUMBNAIL_SIZE` px, and fetches with `UrlFetchApp` (auth via `ScriptApp.getOAuthToken()`). Any failure falls back to the original blob, so export never regresses to a blanked placeholder. Keep `EXPORT_IMAGE_THUMBNAIL_SIZE` above `IMAGE_MAX_HEIGHT`: the thumbnail governs stored bytes, `IMAGE_MAX_HEIGHT` the displayed height, and a smaller thumbnail would look soft.

## XLSX export

The "Export XLSX" toolbar button exports the checkbox selection, restricted to currently-visible columns, as **one** `.xlsx` (unlike one-Doc-per-row above). `runExportXlsx()` orders rows via `sortRowsBySourceOrder()` (as `applyFilters()` does) so output matches on-screen order. Image/PDF/folder cells become clickable `HYPERLINK()` formulas (no image embedding, no blob fetching); `*-table` cells are written verbatim in their pipe/newline storage format.

Apps Script can't author `.xlsx` bytes and there's no build step to bundle a library, so `exportXLSX(rowEntries, visibleColumnIndices)` builds the grid in a temporary `SpreadsheetApp.create()`d sheet and converts it with `_fetchXlsxExportBlob()` — `UrlFetchApp` on `https://docs.google.com/spreadsheets/d/{id}/export?format=xlsx`, authorized with `ScriptApp.getOAuthToken()`. The temp spreadsheet is always trashed in a `finally` block (inner try/catch so cleanup can't mask the real error).

**Gotchas — do not "simplify" these:**

- **Never use `Blob.getAs()`/`File.getAs()` for Sheet→xlsx.** It only supports narrow conversions (mostly PDF/images) and throws `"Конвертування ... не підтримується"`.
- **Never write a mix of text and formulas with one `Range.setFormulas()`.** It parses *every* cell as if typed in the formula bar even without `=`: a bare header word becomes a named-range lookup (`#NAME?`), multi-word text fails to parse (`#ERROR!`). The grid is written in two passes: (1) everything as plain values via `setValues()` (never reinterprets strings), then (2) individual `setFormula()` calls only for cells that resolved to `HYPERLINK()`.
- **New spreadsheets default to Automatic number formatting**, which mangles text-shaped values (phone/`tin`/`number` lose leading zeros, date-like strings become serials). `exportXLSX()` force-formats the whole data range as plain text (`'@'`) before `setValues()`, then resets only the hyperlink cells to `'General'` (one `sheet.getRangeList(a1Addresses).setNumberFormat('General')` call) so their formulas evaluate.

Link resolution reuses `getImagesDataUrls()`'s mimetype classification, extracted as `_classifyDriveFile(file, fileId)` (`Code.js`). `resolveDriveFileForExport(fileId)` is the blob-free wrapper `exportXLSX()` calls (cached per fileId via `_buildXlsxLinkCell()`'s `resolveDriveInfo` closure). Rule per cell: `image`-type columns always attempt resolution (bare Drive ID or URL); other non-table columns only when the raw value `looksLikeDriveUrl()` (shares `DRIVE_URL_REGEX` with `parseDriveId()`); `*-table` columns never. An unresolvable file falls back to the raw value as plain text.

Unlike `_exportDoc()`, there is no partial/continuation result: `exportXLSX()` throws if `EXPORT_TIME_LIMIT_MS` is exceeded, since a half-built spreadsheet has no standalone value. This should rarely trigger — there's no per-row blob fetch.

`_makeSheetDataLoader(localSs)` is the per-spreadsheet cache/loader factory shared by `_exportDoc()`, `exportXLSX()`, and `startPhotoExport()`, so a multi-source export reads each spreadsheet at most once. Reuse it for any new export that reads each selected row's source spreadsheet.

## Photo export for S-КАДР

The "Export Photos for S-КАДР" menu item is shown only in Master Mode (`getMasterMode()`). It calls `openPhotoExport()` (`Code.js`), which opens the *same* `WebEditor` modal with `template.mode = 'photoExport'`; `init()` in `WebEditor.js.html` detects `INITIAL_MODE` and drives the shared `#export-overlay` progress UI instead of the normal bootstrap. Because it never calls `getSchemaAndData()`, `schema` stays `null` for the life of the dialog — none of the photo-export functions may read it. `#btn-export-close` closes the whole dialog (`google.script.host.close()`) in this mode.

There is no manual selection: it sweeps **every** row in the Actual Personnel list across the local `Database` sheet and every Master Mode source. To stay under the 6-minute ceiling it works in two phases:

1. **Discovery** — `startPhotoExport()` creates the destination folder (`PHOTO_EXPORT_FOLDER_PREFIX` + `formatDateDDMMYYYY()`, inside `Handbook!A13`), then walks local + every `getMasterSources()` ID **in that fixed order** — deliberately not a concurrent pool, because duplicate-ID resolution needs reproducible order. Per source (via `_makeSheetDataLoader()`), it filters to Actual Personnel and reads `COL_PHOTO`/`COL_CARD_ID` via `getFieldByPattern()`. A person missing either value is skipped with a reason; a repeated S-КАДР ID (e.g. across two sources) is skipped as a duplicate — first occurrence wins. No image bytes are fetched, so the full eligible list is known before any Drive I/O. Like `exportXLSX()` it has no continuation and throws past `EXPORT_TIME_LIMIT_MS`.
2. **Copy/convert** — `copyPhotosBatch(folderId, entries)` uses the same time-boxed `{results, remaining}` continuation as `_exportDoc()`, plus a `skipped` field (per-item failures must be shown to the user). It fetches the **full-resolution** original (`DriveApp.getFileById(fileId).getBlob()` — deliberately *not* the 800px `_getExportImageBlob()` path, since this feed needs source quality), force-converts with `blob.getAs('image/jpeg')`, and saves `{S-КАДР ID}.jpg`. A fetch/convert failure becomes a skip entry, not an abort.

Client flow: `startPhotoExportFlow()` → `onPhotoDiscoverySuccess()` (skips the loop if nothing is eligible) → `runPhotoCopyBatch()` (recurses on `response.remaining`, like `startExportProgress()`'s `fire()`) → `finishPhotoExport()`, which links the destination folder and renders the skip/duplicate list via `renderPhotoSkipLog()` into `#export-skip-log` (`.log-list`).

## Award import from S-КАДР (`Import.js`)

The "Import awards from S-КАДР" menu item (Master Mode only, next to photo export) calls `importAwards()`. It uses plain `ui.prompt()`/`ui.alert()` rather than the `WebEditor` modal: it's a synchronous sheet-to-sheet scan with no per-row Drive I/O, so it finishes in one execution.

- **Source**: prompts for the sheet's full URL, resolves it via `parseDriveId()` + `openSpreadsheetSafely()`, and picks the tab by matching `?gid=`/`#gid=` (`parseGidFromUrl()`) against `sheet.getSheetId()`, falling back to the first tab.
- **Layout is fixed by column position**, not header text: `AWARDS_IMPORT_ID_COL`/`_NAME_COL`/`_ORDER_NUMBER_COL`/`_ORDER_DATE_COL` (A1 letters `'A'`/`'F'`/`'G'`/`'H'`), data from row 2. `_groupAwardsById()` converts letters once via `columnLetterToIndex()`, skips rows without ID or award name, capitalizes the name's first letter, and cleans the order number with `AWARDS_ORDER_NUMBER_CLEAN_REGEX` (`\W` is ASCII-only, so it strips Cyrillic text, spaces and `№` plus a trailing `/2022`-style suffix: `"Указ Президента України № 559/2022"` → `"559"`). Result: `{ id: [name, orderNumber, orderDate][] }`.
- **Matching** searches the local `Database` plus every `getMasterSources()` spreadsheet (`_collectAwardTargetSources()` → `_addAwardSource()`), locating `S-КАДР ID`/`Нагороди` via `findColumnIndex()` (`Utils.js`, also used by `fixPhoneNumbers()`/`fixFullNames()`) and precomputing an `id → row index` map per source.
- **Applying**: `_applyAwards()` resolves each ID to the first source that has it (local first, then `getMasterSources()` order — same "first wins" as photo export), decodes the existing `Нагороди` cell via `_parseSubTable()`, appends new entries (skipping exact `[name, number, date]` duplicates), and writes back via `_encodeSubTable()` (inverse of `_parseSubTable()`). Their separators are `TABLE_FIELD_SEP`/`TABLE_ROW_SEP` in `Config.js`; the client (`WebEditor.tables.js.html`) declares its own copies of the same two literals because it runs in a separate script — keep them in sync.
- **Reporting**: the final `ui.alert()` shows counts (updated, added, duplicates skipped, not found) plus up to `AWARDS_IMPORT_NOT_FOUND_DISPLAY_LIMIT` not-found IDs by name; the full per-ID log goes to `Logger.log()` since alerts aren't scrollable.
