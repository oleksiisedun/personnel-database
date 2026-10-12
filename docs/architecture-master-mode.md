# Master Mode, Actual personnel filter, Move personnel

Read this before touching Master Mode loading, the "All Units"/"Actual personnel" filters, Move, or per-unit Drive folder resolution. Paths are relative to `src/`.

## Master Mode

When `Handbook!A2` is `true`, `getSchemaAndData()` reads source spreadsheet IDs from `Handbook!B2:B` via `getMasterSources()` and returns them as `masterSourceIds: string[]` — it does **not** open any remote spreadsheet itself. Local rows have no `spreadsheetId` property; `masterSources` is seeded with just the local entry (`{ id: null, name: <current spreadsheet name> }`).

After the initial render, the client (`onDataLoaded()` in `WebEditor.js.html`) calls `queueMasterSourceFetch()`, which fetches each remote source's rows one at a time through a worker pool (size `masterModeFetchConcurrency`/`MASTER_MODE_FETCH_CONCURRENCY`, mirroring the image-fetch pool) via `getMasterSourceRows(spreadsheetId)`. Rows may still be arriving for a short time after `onDataLoaded()` fires — code that needs "all master-mode rows" must account for that.

Before returning rows, `getMasterSourceRows()` validates the remote column schema against the local `Database` sheet with `compareColumnSchemas()` (`SchemaHelpers.js`; compares names and types at every index, skipping trailing blanks from `getDataRange()` overreach). On mismatch it returns `rows: []` plus `columnMismatches: Array<{colIndex, localName, localType, remoteName, remoteType}>`. The client records these in `schema.sourceColumnMismatches` (keyed by `spreadsheetId`) and reveals the amber `#btn-schema-warning` ⚠ button, whose overlay lists each mismatched source and column. Inaccessible sources are skipped silently with `columnMismatches: null` — never flagged as mismatched. The `{id, name}` entry is always pushed onto `schema.masterSources` regardless of mismatch, so the overlay can resolve IDs to names.

### Refresh and Reset

- **`#btn-refresh`** (green, always visible — not Master-Mode-gated, since multiple users can share one spreadsheet) → `refreshAll()` → `onRefreshSuccess()`. Unlike `onDataLoaded()` it does **not** call `buildListView()` (would wipe typed filter text) or re-attach the `#chk-all-units`/`#chk-actual` listeners (would stack duplicates). It only replaces the data fields on `schema` (`rows`, `masterSources`, `masterSourceIds`, `sourceColumnMismatches`, `actualPersonnelNames`/`actualPersonnelNamesSet`), clears `selectedKeys` (row positions may have shifted), and re-renders via `applyFilters()`. Both paths share `buildActualPersonnelNamesSet(names)` and `maybeQueueMasterSourceFetch(data)` — reuse them instead of re-inlining.
- **Stale-run guard**: `queueMasterSourceFetch()` captures `refreshGeneration` (bumped by every `refreshAll()`) at start and checks it before every mutation of `schema`; a stale run silently stops instead of pushing duplicate/outdated rows.
- While a refresh is in flight, `#btn-refresh` is disabled and `#data-table` gets the `.filtering` dimming class; both are cleared in `onRefreshSuccess()` and `onRefreshError()`.
- **`#btn-reset`** (orange, always visible, between ⚠ and Refresh) → `resetListState()`. Clears client-only view state — column filter text, the `useRegex`/`showAllUnits`/`filterActualPersonnel` toggles (and checkboxes), `hiddenColumns`, `selectedKeys` — then `buildColumnPanel()` + `applyFilters()`. Never touches `schema.rows`, no server call.
- `dimTableWhile(fn)` (`WebEditor.js.html`) disables the button and dims the table around synchronous work via a double-`requestAnimationFrame` so the state paints before the work blocks. Also used by the `#chk-all-units`/`#chk-actual` listeners; reuse it rather than re-inlining the class-toggle/rAF pattern.

### Ordering and routing

- `applyFilters()` sorts `filteredRows` by source order (local first, then each remote source in `masterSourceIds` order) and by `rowIndex` within a source (`sortRowsBySourceOrder()`), because `schema.rows` is appended in non-deterministic arrival order. `schema.rows` itself is never sorted — edit/delete/move logic always uses `rowIndex` + `spreadsheetId`.
- `updateRow(rowIndex, values, spreadsheetId)` and `deleteRow(rowIndex, spreadsheetId)` route by `spreadsheetId` — `openSpreadsheetSafely()` for remote (throwing a clean `Error` if inaccessible), `getActiveSpreadsheet()` for local. New rows go to the local `Database` sheet by default; in Master Mode a spreadsheet selector left of Save in the Add person view picks any loaded source, and `addRowWithData(values, spreadsheetId)` routes accordingly. Editing is never restricted by Master Mode.
- `masterSources: Array<{id, name}>` populates the Move destination dropdown (current spreadsheet is `id: null`; inaccessible sources show the raw ID) and grows as remote sources resolve.

## Actual personnel filter

`getActualPersonnelNames()` reads the spreadsheet link from `Handbook!A6` and range address from `Handbook!A7`, opens it via `openSpreadsheetSafely()`, and returns the flat list of non-empty names — or `null` if either cell is empty or the spreadsheet is inaccessible. `getSchemaAndData()` returns it as `actualPersonnelNames: string[]|null`.

- The **"Actual personnel"** checkbox is enabled only when the array is non-null and non-empty, and defaults to **checked** whenever enabled (both in `onDataLoaded()` and `resetListState()`).
- When checked, `applyFilters()` additionally requires `row.values[0]` (full name) to be in the list, via the shared `isActualPersonnelRow(row)`. Composes with column text filters and the regex toggle.
- `getSelectedRows()`/`getSelectedRowEntries()` also apply `isActualPersonnelRow()` when the filter is on, excluding checked rows not on the list even if selected under a different filter state. This is a deliberate exception to "selection persists across filter changes": the Actual Personnel filter is a correctness gate (who may appear in an export/move/delete), not a display convenience, so a stale off-list selection must never leak into an exported file.

## Move personnel (`movePersonnel`)

`movePersonnel(rowEntries, destinationSpreadsheetId)` is available only in Master Mode.

1. Groups `rowEntries` by source `spreadsheetId`, each group sorted in descending `rowIndex` order (so deleting lower rows doesn't shift higher ones).
2. Per row: reads the data, appends it to the destination `Database` sheet, then **hard-deletes** it from the source (`sheet.deleteRow()` — not the soft delete `deleteRow()` does). Returns the new `rowIndex` and `values` so the client can update `schema.rows` in place.
3. Rows where source === destination, or whose source spreadsheet/sheet is inaccessible, are skipped (logged, not moved/deleted) and returned in `skippedEntries: Array<{rowIndex, spreadsheetId}>`.
4. After each row move, it tries to move the person's Drive folder. The folder is located from the row's photo-column link (`resolvePersonFolderId()`, `DriveHelpers.js`): a folder link is the person's folder, a file link resolves to its single parent folder. The destination folder is the parent of the first resolvable person folder among the first `MOVE_DEST_FOLDER_PROBE_ROWS` (20) destination rows (`findUnitFolderIdFromPhotos()`), read before any row is appended. There is no name matching and no fallback: if the photo link is missing/inaccessible/ambiguous, or the destination has no photo to probe, the row move still completes and the log note says why the folder was not moved. A folder already in the destination is left alone.

Returns `{ log, movedRows, skippedEntries }`. The client (`runMove()`/`onMoveSuccess()`) stashes the sent entries as `pendingMoveEntries`, then subtracts `skippedEntries` to get what was *actually* hard-deleted. Two consequences:

- Skipped rows must stay in `schema.rows` — they're still in the source sheet.
- Every other remaining row sharing a deleted row's source `spreadsheetId` needs its cached `rowIndex` decremented by the number of actually-deleted rows above it, otherwise a later move/edit/delete addresses the wrong physical row (`sheet.deleteRow()` shifts everything below up). Any new client code that hard-deletes source rows must apply the same adjustment (see `onDeleteSuccess()` for the single-row equivalent).
