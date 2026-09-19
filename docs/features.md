# Features

What the web editor, the custom menu and the Master Mode features do. Back to the [README](../README.md); see also [Exports](exports.md).

## Master Mode

When the **Master Mode** checkbox (`Handbook!A2`) is checked, `getSchemaAndData()` returns the source spreadsheet IDs listed in `Handbook!B2:B` (`masterSourceIds`) without opening any of them itself, so the local `Database` rows render immediately. The client then calls `queueMasterSourceFetch()`, which fetches each source's rows one at a time through a concurrency-limited worker pool (`MASTER_MODE_FETCH_CONCURRENCY` in `Config.js`) via `getMasterSourceRows(spreadsheetId)`. Each remote row carries a `spreadsheetId` property so saves and deletes are routed back to the correct spreadsheet. As each source resolves, its rows are appended to the list and the view is silently re-filtered — so remote rows stream in over the following seconds while the local-only list is already visible. Inaccessible sources are skipped silently.

Before accepting a remote source's rows, `getMasterSourceRows()` compares the remote sheet's column names and types (rows 1 and 2) against the local `Database` sheet. If they differ, the source's rows are **not loaded** and an amber **⚠** warning button appears to the left of the "All Units" checkbox in the toolbar. Clicking it opens a dialog listing each mismatched source by name and each differing column by its spreadsheet letter (A, B, C…) alongside the local and remote header names. Sources with a matching schema load normally.

The displayed list is always sorted: local spreadsheet rows first, then each remote source in the order it appears in `Handbook!B2:B`, with rows within each source in their original sheet row order. This sort runs on every filter pass, so the order is stable even while sources are still loading.

The **"All units"** toolbar checkbox (visible only in Master Mode, checked by default) hides all streamed-in remote rows so only the local `Database` sheet's own rows are shown.

The green **Refresh** button (visible only in Master Mode, to the left of "All units") re-runs the same load — local rows plus every Master Mode source — without closing the dialog, so updates made by other admins in a source spreadsheet since the dialog was opened become visible. It preserves current filter text and column visibility, but clears the row selection checkboxes (row positions may have shifted upstream). Clicking it again while a previous refresh's remote sources are still streaming in safely discards the stale results instead of duplicating rows.

When Master Mode is **OFF**, only the local `Database` sheet is shown. Editing (add / edit / delete) is always available regardless of Master Mode.

New rows added via **Add person** always go to the local `Database` sheet, never to a remote source.

### Move personnel (Master Mode only)

The **Move** toolbar button appears only when Master Mode is on. Select one or more rows with the checkbox column, click **Move**, pick a destination spreadsheet from the dropdown, and confirm.

What happens on the server:
1. The row is appended to the destination `Database` sheet and **hard-deleted** from the source (not sent to Trash).
2. The person's Drive folder — searched by full name (first column) inside the source spreadsheet's own unit folder (resolved as the subfolder of `DATA_FOLDER` matching the unit name extracted from that spreadsheet's title, or `DATA_FOLDER` itself directly if Master Mode is on for that spreadsheet) — is moved to the destination spreadsheet's own unit folder using `DriveApp`. If either unit folder can't be resolved, or the person's folder is not found inside it, the row move still completes; the result dialog shows a per-person note.
3. Rows that already belong to the destination spreadsheet are skipped.

After a successful move the rows reappear in the list immediately under their new spreadsheet, without reopening the webview.

## Custom menu

Opening the web editor (either "Open Web Editor" or "Export photos for S-КАДР" below) first checks that the `Database` and `Handbook` sheets exist. If either is missing, a dialog reports it instead of the editor opening in a silently broken state.

The Sheets **More... ⭐️** menu (added by `onOpen()`) always has three items, plus two more shown only when Master Mode (`Handbook!A2`) is on:

| Item | Action |
|------|--------|
| Open Web Editor | Opens the web editor dialog described below |
| Fix phone numbers | Scans the `Database` sheet's `Номер телефону` column and rewrites two malformed shapes in place: bare 9-digit numbers missing the leading `0`, and 12-digit numbers carrying a `38` country-code prefix. Numbers already in canonical 10-digit form are left untouched. Runs synchronously over the whole sheet and reports the fixed count via a dialog. No undo beyond manual edit or `Trash` recovery. |
| Fix full names | Scans the `Database` sheet's `ПІБ` column and rewrites each value: trims surrounding whitespace, collapses internal whitespace runs (including newlines) to a single space, and uppercases the surname (first word). Already-normalized values are left untouched. Runs synchronously over the whole sheet and reports the fixed count via a dialog. No undo beyond manual edit or `Trash` recovery. |
| Export photos for S-КАДР *(Master Mode only)* | Opens the same web editor dialog in a dedicated progress view and copies photos for every Actual Personnel row across the local sheet and all Master Mode sources — see [Photo export for S-КАДР](#photo-export-for-s-кадр) |
| Import awards from S-КАДР *(Master Mode only)* | Prompts for an external Google Sheets URL and merges its award rows into the local `Нагороди` column across the local sheet and all Master Mode sources — see [Award import from S-КАДР](#award-import-from-s-кадр) |

## Photo export for S-КАДР

Triggered by the **Export photos for S-КАДР** menu item (Master Mode only). It opens the same `WebEditor` dialog as **Open Web Editor**, but in a dedicated progress view instead of the normal list — there is no manual row selection here.

Instead of operating on a checkbox selection, it sweeps **every** row on the Actual Personnel list (`Handbook!A6`/`A7`) across the local `Database` sheet and every Master Mode source, in two phases:

1. **Discovery** — walks the local sheet, then each `Handbook!B2:B` source in order, reading only the `Фото` and `S-КАДР ID` columns (no image bytes fetched yet). A person missing either value is skipped with a reason. If the same `S-КАДР ID` appears more than once (e.g. duplicated across sources), the first occurrence found wins and later ones are recorded as duplicate skips.
2. **Copy/convert** — for each eligible person, fetches the full-resolution photo from Drive, force-converts it to JPEG, and saves it as `{S-КАДР ID}.jpg` in a new destination folder named `Photos for S-КАДР DD.MM.YYYY` inside `Handbook!A13` (`EXPORT_FOLDER_CELL`). Runs in automatic time-boxed batches the same way F-1/WC export does, so large personnel lists don't hit the Apps Script execution limit.

When finished, the dialog shows a link to the destination folder plus the accumulated list of skipped/duplicate entries with their reasons. Closing the dialog in this mode closes the whole window, since there's no list view to return to.

## Award import from S-КАДР

Triggered by the **Import awards from S-КАДР** menu item (Master Mode only). Unlike every other Master Mode feature, this one uses plain Sheets `ui.prompt()`/`ui.alert()` dialogs instead of the web editor — the whole operation is a synchronous sheet-to-sheet scan with no per-row Drive I/O, so it comfortably finishes within one execution.

1. Prompts for the full URL of the external S-КАДР Google Sheet. If the URL contains a `gid` parameter it opens that exact tab; otherwise it uses the first tab.
2. Reads award rows by **fixed column position** (not header text) — ID, award name, order number, and order date columns, configured in `Config.js` (`AWARDS_IMPORT_ID_COL`/`_NAME_COL`/`_ORDER_NUMBER_COL`/`_ORDER_DATE_COL`, data starting at row `AWARDS_IMPORT_DATA_START_ROW`). Rows with no ID or no award name are skipped. The award name's first letter is capitalized, and the order number is cleaned of surrounding text (e.g. `"Указ Президента України № 559/2022"` → `"559"`).
3. Matches each imported ID against the local `Database` sheet and every Master Mode source's `S-КАДР ID` column. An ID found in more than one source resolves to the first match (local first, then sources in `Handbook!B2:B` order).
4. Merges the new award entries into the matched person's `Нагороди` cell, skipping any entry that exactly duplicates one already there, and writes the merged cell back.

The final alert shows counts (people updated, entries added, duplicates skipped, IDs not found) plus up to `AWARDS_IMPORT_NOT_FOUND_DISPLAY_LIMIT` not-found IDs by name. The full per-ID log is written to the Apps Script execution log (**Executions** in the Apps Script editor), since the alert dialog isn't scrollable.

## Web editor features

- **List view** — full-screen table with all columns and data
- **Filtering** — debounced live filter input above every column; supports plain text and regular expressions (toggle per session); for `image` columns the search matches the raw Drive URL/ID (`""` to filter empty); for `*-table` columns the search runs against the raw encoded cell content, so any sub-field value is matched
- **Actual personnel filter** — "Actual personnel" checkbox in the toolbar (enabled only when `Handbook!A6`/`A7` are configured and accessible), **checked by default** whenever it's enabled; when checked, only rows whose first-column value (full name) appears in the external personnel list are shown; composes with all other filters
- **All units filter** (Master Mode only) — "All units" checkbox in the toolbar, checked by default; uncheck to hide rows streamed in from Master Mode source spreadsheets and show only the local `Database` sheet's rows; composes with all other filters. An amber ⚠ button appears to its left when one or more remote sources have a column schema mismatch; clicking it shows which columns differ so the issue can be fixed in the remote sheet
- **Refresh** (Master Mode only) — green button to the left of "All units"; re-fetches local rows and every Master Mode source without closing the dialog, so changes made by other admins become visible; preserves filter text and column visibility, clears row selection
- **Reset** — orange button to the left of Refresh, always visible; clears all column filters, the regex/"Actual personnel"/"All units" toggles, hidden-column visibility, and the row selection checkboxes back to their defaults, without re-fetching data from the server
- **Add person** — appends a new empty row to the local `Database` sheet and opens it in the edit view immediately
- **Delete** — red "Delete" button in the edit view moves the record to the `Trash` sheet of its source spreadsheet (not available for unsaved new rows)
- **Unsaved changes confirmation** — clicking "Back" in the edit view with pending edits (including `*-table` fields) shows a "Keep Editing" / "Discard Changes" prompt instead of silently discarding them; no prompt if nothing changed
- **Column visibility** — "Columns ▾" button to hide/show individual columns; first column is always visible
- **Image thumbnails** — loaded asynchronously; persisted in IndexedDB so subsequent opens display instantly
- **Lightbox** — click any thumbnail to view the full image
- **Row selection** — checkbox column at the left of the table; master checkbox in the filter row selects/deselects all visible rows; indeterminate state when a subset is selected; drives which rows are exported and moved
- **Export XLSX** — exports the selected rows, restricted to currently visible columns, as a single `.xlsx` file (see [XLSX export](exports.md#xlsx-export))
- **Move** (Master Mode only) — moves selected rows to another spreadsheet and relocates the person's Drive folder; button is hidden when Master Mode is off
- **Edit view** — click a name in the first column to open a per-record editor. All regular (non-`*-table`, non-`image`) fields share one "main" tab; every `image`-type field is grouped into its own dedicated tab (labeled from `Config.js`'s `DOCUMENT_PHOTO_TAB_NAME`, shown only if the record has at least one `image` column); each `*-table` column (e.g. service history, close relatives, awards) also gets its own tab in the header next to the person's name, so wide sub-tables and photo fields don't crowd the main form. Switching tabs never discards edits — every tab stays mounted in the background for the life of the edit session
