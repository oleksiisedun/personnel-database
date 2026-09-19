# Spreadsheet setup

How the `Database` and `Handbook` sheets are laid out, what each column type does, and the sample files for a new deployment. Back to the [README](../README.md).

## Spreadsheet structure

### `Database` sheet

| Row | Purpose |
|-----|---------|
| 1 | Column names |
| 2 | Column types (`text`, `image`, `date`, `number`, `unit`, `tin`, `relatives-table`, `service-table`, …) |
| 3+ | Data rows |

### `Trash` sheet

Same structure as `Database` (row 1 = column names, row 2 = column types, row 3+ = data rows). Deleted records are appended here instead of being permanently removed. The sheet is created automatically on first delete if it does not exist.

### `Handbook` sheet

Single-value config lives in column A as a vertical list (a documentation label, then its value on the row(s) below):

| Cell | Purpose |
|------|---------|
| `A2` (`MASTER_MODE_CELL`) | Master Mode toggle (checkbox) — when checked, the webview aggregates data from all source spreadsheets listed in `B2:B` |
| `A4` (`DATA_FOLDER`) | Google Drive folder ID of the shared "UNITS" parent folder (same value everywhere, imported like most other Handbook config) — each unit's own person images/PDFs subfolder inside it is resolved per spreadsheet by matching the unit name extracted from the spreadsheet's own title (e.g. `"УСТАНОВЧІ ДАНІ О/С - 7 РОП"` → `"7 РОП"`), *unless* Master Mode is on for that spreadsheet, in which case `DATA_FOLDER` is used directly as that spreadsheet's own dedicated folder |
| `A6` (`ACTUAL_PERSONNEL_SPREADSHEET_CELL`) | Link or bare ID of the spreadsheet that holds the authoritative personnel list |
| `A7` (`ACTUAL_PERSONNEL_RANGE_CELL`) | Range address within that spreadsheet (e.g. `Sheet1!A:A`) containing full names |
| `A9` (`EXPORT_F1_TEMPLATE_CELL`) | Google Drive file ID (or shareable link) of the F-1 Docs template |
| `A11` (`EXPORT_WC_TEMPLATE_CELL`) | Google Drive file ID (or shareable link) of the Wanted Card Docs template |
| `A13` (`EXPORT_FOLDER_CELL`) | Google Drive folder ID (or shareable link) where exported documents are saved |
| `B2:B` (`MASTER_MODE_SOURCES_RANGE`) | Source spreadsheet IDs/URLs — one per row; used when Master Mode is ON |

Three more tables sit side by side across the rest of row 1 onward — each a header row followed by open-ended data rows (no fixed row cap; just add a row):

| Range | Purpose |
|-------|---------|
| `D2:F` (`HANDBOOK_CORR_RANGE`) | Placeholder correspondence table for document exports (see [Correspondence table columns](#correspondence-table-columns)) |
| `H2:J` (`HANDBOOK_TABLE_COLUMNS_RANGE`) | `*-table` sub-column definitions: `Table Type \| Column Name \| Column Type`, one row per sub-column |
| `L2:AL` (`HANDBOOK_DATA_TYPES_RANGE`) | Data type definitions: `Data Type \| Allowed Values...`, one row per type name. A row with values (e.g. `unit`, `blood-type`) is a dropdown type; a row with no values (e.g. `text`, `date`) is documentation only |

Adding a new dropdown type, or a new `*-table` type/column, is a Handbook edit — add a row to the relevant table above. No code change needed. See [Dropdown and table column types](#dropdown-and-table-column-types) below.

#### Correspondence table columns

| Column | Purpose |
|--------|---------|
| D | Placeholder name (used as `{placeholder}` in the template) |
| E | Source Database column name (value is copied directly) |
| F | Computed value key (see [Computed values](exports.md#computed-values) below) |

Exactly one of E or F should be filled per row.

## Column types

| Type | List view | Edit view |
|------|-----------|-----------|
| `text` | Plain text | Text input |
| `image` | Thumbnail (click to enlarge) | Google Drive link input with live preview |
| `date` | Plain text | Text input with `DD.MM.YYYY` format validation |
| `tin` | Plain text | Text input validated as exactly 10 digits |
| `number` | Plain text | Text input validated as digits only |
| *(any dropdown type, e.g.* `unit`*,* `origin`*,* `marital-status`*,* `sex`*,* `blood-type`*)* | Plain text | Dropdown of allowed values from the Handbook's Data Types table |
| `*-table` | Decoded mini-table | Row/column editor with add & delete — a sub-column can itself be a dropdown type |

### `*-table` encoding format

Table data is stored in a single cell as a pipe-and-newline delimited string:

```
value1 | value2 | value3
value1 | value2 | value3
```

### Image columns

Images are stored as Google Drive sharing links. Supported URL formats:

- `https://drive.google.com/file/d/FILE_ID/view`
- `https://drive.google.com/drive/folders/FOLDER_ID`
- `https://drive.google.com/open?id=FILE_ID`
- Bare file ID

Images are fetched server-side (via `DriveApp`) and returned as base64 data URLs, so all users with access to the spreadsheet can view images regardless of their personal Drive session.

Fetched images are persisted in an **IndexedDB** database (`pdb_images`) so subsequent dialog opens display all thumbnails immediately without any server round-trips. Entries expire after `IMAGE_CACHE_TTL_DAYS` (default 7 days). To force a full re-fetch, clear the site data for the script origin in browser DevTools.

If the linked file is a **PDF**, the cell shows a red "PDF" badge instead of a thumbnail; clicking it opens the file in Drive in a new tab. If the link points to a **Drive folder**, a blue "Folder" badge is shown instead. If the script owner does not have access to the linked file, a gray **"No access"** badge is shown.

### Date columns

Values are stored as plain text in `DD.MM.YYYY` format. In the edit view, the input validates the format on every keystroke and highlights the field in red with an error hint if the format is wrong.

### TIN columns

Values are stored as plain text. In the edit view, the input validates on every keystroke that the value is exactly 10 digits (digits only, no spaces or other characters).

### Save-time validation errors

Clicking **Save** re-validates every `date`/`tin`/`number` field (plain columns and `*-table` sub-columns alike) across **all** edit-view tabs, not just the one currently visible — including fields the user never touched, such as a malformed value entered directly in the sheet before the record was opened in the editor (empty values always pass). If any field fails, the editor switches to the tab containing the first invalid field and shows an alert listing every invalid field together with its tab and the expected format, instead of a generic banner — this is deliberate so a stale bad value on a tab the user isn't looking at (e.g. a birth date typed wrong directly in the sheet) isn't mistaken for a problem with whatever field the user was just editing. No save request is sent while any field is invalid.

Values entered directly in the sheet that do not match a validated column's expected format are displayed as-is in the list view; no validation is applied there — only the edit view enforces format, at save time as described above.

### Dropdown and table column types

Dropdown types are entirely defined in the Handbook's Data Types table (`L2:AL`, `Data Type | Allowed Values...`) — no code change is needed to add one. A type name with at least one value in its row (e.g. `unit`, `origin`, `marital-status`, `sex`, `blood-type`) automatically becomes a dropdown: at page load its values are read and served to the client as part of the schema, and in the edit view the field renders as a `<select>` containing only those values. Values entered directly in the sheet that are not in the allowed list are appended to the dropdown as an extra option and shown selected, so no data is lost. A type name with no values in its row (e.g. `text`, `date`, `image`, `tin`, `number`) is documentation only.

`*-table` sub-columns can be dropdown types too. The Table Columns table (`H2:J`, `Table Type | Column Name | Column Type`) assigns a type to every sub-column of every `*-table` — if that type has values in the Data Types table, the sub-column renders as a dropdown in the table editor; otherwise it's a plain text input. This is how, for example, a `medical-table`'s "Група крові" (blood type) column can be a dropdown while its other sub-columns stay free text.

## Sample files

The `samples/` directory contains example files for setting up a new deployment:

| File | Purpose |
|------|---------|
| `УСТАНОВЧІ ДАНІ ОС.xlsx` | Sample `Database` sheet layout — column headers matching the schema described in [Spreadsheet structure](#spreadsheet-structure), usable as a starting point for a new spreadsheet |
| `ДОВІДКА (Ф-1).docx` | Sample F-1 export template (see [Document export](exports.md#document-export)) with `{Column Name}` placeholders matching this schema |
| `РОЗШУКОВА КАРТКА.docx` | Sample Wanted Card export template (see [Document export](exports.md#document-export)) with `{Column Name}` placeholders matching this schema |

To use a template, upload it to Google Drive (converting to Google Docs format if needed) and set its file ID or link in `Handbook!A9` (`EXPORT_F1_TEMPLATE_CELL`) for the F-1 template, or `Handbook!A11` (`EXPORT_WC_TEMPLATE_CELL`) for the Wanted Card template.
