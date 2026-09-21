# Exports

F-1 / Wanted Card document export and XLSX export. Back to the [README](../README.md). The S-КАДР photo export is covered in [Features](features.md#photo-export-for-s-кадр).

## Document export

Two export types are available from the toolbar. Both operate on the **selected** rows (rows checked via the checkbox column). The export buttons are disabled until at least one row is selected.

| Button | Template cell | Output prefix |
|--------|---------------|---------------|
| Export F-1 | `Handbook!A9` (`EXPORT_F1_TEMPLATE_CELL`) | `Ф-1 ` |
| Export WC | `Handbook!A11` (`EXPORT_WC_TEMPLATE_CELL`) | `РК ` |

Exported files are saved to the Google Drive folder configured in `Handbook!A13` (`EXPORT_FOLDER_CELL`). Each cell accepts either a bare Drive ID or a full shareable link.

### How export works

Placeholders in the template use the format `{Column Name}`. Four passes run per document:

1. **Images** — `image`-type column placeholders are replaced with a compressed image blob
2. **Service history table** — `{Проходження служби}` is expanded into one table row per entry
3. **Direct text** — remaining `{Column Name}` placeholders are replaced with cell values
4. **Correspondence table** — Handbook-defined aliases and computed values fill any remaining placeholders

After all passes, the marital status line is underlined based on the `Сімейний стан` value.

### Image compression

Google Docs embeds an inserted image's actual bytes regardless of the display size set afterward, so full-resolution photos (often several MB) would bloat exported documents even though they display at only `IMAGE_MAX_HEIGHT` px tall. Image placeholders are therefore filled with a resized thumbnail fetched via the Advanced Drive Service (`Drive.Files.get(fileId, {fields: 'thumbnailLink'})`, enabled in `appsscript.json`) requested at `EXPORT_IMAGE_THUMBNAIL_SIZE` px wide, rather than the original file. If the Drive advanced service or thumbnail fetch fails for any reason, export falls back to the original full-resolution blob, so this never breaks or blanks an export.

### Computed values

These keys can be placed in column C of the correspondence table:

| Key | Description |
|-----|-------------|
| `totalServiceLength` | Duration from last date in `Дата призову` to today, e.g. `3 роки, 8 місяців, 17 днів (станом на 09.05.2026)`. Whole years and months are counted from the start date (a 31st clamps to the last day of a shorter month), then the leftover days; the parts are never negative |
| `contractSignDate` | First date in `Дата призову` + unit number from first service entry, e.g. `07.05.2015 з в/ч 3011` |
| `currentPosition` | Position title from the last entry in `Проходження служби` |
| `currentPositionStartDate` | Start date of the last entry in `Проходження служби` |
| `motherFullName` | Full name of relative with relation `мати` |
| `fatherFullName` | Full name of relative with relation `батько` |
| `spouseFullName` | Full name of relative with relation `дружина` or `чоловік` |
| `motherActualAddress` | Address of relative with relation `мати` |
| `fatherActualAddress` | Address of relative with relation `батько` |
| `spouseActualAddress` | Address of relative with relation `дружина` or `чоловік` |
| `motherPhoneNumber` | Phone of relative with relation `мати` |
| `fatherPhoneNumber` | Phone of relative with relation `батько` |
| `spousePhoneNumber` | Phone of relative with relation `дружина` or `чоловік` |
| `childrenNamesBirthDates` | Numbered list of children's names and birth dates |
| `childrenPhoneNumbers` | Comma-separated phone numbers of all children |
| `relativesWithPhoneNumbers` | Semicolon-separated list of all relatives with a phone number, formatted as `relation, name, address, phone` |
| `awardsList` | Semicolon-joined sentence built from the `Нагороди` sub-table, each entry formatted `{name} №{order number} від {order date}`, with the number/date parts individually omitted when empty |

### Large exports

When more than `EXPORT_CONFIRM_THRESHOLD` (default 10) rows are selected, clicking an export button shows a confirmation dialog with a time estimate before the export begins. This prevents accidental long-running exports, since there is no way to cancel once started. The estimate is computed as `total × EXPORT_SECONDS_PER_DOC` (default 6 s/doc, so 5 docs ≈ 30 s).

Exports run in automatic batches capped at `EXPORT_TIME_LIMIT_MS` (5 minutes) to stay within the Google Apps Script execution limit. The client automatically fires the next batch until all rows are done — no user interaction required. The progress bar in the export dialog shows real-time progress across batches.

## XLSX export

The "Export XLSX" toolbar button exports the **selected** rows (checkbox column), restricted to the columns currently **visible** in the list view (toggle via "Columns ▾"), as a single `.xlsx` file:

- Regular columns → cell value as-is
- `*-table` columns → the raw pipe/newline-encoded storage string, unchanged
- `image` columns, and any other column whose value looks like a Drive sharing URL → a clickable `HYPERLINK()` formula pointing at the file/folder's Drive view URL (no image embedding, no blob fetch — link only)

The file is named `Export DD.MM.YYYY.xlsx` (today's date) and saved to the same Drive folder as F-1/WC exports (`Handbook!A13`). Repeated exports on the same day are saved as separate files — Drive allows duplicate filenames, so no overwrite/suffix logic is applied.

Since Apps Script has no way to author `.xlsx` bytes directly and this project has no build step (so no bundling a library like ExcelJS), the export is built as a temporary Google Sheet and converted by fetching the Sheets export URL (`.../export?format=xlsx`) via `UrlFetchApp`, authorized with the script's own OAuth token (`Blob.getAs()` doesn't support this conversion); the temp sheet is always deleted afterward, even on error. Unlike F-1/WC export, this is a single `google.script.run` call with no batching — it produces one file for the whole selection, not one file per row, so there's no partial result to resume. See `exportXLSX()` in `ExportXlsx.js` and `runExportXlsx()` in `WebEditor.export.js.html`.
