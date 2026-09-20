# Configuration (`Config.js`)

All tuneable constants live in `src/Config.js`. Back to the [README](../README.md).

| Constant | Default | Purpose |
|----------|---------|---------|
| `SHEET_DATABASE` | `'Database'` | Name of the data sheet |
| `SHEET_HANDBOOK` | `'Handbook'` | Name of the handbook sheet |
| `SHEET_TRASH` | `'Trash'` | Name of the trash sheet (created automatically on first delete) |
| `MASTER_MODE_CELL` | `'A2'` | Cell that holds the Master Mode checkbox |
| `DATA_FOLDER` | `'A4'` | Cell that holds the Google Drive folder ID of the shared "UNITS" parent folder |
| `MASTER_MODE_SOURCES_RANGE` | `'B2:B'` | Range of source spreadsheet IDs/URLs for Master Mode |
| `ACTUAL_PERSONNEL_SPREADSHEET_CELL` | `'A6'` | Cell holding the link/ID of the external spreadsheet with the actual personnel list |
| `ACTUAL_PERSONNEL_RANGE_CELL` | `'A7'` | Cell holding the range address within that spreadsheet (e.g. `Sheet1!A:A`) |
| `EXPORT_F1_TEMPLATE_CELL` | `'A9'` | Cell holding the Drive ID or link of the F-1 Docs template |
| `EXPORT_WC_TEMPLATE_CELL` | `'A11'` | Cell holding the Drive ID or link of the Wanted Card Docs template |
| `EXPORT_FOLDER_CELL` | `'A13'` | Cell holding the Drive ID or link of the folder for exported documents |
| `HANDBOOK_CORR_RANGE` | `'D2:F'` | Range of the placeholder correspondence table |
| `HANDBOOK_TABLE_COLUMNS_RANGE` | `'H2:J'` | Range of `*-table` sub-column definitions (`Table Type \| Column Name \| Column Type`) |
| `HANDBOOK_DATA_TYPES_RANGE` | `'L2:AL'` | Range of data type definitions (`Data Type \| Allowed Values...`) — a type with values is a dropdown type |
| `COL_DRAFT_DATE` / `_SERVICE_HISTORY` / `_CLOSE_RELATIVES` / `_AWARDS` / `_MARITAL_STATUS` / `_CONTRACT_UNTIL` / `_PHONE_NUMBER` / `_FULL_NAME` / `_PHOTO` / `_CARD_ID` | case-insensitive regexes, e.g. `/дата призову/i` | Match `Database` column headers by name (not exact string) so several features keep working even if header casing drifts across Master Mode sources. The `Database` sheet's headers must contain text matching each of these for the corresponding feature to work: `Дата призову` (service length), `Проходження служби` (service table), `Близькі родичі` (relative lookups), `Нагороди` (award import/export), `Сімейний стан` (export underline), `Контракт укладено до`, `Номер телефону` (Fix phone numbers), `ПІБ` (Fix full names), `Фото`/`S-КАДР ID` (photo export) — the sample `Database` layout in `samples/` already uses matching headers |
| `DOCUMENT_PHOTO_TAB_NAME` | `'Фото документи'` | Tab label grouping all `image`-type fields together in the edit view (see [Web editor features](features.md#web-editor-features)) |
| `EXPORT_TIME_LIMIT_MS` | `300000` | Max server execution time per batch (ms) |
| `EXPORT_CONFIRM_THRESHOLD` | `10` | Row count above which a confirmation dialog is shown before export starts |
| `EXPORT_SECONDS_PER_DOC` | `6` | Seconds per document used to estimate export duration in the confirmation dialog |
| `F1_DOC_PREFIX` | `'Ф-1 '` | Filename prefix for F-1 exports |
| `WC_DOC_PREFIX` | `'РК '` | Filename prefix for Wanted Card exports |
| `DEFAULT_UNIT_NUMBER` | `'3102'` | Fallback military unit number for `contractSignDate` |
| `IMAGE_MAX_HEIGHT` | `500` | Max image height (px) when inserting into a document |
| `EXPORT_IMAGE_THUMBNAIL_SIZE` | `800` | Width (px) requested from Drive's thumbnail service for export images, before the `IMAGE_MAX_HEIGHT` display clamp is applied |
| `PHOTO_EXPORT_FOLDER_PREFIX` | `'Photos for S-КАДР '` | Destination folder name prefix for [Photo export for S-КАДР](features.md#photo-export-for-s-кадр), combined with today's date |
| `GID_REGEX` | `/[?&#]gid=(\d+)/` | Extracts the tab id from a Google Sheets URL's `gid` parameter; used by [Award import](features.md#award-import-from-s-кадр) to pick the correct tab |
| `UNIT_NAME_SEPARATOR` | `' - '` | Separator between a unit spreadsheet's fixed title prefix and its actual unit name (e.g. `"УСТАНОВЧІ ДАНІ О/С - 7 РОП"` → `"7 РОП"`); used by `extractUnitName()` to match against Drive folder names in `getUnitDataFolder()` |
| `AWARDS_IMPORT_ID_COL` / `_NAME_COL` / `_ORDER_NUMBER_COL` / `_ORDER_DATE_COL` | `'A'` / `'F'` / `'G'` / `'H'` | Fixed column letters (A1 notation) for the ID/name/order-number/order-date fields in the external award import sheet |
| `AWARDS_IMPORT_DATA_START_ROW` | `2` | First data row (after the header) in the external award import sheet |
| `AWARDS_ORDER_NUMBER_CLEAN_REGEX` | `/\/\d+\|[\W]+/g` | Strips surrounding text/punctuation from the import sheet's free-text order-number field, keeping just the leading number |
| `AWARDS_IMPORT_NOT_FOUND_DISPLAY_LIMIT` | `20` | Max "not found" IDs listed by name in the award import summary alert before collapsing the rest into a `(+N more)` suffix |
| `TABLE_FIELD_SEP` | `' \| '` | Field separator used to encode/decode `*-table` cell values (shared by `Export.js` and `Import.js`; the client declares its own copy) |
| `TABLE_ROW_SEP` | `'\n'` | Row separator used to encode/decode `*-table` cell values (same sharing as `TABLE_FIELD_SEP`) |
| `COLUMN_MIN_WIDTHS` | `{ text: 150, image: 150, table: 900 }` | Minimum column widths (px) in the list view |
| `COLUMN_MAX_WIDTHS` | `{ image: 250 }` | Maximum column widths (px) in the list view |
| `FILTER_DEBOUNCE_MS` | `500` | Debounce delay (ms) for filter text inputs |
| `IMAGE_FETCH_BATCH_SIZE` | `10` | Number of Drive files resolved per `google.script.run` call |
| `IMAGE_FETCH_CONCURRENCY` | `3` | Number of image-fetch batches running in parallel; raising it speeds up large lists but risks the Apps Script 30-concurrent-execution limit |
| `MASTER_MODE_FETCH_CONCURRENCY` | `3` | Number of remote Master Mode source spreadsheets fetched in parallel after the initial local-only row set has rendered |
| `IMAGE_CACHE_TTL_DAYS` | `7` | How many days a cached image entry survives in IndexedDB before being re-fetched |
| `DRIVE_URL_REGEX` | `/(?:\/folders\/\|\/d\/\|[?&]id=)([-\w]+)/` | Extracts a Drive file/folder ID from a sharing URL; shared by `parseDriveId()` and `looksLikeDriveUrl()` |
| `XLSX_EXPORT_FILENAME_PREFIX` | `'Export '` | Filename prefix for XLSX exports, e.g. `Export 11.08.2026.xlsx` |
| `XLSX_EXPORT_SECONDS_PER_ROW` | `0.2` | Seconds per row used to estimate XLSX export duration in the confirmation dialog |
