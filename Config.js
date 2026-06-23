// Sheet names
const SHEET_DATABASE = 'Database';
const SHEET_HANDBOOK = 'Handbook';
const SHEET_TRASH = 'Trash';

// Handbook layout — Master Mode toggle cell and source spreadsheet IDs range
const MASTER_MODE_CELL = 'M2';
const MASTER_MODE_SOURCES_RANGE = 'N2:N';

// Handbook layout — data folder cell
const DATA_FOLDER = 'M4';

// Handbook layout — table-type definitions
const HANDBOOK_TYPES_RANGE = 'A2:K15';

// Handbook layout — placeholder correspondence table
const HANDBOOK_CORR_RANGE = 'A17:C40';

// Handbook layout — allowed values for "unit" type columns
const HANDBOOK_UNIT_RANGE = 'D17:D40';

// Handbook layout — allowed values for "origin" type columns
const HANDBOOK_ORIGIN_RANGE = 'E17:E40';

// Handbook layout — allowed values for "marital-status" type columns
const HANDBOOK_MARITAL_STATUS_RANGE = 'F17:F40';

// Handbook layout — allowed values for "sex" type columns
const HANDBOOK_SEX_RANGE = 'G17:G40';

// Handbook layout — actual personnel list (external spreadsheet link and range address)
const ACTUAL_PERSONNEL_SPREADSHEET_CELL = 'M6';
const ACTUAL_PERSONNEL_RANGE_CELL = 'M7';

// Handbook layout — export template and destination folder IDs
const EXPORT_F1_TEMPLATE_CELL = 'M9';
const EXPORT_WC_TEMPLATE_CELL = 'M11';
const EXPORT_FOLDER_CELL = 'M13';

// Database column names
const COL_DRAFT_DATE = 'Дата призову';
const COL_SERVICE_HISTORY = 'Проходження служби';
const COL_CLOSE_RELATIVES = 'Близькі родичі';
const COL_MARITAL_STATUS = 'Сімейний стан';
const COL_CONTRACT_UNTIL = 'Контракт укладено до';

// Export / document settings
const EXPORT_TIME_LIMIT_MS = 5 * 60 * 1000; // 5 min — 1 min safety margin before GAS 6-min hard limit
const F1_DOC_PREFIX = 'Ф-1 ';
const WC_DOC_PREFIX = 'Розшукова картка ';
const DEFAULT_UNIT_NUMBER = '3102';
const IMAGE_MAX_HEIGHT = 500;

// Web editor UI settings
const COLUMN_MIN_WIDTHS = { text: 150, image: 150, table: 900 };
const COLUMN_MAX_WIDTHS = { image: 250 };
const FILTER_DEBOUNCE_MS = 500;

// Image fetch concurrency pool.
// BATCH: how many Drive files are resolved in a single google.script.run call.
//   Larger → fewer round trips, but bigger response payload per call.
// CONCURRENCY: how many batches run in parallel.
//   Raising this speeds up large lists, but too high risks Apps Script's
//   30-concurrent-execution quota and re-introduces dropped images.
const IMAGE_FETCH_BATCH_SIZE = 10;
const IMAGE_FETCH_CONCURRENCY = 3;

// Master Mode source fetch concurrency — how many remote spreadsheets are
// opened in parallel while streaming in their rows after the initial
// (local-only) row set has already rendered.
const MASTER_MODE_FETCH_CONCURRENCY = 3;

// Persistent IndexedDB image cache — how long an entry is kept before re-fetching from Drive.
const IMAGE_CACHE_TTL_DAYS = 7;
