// Sheet names
const SHEET_DATABASE = 'Database';
const SHEET_HANDBOOK = 'Handbook';
const SHEET_TRASH = 'Trash';

// Named cell / range addresses
const EDIT_MODE_CELL = 'I2';

// Handbook layout — table-type definitions
const HANDBOOK_TYPES_RANGE = 'A2:G10';

// Handbook layout — placeholder correspondence table
const HANDBOOK_CORR_RANGE = 'A12:C40';

// Handbook layout — allowed values for "unit" type columns
const HANDBOOK_UNIT_RANGE = 'D12:D40';

// Handbook layout — allowed values for "origin" type columns
const HANDBOOK_ORIGIN_RANGE = 'E12:E40';

// Handbook layout — allowed values for "marital-status" type columns
const HANDBOOK_MARITAL_STATUS_RANGE = 'F12:F40';

// Database column names
const COL_DRAFT_DATE = 'Дата призову';
const COL_SERVICE_HISTORY = 'Проходження служби';
const COL_CLOSE_RELATIVES = 'Близькі родичі';
const COL_MARITAL_STATUS = 'Сімейний стан';
const COL_CONTRACT_UNTIL = 'Контракт укладено до';

// Google Drive IDs
const F1_TEMPLATE_ID = '16yktSuOPgjNxQap-SCQZkmcISOftxWPZuSovVcAFkYI';
const WC_TEMPLATE_ID = '10WKfn_cUP-T_C_TWr3pPeR3sELLVgUFFDan8VFpCoPo';
const EXPORT_FOLDER_ID = '1mG3vDgV9fCIYAt1S4Aj-IEUhKTw9O-Ki';

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
