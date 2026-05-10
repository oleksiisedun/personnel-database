// Sheet names
const SHEET_DATABASE = 'Database';
const SHEET_HANDBOOK = 'Handbook';

// Named cell / range addresses
const EDIT_MODE_CELL = 'I2';

// Handbook layout — table-type definitions
const HANDBOOK_TYPES_RANGE = 'A2:G10';

// Handbook layout — placeholder correspondence table
const HANDBOOK_CORR_RANGE = 'A12:C40';

// Handbook layout — allowed values for "unit" type columns
const HANDBOOK_UNIT_RANGE = 'D12:D40';

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
