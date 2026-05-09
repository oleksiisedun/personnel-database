// Sheet names
const SHEET_DATABASE = 'Database';
const SHEET_HANDBOOK = 'Handbook';

// Named cell / range addresses
const EDIT_MODE_CELL = 'M1';

// Handbook layout — table-type definitions (rows 2–10, col A, all columns)
const HANDBOOK_TYPES_ROW_START = 2;
const HANDBOOK_TYPES_ROW_COUNT = 9;

// Handbook layout — placeholder correspondence table (B11:D50, first two rows are headers)
const HANDBOOK_CORR_ROW_START = 11;
const HANDBOOK_CORR_COL_START = 2;
const HANDBOOK_CORR_ROW_COUNT = 40;
const HANDBOOK_CORR_COL_COUNT = 3;

// Database column names
const COL_DRAFT_DATE = 'Дата призову';
const COL_SERVICE_HISTORY = 'Проходження служби';
const COL_CLOSE_RELATIVES = 'Близькі родичі';
const COL_MARITAL_STATUS = 'Сімейний стан';
const COL_CONTRACT_UNTIL = 'Контракт укладено до';

// Google Drive IDs
const F1_TEMPLATE_ID = '16yktSuOPgjNxQap-SCQZkmcISOftxWPZuSovVcAFkYI';
const EXPORT_FOLDER_ID = '1mG3vDgV9fCIYAt1S4Aj-IEUhKTw9O-Ki';

// Export / document settings
const EXPORT_DOC_PREFIX = 'Ф-1 ';
const DEFAULT_UNIT_NUMBER = '3102';
const IMAGE_MAX_HEIGHT = 600;

// Web editor UI settings
const COLUMN_MIN_WIDTHS = { text: 150, image: 150, table: 900 };
const WEB_EDITOR_WIDTH = 800;
const WEB_EDITOR_HEIGHT = 600;
