/**
 * Returns the Handbook sheet for the active spreadsheet, or null if not found.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getHandbookSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HANDBOOK);
}

/**
 * Reads the Handbook's Data Types & Allowed Values table (HANDBOOK_DATA_TYPES_RANGE) into a
 * map from data type name (lowercased) to its ordered list of allowed values. Only rows with
 * at least one non-empty value produce an entry — a type row left with no values (e.g. "text",
 * "date") is documentation only and correctly gets no dropdown. Used by getSchemaAndData() to
 * attach col.dropdownOptions to every column — top-level or *-table sub-column — whose type
 * matches a row here.
 * @returns {Object.<string, string[]>}
 */
function getDataTypeOptionsMap() {
  const sheet = getHandbookSheet();
  if (!sheet) return {};
  /** @type {Object.<string, string[]>} */
  const map = {};
  try {
    sheet.getRange(HANDBOOK_DATA_TYPES_RANGE).getValues().forEach(row => {
      const type = String(row[0]).trim().toLowerCase();
      if (!type) return;
      const values = row.slice(1).map(v => String(v).trim()).filter(v => v !== '');
      if (values.length) map[type] = values;
    });
  } catch (e) {
    // skip if Handbook data-types range is inaccessible
  }
  return map;
}

/**
 * Reads the Handbook's Table Columns table (HANDBOOK_TABLE_COLUMNS_RANGE) into a map from
 * *-table column type (lowercased) to its ordered sub-column definitions. Used by
 * getSchemaAndData() to attach col.tableHeaders to every *-table column.
 * @returns {Object.<string, Array<{name: string, type: string}>>}
 */
function getTableColumnsMap() {
  const sheet = getHandbookSheet();
  if (!sheet) return {};
  /** @type {Object.<string, {name: string, type: string}[]>} */
  const map = {};
  try {
    sheet.getRange(HANDBOOK_TABLE_COLUMNS_RANGE).getValues().forEach(row => {
      const tableType = String(row[0]).trim().toLowerCase();
      const name = String(row[1]).trim();
      const colType = String(row[2]).trim().toLowerCase();
      if (!tableType || !name) return;
      if (!map[tableType]) map[tableType] = [];
      map[tableType].push({ name, type: colType });
    });
  } catch (e) {
    // skip if Handbook table-columns range is inaccessible
  }
  return map;
}

/**
 * Validates that the active spreadsheet's Database and Handbook sheets exist.
 * Most Handbook reads in getSchemaAndData() individually fall back to a safe
 * empty default on error rather than throwing, so a missing Handbook sheet
 * would otherwise silently degrade the editor instead of surfacing an error.
 * Called by openWebEditor()/openPhotoExport() before the modal is shown.
 *
 * @returns {{ ok: boolean, errors: string[] }}
 */
function handbookCheck() {
  const errors = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_DATABASE)) errors.push(`Sheet "${SHEET_DATABASE}" not found.`);
  if (!ss.getSheetByName(SHEET_HANDBOOK)) errors.push(`Sheet "${SHEET_HANDBOOK}" not found.`);

  return { ok: errors.length === 0, errors };
}

/**
 * Resolves a spreadsheet by optional ID. Returns the remote spreadsheet when an
 * ID is supplied (via openSpreadsheetSafely), or the active spreadsheet when ID
 * is null / undefined / empty. Returns null only when a non-empty ID is provided
 * but the spreadsheet is inaccessible.
 *
 * @param {string|null|undefined} spreadsheetId
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet|null}
 */
function resolveSpreadsheet(spreadsheetId) {
  return spreadsheetId
    ? openSpreadsheetSafely(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Resolves the "Database" sheet for the given spreadsheet. Throws if the
 * spreadsheet is inaccessible or the sheet is missing.
 *
 * @param {string|null|undefined} spreadsheetId
 * @returns {{ ss: GoogleAppsScript.Spreadsheet.Spreadsheet, sheet: GoogleAppsScript.Spreadsheet.Sheet }}
 */
function getDatabaseSheet(spreadsheetId) {
  const ss = resolveSpreadsheet(spreadsheetId);
  if (!ss) throw new Error('Spreadsheet is not accessible.');
  const sheet = ss.getSheetByName(SHEET_DATABASE);
  if (!sheet) throw new Error(`Sheet "${SHEET_DATABASE}" not found.`);
  return { ss, sheet };
}

/**
 * Reads a Handbook cell holding a Drive ID (or sharing URL) and parses it via
 * parseDriveId(). Returns '' if the sheet is missing.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|null} handbookSheet
 * @param {string} cellAddress
 * @returns {string}
 */
function getDriveIdFromHandbook(handbookSheet, cellAddress) {
  return handbookSheet ? parseDriveId(String(handbookSheet.getRange(cellAddress).getValue()).trim()) : '';
}

/**
 * Reads the Master Mode toggle (MASTER_MODE_CELL) directly from a given
 * Handbook sheet, rather than resolving the active spreadsheet's own Handbook
 * via getHandbookSheet() the way getMasterMode() (Code.js) does. Shared by
 * getMasterMode() and getUnitDataFolder() (DriveHelpers.js), which already holds a
 * Handbook sheet reference for a spreadsheet that isn't necessarily the
 * active one (a Master Mode source, or the destination in movePersonnel()).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|null} handbookSheet
 * @returns {boolean}
 */
function readMasterModeFromSheet(handbookSheet) {
  if (!handbookSheet) return false;
  try {
    return handbookSheet.getRange(MASTER_MODE_CELL).getValue() === true;
  } catch (e) {
    return false;
  }
}

/**
 * Groups an array of row entries by spreadsheetId and sorts each group in
 * descending rowIndex order so that rows can be deleted sequentially without
 * the deletion of one row shifting the index of rows below it.
 *
 * @param {Array<{rowIndex: number, spreadsheetId: string|null}>} rowEntries
 * @returns {Map<string|null, Array<{rowIndex: number, spreadsheetId: string|null}>>}
 */
function groupAndSortBySpreadsheetId(rowEntries) {
  const groups = new Map();
  rowEntries.forEach(entry => {
    const key = entry.spreadsheetId ?? null;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(entry);
  });
  groups.forEach(entries => entries.sort((a, b) => b.rowIndex - a.rowIndex));
  return groups;
}

/**
 * Returns the Trash sheet for the given spreadsheet, creating it (with header rows
 * copied from dbSheet) if it does not exist yet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dbSheet - The Database sheet to copy headers from.
 * @param {number} numCols - Number of columns for the header copy range.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureTrashSheetExists(ss, dbSheet, numCols) {
  let trashSheet = ss.getSheetByName(SHEET_TRASH);
  if (!trashSheet) {
    trashSheet = ss.insertSheet(SHEET_TRASH);
    trashSheet.getRange(1, 1, 2, numCols)
      .setValues(dbSheet.getRange(1, 1, 2, numCols).getValues());
  }
  return trashSheet;
}

/**
 * Parses a pipe-and-newline encoded sub-table cell (the storage format used by
 * all *-table columns) into an array of field arrays.
 * Field separator: TABLE_FIELD_SEP  Row separator: TABLE_ROW_SEP
 * Trailing empty fields from a trailing separator are preserved but harmless.
 *
 * @param {string} rawValue - Raw encoded cell string.
 * @returns {string[][]} Array of rows, each row being an array of field strings.
 */
function _parseSubTable(rawValue) {
  if (!rawValue) return [];
  return rawValue.split(TABLE_ROW_SEP)
    .map(row => row.split(TABLE_FIELD_SEP))
    .filter(fields => fields[0] && fields[0].trim());
}

/**
 * Encodes an array of field arrays back into a sub-table cell string — the
 * inverse of _parseSubTable().
 *
 * @param {string[][]} rows - Array of rows, each row being an array of field strings.
 * @returns {string} Pipe-and-newline encoded cell string.
 */
function _encodeSubTable(rows) {
  return rows.map(fields => fields.join(TABLE_FIELD_SEP)).join(TABLE_ROW_SEP);
}
