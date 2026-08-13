let _sharedHandbookSheetCache; // undefined = not yet resolved this execution

/**
 * Opens the shared fallback Handbook spreadsheet (SHARED_HANDBOOK_SPREADSHEET_ID)
 * and returns its "Handbook" sheet, memoized for the lifetime of the current
 * execution since multiple call sites (e.g. movePersonnel's source + destination
 * lookups) may need it in the same request. Returns null if the constant is
 * empty, the spreadsheet is inaccessible, or it has no Handbook sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getSharedHandbookSheet() {
  if (_sharedHandbookSheetCache === undefined) {
    const ss = SHARED_HANDBOOK_SPREADSHEET_ID ? openSpreadsheetSafely(SHARED_HANDBOOK_SPREADSHEET_ID) : null;
    _sharedHandbookSheetCache = ss ? ss.getSheetByName(SHEET_HANDBOOK) : null;
  }
  return _sharedHandbookSheetCache;
}

/**
 * Resolves the Handbook sheet for a given spreadsheet: its own local
 * "Handbook" sheet if present, otherwise the shared fallback Handbook
 * spreadsheet's sheet (see SHARED_HANDBOOK_SPREADSHEET_ID in Config.js).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function resolveHandbookSheet(ss) {
  return ss.getSheetByName(SHEET_HANDBOOK) || getSharedHandbookSheet();
}

/**
 * Returns the Handbook sheet for the active spreadsheet — its own local
 * "Handbook" sheet if present, otherwise the shared fallback (see
 * resolveHandbookSheet()). Returns null if neither is available.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getHandbookSheet() {
  return resolveHandbookSheet(SpreadsheetApp.getActiveSpreadsheet());
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
  if (!sheet) throw new Error('Sheet "Database" not found.');
  return { ss, sheet };
}

/**
 * Converts every cell in a raw sheet row to a string, treating null and undefined
 * as empty string.
 *
 * @param {Array<*>} row - Raw row array as returned by Sheet.getValues().
 * @returns {string[]}
 */
function stringifyRowValues(row) {
  return row.map(c => String(c == null ? '' : c));
}

/**
 * Formats a Date as DD.MM.YYYY (zero-padded).
 *
 * @param {Date} date
 * @returns {string}
 */
function formatDateDDMMYYYY(date) {
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  return `${dd}.${mm}.${date.getFullYear()}`;
}

/**
 * Trims, collapses internal whitespace runs to a single space, and
 * uppercases the first word (surname) of a full name.
 * @param {string} value
 * @returns {string}
 */
function normalizeFullName(value) {
  const collapsed = String(value).trim().replace(WHITESPACE_RUN_REGEX, ' ');
  if (!collapsed) return collapsed;
  const [surname, ...rest] = collapsed.split(' ');
  return [surname.toUpperCase(), ...rest].join(' ');
}

/**
 * Trims or pads a values array to exactly numCols elements. Excess values are
 * dropped; missing values are filled with empty strings.
 *
 * @param {Array<*>} values - Source row values.
 * @param {number} numCols - Target column count.
 * @returns {Array<*>}
 */
function padRowToColumnCount(values, numCols) {
  const padded = values.slice(0, numCols);
  while (padded.length < numCols) padded.push('');
  return padded;
}

/**
 * Derives the column schema array from raw sheet data (all rows including headers).
 * Row 0 is column names, row 1 is column types.
 *
 * @param {Array<Array<*>>} allData - Full sheet data including header rows.
 * @returns {Array<{name: string, type: string}>}
 */
function extractColumnSchema(allData) {
  return allData[0].map((name, i) => ({
    name: String(name),
    type: String(allData[1]?.[i] ?? '').toLowerCase(),
  }));
}

/**
 * Converts an A1 column letter (e.g. "A", "F", "AA") to a 0-based column index.
 *
 * @param {string} letter - Column letter(s), case-insensitive.
 * @returns {number} 0-based column index.
 */
function columnLetterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.toUpperCase().charCodeAt(i) - 64);
  }
  return index - 1;
}

/**
 * Finds the index of the first header cell matching the given pattern.
 *
 * @param {Array<*>} headerRow - Raw header row values (e.g. sheet row 1).
 * @param {RegExp} pattern - Pattern to test each trimmed header against.
 * @returns {number} 0-based column index, or -1 if no header matches.
 */
function findColumnIndex(headerRow, pattern) {
  return headerRow.findIndex(h => pattern.test(String(h).trim()));
}

/**
 * Extracts the gid (tab id) from a Google Sheets URL's ?gid=... / #gid=... param.
 *
 * @param {string} url
 * @returns {number|null} The parsed gid, or null if the URL has none.
 */
function parseGidFromUrl(url) {
  const m = GID_REGEX.exec(String(url));
  return m ? Number(m[1]) : null;
}

/**
 * Finds the first key in a plain object whose text matches the given pattern.
 *
 * @param {Object.<string, *>} obj - Plain object keyed by column header text.
 * @param {RegExp} pattern - Pattern to test each key against.
 * @returns {string|null} The first matching key, or null if none match.
 */
function findKeyByPattern(obj, pattern) {
  return Object.keys(obj).find(key => pattern.test(key)) ?? null;
}

/**
 * Returns the value of the first entry in a row-data map whose key matches
 * the given column pattern.
 *
 * @param {Object.<string, string>} data - Row data map keyed by column header text.
 * @param {RegExp} pattern - Column-name pattern (one of the COL_* constants).
 * @returns {string} The matched cell value, or '' if no column matches.
 */
function getFieldByPattern(data, pattern) {
  const key = findKeyByPattern(data, pattern);
  return key !== null ? data[key] : '';
}

/**
 * Compares two column schemas and returns an entry for every position where
 * name or type differs. Positions where both sides have an empty name are
 * skipped — those are trailing blank columns from getDataRange() expanding
 * past the real schema extent, not genuine mismatches.
 *
 * @param {Array<{name: string, type: string}>} localSchema
 * @param {Array<{name: string, type: string}>} remoteSchema
 * @returns {Array<{colIndex: number, localName: string, localType: string, remoteName: string, remoteType: string}>}
 */
function compareColumnSchemas(localSchema, remoteSchema) {
  const len = Math.max(localSchema.length, remoteSchema.length);
  const mismatches = [];
  for (let i = 0; i < len; i++) {
    const local  = localSchema[i]  ?? { name: '', type: '' };
    const remote = remoteSchema[i] ?? { name: '', type: '' };
    if (!local.name && !remote.name) continue;
    if (local.name !== remote.name || local.type !== remote.type) {
      mismatches.push({ colIndex: i, localName: local.name, localType: local.type,
                        remoteName: remote.name, remoteType: remote.type });
    }
  }
  return mismatches;
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
