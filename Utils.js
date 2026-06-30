/**
 * Returns the Handbook sheet for the active spreadsheet, or null if not found.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getHandbookSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HANDBOOK);
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
