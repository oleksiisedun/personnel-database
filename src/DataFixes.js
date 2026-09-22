/**
 * Rewrites every data cell of the Database sheet's column whose header matches
 * `columnPattern` through `normalize`, then alerts how many cells changed.
 * Shared by the menu commands fixPhoneNumbers() and fixFullNames().
 * @param {RegExp} columnPattern - Header pattern locating the column in row 1.
 * @param {(cell: string) => string} normalize - Returns the cell text unchanged
 *   when nothing needs fixing (so it isn't counted or rewritten), else the fixed text.
 * @param {string} noun - Singular label for the summary alert, e.g. "phone number".
 * @returns {void}
 */
function _normalizeDatabaseColumn(columnPattern, normalize, noun) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATABASE);
  if (!sheet) {
    ui.alert(`Sheet "${SHEET_DATABASE}" not found.`);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = findColumnIndex(headers, columnPattern);
  if (colIndex === -1) {
    ui.alert(`Column matching "${columnPattern.source}" not found in row 1.`);
    return;
  }

  const numRows = sheet.getLastRow() - 2;
  if (numRows <= 0) {
    ui.alert('No data rows to process.');
    return;
  }

  const range = sheet.getRange(3, colIndex + 1, numRows, 1);
  let fixedCount = 0;

  const result = range.getValues().map(row => {
    const original = String(row[0]);
    const normalized = normalize(original);
    if (normalized === original) return row;
    fixedCount++;
    return [normalized];
  });

  range.setValues(result);
  ui.alert(`Fixed ${fixedCount} ${noun}(s).`);
}

/**
 * Normalizes phone numbers in the Database sheet's "Номер телефону" column:
 * adds a leading zero to bare 9-digit numbers, and strips the "38" country
 * prefix from 12-digit numbers. Triggered only from the custom menu.
 * @returns {void}
 */
function fixPhoneNumbers() {
  _normalizeDatabaseColumn(COL_PHONE_NUMBER, cell => {
    const phone = cell.trim();
    const fixed = normalizePhoneNumber(phone);
    return fixed === phone ? cell : fixed;
  }, 'phone number');
}

/**
 * Normalizes full names in the Database sheet's "ПІБ" column: trims
 * surrounding whitespace, collapses internal whitespace runs (including
 * newlines) to a single space, and uppercases the surname (first word).
 * Triggered only from the custom menu.
 * @returns {void}
 */
function fixFullNames() {
  _normalizeDatabaseColumn(COL_FULL_NAME, normalizeFullName, 'full name');
}
