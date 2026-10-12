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

  const result = range.getValues().map((row) => {
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
  _normalizeDatabaseColumn(
    COL_PHONE_NUMBER,
    (cell) => {
      const phone = cell.trim();
      const fixed = normalizePhoneNumber(phone);
      return fixed === phone ? cell : fixed;
    },
    'phone number',
  );
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

/**
 * Fills empty cells of the Database sheet's sex column (the one typed `sex` in
 * row 2) from the patronymic in the "ПІБ" column. Cells already set, and rows
 * whose patronymic is unrecognizable, are left untouched. Triggered only from
 * the custom menu.
 * @returns {void}
 */
function addSex() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATABASE);
  if (!sheet) {
    ui.alert(`Sheet "${SHEET_DATABASE}" not found.`);
    return;
  }

  const [headers, types] = sheet.getRange(1, 1, 2, sheet.getLastColumn()).getValues();
  const sexCol = types.findIndex((t) => String(t).trim().toLowerCase() === COL_SEX_TYPE);
  if (sexCol === -1) {
    ui.alert(`No column of type "${COL_SEX_TYPE}" found in row 2.`);
    return;
  }
  const nameCol = findColumnIndex(headers, COL_FULL_NAME);
  if (nameCol === -1) {
    ui.alert(`Column matching "${COL_FULL_NAME.source}" not found in row 1.`);
    return;
  }

  const numRows = sheet.getLastRow() - 2;
  if (numRows <= 0) {
    ui.alert('No data rows to process.');
    return;
  }

  const names = sheet.getRange(3, nameCol + 1, numRows, 1).getValues();
  const sexRange = sheet.getRange(3, sexCol + 1, numRows, 1);
  let addedCount = 0;

  const result = sexRange.getValues().map((row, i) => {
    if (String(row[0]).trim()) return row;
    const sex = inferSexFromFullName(names[i][0]);
    if (!sex) return row;
    addedCount++;
    return [sex];
  });

  sexRange.setValues(result);
  ui.alert(`Added sex to ${addedCount} row(s).`);
}
