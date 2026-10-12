/**
 * Converts every cell in a raw sheet row to a string, treating null and undefined
 * as empty string.
 *
 * @param {Array<*>} row - Raw row array as returned by Sheet.getValues().
 * @returns {string[]}
 */
function stringifyRowValues(row) {
  return row.map((c) => String(c == null ? '' : c));
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
 * Infers sex from the patronymic (third word) of a full name.
 * @param {string} fullName
 * @returns {string} SEX_MALE, SEX_FEMALE, or '' when there is no recognizable patronymic.
 */
function inferSexFromFullName(fullName) {
  const patronymic = String(fullName).trim().split(WHITESPACE_RUN_REGEX)[PATRONYMIC_WORD_INDEX] || '';
  if (PATRONYMIC_MALE_REGEX.test(patronymic)) return SEX_MALE;
  if (PATRONYMIC_FEMALE_REGEX.test(patronymic)) return SEX_FEMALE;
  return '';
}

/**
 * Returns a copy of a row with its full-name cell (the column whose header matches
 * COL_FULL_NAME) passed through normalizeFullName(). Rows without such a column,
 * or with a short row not reaching it, are returned unchanged.
 * @param {string[]} headerRow - Column names (Database row 1).
 * @param {string[]} values - Cell values, one per column.
 * @returns {string[]}
 */
function normalizeRowFullName(headerRow, values) {
  const idx = findColumnIndex(headerRow, COL_FULL_NAME);
  const result = values.slice();
  if (idx >= 0 && idx < result.length) result[idx] = normalizeFullName(result[idx]);
  return result;
}

/**
 * Normalizes a Ukrainian phone number: adds a leading zero to a bare 9-digit
 * number, and strips the "38" country prefix from a 12-digit number. Anything
 * else is returned unchanged.
 *
 * @param {string} phone - Already-trimmed phone number.
 * @returns {string}
 */
function normalizePhoneNumber(phone) {
  if (PHONE_REGEX_9DIGIT.test(phone) && phone[0] !== '0') return '0' + phone;
  if (PHONE_REGEX_COUNTRY.test(phone)) return phone.slice(2);
  return phone;
}
