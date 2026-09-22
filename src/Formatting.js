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
