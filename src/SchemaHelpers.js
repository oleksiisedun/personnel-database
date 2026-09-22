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
 * @returns {Array<{name: string, type: string, dropdownOptions?: string[], tableHeaders?: Array<{name: string, type: string, dropdownOptions?: string[]}>}>}
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
