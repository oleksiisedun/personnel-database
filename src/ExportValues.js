/**
 * Dispatches to the appropriate computed-value function by key name.
 *
 * @param {string} name - Computed value key from the correspondence table.
 * @param {Object.<string, string>} data - Map of column name to cell value for the current row.
 * @returns {string} Computed replacement value, or empty string if key is unknown.
 */
function _computeValue(name, data) {
  if (name === 'totalServiceLength') return _computeTotalServiceLength(data);
  if (name === 'motherFullName') return _findRelativeField(data, 'мати', 1);
  if (name === 'fatherFullName') return _findRelativeField(data, 'батько', 1);
  if (name === 'spouseFullName') return _findSpouseField(data, 1);
  if (name === 'spouseActualAddress') return _findSpouseField(data, 2);
  if (name === 'motherPhoneNumber') return _findRelativeField(data, 'мати', 3);
  if (name === 'fatherPhoneNumber') return _findRelativeField(data, 'батько', 3);
  if (name === 'motherActualAddress') return _findRelativeField(data, 'мати', 2);
  if (name === 'fatherActualAddress') return _findRelativeField(data, 'батько', 2);
  if (name === 'spousePhoneNumber') return _findSpouseField(data, 3);
  if (name === 'childrenNamesBirthDates') return _computeChildrenNamesBirthDates(data);
  if (name === 'childrenPhoneNumbers') return _computeChildrenPhoneNumbers(data);
  if (name === 'currentPosition') return _computeCurrentPosition(data);
  if (name === 'currentPositionStartDate') return _computeCurrentPositionStartDate(data);
  if (name === 'contractSignDate') return _computeContractSignDate(data);
  if (name === 'relativesWithPhoneNumbers') return _computeRelativesWithPhoneNumbers(data);
  if (name === 'awardsList') return _computeAwardsList(data);
  return '';
}

/**
 * Parses the service history sub-table for the given row data.
 * @param {Object.<string, string>} data - Row data map.
 * @returns {string[][]}
 */
function _getServiceHistoryRows(data) {
  return _parseSubTable(getFieldByPattern(data, COL_SERVICE_HISTORY));
}

/**
 * Parses the close relatives sub-table for the given row data.
 * @param {Object.<string, string>} data - Row data map.
 * @returns {string[][]}
 */
function _getRelativesRows(data) {
  return _parseSubTable(getFieldByPattern(data, COL_CLOSE_RELATIVES));
}

/**
 * Parses the awards sub-table for the given row data.
 * @param {Object.<string, string>} data - Row data map.
 * @returns {string[][]}
 */
function _getAwardsRows(data) {
  return _parseSubTable(getFieldByPattern(data, COL_AWARDS));
}

/**
 * Filters a relatives table row array to only rows where the first field starts
 * with "дитина" (case-insensitive).
 * @param {string[][]} rows - Parsed relatives sub-table rows.
 * @returns {string[][]}
 */
function _filterChildrenRows(rows) {
  return rows.filter(fields => (fields[0] || '').trim().toLowerCase().startsWith('дитина'));
}

/**
 * Computes total military service length from the "Дата призову" column value.
 * Extracts the last DD.MM.YYYY date found in the cell, then calculates the
 * calendar-accurate duration from that date to today.
 *
 * @param {Object.<string, string>} data - Row data map.
 * @param {Date} [today] - Reference date; defaults to the current date. Injectable for tests.
 * @returns {string} Formatted string, e.g. "3 роки, 8 місяців, 17 днів (станом на 09.05.2026)",
 *                   or empty string if no valid date is found.
 */
function _computeTotalServiceLength(data, today = new Date()) {
  const raw = getFieldByPattern(data, COL_DRAFT_DATE);
  const matches = raw.match(DATE_REGEX);
  if (!matches) return '';

  const parts = matches[matches.length - 1].split('.');
  const start = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const { years, months, days } = _calendarDuration(start, end);

  return `${_pluralizeUk(years, 'year')}, ${_pluralizeUk(months, 'month')}, ${_pluralizeUk(days, 'day')} (станом на ${formatDateDDMMYYYY(end)})`;
}

/**
 * Calendar-accurate difference between two dates as whole years, whole months
 * and leftover days: the whole years+months are added to `start` (clamping to
 * the last day of a shorter month, so 31 Jan + 1 month = 28/29 Feb), and the
 * days are counted from that anchor to `end`. Never returns negative parts.
 * Returns zeros if `end` is before `start`.
 *
 * @param {Date} start - Local-midnight start date.
 * @param {Date} end - Local-midnight end date.
 * @returns {{years: number, months: number, days: number}}
 */
function _calendarDuration(start, end) {
  if (end < start) return { years: 0, months: 0, days: 0 };

  let totalMonths = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
  if (end.getDate() < start.getDate()) totalMonths--;

  const anchorMonthIndex = start.getMonth() + totalMonths;
  const lastDayOfAnchorMonth = new Date(start.getFullYear(), anchorMonthIndex + 1, 0).getDate();
  const anchor = new Date(start.getFullYear(), anchorMonthIndex, Math.min(start.getDate(), lastDayOfAnchorMonth));

  // UTC arithmetic keeps a DST change between the two dates from skewing the day count.
  const msPerDay = 24 * 60 * 60 * 1000;
  const days = Math.round((Date.UTC(end.getFullYear(), end.getMonth(), end.getDate())
    - Date.UTC(anchor.getFullYear(), anchor.getMonth(), anchor.getDate())) / msPerDay);

  return { years: Math.floor(totalMonths / 12), months: totalMonths % 12, days };
}

/**
 * Returns the contract sign date and military unit for the F-1 form.
 * Returns empty string if the person was mobilised ("мобілізований/а").
 * The date is taken from the first DD.MM.YYYY date in "Дата призову".
 * The unit number is the first 4-digit number found in the first service
 * position record; falls back to "3102" if none found.
 *
 * @param {Object.<string, string>} data - Row data map.
 * @returns {string} e.g. "07.05.2015 з в/ч 3011", or empty string.
 */
function _computeContractSignDate(data) {
  const contractField = getFieldByPattern(data, COL_CONTRACT_UNTIL).toLowerCase();
  if (contractField.includes('мобілізований') || contractField.includes('мобілізована')) return '';

  const dateMatch = getFieldByPattern(data, COL_DRAFT_DATE).match(DATE_REGEX);
  if (!dateMatch) return '';

  const rows = _getServiceHistoryRows(data);
  let unitNumber = DEFAULT_UNIT_NUMBER;
  if (rows.length) {
    const unitMatch = (rows[0][1] || '').match(UNIT_NUMBER_REGEX);
    if (unitMatch) unitNumber = unitMatch[1];
  }

  return `${dateMatch[0]} з в/ч ${unitNumber}`;
}

/**
 * Returns the start date of the last (current) entry in the "Проходження служби"
 * column by extracting the first DD.MM.YYYY date from the period field.
 *
 * @param {Object.<string, string>} data - Row data map.
 * @returns {string} Date string e.g. "01.01.2025", or empty string if not found.
 */
function _computeCurrentPositionStartDate(data) {
  const rows = _getServiceHistoryRows(data);
  if (!rows.length) return '';
  const period = rows[rows.length - 1][0] || '';
  const m = period.match(DATE_REGEX);
  return m ? m[0] : '';
}

/**
 * Returns the position title from the last entry in the "Проходження служби"
 * column (the most recent / current position).
 * Each row is encoded as "period | position title".
 *
 * @param {Object.<string, string>} data - Row data map.
 * @returns {string} Trimmed position title, or empty string if not found.
 */
function _computeCurrentPosition(data) {
  const rows = _getServiceHistoryRows(data);
  if (!rows.length) return '';
  return (rows[rows.length - 1][1] || '').trim();
}

/**
 * Returns a comma-separated list of phone numbers for all children found in the
 * "Близькі родичі" sub-table. Children with no phone number are skipped.
 *
 * @param {Object.<string, string>} data - Row data map.
 * @returns {string} e.g. "0671234567, 0991234567", or empty string if none.
 */
function _computeChildrenPhoneNumbers(data) {
  return _filterChildrenRows(_getRelativesRows(data))
    .map(fields => (fields[3] || '').trim())
    .filter(phone => phone)
    .join(', ');
}

/**
 * Builds a numbered list of children's full names and birth dates from the
 * "Близькі родичі" sub-table. All rows with relation type "дитина" are included
 * in the order they appear.
 *
 * @param {Object.<string, string>} data - Row data map.
 * @returns {string} Multi-line string, e.g. "1 дитина: Іванова Анна Іванівна 20.01.2003\n2 дитина: ..."
 *                   or empty string if no children found.
 */
function _computeChildrenNamesBirthDates(data) {
  const children = _filterChildrenRows(_getRelativesRows(data));
  if (!children.length) return '';
  return children.map((fields, i) => {
    const name = (fields[1] || '').trim();
    const birthDate = (fields[4] || '').trim();
    return `${i + 1} дитина: ${name}${birthDate ? ` ${birthDate}` : ''}`;
  }).join('\n');
}

/**
 * Returns a semicolon-separated list of all relatives who have a phone number.
 * Each entry is formatted as: "relation, full name, address, phone".
 *
 * @param {Object.<string, string>} data - Row data map.
 * @returns {string} e.g. "мати, Іванова Марія, Київ, 0671234567; батько, Іванов Петро, Львів, 0991234567"
 *                   or empty string if no relatives have a phone number.
 */
function _computeRelativesWithPhoneNumbers(data) {
  const rows = _getRelativesRows(data);
  return rows
    .filter(fields => (fields[3] || '').trim())
    .map(fields => [fields[0], fields[1], fields[2], fields[3]].map(f => (f || '').trim()).join(', '))
    .join('; ');
}

/**
 * Builds a semicolon-separated sentence listing all awards from the "Нагороди"
 * sub-table. Each row is encoded as "award name | order number | order date"
 * (order number and/or date may be empty). Each entry is formatted as the
 * award name (first letter capitalized) followed by "№{order number}" and
 * "від {order date}" when present.
 *
 * @param {Object.<string, string>} data - Row data map.
 * @returns {string} e.g. "Нагрудний знак «За доблесну службу» №421 від 28.11.2023; ..."
 *                   or empty string if no awards found.
 */
function _computeAwardsList(data) {
  return _getAwardsRows(data).map(fields => {
    const name = (fields[0] || '').trim();
    const number = (fields[1] || '').trim();
    const date = (fields[2] || '').trim();
    let text = name.charAt(0).toUpperCase() + name.slice(1);
    if (number) text += ` №${number}`;
    if (date) text += ` від ${date}`;
    return text;
  }).join('; ');
}

/**
 * Looks up a relative by relation type in the "Близькі родичі" sub-table and
 * returns the value of the requested field index.
 * Field layout (by position): 0=relation, 1=full name, 2=address, 3=phone, 4=birth date.
 *
 * @param {Object.<string, string>} data - Row data map.
 * @param {string} relationType - Relation type to match in field 0, e.g. 'мати'.
 * @param {number} fieldIndex - 0-based index of the field to return.
 * @returns {string} Trimmed field value, or empty string if not found.
 */
function _findRelativeField(data, relationType, fieldIndex) {
  const rows = _getRelativesRows(data);
  const row = rows.find(fields => (fields[0] || '').trim().toLowerCase() === relationType.toLowerCase());
  return row ? (row[fieldIndex] || '').trim() : '';
}

/**
 * Looks up the requested field for a spouse, trying both "дружина" (wife) and
 * "чоловік" (husband) relation types since only one will be present per row.
 *
 * @param {Object.<string, string>} data - Row data map.
 * @param {number} fieldIndex - 0-based index of the field to return.
 * @returns {string} Trimmed field value, or empty string if not found.
 */
function _findSpouseField(data, fieldIndex) {
  return _findRelativeField(data, 'дружина', fieldIndex) || _findRelativeField(data, 'чоловік', fieldIndex);
}

/**
 * Returns a Ukrainian pluralized string for a count and unit.
 *
 * @param {number} count
 * @param {'year'|'month'|'day'} unit
 * @returns {string} e.g. "3 роки", "1 місяць", "5 днів"
 */
function _pluralizeUk(count, unit) {
  const mod10 = count % 10;
  const mod100 = count % 100;
  const isTeen = mod100 >= 11 && mod100 <= 19;
  let word;
  if (unit === 'year') {
    if (!isTeen && mod10 === 1) word = 'рік';
    else if (!isTeen && mod10 >= 2 && mod10 <= 4) word = 'роки';
    else word = 'років';
  } else if (unit === 'month') {
    if (!isTeen && mod10 === 1) word = 'місяць';
    else if (!isTeen && mod10 >= 2 && mod10 <= 4) word = 'місяці';
    else word = 'місяців';
  } else {
    if (!isTeen && mod10 === 1) word = 'день';
    else if (!isTeen && mod10 >= 2 && mod10 <= 4) word = 'дні';
    else word = 'днів';
  }
  return `${count} ${word}`;
}
