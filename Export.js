/**
 * Exports an F-1 document for each given row entry.
 *
 * @param {Array<{rowIndex: number, spreadsheetId: string|null}>} rowEntries
 * @returns {{results: Array<{name: string, url: string}>, remaining: Array<{rowIndex: number, spreadsheetId: string|null}>}}
 */
function exportF1(rowEntries) {
  return _exportDoc(rowEntries, F1_TEMPLATE_ID, F1_DOC_PREFIX);
}

/**
 * Exports a Wanted Card document for each given row entry.
 *
 * @param {Array<{rowIndex: number, spreadsheetId: string|null}>} rowEntries
 * @returns {{results: Array<{name: string, url: string}>, remaining: Array<{rowIndex: number, spreadsheetId: string|null}>}}
 */
function exportWC(rowEntries) {
  return _exportDoc(rowEntries, WC_TEMPLATE_ID, WC_DOC_PREFIX);
}

/**
 * Copies a Google Docs template for each given row entry, fills all placeholders
 * with row data, and saves the result to the export folder. Each source spreadsheet
 * is read at most once per call (cached by spreadsheet ID).
 *
 * Four passes run in order:
 *   1. Image-type columns — placeholder replaced with the actual image blob.
 *   2. Service history table — {COL_SERVICE_HISTORY} row expanded into table rows.
 *   3. Direct text columns — {column name} replaced with the cell value.
 *   4. Correspondence table — Handbook-defined aliases and computed values.
 *
 * Placeholders with no match are left untouched. The marital status line is
 * underlined based on the COL_MARITAL_STATUS value.
 *
 * @param {Array<{rowIndex: number, spreadsheetId: string|null}>} rowEntries
 * @param {string} templateId - Google Drive file ID of the Docs template.
 * @param {string} docPrefix - Prefix prepended to the first-column value to form the document name.
 * @returns {{results: Array<{name: string, url: string}>, remaining: Array<{rowIndex: number, spreadsheetId: string|null}>}}
 */
function _exportDoc(rowEntries, templateId, docPrefix) {
  const exportFolder = DriveApp.getFolderById(EXPORT_FOLDER_ID);
  const localSs = SpreadsheetApp.getActiveSpreadsheet();
  const handbookSheet = localSs.getSheetByName(SHEET_HANDBOOK);
  const mappings = handbookSheet ? _loadCorrespondenceTable(handbookSheet) : [];

  // Cache sheet data per spreadsheet to avoid redundant reads within the same export call.
  const sheetCache = new Map();
  /**
   * Returns cached { all, columns } for the given spreadsheet, or null if inaccessible.
   * @param {string|null} spreadsheetId
   * @returns {{all: Array, columns: Array}|null}
   */
  const getSheetData = spreadsheetId => {
    const key = spreadsheetId ?? '';
    if (!sheetCache.has(key)) {
      try {
        const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : localSs;
        const sheet = ss.getSheetByName(SHEET_DATABASE);
        if (!sheet) { sheetCache.set(key, null); return null; }
        const all = sheet.getDataRange().getValues();
        const columns = all[0].map((name, i) => ({
          name: String(name),
          type: String(all[1][i]).toLowerCase()
        }));
        sheetCache.set(key, { all, columns });
      } catch (e) {
        sheetCache.set(key, null);
      }
    }
    return sheetCache.get(key);
  };

  const results = [];
  const startTime = new Date();
  let remaining = [];

  for (let i = 0; i < rowEntries.length; i++) {
    if (new Date() - startTime > EXPORT_TIME_LIMIT_MS) {
      remaining = rowEntries.slice(i);
      break;
    }

    const { rowIndex, spreadsheetId } = rowEntries[i];
    const sheetData = getSheetData(spreadsheetId ?? null);
    if (!sheetData) continue;
    const { all, columns } = sheetData;

    const rowValues = all[rowIndex - 1];
    if (!rowValues) continue;
    const data = {};
    columns.forEach((col, j) => { data[col.name] = String(rowValues[j] == null ? '' : rowValues[j]); });

    const docName = docPrefix + (data[columns[0].name] || 'Unknown');
    const copy = DriveApp.getFileById(templateId).makeCopy(docName, exportFolder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    // Pass 1: image column placeholders.
    columns.forEach(col => {
      if (col.type !== 'image') return;
      const placeholder = '{' + col.name + '}';
      const fileId = parseDriveId(data[col.name] || '');
      if (fileId) {
        try {
          _replacePlaceholderWithImage(body, placeholder, DriveApp.getFileById(fileId).getBlob());
        } catch (e) {
          body.replaceText(_escapeRegex(placeholder), '');
        }
      } else {
        body.replaceText(_escapeRegex(placeholder), '');
      }
    });

    // Pass 2: service history table — must run before direct text replacement
    // so the {Проходження служби} placeholder is still intact.
    _fillServiceHistoryTable(body, data);

    // Pass 3: direct text column placeholders.
    columns.forEach(col => {
      if (col.type === 'image') return;
      body.replaceText(_escapeRegex('{' + col.name + '}'), _escapeReplacement(data[col.name] || ''));
    });

    // Pass 4: correspondence table mappings (source column aliases and computed values).
    mappings.forEach(({ placeholder, sourceCol, computedKey }) => {
      const value = sourceCol
        ? (data[sourceCol] || '')
        : _computeValue(computedKey, data);
      body.replaceText(_escapeRegex('{' + placeholder + '}'), _escapeReplacement(value));
    });

    _underlineMaritalStatus(body, data);

    doc.saveAndClose();
    results.push({ name: docName, url: copy.getUrl() });
  }

  return { results, remaining };
}

/**
 * Reads the placeholder correspondence table from Handbook (range defined by
 * HANDBOOK_CORR_RANGE, data only — header rows are excluded from the range).
 * Each data row maps a template placeholder to either a source Database column
 * header (column C) or a computed value key (column D).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} handbookSheet
 * @returns {Array<{placeholder: string, sourceCol: string, computedKey: string}>}
 */
function _loadCorrespondenceTable(handbookSheet) {
  const rows = handbookSheet.getRange(HANDBOOK_CORR_RANGE).getValues();
  const mappings = [];
  for (let i = 0; i < rows.length; i++) {
    const placeholder = String(rows[i][0]).trim();
    if (!placeholder) continue;
    mappings.push({
      placeholder,
      sourceCol: String(rows[i][1]).trim(),
      computedKey: String(rows[i][2]).trim()
    });
  }
  return mappings;
}

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
  if (name === 'spouseFullName') return _findRelativeField(data, 'дружина', 1) || _findRelativeField(data, 'чоловік', 1);
  if (name === 'spouseActualAddress') return _findRelativeField(data, 'дружина', 2) || _findRelativeField(data, 'чоловік', 2);
  if (name === 'motherPhoneNumber') return _findRelativeField(data, 'мати', 3);
  if (name === 'fatherPhoneNumber') return _findRelativeField(data, 'батько', 3);
  if (name === 'motherActualAddress') return _findRelativeField(data, 'мати', 2);
  if (name === 'fatherActualAddress') return _findRelativeField(data, 'батько', 2);
  if (name === 'spousePhoneNumber') return _findRelativeField(data, 'дружина', 3) || _findRelativeField(data, 'чоловік', 3);
  if (name === 'childrenNamesBirthDates') return _computeChildrenNamesBirthDates(data);
  if (name === 'childrenPhoneNumbers') return _computeChildrenPhoneNumbers(data);
  if (name === 'currentPosition') return _computeCurrentPosition(data);
  if (name === 'currentPositionStartDate') return _computeCurrentPositionStartDate(data);
  if (name === 'contractSignDate') return _computeContractSignDate(data);
  if (name === 'relativesWithPhoneNumbers') return _computeRelativesWithPhoneNumbers(data);
  return '';
}

/**
 * Computes total military service length from the "Дата призову" column value.
 * Extracts the last DD.MM.YYYY date found in the cell, then calculates the
 * calendar-accurate duration from that date to today.
 *
 * @param {Object.<string, string>} data - Row data map.
 * @returns {string} Formatted string, e.g. "3 роки, 8 місяців, 17 днів (станом на 09.05.2026)",
 *                   or empty string if no valid date is found.
 */
function _computeTotalServiceLength(data) {
  const raw = data[COL_DRAFT_DATE] || '';
  const matches = raw.match(/\d{2}\.\d{2}\.\d{4}/g);
  if (!matches) return '';

  const parts = matches[matches.length - 1].split('.');
  const start = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let years = today.getFullYear() - start.getFullYear();
  let months = today.getMonth() - start.getMonth();
  let days = today.getDate() - start.getDate();

  if (days < 0) {
    months--;
    const prevMonth = new Date(today.getFullYear(), today.getMonth(), 0);
    days += prevMonth.getDate();
  }
  if (months < 0) {
    years--;
    months += 12;
  }

  const dd = String(today.getDate()).padStart(2, '0');
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const yyyy = today.getFullYear();

  return `${_pluralizeUk(years, 'year')}, ${_pluralizeUk(months, 'month')}, ${_pluralizeUk(days, 'day')} (станом на ${dd}.${mm}.${yyyy})`;
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
  const contractField = (data[COL_CONTRACT_UNTIL] || '').toLowerCase();
  if (contractField.includes('мобілізований') || contractField.includes('мобілізована')) return '';

  const dateMatch = (data[COL_DRAFT_DATE] || '').match(/\d{2}\.\d{2}\.\d{4}/);
  if (!dateMatch) return '';

  const rows = _parseSubTable(data[COL_SERVICE_HISTORY] || '');
  let unitNumber = DEFAULT_UNIT_NUMBER;
  if (rows.length) {
    const unitMatch = (rows[0][1] || '').match(/\b(\d{4})\b/);
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
  const rows = _parseSubTable(data[COL_SERVICE_HISTORY] || '');
  if (!rows.length) return '';
  const period = rows[rows.length - 1][0] || '';
  const m = period.match(/\d{2}\.\d{2}\.\d{4}/);
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
  const rows = _parseSubTable(data[COL_SERVICE_HISTORY] || '');
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
  const rows = _parseSubTable(data[COL_CLOSE_RELATIVES] || '');
  return rows
    .filter(fields => (fields[0] || '').trim().toLowerCase().startsWith('дитина'))
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
  const rows = _parseSubTable(data[COL_CLOSE_RELATIVES] || '');
  const children = rows.filter(fields => (fields[0] || '').trim().toLowerCase().startsWith('дитина'));
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
  const rows = _parseSubTable(data[COL_CLOSE_RELATIVES] || '');
  return rows
    .filter(fields => (fields[3] || '').trim())
    .map(fields => [fields[0], fields[1], fields[2], fields[3]].map(f => (f || '').trim()).join(', '))
    .join('; ');
}

/**
 * Underlines the appropriate part of the marital status line
 * "одружений (заміжня)/ неодружений (незаміжня)" based on the "Сімейний стан"
 * column value. Married values (одружений/а, заміжня) underline the first part;
 * unmarried values (неодружений/а, незаміжня) underline the second part.
 *
 * @param {GoogleAppsScript.Document.Body} body - Document body to search.
 * @param {Object.<string, string>} data - Row data map.
 */
function _underlineMaritalStatus(body, data) {
  const status = (data[COL_MARITAL_STATUS] || '').toLowerCase().trim();
  let pattern;
  if (status.includes('неодружен') || status.includes('незаміжн')) {
    pattern = 'неодружений \\(незаміжня\\)';
  } else if (status.includes('одружен') || status.includes('заміжн')) {
    pattern = 'одружений \\(заміжня\\)';
  } else {
    return;
  }
  const found = body.findText(pattern);
  if (!found) return;
  found.getElement().asText().setUnderline(found.getStartOffset(), found.getEndOffsetInclusive(), true);
}

/**
 * Fills the "ПРОХОДЖЕННЯ СЛУЖБИ" table in the document by expanding the single
 * {Проходження служби} placeholder row into one row per service entry.
 * The placeholder row is reused for the first entry; additional rows are
 * inserted after it. Each row gets the period (field 0) in column 1 and the
 * position title (field 1) in column 2.
 *
 * @param {GoogleAppsScript.Document.Body} body - Document body to search.
 * @param {Object.<string, string>} data - Row data map.
 */
function _fillServiceHistoryTable(body, data) {
  const entries = _parseSubTable(data[COL_SERVICE_HISTORY] || '');

  const found = body.findText(_escapeRegex('{' + COL_SERVICE_HISTORY + '}'));
  if (!found) return;

  // Navigate up: Text → Paragraph → TableCell → TableRow → Table
  const placeholderRow = found.getElement().getParent().getParent().getParent();
  const table = placeholderRow.getParent();
  const rowIndex = table.getChildIndex(placeholderRow);

  if (entries.length === 0) {
    placeholderRow.getCell(0).setText('');
    placeholderRow.getCell(1).setText('');
    return;
  }

  // Fill the placeholder row with the first entry.
  placeholderRow.getCell(0).setText((entries[0][0] || '').trim());
  placeholderRow.getCell(1).setText((entries[0][1] || '').trim());

  // Insert remaining entries in reverse order at rowIndex+1 to preserve order.
  for (let i = entries.length - 1; i >= 1; i--) {
    const newRow = table.insertTableRow(rowIndex + 1);
    newRow.appendTableCell().setText((entries[i][0] || '').trim());
    newRow.appendTableCell().setText((entries[i][1] || '').trim());
  }
}

/**
 * Parses a pipe-and-newline encoded sub-table cell (the storage format used by
 * all *-table columns) into an array of field arrays.
 * Field separator: ' | '  Row separator: '\n'
 * Trailing empty fields from a trailing separator are preserved but harmless.
 *
 * @param {string} rawValue - Raw encoded cell string.
 * @returns {string[][]} Array of rows, each row being an array of field strings.
 */
function _parseSubTable(rawValue) {
  if (!rawValue) return [];
  return rawValue.split('\n')
    .map(row => row.split(' | '))
    .filter(fields => fields[0] && fields[0].trim());
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
  const rows = _parseSubTable(data[COL_CLOSE_RELATIVES] || '');
  const row = rows.find(fields => (fields[0] || '').trim().toLowerCase() === relationType.toLowerCase());
  return row ? (row[fieldIndex] || '').trim() : '';
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


/**
 * Escapes all Java regex metacharacters in a string so it can be passed as a
 * literal pattern to Body.replaceText() or Body.findText().
 *
 * @param {string} str - Raw string to escape.
 * @returns {string} Escaped pattern string.
 */
function _escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Escapes characters that are special in Java regex replacement strings
 * (backslash and dollar sign) so that Body.replaceText() treats the
 * replacement as a plain literal rather than a back-reference expression.
 *
 * @param {string} str - Replacement text to escape.
 * @returns {string} Escaped replacement string.
 */
function _escapeReplacement(str) {
  return str.replace(/\\/g, '\\\\').replace(/\$/g, '\\$');
}

/**
 * Finds the first paragraph in the document body that contains the given
 * placeholder text, clears it, and inserts the provided image blob inline.
 * If the inserted image's height exceeds 600 px the image is scaled down
 * proportionally to fit within that limit.
 *
 * @param {GoogleAppsScript.Document.Body} body - Document body to search.
 * @param {string} placeholder - Literal placeholder string, e.g. "{Фото}".
 * @param {GoogleAppsScript.Base.Blob} blob - Image blob to insert.
 */
function _replacePlaceholderWithImage(body, placeholder, blob) {
  const found = body.findText(_escapeRegex(placeholder));
  if (!found) return;
  const para = found.getElement().getParent().asParagraph();
  para.clear();
  const img = para.appendInlineImage(blob);
  const h = img.getHeight();
  if (h > IMAGE_MAX_HEIGHT) {
    img.setWidth(Math.round(img.getWidth() * IMAGE_MAX_HEIGHT / h));
    img.setHeight(IMAGE_MAX_HEIGHT);
  }
}
