/**
 * Exports an F-1 document for each given row index by copying the template,
 * replacing {column name} placeholders with row data, and saving to the export
 * folder. Image-type columns have their placeholder replaced with the actual image.
 * Placeholders mapped via the Handbook correspondence table are resolved before
 * the direct-column pass. Placeholders with no match are left untouched.
 *
 * @param {number[]} rowIndices - 1-based sheet row numbers to export.
 * @returns {Array<{name: string, url: string}>} Names and URLs of created documents.
 */
function exportF1(rowIndices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DATABASE);
  if (!sheet) throw new Error('Sheet "Database" not found.');
  const all = sheet.getDataRange().getValues();
  const columns = all[0].map((name, i) => ({
    name: String(name),
    type: String(all[1][i]).toLowerCase()
  }));
  const exportFolder = DriveApp.getFolderById(EXPORT_FOLDER_ID);

  const handbookSheet = ss.getSheetByName(SHEET_HANDBOOK);
  const mappings = handbookSheet ? _loadCorrespondenceTable(handbookSheet) : [];

  const results = [];

  rowIndices.forEach(rowIndex => {
    const rowValues = all[rowIndex - 1];
    const data = {};
    columns.forEach((col, i) => { data[col.name] = String(rowValues[i] == null ? '' : rowValues[i]); });

    const docName = EXPORT_DOC_PREFIX + (data[columns[0].name] || 'Unknown');
    const copy = DriveApp.getFileById(F1_TEMPLATE_ID).makeCopy(docName, exportFolder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    // Pass 1: image column placeholders.
    columns.forEach(col => {
      if (col.type !== 'image') return;
      const placeholder = '{' + col.name + '}';
      const fileId = _extractDriveId(data[col.name] || '');
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
  });

  return results;
}

/**
 * Reads the placeholder correspondence table from Handbook!B11:D50.
 * The first two rows (B11:D12) are headers and are skipped.
 * Each data row maps a template placeholder to either a source Database column
 * header (column C) or a computed value key (column D).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} handbookSheet
 * @returns {Array<{placeholder: string, sourceCol: string, computedKey: string}>}
 */
function _loadCorrespondenceTable(handbookSheet) {
  const rows = handbookSheet.getRange(HANDBOOK_CORR_ROW_START, HANDBOOK_CORR_COL_START, HANDBOOK_CORR_ROW_COUNT, HANDBOOK_CORR_COL_COUNT).getValues();
  const mappings = [];
  for (let i = 2; i < rows.length; i++) { // skip 2 header rows
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

  return (
    _pluralizeUk(years, 'year') + ', ' +
    _pluralizeUk(months, 'month') + ', ' +
    _pluralizeUk(days, 'day') +
    ' (станом на ' + dd + '.' + mm + '.' + yyyy + ')'
  );
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

  return dateMatch[0] + ' з в/ч ' + unitNumber;
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
    return (i + 1) + ' дитина: ' + name + (birthDate ? ' ' + birthDate : '');
  }).join('\n');
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
  return count + ' ' + word;
}

/**
 * Extracts a Google Drive file ID from a sharing URL or returns the input
 * unchanged if it already looks like a bare file ID.
 * Supported formats:
 *   https://drive.google.com/file/d/FILE_ID/view
 *   https://drive.google.com/open?id=FILE_ID
 *   FILE_ID (bare alphanumeric string ≥ 10 chars)
 *
 * @param {string} url - Drive sharing URL or raw file ID.
 * @returns {string|null} Extracted file ID, or null if not parseable.
 */
function _extractDriveId(url) {
  if (!url) return null;
  var m = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(url.trim())) return url.trim();
  return null;
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
