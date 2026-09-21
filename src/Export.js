/**
 * Exports an F-1 document for each given row entry.
 *
 * @param {Array<{rowIndex: number, spreadsheetId: string|null}>} rowEntries
 * @returns {{results: Array<{name: string, url: string}>, remaining: Array<{rowIndex: number, spreadsheetId: string|null}>}}
 */
function exportF1(rowEntries) {
  return _exportDoc(rowEntries, EXPORT_F1_TEMPLATE_CELL, F1_DOC_PREFIX);
}

/**
 * Exports a Wanted Card document for each given row entry.
 *
 * @param {Array<{rowIndex: number, spreadsheetId: string|null}>} rowEntries
 * @returns {{results: Array<{name: string, url: string}>, remaining: Array<{rowIndex: number, spreadsheetId: string|null}>}}
 */
function exportWC(rowEntries) {
  return _exportDoc(rowEntries, EXPORT_WC_TEMPLATE_CELL, WC_DOC_PREFIX);
}

/**
 * Builds a per-spreadsheet cache-and-loader for Database sheet data, shared by
 * document export (_exportDoc) and XLSX export (exportXLSX) so each distinct
 * source spreadsheet (by ID) is read from Sheets at most once per call.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} localSs - Used when spreadsheetId is null/undefined.
 * @returns {(spreadsheetId: string|null) => ({all: Array<Array<*>>, columns: Array<{name: string, type: string}>}|null)}
 */
function _makeSheetDataLoader(localSs) {
  const sheetCache = new Map();
  return spreadsheetId => {
    const key = spreadsheetId ?? '';
    if (!sheetCache.has(key)) {
      const ss = spreadsheetId ? openSpreadsheetSafely(spreadsheetId) : localSs;
      if (!ss) {
        sheetCache.set(key, null);
      } else {
        try {
          const sheet = ss.getSheetByName(SHEET_DATABASE);
          if (!sheet) {
            sheetCache.set(key, null);
          } else {
            const all = sheet.getDataRange().getValues();
            const columns = extractColumnSchema(all);
            sheetCache.set(key, { all, columns });
          }
        } catch (e) {
          sheetCache.set(key, null);
        }
      }
    }
    return sheetCache.get(key);
  };
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
 *      *-table columns are skipped here (their raw pipe/newline-encoded storage
 *      string is never a meaningful placeholder replacement); they're instead
 *      handled by pass 2 (service history) or pass 4 computed values (e.g.
 *      awardsList, relativesWithPhoneNumbers).
 *   4. Correspondence table — Handbook-defined aliases and computed values.
 *
 * Placeholders with no match are left untouched. The marital status line is
 * underlined based on the COL_MARITAL_STATUS value.
 *
 * @param {Array<{rowIndex: number, spreadsheetId: string|null}>} rowEntries
 * @param {string} templateCell - Handbook cell address (e.g. 'A9') holding the Docs template Drive ID.
 * @param {string} docPrefix - Prefix prepended to the first-column value to form the document name.
 * @returns {{results: Array<{name: string, url: string}>, remaining: Array<{rowIndex: number, spreadsheetId: string|null}>}}
 */
function _exportDoc(rowEntries, templateCell, docPrefix) {
  const localSs = SpreadsheetApp.getActiveSpreadsheet();
  const handbookSheet = localSs.getSheetByName(SHEET_HANDBOOK);
  const templateId = getDriveIdFromHandbook(handbookSheet, templateCell);
  const exportFolderId = getDriveIdFromHandbook(handbookSheet, EXPORT_FOLDER_CELL);
  const exportFolder = getExportFolderSafely(exportFolderId);
  const mappings = handbookSheet ? _loadCorrespondenceTable(handbookSheet) : [];
  const getSheetData = _makeSheetDataLoader(localSs);

  const results = [];
  const startTime = Date.now();
  let remaining = [];

  for (let i = 0; i < rowEntries.length; i++) {
    if (Date.now() - startTime > EXPORT_TIME_LIMIT_MS) {
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
    const strValues = stringifyRowValues(rowValues);
    columns.forEach((col, j) => { data[col.name] = strValues[j]; });

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
          _replacePlaceholderWithImage(body, placeholder, _getExportImageBlob(fileId));
        } catch (e) {
          body.replaceText(_escapeRegex(placeholder), '');
        }
      } else {
        body.replaceText(_escapeRegex(placeholder), '');
      }
    });

    // Pass 2: service history table — must run before pass 3 or the placeholder
    // text would be consumed before the table handler can locate it.
    _fillServiceHistoryTable(body, data);

    // Pass 3: direct text column placeholders. *-table columns are skipped —
    // their raw pipe/newline-encoded storage string would otherwise leak into
    // the document, and (for columns whose header matches a correspondence
    // table placeholder, e.g. awardsList) would also consume the placeholder
    // before pass 4 gets a chance to replace it with the computed value.
    columns.forEach(col => {
      if (col.type === 'image' || col.type.endsWith('-table')) return;
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
 * Underlines the appropriate part of the marital status line
 * "одружений (заміжня)/ неодружений (незаміжня)" based on the "Сімейний стан"
 * column value. Married values (одружений/а, заміжня) underline the first part;
 * unmarried values (неодружений/а, незаміжня) underline the second part.
 *
 * @param {GoogleAppsScript.Document.Body} body - Document body to search.
 * @param {Object.<string, string>} data - Row data map.
 */
function _underlineMaritalStatus(body, data) {
  const status = getFieldByPattern(data, COL_MARITAL_STATUS).toLowerCase().trim();
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
  const entries = _getServiceHistoryRows(data);

  const serviceHistoryKey = findKeyByPattern(data, COL_SERVICE_HISTORY);
  if (!serviceHistoryKey) return;

  const found = body.findText(_escapeRegex('{' + serviceHistoryKey + '}'));
  if (!found) return;

  // Navigate up: Text → Paragraph → TableCell → TableRow → Table
  const placeholderRow = /** @type {GoogleAppsScript.Document.TableRow} */ (found.getElement().getParent().getParent().getParent());
  const table = placeholderRow.getParent().asTable();
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
 * Returns a compressed image blob for the given Drive file ID, suitable for
 * inserting into an exported document. Requests a resized thumbnail via the
 * Drive advanced service (EXPORT_IMAGE_THUMBNAIL_SIZE px wide) instead of the
 * full-resolution original, since Google Docs embeds the blob's actual bytes
 * regardless of the display size set later. Falls back to the original
 * full-resolution blob if the thumbnail path fails for any reason.
 *
 * @param {string} fileId - Google Drive file ID of the image.
 * @returns {GoogleAppsScript.Base.Blob} Compressed thumbnail blob, or the
 *   original full-resolution blob if compression was not possible.
 */
function _getExportImageBlob(fileId) {
  try {
    const meta = Drive.Files.get(fileId, { fields: 'thumbnailLink' });
    const link = meta && meta.thumbnailLink;
    if (link) {
      const resized = link.replace(/=s\d+/, '=s' + EXPORT_IMAGE_THUMBNAIL_SIZE);
      const response = UrlFetchApp.fetch(resized, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      if (response.getResponseCode() === 200) {
        const blob = response.getBlob();
        if (blob && blob.getBytes().length > 0) return blob;
      }
    }
  } catch (e) {
    // Fall through to full-resolution fetch below.
  }
  return DriveApp.getFileById(fileId).getBlob();
}

/**
 * Finds the first paragraph in the document body that contains the given
 * placeholder text, clears it, and inserts the provided image blob inline.
 * If the inserted image's height exceeds IMAGE_MAX_HEIGHT px the image is
 * scaled down proportionally to fit within that limit.
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
