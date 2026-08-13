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
 * @returns {function(string|null): ({all: Array<Array<*>>, columns: Array<{name: string, type: string}>}|null)}
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
 * @param {string} templateCell - Handbook cell address (e.g. 'M9') holding the Docs template Drive ID.
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
 * Exports the given rows, restricted to visibleColumnIndices, as a single .xlsx
 * file saved into the export folder (Handbook!EXPORT_FOLDER_CELL). Apps Script
 * has no way to author .xlsx bytes directly, so the grid is built in a
 * temporary Google Sheet and converted via _fetchXlsxExportBlob() (the Sheets
 * export URL — Blob.getAs() does NOT support Sheet→xlsx conversion, only a
 * narrow set of formats such as PDF/image); the temp spreadsheet is always
 * trashed (try/finally), even on error.
 *
 * image-type column cells, and any other non-table column's cell whose raw
 * value looks like a Drive sharing URL (looksLikeDriveUrl()), are written as
 * =HYPERLINK("driveViewUrl","rawValue") formulas instead of plain text.
 * *-table column cells are always written verbatim — their pipe/newline-encoded
 * string is the desired xlsx cell content per spec, never reformatted.
 *
 * The grid is written in two passes, deliberately never mixed into a single
 * Range.setFormulas() call: (1) the whole grid as plain VALUES via
 * setValues(), which never reinterprets a string as a formula regardless of
 * its content — unlike setFormulas(), which parses every cell as though it
 * were typed into the formula bar even without a leading '=' (a bare word
 * like a column header gets looked up as a named range and fails with
 * #NAME?, multi-word text fails to parse as #ERROR!); (2) individual
 * setFormula() calls only for the specific cells that resolved to a
 * HYPERLINK() formula, overwriting their pass-1 placeholder value.
 *
 * Because a brand-new spreadsheet defaults to Automatic number formatting
 * (which would silently mangle text-shaped values like phone numbers, TINs, or
 * dates-as-text — e.g. dropping a leading zero), the destination data range is
 * force-formatted as plain text ('@') before the setValues() pass, then only
 * the specific cells that end up holding a HYPERLINK() formula are reset to
 * 'General' format so those formulas evaluate instead of displaying as
 * literal text.
 *
 * Each distinct source spreadsheet (by spreadsheetId) is read at most once, and
 * each distinct Drive fileId's link classification is resolved at most once,
 * within a single call.
 *
 * Unlike _exportDoc(), there is no partial/continuation result: if
 * EXPORT_TIME_LIMIT_MS is exceeded while building the grid this throws (nothing
 * is saved, but the temp spreadsheet is still cleaned up) — a half-built xlsx
 * has no standalone value the way a partial batch of Docs does.
 *
 * @param {Array<{rowIndex: number, spreadsheetId: string|null}>} rowEntries - Rows to export, in the
 *   exact order they should appear in the sheet (caller orders them — see
 *   sortRowsBySourceOrder() in WebEditor.js.html).
 * @param {number[]} visibleColumnIndices - Indices into the local Database sheet's column
 *   schema to include, in order (columns the user has not hidden).
 * @returns {{name: string, url: string}} The saved .xlsx file's name and Drive URL.
 */
function exportXLSX(rowEntries, visibleColumnIndices) {
  const localSs = SpreadsheetApp.getActiveSpreadsheet();
  const handbookSheet = localSs.getSheetByName(SHEET_HANDBOOK);
  const exportFolderId = getDriveIdFromHandbook(handbookSheet, EXPORT_FOLDER_CELL);
  const exportFolder = getExportFolderSafely(exportFolderId);

  const getSheetData = _makeSheetDataLoader(localSs);
  const localData = getSheetData(null);
  if (!localData) throw new Error('Local "Database" sheet not accessible.');

  const colIndices = visibleColumnIndices.filter(i => Number.isInteger(i) && i >= 0 && i < localData.columns.length);
  if (!colIndices.length) throw new Error('No columns selected for export.');
  const headerRow = colIndices.map(i => localData.columns[i].name);

  const driveInfoCache = new Map();
  const resolveDriveInfo = fileId => {
    if (!driveInfoCache.has(fileId)) driveInfoCache.set(fileId, resolveDriveFileForExport(fileId));
    return driveInfoCache.get(fileId);
  };

  const grid = [headerRow];
  const linkCells = []; // { a1: string, row: number, col: number, formula: string }
  const startTime = new Date();

  rowEntries.forEach(entry => {
    if (new Date() - startTime > EXPORT_TIME_LIMIT_MS) {
      throw new Error(`XLSX export timed out after building ${grid.length - 1} of ${rowEntries.length} rows. Select fewer rows and try again.`);
    }
    const sheetData = getSheetData(entry.spreadsheetId ?? null);
    if (!sheetData) return;
    const rowValues = sheetData.all[entry.rowIndex - 1];
    if (!rowValues) return;
    const strValues = stringifyRowValues(rowValues);
    const rowNum = grid.length + 1; // 1-based sheet row this data row will occupy
    const rowCells = colIndices.map((colIdx, cellIdx) => {
      const col = sheetData.columns[colIdx];
      const raw = strValues[colIdx] || '';
      const linkFormula = _buildXlsxLinkCell(col, raw, resolveDriveInfo);
      if (linkFormula !== null) {
        const colNum = cellIdx + 1;
        linkCells.push({ a1: _colIndexToA1Column(cellIdx) + rowNum, row: rowNum, col: colNum, formula: linkFormula });
      }
      return raw; // always the plain value — link cells get overwritten with a formula in the second pass below
    });
    grid.push(rowCells);
  });

  const fileName = XLSX_EXPORT_FILENAME_PREFIX + formatDateDDMMYYYY(new Date()) + '.xlsx';
  const tempName = 'tmp-xlsx-export-' + Utilities.getUuid();
  let tempSs = null;
  try {
    tempSs = SpreadsheetApp.create(tempName);
    const sheet = tempSs.getSheets()[0];
    const range = sheet.getRange(1, 1, grid.length, headerRow.length);
    range.setNumberFormat('@');
    range.setValues(grid);
    if (linkCells.length) {
      sheet.getRangeList(linkCells.map(c => c.a1)).setNumberFormat('General');
      linkCells.forEach(({ row, col, formula }) => sheet.getRange(row, col).setFormula(formula));
    }
    SpreadsheetApp.flush();

    const blob = _fetchXlsxExportBlob(tempSs.getId()).setName(fileName);
    const file = exportFolder.createFile(blob);
    return { name: fileName, url: file.getUrl() };
  } finally {
    if (tempSs) {
      try {
        DriveApp.getFileById(tempSs.getId()).setTrashed(true);
      } catch (e) {
        // best-effort cleanup — never let this mask the real error
      }
    }
  }
}

/**
 * Discovers every Actual Personnel person (local Database + every Master Mode
 * source in Handbook!N2:N, processed in that fixed order so duplicate S-КАДР
 * ID resolution is reproducible run to run) who has both a non-empty "Фото"
 * and "S-КАДР ID" value, creates the destination Drive folder, and returns
 * the full eligible list plus a skip list. No image bytes are fetched here —
 * only sheet values — so the complete eligible list (and therefore duplicate
 * detection) is known before any photo copying starts. Unlike _exportDoc(),
 * this has no continuation: a half-scanned personnel list has no standalone
 * value, so it throws instead of returning partial results if the scan
 * exceeds EXPORT_TIME_LIMIT_MS (same choice exportXLSX() makes).
 *
 * @returns {{
 *   folderId: string,
 *   folderUrl: string,
 *   eligible: Array<{rowIndex: number, spreadsheetId: string|null, fullName: string, cardId: string, fileId: string}>,
 *   skipped: Array<{name: string, reason: string}>
 * }}
 */
function startPhotoExport() {
  const localSs = SpreadsheetApp.getActiveSpreadsheet();
  const handbookSheet = localSs.getSheetByName(SHEET_HANDBOOK);

  const exportFolderId = getDriveIdFromHandbook(handbookSheet, EXPORT_FOLDER_CELL);
  const parentFolder = getExportFolderSafely(exportFolderId);
  const destFolder = parentFolder.createFolder(PHOTO_EXPORT_FOLDER_PREFIX + formatDateDDMMYYYY(new Date()));

  const actualNames = getActualPersonnelNames();
  if (!actualNames) {
    throw new Error('Actual Personnel list (Handbook!M6/M7) is not configured or not accessible — cannot determine who to export.');
  }
  const actualSet = new Set(actualNames.map(n => n.trim().toUpperCase()));

  const sourceIds = [null, ...getMasterSources()]; // null = local Database; fixed, deterministic order
  const getSheetData = _makeSheetDataLoader(localSs);

  const seenCardIds = new Map(); // cardId -> fullName of first winner
  const eligible = [];
  const skipped = [];
  const startTime = new Date();

  sourceIds.forEach((spreadsheetId, idx) => {
    if (new Date() - startTime > EXPORT_TIME_LIMIT_MS) {
      throw new Error(`Timed out scanning personnel (${idx} of ${sourceIds.length} sources completed).`);
    }
    const sheetData = getSheetData(spreadsheetId);
    if (!sheetData) return; // inaccessible/missing source — silently skipped, same convention as getMasterSourceRows()
    const { all, columns } = sheetData;

    for (let i = 2; i < all.length; i++) {
      const strValues = stringifyRowValues(all[i]);
      if (strValues.every(v => v === '')) continue;
      const data = {};
      columns.forEach((col, j) => { data[col.name] = strValues[j]; });

      const fullName = (data[columns[0].name] || '').trim();
      if (!fullName || !actualSet.has(fullName.toUpperCase())) continue;

      const photoRaw = getFieldByPattern(data, COL_PHOTO).trim();
      const cardId = getFieldByPattern(data, COL_CARD_ID).trim();

      if (!photoRaw || !cardId) {
        const reason = !photoRaw && !cardId ? 'Missing photo and S-КАДР ID'
                     : !photoRaw ? 'Missing photo' : 'Missing S-КАДР ID';
        skipped.push({ name: fullName, reason });
        continue;
      }
      if (seenCardIds.has(cardId)) {
        skipped.push({ name: fullName, reason: `Duplicate S-КАДР ID "${cardId}" (already used by ${seenCardIds.get(cardId)})` });
        continue;
      }
      seenCardIds.set(cardId, fullName);
      eligible.push({ rowIndex: i + 1, spreadsheetId, fullName, cardId, fileId: parseDriveId(photoRaw) });
    }
  });

  return { folderId: destFolder.getId(), folderUrl: destFolder.getUrl(), eligible, skipped };
}

/**
 * Copies and JPEG-converts photos for one batch of eligible entries (as
 * produced by startPhotoExport()) into the given folder, continuing across
 * multiple calls via the same time-boxed pattern _exportDoc() uses. Fetches
 * the full-resolution original blob (not the 800px thumbnail path
 * _getExportImageBlob() uses for Doc exports) and force-converts it to JPEG,
 * since the source file's original format is not guaranteed to be JPEG.
 *
 * @param {string} folderId - Destination Drive folder ID (from startPhotoExport()).
 * @param {Array<{rowIndex: number, spreadsheetId: string|null, fullName: string, cardId: string, fileId: string}>} entries
 * @returns {{
 *   results: Array<{name: string, url: string}>,
 *   skipped: Array<{name: string, reason: string}>,
 *   remaining: Array<{rowIndex: number, spreadsheetId: string|null, fullName: string, cardId: string, fileId: string}>
 * }}
 */
function copyPhotosBatch(folderId, entries) {
  const folder = DriveApp.getFolderById(folderId);
  const results = [];
  const skipped = [];
  const startTime = new Date();
  let remaining = [];

  for (let i = 0; i < entries.length; i++) {
    if (new Date() - startTime > EXPORT_TIME_LIMIT_MS) {
      remaining = entries.slice(i);
      break;
    }
    const { fullName, cardId, fileId } = entries[i];

    let blob;
    try {
      blob = DriveApp.getFileById(fileId).getBlob();
    } catch (e) {
      skipped.push({ name: fullName, reason: 'Photo file not accessible' });
      continue;
    }

    let jpegBlob;
    try {
      jpegBlob = blob.getAs('image/jpeg'); // force-convert regardless of source format
    } catch (e) {
      skipped.push({ name: fullName, reason: 'Could not convert photo to JPEG: ' + e.message });
      continue;
    }

    jpegBlob.setName(cardId + '.jpg');
    const file = folder.createFile(jpegBlob);
    results.push({ name: `${fullName} (${cardId}.jpg)`, url: file.getUrl() });
  }

  return { results, skipped, remaining };
}

/**
 * Fetches a spreadsheet's .xlsx bytes via the Sheets export URL, authorized
 * with the script's own OAuth token — the same authenticated-fetch pattern
 * _getExportImageBlob() uses for Drive thumbnails. This is necessary because
 * Blob.getAs() does not support converting a Google Sheet to
 * application/vnd.openxmlformats-officedocument.spreadsheetml.sheet (it only
 * supports a narrow set of conversions, mainly to PDF/image formats) — calling
 * it throws "Конвертування ... не підтримується."
 *
 * @param {string} spreadsheetId
 * @returns {GoogleAppsScript.Base.Blob}
 */
function _fetchXlsxExportBlob(spreadsheetId) {
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=xlsx';
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Failed to convert spreadsheet to XLSX (HTTP ' + response.getResponseCode() + ').');
  }
  return response.getBlob();
}

/**
 * Builds a HYPERLINK() formula string for a cell whose raw value resolves to a
 * Drive file/folder, or returns null when the cell should be written as a
 * plain value. image-type columns always attempt Drive resolution (their
 * stored value is a bare Drive ID or a Drive URL, same convention as
 * _exportDoc()'s pass 1). Any other non-table column is only treated as a link
 * when its raw value itself looks like a Drive sharing URL. *-table columns
 * are never linkified — their encoded string is always written verbatim.
 * When the fileId can't be resolved (no-access / lookup error), falls back to
 * null (raw value stays as plain text) rather than dropping the cell content.
 *
 * @param {{name: string, type: string}} col
 * @param {string} raw - Raw stringified cell value.
 * @param {function(string): ({type: string, viewUrl?: string})} resolveDriveInfo - Cached per-fileId Drive lookup.
 * @returns {string|null} An `=HYPERLINK("url","display")` formula, or null.
 */
function _buildXlsxLinkCell(col, raw, resolveDriveInfo) {
  if (!raw || col.type.endsWith('-table')) return null;
  let fileId = null;
  if (col.type === 'image') {
    fileId = parseDriveId(raw);
  } else if (looksLikeDriveUrl(raw)) {
    fileId = parseDriveId(raw);
  }
  if (!fileId) return null;
  const info = resolveDriveInfo(fileId);
  if (!info || info.type === 'no-access' || !info.viewUrl) return null;
  return `=HYPERLINK("${_escapeFormulaString(info.viewUrl)}","${_escapeFormulaString(raw)}")`;
}

/**
 * Escapes double quotes for embedding a string as a literal inside a Sheets
 * formula (e.g. a HYPERLINK() argument), by doubling them per formula-string
 * escaping rules.
 *
 * @param {string} str
 * @returns {string}
 */
function _escapeFormulaString(str) {
  return String(str).replace(/"/g, '""');
}

/**
 * Converts a 0-based column index to its spreadsheet column letter(s)
 * (0 → "A", 25 → "Z", 26 → "AA"), for building A1 cell addresses.
 *
 * @param {number} index - 0-based column index.
 * @returns {string}
 */
function _colIndexToA1Column(index) {
  let letters = '';
  let n = index;
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return letters;
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
 * @returns {string} Formatted string, e.g. "3 роки, 8 місяців, 17 днів (станом на 09.05.2026)",
 *                   or empty string if no valid date is found.
 */
function _computeTotalServiceLength(data) {
  const raw = getFieldByPattern(data, COL_DRAFT_DATE);
  const matches = raw.match(DATE_REGEX);
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
 * Field separator: TABLE_FIELD_SEP  Row separator: TABLE_ROW_SEP
 * Trailing empty fields from a trailing separator are preserved but harmless.
 *
 * @param {string} rawValue - Raw encoded cell string.
 * @returns {string[][]} Array of rows, each row being an array of field strings.
 */
function _parseSubTable(rawValue) {
  if (!rawValue) return [];
  return rawValue.split(TABLE_ROW_SEP)
    .map(row => row.split(TABLE_FIELD_SEP))
    .filter(fields => fields[0] && fields[0].trim());
}

/**
 * Encodes an array of field arrays back into a sub-table cell string — the
 * inverse of _parseSubTable().
 *
 * @param {string[][]} rows - Array of rows, each row being an array of field strings.
 * @returns {string} Pipe-and-newline encoded cell string.
 */
function _encodeSubTable(rows) {
  return rows.map(fields => fields.join(TABLE_FIELD_SEP)).join(TABLE_ROW_SEP);
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
