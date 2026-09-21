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
 *   sortRowsBySourceOrder() in WebEditor.list.js.html).
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
  const startTime = Date.now();

  rowEntries.forEach(entry => {
    if (Date.now() - startTime > EXPORT_TIME_LIMIT_MS) {
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
 * @param {(fileId: string) => ({type: string, viewUrl?: string})} resolveDriveInfo - Cached per-fileId Drive lookup.
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
