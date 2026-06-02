/**
 * Returns the list of source spreadsheet IDs from Handbook N2:N.
 * Used by Master Mode to aggregate data from multiple spreadsheets.
 *
 * @returns {string[]}
 */
function getMasterSources() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HANDBOOK);
  if (!sheet) return [];
  return sheet.getRange(MASTER_MODE_SOURCES_RANGE).getValues()
    .map(r => String(r[0]).trim())
    .filter(v => v !== '');
}

/**
 * Reads the Master Mode toggle from Handbook!M2.
 * When true, the webview aggregates data from all source spreadsheets listed in N2:N.
 *
 * @returns {boolean}
 */
function getMasterMode() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HANDBOOK);
  if (!sheet) return false;
  return sheet.getRange(MASTER_MODE_CELL).getValue() === true;
}

/**
 * Returns the column minimum width config for use in HTML template scriptlets.
 *
 * @returns {{ text: number, image: number, table: number }}
 */
function getColumnMinWidths() { return COLUMN_MIN_WIDTHS; }

/**
 * Returns the column maximum width config for use in HTML template scriptlets.
 *
 * @returns {{ image: number }}
 */
function getColumnMaxWidths() { return COLUMN_MAX_WIDTHS; }

/**
 * Simple trigger that runs when the spreadsheet is opened.
 * Adds the "More... ⭐️" custom menu to the Google Sheets toolbar.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('More... ⭐️')
    .addItem('Open Web Editor', 'openWebEditor')
    .addToUi();
}

/**
 * Opens the web editor as a full-screen modal dialog.
 * Uses createTemplateFromFile so that <?!= ?> scriptlet includes in
 * WebEditor.html are evaluated before the HTML is served to the client.
 */
function openWebEditor() {
  const html = HtmlService.createTemplateFromFile('WebEditor').evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, 'Web Editor');
}

/**
 * Returns the full schema and data from the "Database" sheet, plus sub-column
 * headers for any *-table columns resolved from the "Handbook" sheet.
 *
 * Sheet layout:
 *   Row 1 — column names
 *   Row 2 — column types (text | image | *-table)
 *   Row 3+ — data rows
 *
 * Handbook layout (row 1 is a header and is skipped):
 *   Column A — data type name (e.g. relatives-table)
 *   Column B+ — sub-column headers for that type
 *
 * @returns {{
 *   columns: Array<{name: string, type: string, tableHeaders?: string[]}>,
 *   rows: Array<{rowIndex: number, values: string[]}>
 * }}
 */
function getSchemaAndData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DATABASE);
  if (!sheet) throw new Error('Sheet "Database" not found.');
  const all = sheet.getDataRange().getValues();
  if (all.length < 2) throw new Error('Sheet must have at least 2 rows (names + types).');
  const columns = all[0].map((name, i) => ({
    name: String(name),
    type: String(all[1][i]).toLowerCase()
  }));
  const rows = [];
  for (let i = 2; i < all.length; i++) {
    const values = all[i].map(c => String(c == null ? '' : c));
    if (values.every(v => v === '')) continue;
    rows.push({ rowIndex: i + 1, values });
  }

  const masterMode = getMasterMode();
  if (masterMode) {
    getMasterSources().forEach(spreadsheetId => {
      try {
        const remoteSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(SHEET_DATABASE);
        if (!remoteSheet) return;
        const remoteAll = remoteSheet.getDataRange().getValues();
        for (let i = 2; i < remoteAll.length; i++) {
          const values = remoteAll[i].map(c => String(c == null ? '' : c));
          if (values.every(v => v === '')) continue;
          rows.push({ rowIndex: i + 1, values, spreadsheetId });
        }
      } catch (e) {
        // skip inaccessible spreadsheets
      }
    });
  }

  const tableHeadersMap = {};
  const handbookSheet = ss.getSheetByName(SHEET_HANDBOOK);
  if (handbookSheet) {
    const hbData = handbookSheet.getRange(HANDBOOK_TYPES_RANGE).getValues();
    for (let r = 0; r < hbData.length; r++) {
      const dataType = String(hbData[r][0]).trim().toLowerCase();
      if (!dataType) continue;
      const headers = hbData[r].slice(1).map(h => String(h)).filter(h => h !== '');
      if (headers.length) tableHeadersMap[dataType] = headers;
    }
  }
  columns.forEach(col => {
    if (col.type.endsWith('-table')) col.tableHeaders = tableHeadersMap[col.type] || [];
  });

  let unitOptions = [];
  if (handbookSheet) {
    unitOptions = handbookSheet.getRange(HANDBOOK_UNIT_RANGE).getValues()
      .map(r => String(r[0])).filter(v => v !== '');
  }
  columns.forEach(col => {
    if (col.type === 'unit') col.unitOptions = unitOptions;
  });

  let originOptions = [];
  if (handbookSheet) {
    originOptions = handbookSheet.getRange(HANDBOOK_ORIGIN_RANGE).getValues()
      .map(r => String(r[0])).filter(v => v !== '');
  }
  columns.forEach(col => {
    if (col.type === 'origin') col.originOptions = originOptions;
  });

  let maritalStatusOptions = [];
  if (handbookSheet) {
    maritalStatusOptions = handbookSheet.getRange(HANDBOOK_MARITAL_STATUS_RANGE).getValues()
      .map(r => String(r[0])).filter(v => v !== '');
  }
  columns.forEach(col => {
    if (col.type === 'marital-status') col.maritalStatusOptions = maritalStatusOptions;
  });

  let sexOptions = [];
  if (handbookSheet) {
    sexOptions = handbookSheet.getRange(HANDBOOK_SEX_RANGE).getValues()
      .map(r => String(r[0])).filter(v => v !== '');
  }
  columns.forEach(col => {
    if (col.type === 'sex') col.sexOptions = sexOptions;
  });

  return { columns, rows, masterMode, filterDebounceMs: FILTER_DEBOUNCE_MS,
           imageFetchBatchSize: IMAGE_FETCH_BATCH_SIZE, imageFetchConcurrency: IMAGE_FETCH_CONCURRENCY };
}

/**
 * Fetches Google Drive files by ID server-side. Images are returned as base64
 * data URLs; PDFs are returned as view URLs (no blob download). Running
 * server-side means the script owner's OAuth token is used, so all editor
 * users can view files regardless of their own Drive session.
 *
 * @param {string[]} fileIds - Array of Google Drive file IDs to fetch.
 * @returns {Object.<string, {type:string, dataUrl?:string, viewUrl?:string}|null>}
 *   Map of fileId → { type:'image', dataUrl } | { type:'pdf', viewUrl } | null on error.
 */
function getImagesDataUrls(fileIds) {
  const result = {};
  fileIds.forEach(fileId => {
    if (!fileId) return;
    try {
      const file = DriveApp.getFileById(fileId);
      const mimeType = file.getMimeType();
      if (mimeType === 'application/pdf') {
        result[fileId] = { type: 'pdf', viewUrl: 'https://drive.google.com/file/d/' + fileId + '/view' };
      } else if (mimeType === 'application/vnd.google-apps.folder') {
        result[fileId] = { type: 'folder', viewUrl: 'https://drive.google.com/drive/folders/' + fileId };
      } else {
        const blob = file.getBlob();
        result[fileId] = { type: 'image', dataUrl: 'data:' + (mimeType || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
      }
    } catch (e) {
      result[fileId] = { type: 'no-access' };
    }
  });
  return result;
}

/**
 * Appends a new empty row to the "Database" sheet and returns its 1-based row index.
 *
 * @returns {number} 1-based row index of the newly created row.
 */
function addRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATABASE);
  if (!sheet) throw new Error('Sheet "Database" not found.');
  const newRowIndex = sheet.getLastRow() + 1;
  const numCols = sheet.getLastColumn();
  sheet.getRange(newRowIndex, 1, 1, numCols).setValues([new Array(numCols).fill('')]);
  return newRowIndex;
}

/**
 * Writes new values for a single data row back to the "Database" sheet.
 * When spreadsheetId is provided, writes to that remote spreadsheet instead.
 *
 * @param {number} rowIndex - 1-based spreadsheet row number to update.
 * @param {string[]} values - Array of cell values, one per column.
 * @param {string|null} spreadsheetId - Remote spreadsheet ID, or null for local.
 * @returns {boolean} Always true; thrown errors propagate to the client failure handler.
 */
function updateRow(rowIndex, values, spreadsheetId) {
  const ss = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DATABASE);
  if (!sheet) throw new Error('Sheet "Database" not found.');
  sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
  return true;
}

/**
 * Moves a row from the "Database" sheet to the "Trash" sheet.
 * If the Trash sheet does not yet exist it is created with the same
 * header rows (rows 1 and 2) as the Database sheet.
 * When spreadsheetId is provided, operates on that remote spreadsheet instead.
 *
 * @param {number} rowIndex - 1-based spreadsheet row number to delete.
 * @param {string|null} spreadsheetId - Remote spreadsheet ID, or null for local.
 * @returns {boolean} Always true; thrown errors propagate to the client failure handler.
 */
function deleteRow(rowIndex, spreadsheetId) {
  const ss = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEET_DATABASE);
  if (!dbSheet) throw new Error('Sheet "Database" not found.');

  const numCols = dbSheet.getLastColumn();
  const rowData = dbSheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];

  let trashSheet = ss.getSheetByName(SHEET_TRASH);
  if (!trashSheet) {
    trashSheet = ss.insertSheet(SHEET_TRASH);
    const headers = dbSheet.getRange(1, 1, 2, numCols).getValues();
    trashSheet.getRange(1, 1, 2, numCols).setValues(headers);
  }

  const trashLastRow = Math.max(trashSheet.getLastRow(), 2);
  trashSheet.getRange(trashLastRow + 1, 1, 1, numCols).setValues([rowData]);

  dbSheet.deleteRow(rowIndex);
  return true;
}
