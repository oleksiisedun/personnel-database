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
  const html = HtmlService.createTemplateFromFile('WebEditor').evaluate()
    .setWidth(800)
    .setHeight(600);
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
  const sheet = ss.getSheetByName('Database');
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

  const tableHeadersMap = {};
  const handbookSheet = ss.getSheetByName('Handbook');
  if (handbookSheet) {
    const hbData = handbookSheet.getDataRange().getValues();
    for (let r = 1; r < hbData.length; r++) {
      const dataType = String(hbData[r][0]).trim().toLowerCase();
      if (!dataType) continue;
      const headers = hbData[r].slice(1).map(h => String(h)).filter(h => h !== '');
      if (headers.length) tableHeadersMap[dataType] = headers;
    }
  }
  columns.forEach(col => {
    if (col.type.endsWith('-table')) col.tableHeaders = tableHeadersMap[col.type] || [];
  });

  return { columns, rows };
}

/**
 * Fetches Google Drive files by ID server-side and returns them as base64
 * data URLs. Running server-side means the script owner's OAuth token is used,
 * so all editor users can view images regardless of their own Drive session.
 * Files that cannot be accessed (wrong ID, no permission) are returned as null.
 *
 * Called once per image from the client to keep individual response payloads
 * small and avoid HtmlService JSON size limits.
 *
 * @param {string[]} fileIds - Array of Google Drive file IDs to fetch.
 * @returns {Object.<string, string|null>} Map of fileId → data URL (or null on error).
 */
function getImagesDataUrls(fileIds) {
  const result = {};
  fileIds.forEach(fileId => {
    if (!fileId) return;
    try {
      const blob = DriveApp.getFileById(fileId).getBlob();
      result[fileId] = `data:${blob.getContentType() || 'image/jpeg'};base64,${Utilities.base64Encode(blob.getBytes())}`;
    } catch (e) {
      result[fileId] = null;
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Database');
  if (!sheet) throw new Error('Sheet "Database" not found.');
  const newRowIndex = sheet.getLastRow() + 1;
  const numCols = sheet.getLastColumn();
  sheet.getRange(newRowIndex, 1, 1, numCols).setValues([new Array(numCols).fill('')]);
  return newRowIndex;
}

/**
 * Writes new values for a single data row back to the "Database" sheet.
 *
 * @param {number} rowIndex - 1-based spreadsheet row number to update.
 * @param {string[]} values - Array of cell values, one per column.
 * @returns {boolean} Always true; thrown errors propagate to the client failure handler.
 */
function updateRow(rowIndex, values) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Database');
  if (!sheet) throw new Error('Sheet "Database" not found.');
  sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
  return true;
}
