/**
 * Returns the list of source spreadsheet IDs from Handbook N2:N.
 * Used by Master Mode to aggregate data from multiple spreadsheets.
 *
 * @returns {string[]}
 */
function getMasterSources() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HANDBOOK);
  if (!sheet) return [];
  try {
    return sheet.getRange(MASTER_MODE_SOURCES_RANGE).getValues()
      .map(r => parseDriveId(String(r[0]).trim()))
      .filter(v => v !== '');
  } catch (e) {
    return [];
  }
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
  try {
    return sheet.getRange(MASTER_MODE_CELL).getValue() === true;
  } catch (e) {
    return false;
  }
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
 * Returns an array of { id, name } for the current spreadsheet (id = null) and
 * every source listed in MASTER_MODE_SOURCES_RANGE. Falls back to the raw ID
 * string if a remote spreadsheet cannot be opened.
 *
 * @returns {Array<{id: string|null, name: string}>}
 */
function getSourceSpreadsheetInfos() {
  const currentSs = SpreadsheetApp.getActiveSpreadsheet();
  const result = [{ id: null, name: currentSs.getName() }];
  getMasterSources().forEach(id => {
    try {
      result.push({ id, name: SpreadsheetApp.openById(id).getName() });
    } catch (e) {
      result.push({ id, name: id });
    }
  });
  return result;
}

/**
 * Extracts a bare Drive resource ID from a raw ID string or any Drive URL form:
 *   https://drive.google.com/drive/folders/<ID>
 *   https://drive.google.com/file/d/<ID>/view
 *   https://drive.google.com/open?id=<ID>
 * Returns the input unchanged when it does not look like a URL.
 *
 * @param {string} value
 * @returns {string}
 */
function parseDriveId(value) {
  const m = value.match(/(?:\/folders\/|\/d\/|[?&]id=)([-\w]+)/);
  return m ? m[1] : value;
}

/**
 * Reads the actual personnel name list from the external spreadsheet configured
 * in Handbook M6 (spreadsheet link/ID) and M7 (range address).
 * Returns null if not configured or if the spreadsheet is inaccessible.
 *
 * @returns {string[]|null}
 */
function getActualPersonnelNames() {
  const handbook = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HANDBOOK);
  if (!handbook) return null;
  const link  = String(handbook.getRange(ACTUAL_PERSONNEL_SPREADSHEET_CELL).getValue()).trim();
  const range = String(handbook.getRange(ACTUAL_PERSONNEL_RANGE_CELL).getValue()).trim();
  const id = parseDriveId(link);
  if (!id || !range) return null;
  try {
    return SpreadsheetApp.openById(id).getRange(range).getValues()
      .map(r => String(r[0]).trim()).filter(v => v !== '');
  } catch (e) {
    return null;
  }
}


/**
 * Moves rows from their source spreadsheets into the destination spreadsheet.
 * Rows are appended to the destination Database sheet and hard-deleted from
 * the source. Within each source the rows are processed in descending rowIndex
 * order so that deleting one row does not shift the indices of others.
 *
 * After moving each row the function tries to find the person's Drive folder
 * (by full name, first column) in the source DATA_FOLDER and move it to the
 * destination DATA_FOLDER. If the folder is not found or DATA_FOLDER is not
 * configured the row move still succeeds and a note is added to the log.
 *
 * @param {Array<{rowIndex: number, spreadsheetId: string|null}>} rowEntries
 * @param {string|null} destinationSpreadsheetId
 * @returns {{ log: Array<{name: string, folderNote: string}> }}
 */
function movePersonnel(rowEntries, destinationSpreadsheetId) {
  const destSs = destinationSpreadsheetId
    ? SpreadsheetApp.openById(destinationSpreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSs.getSheetByName(SHEET_DATABASE);
  if (!destSheet) throw new Error('Destination sheet "Database" not found.');

  const destHandbook = destSs.getSheetByName(SHEET_HANDBOOK);
  const destFolderId = parseDriveId(destHandbook
    ? String(destHandbook.getRange(DATA_FOLDER).getValue()).trim()
    : '');

  // Group entries by source spreadsheetId, sort each group descending by rowIndex.
  const groups = new Map();
  rowEntries.forEach(entry => {
    const key = entry.spreadsheetId ?? null;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(entry);
  });
  groups.forEach(entries => entries.sort((a, b) => b.rowIndex - a.rowIndex));

  const log = [];
  const movedRows = [];

  groups.forEach((entries, spreadsheetId) => {
    const srcSs = spreadsheetId
      ? SpreadsheetApp.openById(spreadsheetId)
      : SpreadsheetApp.getActiveSpreadsheet();
    const srcSheet = srcSs.getSheetByName(SHEET_DATABASE);
    if (!srcSheet) {
      entries.forEach(() => log.push({ name: '?', folderNote: 'Source sheet not found — skipped' }));
      return;
    }

    const srcHandbook = srcSs.getSheetByName(SHEET_HANDBOOK);
    const srcFolderId = parseDriveId(srcHandbook
      ? String(srcHandbook.getRange(DATA_FOLDER).getValue()).trim()
      : '');

    const destNumCols = destSheet.getLastColumn();

    entries.forEach(({ rowIndex }) => {
      const srcNumCols = srcSheet.getLastColumn();
      const rowData = srcSheet.getRange(rowIndex, 1, 1, srcNumCols).getValues()[0];
      const fullName = String(rowData[0]).trim();

      if (spreadsheetId === destinationSpreadsheetId) {
        log.push({ name: fullName, folderNote: 'Same spreadsheet — skipped' });
        return;
      }

      // Pad or trim row to match destination column count.
      const paddedData = rowData.slice(0, destNumCols);
      while (paddedData.length < destNumCols) paddedData.push('');

      const newRowIndex = Math.max(destSheet.getLastRow(), 2) + 1;
      destSheet.getRange(newRowIndex, 1, 1, destNumCols).setValues([paddedData]);

      srcSheet.deleteRow(rowIndex);

      const newRow = { rowIndex: newRowIndex, values: paddedData.map(c => String(c == null ? '' : c)) };
      if (destinationSpreadsheetId) newRow.spreadsheetId = destinationSpreadsheetId;
      movedRows.push(newRow);

      // Try to move the person's Drive folder.
      let folderNote = '';
      if (!srcFolderId || !destFolderId) {
        folderNote = 'DATA_FOLDER not configured — folder not moved';
      } else {
        try {
          const srcDataFolder = DriveApp.getFolderById(srcFolderId);
          const folderIter = srcDataFolder.getFoldersByName(fullName);
          if (!folderIter.hasNext()) {
            folderNote = 'Folder not found — skipped';
          } else {
            const personFolder = folderIter.next();
            const destDataFolder = DriveApp.getFolderById(destFolderId);
            personFolder.moveTo(destDataFolder);
            folderNote = 'Folder moved';
          }
        } catch (e) {
          folderNote = `Folder move failed: ${e.message}`;
        }
      }

      log.push({ name: fullName, folderNote });
    });
  });

  return { log, movedRows };
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
 *   rows: Array<{rowIndex: number, values: string[]}>,
 *   masterMode: boolean,
 *   masterSources: Array<{id: string|null, name: string}>|undefined,
 *   filterDebounceMs: number,
 *   imageFetchBatchSize: number,
 *   imageFetchConcurrency: number,
 *   imageCacheTtlDays: number
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
    try {
      const hbData = handbookSheet.getRange(HANDBOOK_TYPES_RANGE).getValues();
      for (let r = 0; r < hbData.length; r++) {
        const dataType = String(hbData[r][0]).trim().toLowerCase();
        if (!dataType) continue;
        const headers = hbData[r].slice(1).map(h => String(h)).filter(h => h !== '');
        if (headers.length) tableHeadersMap[dataType] = headers;
      }
    } catch (e) {
      // skip if Handbook table-types range is inaccessible
    }
  }
  columns.forEach(col => {
    if (col.type.endsWith('-table')) col.tableHeaders = tableHeadersMap[col.type] || [];
  });

  const DROPDOWN_TYPES = [
    { type: 'unit',           range: HANDBOOK_UNIT_RANGE,           key: 'unitOptions' },
    { type: 'origin',         range: HANDBOOK_ORIGIN_RANGE,         key: 'originOptions' },
    { type: 'marital-status', range: HANDBOOK_MARITAL_STATUS_RANGE, key: 'maritalStatusOptions' },
    { type: 'sex',            range: HANDBOOK_SEX_RANGE,            key: 'sexOptions' },
  ];
  DROPDOWN_TYPES.forEach(({ type, range, key }) => {
    let options = [];
    if (handbookSheet) {
      try {
        options = handbookSheet.getRange(range).getValues().map(r => String(r[0])).filter(v => v !== '');
      } catch (e) {
        // skip if this Handbook range is inaccessible
      }
    }
    columns.forEach(col => { if (col.type === type) col[key] = options; });
  });

  const masterSources = masterMode ? getSourceSpreadsheetInfos() : undefined;
  return { columns, rows, masterMode, masterSources,
           actualPersonnelNames: getActualPersonnelNames(),
           filterDebounceMs: FILTER_DEBOUNCE_MS,
           imageFetchBatchSize: IMAGE_FETCH_BATCH_SIZE, imageFetchConcurrency: IMAGE_FETCH_CONCURRENCY,
           imageCacheTtlDays: IMAGE_CACHE_TTL_DAYS };
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
 * Appends a new row populated with the given values to the "Database" sheet
 * and returns its 1-based row index. Always writes to the active (local) spreadsheet.
 *
 * @param {string[]} values - Array of cell values, one per column.
 * @returns {number} 1-based row index of the newly created row.
 */
function addRowWithData(values) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATABASE);
  if (!sheet) throw new Error('Sheet "Database" not found.');
  const newRowIndex = sheet.getLastRow() + 1;
  const numCols = sheet.getLastColumn();
  const padded = values.slice(0, numCols);
  while (padded.length < numCols) padded.push('');
  sheet.getRange(newRowIndex, 1, 1, numCols).setValues([padded]);
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
