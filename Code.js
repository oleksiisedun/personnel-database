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
    .addItem('Fix phone numbers', 'fixPhoneNumbers')
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
 * Normalizes phone numbers in the Database sheet's "Номер телефону" column:
 * adds a leading zero to bare 9-digit numbers, and strips the "38" country
 * prefix from 12-digit numbers. Triggered only from the custom menu.
 * @returns {void}
 */
function fixPhoneNumbers() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATABASE);
  if (!sheet) {
    ui.alert(`Sheet "${SHEET_DATABASE}" not found.`);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.findIndex(h => String(h).trim() === COL_PHONE_NUMBER);
  if (colIndex === -1) {
    ui.alert(`Column "${COL_PHONE_NUMBER}" not found in row 1.`);
    return;
  }

  const numRows = sheet.getLastRow() - 2;
  if (numRows <= 0) {
    ui.alert('No data rows to process.');
    return;
  }

  const range = sheet.getRange(3, colIndex + 1, numRows, 1);
  const values = range.getValues();
  let fixedCount = 0;

  const result = values.map(row => {
    const phone = String(row[0]).trim();
    if (PHONE_REGEX_9DIGIT.test(phone) && phone[0] !== '0') {
      fixedCount++;
      return ['0' + phone];
    }
    if (PHONE_REGEX_COUNTRY.test(phone)) {
      fixedCount++;
      return [phone.slice(2)];
    }
    return row;
  });

  range.setValues(result);
  ui.alert(`Fixed ${fixedCount} phone number(s).`);
}

/**
 * Reads one Master Mode source spreadsheet's name and Database rows. Called
 * by the client once per source, after the initial local-only render, so
 * remote rows stream in instead of blocking getSchemaAndData(). Falls back to
 * the raw ID as the name and an empty row list if the spreadsheet is
 * inaccessible or has no Database sheet.
 *
 * @param {string} spreadsheetId
 * @returns {{ id: string, name: string, rows: Array<{rowIndex: number, values: string[], spreadsheetId: string}> }}
 */
function getMasterSourceRows(spreadsheetId) {
  const remoteSs = openSpreadsheetSafely(spreadsheetId);
  if (!remoteSs) return { id: spreadsheetId, name: spreadsheetId, rows: [] };
  try {
    const name = remoteSs.getName();
    const remoteSheet = remoteSs.getSheetByName(SHEET_DATABASE);
    if (!remoteSheet) return { id: spreadsheetId, name, rows: [] };
    const remoteAll = remoteSheet.getDataRange().getValues();
    const rows = [];
    for (let i = 2; i < remoteAll.length; i++) {
      const values = stringifyRowValues(remoteAll[i]);
      if (values.every(v => v === '')) continue;
      rows.push({ rowIndex: i + 1, values, spreadsheetId });
    }
    return { id: spreadsheetId, name, rows };
  } catch (e) {
    return { id: spreadsheetId, name: spreadsheetId, rows: [] };
  }
}

/**
 * Extracts a bare Drive resource ID from a raw ID string or any Drive URL form:
 *   https://drive.google.com/drive/folders/<ID>
 *   https://drive.google.com/file/d/<ID>/view
 *   https://drive.google.com/open?id=<ID>
 * Returns the input unchanged when it does not look like a URL.
 * Client-side counterpart: extractDriveId() in WebEditor.js.html — keep URL patterns aligned.
 *
 * @param {string} value
 * @returns {string}
 */
function parseDriveId(value) {
  const m = value.match(/(?:\/folders\/|\/d\/|[?&]id=)([-\w]+)/);
  return m ? m[1] : value;
}

/**
 * Opens a spreadsheet by ID, first probing Drive access so a permission error
 * on an inaccessible spreadsheet can't poison the overall execution status
 * (see CLAUDE.md "Actual personnel filter" for background).
 *
 * @param {string} id - Spreadsheet (Drive file) ID.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet|null} The opened spreadsheet, or null if inaccessible.
 */
function openSpreadsheetSafely(id) {
  try {
    DriveApp.getFileById(id);
  } catch (e) {
    return null;
  }
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    return null;
  }
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
  const ss = openSpreadsheetSafely(id);
  if (!ss) return null;
  try {
    return ss.getRange(range).getValues()
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
 * @returns {{
 *   log: Array<{name: string, folderNote: string}>,
 *   movedRows: Array<{rowIndex: number, values: string[], spreadsheetId?: string}>,
 *   skippedEntries: Array<{rowIndex: number, spreadsheetId: string|null}>
 * }}
 */
function movePersonnel(rowEntries, destinationSpreadsheetId) {
  const { ss: destSs, sheet: destSheet } = getDatabaseSheet(destinationSpreadsheetId);

  const destHandbook = destSs.getSheetByName(SHEET_HANDBOOK);
  const destFolderId = parseDriveId(destHandbook
    ? String(destHandbook.getRange(DATA_FOLDER).getValue()).trim()
    : '');

  const groups = groupAndSortBySpreadsheetId(rowEntries);

  const log = [];
  const movedRows = [];
  const skippedEntries = [];

  groups.forEach((entries, spreadsheetId) => {
    const srcSs = resolveSpreadsheet(spreadsheetId);
    if (!srcSs) {
      entries.forEach(({ rowIndex }) => {
        log.push({ name: '?', folderNote: 'Source spreadsheet not accessible — skipped' });
        skippedEntries.push({ rowIndex, spreadsheetId });
      });
      return;
    }
    const srcSheet = srcSs.getSheetByName(SHEET_DATABASE);
    if (!srcSheet) {
      entries.forEach(({ rowIndex }) => {
        log.push({ name: '?', folderNote: 'Source sheet not found — skipped' });
        skippedEntries.push({ rowIndex, spreadsheetId });
      });
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
        skippedEntries.push({ rowIndex, spreadsheetId });
        return;
      }

      const paddedData = padRowToColumnCount(rowData, destNumCols);

      const newRowIndex = Math.max(destSheet.getLastRow(), 2) + 1;
      destSheet.getRange(newRowIndex, 1, 1, destNumCols).setValues([paddedData]);

      srcSheet.deleteRow(rowIndex);

      const newRow = { rowIndex: newRowIndex, values: stringifyRowValues(paddedData) };
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

  return { log, movedRows, skippedEntries };
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
 *   masterSourceIds: string[],
 *   masterSources: Array<{id: string|null, name: string}>|undefined,
 *   filterDebounceMs: number,
 *   imageFetchBatchSize: number,
 *   imageFetchConcurrency: number,
 *   masterModeFetchConcurrency: number,
 *   imageCacheTtlDays: number,
 *   exportConfirmThreshold: number,
 *   exportSecondsPerDoc: number
 * }}
 */
function getSchemaAndData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DATABASE);
  if (!sheet) throw new Error('Sheet "Database" not found.');
  const all = sheet.getDataRange().getValues();
  if (all.length < 2) throw new Error('Sheet must have at least 2 rows (names + types).');
  const columns = extractColumnSchema(all);
  const rows = [];
  for (let i = 2; i < all.length; i++) {
    const values = stringifyRowValues(all[i]);
    if (values.every(v => v === '')) continue;
    rows.push({ rowIndex: i + 1, values });
  }

  const masterMode = getMasterMode();
  const masterSourceIds = masterMode ? getMasterSources() : [];

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

  const masterSources = masterMode ? [{ id: null, name: ss.getName() }] : undefined;
  return { columns, rows, masterMode, masterSourceIds, masterSources,
           actualPersonnelNames: getActualPersonnelNames(),
           filterDebounceMs: FILTER_DEBOUNCE_MS,
           imageFetchBatchSize: IMAGE_FETCH_BATCH_SIZE, imageFetchConcurrency: IMAGE_FETCH_CONCURRENCY,
           masterModeFetchConcurrency: MASTER_MODE_FETCH_CONCURRENCY,
           imageCacheTtlDays: IMAGE_CACHE_TTL_DAYS,
           exportConfirmThreshold: EXPORT_CONFIRM_THRESHOLD,
           exportSecondsPerDoc: EXPORT_SECONDS_PER_DOC };
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
 * and returns its 1-based row index. When spreadsheetId is provided, writes to
 * that remote spreadsheet via openSpreadsheetSafely(); otherwise writes locally.
 *
 * @param {string[]} values - Array of cell values, one per column.
 * @param {string|null} [spreadsheetId] - Remote spreadsheet ID, or null/omitted for local.
 * @returns {number} 1-based row index of the newly created row.
 */
function addRowWithData(values, spreadsheetId) {
  const { sheet } = getDatabaseSheet(spreadsheetId);
  const newRowIndex = sheet.getLastRow() + 1;
  const numCols = sheet.getLastColumn();
  const padded = padRowToColumnCount(values, numCols);
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
  const { sheet } = getDatabaseSheet(spreadsheetId);
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
  const { ss, sheet: dbSheet } = getDatabaseSheet(spreadsheetId);

  const numCols = dbSheet.getLastColumn();
  const rowData = dbSheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];

  const trashSheet = ensureTrashSheetExists(ss, dbSheet, numCols);
  const trashLastRow = Math.max(trashSheet.getLastRow(), 2);
  trashSheet.getRange(trashLastRow + 1, 1, 1, numCols).setValues([rowData]);

  dbSheet.deleteRow(rowIndex);
  return true;
}

/**
 * Moves multiple rows to the Trash sheet (soft delete), processing each spreadsheet's rows
 * in descending rowIndex order to avoid row-shift bugs during sequential deletion.
 * @param {Array<{rowIndex: number, spreadsheetId: string|null}>} rowEntries
 * @returns {boolean} Always true; thrown errors propagate to the client failure handler.
 */
function deleteRows(rowEntries) {
  const groups = groupAndSortBySpreadsheetId(rowEntries);
  for (const [spreadsheetId, entries] of groups) {
    const { ss, sheet: dbSheet } = getDatabaseSheet(spreadsheetId);

    const numCols = dbSheet.getLastColumn();
    const trashSheet = ensureTrashSheetExists(ss, dbSheet, numCols);
    for (const { rowIndex } of entries) {
      const rowData = dbSheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];
      const trashLastRow = Math.max(trashSheet.getLastRow(), 2);
      trashSheet.getRange(trashLastRow + 1, 1, 1, numCols).setValues([rowData]);
      dbSheet.deleteRow(rowIndex);
    }
  }
  return true;
}
