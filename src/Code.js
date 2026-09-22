/**
 * Returns the list of source spreadsheet IDs from Handbook!B2:B.
 * Used by Master Mode to aggregate data from multiple spreadsheets.
 *
 * @returns {string[]}
 */
function getMasterSources() {
  const sheet = getHandbookSheet();
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
 * Reads the Master Mode toggle from Handbook!A2.
 * When true, the webview aggregates data from all source spreadsheets listed in B2:B.
 *
 * @returns {boolean}
 */
function getMasterMode() {
  return readMasterModeFromSheet(getHandbookSheet());
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
 * Adds the "More... ⭐️" custom menu to the Google Sheets toolbar. The photo
 * export and awards import items are only shown when Master Mode is on,
 * since both are meaningless for a spreadsheet with no other sources to
 * aggregate across.
 */
function onOpen() {
  const menu = SpreadsheetApp.getUi()
    .createMenu('More... ⭐️')
    .addItem('Open Web Editor', 'openWebEditor')
    .addSeparator()
    .addItem('Fix phone numbers', 'fixPhoneNumbers')
    .addItem('Fix full names', 'fixFullNames');
  if (getMasterMode()) {
    menu
      .addSeparator()
      .addItem('Export photos for S-КАДР', 'openPhotoExport')
      .addItem('Import awards from S-КАДР', 'importAwards');
  }
  menu.addToUi();
}

/**
 * Runs handbookCheck() and, on failure, alerts the user and returns false
 * without opening any dialog. Shared by openWebEditor() and openPhotoExport()
 * so both entry points fail the same way.
 *
 * @param {string} actionName - Human-readable name of the action being attempted, shown in the alert.
 * @returns {boolean} True if the healthcheck passed.
 */
function assertHandbookHealthy(actionName) {
  const result = handbookCheck();
  if (!result.ok) {
    SpreadsheetApp.getUi().alert(`${actionName} cannot be opened:\n\n${result.errors.join('\n')}`);
  }
  return result.ok;
}

/**
 * Opens the web editor as a full-screen modal dialog.
 * Uses createTemplateFromFile so that <?!= ?> scriptlet includes in
 * WebEditor.html are evaluated before the HTML is served to the client.
 * Bails out (via assertHandbookHealthy()) with an alert instead of opening
 * the modal if the Database/Handbook sheets are missing or corrupted.
 */
function openWebEditor() {
  if (!assertHandbookHealthy('Web Editor')) return;
  const template = HtmlService.createTemplateFromFile('WebEditor');
  template.mode = null;
  SpreadsheetApp.getUi().showModalDialog(template.evaluate(), 'Web Editor');
}

/**
 * Opens the web editor modal directly in "photo export" mode, which skips the
 * normal list-view bootstrap and immediately drives the shared export
 * progress overlay to run startPhotoExport()/copyPhotosBatch(). Only exposed
 * on the menu when Master Mode is on (see onOpen()). Bails out (via
 * assertHandbookHealthy()) with an alert instead of opening the modal if the
 * Database/Handbook sheets are missing or corrupted.
 */
function openPhotoExport() {
  if (!assertHandbookHealthy('Export Photos for S-КАДР')) return;
  const template = HtmlService.createTemplateFromFile('WebEditor');
  template.mode = 'photoExport';
  SpreadsheetApp.getUi().showModalDialog(template.evaluate(), 'Export Photos for S-КАДР');
}

/**
 * Reads one Master Mode source spreadsheet's name and Database rows. Called
 * by the client once per source, after the initial local-only render, so
 * remote rows stream in instead of blocking getSchemaAndData(). Falls back to
 * the raw ID as the name and an empty row list if the spreadsheet is
 * inaccessible or has no Database sheet.
 *
 * @param {string} spreadsheetId
 * @returns {{ id: string, name: string, rows: Array<{rowIndex: number, values: string[], spreadsheetId: string}>, columnMismatches: Array<{colIndex: number, localName: string, localType: string, remoteName: string, remoteType: string}>|null }}
 */
function getMasterSourceRows(spreadsheetId) {
  const remoteSs = openSpreadsheetSafely(spreadsheetId);
  if (!remoteSs) return { id: spreadsheetId, name: spreadsheetId, rows: [], columnMismatches: null };
  try {
    const name = remoteSs.getName();
    const remoteSheet = remoteSs.getSheetByName(SHEET_DATABASE);
    if (!remoteSheet) return { id: spreadsheetId, name, rows: [], columnMismatches: null };
    const remoteAll = remoteSheet.getDataRange().getValues();
    if (remoteAll.length < 2) return { id: spreadsheetId, name, rows: [], columnMismatches: null };
    const { sheet: localSheet } = getDatabaseSheet(null);
    const localSchema = extractColumnSchema(localSheet.getDataRange().getValues());
    const remoteSchema = extractColumnSchema(remoteAll);
    const columnMismatches = compareColumnSchemas(localSchema, remoteSchema);
    if (columnMismatches.length > 0) return { id: spreadsheetId, name, rows: [], columnMismatches };
    const rows = [];
    for (let i = 2; i < remoteAll.length; i++) {
      const values = stringifyRowValues(remoteAll[i]);
      if (values.every(v => v === '')) continue;
      rows.push({ rowIndex: i + 1, values, spreadsheetId });
    }
    return { id: spreadsheetId, name, rows, columnMismatches: null };
  } catch (e) {
    return { id: spreadsheetId, name: spreadsheetId, rows: [], columnMismatches: null };
  }
}

/**
 * Opens a spreadsheet by ID, first probing Drive access so a permission error
 * on an inaccessible spreadsheet can't poison the overall execution status
 * (see docs/architecture-master-mode.md "Actual personnel filter" for background).
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
    // eslint-disable-next-line no-restricted-properties -- this is the safe wrapper itself
    return SpreadsheetApp.openById(id);
  } catch (e) {
    return null;
  }
}

/**
 * Reads the actual personnel name list from the external spreadsheet configured
 * in Handbook A6 (spreadsheet link/ID) and A7 (range address).
 * Returns null if not configured or if the spreadsheet is inaccessible.
 *
 * @returns {string[]|null}
 */
function getActualPersonnelNames() {
  const handbook = getHandbookSheet();
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
 * (by full name, first column) inside the source spreadsheet's own unit
 * folder and move it into the destination spreadsheet's own unit folder —
 * both resolved via getUnitDataFolder(), which looks up a subfolder matching
 * the spreadsheet's own name under the shared Handbook!DATA_FOLDER parent
 * folder. If that unit folder can't be resolved on either side, or the named
 * folder is not found inside it, the row move still succeeds and a note is
 * added to the log.
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
  const destFolderId = getUnitDataFolder(destHandbook, destSs);

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
    const srcFolderId = getUnitDataFolder(srcHandbook, srcSs);

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
      let folderNote;
      if (!srcFolderId || !destFolderId) {
        folderNote = 'Unit data folder not found — folder not moved';
      } else {
        try {
          const srcDataFolder = DriveApp.getFolderById(srcFolderId);
          const personFolder = findFolderByNameCaseInsensitive(srcDataFolder, fullName);
          if (!personFolder) {
            folderNote = 'Folder not found — skipped';
          } else {
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
 *   Column A — single-value config (vertical label/value list; see Config.js constants)
 *   Column B — Master Mode source spreadsheet list
 *   Columns D:F — Export Correspondence table (Template Placeholder | Database Column | Computed Value)
 *   Columns H:J — Table Columns table (Table Type | Column Name | Column Type)
 *   Columns L:AL — Data Types & Allowed Values table (Data Type | Allowed Values...); a type
 *     row with values is a dropdown type, a type row with no values is documentation only
 *
 * @returns {{
 *   columns: Array<{
 *     name: string, type: string,
 *     dropdownOptions?: string[],
 *     tableHeaders?: Array<{name: string, type: string, dropdownOptions?: string[]}>
 *   }>,
 *   rows: Array<{rowIndex: number, values: string[]}>,
 *   masterMode: boolean,
 *   masterSourceIds: string[],
 *   masterSources: Array<{id: string|null, name: string}>|undefined,
 *   actualPersonnelNames: string[]|null,
 *   filterDebounceMs: number,
 *   imageFetchBatchSize: number,
 *   imageFetchConcurrency: number,
 *   masterModeFetchConcurrency: number,
 *   imageCacheTtlDays: number,
 *   exportConfirmThreshold: number,
 *   exportSecondsPerDoc: number,
 *   xlsxExportSecondsPerRow: number,
 *   documentPhotoTabName: string
 * }}
 */
function getSchemaAndData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DATABASE);
  if (!sheet) throw new Error(`Sheet "${SHEET_DATABASE}" not found.`);
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

  const dataTypeOptionsMap = getDataTypeOptionsMap();
  const tableColumnsMap = getTableColumnsMap();
  columns.forEach(col => {
    if (col.type.endsWith('-table')) {
      col.tableHeaders = (tableColumnsMap[col.type] || []).map(sub => {
        const subCol = { name: sub.name, type: sub.type };
        if (dataTypeOptionsMap[sub.type]) subCol.dropdownOptions = dataTypeOptionsMap[sub.type];
        return subCol;
      });
    } else if (dataTypeOptionsMap[col.type]) {
      col.dropdownOptions = dataTypeOptionsMap[col.type];
    }
  });

  const masterSources = masterMode ? [{ id: null, name: ss.getName() }] : undefined;
  return { columns, rows, masterMode, masterSourceIds, masterSources,
           actualPersonnelNames: getActualPersonnelNames(),
           filterDebounceMs: FILTER_DEBOUNCE_MS,
           imageFetchBatchSize: IMAGE_FETCH_BATCH_SIZE, imageFetchConcurrency: IMAGE_FETCH_CONCURRENCY,
           masterModeFetchConcurrency: MASTER_MODE_FETCH_CONCURRENCY,
           imageCacheTtlDays: IMAGE_CACHE_TTL_DAYS,
           exportConfirmThreshold: EXPORT_CONFIRM_THRESHOLD,
           exportSecondsPerDoc: EXPORT_SECONDS_PER_DOC,
           xlsxExportSecondsPerRow: XLSX_EXPORT_SECONDS_PER_ROW,
           documentPhotoTabName: DOCUMENT_PHOTO_TAB_NAME };
}
