/**
 * Imports award records from an external S-КАДР Google Sheet into the local
 * "Нагороди" *-table column, matching people by "S-КАДР ID" across the local
 * Database sheet and every Master Mode source. Triggered only from the
 * custom menu, and only shown there when Master Mode is on (see onOpen()).
 * @returns {void}
 */
function importAwards() {
  const ui = SpreadsheetApp.getUi();
  if (!getMasterMode()) {
    ui.alert('Master Mode is not enabled.');
    return;
  }

  const resp = ui.prompt('Import awards from S-КАДР', 'Paste the full Google Sheets URL:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const url = resp.getResponseText().trim();
  if (!url) return;

  const spreadsheetId = parseDriveId(url);
  const gid = parseGidFromUrl(url);
  const importSs = openSpreadsheetSafely(spreadsheetId);
  if (!importSs) {
    ui.alert('Could not open the import spreadsheet. Check the URL and your access.');
    return;
  }

  const importSheet = gid !== null
    ? importSs.getSheets().find(s => s.getSheetId() === gid)
    : importSs.getSheets()[0];
  if (!importSheet) {
    ui.alert('Could not find the specified sheet tab.');
    return;
  }

  const importData = importSheet.getDataRange().getValues();
  const awardsByPersonId = _groupAwardsById(importData);
  if (Object.keys(awardsByPersonId).length === 0) {
    ui.alert('No award rows found to import.');
    return;
  }

  const sources = _collectAwardTargetSources();
  const result = _applyAwards(awardsByPersonId, sources);

  Logger.log(result.fullLog.join('\n'));
  ui.alert(_buildImportSummary(result));
}

/**
 * Groups the external import sheet's rows into award entries keyed by
 * personnel ID, skipping rows with no ID or no award name. Cleans the order
 * number and capitalizes the award name's first letter along the way.
 *
 * @param {Array<Array<*>>} importData - Raw values from the import sheet (row 1 = header).
 * @returns {Object.<string, string[][]>} Map of ID -> array of [name, order number, order date].
 */
function _groupAwardsById(importData) {
  const idColIndex = columnLetterToIndex(AWARDS_IMPORT_ID_COL);
  const nameColIndex = columnLetterToIndex(AWARDS_IMPORT_NAME_COL);
  const orderNumberColIndex = columnLetterToIndex(AWARDS_IMPORT_ORDER_NUMBER_COL);
  const orderDateColIndex = columnLetterToIndex(AWARDS_IMPORT_ORDER_DATE_COL);

  const map = {};
  for (let i = AWARDS_IMPORT_DATA_START_ROW - 1; i < importData.length; i++) {
    const row = importData[i];
    const id = String(row[idColIndex] || '').trim();
    const name = String(row[nameColIndex] || '').trim();
    if (!id || !name) continue;

    const capitalizedName = name.charAt(0).toUpperCase() + name.slice(1);
    const orderNumber = String(row[orderNumberColIndex] || '').replace(AWARDS_ORDER_NUMBER_CLEAN_REGEX, '');
    // Date-formatted cells come back from getValues() as real Date objects,
    // not strings — String(date) would yield "Wed Dec 31 2025 10:00:00
    // GMT+0200 (...)" instead of the DD.MM.YYYY format stored everywhere else.
    const orderDateRaw = row[orderDateColIndex];
    const orderDate = orderDateRaw instanceof Date
      ? formatDateDDMMYYYY(orderDateRaw)
      : String(orderDateRaw || '').trim();

    if (!map[id]) map[id] = [];
    map[id].push([capitalizedName, orderNumber, orderDate]);
  }
  return map;
}

/**
 * Builds the list of Database sheets to search for matching personnel IDs:
 * the local spreadsheet plus every accessible Master Mode source.
 *
 * @returns {Array<{spreadsheetId: string|null, sheet: GoogleAppsScript.Spreadsheet.Sheet, values: Array<Array<*>>, idColIndex: number, awardsColIndex: number, rowIndexById: Object.<string, number>}>}
 */
function _collectAwardTargetSources() {
  const sources = [];
  _addAwardSource(sources, null, SpreadsheetApp.getActiveSpreadsheet());
  getMasterSources().forEach(id => {
    const remoteSs = openSpreadsheetSafely(id);
    if (remoteSs) _addAwardSource(sources, id, remoteSs);
  });
  return sources;
}

/**
 * Reads one spreadsheet's Database sheet and, if it has both an "S-КАДР ID"
 * and a "Нагороди" column, appends it to the sources list with a precomputed
 * ID -> row-index map. Sources missing either column, or with no Database
 * sheet, are silently skipped.
 *
 * @param {Array<Object>} sources - Accumulator array, mutated in place.
 * @param {string|null} spreadsheetId
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {void}
 */
function _addAwardSource(sources, spreadsheetId, ss) {
  const sheet = ss.getSheetByName(SHEET_DATABASE);
  if (!sheet) return;

  const values = sheet.getDataRange().getValues();
  const headerRow = values[0] || [];
  const idColIndex = findColumnIndex(headerRow, COL_CARD_ID);
  const awardsColIndex = findColumnIndex(headerRow, COL_AWARDS);
  if (idColIndex === -1 || awardsColIndex === -1) return;

  const rowIndexById = {};
  for (let i = 2; i < values.length; i++) {
    const id = String(values[i][idColIndex] || '').trim();
    if (id) rowIndexById[id] = i;
  }

  sources.push({ spreadsheetId, sheet, values, idColIndex, awardsColIndex, rowIndexById });
}

/**
 * Matches each imported ID against the collected sources, merges its award
 * entries into the matched person's "Нагороди" cell (skipping entries that
 * exactly duplicate an existing one), and writes the result back to the
 * sheet. An ID present in more than one source resolves to the first source
 * that has it (local first, then Master Mode sources in order).
 *
 * @param {Object.<string, string[][]>} awardsByPersonId - Map of ID -> array of [name, order number, order date].
 * @param {Array<Object>} sources - Sources built by _collectAwardTargetSources().
 * @returns {{peopleUpdated: number, entriesAdded: number, entriesDuplicate: number, notFound: string[], fullLog: string[]}}
 */
function _applyAwards(awardsByPersonId, sources) {
  let peopleUpdated = 0;
  let entriesAdded = 0;
  let entriesDuplicate = 0;
  const notFound = [];
  const fullLog = [];

  Object.keys(awardsByPersonId).forEach(id => {
    const newEntries = awardsByPersonId[id];
    const source = sources.find(s => Object.prototype.hasOwnProperty.call(s.rowIndexById, id));
    if (!source) {
      notFound.push(id);
      fullLog.push(`${id}: not found in any source, skipped`);
      return;
    }

    const rowIndex = source.rowIndexById[id];
    const existingRows = _parseSubTable(source.values[rowIndex][source.awardsColIndex]);
    const existingKeys = new Set(existingRows.map(r => r.join('')));

    let addedForPerson = 0;
    newEntries.forEach(entry => {
      const key = entry.join('');
      if (existingKeys.has(key)) {
        entriesDuplicate++;
        return;
      }
      existingRows.push(entry);
      existingKeys.add(key);
      addedForPerson++;
      entriesAdded++;
    });

    if (addedForPerson > 0) {
      source.sheet.getRange(rowIndex + 1, source.awardsColIndex + 1).setValue(_encodeSubTable(existingRows));
      peopleUpdated++;
      fullLog.push(`${id}: added ${addedForPerson} award(s)`);
    } else {
      fullLog.push(`${id}: no new awards (all duplicates)`);
    }
  });

  return { peopleUpdated, entriesAdded, entriesDuplicate, notFound, fullLog };
}

/**
 * Formats a capped, human-readable summary of an award import for display in
 * a ui.alert() dialog. The full per-ID log is expected to already have been
 * written to the execution log separately (see importAwards()).
 *
 * @param {{peopleUpdated: number, entriesAdded: number, entriesDuplicate: number, notFound: string[]}} result
 * @returns {string}
 */
function _buildImportSummary(result) {
  const lines = [
    `People updated: ${result.peopleUpdated}`,
    `Award entries added: ${result.entriesAdded}`,
    `Duplicate entries skipped: ${result.entriesDuplicate}`,
    `IDs not found: ${result.notFound.length}`,
  ];
  if (result.notFound.length > 0) {
    const shown = result.notFound.slice(0, AWARDS_IMPORT_NOT_FOUND_DISPLAY_LIMIT);
    const suffix = result.notFound.length > shown.length ? ` (+${result.notFound.length - shown.length} more)` : '';
    lines.push('', 'Not found: ' + shown.join(', ') + suffix);
  }
  lines.push('', 'Full details logged to the Apps Script execution log.');
  return lines.join('\n');
}
