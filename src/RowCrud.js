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
  return deleteRows([{ rowIndex, spreadsheetId }]);
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
