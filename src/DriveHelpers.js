/**
 * Extracts a bare Drive resource ID from a raw ID string or any Drive URL form:
 *   https://drive.google.com/drive/folders/<ID>
 *   https://drive.google.com/file/d/<ID>/view
 *   https://drive.google.com/open?id=<ID>
 * Returns the input unchanged when it does not look like a URL.
 * Client-side counterpart: extractDriveId() in WebEditor.images.js.html — keep URL patterns aligned.
 *
 * @param {string} value
 * @returns {string}
 */
function parseDriveId(value) {
  const m = value.match(DRIVE_URL_REGEX);
  return m ? m[1] : value;
}

/**
 * Tests whether a raw cell value looks like a Drive sharing URL (as opposed to
 * a bare ID or ordinary text). Used by XLSX export to decide whether a
 * non-image column's cell should be linkified — image columns always attempt
 * resolution regardless (see _buildXlsxLinkCell() in Export.js).
 *
 * @param {string} value
 * @returns {boolean}
 */
function looksLikeDriveUrl(value) {
  return DRIVE_URL_REGEX.test(value);
}

/**
 * Extracts the unit name portion of a unit spreadsheet's own file name.
 * Spreadsheet names follow a "<fixed prefix>UNIT_NAME_SEPARATOR<unit name>"
 * convention, e.g. "УСТАНОВЧІ ДАНІ О/С - 7 РОП" → "7 РОП", which is what the
 * unit's actual Drive folder (a direct subfolder of the shared UNITS parent)
 * is named — matching the full spreadsheet name against the folder name
 * would never succeed. Splits on the *last* UNIT_NAME_SEPARATOR occurrence
 * and trims; falls back to the full name unchanged if the separator isn't
 * present, so an unexpected spreadsheet name degrades to the old
 * whole-name-match behavior instead of throwing.
 *
 * @param {string} spreadsheetName
 * @returns {string}
 */
function extractUnitName(spreadsheetName) {
  const idx = spreadsheetName.lastIndexOf(UNIT_NAME_SEPARATOR);
  return idx === -1 ? spreadsheetName.trim() : spreadsheetName.slice(idx + UNIT_NAME_SEPARATOR.length).trim();
}

/**
 * Finds a direct subfolder of parentFolder matching name, case-insensitively.
 * DriveApp's Folder.getFoldersByName() only does exact, case-sensitive matches.
 * @param {GoogleAppsScript.Drive.Folder} parentFolder
 * @param {string} name
 * @returns {GoogleAppsScript.Drive.Folder|null}
 */
function findFolderByNameCaseInsensitive(parentFolder, name) {
  const target = name.toLowerCase();
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName().toLowerCase() === target) return folder;
  }
  return null;
}

/**
 * Resolves a spreadsheet's own unit-specific person-photo folder.
 * Only one spreadsheet is expected to run in Master Mode — the central one
 * that aggregates every other unit — so Handbook!MASTER_MODE_CELL doubles as
 * the signal for "this spreadsheet's own folder isn't organized under the
 * shared UNITS tree": when it's true, Handbook!DATA_FOLDER is returned
 * directly as this spreadsheet's own dedicated folder (e.g. the master
 * spreadsheet's separate headquarters personnel folder). When Master Mode is
 * off, DATA_FOLDER is treated as the shared parent "UNITS" folder — the same
 * ID across every unit spreadsheet, since it's imported via IMPORTRANGE like
 * most other Handbook config — and searched for a direct subfolder matching
 * extractUnitName(ss.getName()), case-insensitively (via
 * findFolderByNameCaseInsensitive()). Returns '' if DATA_FOLDER is
 * unconfigured, the folder is inaccessible, or (Master Mode off) no subfolder
 * matches the extracted unit name.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|null} handbookSheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {string}
 */
function getUnitDataFolder(handbookSheet, ss) {
  const folderId = getDriveIdFromHandbook(handbookSheet, DATA_FOLDER);
  if (!folderId) return '';

  if (readMasterModeFromSheet(handbookSheet)) return folderId;

  try {
    const parentFolder = DriveApp.getFolderById(folderId);
    const unitFolder = findFolderByNameCaseInsensitive(parentFolder, extractUnitName(ss.getName()));
    return unitFolder ? unitFolder.getId() : '';
  } catch (e) {
    return '';
  }
}

/**
 * Opens the export folder (Handbook!A13 / EXPORT_FOLDER_CELL) by ID, throwing
 * a clear, actionable error instead of DriveApp's raw "Unexpected error while
 * getting the method or property getFolderById on object DriveApp" — either
 * because the cell is empty, or because the current user's account lacks
 * access to the configured folder.
 *
 * @param {string} exportFolderId
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getExportFolderSafely(exportFolderId) {
  if (!exportFolderId) {
    throw new Error(`Export folder is not configured (Handbook!${EXPORT_FOLDER_CELL} is empty).`);
  }
  try {
    return DriveApp.getFolderById(exportFolderId);
  } catch (e) {
    throw new Error(`Cannot access the export folder (Handbook!${EXPORT_FOLDER_CELL}, Drive ID "${exportFolderId}"). ` +
      `Check that your Google account has been granted access to this folder. Original error: ${e.message}`);
  }
}

/**
 * Extracts the gid (tab id) from a Google Sheets URL's ?gid=... / #gid=... param.
 *
 * @param {string} url
 * @returns {number|null} The parsed gid, or null if the URL has none.
 */
function parseGidFromUrl(url) {
  const m = GID_REGEX.exec(String(url));
  return m ? Number(m[1]) : null;
}
