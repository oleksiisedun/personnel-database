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
 * @typedef {{isFolder: boolean, parentIds: string[]}} DriveItemInfo
 */

/**
 * Describes a Drive file/folder for resolvePersonFolderId(). Returns null
 * when the item can't be opened (missing, or no access for the script owner).
 *
 * @param {string} id
 * @returns {DriveItemInfo|null}
 */
function lookupDriveItem(id) {
  try {
    const file = DriveApp.getFileById(id);
    const parentIds = [];
    const parents = file.getParents();
    while (parents.hasNext()) parentIds.push(parents.next().getId());
    return { isFolder: file.getMimeType() === 'application/vnd.google-apps.folder', parentIds };
  } catch (e) {
    return null;
  }
}

/**
 * Resolves a person's Drive folder from the raw cell of the photo column: a
 * folder link is the person's folder itself; a file link is a photo assumed
 * to sit directly in it, so its single parent is. Never guesses — any
 * ambiguity returns a reason for the move log instead of an ID.
 *
 * @param {string} photoCell - Raw photo-column cell (bare ID or Drive URL).
 * @param {(id: string) => (DriveItemInfo|null)} lookup - Usually lookupDriveItem; injected for tests.
 * @returns {{id: string, reason: string}}
 */
function resolvePersonFolderId(photoCell, lookup) {
  const raw = String(photoCell ?? '').trim();
  if (!raw) return { id: '', reason: 'No photo link' };
  const item = lookup(parseDriveId(raw));
  if (!item) return { id: '', reason: 'Photo not accessible' };
  if (item.isFolder) return { id: parseDriveId(raw), reason: '' };
  if (item.parentIds.length !== 1) return { id: '', reason: 'Photo has no single parent folder' };
  return { id: item.parentIds[0], reason: '' };
}

/**
 * Finds a unit's shared data folder from existing people: the parent of the
 * first person folder that resolves from the given photo cells.
 *
 * @param {string[]} photoCells - Raw photo-column cells of existing records.
 * @param {(id: string) => (DriveItemInfo|null)} lookup
 * @returns {string} Folder ID, or '' if no cell resolves to a person folder with a single parent.
 */
function findUnitFolderIdFromPhotos(photoCells, lookup) {
  for (const cell of photoCells) {
    const person = resolvePersonFolderId(cell, lookup);
    if (!person.id) continue;
    const parentIds = (lookup(person.id) ?? { parentIds: [] }).parentIds;
    if (parentIds.length === 1) return parentIds[0];
  }
  return '';
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
    throw new Error(
      `Cannot access the export folder (Handbook!${EXPORT_FOLDER_CELL}, Drive ID "${exportFolderId}"). ` +
        `Check that your Google account has been granted access to this folder. Original error: ${e.message}`,
    );
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

/**
 * Builds the Drive "view" URL for a file or folder.
 *
 * @param {boolean} isFolder - True for a folder ID, false for a file ID.
 * @param {string} id - Drive file or folder ID.
 * @returns {string}
 */
function driveViewUrl(isFolder, id) {
  return isFolder ? 'https://drive.google.com/drive/folders/' + id : 'https://drive.google.com/file/d/' + id + '/view';
}
