/**
 * Classifies an already-opened Drive file by mimetype, returning its export
 * type and Drive "view" URL. Shared by getImagesDataUrls() (which additionally
 * fetches + base64-encodes the blob for the 'image' case) and
 * resolveDriveFileForExport() (XLSX export's link-only resolver, which never
 * fetches the blob).
 *
 * @param {GoogleAppsScript.Drive.File} file - Already-opened Drive file.
 * @param {string} fileId
 * @returns {{type: 'image'|'pdf'|'folder', viewUrl: string, mimeType: string}}
 */
function _classifyDriveFile(file, fileId) {
  const mimeType = file.getMimeType();
  if (mimeType === 'application/pdf') {
    return { type: 'pdf', viewUrl: 'https://drive.google.com/file/d/' + fileId + '/view', mimeType };
  }
  if (mimeType === 'application/vnd.google-apps.folder') {
    return { type: 'folder', viewUrl: 'https://drive.google.com/drive/folders/' + fileId, mimeType };
  }
  return { type: 'image', viewUrl: 'https://drive.google.com/file/d/' + fileId + '/view', mimeType };
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
  /** @type {Object.<string, {type: string, dataUrl?: string, viewUrl?: string}>} */
  const result = {};
  fileIds.forEach(fileId => {
    if (!fileId) return;
    try {
      const file = DriveApp.getFileById(fileId);
      const info = _classifyDriveFile(file, fileId);
      if (info.type === 'image') {
        const blob = file.getBlob();
        result[fileId] = { type: 'image', dataUrl: 'data:' + (info.mimeType || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
      } else {
        result[fileId] = { type: info.type, viewUrl: info.viewUrl };
      }
    } catch (e) {
      result[fileId] = { type: 'no-access' };
    }
  });
  return result;
}

/**
 * Resolves a Drive file/folder's export link type without fetching its blob —
 * the lighter-weight counterpart to getImagesDataUrls(), used only by XLSX
 * export (Export.js), which needs a clickable Drive view URL but never embeds
 * image bytes.
 *
 * @param {string} fileId
 * @returns {{type: 'image'|'pdf'|'folder', viewUrl: string}|{type: 'no-access'}}
 */
function resolveDriveFileForExport(fileId) {
  try {
    return _classifyDriveFile(DriveApp.getFileById(fileId), fileId);
  } catch (e) {
    return { type: 'no-access' };
  }
}
