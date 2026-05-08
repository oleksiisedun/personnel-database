/**
 * Exports an F-1 document for each given row index by copying the template,
 * replacing {column name} placeholders with row data, and saving to the export
 * folder. Image-type columns have their placeholder replaced with the actual image.
 * Placeholders with no matching column are left untouched.
 *
 * @param {number[]} rowIndices - 1-based sheet row numbers to export.
 * @returns {Array<{name: string, url: string}>} Names and URLs of created documents.
 */
function exportF1(rowIndices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Database');
  if (!sheet) throw new Error('Sheet "Database" not found.');
  const all = sheet.getDataRange().getValues();
  const columns = all[0].map((name, i) => ({
    name: String(name),
    type: String(all[1][i]).toLowerCase()
  }));
  const exportFolder = DriveApp.getFolderById(EXPORT_FOLDER_ID);
  const results = [];

  rowIndices.forEach(rowIndex => {
    const rowValues = all[rowIndex - 1];
    const data = {};
    columns.forEach((col, i) => { data[col.name] = String(rowValues[i] == null ? '' : rowValues[i]); });

    const docName = 'F-1 ' + (data[columns[0].name] || 'Unknown');
    const copy = DriveApp.getFileById(F1_TEMPLATE_ID).makeCopy(docName, exportFolder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    // Replace image-type placeholders with actual images first.
    columns.forEach(col => {
      if (col.type !== 'image') return;
      const placeholder = '{' + col.name + '}';
      const fileId = _extractDriveId(data[col.name] || '');
      if (fileId) {
        try {
          _replacePlaceholderWithImage(body, placeholder, DriveApp.getFileById(fileId).getBlob());
        } catch (e) {
          body.replaceText(_escapeRegex(placeholder), '');
        }
      } else {
        body.replaceText(_escapeRegex(placeholder), '');
      }
    });

    // Replace all remaining text placeholders.
    columns.forEach(col => {
      if (col.type === 'image') return;
      body.replaceText(_escapeRegex('{' + col.name + '}'), _escapeReplacement(data[col.name] || ''));
    });

    doc.saveAndClose();
    results.push({ name: docName, url: copy.getUrl() });
  });

  return results;
}

/**
 * Extracts a Google Drive file ID from a sharing URL or returns the input
 * unchanged if it already looks like a bare file ID.
 * Supported formats:
 *   https://drive.google.com/file/d/FILE_ID/view
 *   https://drive.google.com/open?id=FILE_ID
 *   FILE_ID (bare alphanumeric string ≥ 10 chars)
 *
 * @param {string} url - Drive sharing URL or raw file ID.
 * @returns {string|null} Extracted file ID, or null if not parseable.
 */
function _extractDriveId(url) {
  if (!url) return null;
  var m = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(url.trim())) return url.trim();
  return null;
}

/**
 * Escapes all Java regex metacharacters in a string so it can be passed as a
 * literal pattern to Body.replaceText() or Body.findText().
 *
 * @param {string} str - Raw string to escape.
 * @returns {string} Escaped pattern string.
 */
function _escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Escapes characters that are special in Java regex replacement strings
 * (backslash and dollar sign) so that Body.replaceText() treats the
 * replacement as a plain literal rather than a back-reference expression.
 *
 * @param {string} str - Replacement text to escape.
 * @returns {string} Escaped replacement string.
 */
function _escapeReplacement(str) {
  return str.replace(/\\/g, '\\\\').replace(/\$/g, '\\$');
}

/**
 * Finds the first paragraph in the document body that contains the given
 * placeholder text, clears it, and inserts the provided image blob inline.
 * If the inserted image's height exceeds 600 px the image is scaled down
 * proportionally to fit within that limit.
 *
 * @param {GoogleAppsScript.Document.Body} body - Document body to search.
 * @param {string} placeholder - Literal placeholder string, e.g. "{Фото}".
 * @param {GoogleAppsScript.Base.Blob} blob - Image blob to insert.
 */
function _replacePlaceholderWithImage(body, placeholder, blob) {
  const found = body.findText(_escapeRegex(placeholder));
  if (!found) return;
  const para = found.getElement().getParent().asParagraph();
  para.clear();
  const img = para.appendInlineImage(blob);
  const h = img.getHeight();
  if (h > 600) {
    img.setWidth(Math.round(img.getWidth() * 600 / h));
    img.setHeight(600);
  }
}
