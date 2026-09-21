/**
 * Discovers every Actual Personnel person (local Database + every Master Mode
 * source in Handbook!B2:B, processed in that fixed order so duplicate S-КАДР
 * ID resolution is reproducible run to run) who has both a non-empty "Фото"
 * and "S-КАДР ID" value, creates the destination Drive folder, and returns
 * the full eligible list plus a skip list. No image bytes are fetched here —
 * only sheet values — so the complete eligible list (and therefore duplicate
 * detection) is known before any photo copying starts. Unlike _exportDoc(),
 * this has no continuation: a half-scanned personnel list has no standalone
 * value, so it throws instead of returning partial results if the scan
 * exceeds EXPORT_TIME_LIMIT_MS (same choice exportXLSX() makes).
 *
 * @returns {{
 *   folderId: string,
 *   folderUrl: string,
 *   eligible: Array<{rowIndex: number, spreadsheetId: string|null, fullName: string, cardId: string, fileId: string}>,
 *   skipped: Array<{name: string, reason: string}>
 * }}
 */
function startPhotoExport() {
  const localSs = SpreadsheetApp.getActiveSpreadsheet();
  const handbookSheet = localSs.getSheetByName(SHEET_HANDBOOK);

  const exportFolderId = getDriveIdFromHandbook(handbookSheet, EXPORT_FOLDER_CELL);
  const parentFolder = getExportFolderSafely(exportFolderId);
  const destFolder = parentFolder.createFolder(PHOTO_EXPORT_FOLDER_PREFIX + formatDateDDMMYYYY(new Date()));

  const actualNames = getActualPersonnelNames();
  if (!actualNames) {
    throw new Error('Actual Personnel list (Handbook!A6/A7) is not configured or not accessible — cannot determine who to export.');
  }
  const actualSet = new Set(actualNames.map(n => n.trim().toUpperCase()));

  const sourceIds = [null, ...getMasterSources()]; // null = local Database; fixed, deterministic order
  const getSheetData = _makeSheetDataLoader(localSs);

  const seenCardIds = new Map(); // cardId -> fullName of first winner
  const eligible = [];
  const skipped = [];
  const startTime = Date.now();

  sourceIds.forEach((spreadsheetId, idx) => {
    if (Date.now() - startTime > EXPORT_TIME_LIMIT_MS) {
      throw new Error(`Timed out scanning personnel (${idx} of ${sourceIds.length} sources completed).`);
    }
    const sheetData = getSheetData(spreadsheetId);
    if (!sheetData) return; // inaccessible/missing source — silently skipped, same convention as getMasterSourceRows()
    const { all, columns } = sheetData;

    for (let i = 2; i < all.length; i++) {
      const strValues = stringifyRowValues(all[i]);
      if (strValues.every(v => v === '')) continue;
      const data = {};
      columns.forEach((col, j) => { data[col.name] = strValues[j]; });

      const fullName = (data[columns[0].name] || '').trim();
      if (!fullName || !actualSet.has(fullName.toUpperCase())) continue;

      const photoRaw = getFieldByPattern(data, COL_PHOTO).trim();
      const cardId = getFieldByPattern(data, COL_CARD_ID).trim();

      if (!photoRaw || !cardId) {
        const reason = !photoRaw && !cardId ? 'Missing photo and S-КАДР ID'
                     : !photoRaw ? 'Missing photo' : 'Missing S-КАДР ID';
        skipped.push({ name: fullName, reason });
        continue;
      }
      if (seenCardIds.has(cardId)) {
        skipped.push({ name: fullName, reason: `Duplicate S-КАДР ID "${cardId}" (already used by ${seenCardIds.get(cardId)})` });
        continue;
      }
      seenCardIds.set(cardId, fullName);
      eligible.push({ rowIndex: i + 1, spreadsheetId, fullName, cardId, fileId: parseDriveId(photoRaw) });
    }
  });

  return { folderId: destFolder.getId(), folderUrl: destFolder.getUrl(), eligible, skipped };
}

/**
 * Copies and JPEG-converts photos for one batch of eligible entries (as
 * produced by startPhotoExport()) into the given folder, continuing across
 * multiple calls via the same time-boxed pattern _exportDoc() uses. Fetches
 * the full-resolution original blob (not the 800px thumbnail path
 * _getExportImageBlob() uses for Doc exports) and force-converts it to JPEG,
 * since the source file's original format is not guaranteed to be JPEG.
 *
 * @param {string} folderId - Destination Drive folder ID (from startPhotoExport()).
 * @param {Array<{rowIndex: number, spreadsheetId: string|null, fullName: string, cardId: string, fileId: string}>} entries
 * @returns {{
 *   results: Array<{name: string, url: string}>,
 *   skipped: Array<{name: string, reason: string}>,
 *   remaining: Array<{rowIndex: number, spreadsheetId: string|null, fullName: string, cardId: string, fileId: string}>
 * }}
 */
function copyPhotosBatch(folderId, entries) {
  const folder = DriveApp.getFolderById(folderId);
  const results = [];
  const skipped = [];
  const startTime = Date.now();
  let remaining = [];

  for (let i = 0; i < entries.length; i++) {
    if (Date.now() - startTime > EXPORT_TIME_LIMIT_MS) {
      remaining = entries.slice(i);
      break;
    }
    const { fullName, cardId, fileId } = entries[i];

    let blob;
    try {
      blob = DriveApp.getFileById(fileId).getBlob();
    } catch (e) {
      skipped.push({ name: fullName, reason: 'Photo file not accessible' });
      continue;
    }

    let jpegBlob;
    try {
      jpegBlob = blob.getAs('image/jpeg'); // force-convert regardless of source format
    } catch (e) {
      skipped.push({ name: fullName, reason: 'Could not convert photo to JPEG: ' + e.message });
      continue;
    }

    jpegBlob.setName(cardId + '.jpg');
    const file = folder.createFile(jpegBlob);
    results.push({ name: `${fullName} (${cardId}.jpg)`, url: file.getUrl() });
  }

  return { results, skipped, remaining };
}
