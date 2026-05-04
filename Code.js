function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WebEditor')
    .addItem('Open Editor', 'openWebEditor')
    .addToUi();
}

function openWebEditor() {
  const html = HtmlService.createTemplateFromFile('WebEditor').evaluate()
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'WebEditor');
}

function getSchemaAndData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Database');
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

  const tableHeadersMap = {};
  const handbookSheet = ss.getSheetByName('Handbook');
  if (handbookSheet) {
    const hbData = handbookSheet.getDataRange().getValues();
    for (let r = 1; r < hbData.length; r++) {
      const dataType = String(hbData[r][0]).trim().toLowerCase();
      if (!dataType) continue;
      const headers = hbData[r].slice(1).map(h => String(h)).filter(h => h !== '');
      if (headers.length) tableHeadersMap[dataType] = headers;
    }
  }
  columns.forEach(col => {
    if (col.type.endsWith('-table')) col.tableHeaders = tableHeadersMap[col.type] || [];
  });

  return { columns, rows };
}

function getImagesDataUrls(fileIds) {
  const result = {};
  fileIds.forEach(fileId => {
    if (!fileId) return;
    try {
      const blob = DriveApp.getFileById(fileId).getBlob();
      result[fileId] = `data:${blob.getContentType() || 'image/jpeg'};base64,${Utilities.base64Encode(blob.getBytes())}`;
    } catch (e) {
      result[fileId] = null;
    }
  });
  return result;
}

function updateRow(rowIndex, values) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Database');  // keep separate for simplicity
  if (!sheet) throw new Error('Sheet "Database" not found.');
  sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
  return true;
}
