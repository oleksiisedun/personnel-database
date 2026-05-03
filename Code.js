function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WebEditor')
    .addItem('Open Editor', 'openWebEditor')
    .addToUi();
}

function openWebEditor() {
  const html = HtmlService.createHtmlOutputFromFile('WebEditor')
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'WebEditor');
}

function getSchemaAndData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Database');
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
  return { columns, rows };
}

function updateRow(rowIndex, values) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Database');
  if (!sheet) throw new Error('Sheet "Database" not found.');
  sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
  return true;
}
