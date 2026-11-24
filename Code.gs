function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Center Admissions')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAdmissions(centerCode) {
  if (typeof centerCode !== 'string') return null;
  const code = centerCode.trim().toUpperCase();
  if (!code || code.length > 64) return null;

  const sheetId = PropertiesService.getScriptProperties().getProperty('ADMISSIONS_SHEET_ID');
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const januaryRow = findAdmissionsRow_(spreadsheet.getSheetByName('Jan 2026'), code, 2, 1, 2);
  const julyRow = findAdmissionsRow_(spreadsheet.getSheetByName('July 2026'), code, 3, 2, 12);
  if (!januaryRow && !julyRow) return null;

  // January lists VGU first. July lists SVGU first.
  const january = januaryRow && {svgu: januaryRow.values[1], vgu: januaryRow.values[0], total: januaryRow.values[2]};
  const july = julyRow && {svgu: julyRow.values[0], vgu: julyRow.values[1], total: julyRow.values[2]};
  return {code: (januaryRow || julyRow).code, january, july};
}

function findAdmissionsRow_(sheet, code, firstRow, codeColumn, totalsColumn) {
  if (!sheet) throw new Error('Admissions sheet is unavailable.');
  const rowCount = sheet.getLastRow() - firstRow + 1;
  if (rowCount < 1) return null;
  const codes = sheet.getRange(firstRow, codeColumn, rowCount, 1).getDisplayValues();
  const index = codes.findIndex(([value]) => value.trim().toUpperCase() === code);
  if (index === -1) return null;
  const [values] = sheet.getRange(firstRow + index, totalsColumn, 1, 3).getDisplayValues();
  return {code: codes[index][0].trim(), values};
}
