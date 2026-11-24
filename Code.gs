function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Center Admissions').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAdmissions(centerCode) {
  if (typeof centerCode !== 'string') return null;
  const code = centerCode.trim().toUpperCase();
  if (!code || code.length > 64) return null;
  const sheetId = PropertiesService.getScriptProperties().getProperty('ADMISSIONS_SHEET_ID');
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Sheet1');
  if (!sheet) throw new Error('Admissions sheet is unavailable.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return null;
  const codes = sheet.getRange(3, 2, lastRow - 2, 1).getDisplayValues();
  const index = codes.findIndex(([value]) => value.trim().toUpperCase() === code);
  if (index === -1) return null;
  const [[svgu, vgu, total]] = sheet.getRange(index + 3, 12, 1, 3).getDisplayValues();
  return {code: codes[index][0].trim(), svgu, vgu, total};
}
