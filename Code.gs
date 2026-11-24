const PROFILE_FIELDS = ['company', 'owner', 'mobile', 'email', 'address', 'vguId', 'sgvuId', 'activeBy'];
const SESSION_FIELDS = ['svgu', 'vgu', 'total'];
const GROUPS = {
  profile: {tab: 'July 2026', first: 3, code: 2, columns: [3,4,5,6,7,8,9,10], keys: PROFILE_FIELDS},
  january: {tab: 'Jan 2026', first: 2, code: 1, columns: [3,2,4], keys: SESSION_FIELDS},
  july: {tab: 'July 2026', first: 3, code: 2, columns: [12,13,14], keys: SESSION_FIELDS},
  nextJanuary: {tab: 'Jan 2027', first: 2, code: 1, columns: [3,2,4], keys: SESSION_FIELDS}
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Center Admissions')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAdmissions(centerCode) {
  const code = normalizeCode_(centerCode);
  if (!code) return null;
  return withLock_(() => readCenter_(openSheet_(), code));
}

function updateCenter(centerCode, groupKey, key, operation, value, expected) {
  const code = normalizeCode_(centerCode);
  const group = Object.prototype.hasOwnProperty.call(GROUPS, groupKey) ? GROUPS[groupKey] : null;
  const index = group ? group.keys.indexOf(key) : -1;
  if (!code || index < 0 || !['save', 'delete', 'restore'].includes(operation)) {
    throw new Error('Invalid request.');
  }
  return withLock_(() => {
    const book = openSheet_();
    if (!readCenter_(book, code)) throw new Error('Center code not found.');
    let sheet = book.getSheetByName(group.tab);
    let row = findRow_(sheet, group, code);
    let cell = row ? sheet.getRange(row, group.columns[index]) : null;
    const current = cell ? cell.getDisplayValue() : '';
    if (current !== expected) throw new Error('This field changed. Reload the center and try again.');
    const backup = getBackup_(book, code, groupKey, key);
    if (operation === 'restore' && (current || !backup)) throw new Error('Nothing to restore.');

    if (operation === 'save') {
      const next = validateValue_(groupKey, value);
      if (cell && cell.getFormula()) throw new Error('This field is calculated in the sheet.');
      if (!next) throw new Error('Use Delete to clear this field.');
      if (!sheet) {
        if (groupKey !== 'nextJanuary') throw new Error('Session sheet is unavailable.');
        sheet = book.insertSheet(group.tab);
        sheet.getRange('A1:D1').setValues([['Center Code', 'VGU', 'SGVU', 'Total']]);
      }
      if (!row) {
        row = Math.max(sheet.getLastRow() + 1, group.first);
        sheet.getRange(row, group.code).setValue("'" + code);
      }
      cell = sheet.getRange(row, group.columns[index]);
      cell.setValue(groupKey === 'profile' ? "'" + next : Number(next));
    } else if (operation === 'delete' && current) {
      if (cell.getFormula()) throw new Error('This field is calculated in the sheet.');
      saveBackup_(book, code, groupKey, key, cell.getValue());
      SpreadsheetApp.flush();
      cell.clearContent();
    } else if (operation === 'restore') {
      const restored = JSON.parse(backup.payload);
      cell.setValue(typeof restored === 'string' ? "'" + restored : restored);
    }
    SpreadsheetApp.flush();
    return readCenter_(book, code);
  });
}

function validateValue_(groupKey, value) {
  if (typeof value !== 'string') throw new Error('Enter a valid value.');
  const trimmed = value.trim();
  if (groupKey === 'profile') {
    if (trimmed.length > 1000) throw new Error('Use 1,000 characters or fewer.');
  } else if (!/^\d{1,12}$/.test(trimmed)) {
    throw new Error('Enter a whole admission count of zero or more.');
  }
  return trimmed;
}

function readCenter_(book, code) {
  const result = {code, groups: {}};
  const backups = readBackups_(book, code);
  let found = false;
  for (const [groupKey, group] of Object.entries(GROUPS)) {
    const sheet = book.getSheetByName(group.tab);
    const row = findRow_(sheet, group, code);
    found = found || Boolean(row);
    const start = Math.min(...group.columns);
    const width = Math.max(...group.columns) - start + 1;
    const range = row ? sheet.getRange(row, start, 1, width) : null;
    const values = range ? range.getDisplayValues()[0] : [];
    const formulas = range ? range.getFormulas()[0] : [];
    result.groups[groupKey] = Object.fromEntries(group.keys.map((key, i) => {
      const column = group.columns[i] - start;
      const value = values[column] || '';
      return [key, {value, calculated: Boolean(formulas[column]), canRestore: !value && backups.has(groupKey + '/' + key)}];
    }));
  }
  return found ? result : null;
}

function findRow_(sheet, group, code) {
  if (!sheet || sheet.getLastRow() < group.first) return 0;
  const codes = sheet.getRange(group.first, group.code, sheet.getLastRow() - group.first + 1, 1).getDisplayValues();
  const index = codes.findIndex(([value]) => normalizeCode_(value) === code);
  return index < 0 ? 0 : group.first + index;
}

function normalizeCode_(code) {
  return typeof code === 'string' && code.trim().length <= 64 ? code.trim().toUpperCase() : '';
}

function openSheet_() {
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMISSIONS_SHEET_ID'));
}

function withLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return callback(); } finally { lock.releaseLock(); }
}

function readBackups_(book, code) {
  const sheet = book.getSheetByName('Portal Restore History');
  const backups = new Map();
  if (!sheet || sheet.getLastRow() < 2) return backups;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getDisplayValues();
  for (const [savedCode, group, field, payload] of rows) {
    if (savedCode === code) backups.set(group + '/' + field, {payload});
  }
  return backups;
}

function getBackup_(book, code, group, field) {
  return readBackups_(book, code).get(group + '/' + field);
}

function saveBackup_(book, code, group, field, value) {
  let sheet = book.getSheetByName('Portal Restore History');
  if (!sheet) {
    sheet = book.insertSheet('Portal Restore History');
    sheet.appendRow(['Center code', 'Group', 'Field', 'Value (JSON)', 'Deleted at']);
  }
  sheet.appendRow(["'" + code, group, field, JSON.stringify(value), new Date()]);
}
