const CONFIG = Object.freeze({
  SPREADSHEET_ID: '1Juvs_pH3frP7bUOe_-9NMp0WTj5FBcRj5pHEV_vYHfU',
  LIST_SHEET: 'LIST',
  RANK_SHEET: 'RANK',
  OFFICE_SHEET: 'OFFICE_DIRECTORY'
});

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('NBP Personnel Profile Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getFormOptions() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return {
    ranks: readRankOptions_(ss),
    offices: readOfficeOptions_(ss),
    camps: ['NBP', 'MAXIMUM', 'MEDIUM', 'MINIMUM', 'RDC'],
    genders: ['MALE', 'FEMALE'],
    suffixes: ['', 'JR.', 'SR.', 'II', 'III', 'IV', 'V']
  };
}

function submitPersonnel(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.LIST_SHEET);
    if (!sheet) throw new Error('LIST sheet was not found.');

    const record = normalizePayload_(payload || {});
    validateRecord_(record);

    const headers = ensureHeaders_(sheet);
    const duplicate = detectDuplicate_(sheet, headers, record);
    const rowObject = {
      'RECORD ID': Utilities.getUuid(),
      'BADGE NUMBER': record.badgeNumber,
      'RANK': record.rank,
      'LAST NAME': record.lastName,
      'FIRST NAME': record.firstName,
      'MIDDLE NAME': record.middleName,
      'SUFFIX': record.suffix,
      'CAMP': record.camp,
      'OFFICE': record.office,
      'GENDER': record.gender,
      'CLASSIFICATION': record.classification,
      'TYPE': record.type,
      'DUPLICATE STATUS': duplicate.status,
      'DUPLICATE TYPE': duplicate.type,
      'CREATED AT': new Date(),
      'UPDATED AT': new Date()
    };

    const row = headers.map(header => rowObject[header] !== undefined ? rowObject[header] : '');
    sheet.appendRow(row);
    const rowNumber = sheet.getLastRow();
    highlightDuplicateStatus_(sheet, headers, rowNumber, duplicate.status);

    return {
      ok: true,
      duplicateStatus: duplicate.status,
      duplicateType: duplicate.type
    };
  } finally {
    lock.releaseLock();
  }
}

function readRankOptions_(ss) {
  const sheet = ss.getSheetByName(CONFIG.RANK_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift().map(normalizeHeader_);
  const rankIndex = findHeaderIndex_(headers, ['RANK']);
  const classificationIndex = findHeaderIndex_(headers, ['CLASSIFICATION', 'COMMISSIONED OR NON-COMMISSIONED']);
  const typeIndex = findHeaderIndex_(headers, ['TYPE', 'CO OR CTO']);
  const activeIndex = findHeaderIndex_(headers, ['ACTIVE']);
  const sortIndex = findHeaderIndex_(headers, ['SORT ORDER']);

  return values
    .filter(row => row[rankIndex])
    .filter(row => activeIndex < 0 || String(row[activeIndex]).toUpperCase() !== 'FALSE')
    .map(row => ({
      rank: cleanUpper_(row[rankIndex]),
      classification: cleanText_(row[classificationIndex]),
      type: cleanUpper_(row[typeIndex]),
      sortOrder: Number(row[sortIndex]) || 9999
    }))
    .sort((a, b) => a.sortOrder - b.sortOrder || a.rank.localeCompare(b.rank));
}

function readOfficeOptions_(ss) {
  const sheet = ss.getSheetByName(CONFIG.OFFICE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift().map(normalizeHeader_);
  const campIndex = findHeaderIndex_(headers, ['CAMP']);
  const officeIndex = findHeaderIndex_(headers, ['OFFICE']);
  const activeIndex = findHeaderIndex_(headers, ['ACTIVE']);
  const sortIndex = findHeaderIndex_(headers, ['SORT ORDER']);

  return values
    .filter(row => row[campIndex] && row[officeIndex])
    .filter(row => activeIndex < 0 || String(row[activeIndex]).toUpperCase() !== 'FALSE')
    .map(row => ({
      camp: cleanUpper_(row[campIndex]),
      office: cleanUpper_(row[officeIndex]),
      sortOrder: Number(row[sortIndex]) || 9999
    }))
    .sort((a, b) => a.camp.localeCompare(b.camp) || a.sortOrder - b.sortOrder || a.office.localeCompare(b.office));
}

function normalizePayload_(payload) {
  return {
    badgeNumber: String(payload.badgeNumber || '').trim(),
    rank: cleanUpper_(payload.rank),
    lastName: cleanUpper_(payload.lastName),
    firstName: cleanUpper_(payload.firstName),
    middleName: cleanUpper_(payload.middleName),
    suffix: cleanUpper_(payload.suffix),
    camp: cleanUpper_(payload.camp),
    office: cleanUpper_(payload.office),
    gender: cleanUpper_(payload.gender),
    classification: cleanText_(payload.classification),
    type: cleanUpper_(payload.type)
  };
}

function validateRecord_(record) {
  const required = {
    'Badge number': record.badgeNumber,
    Rank: record.rank,
    'Last name': record.lastName,
    'First name': record.firstName,
    Camp: record.camp,
    Office: record.office,
    Gender: record.gender
  };

  const missing = Object.keys(required).filter(key => !required[key]);
  if (missing.length) throw new Error(`Please complete: ${missing.join(', ')}.`);
}

function ensureHeaders_(sheet) {
  const requiredHeaders = [
    'RECORD ID',
    'BADGE NUMBER',
    'RANK',
    'LAST NAME',
    'FIRST NAME',
    'MIDDLE NAME',
    'SUFFIX',
    'CAMP',
    'OFFICE',
    'GENDER',
    'CLASSIFICATION',
    'TYPE',
    'DUPLICATE STATUS',
    'DUPLICATE TYPE',
    'CREATED AT',
    'UPDATED AT'
  ];

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const hasDataHeaders = existing.some(Boolean);

  if (!hasDataHeaders) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    styleHeader_(sheet, requiredHeaders.length);
    return requiredHeaders;
  }

  const headers = existing.map(value => normalizeHeader_(value));
  const missing = requiredHeaders.filter(header => !headers.includes(header));
  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers.push(...missing);
    styleHeader_(sheet, headers.length);
  }
  return headers;
}

function detectDuplicate_(sheet, headers, record) {
  if (sheet.getLastRow() < 2) return { status: 'UNIQUE', type: 'NONE' };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getDisplayValues();
  const badgeIndex = headers.indexOf('BADGE NUMBER');
  const lastIndex = headers.indexOf('LAST NAME');
  const firstIndex = headers.indexOf('FIRST NAME');
  const middleIndex = headers.indexOf('MIDDLE NAME');
  const suffixIndex = headers.indexOf('SUFFIX');

  const incomingBadge = normalizeBadge_(record.badgeNumber);
  const incomingName = normalizeNameKey_(record.lastName, record.firstName, record.middleName, record.suffix);
  let badgeMatch = false;
  let nameMatch = false;

  data.forEach(row => {
    if (badgeIndex >= 0 && normalizeBadge_(row[badgeIndex]) === incomingBadge) badgeMatch = true;
    const rowName = normalizeNameKey_(
      lastIndex >= 0 ? row[lastIndex] : '',
      firstIndex >= 0 ? row[firstIndex] : '',
      middleIndex >= 0 ? row[middleIndex] : '',
      suffixIndex >= 0 ? row[suffixIndex] : ''
    );
    if (rowName && rowName === incomingName) nameMatch = true;
  });

  if (badgeMatch && nameMatch) return { status: 'DUPLICATE', type: 'BADGE NUMBER + FULL NAME' };
  if (badgeMatch) return { status: 'DUPLICATE', type: 'BADGE NUMBER' };
  if (nameMatch) return { status: 'POSSIBLE DUPLICATE', type: 'FULL NAME' };
  return { status: 'UNIQUE', type: 'NONE' };
}

function highlightDuplicateStatus_(sheet, headers, rowNumber, status) {
  const column = headers.indexOf('DUPLICATE STATUS') + 1;
  if (!column) return;

  const cell = sheet.getRange(rowNumber, column);
  cell.setFontWeight('bold').setHorizontalAlignment('center');

  if (status === 'DUPLICATE') {
    cell.setBackground('#F4CCCC').setFontColor('#990000');
  } else if (status === 'POSSIBLE DUPLICATE') {
    cell.setBackground('#FFF2CC').setFontColor('#7F6000');
  } else {
    cell.setBackground(null).setFontColor('#38761D');
  }
}

function styleHeader_(sheet, columnCount) {
  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground('#174D36')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

function normalizeHeader_(value) {
  return cleanUpper_(value).replace(/\s+/g, ' ');
}

function findHeaderIndex_(headers, candidates) {
  for (const candidate of candidates) {
    const index = headers.indexOf(normalizeHeader_(candidate));
    if (index >= 0) return index;
  }
  return -1;
}

function cleanUpper_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toUpperCase();
}

function cleanText_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ');
}

function normalizeBadge_(value) {
  return cleanUpper_(value).replace(/\s+/g, '');
}

function normalizeNameKey_(lastName, firstName, middleName, suffix) {
  return [lastName, firstName, middleName, suffix]
    .map(cleanUpper_)
    .filter(Boolean)
    .join('|');
}
