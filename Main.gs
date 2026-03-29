// ============================================================
//  CONSTANTS
// ============================================================
const CLIENTS_FOLDER_ID        = '19iNsbyldpjF8pVuc9WUGIHB8bibYbu0J';
const LEADS_FOLDER_ID = '1DbnsDxGgymXsFMnCK8uUIUP0UFXuMv3z';
const ARCHIVED_FOLDER_ID   = '1TQAW86XAPJB2PdXWYYQoTNSdRwpahY-l';

// ============================================================
//  COLUMN INDEXES — Accounts (0-based)
// ============================================================
const A = {
  ACCOUNT_ID:    0,
  ACCOUNT_NAME:  1,
  SOURCE:        2,
  CHANNEL:       3,
  COUNTRY:       4,
  INDUSTRY:      5,
  IS_REGISTERED: 6,
  STATUS:        7,
  CREATED_BY:    8,
  CREATED_AT:    9,
  FOLDER_URL:    10,
  NOTES:         11
};

// ============================================================
//  COLUMN INDEXES — Quotations (0-based)
// ============================================================
const Q = {
  QUOTATION_ID:      0,
  CURRENT_VERSION:   1,
  ACCOUNT_ID:        2,
  ACCOUNT_NAME:      3,
  PROJECT_NAME:      4,
  PROJECT_DESC:      5,
  DATE_ISSUED:       6,
  MIN_DAYS:          7,
  MAX_DAYS:          8,
  DELIVERY_DEADLINE: 9,
  PRICING_MODE:      10,
  CURRENCY:          11,
  SUBTOTAL:          12,
  DISCOUNTED:        13,
  DISCOUNT_PERCENT:  14,
  DISCOUNT_AMOUNT:   15,
  TAXED:             16,
  TAX_PERCENT:       17,
  TAX_AMOUNT:        18,
  TOTAL:             19,
  STATUS:            20,
  NOTES:             21,
  FOLDER_URL:        22,
  CREATED_BY:        23,
  CREATED_AT:        24,
  LAST_UPDATED_BY:   25,
  LAST_UPDATED_AT:   26
};

// ============================================================
//  COLUMN INDEXES — Quote_Items (0-based)
// ============================================================
const QI = {
  ITEM_ID:         0,
  QUOTATION_ID:    1,
  ITEM_INDEX:      2,
  ITEM_NAME:       3,
  QUANTITY:        4,
  DESCRIPTION:     5,
  NOTES:           6,
  UNIT_PRICE:      7,
  SUBTOTAL:        8,
  STATUS:          9,
  LAST_UPDATED_BY: 10,
  LAST_UPDATED_AT: 11
};

// ============================================================
//  COLUMN INDEXES — Projects (0-based)
// ============================================================
const P = {
  PROJECT_ID:        0,
  QUOTATION_ID:      1,
  ACCOUNT_ID:        2,
  ACCOUNT_NAME:      3,
  PROJECT_NAME:      4,
  PROJECT_DESC:      5,
  DELIVERY_DEADLINE: 6,
  DUE_DATE:          7,
  STATUS:            8,
  INTERNAL_NOTES:    9,
  CREATED_AT:        10,
  COMPLETED_AT:      11
};

// ============================================================
//  COLUMN INDEXES — Project_Items (0-based)
// ============================================================
const PI = {
  ITEM_ID:           0,
  PROJECT_ID:        1,
  QUOTATION_ID:      2,
  ACCOUNT_NAME:      3,
  PROJECT_NAME:      4,
  ITEM_NAME:         5,
  QUANTITY:          6,
  DESCRIPTION:       7,
  NOTES:             8,
  DELIVERY_STATUS:   9,
  ASSIGNED_TO:       10,
  REDO_COUNT:        11,
  UPLOADED_FILE_URL: 12,
  INTERNAL_NOTES:    13,
  DUE_DATE:          14,
  COMPLETED_AT:      15,
  CREATED_AT:        16
};

// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CRM System')
    .addItem('Accounts',   'openAccountsModule')
    .addItem('Quotations', 'openQuotationsModule')
    .addToUi();
}

function openAccountsModule() {
  const html = HtmlService.createHtmlOutputFromFile('AccountsForm')
    .setWidth(960)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Accounts');
}

function openQuotationsModule() {
  const html = HtmlService.createHtmlOutputFromFile('QuotationsForm')
    .setWidth(960)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Quotations');
}

// ============================================================
//  SHEET HELPER
// ============================================================
function getSheet(name) {
  return SpreadsheetApp.getActive().getSheetByName(name);
}

// ============================================================
//  DRIVE HELPER
// ============================================================
function moveToDeleted(fileUrl) {
  if (!fileUrl) return;
  try {
    const fileId      = fileUrl.match(/[-\w]{25,}/)[0];
    const file        = DriveApp.getFileById(fileId);
    const trashFolder = DriveApp.getFolderById(DELETED_FILES_FOLDER_ID);
    const parents     = file.getParents();
    trashFolder.addFile(file);
    while (parents.hasNext()) parents.next().removeFile(file);
  } catch(e) {
    Logger.log('moveToDeleted error: ' + e);
  }
}

// ============================================================
//  LOGGING HELPER
// ============================================================
function writeLog(sheetName, module, recordId, recordLabel,
                  accountId, accountName, fieldChanged,
                  oldValue, newValue, notes) {
  const sheet = getSheet(sheetName);
  if (!sheet) return;
  sheet.appendRow([
    Utilities.getUuid(),
    module,
    String(recordId    || ''),
    String(recordLabel || ''),
    String(accountId   || ''),
    String(accountName || ''),
    String(fieldChanged|| ''),
    oldValue !== null && oldValue !== undefined ? String(oldValue) : '',
    newValue !== null && newValue !== undefined ? String(newValue) : '',
    Session.getActiveUser().getEmail(),
    new Date(),
    String(notes || '')
  ]);
}

// ============================================================
//  DROPDOWN OPTIONS — shared by all modules
// ============================================================
function getDropdownOptions(sheetName) {
  try {
    const sheet = getSheet(sheetName);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 2, lastRow - 1, 1)
      .getValues().flat().filter(v => v);
  } catch(e) {
    Logger.log('getDropdownOptions error: ' + e);
    return [];
  }
}
