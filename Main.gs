// ============================================================
//  CONSTANTS
// ============================================================
const CLIENTS_FOLDER_ID  = '19iNsbyldpjF8pVuc9WUGIHB8bibYbu0J';
const LEADS_FOLDER_ID    = '1DbnsDxGgymXsFMnCK8uUIUP0UFXuMv3z';
const ARCHIVED_FOLDER_ID = '1TQAW86XAPJB2PdXWYYQoTNSdRwpahY-l';
const RUYA_TEAM_FOLDER_ID = '1S_nrzYMIJdpwBb84t-73D-NU7WDt33Nc';
const RUYA_QUOTATIONS_PIPELINE_ID  = '14_BP1bYHWJ3hqXwHE5WfYxp6hskjqWqj';
const RUYA_QUOTATIONS_CONFIRMED_ID = '1StmfEKC-cUVK-Y-5WrJI1avmgvXQBe5s';
const RUYA_QUOTATIONS_FULFILLED_ID = '1OBZ6QVCq_oG-b0nFK3Ejmaol8Iz7JjTn';

// ============================================================
//  COLUMN INDEXES — Payments (0-based)
// ============================================================
const PAY = {
  PAYMENT_ID:   0,
  PROJECT_ID:   1,
  QUOTATION_ID: 2,
  ACCOUNT_NAME: 3,
  PROJECT_NAME: 4,
  AMOUNT:       5,
  DATE:         6,
  METHOD:       7,
  NOTES:        8,
  RECORDED_BY:  9,
  RECORDED_AT:  10
};

// ============================================================
//  COLUMN INDEXES — Revenue_Distribution (0-based)
// ============================================================
const DIST = {
  DISTRIBUTION_ID: 0,
  PROJECT_ID:      1,
  QUOTATION_ID:    2,
  ACCOUNT_NAME:    3,
  PROJECT_NAME:    4,
  PERSON_ID:       5,
  PERSON_NAME:     6,
  PERCENT:         7,
  AMOUNT:          8,
  CURRENCY:        9,
  NOTES:           10,
  CREATED_BY:      11,
  CREATED_AT:      12
};

// ============================================================
//  PAYMENT METHODS — hardcoded (small, stable list)
// ============================================================
const PAYMENT_METHODS = ['Cash', 'Bank Transfer', 'Instapay', 'Vodafone Cash', 'Other'];

function getPaymentMethods() {
  return PAYMENT_METHODS;
}
// ============================================================
//  COLUMN INDEXES — Team (0-based)
//
//  Col 7 (H) added: IsPartner — checkbox, marks whether this
//  team member is a business Partner (Owner). Used to filter
//  who can appear in Expense/Drawing splits. Kept separate from
//  the free-text Role field so filtering never depends on exact
//  wording. Added at the very end so no existing index shifts.
// ============================================================
const TM = {
  TEAM_ID:    0,
  NAME:       1,
  PHONE:      2,
  EMAIL:      3,
  ROLE:       4,
  FOLDER_URL: 5,
  CREATED_AT: 6,
  IS_PARTNER: 7
};
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
//
//  Col 27 (AB) added: IncludeBankDetails — checkbox, per-quotation
//  toggle for whether the Payment Details section renders in the PDF.
//  Added at the very end so no existing index shifts.
// ============================================================
const Q = {
  QUOTATION_ID:         0,
  CURRENT_VERSION:      1,
  ACCOUNT_ID:           2,
  ACCOUNT_NAME:         3,
  PROJECT_NAME:         4,
  PROJECT_DESC:         5,
  DATE_ISSUED:          6,
  MIN_DAYS:             7,
  MAX_DAYS:             8,
  DELIVERY_DEADLINE:    9,
  PRICING_MODE:         10,
  CURRENCY:             11,
  SUBTOTAL:             12,
  DISCOUNTED:           13,
  DISCOUNT_PERCENT:     14,
  DISCOUNT_AMOUNT:      15,
  TAXED:                16,
  TAX_PERCENT:          17,
  TAX_AMOUNT:           18,
  TOTAL:                19,
  STATUS:               20,
  NOTES:                21,
  FOLDER_URL:           22,
  CREATED_BY:           23,
  CREATED_AT:           24,
  LAST_UPDATED_BY:      25,
  LAST_UPDATED_AT:      26,
  INCLUDE_BANK_DETAILS: 27
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
//
//  Sheet headers (row 1):
//  Project ID | Quotation ID | Account ID | Account Name |
//  Project Name | Project Description | Delivery Deadline |
//  Due Date | Status | Internal Notes | Created At |
//  Completed At | Total Amount | Currency |
//  Total Commission | Remaining Amount
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
  COMPLETED_AT:      11,
  TOTAL_AMOUNT:      12,
  CURRENCY:          13,
  TOTAL_COMMISSION:  14,
  REMAINING_AMOUNT:  15
};

// ============================================================
//  COLUMN INDEXES — Project_Items (0-based)
//
//  Sheet headers (row 1):
//  Item ID | Project ID | Quotation ID | Account Name |
//  Project Name | Project Description | Item Name |
//  Quantity | Description | Notes | Created At
// ============================================================
const PI = {
  ITEM_ID:     0,
  PROJECT_ID:  1,
  QUOTATION_ID:2,
  ACCOUNT_NAME:3,
  PROJECT_NAME:4,
  PROJECT_DESC:5,
  ITEM_NAME:   6,
  QUANTITY:    7,
  DESCRIPTION: 8,
  NOTES:       9,
  CREATED_AT:  10
};

// ============================================================
//  COLUMN INDEXES — Bank_Details (0-based)
//
//  Sheet headers (row 1):
//  Label | Account Name | Bank Name | IBAN | SWIFT Code |
//  Nationality | Branch Name | Branch Code | Address |
//  Instapay Payment Address | Instapay Mobile Number
//
//  Only row 2 (the first data row) is read — one active bank
//  details set is assumed for now.
// ============================================================
const BD = {
  LABEL:            0,
  ACCOUNT_NAME:     1,
  BANK_NAME:        2,
  IBAN:             3,
  SWIFT_CODE:       4,
  NATIONALITY:      5,
  BRANCH_NAME:      6,
  BRANCH_CODE:      7,
  ADDRESS:          8,
  INSTAPAY_ADDRESS: 9,
  INSTAPAY_MOBILE:  10
};

// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CRM System')
    .addItem('Accounts',    'openAccountsModule')
    .addItem('Quotations',  'openQuotationsModule')
    .addItem('Projects',    'openProjectsModule')
    .addItem('Team',       'openTeamModule')
    .addToUi();
}

function openAccountsModule() {
  const html = HtmlService.createHtmlOutputFromFile('AccountsForm')
    .setWidth(960).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Accounts');
}

function openQuotationsModule() {
  const html = HtmlService.createHtmlOutputFromFile('QuotationsForm')
    .setWidth(960).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Quotations');
}

function openProjectsModule() {
  const html = HtmlService.createHtmlOutputFromFile('ProjectsForm')
    .setWidth(1100).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Projects');
}

function openTeamModule() {
  const html = HtmlService.createHtmlOutputFromFile('TeamForm')
    .setWidth(960).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Team');
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
function moveToArchived(fileOrFolderUrl) {
  if (!fileOrFolderUrl) return;
  try {
    const id           = fileOrFolderUrl.match(/[-\w]{25,}/)[0];
    const item         = DriveApp.getFileById(id);
    const archivedRoot = DriveApp.getFolderById(ARCHIVED_FOLDER_ID);
    const parents      = item.getParents();
    archivedRoot.addFile(item);
    while (parents.hasNext()) parents.next().removeFile(item);
  } catch(e) {
    Logger.log('moveToArchived error: ' + e);
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
    String(recordId     || ''),
    String(recordLabel  || ''),
    String(accountId    || ''),
    String(accountName  || ''),
    String(fieldChanged || ''),
    oldValue !== null && oldValue !== undefined ? String(oldValue) : '',
    newValue !== null && newValue !== undefined ? String(newValue) : '',
    Session.getActiveUser().getEmail(),
    new Date(),
    String(notes || '')
  ]);
}

// ============================================================
//  DROPDOWN OPTIONS
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

// ============================================================
//  HELPER — extract plain URL from formula or raw string
// ============================================================
function extractUrl(cellValue) {
  if (!cellValue) return '';
  const str = String(cellValue);
  if (str.toUpperCase().startsWith('=HYPERLINK')) {
    const match = str.match(/\"(https?:\/\/[^\"]+)\"/);
    return match ? match[1] : '';
  }
  return str;
}
