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
//  COMPANY ACCOUNT — the shared pooled account. Not a separate
//  entity with its own money; it's only ever partner balances.
//  Used as a special PaidBy value meaning "no one personally
//  fronted cash — the split directly debits each partner's
//  notional balance in the pool."
// ============================================================
const COMPANY_ACCOUNT_ID   = 'COMPANY';
const COMPANY_ACCOUNT_NAME = 'Company Account';

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
  RECORDED_AT:  10,
  EGP_EQUIVALENT: 11,
  EXCHANGE_RATE:  12
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
//  EXPENSE FREQUENCIES — hardcoded for now; only Monthly is
//  needed today, kept as a list so it's a one-line extension
//  later (e.g. Weekly, Yearly) without touching the schema.
// ============================================================
const EXPENSE_FREQUENCIES = ['Monthly'];

function getExpenseFrequencies() {
  return EXPENSE_FREQUENCIES;
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
//  toggle for whether the Bank Transfer sub-section renders in
//  the PDF.
//  Col 28 (AC) added: IncludeInstapayDetails — checkbox, same
//  idea but for the Instapay sub-section. Independent toggle so
//  a quotation can show either, both, or neither. Added at the
//  very end so no existing index shifts.
// ============================================================
const Q = {
  QUOTATION_ID:            0,
  CURRENT_VERSION:         1,
  ACCOUNT_ID:              2,
  ACCOUNT_NAME:            3,
  PROJECT_NAME:            4,
  PROJECT_DESC:            5,
  DATE_ISSUED:             6,
  MIN_DAYS:                7,
  MAX_DAYS:                8,
  DELIVERY_DEADLINE:       9,
  PRICING_MODE:            10,
  CURRENCY:                11,
  SUBTOTAL:                12,
  DISCOUNTED:              13,
  DISCOUNT_PERCENT:        14,
  DISCOUNT_AMOUNT:         15,
  TAXED:                   16,
  TAX_PERCENT:             17,
  TAX_AMOUNT:              18,
  TOTAL:                   19,
  STATUS:                  20,
  NOTES:                   21,
  FOLDER_URL:              22,
  CREATED_BY:              23,
  CREATED_AT:              24,
  LAST_UPDATED_BY:         25,
  LAST_UPDATED_AT:         26,
  INCLUDE_BANK_DETAILS:    27,
  INCLUDE_INSTAPAY_DETAILS: 28
};

// ============================================================
//  COLUMN INDEXES — Quote_Items (0-based)
//
//  Sheet headers (row 1):
//  Item ID | Quotation ID | Item Index | Display Value |
//  Item Name | Quantity | Description | Notes | Unit Price |
//  Subtotal | Status | Last Updated By | Last Updated At
//
//  ITEM_NAME is the admin-selected Category (from the Items
//  sheet) — mandatory. DISPLAY_VALUE is optional free text the
//  admin can enter to control what the client actually sees on
//  the PDF; when blank, ITEM_NAME (the Category) is shown instead.
// ============================================================
const QI = {
  ITEM_ID:         0,
  QUOTATION_ID:    1,
  ITEM_INDEX:      2,
  DISPLAY_VALUE:   3,
  ITEM_NAME:       4,
  QUANTITY:        5,
  DESCRIPTION:     6,
  NOTES:           7,
  UNIT_PRICE:      8,
  SUBTOTAL:        9,
  STATUS:          10,
  LAST_UPDATED_BY: 11,
  LAST_UPDATED_AT: 12
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
//  Project Name | Project Description | Display Value |
//  Item Name | Quantity | Description | Notes | Created At
//
//  ITEM_NAME is the Category carried over from the quotation.
//  DISPLAY_VALUE is the client-facing text carried over from the
//  quotation's Quote_Items row (falls back to Category when blank).
// ============================================================
const PI = {
  ITEM_ID:      0,
  PROJECT_ID:   1,
  QUOTATION_ID: 2,
  ACCOUNT_NAME: 3,
  PROJECT_NAME: 4,
  PROJECT_DESC: 5,
  DISPLAY_VALUE:6,
  ITEM_NAME:    7,
  QUANTITY:     8,
  DESCRIPTION:  9,
  NOTES:        10,
  CREATED_AT:   11
};

// ============================================================
//  COLUMN INDEXES — Bank_Details (0-based)
//
//  Sheet headers (row 1):
//  Label | Account Name | Account Number | Bank Name | IBAN |
//  SWIFT Code | Branch Name
//
//  Instapay now lives on its own sheet (see IP below) — this
//  sheet is purely bank-transfer details. Only row 2 (the first
//  data row) is read — one active bank details set is assumed.
// ============================================================
const BD = {
  LABEL:          0,
  ACCOUNT_NAME:   1,
  ACCOUNT_NUMBER: 2,
  BANK_NAME:      3,
  IBAN:           4,
  SWIFT_CODE:     5,
  BRANCH_NAME:    6
};

// ============================================================
//  COLUMN INDEXES — Instapay Details (0-based)
//
//  Sheet headers (row 1):
//  Label | Display Value | Instapay Link
//
//  DISPLAY_VALUE is what's shown to the client on the PDF (e.g.
//  a phone number or handle); LINK is the actual Instapay URL
//  the display value is hyperlinked to — same pattern as the
//  Website/Instagram footer links. Only row 2 is read.
// ============================================================
const IP = {
  LABEL:         0,
  DISPLAY_VALUE: 1,
  LINK:          2
};

// ============================================================
//  COLUMN INDEXES — Expense_Catalog (0-based)
//
//  The master list of expense TYPES — both recurring and
//  one-time. Doubles as the "template" for recurring expenses:
//  Last* fields always reflect the most recently logged
//  instance, and are only ever used as a pre-fill suggestion —
//  never an authority. The real price history lives in Expenses.
// ============================================================
const EC = {
  CATALOG_ID:        0,
  NAME:              1,
  CATEGORY:          2,
  VENDOR:            3,
  IS_RECURRING:      4,
  FREQUENCY:         5,
  LAST_AMOUNT:       6,
  LAST_CURRENCY:     7,
  LAST_PAID_BY:      8,
  LAST_PAID_BY_NAME: 9,
  LAST_SPLIT_JSON:   10,
  STATUS:            11,
  NOTES:             12,
  CREATED_BY:        13,
  CREATED_AT:        14
};

// ============================================================
//  COLUMN INDEXES — Expenses (0-based)
//
//  One row = one real transaction in time (= the price history
//  for its CatalogId). PeriodKey ("2026-07") is set only for
//  recurring instances, used to detect what's already logged.
//
//  Cols 13-14 (N-O) added: EgpEquivalent / ExchangeRate — mirrors
//  the same pattern used on Payments (PAY.EGP_EQUIVALENT /
//  PAY.EXCHANGE_RATE). For EGP-currency expenses these just equal
//  the amount / 1. Added at the very end so no existing index shifts.
// ============================================================
const EXP = {
  EXPENSE_ID:     0,
  CATALOG_ID:     1,
  CATALOG_NAME:   2,
  CATEGORY:       3,
  AMOUNT:         4,
  CURRENCY:       5,
  DATE:           6,
  PAID_BY:        7,
  PAID_BY_NAME:   8,
  PERIOD_KEY:     9,
  NOTES:          10,
  CREATED_BY:     11,
  CREATED_AT:     12,
  EGP_EQUIVALENT: 13,
  EXCHANGE_RATE:  14
};

// ============================================================
//  COLUMN INDEXES — Expense_Split (0-based)
//  Mirrors Revenue_Distribution's shape.
// ============================================================
const ESPL = {
  SPLIT_ID:    0,
  EXPENSE_ID:  1,
  CATALOG_ID:  2,
  PERSON_ID:   3,
  PERSON_NAME: 4,
  PERCENT:     5,
  AMOUNT:      6,
  CURRENCY:    7,
  CREATED_BY:  8,
  CREATED_AT:  9
};

// ============================================================
//  COLUMN INDEXES — Drawings (0-based)
//  A partner pulling their own already-earned balance out.
//  Always 100% to one person — no split table needed.
// ============================================================
const DRW = {
  DRAWING_ID:  0,
  PERSON_ID:   1,
  PERSON_NAME: 2,
  AMOUNT:      3,
  CURRENCY:    4,
  DATE:        5,
  METHOD:      6,
  NOTES:       7,
  CREATED_BY:  8,
  CREATED_AT:  9,
  EGP_EQUIVALENT: 10,
  EXCHANGE_RATE:  11
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
    .addItem('Team',        'openTeamModule')
    .addItem('Expenses',    'openExpensesModule')
    .addItem('Drawings',    'openDrawingsModule')
    .addToUi();

  // Home is always the landing tab, table freshly synced — fully
  // automatic, no menu item. See Index.gs.
  try {
    refreshHome();
  } catch (e) {
    Logger.log('refreshHome error on open: ' + e);
  }
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

function openExpensesModule() {
  const html = HtmlService.createHtmlOutputFromFile('ExpensesForm')
    .setWidth(1000).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Expenses');
}

function openDrawingsModule() {
  const html = HtmlService.createHtmlOutputFromFile('DrawingsForm')
    .setWidth(880).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Drawings');
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
