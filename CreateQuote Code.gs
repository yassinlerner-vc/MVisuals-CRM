// ============================================================
//  CONSTANTS
// ============================================================
const ACCOUNTS_FOLDER_ID        = '19iNsbyldpjF8pVuc9WUGIHB8bibYbu0J';
const GLOBAL_PIPELINE_FOLDER_ID = '1DbnsDxGgymXsFMnCK8uUIUP0UFXuMv3z';

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
  ITEM_ID:          0,
  QUOTATION_ID:     1,
  ITEM_INDEX:       2,
  ITEM_NAME:        3,
  QUANTITY:         4,
  DESCRIPTION:      5,
  NOTES:            6,
  UNIT_PRICE:       7,
  SUBTOTAL:         8,
  STATUS:           9,
  LAST_UPDATED_BY:  10,
  LAST_UPDATED_AT:  11
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
  ITEM_ID:          0,
  PROJECT_ID:       1,
  QUOTATION_ID:     2,
  ACCOUNT_NAME:     3,
  PROJECT_NAME:     4,
  ITEM_NAME:        5,
  QUANTITY:         6,
  DESCRIPTION:      7,
  NOTES:            8,
  DELIVERY_STATUS:  9,
  ASSIGNED_TO:      10,
  REDO_COUNT:       11,
  UPLOADED_FILE_URL:12,
  INTERNAL_NOTES:   13,
  DUE_DATE:         14,
  COMPLETED_AT:     15,
  CREATED_AT:       16
};

// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CRM System')
    .addItem('Quotations', 'openQuotationsModule')
    .addToUi();
}

function openQuotationsModule() {
  const html = HtmlService.createHtmlOutputFromFile('quotation_form')
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
    recordId,
    recordLabel,
    accountId || '',
    accountName || '',
    fieldChanged,
    oldValue !== undefined && oldValue !== null ? String(oldValue) : '',
    newValue  !== undefined && newValue  !== null ? String(newValue)  : '',
    Session.getActiveUser().getEmail(),
    new Date(),
    notes || ''
  ]);
}

// ============================================================
//  ACCOUNTS
// ============================================================
function getAccounts() {
  const data = getSheet('Accounts').getDataRange().getValues();
  return data.slice(1)
    .filter(row => row[0])
    .map(row => ({ id: row[0], name: row[1] }));
}

function getAccountRow(accountId) {
  const sheet = getSheet('Accounts');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(accountId)) {
      return { row: i + 1, data: data[i] };
    }
  }
  return null;
}

function promoteAccountToClient(accountId, accountName, quotationId) {
  const sheet = getSheet('Accounts');
  const found = getAccountRow(accountId);
  if (!found) return;

  const { row, data } = found;
  const currentStatus = data[7];
  if (currentStatus === 'Client') return;

  const folderUrl = sheet.getRange(row, 11).getRichTextValue().getLinkUrl();
  if (folderUrl) {
    try {
      const folderId       = folderUrl.match(/[-\w]{25,}/)[0];
      const folder         = DriveApp.getFolderById(folderId);
      const accountsFolder = DriveApp.getFolderById(ACCOUNTS_FOLDER_ID);
      accountsFolder.addFolder(folder);
      DriveApp.getFolderById(GLOBAL_PIPELINE_FOLDER_ID).removeFolder(folder);
    } catch(e) {
      Logger.log('Folder move error: ' + e);
    }
  }

  sheet.getRange(row, 8).setValue('Client');
}

// ============================================================
//  SETTINGS
// ============================================================
function getSettings() {
  const sheet = getSheet('Quotation Settings');
  if (!sheet) return { currencies: [], pricingModes: [] };

  const lastRow = sheet.getLastRow();

  const currencies = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String)
    : [];

  const pricingModes = lastRow >= 2
    ? sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().filter(String)
    : [];

  return { currencies, pricingModes };
}

// ============================================================
//  ITEMS
// ============================================================
function getItems() {
  const data = getSheet('Items').getDataRange().getValues();
  return data.slice(1)
    .filter(row => row[0])
    .map(row => ({ name: row[0], description: row[1], price: row[2] }));
}

// ============================================================
//  QUOTATIONS — GENERATE ID
// ============================================================
function generateQuotationId() {
  const sheet = getSheet('Quotations');
  const data  = sheet.getDataRange().getValues();
  let maxNum  = 0;
  data.slice(1).forEach(row => {
    const qId   = String(row[Q.QUOTATION_ID] || '');
    const match = qId.match(/^Q-(\d+)$/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  });
  return 'Q-' + Utilities.formatString('%04d', maxNum + 1);
}

// ============================================================
//  QUOTATIONS — LIST
// ============================================================
function getQuotations() {
  const sheet = getSheet('Quotations');
  const data  = sheet.getDataRange().getValues();
  const tz    = Session.getScriptTimeZone();

  return data.slice(1)
    .filter(row => row[Q.QUOTATION_ID])
    .map(row => ({
      id:          String(row[Q.QUOTATION_ID]),
      version:     row[Q.CURRENT_VERSION] || 1,
      accountId:   row[Q.ACCOUNT_ID],
      accountName: row[Q.ACCOUNT_NAME],
      projectName: row[Q.PROJECT_NAME],
      dateIssued:  row[Q.DATE_ISSUED]
        ? Utilities.formatDate(new Date(row[Q.DATE_ISSUED]), tz, 'yyyy-MM-dd') : '',
      total:       row[Q.TOTAL],
      currency:    row[Q.CURRENCY],
      status:      row[Q.STATUS]
    }))
    .sort((a, b) => a.id.localeCompare(b.id));
}

// ============================================================
//  QUOTATIONS — GET SINGLE
// ============================================================
function getQuotationById(quotationId) {
  const qSheet = getSheet('Quotations');
  const iSheet = getSheet('Quote_Items');
  const qData  = qSheet.getDataRange().getValues();
  const iData  = iSheet.getDataRange().getValues();
  const tz     = Session.getScriptTimeZone();

  let qRow = null;
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      qRow = qData[i]; break;
    }
  }
  if (!qRow) return null;

  const items = iData.slice(1)
    .filter(r => String(r[QI.QUOTATION_ID]) === String(quotationId)
              && r[QI.STATUS] !== 'Deleted')
    .sort((a, b) => a[QI.ITEM_INDEX] - b[QI.ITEM_INDEX])
    .map(r => ({
      itemId:      r[QI.ITEM_ID],
      name:        r[QI.ITEM_NAME],
      quantity:    r[QI.QUANTITY],
      description: r[QI.DESCRIPTION],
      notes:       r[QI.NOTES],
      unitPrice:   r[QI.UNIT_PRICE],
      subtotal:    r[QI.SUBTOTAL]
    }));

  return {
    id:                 String(qRow[Q.QUOTATION_ID]),
    version:            qRow[Q.CURRENT_VERSION] || 1,
    accountId:          qRow[Q.ACCOUNT_ID],
    accountName:        qRow[Q.ACCOUNT_NAME],
    projectName:        qRow[Q.PROJECT_NAME],
    projectDescription: qRow[Q.PROJECT_DESC],
    dateIssued:         qRow[Q.DATE_ISSUED]
      ? Utilities.formatDate(new Date(qRow[Q.DATE_ISSUED]), tz, 'yyyy-MM-dd') : '',
    minDays:            qRow[Q.MIN_DAYS],
    maxDays:            qRow[Q.MAX_DAYS],
    deliveryDeadline:   qRow[Q.DELIVERY_DEADLINE]
      ? Utilities.formatDate(new Date(qRow[Q.DELIVERY_DEADLINE]), tz, 'yyyy-MM-dd') : '',
    pricingMode:        qRow[Q.PRICING_MODE],
    currency:           qRow[Q.CURRENCY],
    subtotal:           qRow[Q.SUBTOTAL],
    discounted:         qRow[Q.DISCOUNTED],
    discountPercent:    qRow[Q.DISCOUNT_PERCENT],
    discountAmount:     qRow[Q.DISCOUNT_AMOUNT],
    taxed:              qRow[Q.TAXED],
    taxPercent:         qRow[Q.TAX_PERCENT],
    taxAmount:          qRow[Q.TAX_AMOUNT],
    total:              qRow[Q.TOTAL],
    status:             qRow[Q.STATUS],
    notes:              qRow[Q.NOTES],
    folderUrl:          qRow[Q.FOLDER_URL] || '',
    items
  };
}

// ============================================================
//  QUOTATIONS — GET VERSION HISTORY (from log)
// ============================================================
function getQuotationHistory(quotationId) {
  const sheet = getSheet('Quotations_Log');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const tz   = Session.getScriptTimeZone();

  return data.slice(1)
    .filter(row => String(row[2]) === String(quotationId)
                && row[6] === 'Version')
    .map(row => ({
      version:   row[8],
      changedBy: row[9],
      changedAt: row[10]
        ? Utilities.formatDate(new Date(row[10]), tz, 'dd MMM yyyy HH:mm') : '',
      folderUrl: row[11] || ''
    }))
    .sort((a, b) => b.version - a.version);
}

// ============================================================
//  QUOTATIONS — CREATE
// ============================================================
function createQuotation(data) {
  const ss            = SpreadsheetApp.getActive();
  const qSheet        = ss.getSheetByName('Quotations');
  const iSheet        = ss.getSheetByName('Quote_Items');
  const accountsSheet = ss.getSheetByName('Accounts');
  const brandingSheet = ss.getSheetByName('Branding');

  const quotationId = generateQuotationId();
  const now         = new Date();
  const user        = Session.getActiveUser().getEmail();
  const version     = 1;
  const tz          = Session.getScriptTimeZone();

  // Get account details
  const accounts = accountsSheet.getDataRange().getValues();
  let accountName = '', accountFolderUrl = '', accountRowIndex = -1;

  for (let i = 1; i < accounts.length; i++) {
    if (String(accounts[i][0]) === String(data.accountId)) {
      accountName      = accounts[i][1];
      accountFolderUrl = accountsSheet.getRange(i + 1, 11).getRichTextValue().getLinkUrl();
      accountRowIndex  = i + 1;
      break;
    }
  }

  // Get or create Pipeline folder
  let pipelineFolder;
  if (accountFolderUrl) {
    const accountFolderId = accountFolderUrl.match(/[-\w]{25,}/)[0];
    const accountFolder   = DriveApp.getFolderById(accountFolderId);
    const existing        = accountFolder.getFoldersByName('Pipeline');
    pipelineFolder        = existing.hasNext()
      ? existing.next()
      : accountFolder.createFolder('Pipeline');
  } else {
    const globalPipeline = DriveApp.getFolderById(GLOBAL_PIPELINE_FOLDER_ID);
    const accountFolder  = globalPipeline.createFolder(accountName);
    pipelineFolder       = accountFolder.createFolder('Pipeline');
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(accountName)
      .setLinkUrl('https://drive.google.com/drive/folders/' + accountFolder.getId())
      .build();
    accountsSheet.getRange(accountRowIndex, 11).setRichTextValue(richText);
  }

  // Calculations
  let subtotal = data.subtotal;
  if (data.pricingMode === 'Itemized') {
    subtotal = data.items.reduce((sum, item) => sum + (item.subtotal || 0), 0);
  }
  const discountAmount     = data.discounted ? subtotal * (data.discountPercent / 100) : 0;
  const discountedSubtotal = subtotal - discountAmount;
  const taxAmount          = data.taxed ? discountedSubtotal * (data.taxPercent / 100) : 0;
  const total              = discountedSubtotal + taxAmount;

  // Generate PDF
  const branding = brandingSheet.getDataRange().getValues()[1] || [];
  const html     = generateQuotationHTML({
    quotationId, accountName, ...data,
    subtotal, taxAmount, discountAmount, total,
    logoUrl: branding[0] || '', companyName: branding[1] || '', version
  });
  const blob    = Utilities.newBlob(html, 'text/html', 'quotation.html')
    .getAs('application/pdf')
    .setName(`${quotationId} - ${data.projectName} - v${version}.pdf`);
  const pdfFile = pipelineFolder.createFile(blob);
  const pdfUrl  = pdfFile.getUrl();

  // Write quotation row
  qSheet.appendRow([
    quotationId, version, data.accountId, accountName,
    data.projectName, data.projectDescription,
    data.dateIssued, data.minDays, data.maxDays, data.deliveryDeadline,
    data.pricingMode, data.currency,
    subtotal, data.discounted, data.discountPercent, discountAmount,
    data.taxed, data.taxPercent, taxAmount,
    total, 'Drafted', data.notes, pdfUrl,
    user, now, user, now
  ]);

  // Write items
  data.items.forEach((item, index) => {
    iSheet.appendRow([
      Utilities.getUuid(), quotationId, index + 1,
      item.name, item.quantity, item.description, item.notes,
      data.pricingMode === 'Itemized' ? item.unitPrice : '',
      data.pricingMode === 'Itemized' ? item.subtotal  : '',
      'Active', user, now
    ]);
  });

  // Log creation
  writeLog('Quotations_Log', 'Quotations', quotationId,
    `${quotationId} — ${data.projectName}`,
    data.accountId, accountName,
    'Version', '', version, 'Quotation created');

  return { success: true, quotationId };
}

// ============================================================
//  QUOTATIONS — EDIT (in-place update, new version)
// ============================================================
function editQuotation(data) {
  const ss            = SpreadsheetApp.getActive();
  const qSheet        = ss.getSheetByName('Quotations');
  const iSheet        = ss.getSheetByName('Quote_Items');
  const accountsSheet = ss.getSheetByName('Accounts');
  const brandingSheet = ss.getSheetByName('Branding');

  const now  = new Date();
  const user = Session.getActiveUser().getEmail();

  // Find quotation row
  const qData = qSheet.getDataRange().getValues();
  let qRowIndex = -1, currentVersion = 1, currentPdfUrl = '';

  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(data.id)) {
      qRowIndex      = i + 1;
      currentVersion = qData[i][Q.CURRENT_VERSION] || 1;
      currentPdfUrl  = qData[i][Q.FOLDER_URL] || '';
      break;
    }
  }
  if (qRowIndex === -1) return { success: false, error: 'Quotation not found.' };

  const newVersion = Number(currentVersion) + 1;

  // Trash old PDF
  if (currentPdfUrl) {
    try {
      const fileId = currentPdfUrl.match(/[-\w]{25,}/)[0];
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch(e) { Logger.log('Old PDF delete: ' + e); }
  }

  // Get account details
  const accounts = accountsSheet.getDataRange().getValues();
  let accountName = '', pipelineFolder = null;

  for (let i = 1; i < accounts.length; i++) {
    if (String(accounts[i][0]) === String(data.accountId)) {
      accountName = accounts[i][1];
      const folderUrl = accountsSheet.getRange(i + 1, 11).getRichTextValue().getLinkUrl();
      if (folderUrl) {
        const accountFolderId = folderUrl.match(/[-\w]{25,}/)[0];
        const accountFolder   = DriveApp.getFolderById(accountFolderId);
        const existing        = accountFolder.getFoldersByName('Pipeline');
        pipelineFolder        = existing.hasNext()
          ? existing.next()
          : accountFolder.createFolder('Pipeline');
      }
      break;
    }
  }

  // Calculations
  let subtotal = data.subtotal;
  if (data.pricingMode === 'Itemized') {
    subtotal = data.items.reduce((sum, item) => sum + (item.subtotal || 0), 0);
  }
  const discountAmount     = data.discounted ? subtotal * (data.discountPercent / 100) : 0;
  const discountedSubtotal = subtotal - discountAmount;
  const taxAmount          = data.taxed ? discountedSubtotal * (data.taxPercent / 100) : 0;
  const total              = discountedSubtotal + taxAmount;

  // New PDF
  const branding = brandingSheet.getDataRange().getValues()[1] || [];
  const html     = generateQuotationHTML({
    quotationId: data.id, accountName, ...data,
    subtotal, taxAmount, discountAmount, total,
    logoUrl: branding[0] || '', companyName: branding[1] || '',
    version: newVersion
  });
  const blob    = Utilities.newBlob(html, 'text/html', 'quotation.html')
    .getAs('application/pdf')
    .setName(`${data.id} - ${data.projectName} - v${newVersion}.pdf`);
  const pdfFile = pipelineFolder ? pipelineFolder.createFile(blob) : null;
  const pdfUrl  = pdfFile ? pdfFile.getUrl() : '';

  // Update quotation row in place
  const qRange = qSheet.getRange(qRowIndex, 1, 1, 27);
  qRange.setValues([[
    data.id, newVersion, data.accountId, accountName,
    data.projectName, data.projectDescription,
    data.dateIssued, data.minDays, data.maxDays, data.deliveryDeadline,
    data.pricingMode, data.currency,
    subtotal, data.discounted, data.discountPercent, discountAmount,
    data.taxed, data.taxPercent, taxAmount,
    total, 'Drafted', data.notes, pdfUrl,
    qData[qRowIndex - 1][Q.CREATED_BY],
    qData[qRowIndex - 1][Q.CREATED_AT],
    user, now
  ]]);

  // Snapshot old items for log
  const iData     = iSheet.getDataRange().getValues();
  const oldItems  = iData.slice(1)
    .filter(r => String(r[QI.QUOTATION_ID]) === String(data.id)
              && r[QI.STATUS] !== 'Deleted')
    .map(r => ({
      name: r[QI.ITEM_NAME], qty: r[QI.QUANTITY],
      desc: r[QI.DESCRIPTION], notes: r[QI.NOTES],
      unitPrice: r[QI.UNIT_PRICE], subtotal: r[QI.SUBTOTAL]
    }));

  // Update items in place — match by ItemIndex, add new, remove extras
  const existingItemRows = [];
  for (let i = 1; i < iData.length; i++) {
    if (String(iData[i][QI.QUOTATION_ID]) === String(data.id)
     && iData[i][QI.STATUS] !== 'Deleted') {
      existingItemRows.push({ sheetRow: i + 1, data: iData[i] });
    }
  }

  // Update or create item rows
  data.items.forEach((item, index) => {
    const itemIndex  = index + 1;
    const existing   = existingItemRows.find(r => r.data[QI.ITEM_INDEX] === itemIndex);
    const unitPrice  = data.pricingMode === 'Itemized' ? item.unitPrice : '';
    const itemSubtotal = data.pricingMode === 'Itemized' ? item.subtotal : '';

    if (existing) {
      iSheet.getRange(existing.sheetRow, 1, 1, 12).setValues([[
        existing.data[QI.ITEM_ID], data.id, itemIndex,
        item.name, item.quantity, item.description, item.notes,
        unitPrice, itemSubtotal, 'Active', user, now
      ]]);
    } else {
      iSheet.appendRow([
        Utilities.getUuid(), data.id, itemIndex,
        item.name, item.quantity, item.description, item.notes,
        unitPrice, itemSubtotal, 'Active', user, now
      ]);
    }
  });

  // Mark extra item rows as Deleted if quotation now has fewer items
  if (existingItemRows.length > data.items.length) {
    for (let i = data.items.length; i < existingItemRows.length; i++) {
      iSheet.getRange(existingItemRows[i].sheetRow, QI.STATUS + 1).setValue('Deleted');
      iSheet.getRange(existingItemRows[i].sheetRow, QI.LAST_UPDATED_BY + 1).setValue(user);
      iSheet.getRange(existingItemRows[i].sheetRow, QI.LAST_UPDATED_AT + 1).setValue(now);
    }
  }

  // Log version bump + items snapshot
  writeLog('Quotations_Log', 'Quotations', data.id,
    `${data.id} — ${data.projectName}`,
    data.accountId, accountName,
    'Version', currentVersion, newVersion, '');

  writeLog('Quotations_Log', 'Quote_Items', data.id,
    `${data.id} — ${data.projectName}`,
    data.accountId, accountName,
    'Items',
    JSON.stringify(oldItems),
    JSON.stringify(data.items.map(i => ({
      name: i.name, qty: i.quantity,
      desc: i.description, notes: i.notes,
      unitPrice: i.unitPrice, subtotal: i.subtotal
    }))),
    `Items updated on v${newVersion}`);

  return { success: true, quotationId: data.id, version: newVersion };
}

// ============================================================
//  QUOTATIONS — DELETE
// ============================================================
function deleteQuotation(quotationId) {
  const qSheet = getSheet('Quotations');
  const qData  = qSheet.getDataRange().getValues();
  const user   = Session.getActiveUser().getEmail();

  let qRowIndex = -1, pdfUrl = '', accountId = '',
      accountName = '', projectName = '';

  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      if (qData[i][Q.STATUS] === 'Confirmed') {
        return { success: false, error: 'Cannot delete a confirmed quotation.' };
      }
      qRowIndex   = i + 1;
      pdfUrl      = qData[i][Q.FOLDER_URL] || '';
      accountId   = qData[i][Q.ACCOUNT_ID];
      accountName = qData[i][Q.ACCOUNT_NAME];
      projectName = qData[i][Q.PROJECT_NAME];
      break;
    }
  }
  if (qRowIndex === -1) return { success: false, error: 'Quotation not found.' };

  // Trash PDF
  if (pdfUrl) {
    try {
      const fileId = pdfUrl.match(/[-\w]{25,}/)[0];
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch(e) { Logger.log('PDF delete: ' + e); }
  }

  // Delete quotation row
  qSheet.deleteRow(qRowIndex);

  // Delete items
  const iSheet = getSheet('Quote_Items');
  const iData  = iSheet.getDataRange().getValues();
  for (let i = iData.length - 1; i >= 1; i--) {
    if (String(iData[i][QI.QUOTATION_ID]) === String(quotationId)) {
      iSheet.deleteRow(i + 1);
    }
  }

  // Log deletion
  writeLog('Quotations_Log', 'Quotations', quotationId,
    `${quotationId} — ${projectName}`,
    accountId, accountName,
    'Status', 'Drafted', 'Deleted', 'Quotation deleted');

  return { success: true };
}

// ============================================================
//  QUOTATIONS — CONFIRM
// ============================================================
function confirmQuotation(quotationId) {
  const ss            = SpreadsheetApp.getActive();
  const qSheet        = ss.getSheetByName('Quotations');
  const iSheet        = ss.getSheetByName('Quote_Items');
  const accountsSheet = ss.getSheetByName('Accounts');
  const pSheet        = ss.getSheetByName('Projects');
  const piSheet       = ss.getSheetByName('Project_Items');

  const qData  = qSheet.getDataRange().getValues();
  const user   = Session.getActiveUser().getEmail();
  const now    = new Date();
  const tz     = Session.getScriptTimeZone();

  let qRowIndex = -1, qRow = null;
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      qRowIndex = i + 1;
      qRow      = qData[i];
      break;
    }
  }
  if (!qRow) return { success: false, error: 'Quotation not found.' };
  if (qRow[Q.STATUS] === 'Confirmed') {
    return { success: false, error: 'Quotation is already confirmed.' };
  }

  const accountId      = qRow[Q.ACCOUNT_ID];
  const accountName    = qRow[Q.ACCOUNT_NAME];
  const projectName    = qRow[Q.PROJECT_NAME];
  const projectDesc    = qRow[Q.PROJECT_DESC];
  const pdfUrl         = qRow[Q.FOLDER_URL] || '';
  const deliveryDdl    = qRow[Q.DELIVERY_DEADLINE];
  const dateIssued     = qRow[Q.DATE_ISSUED];
  const minDays        = qRow[Q.MIN_DAYS];

  // Calculate DueDate = earlier of (DeliveryDeadline - 1) or (DateIssued + MinDays)
  let dueDate = null;
  try {
    const ddl      = new Date(deliveryDdl);
    const ddlMinus1 = new Date(ddl); ddlMinus1.setDate(ddl.getDate() - 1);
    const issued   = new Date(dateIssued);
    const minTarget = new Date(issued); minTarget.setDate(issued.getDate() + Number(minDays));
    dueDate = ddlMinus1 < minTarget ? ddlMinus1 : minTarget;
  } catch(e) { Logger.log('DueDate calc: ' + e); }

  // Get account folder URL
  const accounts = accountsSheet.getDataRange().getValues();
  let accountFolderUrl = '';
  for (let i = 1; i < accounts.length; i++) {
    if (String(accounts[i][0]) === String(accountId)) {
      accountFolderUrl = accountsSheet.getRange(i + 1, 11).getRichTextValue().getLinkUrl();
      break;
    }
  }

  // Move PDF: Pipeline → Q-XXXX folder, create Deliverables subfolder
  if (pdfUrl && accountFolderUrl) {
    try {
      const pdfFileId       = pdfUrl.match(/[-\w]{25,}/)[0];
      const pdfFile         = DriveApp.getFileById(pdfFileId);
      const accountFolderId = accountFolderUrl.match(/[-\w]{25,}/)[0];
      const accountFolder   = DriveApp.getFolderById(accountFolderId);
      const qFolder         = accountFolder.createFolder(quotationId);
      qFolder.addFile(pdfFile);
      qFolder.createFolder('Deliverables');
      const pipelineIt = accountFolder.getFoldersByName('Pipeline');
      if (pipelineIt.hasNext()) pipelineIt.next().removeFile(pdfFile);
    } catch(e) { Logger.log('PDF move error: ' + e); }
  }

  // Update quotation status in place
  qSheet.getRange(qRowIndex, Q.STATUS + 1).setValue('Confirmed');
  qSheet.getRange(qRowIndex, Q.LAST_UPDATED_BY + 1).setValue(user);
  qSheet.getRange(qRowIndex, Q.LAST_UPDATED_AT + 1).setValue(now);

  // Update Quote_Items to Approved
  const iData = iSheet.getDataRange().getValues();
  for (let i = 1; i < iData.length; i++) {
    if (String(iData[i][QI.QUOTATION_ID]) === String(quotationId)
     && iData[i][QI.STATUS] === 'Active') {
      iSheet.getRange(i + 1, QI.STATUS + 1).setValue('Approved');
      iSheet.getRange(i + 1, QI.LAST_UPDATED_BY + 1).setValue(user);
      iSheet.getRange(i + 1, QI.LAST_UPDATED_AT + 1).setValue(now);
    }
  }

  // Create Projects row
  const projectId = Utilities.getUuid();
  pSheet.appendRow([
    projectId, quotationId, accountId, accountName,
    projectName, projectDesc,
    deliveryDdl, dueDate,
    'Active', '', now, ''
  ]);

  // Create Project_Items rows from approved Quote_Items
  const approvedItems = iData.slice(1).filter(r =>
    String(r[QI.QUOTATION_ID]) === String(quotationId)
    && r[QI.STATUS] === 'Approved'
  );

  approvedItems.forEach(item => {
    piSheet.appendRow([
      Utilities.getUuid(), projectId, quotationId,
      accountName, projectName,
      item[QI.ITEM_NAME], item[QI.QUANTITY],
      item[QI.DESCRIPTION], item[QI.NOTES],
      'Pending', '', 0, '', '',
      dueDate, '', now
    ]);
  });

  // Log confirmation
  writeLog('Quotations_Log', 'Quotations', quotationId,
    `${quotationId} — ${projectName}`,
    accountId, accountName,
    'Status', 'Drafted', 'Confirmed', 'Quotation confirmed');

  writeLog('Projects_Log', 'Projects', projectId,
    `${quotationId} — ${projectName}`,
    accountId, accountName,
    'Status', '', 'Active', 'Project created on confirmation');

  // Promote account to Client
  promoteAccountToClient(accountId, accountName, quotationId);

  return { success: true };
}

// ============================================================
//  PDF GENERATION
// ============================================================
function generateQuotationHTML(data) {
  const showPricing = data.pricingMode === 'Itemized';
  const itemsRows   = (data.items || []).map((item, i) => `
    <tr>
      <td>${i + 1}</td><td>${item.name}</td><td>${item.quantity}</td>
      <td>${item.description || ''}</td>
      ${showPricing ? `<td>${item.unitPrice}</td><td>${item.subtotal}</td>` : ''}
    </tr>`).join('');

  return `<html><body style="font-family:Arial;padding:20px;">
    <h2>${data.companyName || 'Quotation'}</h2>
    ${data.logoUrl ? `<img src="${data.logoUrl}" height="60"/>` : ''}
    <h3>Quotation #${data.quotationId} — Version ${data.version}</h3>
    <p><b>Account:</b> ${data.accountName}</p>
    <p><b>Project:</b> ${data.projectName}</p>
    <p><b>Description:</b> ${data.projectDescription || ''}</p>
    <p><b>Date Issued:</b> ${data.dateIssued}</p>
    <p><b>Delivery:</b> ${data.minDays}–${data.maxDays} days</p>
    <table border="1" cellspacing="0" cellpadding="5" width="100%">
      <tr><th>#</th><th>Item</th><th>Qty</th><th>Description</th>
      ${showPricing ? '<th>Unit Price</th><th>Subtotal</th>' : ''}</tr>
      ${itemsRows}
    </table>
    <h3>Summary</h3>
    <p>Subtotal: ${data.subtotal} ${data.currency}</p>
    ${data.discounted ? `<p>Discount (${data.discountPercent}%): -${data.discountAmount}</p>` : ''}
    ${data.taxed      ? `<p>Tax (${data.taxPercent}%): +${data.taxAmount}</p>` : ''}
    <h2>Total: ${data.total} ${data.currency}</h2>
  </body></html>`;
}
