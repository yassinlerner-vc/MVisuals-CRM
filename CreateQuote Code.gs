// ============================================================
//  CONSTANTS
// ============================================================
const ACCOUNTS_FOLDER_ID        = '19iNsbyldpjF8pVuc9WUGIHB8bibYbu0J';
const GLOBAL_PIPELINE_FOLDER_ID = '1DbnsDxGgymXsFMnCK8uUIUP0UFXuMv3z';

// Column indexes for Quotations sheet (0-based)
const Q_COL = {
  ROW_ID:           0,
  QUOTATION_ID:     1,
  VERSION:          2,
  ACCOUNT_ID:       3,
  ACCOUNT_NAME:     4,
  DATE_ISSUED:      5,
  PROJECT_NAME:     6,
  PROJECT_DESC:     7,
  MIN_DAYS:         8,
  MAX_DAYS:         9,
  DELIVERY:         10,
  PRICING_MODE:     11,
  CURRENCY:         12,
  SUBTOTAL:         13,
  TAXED:            14,
  TAX_PERCENT:      15,
  TAX_AMOUNT:       16,
  DISCOUNTED:       17,
  DISCOUNT_PERCENT: 18,
  DISCOUNT_AMOUNT:  19,
  TOTAL:            20,
  STATUS:           21,
  NOTES:            22,
  CREATED_AT:       23,
  CREATED_BY:       24,
  FOLDER_URL:       25
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
//  SHEET HELPERS
// ============================================================
function getSheet(name) {
  return SpreadsheetApp.getActive().getSheetByName(name);
}

// ============================================================
//  ACCOUNTS
// ============================================================
function getAccounts() {
  const data = getSheet('Accounts').getDataRange().getValues();
  return data.slice(1).map(row => ({ id: row[0], name: row[1] }));
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

function promoteAccountToClient(accountId) {
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
//  ITEMS & SETTINGS
// ============================================================
function getItems() {
  const data = getSheet('Items').getDataRange().getValues();
  return data.slice(1).map(row => ({
    name: row[0], description: row[1], price: row[2]
  }));
}

function getSettings() {
  const sheet = getSheet('Settings');
  return {
    currencies:   sheet.getRange('A2:A20').getValues().flat().filter(String),
    pricingModes: sheet.getRange('D2:D20').getValues().flat().filter(String)
  };
}

// ============================================================
//  QUOTATIONS — LIST
// ============================================================
function getQuotations() {
  const sheet = getSheet('Quotations');
  const data  = sheet.getDataRange().getValues();
  const tz    = Session.getScriptTimeZone();

  const map = {};
  data.slice(1).forEach(row => {
    const qId    = String(row[Q_COL.QUOTATION_ID]);
    const status = String(row[Q_COL.STATUS]);
    const ver    = Number(row[Q_COL.VERSION]) || 1;
    if (!qId || status === 'Deleted') return;
    if (!map[qId] || ver > map[qId].ver) {
      map[qId] = {
        rowId:       row[Q_COL.ROW_ID],
        id:          qId,
        accountId:   row[Q_COL.ACCOUNT_ID],
        accountName: row[Q_COL.ACCOUNT_NAME],
        dateIssued:  row[Q_COL.DATE_ISSUED]
          ? Utilities.formatDate(new Date(row[Q_COL.DATE_ISSUED]), tz, 'yyyy-MM-dd') : '',
        projectName: row[Q_COL.PROJECT_NAME],
        total:       row[Q_COL.TOTAL],
        currency:    row[Q_COL.CURRENCY],
        status,
        ver
      };
    }
  });

  return Object.values(map).sort((a, b) => a.id.localeCompare(b.id));
}

// ============================================================
//  QUOTATIONS — GET VERSION HISTORY
// ============================================================
function getQuotationHistory(quotationId) {
  const sheet = getSheet('Quotations');
  const data  = sheet.getDataRange().getValues();
  const tz    = Session.getScriptTimeZone();

  const rows = [];
  data.slice(1).forEach(row => {
    if (String(row[Q_COL.QUOTATION_ID]) !== String(quotationId)) return;
    rows.push({
      version:   Number(row[Q_COL.VERSION]) || 1,
      status:    String(row[Q_COL.STATUS]),
      createdAt: row[Q_COL.CREATED_AT]
        ? Utilities.formatDate(new Date(row[Q_COL.CREATED_AT]), tz, 'dd MMM yyyy') : '',
      folderUrl: row[Q_COL.FOLDER_URL] || ''
    });
  });

  return rows.sort((a, b) => b.version - a.version);
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

  let best = null;
  for (let i = 1; i < qData.length; i++) {
    const row    = qData[i];
    const qId    = String(row[Q_COL.QUOTATION_ID]);
    const status = String(row[Q_COL.STATUS]);
    const ver    = Number(row[Q_COL.VERSION]) || 1;
    if (qId !== String(quotationId) || status === 'Deleted') continue;
    if (!best || ver > Number(best[Q_COL.VERSION])) best = row;
  }
  if (!best) return null;

  const items = iData.slice(1)
    .filter(r => String(r[1]) === String(quotationId) && r[9] !== 'Deleted')
    .map(r => ({
      name: r[3], quantity: r[4], description: r[5],
      notes: r[6], unitPrice: r[7], subtotal: r[8]
    }));

  return {
    rowId:              best[Q_COL.ROW_ID],
    id:                 best[Q_COL.QUOTATION_ID],
    accountId:          best[Q_COL.ACCOUNT_ID],
    accountName:        best[Q_COL.ACCOUNT_NAME],
    dateIssued:         best[Q_COL.DATE_ISSUED]
      ? Utilities.formatDate(new Date(best[Q_COL.DATE_ISSUED]), tz, 'yyyy-MM-dd') : '',
    projectName:        best[Q_COL.PROJECT_NAME],
    projectDescription: best[Q_COL.PROJECT_DESC],
    minDays:            best[Q_COL.MIN_DAYS],
    maxDays:            best[Q_COL.MAX_DAYS],
    deliveryDeadline:   best[Q_COL.DELIVERY]
      ? Utilities.formatDate(new Date(best[Q_COL.DELIVERY]), tz, 'yyyy-MM-dd') : '',
    pricingMode:        best[Q_COL.PRICING_MODE],
    currency:           best[Q_COL.CURRENCY],
    subtotal:           best[Q_COL.SUBTOTAL],
    taxed:              best[Q_COL.TAXED],
    taxPercent:         best[Q_COL.TAX_PERCENT],
    taxAmount:          best[Q_COL.TAX_AMOUNT],
    discounted:         best[Q_COL.DISCOUNTED],
    discountPercent:    best[Q_COL.DISCOUNT_PERCENT],
    discountAmount:     best[Q_COL.DISCOUNT_AMOUNT],
    total:              best[Q_COL.TOTAL],
    status:             best[Q_COL.STATUS],
    notes:              best[Q_COL.NOTES],
    version:            best[Q_COL.VERSION] || 1,
    folderUrl:          best[Q_COL.FOLDER_URL] || '',
    items
  };
}

// ============================================================
//  QUOTATIONS — GENERATE ID (max + 1, gap-safe)
// ============================================================
function generateQuotationId() {
  const sheet  = getSheet('Quotations');
  const data   = sheet.getDataRange().getValues();
  let maxNum   = 0;
  data.slice(1).forEach(row => {
    const qId   = String(row[Q_COL.QUOTATION_ID] || '');
    const match = qId.match(/^Q-(\d+)$/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  });
  return 'Q-' + Utilities.formatString('%04d', maxNum + 1);
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

  const rowId       = Utilities.getUuid();
  const quotationId = generateQuotationId();
  const now         = new Date();
  const user        = Session.getActiveUser().getEmail();
  const version     = 1;

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

  let subtotal = data.subtotal;
  if (data.pricingMode === 'Itemized') {
    subtotal = data.items.reduce((sum, item) => sum + (item.subtotal || 0), 0);
  }
  const discountAmount     = data.discounted ? subtotal * (data.discountPercent / 100) : 0;
  const discountedSubtotal = subtotal - discountAmount;
  const taxAmount          = data.taxed ? discountedSubtotal * (data.taxPercent / 100) : 0;
  const total              = discountedSubtotal + taxAmount;

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

  qSheet.appendRow([
    rowId, quotationId, version, data.accountId, accountName,
    data.dateIssued, data.projectName, data.projectDescription,
    data.minDays, data.maxDays, data.deliveryDeadline,
    data.pricingMode, data.currency,
    subtotal, data.taxed, data.taxPercent, taxAmount,
    data.discounted, data.discountPercent, discountAmount,
    total, 'Drafted', data.notes, now, user, pdfUrl
  ]);

  data.items.forEach((item, index) => {
    iSheet.appendRow([
      Utilities.getUuid(), quotationId, index + 1,
      item.name, item.quantity, item.description, item.notes,
      data.pricingMode === 'Itemized' ? item.unitPrice : '',
      data.pricingMode === 'Itemized' ? item.subtotal  : '',
      'Active'
    ]);
  });

  return { success: true, quotationId };
}

// ============================================================
//  QUOTATIONS — EDIT
// ============================================================
function editQuotation(data) {
  const ss            = SpreadsheetApp.getActive();
  const qSheet        = ss.getSheetByName('Quotations');
  const iSheet        = ss.getSheetByName('Quote_Items');
  const accountsSheet = ss.getSheetByName('Accounts');
  const brandingSheet = ss.getSheetByName('Branding');

  const now  = new Date();
  const user = Session.getActiveUser().getEmail();

  const qData = qSheet.getDataRange().getValues();
  let currentRow = -1, currentVersion = 1, currentPdfUrl = '';

  for (let i = 1; i < qData.length; i++) {
    const row    = qData[i];
    const qId    = String(row[Q_COL.QUOTATION_ID]);
    const status = String(row[Q_COL.STATUS]);
    const ver    = Number(row[Q_COL.VERSION]) || 1;
    if (qId !== String(data.id) || status === 'Deleted') continue;
    if (currentRow === -1 || ver > currentVersion) {
      currentRow     = i + 1;
      currentVersion = ver;
      currentPdfUrl  = row[Q_COL.FOLDER_URL] || '';
    }
  }
  if (currentRow === -1) return { success: false, error: 'Quotation not found.' };

  const newVersion = currentVersion + 1;

  if (currentPdfUrl) {
    try {
      const fileId = currentPdfUrl.match(/[-\w]{25,}/)[0];
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch(e) { Logger.log('Old PDF delete: ' + e); }
  }

  qSheet.getRange(currentRow, Q_COL.STATUS + 1).setValue('Deleted');

  const iData = iSheet.getDataRange().getValues();
  for (let i = 1; i < iData.length; i++) {
    if (String(iData[i][1]) === String(data.id) && iData[i][9] !== 'Deleted') {
      iSheet.getRange(i + 1, 10).setValue('Deleted');
    }
  }

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

  let subtotal = data.subtotal;
  if (data.pricingMode === 'Itemized') {
    subtotal = data.items.reduce((sum, item) => sum + (item.subtotal || 0), 0);
  }
  const discountAmount     = data.discounted ? subtotal * (data.discountPercent / 100) : 0;
  const discountedSubtotal = subtotal - discountAmount;
  const taxAmount          = data.taxed ? discountedSubtotal * (data.taxPercent / 100) : 0;
  const total              = discountedSubtotal + taxAmount;

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

  qSheet.appendRow([
    Utilities.getUuid(), data.id, newVersion, data.accountId, accountName,
    data.dateIssued, data.projectName, data.projectDescription,
    data.minDays, data.maxDays, data.deliveryDeadline,
    data.pricingMode, data.currency,
    subtotal, data.taxed, data.taxPercent, taxAmount,
    data.discounted, data.discountPercent, discountAmount,
    total, 'Drafted', data.notes, now, user, pdfUrl
  ]);

  data.items.forEach((item, index) => {
    iSheet.appendRow([
      Utilities.getUuid(), data.id, index + 1,
      item.name, item.quantity, item.description, item.notes,
      data.pricingMode === 'Itemized' ? item.unitPrice : '',
      data.pricingMode === 'Itemized' ? item.subtotal  : '',
      'Active'
    ]);
  });

  return { success: true, quotationId: data.id, version: newVersion };
}

// ============================================================
//  QUOTATIONS — DELETE
// ============================================================
function deleteQuotation(quotationId) {
  const qSheet = getSheet('Quotations');
  const qData  = qSheet.getDataRange().getValues();

  for (let i = 1; i < qData.length; i++) {
    const row    = qData[i];
    const qId    = String(row[Q_COL.QUOTATION_ID]);
    const status = String(row[Q_COL.STATUS]);
    if (qId !== String(quotationId) || status === 'Deleted') continue;

    if (status === 'Confirmed') {
      return { success: false, error: 'Cannot delete a confirmed quotation.' };
    }

    const pdfUrl = row[Q_COL.FOLDER_URL] || '';
    if (pdfUrl) {
      try {
        const fileId = pdfUrl.match(/[-\w]{25,}/)[0];
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch(e) { Logger.log('PDF delete: ' + e); }
    }

    qSheet.getRange(i + 1, Q_COL.STATUS + 1).setValue('Deleted');
  }

  const iSheet = getSheet('Quote_Items');
  const iData  = iSheet.getDataRange().getValues();
  for (let i = 1; i < iData.length; i++) {
    if (String(iData[i][1]) === String(quotationId) && iData[i][9] !== 'Deleted') {
      iSheet.getRange(i + 1, 10).setValue('Deleted');
    }
  }

  return { success: true };
}

// ============================================================
//  QUOTATIONS — CONFIRM
// ============================================================
function confirmQuotation(quotationId) {
  const ss            = SpreadsheetApp.getActive();
  const qSheet        = ss.getSheetByName('Quotations');
  const accountsSheet = ss.getSheetByName('Accounts');
  const qData         = qSheet.getDataRange().getValues();

  let targetRow = -1, accountId = '', pdfUrl = '',
      accountFolderUrl = '', currentVersion = 0;

  for (let i = 1; i < qData.length; i++) {
    const row    = qData[i];
    const qId    = String(row[Q_COL.QUOTATION_ID]);
    const status = String(row[Q_COL.STATUS]);
    const ver    = Number(row[Q_COL.VERSION]) || 1;
    if (qId !== String(quotationId) || status === 'Deleted') continue;
    if (ver > currentVersion) {
      targetRow      = i + 1;
      currentVersion = ver;
      accountId      = row[Q_COL.ACCOUNT_ID];
      pdfUrl         = row[Q_COL.FOLDER_URL] || '';
    }
  }
  if (targetRow === -1) return { success: false, error: 'Quotation not found.' };

  const accounts = accountsSheet.getDataRange().getValues();
  for (let i = 1; i < accounts.length; i++) {
    if (String(accounts[i][0]) === String(accountId)) {
      accountFolderUrl = accountsSheet.getRange(i + 1, 11).getRichTextValue().getLinkUrl();
      break;
    }
  }

  if (pdfUrl && accountFolderUrl) {
    try {
      const pdfFileId       = pdfUrl.match(/[-\w]{25,}/)[0];
      const pdfFile         = DriveApp.getFileById(pdfFileId);
      const accountFolderId = accountFolderUrl.match(/[-\w]{25,}/)[0];
      const accountFolder   = DriveApp.getFolderById(accountFolderId);
      const qFolder         = accountFolder.createFolder(quotationId);
      qFolder.addFile(pdfFile);
      const pipelineIt = accountFolder.getFoldersByName('Pipeline');
      if (pipelineIt.hasNext()) pipelineIt.next().removeFile(pdfFile);
    } catch(e) { Logger.log('PDF move error: ' + e); }
  }

  qSheet.getRange(targetRow, Q_COL.STATUS + 1).setValue('Confirmed');
  promoteAccountToClient(accountId);

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
