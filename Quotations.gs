// ============================================================
//  FIELD-CHANGE LOG HELPER
//  Compares old row array vs new data object and logs each change
// ============================================================
function logQuotationFieldChanges(oldRow, newData, accountId, accountName) {
  const tz    = Session.getScriptTimeZone();
  const qId   = String(oldRow[Q.QUOTATION_ID]);
  const label = qId + ' — ' + (newData.projectName || oldRow[Q.PROJECT_NAME]);

  const fields = [
    { key: 'ProjectName',        old: oldRow[Q.PROJECT_NAME],      nw: newData.projectName },
    { key: 'ProjectDescription', old: oldRow[Q.PROJECT_DESC],      nw: newData.projectDescription },
    { key: 'DateIssued',
      old: oldRow[Q.DATE_ISSUED]
        ? Utilities.formatDate(new Date(oldRow[Q.DATE_ISSUED]), tz, 'yyyy-MM-dd') : '',
      nw: newData.dateIssued },
    { key: 'MinDays',            old: oldRow[Q.MIN_DAYS],          nw: newData.minDays },
    { key: 'MaxDays',            old: oldRow[Q.MAX_DAYS],          nw: newData.maxDays },
    { key: 'DeliveryDeadline',
      old: oldRow[Q.DELIVERY_DEADLINE]
        ? Utilities.formatDate(new Date(oldRow[Q.DELIVERY_DEADLINE]), tz, 'yyyy-MM-dd') : '',
      nw: newData.deliveryDeadline },
    { key: 'PricingMode',        old: oldRow[Q.PRICING_MODE],      nw: newData.pricingMode },
    { key: 'Currency',           old: oldRow[Q.CURRENCY],          nw: newData.currency },
    { key: 'Subtotal',           old: oldRow[Q.SUBTOTAL],          nw: newData.subtotal },
    { key: 'Discounted',         old: oldRow[Q.DISCOUNTED],        nw: newData.discounted },
    { key: 'DiscountPercent',    old: oldRow[Q.DISCOUNT_PERCENT],  nw: newData.discountPercent },
    { key: 'Taxed',              old: oldRow[Q.TAXED],             nw: newData.taxed },
    { key: 'TaxPercent',         old: oldRow[Q.TAX_PERCENT],       nw: newData.taxPercent },
    { key: 'Total',              old: oldRow[Q.TOTAL],             nw: newData.total },
    { key: 'Notes',              old: oldRow[Q.NOTES],             nw: newData.notes }
  ];

  fields.forEach(f => {
    if (String(f.old) !== String(f.nw)) {
      writeLog('Quotations_Log', 'Quotations', qId, label,
        accountId, accountName, f.key, f.old, f.nw, '');
    }
  });
}
// ============================================================
//  SETTINGS — currencies only from sheet
// ============================================================
function getSettings() {
  const sheet = getSheet('Quotation Settings');
  if (!sheet) return { currencies: [] };
  const lastRow = sheet.getLastRow();
  const currencies = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String)
    : [];
  return { currencies };
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
    const match = String(row[Q.QUOTATION_ID] || '').match(/^Q-(\d+)$/);
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
//  QUOTATIONS — VERSION HISTORY (from log)
// ============================================================
function getQuotationHistory(quotationId) {
  const logSheet = getSheet('Quotations_Log');
  const qSheet   = getSheet('Quotations');
  if (!logSheet) return [];

  const logData = logSheet.getDataRange().getValues();
  const qData   = qSheet.getDataRange().getValues();
  const tz      = Session.getScriptTimeZone();

  // Get current version number and PDF URL from the Quotations sheet
  let currentFolderUrl = '';
  let currentVersion   = 1;
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      currentFolderUrl = qData[i][Q.FOLDER_URL] || '';
      currentVersion   = qData[i][Q.CURRENT_VERSION] || 1;
      break;
    }
  }

  // Gather Version log entries for this specific quotation only (col C = recordId = quotationId)
  // col G (index 6) = FieldChanged = 'Version'
  // col H (index 7) = OldValue = old version number
  // col I (index 8) = NewValue = new version number
  // col J (index 9) = changedBy email
  // col K (index 10) = changedAt timestamp
  const versionRows = logData.slice(1)
    .filter(row =>
      String(row[2]) === String(quotationId) &&
      row[6] === 'Version'
    )
    .map(row => ({
      version:   Number(row[8]) || 0,   // NewValue = the version that was saved
      changedBy: String(row[9]  || ''),
      changedAt: row[10]
        ? Utilities.formatDate(new Date(row[10]), tz, 'dd MMM yyyy HH:mm') : ''
    }))
    .sort((a, b) => b.version - a.version);

  // For each version entry, gather the PDF Link log row that was written at the same edit
  // col G = 'PDF Link', col H = old PDF URL (in Deleted Files), col I = new PDF URL
  const pdfRows = logData.slice(1)
    .filter(row =>
      String(row[2]) === String(quotationId) &&
      row[6] === 'PDF Link'
    )
    .map(row => ({
      version:   Number(row[7]) || 0,   // OldValue = old version number stored as reference
      newPdfUrl: String(row[8]  || '')  // NewValue = new PDF URL for that version
    }));

  // Build the history list — attach correct PDF URL per version
  return versionRows.map(r => {
    const isCurrent = r.version === currentVersion;
    let folderUrl   = '';

    if (isCurrent) {
      folderUrl = currentFolderUrl;
    } else {
      // For old versions, find the PDF Link row where the NEW pdf was saved at that version,
      // which is now in the Deleted Files folder
      const pdfEntry = pdfRows.find(p => p.version === r.version - 1);
      folderUrl = pdfEntry ? pdfEntry.newPdfUrl : '';
    }

    return {
      version:   r.version,
      changedBy: r.changedBy,
      changedAt: r.changedAt,
      folderUrl,
      isCurrent
    };
  });
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
      ? existing.next() : accountFolder.createFolder('Pipeline');
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

  // Log creation — version entry: old = '' (none), new = 1
  writeLog('Quotations_Log', 'Quotations', quotationId,
    quotationId + ' — ' + data.projectName,
    data.accountId, accountName,
    'Version', '', version,
    'Quotation created');

  // Log initial PDF URL
  writeLog('Quotations_Log', 'Quotations', quotationId,
    quotationId + ' — ' + data.projectName,
    data.accountId, accountName,
    'PDF Link', '', pdfUrl,
    'Initial PDF created');

  // Log initial items snapshot
  writeLog('Quotations_Log', 'Quote_Items', quotationId,
    quotationId + ' — ' + data.projectName,
    data.accountId, accountName,
    'Items', '',
    JSON.stringify(data.items.map(i => ({
      name: i.name, qty: i.quantity, desc: i.description,
      notes: i.notes, unitPrice: i.unitPrice, subtotal: i.subtotal
    }))),
    'Items created');

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

  // Find quotation row
  const qData = qSheet.getDataRange().getValues();
  let qRowIndex = -1, currentVersion = 1, currentPdfUrl = '', oldRow = null;

  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(data.id)) {
      qRowIndex      = i + 1;
      currentVersion = qData[i][Q.CURRENT_VERSION] || 1;
      currentPdfUrl  = qData[i][Q.FOLDER_URL] || '';
      oldRow         = qData[i];
      break;
    }
  }
  if (qRowIndex === -1) return { success: false, error: 'Quotation not found.' };

  const newVersion = Number(currentVersion) + 1;

  // Capture old PDF URL before moving — this is what will be stored as the
  // "deleted" PDF link in the log (Drive URL doesn't change on folder move)
  const oldPdfUrl = currentPdfUrl;

  // Get account details + pipeline folder
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
          ? existing.next() : accountFolder.createFolder('Pipeline');
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

  // Generate new PDF
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

  // Move old PDF to Deleted Files AFTER new PDF is created
  moveToDeleted(oldPdfUrl);

  // Snapshot old items for log
  const iDataBefore = iSheet.getDataRange().getValues();
  const oldItems    = iDataBefore.slice(1)
    .filter(r => String(r[QI.QUOTATION_ID]) === String(data.id)
              && r[QI.STATUS] !== 'Deleted')
    .map(r => ({
      name: r[QI.ITEM_NAME], qty: r[QI.QUANTITY],
      desc: r[QI.DESCRIPTION], notes: r[QI.NOTES],
      unitPrice: r[QI.UNIT_PRICE], subtotal: r[QI.SUBTOTAL]
    }));

  // Log field-level changes BEFORE updating the row
  logQuotationFieldChanges(oldRow, {
    ...data, subtotal, total,
    discountAmount, taxAmount
  }, data.accountId, accountName);

  // Update quotation row in place
  qSheet.getRange(qRowIndex, 1, 1, 27).setValues([[
    data.id, newVersion, data.accountId, accountName,
    data.projectName, data.projectDescription,
    data.dateIssued, data.minDays, data.maxDays, data.deliveryDeadline,
    data.pricingMode, data.currency,
    subtotal, data.discounted, data.discountPercent, discountAmount,
    data.taxed, data.taxPercent, taxAmount,
    total, 'Drafted', data.notes, pdfUrl,
    oldRow[Q.CREATED_BY], oldRow[Q.CREATED_AT],
    user, now
  ]]);

  // Update items in place
  const existingItemRows = [];
  for (let i = 1; i < iDataBefore.length; i++) {
    if (String(iDataBefore[i][QI.QUOTATION_ID]) === String(data.id)
     && iDataBefore[i][QI.STATUS] !== 'Deleted') {
      existingItemRows.push({ sheetRow: i + 1, data: iDataBefore[i] });
    }
  }

  data.items.forEach((item, index) => {
    const itemIndex    = index + 1;
    const existing     = existingItemRows.find(r => r.data[QI.ITEM_INDEX] === itemIndex);
    const unitPrice    = data.pricingMode === 'Itemized' ? item.unitPrice : '';
    const itemSubtotal = data.pricingMode === 'Itemized' ? item.subtotal  : '';

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

  // Mark removed items as Deleted
  if (existingItemRows.length > data.items.length) {
    for (let i = data.items.length; i < existingItemRows.length; i++) {
      iSheet.getRange(existingItemRows[i].sheetRow, QI.STATUS + 1).setValue('Deleted');
      iSheet.getRange(existingItemRows[i].sheetRow, QI.LAST_UPDATED_BY + 1).setValue(user);
      iSheet.getRange(existingItemRows[i].sheetRow, QI.LAST_UPDATED_AT + 1).setValue(now);
    }
  }

  // Log version bump: OldValue = old version number, NewValue = new version number
  writeLog('Quotations_Log', 'Quotations', data.id,
    data.id + ' — ' + data.projectName,
    data.accountId, accountName,
    'Version', currentVersion, newVersion,
    'Quotation edited');

  // Log PDF link change: OldValue = old PDF URL (now in Deleted Files), NewValue = new PDF URL
  writeLog('Quotations_Log', 'Quotations', data.id,
    data.id + ' — ' + data.projectName,
    data.accountId, accountName,
    'PDF Link', oldPdfUrl, pdfUrl,
    'PDF updated on v' + newVersion);

  // Log items snapshot
  writeLog('Quotations_Log', 'Quote_Items', data.id,
    data.id + ' — ' + data.projectName,
    data.accountId, accountName,
    'Items',
    JSON.stringify(oldItems),
    JSON.stringify(data.items.map(i => ({
      name: i.name, qty: i.quantity, desc: i.description,
      notes: i.notes, unitPrice: i.unitPrice, subtotal: i.subtotal
    }))),
    'Items updated on v' + newVersion);

  return { success: true, quotationId: data.id, version: newVersion };
}

// ============================================================
//  QUOTATIONS — DELETE
// ============================================================
function deleteQuotation(quotationId) {
  const qSheet = getSheet('Quotations');
  const iSheet = getSheet('Quote_Items');
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

  // Snapshot items BEFORE deleting
  const iData     = iSheet.getDataRange().getValues();
  const itemSnap  = iData.slice(1)
    .filter(r => String(r[QI.QUOTATION_ID]) === String(quotationId)
              && r[QI.STATUS] !== 'Deleted')
    .map(r => ({
      name: r[QI.ITEM_NAME], qty: r[QI.QUANTITY],
      desc: r[QI.DESCRIPTION], notes: r[QI.NOTES],
      unitPrice: r[QI.UNIT_PRICE], subtotal: r[QI.SUBTOTAL]
    }));

  // Move PDF to Deleted Files
  moveToDeleted(pdfUrl);

  // Delete quotation row
  qSheet.deleteRow(qRowIndex);

  // Delete item rows (iterate backwards)
  const iDataFresh = iSheet.getDataRange().getValues();
  for (let i = iDataFresh.length - 1; i >= 1; i--) {
    if (String(iDataFresh[i][QI.QUOTATION_ID]) === String(quotationId)) {
      iSheet.deleteRow(i + 1);
    }
  }

  // Log quotation deletion
  writeLog('Quotations_Log', 'Quotations', quotationId,
    quotationId + ' — ' + projectName,
    accountId, accountName,
    'Status', 'Drafted', 'Deleted',
    'Quotation deleted');

  // Log items snapshot at time of deletion
  writeLog('Quotations_Log', 'Quote_Items', quotationId,
    quotationId + ' — ' + projectName,
    accountId, accountName,
    'Items',
    JSON.stringify(itemSnap), '',
    'Items deleted with quotation');

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

  const qData = qSheet.getDataRange().getValues();
  const user  = Session.getActiveUser().getEmail();
  const now   = new Date();

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

  const accountId   = qRow[Q.ACCOUNT_ID];
  const accountName = qRow[Q.ACCOUNT_NAME];
  const projectName = qRow[Q.PROJECT_NAME];
  const projectDesc = qRow[Q.PROJECT_DESC];
  const pdfUrl      = qRow[Q.FOLDER_URL] || '';
  const deliveryDdl = qRow[Q.DELIVERY_DEADLINE];
  const dateIssued  = qRow[Q.DATE_ISSUED];
  const minDays     = qRow[Q.MIN_DAYS];

  // Calculate DueDate
  let dueDate = null;
  try {
    const ddl       = new Date(deliveryDdl);
    const ddlMinus1 = new Date(ddl); ddlMinus1.setDate(ddl.getDate() - 1);
    const issued    = new Date(dateIssued);
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

  // Move PDF: Pipeline → Q-XXXX folder + create Deliverables subfolder
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

  // Update quotation status
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

  // Re-read items AFTER status update to get Approved rows correctly
  const iDataFresh   = iSheet.getDataRange().getValues();
  const approvedItems = iDataFresh.slice(1).filter(r =>
    String(r[QI.QUOTATION_ID]) === String(quotationId)
    && r[QI.STATUS] === 'Approved'
  );

  // Create Projects row
  const projectId = Utilities.getUuid();
  pSheet.appendRow([
    projectId, quotationId, accountId, accountName,
    projectName, projectDesc,
    deliveryDdl, dueDate,
    'Active', '', now, ''
  ]);

  // Create Project_Items rows
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
    quotationId + ' — ' + projectName,
    accountId, accountName,
    'Status', 'Drafted', 'Confirmed',
    'Quotation confirmed');

  writeLog('Projects_Log', 'Projects', projectId,
    quotationId + ' — ' + projectName,
    accountId, accountName,
    'Status', '', 'Active',
    'Project created on confirmation');

  // Promote account to Client
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
