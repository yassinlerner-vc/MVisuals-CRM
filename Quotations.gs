// ============================================================
//  FIELD-CHANGE LOG HELPER
// ============================================================
function logQuotationFieldChanges(oldRow, newData, accountId, accountName) {
  const tz    = Session.getScriptTimeZone();
  const qId   = String(oldRow[Q.QUOTATION_ID]);
  const label = qId + ' — ' + (newData.projectName || oldRow[Q.PROJECT_NAME]);

  const fields = [
    { key: 'ProjectName',        old: oldRow[Q.PROJECT_NAME],     nw: newData.projectName },
    { key: 'ProjectDescription', old: oldRow[Q.PROJECT_DESC],     nw: newData.projectDescription },
    { key: 'DateIssued',
      old: oldRow[Q.DATE_ISSUED]
        ? Utilities.formatDate(new Date(oldRow[Q.DATE_ISSUED]), tz, 'yyyy-MM-dd') : '',
      nw: newData.dateIssued },
    { key: 'MinDays',            old: oldRow[Q.MIN_DAYS],         nw: newData.minDays },
    { key: 'MaxDays',            old: oldRow[Q.MAX_DAYS],         nw: newData.maxDays },
    { key: 'DeliveryDeadline',
      old: oldRow[Q.DELIVERY_DEADLINE]
        ? Utilities.formatDate(new Date(oldRow[Q.DELIVERY_DEADLINE]), tz, 'yyyy-MM-dd') : '',
      nw: newData.deliveryDeadline },
    { key: 'PricingMode',        old: oldRow[Q.PRICING_MODE],     nw: newData.pricingMode },
    { key: 'Currency',           old: oldRow[Q.CURRENCY],         nw: newData.currency },
    { key: 'Subtotal',           old: oldRow[Q.SUBTOTAL],         nw: newData.subtotal },
    { key: 'Discounted',         old: oldRow[Q.DISCOUNTED],       nw: newData.discounted },
    { key: 'DiscountPercent',    old: oldRow[Q.DISCOUNT_PERCENT], nw: newData.discountPercent },
    { key: 'Taxed',              old: oldRow[Q.TAXED],            nw: newData.taxed },
    { key: 'TaxPercent',         old: oldRow[Q.TAX_PERCENT],      nw: newData.taxPercent },
    { key: 'Total',              old: oldRow[Q.TOTAL],            nw: newData.total },
    { key: 'Notes',              old: oldRow[Q.NOTES],            nw: newData.notes },
    { key: 'IncludeBankDetails', old: oldRow[Q.INCLUDE_BANK_DETAILS], nw: newData.includeBankDetails }
  ];

  fields.forEach(f => {
    if (String(f.old) !== String(f.nw)) {
      writeLog('Quotations_Log', 'Quotations', qId, label,
        accountId, accountName, f.key, f.old, f.nw, '');
    }
  });
}

// ============================================================
//  SETTINGS
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
//  DRIVE HELPERS — quotation PDF (single file, no per-quotation
//  folder anymore; it lives directly in whichever stage folder
//  it's currently in: Pipeline / Confirmed / Fulfilled)
// ============================================================

/** Filename for a live PDF: "v1 - Account Name - Project Name.pdf" */
function quotationPdfName(version, accountName, projectName) {
  return 'v' + version + ' - ' + accountName + ' - ' + projectName + '.pdf';
}

/** Filename for an archived (superseded) PDF, prefixed with quotation ID
 *  so multiple accounts' archived versions don't collide by name. */
function quotationArchivedPdfName(quotationId, version, accountName, projectName) {
  return quotationId + ' - v' + version + ' - ' + accountName + ' - ' + projectName + '.pdf';
}

/** Resolves a quotation's PDF file from its stored URL. Returns File or null. */
function getQuotationPdfFile(pdfUrl) {
  if (!pdfUrl) return null;
  try {
    const url = extractUrl(pdfUrl);
    if (!url) return null;
    const match = url.match(/[-\w]{25,}/);
    if (!match) return null;
    return DriveApp.getFileById(match[0]);
  } catch(e) {
    Logger.log('getQuotationPdfFile error: ' + e);
    return null;
  }
}

/** Maps a quotation's Status to its current Drive stage folder. */
function getQuotationStageFolder(status) {
  if (status === 'Confirmed')  return DriveApp.getFolderById(RUYA_QUOTATIONS_CONFIRMED_ID);
  if (status === 'Fulfilled') return DriveApp.getFolderById(RUYA_QUOTATIONS_FULFILLED_ID);
  return DriveApp.getFolderById(RUYA_QUOTATIONS_PIPELINE_ID); // Drafted / default
}

/** Gets or creates a named subfolder inside a parent folder. Safe — never duplicates. */
function getOrCreateSubfolder(parentFolder, name) {
  const existing = parentFolder.getFoldersByName(name);
  return existing.hasNext() ? existing.next() : parentFolder.createFolder(name);
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
      itemId:       r[QI.ITEM_ID],
      name:         r[QI.ITEM_NAME],
      displayValue: r[QI.DISPLAY_VALUE] || '',
      quantity:     r[QI.QUANTITY],
      description:  r[QI.DESCRIPTION],
      notes:        r[QI.NOTES],
      unitPrice:    r[QI.UNIT_PRICE],
      subtotal:     r[QI.SUBTOTAL]
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
    includeBankDetails: !!qRow[Q.INCLUDE_BANK_DETAILS],
    pdfUrl:             extractUrl(qRow[Q.FOLDER_URL]) || '', // the live quotation PDF
    items
  };
}

// ============================================================
//  QUOTATIONS — VERSION HISTORY
// ============================================================
function getQuotationHistory(quotationId) {
  const logSheet = getSheet('Quotations_Log');
  const qSheet   = getSheet('Quotations');
  if (!logSheet) return [];

  const logData = logSheet.getDataRange().getValues();
  const qData   = qSheet.getDataRange().getValues();
  const tz      = Session.getScriptTimeZone();

  let currentPdfUrl = '', currentVersion = 1;
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      currentPdfUrl  = extractUrl(qData[i][Q.FOLDER_URL]) || '';
      currentVersion = qData[i][Q.CURRENT_VERSION] || 1;
      break;
    }
  }

  // Version rows: FieldChanged = 'Version', col I (index 8) = new version
  const versionRows = logData.slice(1)
    .filter(row => String(row[2]) === String(quotationId) && row[6] === 'Version')
    .map(row => ({
      version:   Number(row[8]) || 0,
      changedBy: String(row[9]  || ''),
      changedAt: row[10]
        ? Utilities.formatDate(new Date(row[10]), tz, 'dd MMM yyyy HH:mm') : ''
    }))
    .sort((a, b) => b.version - a.version);

  // PDF Link rows: FieldChanged = 'PDF Link', col H (index 7) = version number,
  //                col I (index 8) = archived PDF file URL
  const pdfRows = logData.slice(1)
    .filter(row => String(row[2]) === String(quotationId) && row[6] === 'PDF Link')
    .map(row => ({
      version: Number(row[7]) || 0,
      pdfUrl:  String(row[8]  || '')
    }));

  return versionRows.map(r => {
    const isCurrent = r.version === currentVersion;
    let linkUrl = '';
    if (isCurrent) {
      linkUrl = currentPdfUrl;
    } else {
      const entry = pdfRows.find(p => p.version === r.version);
      linkUrl = entry ? entry.pdfUrl : '';
    }
    return { version: r.version, changedBy: r.changedBy, changedAt: r.changedAt, folderUrl: linkUrl, isCurrent };
  });
}

// ============================================================
//  QUOTATIONS — CREATE
//
//  PDF lifecycle:
//    ON CREATE  → PDF "v1 - Account - Project.pdf" saved directly
//                 into Ruya_Quotations/Pipeline
//    ON EDIT    → new version PDF saved into current stage folder;
//                 old PDF renamed + moved to global Archived
//    ON CONFIRM → PDF moved Pipeline → Confirmed; empty project
//                 folder created under account's Projects/
//    ON FULLY PAID (future) → PDF moved Confirmed → Fulfilled
// ============================================================
function createQuotation(data) {
  const ss            = SpreadsheetApp.getActive();
  const qSheet        = ss.getSheetByName('Quotations');
  const iSheet        = ss.getSheetByName('Quote_Items');
  const accountsSheet = ss.getSheetByName('Accounts');
  const quotationId = generateQuotationId();
  const now         = new Date();
  const user        = Session.getActiveUser().getEmail();
  const version     = 1;

  // ── Get account info ──────────────────────────────────────
  const accounts = accountsSheet.getDataRange().getValues();
  let accountName = '';
  for (let i = 1; i < accounts.length; i++) {
    if (String(accounts[i][0]) === String(data.accountId)) {
      accountName = accounts[i][A.ACCOUNT_NAME];
      break;
    }
  }
  if (!accountName) {
    return { success: false, error: 'Account not found. Please check the account record.' };
  }

  // ── Calculations ──────────────────────────────────────────
  let subtotal = data.subtotal;
  if (data.pricingMode === 'Itemized') {
    subtotal = data.items.reduce((sum, item) => sum + (item.subtotal || 0), 0);
  }
  const discountAmount     = data.discounted ? subtotal * (data.discountPercent / 100) : 0;
  const discountedSubtotal = subtotal - discountAmount;
  const taxAmount          = data.taxed ? discountedSubtotal * (data.taxPercent / 100) : 0;
  const total              = discountedSubtotal + taxAmount;

  // ── Generate PDF → save directly into Ruya_Quotations/Pipeline ──
  const html = generateQuotationHTML({
    quotationId, accountName, ...data,
    subtotal, taxAmount, discountAmount, total, version
  });
  const blob = Utilities.newBlob(html, 'text/html', 'quotation.html')
    .getAs('application/pdf')
    .setName(quotationPdfName(version, accountName, data.projectName));
  const pipelineFolder = DriveApp.getFolderById(RUYA_QUOTATIONS_PIPELINE_ID);
  const pdfFile = pipelineFolder.createFile(blob);
  const pdfUrl  = pdfFile.getUrl();

  const includeBankDetails = !!data.includeBankDetails;

  // ── Write quotation row ───────────────────────────────────
  // FOLDER_URL column now stores the quotation PDF FILE url directly
  qSheet.appendRow([
    quotationId, version, data.accountId, accountName,
    data.projectName, data.projectDescription,
    data.dateIssued, data.minDays, data.maxDays, data.deliveryDeadline,
    data.pricingMode, data.currency,
    subtotal, data.discounted, data.discountPercent, discountAmount,
    data.taxed, data.taxPercent, taxAmount,
    total, 'Drafted', data.notes, pdfUrl,
    user, now, user, now,
    includeBankDetails
  ]);

  // ── Write items ───────────────────────────────────────────
  data.items.forEach((item, index) => {
    iSheet.appendRow([
      Utilities.getUuid(), quotationId, index + 1,
      item.displayValue || '', item.name, item.quantity, item.description, item.notes,
      data.pricingMode === 'Itemized' ? item.unitPrice : '',
      data.pricingMode === 'Itemized' ? item.subtotal  : '',
      'Active', user, now
    ]);
  });

  // ── Logs ─────────────────────────────────────────────────
  writeLog('Quotations_Log', 'Quotations', quotationId,
    quotationId + ' — ' + data.projectName,
    data.accountId, accountName,
    'Version', '', version, 'Quotation created');

  writeLog('Quotations_Log', 'Quotations', quotationId,
    quotationId + ' — ' + data.projectName,
    data.accountId, accountName,
    'PDF', '', pdfUrl, 'Quotation PDF created in Pipeline');

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
//
//  Drive behaviour:
//    • New PDF (vN+1) saved into whichever stage folder the
//      quotation is currently in (Pipeline/Confirmed/Fulfilled).
//    • Old PDF renamed with the quotation ID prefix and moved
//      into the global Archived folder.
//    • Status is preserved as-is — editing never changes stage.
// ============================================================
function editQuotation(data) {
  const ss            = SpreadsheetApp.getActive();
  const qSheet        = ss.getSheetByName('Quotations');
  const iSheet        = ss.getSheetByName('Quote_Items');
  const accountsSheet = ss.getSheetByName('Accounts');

  const now  = new Date();
  const user = Session.getActiveUser().getEmail();

  // ── Find quotation row ────────────────────────────────────
  const qData = qSheet.getDataRange().getValues();
  let qRowIndex = -1, currentVersion = 1, currentPdfUrl = '', oldRow = null;
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(data.id)) {
      qRowIndex      = i + 1;
      currentVersion = qData[i][Q.CURRENT_VERSION] || 1;
      currentPdfUrl  = extractUrl(qData[i][Q.FOLDER_URL]) || '';
      oldRow         = qData[i];
      break;
    }
  }
  if (qRowIndex === -1) return { success: false, error: 'Quotation not found.' };

  const newVersion = Number(currentVersion) + 1;
  const status      = oldRow[Q.STATUS];
  const includeBankDetails = !!data.includeBankDetails;

  // ── Get account name ──────────────────────────────────────
  const accounts = accountsSheet.getDataRange().getValues();
  let accountName = '';
  for (let i = 1; i < accounts.length; i++) {
    if (String(accounts[i][0]) === String(data.accountId)) {
      accountName = accounts[i][A.ACCOUNT_NAME]; break;
    }
  }

  // ── Calculations ──────────────────────────────────────────
  let subtotal = data.subtotal;
  if (data.pricingMode === 'Itemized') {
    subtotal = data.items.reduce((sum, item) => sum + (item.subtotal || 0), 0);
  }
  const discountAmount     = data.discounted ? subtotal * (data.discountPercent / 100) : 0;
  const discountedSubtotal = subtotal - discountAmount;
  const taxAmount          = data.taxed ? discountedSubtotal * (data.taxPercent / 100) : 0;
  const total              = discountedSubtotal + taxAmount;

  // ── Generate new PDF → save into current stage folder ─────
  const stageFolder = getQuotationStageFolder(status);
  const html = generateQuotationHTML({
    quotationId: data.id, accountName, ...data,
    subtotal, taxAmount, discountAmount, total,
    version: newVersion
  });
  const blob = Utilities.newBlob(html, 'text/html', 'quotation.html')
    .getAs('application/pdf')
    .setName(quotationPdfName(newVersion, accountName, data.projectName));
  const newPdfFile = stageFolder.createFile(blob);
  const newPdfUrl  = newPdfFile.getUrl();

  // ── Archive old PDF ────────────────────────────────────────
  let oldPdfUrl = '';
  const oldPdfFile = getQuotationPdfFile(currentPdfUrl);
  if (oldPdfFile) {
    try {
      oldPdfFile.setName(
        quotationArchivedPdfName(data.id, currentVersion, oldRow[Q.ACCOUNT_NAME], oldRow[Q.PROJECT_NAME]));
      oldPdfUrl = oldPdfFile.getUrl();
      const archivedRoot = DriveApp.getFolderById(ARCHIVED_FOLDER_ID);
      const parents = oldPdfFile.getParents();
      archivedRoot.addFile(oldPdfFile);
      while (parents.hasNext()) parents.next().removeFile(oldPdfFile);
    } catch(e) { Logger.log('Old PDF archive error: ' + e); }
  }

  // ── Snapshot old items for log ────────────────────────────
  const iDataBefore = iSheet.getDataRange().getValues();
  const oldItems    = iDataBefore.slice(1)
    .filter(r => String(r[QI.QUOTATION_ID]) === String(data.id) && r[QI.STATUS] !== 'Deleted')
    .map(r => ({
      name: r[QI.ITEM_NAME], qty: r[QI.QUANTITY], desc: r[QI.DESCRIPTION],
      notes: r[QI.NOTES], unitPrice: r[QI.UNIT_PRICE], subtotal: r[QI.SUBTOTAL]
    }));

  // ── Log field-level changes ───────────────────────────────
  logQuotationFieldChanges(oldRow, { ...data, subtotal, total, discountAmount, taxAmount, includeBankDetails },
    data.accountId, accountName);

  // ── Update quotation row — FOLDER_URL now points to the new PDF ──
  // 28 columns now (added IncludeBankDetails at the end).
  qSheet.getRange(qRowIndex, 1, 1, 28).setValues([[
    data.id, newVersion, data.accountId, accountName,
    data.projectName, data.projectDescription,
    data.dateIssued, data.minDays, data.maxDays, data.deliveryDeadline,
    data.pricingMode, data.currency,
    subtotal, data.discounted, data.discountPercent, discountAmount,
    data.taxed, data.taxPercent, taxAmount,
    total, status, data.notes, newPdfUrl,
    oldRow[Q.CREATED_BY], oldRow[Q.CREATED_AT], user, now,
    includeBankDetails
  ]]);

  // ── Update items ──────────────────────────────────────────
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
      iSheet.getRange(existing.sheetRow, 1, 1, 13).setValues([[
        existing.data[QI.ITEM_ID], data.id, itemIndex,
        item.displayValue || '', item.name, item.quantity, item.description, item.notes,
        unitPrice, itemSubtotal, 'Active', user, now
      ]]);
    } else {
      iSheet.appendRow([
        Utilities.getUuid(), data.id, itemIndex,
        item.displayValue || '', item.name, item.quantity, item.description, item.notes,
        unitPrice, itemSubtotal, 'Active', user, now
      ]);
    }
  });
  if (existingItemRows.length > data.items.length) {
    for (let i = data.items.length; i < existingItemRows.length; i++) {
      iSheet.getRange(existingItemRows[i].sheetRow, QI.STATUS + 1).setValue('Deleted');
      iSheet.getRange(existingItemRows[i].sheetRow, QI.LAST_UPDATED_BY + 1).setValue(user);
      iSheet.getRange(existingItemRows[i].sheetRow, QI.LAST_UPDATED_AT + 1).setValue(now);
    }
  }

  // ── Logs ─────────────────────────────────────────────────
  writeLog('Quotations_Log', 'Quotations', data.id,
    data.id + ' — ' + data.projectName, data.accountId, accountName,
    'Version', currentVersion, newVersion, 'Quotation edited');

  // PDF Link: OldValue = version number, NewValue = archived PDF URL
  writeLog('Quotations_Log', 'Quotations', data.id,
    data.id + ' — ' + data.projectName, data.accountId, accountName,
    'PDF Link', currentVersion, oldPdfUrl,
    'v' + currentVersion + ' PDF archived on edit to v' + newVersion);

  writeLog('Quotations_Log', 'Quote_Items', data.id,
    data.id + ' — ' + data.projectName, data.accountId, accountName,
    'Items', JSON.stringify(oldItems),
    JSON.stringify(data.items.map(i => ({
      name: i.name, qty: i.quantity, desc: i.description,
      notes: i.notes, unitPrice: i.unitPrice, subtotal: i.subtotal
    }))),
    'Items updated on v' + newVersion);

  return { success: true, quotationId: data.id, version: newVersion };
}

// ============================================================
//  QUOTATIONS — DELETE
//  Only non-Confirmed quotations can be deleted, so the PDF is
//  always in Pipeline at this point. Renamed + moved to Archived.
// ============================================================
function deleteQuotation(quotationId) {
  const qSheet = getSheet('Quotations');
  const iSheet = getSheet('Quote_Items');
  const qData  = qSheet.getDataRange().getValues();

  let qRowIndex = -1, pdfUrl = '', accountId = '', accountName = '', projectName = '', version = 1;
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      if (qData[i][Q.STATUS] === 'Confirmed' || qData[i][Q.STATUS] === 'Fulfilled') {
  return { success: false, error: 'Cannot delete a ' +
    (qData[i][Q.STATUS] === 'Fulfilled' ? 'fulfilled' : 'confirmed') + ' quotation.' };
}
      qRowIndex   = i + 1;
      pdfUrl      = extractUrl(qData[i][Q.FOLDER_URL]) || '';
      accountId   = qData[i][Q.ACCOUNT_ID];
      accountName = qData[i][Q.ACCOUNT_NAME];
      projectName = qData[i][Q.PROJECT_NAME];
      version     = qData[i][Q.CURRENT_VERSION] || 1;
      break;
    }
  }
  if (qRowIndex === -1) return { success: false, error: 'Quotation not found.' };

  const iData    = iSheet.getDataRange().getValues();
  const itemSnap = iData.slice(1)
    .filter(r => String(r[QI.QUOTATION_ID]) === String(quotationId) && r[QI.STATUS] !== 'Deleted')
    .map(r => ({
      name: r[QI.ITEM_NAME], qty: r[QI.QUANTITY], desc: r[QI.DESCRIPTION],
      notes: r[QI.NOTES], unitPrice: r[QI.UNIT_PRICE], subtotal: r[QI.SUBTOTAL]
    }));

  // Move the quotation PDF to Archived
  const pdfFile = getQuotationPdfFile(pdfUrl);
  if (pdfFile) {
    try {
      pdfFile.setName(quotationArchivedPdfName(quotationId, version, accountName, projectName));
      const archivedRoot = DriveApp.getFolderById(ARCHIVED_FOLDER_ID);
      const parents = pdfFile.getParents();
      archivedRoot.addFile(pdfFile);
      while (parents.hasNext()) parents.next().removeFile(pdfFile);
    } catch(e) { Logger.log('Delete PDF archive error: ' + e); }
  }

  qSheet.deleteRow(qRowIndex);

  const iDataFresh = iSheet.getDataRange().getValues();
  for (let i = iDataFresh.length - 1; i >= 1; i--) {
    if (String(iDataFresh[i][QI.QUOTATION_ID]) === String(quotationId)) iSheet.deleteRow(i + 1);
  }

  writeLog('Quotations_Log', 'Quotations', quotationId,
    quotationId + ' — ' + projectName, accountId, accountName,
    'Status', 'Drafted', 'Deleted', 'Quotation deleted — PDF moved to Archived');

  writeLog('Quotations_Log', 'Quote_Items', quotationId,
    quotationId + ' — ' + projectName, accountId, accountName,
    'Items', JSON.stringify(itemSnap), '', 'Items deleted with quotation');

  return { success: true };
}

// ============================================================
//  QUOTATIONS — CONFIRM
//
//  Drive behaviour on confirm:
//    • PDF moved Ruya_Quotations/Pipeline → Ruya_Quotations/Confirmed
//    • An empty folder "Project Name - Account Name" is created
//      inside the account's Projects/ folder (no PDF inside it —
//      the PDF lives in Ruya_Quotations, not the account folder)
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

  // ── Find quotation ────────────────────────────────────────
  let qRowIndex = -1, qRow = null;
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      qRowIndex = i + 1; qRow = qData[i]; break;
    }
  }
  if (!qRow) return { success: false, error: 'Quotation not found.' };
if (qRow[Q.STATUS] === 'Confirmed' || qRow[Q.STATUS] === 'Fulfilled') {
  return { success: false, error: 'Already confirmed.' };
}
  const accountId   = qRow[Q.ACCOUNT_ID];
  const accountName = qRow[Q.ACCOUNT_NAME];
  const projectName = qRow[Q.PROJECT_NAME];
  const projectDesc = qRow[Q.PROJECT_DESC];
  const pdfUrl      = extractUrl(qRow[Q.FOLDER_URL]) || '';
  const deliveryDdl = qRow[Q.DELIVERY_DEADLINE];
  const dateIssued  = qRow[Q.DATE_ISSUED];
  const minDays     = qRow[Q.MIN_DAYS];

  // ── Calculate DueDate ─────────────────────────────────────
  let dueDate = null;
  try {
    const ddl       = new Date(deliveryDdl);
    const ddlMinus1 = new Date(ddl); ddlMinus1.setDate(ddl.getDate() - 1);
    const issued    = new Date(dateIssued);
    const minTarget = new Date(issued); minTarget.setDate(issued.getDate() + Number(minDays));
    dueDate = ddlMinus1 < minTarget ? ddlMinus1 : minTarget;
  } catch(e) { Logger.log('DueDate calc: ' + e); }

  // ── Approve items ─────────────────────────────────────────
  const iData = iSheet.getDataRange().getValues();
  for (let i = 1; i < iData.length; i++) {
    if (String(iData[i][QI.QUOTATION_ID]) === String(quotationId)
     && iData[i][QI.STATUS] === 'Active') {
      iSheet.getRange(i + 1, QI.STATUS          + 1).setValue('Approved');
      iSheet.getRange(i + 1, QI.LAST_UPDATED_BY + 1).setValue(user);
      iSheet.getRange(i + 1, QI.LAST_UPDATED_AT + 1).setValue(now);
    }
  }
  const iDataFresh    = iSheet.getDataRange().getValues();
  const approvedItems = iDataFresh.slice(1).filter(r =>
    String(r[QI.QUOTATION_ID]) === String(quotationId) && r[QI.STATUS] === 'Approved'
  );

  // ── Drive: move PDF Pipeline → Confirmed ──────────────────
  try {
    const pdfFile = getQuotationPdfFile(pdfUrl);
    if (pdfFile) {
      const confirmedFolder = DriveApp.getFolderById(RUYA_QUOTATIONS_CONFIRMED_ID);
      const parents = pdfFile.getParents();
      confirmedFolder.addFile(pdfFile);
      while (parents.hasNext()) parents.next().removeFile(pdfFile);
    }
  } catch(e) { Logger.log('Confirm PDF move error: ' + e); }

  // ── Drive: create empty project folder inside account's Projects/ ──
  try {
    const accounts2 = accountsSheet.getDataRange().getValues();
    let accountFolderUrl = '';
    for (let i = 1; i < accounts2.length; i++) {
      if (String(accounts2[i][0]) === String(accountId)) {
        accountFolderUrl = accountsSheet.getRange(i + 1, A.FOLDER_URL + 1)
          .getRichTextValue().getLinkUrl()
          || extractUrl(accounts2[i][A.FOLDER_URL]);
        break;
      }
    }
    if (accountFolderUrl) {
      const accFid          = accountFolderUrl.match(/[-\w]{25,}/)[0];
      const accountFolder   = DriveApp.getFolderById(accFid);
      const projectsFolder  = getOrCreateSubfolder(accountFolder, 'Projects');
      getOrCreateSubfolder(projectsFolder, projectName + ' - ' + accountName);
    }
  } catch(e) { Logger.log('Confirm project folder error: ' + e); }

  // ── Update quotation status ───────────────────────────────
  qSheet.getRange(qRowIndex, Q.STATUS          + 1).setValue('Confirmed');
  qSheet.getRange(qRowIndex, Q.LAST_UPDATED_BY + 1).setValue(user);
  qSheet.getRange(qRowIndex, Q.LAST_UPDATED_AT + 1).setValue(now);

  // ── Create Projects row ───────────────────────────────────
  const projectId = Utilities.getUuid();
  pSheet.appendRow([
    projectId, quotationId, accountId, accountName,
    projectName, projectDesc, deliveryDdl, dueDate,
    'Active', '', now, '',              // cols 0-11
    qRow[Q.TOTAL],                      // col 12 Total Amount
    qRow[Q.CURRENCY],                   // col 13 Currency
    0,                                  // col 14 unused
    0                                   // col 15 unused
  ]);

  // ── Create Project_Items rows ─────────────────────────────
  approvedItems.forEach(item => {
    piSheet.appendRow([
      Utilities.getUuid(),          // Item ID
      projectId,                    // Project ID
      quotationId,                  // Quotation ID
      accountName,                  // Account Name
      projectName,                  // Project Name
      projectDesc,                  // Project Description
      item[QI.DISPLAY_VALUE],       // Display Value
      item[QI.ITEM_NAME],           // Item Name
      item[QI.QUANTITY],            // Quantity
      item[QI.DESCRIPTION],         // Description
      item[QI.NOTES],               // Notes
      now                           // Created At
    ]);
  });

  // ── Logs ─────────────────────────────────────────────────
  writeLog('Quotations_Log', 'Quotations', quotationId,
    quotationId + ' — ' + projectName, accountId, accountName,
    'Status', 'Drafted', 'Confirmed', 'Quotation confirmed');

  writeLog('Projects_Log', 'Projects', projectId,
    quotationId + ' — ' + projectName, accountId, accountName,
    'Status', '', 'Active', 'Project created on confirmation');

  // ── Promote account Lead → Client ─────────────────────────
  promoteAccountToClient(accountId);

  return { success: true };
}

// ============================================================
//  BRANDING — read from the Branding sheet
//  Cols: A LogoUrl (unused — logo is hardcoded, see Logo.gs) |
//  B CompanyName | C PrimaryColor | D AccentColor | E Address
//  (unused — not shown on PDF anymore) | F Website | G FooterNote |
//  H Phone | I Email | J Instagram
// ============================================================
function getBranding() {
  const sheet = getSheet('Branding');
  const row   = sheet ? (sheet.getDataRange().getValues()[1] || []) : [];

  const websiteText  = String(row[5] || '');
  const phoneText    = String(row[7] || '');
  const emailText    = String(row[8] || '');
  const instagramText = String(row[9] || '');

  return {
    companyName:  String(row[1] || 'Ruya Studios'),
    primaryColor: String(row[2] || '#4A1D13'),
    accentColor:  String(row[3] || '#4A1D13'),

    // Visible text in the PDF.
    website: websiteText,
    phone: phoneText,
    email: emailText,
    instagram: instagramText,

    // Actual destinations used by clickable PDF links.
    websiteUrl: websiteText
      ? 'https://www.ruyastudios.com/'
      : '',

    instagramUrl: instagramText
      ? 'https://www.instagram.com/ruyastudios.co/'
      : '',

    // WhatsApp number in international digits-only format.
    phoneUrl: phoneText
      ? 'https://wa.me/201158541967'
      : '',

    emailUrl: emailText
      ? 'mailto:' + emailText
      : '',

    footerNote: String(row[6] || '')
  };
}

// ============================================================
//  BANK DETAILS — read from the Bank_Details sheet (row 2)
// ============================================================
function getBankDetails() {
  const sheet = getSheet('Bank_Details');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const row = data[1];
  return {
    accountName:     String(row[BD.ACCOUNT_NAME]     || ''),
    bankName:        String(row[BD.BANK_NAME]        || ''),
    iban:            String(row[BD.IBAN]             || ''),
    swiftCode:       String(row[BD.SWIFT_CODE]       || ''),
    nationality:     String(row[BD.NATIONALITY]      || ''),
    branchName:      String(row[BD.BRANCH_NAME]      || ''),
    branchCode:      String(row[BD.BRANCH_CODE]      || ''),
    address:         String(row[BD.ADDRESS]          || ''),
    instapayAddress: String(row[BD.INSTAPAY_ADDRESS] || ''),
    instapayMobile:  String(row[BD.INSTAPAY_MOBILE]  || '')
  };
}

// ============================================================
//  QUOTATION TERMS — read from the Quotation_Terms sheet.
//  One row per term (Term | Detail); order in the sheet is the
//  order they print in, so Yassin can add/reorder freely.
// ============================================================
function getQuotationTerms() {
  const sheet = getSheet('Quotation_Terms');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1)
    .filter(row => row[0])
    .map(row => ({ term: String(row[0]), detail: String(row[1] || '') }));
}

function fmtAccounting(n) {
  const num = Number(n) || 0;
  return num.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

// ============================================================
//  PDF GENERATION
// ============================================================
function generateQuotationHTML(data) {
  const branding    = getBranding();
  const primary     = branding.primaryColor;
  const accent      = branding.accentColor;
  const showPricing = data.pricingMode === 'Itemized';
  const cur         = data.currency || '';

  const itemsRows = (data.items || []).map((item, i) => `
    <tr style="background:${i % 2 === 0 ? '#ffffff' : '#f8f9ff'};">
      <td style="text-align:center;color:#888;">${i + 1}</td>
      <td><strong>${esc(item.displayValue || item.name)}</strong></td>
      <td style="text-align:center;">${esc(String(item.quantity || ''))}</td>
      <td style="color:#555;">${esc(item.description || '')}</td>
      ${showPricing ? `
      <td style="text-align:right;">${cur} ${fmtAccounting(item.unitPrice)}</td>
      <td style="text-align:right;">${cur} ${fmtAccounting(item.subtotal)}</td>` : ''}
    </tr>`).join('');

  const summaryRows = `
    <tr>
      <td colspan="${showPricing ? 5 : 3}" style="border:none;"></td>
      <td style="padding:6px 10px;color:#555;">Subtotal</td>
      <td style="padding:6px 10px;text-align:right;">${cur} ${fmtAccounting(data.subtotal)}</td>
    </tr>
    ${data.discounted ? `
    <tr>
      <td colspan="${showPricing ? 5 : 3}" style="border:none;"></td>
      <td style="padding:6px 10px;color:#555;">Discount (${esc(String(data.discountPercent || 0))}%)</td>
      <td style="padding:6px 10px;text-align:right;">${cur} ${fmtAccounting(data.discountAmount)}</td>
    </tr>` : ''}
    ${data.taxed ? `
    <tr>
      <td colspan="${showPricing ? 5 : 3}" style="border:none;"></td>
      <td style="padding:6px 10px;color:#555;">Tax (${esc(String(data.taxPercent || 0))}%)</td>
      <td style="padding:6px 10px;text-align:right;">${cur} ${fmtAccounting(data.taxAmount)}</td>
    </tr>` : ''}
    <tr>
      <td colspan="${showPricing ? 5 : 3}" style="border:none;"></td>
      <td style="padding:10px;font-weight:700;font-size:14px;color:#1a1a2e;border-top:2px solid #e2e4ef;">Total</td>
      <td style="padding:10px;text-align:right;font-weight:700;font-size:14px;color:#1a1a2e;border-top:2px solid #e2e4ef;">${cur} ${fmtAccounting(data.total)}</td>
    </tr>`;

  const deliveryText = (data.minDays && data.maxDays)
    ? `${data.minDays}–${data.maxDays} working days`
    : (data.deliveryDeadline || '—');

  // ── Terms block (page-broken to always start fresh — see below) ──
  const terms = getQuotationTerms();
  const termsRows = terms.map(t => `
    <tr>
      <td style="padding:9px 10px;font-weight:700;color:#1a1a2e;border-bottom:1px solid #f0f2fa;width:160px;">${esc(t.term)}</td>
      <td style="padding:9px 10px;color:#444;border-bottom:1px solid #f0f2fa;">${esc(t.detail)}</td>
    </tr>`).join('');
  const termsBlock = terms.length ? `
  <div class="section-block avoid-break">
    <div class="section-label">Terms</div>
    <table class="mini-table">
      <thead><tr><th style="width:160px;">Term</th><th>Detail</th></tr></thead>
      <tbody>${termsRows}</tbody>
    </table>
  </div>` : '';

  // ── Payment Details block (only if toggled on for this quotation) ──
  const bank = data.includeBankDetails ? getBankDetails() : null;
  let paymentBlock = '';
  if (data.includeBankDetails && bank) {
    const hasBankTransfer = bank.accountName || bank.bankName || bank.iban ||
      bank.swiftCode || bank.nationality || bank.branchName || bank.branchCode || bank.address;
    const hasInstapay = bank.instapayAddress || bank.instapayMobile;

    const bankTransferHtml = hasBankTransfer ? `
      <div class="pay-subblock avoid-break">
        <div class="pay-subheader">Bank Transfer</div>
        <table class="pay-table">
          <tr><td class="pay-label">Account Name</td><td class="pay-value">${esc(bank.accountName || '—')}</td><td class="pay-label">Bank Name</td><td class="pay-value">${esc(bank.bankName || '—')}</td></tr>
          <tr><td class="pay-label">IBAN</td><td class="pay-value">${esc(bank.iban || '—')}</td><td class="pay-label">SWIFT Code</td><td class="pay-value">${esc(bank.swiftCode || '—')}</td></tr>
          <tr><td class="pay-label">Nationality</td><td class="pay-value">${esc(bank.nationality || '—')}</td><td class="pay-label">Branch Name</td><td class="pay-value">${esc(bank.branchName || '—')}</td></tr>
          <tr><td class="pay-label">Address</td><td class="pay-value">${esc(bank.address || '—')}</td><td class="pay-label">Branch Code</td><td class="pay-value">${esc(bank.branchCode || '—')}</td></tr>
        </table>
      </div>` : '';

    const instapayHtml = hasInstapay ? `
      <div class="pay-subblock avoid-break">
        <div class="pay-subheader">Instapay</div>
        <table class="pay-table">
          <tr><td class="pay-label">Payment Address</td><td class="pay-value">${esc(bank.instapayAddress || '—')}</td><td class="pay-label">Mobile Number</td><td class="pay-value">${esc(bank.instapayMobile || '—')}</td></tr>
        </table>
      </div>` : '';

    if (bankTransferHtml || instapayHtml) {
      paymentBlock = `
      <div class="section-block avoid-break">
        <div class="section-label">Payment Details</div>
        ${bankTransferHtml}
        ${instapayHtml}
      </div>`;
    }
  }

  // ── Terms + Payment Details always start on a fresh page,
  //    right after the items/notes on the previous page ──
  const termsAndPaymentPage = (termsBlock || paymentBlock)
    ? `<div style="page-break-before: always;">${termsBlock}${paymentBlock}</div>`
    : '';

const ICON_PHONE = `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/></svg>`;
const ICON_WHATSAPP = `
<svg width="10" height="10" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2"
     stroke-linecap="round" stroke-linejoin="round">
  <path d="M20.52 3.48A11.78 11.78 0 0 0 12.08 0
           C5.56 0 .26 5.3 .26 11.82
           c0 2.08 .54 4.11 1.57 5.9L.17 24
           l6.45-1.63a11.8 11.8 0 0 0 5.46 1.34h.01
           c6.52 0 11.82-5.3 11.82-11.82
           0-3.16-1.23-6.13-3.39-8.41Z"/>
  <path d="M8.7 6.9c-.22-.49-.45-.5-.66-.51
           h-.56c-.2 0-.53.08-.81.4
           -.28.31-1.06 1.04-1.06 2.54
           s1.09 2.95 1.21 3.15
           c.15.2 2.1 3.37 5.18 4.59
           2.56 1.01 3.08.81 3.63.76
           .55-.05 1.77-.72 2.02-1.41
           .25-.69.25-1.28.17-1.41
           -.07-.13-.28-.2-.58-.35
           -.3-.15-1.77-.87-2.04-.97
           -.27-.1-.47-.15-.67.15
           -.2.3-.77.97-.94 1.17
           -.17.2-.35.22-.65.07
           -.3-.15-1.26-.46-2.4-1.47
           -.89-.79-1.49-1.77-1.66-2.07
           -.17-.3-.02-.46.13-.61
           .13-.13.3-.35.45-.52
           .15-.17.2-.3.3-.5
           .1-.2.05-.37-.02-.52
           -.07-.15-.62-1.49-.85-2.04Z"/>
</svg>`;
  const ICON_MAIL = `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 6-10 7L2 6"/></svg>`;
  const ICON_INSTAGRAM = `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="2" width="20" height="20" rx="5"/><circle cx="12" cy="12" r="4"/><circle cx="17.5" cy="6.5" r="0.5" fill="currentColor" stroke="none"/></svg>`;
  const ICON_GLOBE = `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>`;
  // ── Footer ───────────────────────────────────────────────
  const instagramHandle = branding.instagram
  ? String(branding.instagram).replace(/^@/, '')
  : '';

const emailHtml = branding.email
  ? pdfLink(
      branding.emailUrl,
      `${ICON_MAIL}${esc(branding.email)}`,
      'footer-item'
    )
  : '';

const instagramHtml = branding.instagram
  ? pdfLink(
      branding.instagramUrl,
      `${ICON_INSTAGRAM}${esc(instagramHandle)}`,
      'footer-item'
    )
  : '';

const phoneHtml = branding.phone
  ? pdfLink(
      branding.phoneUrl,
      `${ICON_WHATSAPP}${esc(branding.phone)}`,
      'footer-item'
    )
  : '';

const websiteHtml = branding.website
  ? pdfLink(
      branding.websiteUrl,
      `${ICON_GLOBE}${esc(branding.website)}`,
      'footer-item'
    )
  : '';

const footerNoteHtml = branding.footerNote
  ? `<div class="footer-note">${esc(branding.footerNote)}</div>`
  : '';
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
    font-size: 12px;
    color: #1a1a2e;
    background: #fff;
    padding: 32px 40px 110px;
  }
  .avoid-break { page-break-inside: avoid; break-inside: avoid; }
  .header { display: table; width: 100%; margin-bottom: 24px; }
  .header-left, .header-right { display: table-cell; vertical-align: middle; }
  .header-right { text-align: right; }
  .logo { max-height: 34px; width: auto; }
  .doc-title { font-size: 22px; font-weight: 700; color: ${primary}; }
  .doc-meta { font-size: 11px; color: #888; margin-top: 4px; line-height: 1.6; }
  .divider { border: none; border-top: 2px solid ${accent}; margin: 0 0 24px; }
  .info-grid { display: table; width: 100%; margin-bottom: 28px; }
  .info-col { display: table-cell; width: 50%; vertical-align: top; padding-right: 20px; }
  .info-col:last-child { padding-right: 0; }
  .info-label { font-size: 10.5px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: ${accent}; margin-bottom: 4px; }
  .info-value { font-size: 12px; color: #1a1a2e; line-height: 1.5; }
  .info-value.large { font-size: 14px; font-weight: 600; }
  .section-label { font-size: 10.5px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: ${accent}; margin-bottom: 8px; }
  .section-block { margin-top: 28px; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 0; }
  thead th { background: none; color: #888; padding: 9px 10px; text-align: left; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; border-bottom: 1px solid #e2e4ef; }
  thead th.num   { width: 32px; text-align: center; }
  thead th.qty   { width: 50px; text-align: center; }
  thead th.price { width: 110px; text-align: right; }
  thead th.sub   { width: 120px; text-align: right; }
  tbody td { padding: 8px 10px; border-bottom: 1px solid #e8eaf0; vertical-align: top; }
  .mini-table thead th { background: #F3EDEB; color: #888; }
  .mini-table tbody tr, .mini-table tr { page-break-inside: avoid; break-inside: avoid; }
  .pay-subheader { background: #F3EDEB; color: #888; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; padding: 6px 10px; margin-top: 12px; margin-bottom: 4px; }
  .pay-table td { padding: 7px 10px; font-size: 12px; border-bottom: 1px solid #f0f2fa; }
  .pay-table tr { page-break-inside: avoid; break-inside: avoid; }
  .pay-label { font-weight: 700; color: #1a1a2e; width: 22%; }
  .pay-value { color: #444; width: 28%; }
  .notes-box { background: #f8f9ff; border-left: 3px solid ${accent}; padding: 10px 14px; font-size: 11px; color: #555; margin-top: 24px; line-height: 1.5; }

  /* ── Footer: fixed to the bottom of every page, full width,
     Contact block on the left, Find Us block on the right ── */
  .footer-fixed { position: fixed; left: 40px; right: 40px; bottom: 20px; }
  .footer-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding-top: 12px;
    border-top: 1px solid #e8eaf0;
    font-size: 10px;
    color: #888;
  }
  .footer-side { display: flex; align-items: center; white-space: nowrap; }
  .footer-side-right { justify-content: flex-end; }
  .footer-heading { font-size: 10.5px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: ${accent}; margin-right: 12px; vertical-align: middle; }
  .footer-item { display: inline-flex; align-items: center; font-size: 10px; color: #888; text-decoration: underline; margin-right: 16px; }
  .footer-item:last-child { margin-right: 0; }
  .footer-item svg { margin-right: 4px; vertical-align: -2px; }
  .footer-note { margin-top: 6px; font-size: 9px; color: #aaa; text-align: center; }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <img class="logo" src="${RUYA_LOGO_BASE64}" alt="${esc(branding.companyName)}">
  </div>
  <div class="header-right">
    <div class="doc-title">QUOTATION</div>
    <div class="doc-meta">
      <strong style="color:#1a1a2e;">${esc(data.quotationId)}</strong>
      &nbsp;·&nbsp; Version ${esc(String(data.version))}<br>
      Date Issued: ${esc(data.dateIssued || '—')}<br>
      Delivery: ${esc(deliveryText)}
    </div>
  </div>
</div>

<hr class="divider">

<div class="info-grid">
  <div class="info-col">
    <div class="info-label">Prepared for</div>
    <div class="info-value large">${esc(data.accountName)}</div>
  </div>
  <div class="info-col">
    <div class="info-label">Project</div>
    <div class="info-value large">${esc(data.projectName)}</div>
    ${data.projectDescription
      ? `<div class="info-value" style="margin-top:4px;color:#666;">${esc(data.projectDescription)}</div>` : ''}
  </div>
</div>

<div class="section-label">Scope of Work</div>
<table>
  <thead>
    <tr>
      <th class="num">#</th>
      <th>Item</th>
      <th class="qty">Qty</th>
      <th>Description</th>
      ${showPricing ? '<th class="price">Unit Price</th><th class="sub">Subtotal</th>' : ''}
    </tr>
  </thead>
  <tbody>
    ${itemsRows}
  </tbody>
  <tfoot>
    ${summaryRows}
  </tfoot>
</table>

${data.notes ? `
<div class="notes-box">
  <strong style="font-size:10px;text-transform:uppercase;letter-spacing:0.06em;color:${accent};">Notes</strong><br>
  ${esc(data.notes)}
</div>` : ''}

${termsAndPaymentPage}

<div class="footer-fixed">
  <div class="footer-row">
    <div class="footer-side footer-side-left">
  <span class="footer-heading">Contact</span>
  ${phoneHtml}
  ${emailHtml}
</div>

<div class="footer-side footer-side-right">
  <span class="footer-heading">Find Us</span>
  ${instagramHtml}
  ${websiteHtml}
</div>
  </div>
  ${footerNoteHtml}
</div>

</body>
</html>`;
}

function esc(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
function escAttr(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/'/g, '&#39;');
}
function pdfLink(url, content, className) {
  if (!url || !content) return '';

  return `<a href="${escAttr(url)}" class="${className || ''}">${content}</a>`;
}
