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
    { key: 'Notes',              old: oldRow[Q.NOTES],            nw: newData.notes }
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
//  DRIVE HELPERS
// ============================================================

/** Standard folder name: "Q-0001 - Project Name" */
function quotationFolderName(quotationId, projectName) {
  return quotationId + ' - ' + projectName;
}

/** Resolves a quotation's Drive folder from its stored URL. Returns Folder or null. */
function getQuotationFolder(folderUrl) {
  if (!folderUrl) return null;
  try {
    const url = extractUrl(folderUrl);
    if (!url) return null;
    const match = url.match(/[-\w]{25,}/);
    if (!match) return null;
    return DriveApp.getFolderById(match[0]);
  } catch(e) {
    Logger.log('getQuotationFolder error: ' + e);
    return null;
  }
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
    folderUrl:          extractUrl(qRow[Q.FOLDER_URL]) || '',
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

  let currentFolderUrl = '', currentVersion = 1;
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      currentFolderUrl = extractUrl(qData[i][Q.FOLDER_URL]) || '';
      currentVersion   = qData[i][Q.CURRENT_VERSION] || 1;
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
      linkUrl = currentFolderUrl;
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
//  Folder lifecycle:
//    ON CREATE  → "Q-XXXX - Project Name/" created inside account's Pipeline/
//                 Archived/ subfolder created immediately (always present)
//                 PDF v1 saved in qFolder root
//    ON EDIT    → new PDF saved in qFolder root; old PDF moved to qFolder/Archived/
//    ON CONFIRM → qFolder moved Pipeline/ → Projects/
//                 Uploads/, Rejected/, Deliverables/ + per-item subfolders created
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
  let accountName = '', accountFolderUrl = '';
  for (let i = 1; i < accounts.length; i++) {
    if (String(accounts[i][0]) === String(data.accountId)) {
      accountName      = accounts[i][A.ACCOUNT_NAME];
      accountFolderUrl = accountsSheet.getRange(i + 1, A.FOLDER_URL + 1)
        .getRichTextValue().getLinkUrl()
        || extractUrl(accounts[i][A.FOLDER_URL]);
      break;
    }
  }

  if (!accountFolderUrl) {
    return { success: false, error: 'Account folder not found. Please check the account record.' };
  }

  // ── Resolve Pipeline subfolder ────────────────────────────
  // Pipeline is guaranteed to exist for every account (createAccount creates it;
  // migrateAccountFolders backfills it). We use getOrCreateSubfolder as a
  // safety net only — we do NOT rely on this to create Pipeline for new accounts.
  let pipelineFolder;
  try {
    const fid           = accountFolderUrl.match(/[-\w]{25,}/)[0];
    const accountFolder = DriveApp.getFolderById(fid);
    pipelineFolder      = getOrCreateSubfolder(accountFolder, 'Pipeline');
  } catch(e) {
    Logger.log('Pipeline folder error: ' + e);
    return { success: false, error: 'Could not access Pipeline folder for this account.' };
  }

  // ── Create quotation folder inside Pipeline ───────────────
  // Structure created upfront so the folder is always predictable:
  //   Q-XXXX - Project Name/
  //     Archived/       ← version history PDFs land here on every edit
  //   (Uploads/, Rejected/, Deliverables/ added on confirmation)
  const folderName = quotationFolderName(quotationId, data.projectName);
  const qFolder    = pipelineFolder.createFolder(folderName);
  qFolder.createFolder('Archived');
  const folderUrl  = qFolder.getUrl();

  // ── Calculations ──────────────────────────────────────────
  let subtotal = data.subtotal;
  if (data.pricingMode === 'Itemized') {
    subtotal = data.items.reduce((sum, item) => sum + (item.subtotal || 0), 0);
  }
  const discountAmount     = data.discounted ? subtotal * (data.discountPercent / 100) : 0;
  const discountedSubtotal = subtotal - discountAmount;
  const taxAmount          = data.taxed ? discountedSubtotal * (data.taxPercent / 100) : 0;
  const total              = discountedSubtotal + taxAmount;

  // ── Generate PDF → save inside qFolder ───────────────────
  const html = generateQuotationHTML({
    quotationId, accountName, ...data,
    subtotal, taxAmount, discountAmount, total, version
  });
  const blob = Utilities.newBlob(html, 'text/html', 'quotation.html')
    .getAs('application/pdf')
    .setName(quotationId + ' - ' + data.projectName + ' - v' + version + '.pdf');
  qFolder.createFile(blob);

  // ── Write quotation row ───────────────────────────────────
  // FOLDER_URL stores the quotation FOLDER url, not the PDF file url
  qSheet.appendRow([
    quotationId, version, data.accountId, accountName,
    data.projectName, data.projectDescription,
    data.dateIssued, data.minDays, data.maxDays, data.deliveryDeadline,
    data.pricingMode, data.currency,
    subtotal, data.discounted, data.discountPercent, discountAmount,
    data.taxed, data.taxPercent, taxAmount,
    total, 'Drafted', data.notes, folderUrl,
    user, now, user, now
  ]);

  // ── Write items ───────────────────────────────────────────
  data.items.forEach((item, index) => {
    iSheet.appendRow([
      Utilities.getUuid(), quotationId, index + 1,
      item.name, item.quantity, item.description, item.notes,
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
    'Folder', '', folderUrl, 'Quotation folder created in Pipeline');

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
//    • Resolves existing qFolder from FOLDER_URL (stays in Pipeline until confirm).
//    • Saves new PDF (vN+1) into qFolder root.
//    • Moves old PDF into qFolder/Archived/ — created on first edit, reused after.
//    • Renames qFolder if project name changed.
//    • FOLDER_URL in sheet stays unchanged throughout.
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
  let qRowIndex = -1, currentVersion = 1, currentFolderUrl = '', oldRow = null;
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(data.id)) {
      qRowIndex        = i + 1;
      currentVersion   = qData[i][Q.CURRENT_VERSION] || 1;
      currentFolderUrl = extractUrl(qData[i][Q.FOLDER_URL]) || '';
      oldRow           = qData[i];
      break;
    }
  }
  if (qRowIndex === -1) return { success: false, error: 'Quotation not found.' };

  const newVersion = Number(currentVersion) + 1;

  // ── Resolve quotation folder ──────────────────────────────
  const qFolder = getQuotationFolder(currentFolderUrl);
  if (!qFolder) {
    return { success: false, error: 'Quotation folder not found in Drive. Cannot save new version.' };
  }

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

  // ── Find old PDF in qFolder root ──────────────────────────
  const oldPdfName = data.id + ' - ' + oldRow[Q.PROJECT_NAME] + ' - v' + currentVersion + '.pdf';
  let oldPdfFile   = null;
  try {
    const files = qFolder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getMimeType() === MimeType.PDF) {
        if (f.getName() === oldPdfName) { oldPdfFile = f; break; }
        if (!oldPdfFile) oldPdfFile = f; // fallback: first PDF in root
      }
    }
  } catch(e) { Logger.log('Old PDF search error: ' + e); }

  // ── Generate new PDF → save into qFolder root ─────────────
  const html = generateQuotationHTML({
    quotationId: data.id, accountName, ...data,
    subtotal, taxAmount, discountAmount, total,
    version: newVersion
  });
  const blob = Utilities.newBlob(html, 'text/html', 'quotation.html')
    .getAs('application/pdf')
    .setName(data.id + ' - ' + data.projectName + ' - v' + newVersion + '.pdf');
  qFolder.createFile(blob);

  // ── Move old PDF → qFolder/Archived/ ─────────────────────
  // Archived/ always exists (created at quotation creation) — no need for get-or-create
  let oldPdfUrl = '';
  if (oldPdfFile) {
    oldPdfUrl = oldPdfFile.getUrl();
    try {
      const archivedFolders = qFolder.getFoldersByName('Archived');
      if (archivedFolders.hasNext()) {
        const archivedFolder = archivedFolders.next();
        archivedFolder.addFile(oldPdfFile);
        qFolder.removeFile(oldPdfFile);
      } else {
        // Safety net: should never happen if createQuotation ran correctly
        Logger.log('Warning: Archived folder missing for ' + data.id + ' — creating it now');
        const archivedFolder = qFolder.createFolder('Archived');
        archivedFolder.addFile(oldPdfFile);
        qFolder.removeFile(oldPdfFile);
      }
    } catch(e) { Logger.log('Old PDF archive error: ' + e); }
  }

  // ── Rename folder if project name changed ─────────────────
  if (String(oldRow[Q.PROJECT_NAME] || '') !== data.projectName) {
    try { qFolder.setName(quotationFolderName(data.id, data.projectName)); }
    catch(e) { Logger.log('Folder rename error: ' + e); }
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
  logQuotationFieldChanges(oldRow, { ...data, subtotal, total, discountAmount, taxAmount },
    data.accountId, accountName);

  // ── Update quotation row — FOLDER_URL stays the same ──────
  qSheet.getRange(qRowIndex, 1, 1, 27).setValues([[
    data.id, newVersion, data.accountId, accountName,
    data.projectName, data.projectDescription,
    data.dateIssued, data.minDays, data.maxDays, data.deliveryDeadline,
    data.pricingMode, data.currency,
    subtotal, data.discounted, data.discountPercent, discountAmount,
    data.taxed, data.taxPercent, taxAmount,
    total, 'Drafted', data.notes, currentFolderUrl,
    oldRow[Q.CREATED_BY], oldRow[Q.CREATED_AT], user, now
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
// ============================================================
function deleteQuotation(quotationId) {
  const qSheet = getSheet('Quotations');
  const iSheet = getSheet('Quote_Items');
  const qData  = qSheet.getDataRange().getValues();

  let qRowIndex = -1, folderUrl = '', accountId = '', accountName = '', projectName = '';
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      if (qData[i][Q.STATUS] === 'Confirmed') {
        return { success: false, error: 'Cannot delete a confirmed quotation.' };
      }
      qRowIndex   = i + 1;
      folderUrl   = extractUrl(qData[i][Q.FOLDER_URL]) || '';
      accountId   = qData[i][Q.ACCOUNT_ID];
      accountName = qData[i][Q.ACCOUNT_NAME];
      projectName = qData[i][Q.PROJECT_NAME];
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

  // Move entire quotation folder to global Archived root
  if (folderUrl) {
    try {
      const fid          = folderUrl.match(/[-\w]{25,}/)[0];
      const qFolder      = DriveApp.getFolderById(fid);
      const archivedRoot = DriveApp.getFolderById(ARCHIVED_FOLDER_ID);
      const parents      = qFolder.getParents();
      archivedRoot.addFolder(qFolder);
      while (parents.hasNext()) parents.next().removeFolder(qFolder);
    } catch(e) { Logger.log('Delete folder move error: ' + e); }
  }

  qSheet.deleteRow(qRowIndex);

  const iDataFresh = iSheet.getDataRange().getValues();
  for (let i = iDataFresh.length - 1; i >= 1; i--) {
    if (String(iDataFresh[i][QI.QUOTATION_ID]) === String(quotationId)) iSheet.deleteRow(i + 1);
  }

  writeLog('Quotations_Log', 'Quotations', quotationId,
    quotationId + ' — ' + projectName, accountId, accountName,
    'Status', 'Drafted', 'Deleted', 'Quotation deleted — folder moved to Archived');

  writeLog('Quotations_Log', 'Quote_Items', quotationId,
    quotationId + ' — ' + projectName, accountId, accountName,
    'Items', JSON.stringify(itemSnap), '', 'Items deleted with quotation');

  return { success: true };
}

// ============================================================
//  QUOTATIONS — CONFIRM
//
//  Drive behaviour:
//    • Moves qFolder from account's Pipeline/ → account's Projects/
//    • Creates Uploads/, Rejected/, Deliverables/ inside qFolder
//    • Creates per-item subfolders inside Uploads/ and Rejected/
//    • Creates Projects + Project_Items sheet rows
//    • Promotes account Lead → Client (moves account folder Leads→Clients root)
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
  if (!qRow)                       return { success: false, error: 'Quotation not found.' };
  if (qRow[Q.STATUS] === 'Confirmed') return { success: false, error: 'Already confirmed.' };

  const accountId   = qRow[Q.ACCOUNT_ID];
  const accountName = qRow[Q.ACCOUNT_NAME];
  const projectName = qRow[Q.PROJECT_NAME];
  const projectDesc = qRow[Q.PROJECT_DESC];
  const folderUrl   = extractUrl(qRow[Q.FOLDER_URL]) || '';
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
      iSheet.getRange(i + 1, QI.STATUS + 1).setValue('Approved');
      iSheet.getRange(i + 1, QI.LAST_UPDATED_BY + 1).setValue(user);
      iSheet.getRange(i + 1, QI.LAST_UPDATED_AT + 1).setValue(now);
    }
  }
  const iDataFresh    = iSheet.getDataRange().getValues();
  const approvedItems = iDataFresh.slice(1).filter(r =>
    String(r[QI.QUOTATION_ID]) === String(quotationId) && r[QI.STATUS] === 'Approved'
  );

  // ── Drive: move qFolder Pipeline → Projects, build subfolders ──
  const qFolder = getQuotationFolder(folderUrl);
  if (qFolder) {
    try {
      // Resolve account folder
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
        const accFid         = accountFolderUrl.match(/[-\w]{25,}/)[0];
        const accountFolder  = DriveApp.getFolderById(accFid);
        const projectsFolder = getOrCreateSubfolder(accountFolder, 'Projects');
        const pipelineFolder = getOrCreateSubfolder(accountFolder, 'Pipeline');

        // Move: Pipeline → Projects
        projectsFolder.addFolder(qFolder);
        pipelineFolder.removeFolder(qFolder);
      }

      // Build workflow subfolders inside qFolder
      const uploadsFolder  = getOrCreateSubfolder(qFolder, 'Uploads');
      const rejectedFolder = getOrCreateSubfolder(qFolder, 'Rejected');
      getOrCreateSubfolder(qFolder, 'Deliverables');

      // Per-item subfolders in Uploads/ and Rejected/
      approvedItems.forEach(item => {
        const itemName = String(item[QI.ITEM_NAME] || 'Unnamed Item');
        getOrCreateSubfolder(uploadsFolder,  itemName);
        getOrCreateSubfolder(rejectedFolder, itemName);
      });

    } catch(e) { Logger.log('Confirm folder error: ' + e); }
  }

  // ── Update quotation status ───────────────────────────────
  qSheet.getRange(qRowIndex, Q.STATUS + 1).setValue('Confirmed');
  qSheet.getRange(qRowIndex, Q.LAST_UPDATED_BY + 1).setValue(user);
  qSheet.getRange(qRowIndex, Q.LAST_UPDATED_AT + 1).setValue(now);

  // ── Create Projects row ───────────────────────────────────
  const projectId = Utilities.getUuid();
  pSheet.appendRow([
    projectId, quotationId, accountId, accountName,
    projectName, projectDesc, deliveryDdl, dueDate,
    'Active', '', now, ''
  ]);

  // ── Create Project_Items rows ─────────────────────────────
  approvedItems.forEach(item => {
    piSheet.appendRow([
      Utilities.getUuid(), projectId, quotationId,
      accountName, projectName, projectDesc,
      item[QI.ITEM_NAME], item[QI.QUANTITY],
      item[QI.DESCRIPTION], item[QI.NOTES],
      'Pending', '', 0, '', '', dueDate, '', now
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
//
//  Expected sheet layout (row 2 = values, row 1 = headers):
//    Col A  Logo Drive file URL   (e.g. https://drive.google.com/file/d/FILE_ID/view)
//    Col B  Company Name          (e.g. MVisuals)
//    Col C  Primary Colour        (hex, e.g. #1a1a2e)   — header bar, total row
//    Col D  Accent Colour         (hex, e.g. #4361ee)   — section labels, borders
//    Col E  Company Address       (plain text, line breaks OK)
//    Col F  Website               (e.g. www.mvisuals.com)
//    Col G  Footer Note           (e.g. "Prices valid for 30 days.")
//
//  If a cell is empty, sensible defaults are used.
//  To rebrand: update row 2 only — the PDF and (later) the UI
//  all read from this single source of truth.
// ============================================================
function getBranding() {
  const sheet = getSheet('Branding');
  const row   = sheet ? (sheet.getDataRange().getValues()[1] || []) : [];
  return {
    logoUrl:     String(row[0] || ''),
    companyName: String(row[1] || 'Quotation'),
    primaryColor:String(row[2] || '#1a1a2e'),
    accentColor: String(row[3] || '#4361ee'),
    address:     String(row[4] || ''),
    website:     String(row[5] || ''),
    footerNote:  String(row[6] || '')
  };
}

/**
 * Fetches a Drive file by its URL/ID and returns a base64 data-URI
 * suitable for embedding in an <img src="..."> tag.
 * Returns '' if the file cannot be fetched (so the PDF still renders).
 *
 * WHY base64: Apps Script's PDF renderer cannot make outbound HTTP
 * requests to fetch external URLs, including Drive share links.
 * Embedding the image as a data-URI is the only reliable way to
 * include it in the generated PDF.
 */
function driveImageToBase64(url) {
  if (!url) return '';
  try {
    const match = url.match(/[-\w]{25,}/);
    if (!match) return '';
    const file     = DriveApp.getFileById(match[0]);
    const blob     = file.getBlob();
    const mime     = blob.getContentType() || 'image/png';
    const base64   = Utilities.base64Encode(blob.getBytes());
    return 'data:' + mime + ';base64,' + base64;
  } catch(e) {
    Logger.log('Logo base64 error: ' + e);
    return '';
  }
}

/**
 * Formats a number in accounting style:
 *   1234567.8 → "1,234,567.80"
 * No sign prefix — caller decides how to label it.
 */
function fmtAccounting(n) {
  const num = Number(n) || 0;
  return num.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

// ============================================================
//  PDF GENERATION
//
//  Design principles:
//    • Single source of truth for colours/logo: Branding sheet
//    • Logo embedded as base64 (required for PDF renderer)
//    • All financial figures in accounting format (1,234.56) — no sign prefixes
//    • Discount and Tax labelled clearly without +/- symbols
//    • Clean two-column header layout, full-width items table
//    • Gracefully degrades if any branding field is missing
// ============================================================
function generateQuotationHTML(data) {
  const branding    = getBranding();
  const primary     = branding.primaryColor;
  const accent      = branding.accentColor;
  const logoDataUri = driveImageToBase64(branding.logoUrl);
  const showPricing = data.pricingMode === 'Itemized';
  const cur         = data.currency || '';

  // ── Items rows ────────────────────────────────────────────
  const itemsRows = (data.items || []).map((item, i) => `
    <tr style="background:${i % 2 === 0 ? '#ffffff' : '#f8f9ff'};">
      <td style="text-align:center;color:#888;">${i + 1}</td>
      <td><strong>${esc(item.name)}</strong></td>
      <td style="text-align:center;">${esc(String(item.quantity || ''))}</td>
      <td style="color:#555;">${esc(item.description || '')}</td>
      ${showPricing ? `
      <td style="text-align:right;">${cur} ${fmtAccounting(item.unitPrice)}</td>
      <td style="text-align:right;">${cur} ${fmtAccounting(item.subtotal)}</td>` : ''}
    </tr>`).join('');

  // ── Pricing summary rows ──────────────────────────────────
  const discountedSubtotal = data.subtotal - (data.discountAmount || 0);

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
    <tr style="background:${primary};color:#fff;">
      <td colspan="${showPricing ? 5 : 3}" style="border:none;background:${primary};"></td>
      <td style="padding:10px;font-weight:700;font-size:14px;">Total</td>
      <td style="padding:10px;text-align:right;font-weight:700;font-size:14px;">${cur} ${fmtAccounting(data.total)}</td>
    </tr>`;

  // ── Delivery text ─────────────────────────────────────────
  const deliveryText = (data.minDays && data.maxDays)
    ? `${data.minDays}–${data.maxDays} working days`
    : (data.deliveryDeadline || '—');

  // ── Footer ────────────────────────────────────────────────
  const footerNote = branding.footerNote
    ? `<p style="font-size:11px;color:#888;margin-top:4px;">${esc(branding.footerNote)}</p>` : '';
  const websiteText = branding.website
    ? `<span style="margin-left:16px;">${esc(branding.website)}</span>` : '';

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
    padding: 32px 40px;
  }

  /* ── Header ── */
  .header {
    display: table;
    width: 100%;
    margin-bottom: 32px;
  }
  .header-left, .header-right {
    display: table-cell;
    vertical-align: middle;
  }
  .header-right { text-align: right; }
  .logo { max-height: 56px; max-width: 180px; }
  .company-name {
    font-size: 20px;
    font-weight: 700;
    color: ${primary};
    letter-spacing: 0.02em;
  }
  .company-meta {
    font-size: 10px;
    color: #888;
    margin-top: 4px;
    line-height: 1.5;
  }
  .doc-title {
    font-size: 22px;
    font-weight: 700;
    color: ${primary};
  }
  .doc-meta {
    font-size: 11px;
    color: #888;
    margin-top: 4px;
    line-height: 1.6;
  }

  /* ── Divider ── */
  .divider {
    border: none;
    border-top: 2px solid ${accent};
    margin: 0 0 24px;
  }

  /* ── Info grid ── */
  .info-grid {
    display: table;
    width: 100%;
    margin-bottom: 28px;
  }
  .info-col {
    display: table-cell;
    width: 50%;
    vertical-align: top;
    padding-right: 20px;
  }
  .info-col:last-child { padding-right: 0; }
  .info-label {
    font-size: 9px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: ${accent};
    margin-bottom: 4px;
  }
  .info-value {
    font-size: 12px;
    color: #1a1a2e;
    line-height: 1.5;
  }
  .info-value.large {
    font-size: 14px;
    font-weight: 600;
  }

  /* ── Section heading ── */
  .section-label {
    font-size: 9px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: ${accent};
    margin-bottom: 8px;
  }

  /* ── Items table ── */
  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
    margin-bottom: 0;
  }
  thead th {
    background: ${primary};
    color: #fff;
    padding: 9px 10px;
    text-align: left;
    font-size: 10px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }
  thead th.num   { width: 32px; text-align: center; }
  thead th.qty   { width: 50px; text-align: center; }
  thead th.price { width: 110px; text-align: right; }
  thead th.sub   { width: 120px; text-align: right; }
  tbody td {
    padding: 8px 10px;
    border-bottom: 1px solid #e8eaf0;
    vertical-align: top;
  }
  tfoot td {
    border-top: 2px solid #e8eaf0;
  }

  /* ── Notes ── */
  .notes-box {
    background: #f8f9ff;
    border-left: 3px solid ${accent};
    padding: 10px 14px;
    font-size: 11px;
    color: #555;
    margin-top: 24px;
    line-height: 1.5;
  }

  /* ── Footer ── */
  .footer {
    margin-top: 40px;
    padding-top: 12px;
    border-top: 1px solid #e8eaf0;
    font-size: 10px;
    color: #aaa;
    display: table;
    width: 100%;
  }
  .footer-left  { display: table-cell; vertical-align: middle; }
  .footer-right { display: table-cell; text-align: right; vertical-align: middle; }
</style>
</head>
<body>

<!-- ═══ HEADER ═══ -->
<div class="header">
  <div class="header-left">
    ${logoDataUri
      ? `<img class="logo" src="${logoDataUri}" alt="${esc(branding.companyName)}">`
      : `<div class="company-name">${esc(branding.companyName)}</div>`}
    ${branding.address
      ? `<div class="company-meta">${esc(branding.address).replace(/\n/g,'<br>')}</div>` : ''}
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

<!-- ═══ INFO GRID ═══ -->
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

<!-- ═══ ITEMS TABLE ═══ -->
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
<!-- ═══ NOTES ═══ -->
<div class="notes-box">
  <strong style="font-size:10px;text-transform:uppercase;letter-spacing:0.06em;color:${accent};">Notes</strong><br>
  ${esc(data.notes)}
</div>` : ''}

<!-- ═══ FOOTER ═══ -->
<div class="footer">
  <div class="footer-left">
    ${esc(branding.companyName)}${websiteText}
    ${footerNote}
  </div>
  <div class="footer-right" style="color:#bbb;font-size:9px;">
    ${esc(data.quotationId)} · v${esc(String(data.version))}
  </div>
</div>

</body>
</html>`;
}

/** HTML-escape helper used inside generateQuotationHTML */
function esc(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
