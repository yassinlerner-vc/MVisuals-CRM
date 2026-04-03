// ============================================================
//  PROJECTS — LIST
// ============================================================
function getProjects() {
  const sheet = getSheet('Projects');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const tz   = Session.getScriptTimeZone();

  return data.slice(1)
    .filter(row => row[P.PROJECT_ID])
    .map(row => ({
      projectId:        String(row[P.PROJECT_ID]),
      quotationId:      String(row[P.QUOTATION_ID]),
      accountId:        String(row[P.ACCOUNT_ID]),
      accountName:      String(row[P.ACCOUNT_NAME]),
      projectName:      String(row[P.PROJECT_NAME]),
      projectDesc:      String(row[P.PROJECT_DESC] || ''),
      deliveryDeadline: row[P.DELIVERY_DEADLINE]
        ? Utilities.formatDate(new Date(row[P.DELIVERY_DEADLINE]), tz, 'yyyy-MM-dd') : '',
      dueDate:          row[P.DUE_DATE]
        ? Utilities.formatDate(new Date(row[P.DUE_DATE]), tz, 'yyyy-MM-dd') : '',
      status:           String(row[P.STATUS] || 'Active'),
      internalNotes:    String(row[P.INTERNAL_NOTES] || ''),
      createdAt:        row[P.CREATED_AT]
        ? Utilities.formatDate(new Date(row[P.CREATED_AT]), tz, 'yyyy-MM-dd') : '',
      completedAt:      row[P.COMPLETED_AT]
        ? Utilities.formatDate(new Date(row[P.COMPLETED_AT]), tz, 'yyyy-MM-dd') : ''
    }))
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

// ============================================================
//  PROJECTS — GET ITEMS FOR A PROJECT
// ============================================================
function getProjectItems(projectId) {
  const piSheet = getSheet('Project_Items');
  const qSheet  = getSheet('Quotations');
  const pSheet  = getSheet('Projects');
  if (!piSheet) return { items: [], folderUrl: '' };

  const piData = piSheet.getDataRange().getValues();
  const tz     = Session.getScriptTimeZone();

  // Get quotationId and folderUrl
  let quotationId = '', folderUrl = '';
  const pData = pSheet.getDataRange().getValues();
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) {
      quotationId = String(pData[i][P.QUOTATION_ID]);
      break;
    }
  }
  if (quotationId) {
    const qData = qSheet.getDataRange().getValues();
    for (let i = 1; i < qData.length; i++) {
      if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
        folderUrl = extractUrl(qData[i][Q.FOLDER_URL]) || '';
        break;
      }
    }
  }

  const items = piData.slice(1)
    .filter(row => String(row[PI.PROJECT_ID]) === String(projectId))
    .map(row => ({
      itemId:             String(row[PI.ITEM_ID]),
      projectId:          String(row[PI.PROJECT_ID]),
      quotationId:        String(row[PI.QUOTATION_ID]),
      accountName:        String(row[PI.ACCOUNT_NAME]),
      projectName:        String(row[PI.PROJECT_NAME]),
      projectDesc:        String(row[PI.PROJECT_DESC] || ''),   // ← fixed: was PI.PROJECT_DESCRIPTION
      itemName:           String(row[PI.ITEM_NAME]),
      quantity:           row[PI.QUANTITY],
      description:        String(row[PI.DESCRIPTION] || ''),
      notes:              String(row[PI.NOTES]        || ''),
      deliveryStatus:     String(row[PI.DELIVERY_STATUS] || 'Pending'),
      assignedTo:         String(row[PI.ASSIGNED_TO]     || ''),
      redoCount:          Number(row[PI.REDO_COUNT]       || 0),
      uploadedFileUrl:    String(row[PI.UPLOADED_FILE_URL]|| ''),
      internalNotes:      String(row[PI.INTERNAL_NOTES]   || ''),
      dueDate:            row[PI.DUE_DATE]
        ? Utilities.formatDate(new Date(row[PI.DUE_DATE]), tz, 'yyyy-MM-dd') : '',
      completedAt:        row[PI.COMPLETED_AT]
        ? Utilities.formatDate(new Date(row[PI.COMPLETED_AT]), tz, 'yyyy-MM-dd') : '',
      createdAt:          row[PI.CREATED_AT]
        ? Utilities.formatDate(new Date(row[PI.CREATED_AT]), tz, 'yyyy-MM-dd') : ''
    }));

  return { items, folderUrl, quotationId };
}

// ============================================================
//  PROJECTS — UPLOAD FILE TO ITEM
//
//  Supports multiple uploads per item:
//    • Each call uploads ONE file (the front-end loops per file).
//    • File URLs are stored newline-separated in UPLOADED_FILE_URL.
//    • On a re-upload after rejection, ALL previous files are cleared
//      (moved to Rejected/) and the cell is reset before appending.
//    • Status is set to 'Uploaded' on the first file of a batch;
//      subsequent calls in the same batch leave it as 'Uploaded'.
// ============================================================
function uploadProjectItemFile(params) {
  // params: { itemId, projectId, fileName, mimeType, base64Data }
  const piSheet = getSheet('Project_Items');
  const qSheet  = getSheet('Quotations');
  const pSheet  = getSheet('Projects');

  if (!piSheet) return { success: false, error: 'Project_Items sheet not found.' };

  const user = Session.getActiveUser().getEmail();
  const now  = new Date();

  // ── Find the Project_Item row ─────────────────────────────
  const piData = piSheet.getDataRange().getValues();
  let piRowIndex = -1, piRow = null;
  for (let i = 1; i < piData.length; i++) {
    if (String(piData[i][PI.ITEM_ID]) === String(params.itemId)) {
      piRowIndex = i + 1;
      piRow      = piData[i];
      break;
    }
  }
  if (!piRow) return { success: false, error: 'Item not found.' };

  const itemName    = String(piRow[PI.ITEM_NAME]);
  const projectId   = String(piRow[PI.PROJECT_ID]);
  const quotationId = String(piRow[PI.QUOTATION_ID]);
  const currentStatus = String(piRow[PI.DELIVERY_STATUS] || 'Pending');

  // ── Resolve the quotation folder ──────────────────────────
  let uploadsItemFolder = null;
  let qFolder           = null;
  try {
    const qData = qSheet.getDataRange().getValues();
    let folderUrl = '';
    for (let i = 1; i < qData.length; i++) {
      if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
        folderUrl = extractUrl(qData[i][Q.FOLDER_URL]) || '';
        break;
      }
    }
    if (!folderUrl) return { success: false, error: 'Quotation folder not found.' };

    const fid          = folderUrl.match(/[-\w]{25,}/)[0];
    qFolder            = DriveApp.getFolderById(fid);
    const uploadsFolder = getOrCreateSubfolder(qFolder, 'Uploads');
    uploadsItemFolder   = getOrCreateSubfolder(uploadsFolder, itemName);

    // ── If this is a re-upload after rejection, clear old files ──
    // A re-upload is identified by: there are existing URLs AND the status
    // indicates the item was previously rejected (not just 'Uploaded' mid-batch).
    const prevUrls = String(piRow[PI.UPLOADED_FILE_URL] || '')
      .split('\n').map(s => s.trim()).filter(Boolean);
    const isRejectedReupload = prevUrls.length > 0
      && (currentStatus === 'Admin Rejected' || currentStatus === 'Client Rejected');

    if (isRejectedReupload) {
      // Move all previously uploaded files to Rejected/
      const rejectedFolder = getOrCreateSubfolder(qFolder, 'Rejected');
      const rejectedItem   = getOrCreateSubfolder(rejectedFolder, itemName);
      prevUrls.forEach(url => {
        try {
          const match = url.match(/[-\w]{25,}/);
          if (match) {
            const prevFile = DriveApp.getFileById(match[0]);
            rejectedItem.addFile(prevFile);
            uploadsItemFolder.removeFile(prevFile);
          }
        } catch(e) { Logger.log('Previous file move error: ' + e); }
      });
      // Clear the stored URLs — we'll rebuild from scratch
      piSheet.getRange(piRowIndex, PI.UPLOADED_FILE_URL + 1).setValue('');
    }

  } catch(e) {
    Logger.log('Upload folder resolve error: ' + e);
    return { success: false, error: 'Could not access upload folder: ' + e.message };
  }

  // ── Decode and save the new file ──────────────────────────
  let fileUrl = '';
  try {
    const bytes   = Utilities.base64Decode(params.base64Data);
    const blob    = Utilities.newBlob(bytes, params.mimeType, params.fileName);
    const newFile = uploadsItemFolder.createFile(blob);
    fileUrl       = newFile.getUrl();
  } catch(e) {
    Logger.log('File save error: ' + e);
    return { success: false, error: 'Could not save file: ' + e.message };
  }

  // ── Append the new URL to the stored list ─────────────────
  // Re-read the cell to get the latest value (another file in the same
  // batch may have just written to it a moment ago).
  const currentUrlsRaw = String(
    piSheet.getRange(piRowIndex, PI.UPLOADED_FILE_URL + 1).getValue() || ''
  ).trim();
  const newUrlsValue = currentUrlsRaw
    ? currentUrlsRaw + '\n' + fileUrl
    : fileUrl;

  piSheet.getRange(piRowIndex, PI.UPLOADED_FILE_URL + 1).setValue(newUrlsValue);
  piSheet.getRange(piRowIndex, PI.DELIVERY_STATUS   + 1).setValue('Uploaded');
  piSheet.getRange(piRowIndex, PI.ASSIGNED_TO       + 1).setValue(user);

  // Increment redo count only on the first file of a rejected re-upload
  const prevCount = Number(piRow[PI.REDO_COUNT] || 0);
  if (currentStatus === 'Admin Rejected' || currentStatus === 'Client Rejected') {
    if (!currentUrlsRaw) {
      // First file of this re-upload batch → increment redo
      piSheet.getRange(piRowIndex, PI.REDO_COUNT + 1).setValue(prevCount + 1);
    }
  }

  // ── Log ──────────────────────────────────────────────────
  writeLog('Projects_Log', 'Project_Items', params.itemId,
    quotationId + ' — ' + itemName,
    '', String(piRow[PI.ACCOUNT_NAME]),
    'DeliveryStatus',
    currentStatus, 'Uploaded',
    'File uploaded by ' + user + ': ' + params.fileName);

  return { success: true, fileUrl };
}

// ============================================================
//  PROJECTS — ADMIN REVIEW ITEM
//
//  On approve:
//    • ALL files in Uploads/[Item]/ are copied → Deliverables/[Item]/
//    • Status → 'Admin Approved'
//
//  On reject:
//    • ALL files in Uploads/[Item]/ are moved → Rejected/[Item]/
//    • UPLOADED_FILE_URL cleared
//    • Status → 'Admin Rejected'
//    • REDO_COUNT incremented
// ============================================================
function reviewProjectItem(params) {
  const piSheet = getSheet('Project_Items');
  const qSheet  = getSheet('Quotations');

  if (!piSheet) return { success: false, error: 'Project_Items sheet not found.' };

  const user = Session.getActiveUser().getEmail();
  const now  = new Date();

  // ── Find row ──────────────────────────────────────────────
  const piData = piSheet.getDataRange().getValues();
  let piRowIndex = -1, piRow = null;
  for (let i = 1; i < piData.length; i++) {
    if (String(piData[i][PI.ITEM_ID]) === String(params.itemId)) {
      piRowIndex = i + 1;
      piRow      = piData[i];
      break;
    }
  }
  if (!piRow) return { success: false, error: 'Item not found.' };

  const itemName    = String(piRow[PI.ITEM_NAME]);
  const quotationId = String(piRow[PI.QUOTATION_ID]);
  const fileUrlsRaw = String(piRow[PI.UPLOADED_FILE_URL] || '');
  const fileUrls    = fileUrlsRaw.split('\n').map(s => s.trim()).filter(Boolean);
  const oldStatus   = String(piRow[PI.DELIVERY_STATUS]);

  if (!fileUrls.length) return { success: false, error: 'No uploaded files to review.' };

  // ── Resolve folders ───────────────────────────────────────
  let qFolder = null;
  try {
    const qData = qSheet.getDataRange().getValues();
    let folderUrl = '';
    for (let i = 1; i < qData.length; i++) {
      if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
        folderUrl = extractUrl(qData[i][Q.FOLDER_URL]) || '';
        break;
      }
    }
    if (folderUrl) {
      qFolder = DriveApp.getFolderById(folderUrl.match(/[-\w]{25,}/)[0]);
    }
  } catch(e) { Logger.log('Folder resolve error: ' + e); }

  let newStatus = oldStatus;

  if (params.action === 'approve') {
    // Copy ALL files → Deliverables/[Item]/
    if (qFolder) {
      try {
        const delivFolder = getOrCreateSubfolder(qFolder, 'Deliverables');
        const delivItem   = getOrCreateSubfolder(delivFolder, itemName);
        fileUrls.forEach(url => {
          try {
            const fileId = url.match(/[-\w]{25,}/)[0];
            const file   = DriveApp.getFileById(fileId);
            file.makeCopy(file.getName(), delivItem);
          } catch(e) { Logger.log('Deliverables copy error for ' + url + ': ' + e); }
        });
      } catch(e) { Logger.log('Deliverables folder error: ' + e); }
    }
    newStatus = 'Admin Approved';

  } else if (params.action === 'reject') {
    // Move ALL files → Rejected/[Item]/
    if (qFolder) {
      try {
        const uploadsFolder  = getOrCreateSubfolder(qFolder, 'Uploads');
        const uploadsItem    = getOrCreateSubfolder(uploadsFolder, itemName);
        const rejectedFolder = getOrCreateSubfolder(qFolder, 'Rejected');
        const rejectedItem   = getOrCreateSubfolder(rejectedFolder, itemName);
        fileUrls.forEach(url => {
          try {
            const fileId = url.match(/[-\w]{25,}/)[0];
            const file   = DriveApp.getFileById(fileId);
            rejectedItem.addFile(file);
            uploadsItem.removeFile(file);
          } catch(e) { Logger.log('Reject move error for ' + url + ': ' + e); }
        });
      } catch(e) { Logger.log('Reject folder error: ' + e); }
    }
    newStatus = 'Admin Rejected';
    const currentRedo = Number(piRow[PI.REDO_COUNT] || 0);
    piSheet.getRange(piRowIndex, PI.REDO_COUNT + 1).setValue(currentRedo + 1);
    piSheet.getRange(piRowIndex, PI.UPLOADED_FILE_URL + 1).setValue('');
  }

  piSheet.getRange(piRowIndex, PI.DELIVERY_STATUS + 1).setValue(newStatus);
  if (params.internalNotes !== undefined && params.internalNotes !== null) {
    piSheet.getRange(piRowIndex, PI.INTERNAL_NOTES + 1).setValue(params.internalNotes);
  }

  writeLog('Projects_Log', 'Project_Items', params.itemId,
    quotationId + ' — ' + itemName,
    '', String(piRow[PI.ACCOUNT_NAME]),
    'DeliveryStatus', oldStatus, newStatus,
    (params.action === 'approve' ? 'Admin approved' : 'Admin rejected') + ' by ' + user
    + ' (' + fileUrls.length + ' file(s))');

  return { success: true, newStatus };
}

// ============================================================
//  PROJECTS — CLIENT REVIEW ITEM
//
//  On approve:  Status → 'Delivered', completedAt set
//  On reject:   All Deliverables/[Item]/ files moved → Rejected/[Item]/
//               UPLOADED_FILE_URL cleared, Status → 'Client Rejected'
// ============================================================
function clientReviewItem(params) {
  const piSheet = getSheet('Project_Items');
  const qSheet  = getSheet('Quotations');

  if (!piSheet) return { success: false, error: 'Project_Items sheet not found.' };

  const user = Session.getActiveUser().getEmail();
  const now  = new Date();

  const piData = piSheet.getDataRange().getValues();
  let piRowIndex = -1, piRow = null;
  for (let i = 1; i < piData.length; i++) {
    if (String(piData[i][PI.ITEM_ID]) === String(params.itemId)) {
      piRowIndex = i + 1;
      piRow      = piData[i];
      break;
    }
  }
  if (!piRow) return { success: false, error: 'Item not found.' };

  const oldStatus = String(piRow[PI.DELIVERY_STATUS]);
  if (oldStatus !== 'Admin Approved') {
    return { success: false, error: 'Item must be Admin Approved before client review.' };
  }

  const itemName    = String(piRow[PI.ITEM_NAME]);
  const quotationId = String(piRow[PI.QUOTATION_ID]);
  const fileUrlsRaw = String(piRow[PI.UPLOADED_FILE_URL] || '');
  const fileUrls    = fileUrlsRaw.split('\n').map(s => s.trim()).filter(Boolean);

  let qFolder = null;
  try {
    const qData = qSheet.getDataRange().getValues();
    let folderUrl = '';
    for (let i = 1; i < qData.length; i++) {
      if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
        folderUrl = extractUrl(qData[i][Q.FOLDER_URL]) || '';
        break;
      }
    }
    if (folderUrl) {
      qFolder = DriveApp.getFolderById(folderUrl.match(/[-\w]{25,}/)[0]);
    }
  } catch(e) { Logger.log('Folder resolve error: ' + e); }

  let newStatus = oldStatus;

  if (params.action === 'approve') {
    newStatus = 'Delivered';
    piSheet.getRange(piRowIndex, PI.COMPLETED_AT + 1).setValue(now);
    autoCompleteProject(String(piRow[PI.PROJECT_ID]));

  } else if (params.action === 'reject') {
    newStatus = 'Client Rejected';
    // Move all files from Deliverables/[Item]/ → Rejected/[Item]/
    if (qFolder && fileUrls.length) {
      try {
        const delivFolder   = getOrCreateSubfolder(qFolder, 'Deliverables');
        const delivItem     = getOrCreateSubfolder(delivFolder, itemName);
        const rejectedFolder = getOrCreateSubfolder(qFolder, 'Rejected');
        const rejectedItem  = getOrCreateSubfolder(rejectedFolder, itemName);
        fileUrls.forEach(url => {
          try {
            const fileId = url.match(/[-\w]{25,}/)[0];
            const file   = DriveApp.getFileById(fileId);
            rejectedItem.addFile(file);
            delivItem.removeFile(file);
          } catch(e) { Logger.log('Client reject move error for ' + url + ': ' + e); }
        });
      } catch(e) { Logger.log('Client reject folder error: ' + e); }
    }
    const currentRedo = Number(piRow[PI.REDO_COUNT] || 0);
    piSheet.getRange(piRowIndex, PI.REDO_COUNT + 1).setValue(currentRedo + 1);
    piSheet.getRange(piRowIndex, PI.UPLOADED_FILE_URL + 1).setValue('');
  }

  piSheet.getRange(piRowIndex, PI.DELIVERY_STATUS + 1).setValue(newStatus);
  if (params.internalNotes !== undefined && params.internalNotes !== null) {
    piSheet.getRange(piRowIndex, PI.INTERNAL_NOTES + 1).setValue(params.internalNotes);
  }

  writeLog('Projects_Log', 'Project_Items', params.itemId,
    quotationId + ' — ' + itemName,
    '', String(piRow[PI.ACCOUNT_NAME]),
    'DeliveryStatus', oldStatus, newStatus,
    (params.action === 'approve' ? 'Client approved' : 'Client rejected') + ' by ' + user);

  return { success: true, newStatus };
}

// ============================================================
//  PROJECTS — AUTO-COMPLETE
// ============================================================
function autoCompleteProject(projectId) {
  const piSheet = getSheet('Project_Items');
  const pSheet  = getSheet('Projects');
  if (!piSheet || !pSheet) return;

  const piData = piSheet.getDataRange().getValues();
  const items  = piData.slice(1).filter(r => String(r[PI.PROJECT_ID]) === String(projectId));
  const allDelivered = items.length > 0
    && items.every(r => String(r[PI.DELIVERY_STATUS]) === 'Delivered');

  if (!allDelivered) return;

  const pData = pSheet.getDataRange().getValues();
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) {
      pSheet.getRange(i + 1, P.STATUS       + 1).setValue('Completed');
      pSheet.getRange(i + 1, P.COMPLETED_AT + 1).setValue(new Date());
      writeLog('Projects_Log', 'Projects', projectId,
        String(pData[i][P.QUOTATION_ID]) + ' — ' + String(pData[i][P.PROJECT_NAME]),
        String(pData[i][P.ACCOUNT_ID]), String(pData[i][P.ACCOUNT_NAME]),
        'Status', 'Active', 'Completed', 'All items delivered');
      break;
    }
  }
}

// ============================================================
//  PROJECTS — UPDATE INTERNAL NOTES
// ============================================================
function updateItemNotes(itemId, notes) {
  const piSheet = getSheet('Project_Items');
  if (!piSheet) return { success: false, error: 'Sheet not found.' };
  const piData = piSheet.getDataRange().getValues();
  for (let i = 1; i < piData.length; i++) {
    if (String(piData[i][PI.ITEM_ID]) === String(itemId)) {
      piSheet.getRange(i + 1, PI.INTERNAL_NOTES + 1).setValue(notes || '');
      return { success: true };
    }
  }
  return { success: false, error: 'Item not found.' };
}

// ============================================================
//  PROJECTS — GET DELIVERABLES FOLDER URL
// ============================================================
function getDeliverablesUrl(quotationId) {
  const qSheet = getSheet('Quotations');
  if (!qSheet) return { success: false, url: '' };

  const qData = qSheet.getDataRange().getValues();
  let folderUrl = '';
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      folderUrl = extractUrl(qData[i][Q.FOLDER_URL]) || '';
      break;
    }
  }
  if (!folderUrl) return { success: false, url: '' };

  try {
    const fid          = folderUrl.match(/[-\w]{25,}/)[0];
    const qFolder      = DriveApp.getFolderById(fid);
    const delivFolders = qFolder.getFoldersByName('Deliverables');
    if (!delivFolders.hasNext()) return { success: false, url: '' };
    return { success: true, url: delivFolders.next().getUrl() };
  } catch(e) {
    Logger.log('getDeliverablesUrl error: ' + e);
    return { success: false, url: '' };
  }
}
