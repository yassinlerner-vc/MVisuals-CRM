// ============================================================
//  PROJECTS — LIST
// ============================================================
function getProjects() {
  try {
    const sheet = getSheet('Projects');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const tz = Session.getScriptTimeZone();

    // Safe date formatter — never throws
    function fmtDate(val) {
      if (!val) return '';
      try { return Utilities.formatDate(new Date(val), tz, 'yyyy-MM-dd'); } catch(e) { return ''; }
    }

    // Assignment summary — always safe, returns {} if anything goes wrong
    let assignmentSummary = {};
    try { assignmentSummary = getAssignmentSummaryByProject(); } catch(e) {
      Logger.log('getProjects: assignmentSummary failed: ' + e);
    }

    const projects = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[P.PROJECT_ID]) continue;
      try {
        const projectId = String(row[P.PROJECT_ID]);
        const summary   = assignmentSummary[projectId] || { totalQty: 0, assignedQty: 0 };
        const status    = String(row[P.STATUS] || 'Active');

        let displayStatus = status;
        if (status === 'Active') {
          displayStatus = (summary.totalQty > 0 && summary.assignedQty >= summary.totalQty)
            ? 'In Progress' : 'Needs Assignment';
        }

        const createdAt = fmtDate(row[P.CREATED_AT]);
        projects.push({
          projectId,
          quotationId:      String(row[P.QUOTATION_ID]   || ''),
          accountId:        String(row[P.ACCOUNT_ID]     || ''),
          accountName:      String(row[P.ACCOUNT_NAME]   || ''),
          projectName:      String(row[P.PROJECT_NAME]   || ''),
          projectDesc:      String(row[P.PROJECT_DESC]   || ''),
          deliveryDeadline: fmtDate(row[P.DELIVERY_DEADLINE]),
          dueDate:          fmtDate(row[P.DUE_DATE]),
          status,
          displayStatus,
          internalNotes:    String(row[P.INTERNAL_NOTES] || ''),
          createdAt,
          completedAt:      fmtDate(row[P.COMPLETED_AT]),
          totalQty:         summary.totalQty,
          assignedQty:      summary.assignedQty
        });
      } catch(e) {
        Logger.log('getProjects: skipping row ' + i + ': ' + e);
      }
    }

    const order = { 'Needs Assignment': 0, 'In Progress': 1, 'Completed': 2 };
    projects.sort((a, b) => {
      const ao = order[a.displayStatus] !== undefined ? order[a.displayStatus] : 1;
      const bo = order[b.displayStatus] !== undefined ? order[b.displayStatus] : 1;
      if (ao !== bo) return ao - bo;
      return b.createdAt.localeCompare(a.createdAt);
    });

    return projects;
  } catch(e) {
    Logger.log('getProjects fatal error: ' + e);
    return [];
  }
}

// ============================================================
//  PROJECTS — GET ASSIGNMENT SUMMARY PER PROJECT
//  Returns { projectId: { totalQty, assignedQty } }
// ============================================================
function getAssignmentSummaryByProject() {
  const summary = {};

  try {
    const piSheet = getSheet('Project_Items');
    if (!piSheet) return summary;

    // Total qty per project from Project_Items
    const piData = piSheet.getDataRange().getValues();
    if (piData.length > 1) {
      piData.slice(1).forEach(row => {
        if (!row[PI.PROJECT_ID]) return;
        const pid = String(row[PI.PROJECT_ID]);
        if (!summary[pid]) summary[pid] = { totalQty: 0, assignedQty: 0 };
        summary[pid].totalQty += Number(row[PI.QUANTITY] || 0);
      });
    }
  } catch(e) {
    Logger.log('getAssignmentSummaryByProject PI error: ' + e);
  }

  try {
    // Assigned qty per project from Assignments sheet (may not exist yet)
    const aSheet = getSheet('Assignments');
    if (aSheet && aSheet.getLastRow() > 1) {
      const aData = aSheet.getDataRange().getValues();
      aData.slice(1).forEach(row => {
        if (!row[AS.PROJECT_ID]) return;
        const pid = String(row[AS.PROJECT_ID]);
        if (!summary[pid]) summary[pid] = { totalQty: 0, assignedQty: 0 };
        summary[pid].assignedQty += Number(row[AS.QUANTITY_ASSIGNED] || 0);
      });
    }
  } catch(e) {
    Logger.log('getAssignmentSummaryByProject AS error: ' + e);
  }

  return summary;
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
      projectDesc:        String(row[PI.PROJECT_DESC] || ''),
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
//  PROJECTS — GET FILE PREVIEWS FOR AN ITEM
//  Returns array of { fileId, name, mimeType, thumbUrl, viewUrl, isImage, isVideo }
//  Called by the review dialog to show inline previews
// ============================================================
function getItemFilePreviews(itemId) {
  const piSheet = getSheet('Project_Items');
  if (!piSheet) return [];

  const piData = piSheet.getDataRange().getValues();
  let fileUrlsRaw = '';

  for (let i = 1; i < piData.length; i++) {
    if (String(piData[i][PI.ITEM_ID]) === String(itemId)) {
      fileUrlsRaw = String(piData[i][PI.UPLOADED_FILE_URL] || '');
      break;
    }
  }

  if (!fileUrlsRaw.trim()) return [];

  const urls = fileUrlsRaw.split('\n').map(s => s.trim()).filter(Boolean);
  const previews = [];

  urls.forEach(url => {
    try {
      const match = url.match(/[-\w]{25,}/);
      if (!match) return;
      const fileId = match[0];
      const file   = DriveApp.getFileById(fileId);
      const mime   = file.getMimeType() || '';
      const name   = file.getName();

      const isImage = mime.startsWith('image/');
      const isVideo = mime.startsWith('video/');

      // Drive thumbnail URL — works for images and videos (shows first frame for video)
      const thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w600';
      const viewUrl  = 'https://drive.google.com/file/d/' + fileId + '/view';

      previews.push({ fileId, name, mimeType: mime, thumbUrl, viewUrl, isImage, isVideo });
    } catch(e) {
      Logger.log('getItemFilePreviews error for url ' + url + ': ' + e);
    }
  });

  return previews;
}

// ============================================================
//  PROJECTS — UPLOAD FILE TO ITEM
//
//  File naming: [MemberID]-[ItemNameSanitized]-[Sequence].[ext]
//  All uploads land in Uploads/[ItemName]/ inside the quotation folder.
// ============================================================
function uploadProjectItemFile(params) {
  // params: { itemId, projectId, fileName, mimeType, base64Data }
  const piSheet = getSheet('Project_Items');
  const qSheet  = getSheet('Quotations');
  const uSheet  = getSheet('Users');

  if (!piSheet) return { success: false, error: 'Project_Items sheet not found.' };

  const user = Session.getActiveUser().getEmail();
  const now  = new Date();

  // ── Get uploader's MemberID ───────────────────────────────
  let memberId = 'M000';
  if (uSheet) {
    const uData = uSheet.getDataRange().getValues();
    for (let i = 1; i < uData.length; i++) {
      if (String(uData[i][U.EMAIL]).toLowerCase() === user.toLowerCase()) {
        memberId = String(uData[i][U.MEMBER_ID] || 'M000').trim() || 'M000';
        break;
      }
    }
  }

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

  const itemName      = String(piRow[PI.ITEM_NAME]);
  const quotationId   = String(piRow[PI.QUOTATION_ID]);
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

    const fid           = folderUrl.match(/[-\w]{25,}/)[0];
    qFolder             = DriveApp.getFolderById(fid);
    const uploadsFolder = getOrCreateSubfolder(qFolder, 'Uploads');
    uploadsItemFolder   = getOrCreateSubfolder(uploadsFolder, itemName);

    // ── If re-uploading after rejection, clear old files ─────
    const prevUrls = String(piRow[PI.UPLOADED_FILE_URL] || '')
      .split('\n').map(s => s.trim()).filter(Boolean);
    const isRejectedReupload = prevUrls.length > 0
      && (currentStatus === 'Admin Rejected' || currentStatus === 'Client Rejected');

    if (isRejectedReupload) {
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
      piSheet.getRange(piRowIndex, PI.UPLOADED_FILE_URL + 1).setValue('');
    }

  } catch(e) {
    Logger.log('Upload folder resolve error: ' + e);
    return { success: false, error: 'Could not access upload folder: ' + e.message };
  }

  // ── Build structured file name ────────────────────────────
  // Format: [MemberID]-[ItemName]-[Sequence].[ext]
  // Sequence = count of existing files in this item's upload folder + 1
  let sequence = 1;
  try {
    const existingFiles = uploadsItemFolder.getFiles();
    let count = 0;
    while (existingFiles.hasNext()) { existingFiles.next(); count++; }
    sequence = count + 1;
  } catch(e) { Logger.log('Sequence count error: ' + e); }

  const ext          = params.fileName.includes('.')
    ? params.fileName.split('.').pop().toLowerCase()
    : '';
  const sanitizedItem = itemName.replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_');
  const seqStr        = String(sequence).padStart(2, '0');
  const newFileName   = memberId + '-' + sanitizedItem + '-' + seqStr + (ext ? '.' + ext : '');

  // ── Decode and save the file ──────────────────────────────
  let fileUrl = '';
  try {
    const bytes   = Utilities.base64Decode(params.base64Data);
    const blob    = Utilities.newBlob(bytes, params.mimeType, newFileName);
    const newFile = uploadsItemFolder.createFile(blob);
    fileUrl       = newFile.getUrl();
  } catch(e) {
    Logger.log('File save error: ' + e);
    return { success: false, error: 'Could not save file: ' + e.message };
  }

  // ── Append URL to stored list ─────────────────────────────
  const currentUrlsRaw = String(
    piSheet.getRange(piRowIndex, PI.UPLOADED_FILE_URL + 1).getValue() || ''
  ).trim();
  const newUrlsValue = currentUrlsRaw ? currentUrlsRaw + '\n' + fileUrl : fileUrl;

  piSheet.getRange(piRowIndex, PI.UPLOADED_FILE_URL + 1).setValue(newUrlsValue);
  piSheet.getRange(piRowIndex, PI.DELIVERY_STATUS   + 1).setValue('Uploaded');
  piSheet.getRange(piRowIndex, PI.ASSIGNED_TO       + 1).setValue(user);

  // Increment redo count only on first file of a rejected re-upload
  const prevCount = Number(piRow[PI.REDO_COUNT] || 0);
  if (currentStatus === 'Admin Rejected' || currentStatus === 'Client Rejected') {
    if (!currentUrlsRaw) {
      piSheet.getRange(piRowIndex, PI.REDO_COUNT + 1).setValue(prevCount + 1);
    }
  }

  writeLog('Projects_Log', 'Project_Items', params.itemId,
    quotationId + ' — ' + itemName,
    '', String(piRow[PI.ACCOUNT_NAME]),
    'DeliveryStatus', currentStatus, 'Uploaded',
    'File uploaded by ' + user + ' as ' + newFileName);

  return { success: true, fileUrl };
}

// ============================================================
//  PROJECTS — ADMIN REVIEW ITEM (per-item)
// ============================================================
function reviewProjectItem(params) {
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

  const itemName    = String(piRow[PI.ITEM_NAME]);
  const quotationId = String(piRow[PI.QUOTATION_ID]);
  const fileUrlsRaw = String(piRow[PI.UPLOADED_FILE_URL] || '');
  const fileUrls    = fileUrlsRaw.split('\n').map(s => s.trim()).filter(Boolean);
  const oldStatus   = String(piRow[PI.DELIVERY_STATUS]);

  if (!fileUrls.length) return { success: false, error: 'No uploaded files to review.' };

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
