// ============================================================
//  ACCOUNTS — GENERATE ID
// ============================================================
function generateAccountId() {
  const sheet = getSheet('Accounts');
  const data  = sheet.getDataRange().getValues();
  let maxNum  = 0;
  data.slice(1).forEach(row => {
    const match = String(row[A.ACCOUNT_ID] || '').match(/^ACC(\d+)$/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  });
  return 'ACC' + String(maxNum + 1).padStart(6, '0');
}

// ============================================================
//  ACCOUNTS — LIST
// ============================================================
function getAccounts() {
  const sheet = getSheet('Accounts');
  const data  = sheet.getDataRange().getValues();
  return data.slice(1)
    .filter(row => row[A.ACCOUNT_ID] && row[A.STATUS] !== 'Archived')
    .map(row => ({
      id:           String(row[A.ACCOUNT_ID]),
      name:         row[A.ACCOUNT_NAME],
      source:       row[A.SOURCE],
      channel:      row[A.CHANNEL],
      country:      row[A.COUNTRY],
      industry:     row[A.INDUSTRY],
      isRegistered: row[A.IS_REGISTERED],
      status:       row[A.STATUS],
      createdBy:    row[A.CREATED_BY],
      createdAt:    row[A.CREATED_AT]
        ? Utilities.formatDate(new Date(row[A.CREATED_AT]),
            Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      folderUrl:    extractUrl(row[A.FOLDER_URL]),
      notes:        row[A.NOTES]
    }))
    .sort((a, b) => b.id.localeCompare(a.id));
}

// ============================================================
//  ACCOUNTS — GET SINGLE
// ============================================================
function getAccountById(accountId) {
  const sheet = getSheet('Accounts');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][A.ACCOUNT_ID]) === String(accountId)) {
      const row = data[i];
      return {
        id:           String(row[A.ACCOUNT_ID]),
        name:         row[A.ACCOUNT_NAME],
        source:       row[A.SOURCE],
        channel:      row[A.CHANNEL],
        country:      row[A.COUNTRY],
        industry:     row[A.INDUSTRY],
        isRegistered: row[A.IS_REGISTERED],
        status:       row[A.STATUS],
        createdBy:    row[A.CREATED_BY],
        createdAt:    row[A.CREATED_AT]
          ? Utilities.formatDate(new Date(row[A.CREATED_AT]),
              Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        folderUrl:    extractUrl(row[A.FOLDER_URL]),
        notes:        row[A.NOTES]
      };
    }
  }
  return null;
}

// ============================================================
//  ACCOUNTS — GET ROW (used internally + by Quotations)
// ============================================================
function getAccountRow(accountId) {
  const sheet = getSheet('Accounts');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][A.ACCOUNT_ID]) === String(accountId)) {
      return { row: i + 1, data: data[i] };
    }
  }
  return null;
}

// ============================================================
//  ACCOUNTS — CREATE
//  Creates account folder in Leads with 2 standard subfolders:
//  Brand Assets, Projects
// ============================================================
function createAccount(formData) {
  const sheet     = getSheet('Accounts');
  const accountId = generateAccountId();
  const now       = new Date();
  const user      = Session.getActiveUser().getEmail();

  // Create account folder inside Leads
  const leadsFolder   = DriveApp.getFolderById(LEADS_FOLDER_ID);
  const accountFolder = leadsFolder.createFolder(accountId + ' - ' + formData.name);

  // Create standard subfolders
  accountFolder.createFolder('Brand Assets');
  accountFolder.createFolder('Projects');

  const folderUrl = accountFolder.getUrl();

  sheet.appendRow([
    accountId,
    formData.name,
    formData.source,
    formData.channel,
    formData.country,
    formData.industry,
    formData.isRegistered,
    'Lead',
    user,
    now,
    folderUrl,
    formData.notes || ''
  ]);

  writeLog('Accounts_Log', 'Accounts', accountId,
    accountId + ' — ' + formData.name,
    accountId, formData.name,
    'Status', '', 'Lead', 'Account created');

  return { success: true, accountId, accountName: formData.name, folderUrl };
}

// ============================================================
//  ACCOUNTS — EDIT
// ============================================================
function editAccount(formData) {
  const sheet = getSheet('Accounts');
  const data  = sheet.getDataRange().getValues();
  const user  = Session.getActiveUser().getEmail();
  const now   = new Date();

  let rowIndex = -1, oldRow = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][A.ACCOUNT_ID]) === String(formData.id)) {
      rowIndex = i + 1;
      oldRow   = data[i];
      break;
    }
  }
  if (rowIndex === -1) return { success: false, error: 'Account not found.' };

  // Log field-level changes
  const fields = [
    { key: 'AccountName', old: oldRow[A.ACCOUNT_NAME],  nw: formData.name     },
    { key: 'Source',      old: oldRow[A.SOURCE],         nw: formData.source   },
    { key: 'Channel',     old: oldRow[A.CHANNEL],        nw: formData.channel  },
    { key: 'Country',     old: oldRow[A.COUNTRY],        nw: formData.country  },
    { key: 'Industry',    old: oldRow[A.INDUSTRY],       nw: formData.industry },
    { key: 'IsRegistered',old: oldRow[A.IS_REGISTERED],  nw: formData.isRegistered },
    { key: 'Notes',       old: oldRow[A.NOTES],          nw: formData.notes    }
  ];

  fields.forEach(f => {
    if (String(f.old) !== String(f.nw)) {
      writeLog('Accounts_Log', 'Accounts',
        formData.id, formData.id + ' — ' + formData.name,
        formData.id, formData.name,
        f.key, f.old, f.nw, '');
    }
  });

  // Update row — preserve ID, status, createdBy, createdAt, folderUrl
  sheet.getRange(rowIndex, A.ACCOUNT_NAME  + 1).setValue(formData.name);
  sheet.getRange(rowIndex, A.SOURCE        + 1).setValue(formData.source);
  sheet.getRange(rowIndex, A.CHANNEL       + 1).setValue(formData.channel);
  sheet.getRange(rowIndex, A.COUNTRY       + 1).setValue(formData.country);
  sheet.getRange(rowIndex, A.INDUSTRY      + 1).setValue(formData.industry);
  sheet.getRange(rowIndex, A.IS_REGISTERED + 1).setValue(formData.isRegistered);
  sheet.getRange(rowIndex, A.NOTES         + 1).setValue(formData.notes || '');

  // Rename Drive folder to reflect new name if name changed
  if (String(oldRow[A.ACCOUNT_NAME]) !== String(formData.name)) {
    try {
      const folderUrl = extractUrl(oldRow[A.FOLDER_URL]);
      if (folderUrl) {
        const folderId = folderUrl.match(/[-\w]{25,}/)[0];
        DriveApp.getFolderById(folderId)
          .setName(formData.id + ' - ' + formData.name);
      }
    } catch(e) { Logger.log('Folder rename error: ' + e); }
  }

  return { success: true };
}

// ============================================================
//  ACCOUNTS — DELETE (soft — sets status to Archived,
//             moves folder to Deleted/Archived Drive folder)
// ============================================================
function deleteAccount(accountId) {
  const sheet = getSheet('Accounts');
  const data  = sheet.getDataRange().getValues();
  const user  = Session.getActiveUser().getEmail();
  const now   = new Date();

  let rowIndex = -1, oldRow = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][A.ACCOUNT_ID]) === String(accountId)) {
      rowIndex = i + 1;
      oldRow   = data[i];
      break;
    }
  }
  if (rowIndex === -1) return { success: false, error: 'Account not found.' };
  if (oldRow[A.STATUS] === 'Archived') {
    return { success: false, error: 'Account is already archived.' };
  }

  // Move account folder to Deleted/Archived
  const folderUrl = extractUrl(oldRow[A.FOLDER_URL]);
  if (folderUrl) {
    try {
      const folderId      = folderUrl.match(/[-\w]{25,}/)[0];
      const folder        = DriveApp.getFolderById(folderId);
      const archivedRoot  = DriveApp.getFolderById(ARCHIVED_FOLDER_ID);
      const parents       = folder.getParents();
      archivedRoot.addFile(folder);
      while (parents.hasNext()) parents.next().removeFile(folder);
    } catch(e) { Logger.log('Archive folder error: ' + e); }
  }

  // Soft delete — update status only
  sheet.getRange(rowIndex, A.STATUS + 1).setValue('Archived');

  writeLog('Accounts_Log', 'Accounts', accountId,
    accountId + ' — ' + oldRow[A.ACCOUNT_NAME],
    accountId, oldRow[A.ACCOUNT_NAME],
    'Status', oldRow[A.STATUS], 'Archived', 'Account archived');

  return { success: true };
}

// ============================================================
//  ACCOUNTS — PROMOTE TO CLIENT (called by Quotations on confirm)
// ============================================================
function promoteAccountToClient(accountId) {
  const sheet = getSheet('Accounts');
  const found = getAccountRow(accountId);
  if (!found) return;

  const { row, data } = found;
  if (data[A.STATUS] === 'Client') return;

  // Move account folder from Leads to Clients
  const folderUrl = extractUrl(data[A.FOLDER_URL]);
  if (folderUrl) {
    try {
      const folderId      = folderUrl.match(/[-\w]{25,}/)[0];
      const folder        = DriveApp.getFolderById(folderId);
      const clientsFolder = DriveApp.getFolderById(CLIENTS_FOLDER_ID);
      const parents       = folder.getParents();
      clientsFolder.addFolder(folder);
      while (parents.hasNext()) parents.next().removeFolder(folder);
    } catch(e) { Logger.log('Promote folder error: ' + e); }
  }

  sheet.getRange(row, A.STATUS + 1).setValue('Client');

  writeLog('Accounts_Log', 'Accounts', accountId,
    accountId + ' — ' + data[A.ACCOUNT_NAME],
    accountId, data[A.ACCOUNT_NAME],
    'Status', data[A.STATUS], 'Client',
    'Promoted on quotation confirmation');
}

// ============================================================
//  ACCOUNTS — MIGRATE EXISTING (kept for safety — backfills
//  accounts that never got an ID/folder)
// ============================================================
function migrateExistingAccounts() {
  const sheet        = getSheet('Accounts');
  const leadsFolder  = DriveApp.getFolderById(LEADS_FOLDER_ID);
  if (!sheet || !leadsFolder) {
    SpreadsheetApp.getUi().alert('Accounts sheet or Leads folder missing.');
    return;
  }

  const data    = sheet.getDataRange().getValues();
  let   updated = 0;

  for (let i = 1; i < data.length; i++) {
    const row         = data[i];
    const existingId  = row[A.ACCOUNT_ID];
    const accountName = row[A.ACCOUNT_NAME];
    if (!existingId && accountName) {
      const accountId     = generateAccountId();
      const accountFolder = leadsFolder.createFolder(accountId + ' - ' + accountName);
      accountFolder.createFolder('Brand Assets');
      accountFolder.createFolder('Projects');
      sheet.getRange(i + 1, A.ACCOUNT_ID  + 1).setValue(accountId);
      sheet.getRange(i + 1, A.FOLDER_URL  + 1).setValue(accountFolder.getUrl());
      sheet.getRange(i + 1, A.STATUS      + 1).setValue(row[A.STATUS] || 'Lead');
      updated++;
    }
  }
  SpreadsheetApp.getUi().alert(
    'Migration complete. ' + updated + ' accounts updated.');
}
// ============================================================
//  HELPER — extract plain URL from either a formula or raw string
// ============================================================
function extractUrl(cellValue) {
  if (!cellValue) return '';
  const str = String(cellValue);
  // Handle =HYPERLINK("url","label") formula
  if (str.toUpperCase().startsWith('=HYPERLINK')) {
    const match = str.match(/\"(https?:\/\/[^\"]+)\"/);
    return match ? match[1] : '';
  }
  return str;
}
