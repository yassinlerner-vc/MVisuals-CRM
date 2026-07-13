// ============================================================
//  TEAM — GENERATE ID
// ============================================================
function generateTeamId() {
  const sheet = getSheet('Team');
  const data  = sheet.getDataRange().getValues();
  let maxNum  = 0;
  data.slice(1).forEach(row => {
    const match = String(row[TM.TEAM_ID] || '').match(/^TM(\d+)$/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  });
  return 'TM' + String(maxNum + 1).padStart(6, '0');
}

// ============================================================
//  TEAM — LIST
// ============================================================
function getTeam() {
  const sheet = getSheet('Team');
  const data  = sheet.getDataRange().getValues();
  const tz    = Session.getScriptTimeZone();
  return data.slice(1)
    .filter(row => row[TM.TEAM_ID])
    .map(row => ({
      id:        String(row[TM.TEAM_ID]),
      name:      row[TM.NAME],
      phone:     row[TM.PHONE],
      email:     row[TM.EMAIL],
      role:      row[TM.ROLE],
      folderUrl: extractUrl(row[TM.FOLDER_URL]),
      createdAt: row[TM.CREATED_AT]
        ? Utilities.formatDate(new Date(row[TM.CREATED_AT]), tz, 'yyyy-MM-dd') : ''
    }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

// ============================================================
//  TEAM — GET SINGLE
// ============================================================
function getTeamById(teamId) {
  const sheet = getSheet('Team');
  const data  = sheet.getDataRange().getValues();
  const tz    = Session.getScriptTimeZone();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][TM.TEAM_ID]) === String(teamId)) {
      const row = data[i];
      return {
        id:        String(row[TM.TEAM_ID]),
        name:      row[TM.NAME],
        phone:     row[TM.PHONE],
        email:     row[TM.EMAIL],
        role:      row[TM.ROLE],
        folderUrl: extractUrl(row[TM.FOLDER_URL]),
        createdAt: row[TM.CREATED_AT]
          ? Utilities.formatDate(new Date(row[TM.CREATED_AT]), tz, 'yyyy-MM-dd') : ''
      };
    }
  }
  return null;
}

// ============================================================
//  TEAM — CREATE
//  Creates "[Name]_Ruya" folder inside Ruya_Team, shares it
//  with the member's email as an editor (folder-only access —
//  Drive sharing never grants visibility into sibling folders).
// ============================================================
function createTeam(formData) {
  const sheet   = getSheet('Team');
  const teamId  = generateTeamId();
  const now     = new Date();

  const teamRoot     = DriveApp.getFolderById(RUYA_TEAM_FOLDER_ID);
  const memberFolder = teamRoot.createFolder(formData.name + '_Ruya');

  try {
    memberFolder.addEditor(formData.email);
  } catch(e) {
    Logger.log('Team share error: ' + e);
    // Folder still gets created even if sharing fails (e.g. bad email) —
    // caller can fix the email and we retry sharing on edit.
  }

  const folderUrl = memberFolder.getUrl();

  sheet.appendRow([
    teamId, formData.name, formData.phone, formData.email, formData.role,
    folderUrl, now
  ]);

  writeLog('Team_Log', 'Team', teamId, teamId + ' — ' + formData.name,
    '', '', 'Status', '', 'Created', 'Team member added');

  return { success: true, teamId, folderUrl };
}

// ============================================================
//  TEAM — EDIT
//  Renames folder if name changed. Re-shares folder if email changed
//  (removes old editor, adds new one).
// ============================================================
function editTeam(formData) {
  const sheet = getSheet('Team');
  const data  = sheet.getDataRange().getValues();

  let rowIndex = -1, oldRow = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][TM.TEAM_ID]) === String(formData.id)) {
      rowIndex = i + 1;
      oldRow   = data[i];
      break;
    }
  }
  if (rowIndex === -1) return { success: false, error: 'Team member not found.' };

  const fields = [
    { key: 'Name',  old: oldRow[TM.NAME],  nw: formData.name  },
    { key: 'Phone', old: oldRow[TM.PHONE], nw: formData.phone },
    { key: 'Email', old: oldRow[TM.EMAIL], nw: formData.email },
    { key: 'Role',  old: oldRow[TM.ROLE],  nw: formData.role  }
  ];
  fields.forEach(f => {
    if (String(f.old) !== String(f.nw)) {
      writeLog('Team_Log', 'Team', formData.id, formData.id + ' — ' + formData.name,
        '', '', f.key, f.old, f.nw, '');
    }
  });

  sheet.getRange(rowIndex, TM.NAME  + 1).setValue(formData.name);
  sheet.getRange(rowIndex, TM.PHONE + 1).setValue(formData.phone);
  sheet.getRange(rowIndex, TM.EMAIL + 1).setValue(formData.email);
  sheet.getRange(rowIndex, TM.ROLE  + 1).setValue(formData.role);

  const folderUrl = extractUrl(oldRow[TM.FOLDER_URL]);
  if (folderUrl) {
    try {
      const fid    = folderUrl.match(/[-\w]{25,}/)[0];
      const folder = DriveApp.getFolderById(fid);

      if (String(oldRow[TM.NAME]) !== String(formData.name)) {
        folder.setName(formData.name + '_Ruya');
      }
      if (String(oldRow[TM.EMAIL]) !== String(formData.email)) {
        try { folder.removeEditor(String(oldRow[TM.EMAIL])); } catch(e) { /* not shared yet, ignore */ }
        try { folder.addEditor(formData.email); } catch(e) { Logger.log('Re-share error: ' + e); }
      }
    } catch(e) { Logger.log('Team folder update error: ' + e); }
  }

  return { success: true };
}

// ============================================================
//  TEAM — DELETE (hard — moves folder to Archived, removes row)
//  No Status field, so removal is permanent in the sheet;
//  the Drive folder is preserved (archived) rather than deleted.
// ============================================================
function deleteTeam(teamId) {
  const sheet = getSheet('Team');
  const data  = sheet.getDataRange().getValues();

  let rowIndex = -1, oldRow = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][TM.TEAM_ID]) === String(teamId)) {
      rowIndex = i + 1;
      oldRow   = data[i];
      break;
    }
  }
  if (rowIndex === -1) return { success: false, error: 'Team member not found.' };

  const folderUrl = extractUrl(oldRow[TM.FOLDER_URL]);
  if (folderUrl) {
    try {
      const fid          = folderUrl.match(/[-\w]{25,}/)[0];
      const folder        = DriveApp.getFolderById(fid);
      const archivedRoot  = DriveApp.getFolderById(ARCHIVED_FOLDER_ID);
      const parents        = folder.getParents();
      archivedRoot.addFolder(folder);
      while (parents.hasNext()) parents.next().removeFolder(folder);
    } catch(e) { Logger.log('Team archive folder error: ' + e); }
  }

  sheet.deleteRow(rowIndex);

  writeLog('Team_Log', 'Team', teamId, teamId + ' — ' + oldRow[TM.NAME],
    '', '', 'Status', 'Active', 'Removed', 'Team member removed, folder archived');

  return { success: true };
}
