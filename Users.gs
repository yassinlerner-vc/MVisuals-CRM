// ============================================================
//  AUTH — GET CURRENT USER
//  Called by every module on init.
//  Returns { email, name, role } or throws if not found / inactive.
// ============================================================
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  const sheet = getSheet('Users');
  if (!sheet) throw new Error('Users sheet not found. Please contact your administrator.');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[U.EMAIL]).toLowerCase().trim() === email.toLowerCase().trim()) {
      if (String(row[U.STATUS]) === 'Inactive') {
        throw new Error('Your account has been deactivated. Please contact your administrator.');
      }
      return {
        email:  String(row[U.EMAIL]),
        name:   String(row[U.NAME]),
        role:   String(row[U.ROLE])   // 'Admin' | 'Team' | 'Lead'
      };
    }
  }
  throw new Error('Access denied. Your Google account (' + email + ') is not registered in this system. Please contact your administrator.');
}

// ============================================================
//  USERS — LIST (Admin only — enforced in UI)
// ============================================================
function getUsers() {
  getCurrentUser(); // auth check
  const sheet = getSheet('Users');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const tz   = Session.getScriptTimeZone();
  return data.slice(1)
    .filter(row => row[U.EMAIL])
    .map(row => ({
      email:     String(row[U.EMAIL]),
      name:      String(row[U.NAME]),
      role:      String(row[U.ROLE]),
      status:    String(row[U.STATUS]),
      createdBy: String(row[U.CREATED_BY] || ''),
      createdAt: row[U.CREATED_AT]
        ? Utilities.formatDate(new Date(row[U.CREATED_AT]),
            Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      notes:     String(row[U.NOTES] || '')
    }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

// ============================================================
//  USERS — CREATE
// ============================================================
function createUser(formData) {
  const actor = getCurrentUser();
  if (actor.role !== 'Admin') return { success: false, error: 'Only Admins can create users.' };

  const sheet = getSheet('Users');
  const data  = sheet.getDataRange().getValues();
  const now   = new Date();

  // Check duplicate email
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][U.EMAIL]).toLowerCase().trim() ===
        String(formData.email).toLowerCase().trim()) {
      return { success: false, error: 'A user with this email already exists.' };
    }
  }

  sheet.appendRow([
    formData.email.trim(),
    formData.name.trim(),
    formData.role,
    'Active',
    actor.email,
    now,
    formData.notes || ''
  ]);

  writeLog('Users_Log', 'Users', formData.email, formData.name,
    '', '', 'Status', '', 'Active',
    'User created with role ' + formData.role + ' by ' + actor.email);

  return { success: true };
}

// ============================================================
//  USERS — EDIT
// ============================================================
function editUser(formData) {
  const actor = getCurrentUser();
  if (actor.role !== 'Admin') return { success: false, error: 'Only Admins can edit users.' };

  const sheet = getSheet('Users');
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][U.EMAIL]).toLowerCase().trim() ===
        String(formData.email).toLowerCase().trim()) {

      const oldRole   = String(data[i][U.ROLE]);
      const oldStatus = String(data[i][U.STATUS]);

      sheet.getRange(i + 1, U.NAME   + 1).setValue(formData.name.trim());
      sheet.getRange(i + 1, U.ROLE   + 1).setValue(formData.role);
      sheet.getRange(i + 1, U.STATUS + 1).setValue(formData.status);
      sheet.getRange(i + 1, U.NOTES  + 1).setValue(formData.notes || '');

      if (oldRole !== formData.role) {
        writeLog('Users_Log', 'Users', formData.email, formData.name,
          '', '', 'Role', oldRole, formData.role, 'Role changed by ' + actor.email);
      }
      if (oldStatus !== formData.status) {
        writeLog('Users_Log', 'Users', formData.email, formData.name,
          '', '', 'Status', oldStatus, formData.status, 'Status changed by ' + actor.email);
      }
      return { success: true };
    }
  }
  return { success: false, error: 'User not found.' };
}

// ============================================================
//  USERS — DEACTIVATE
// ============================================================
function deactivateUser(email) {
  const actor = getCurrentUser();
  if (actor.role !== 'Admin') return { success: false, error: 'Only Admins can deactivate users.' };
  if (actor.email.toLowerCase() === email.toLowerCase()) {
    return { success: false, error: 'You cannot deactivate your own account.' };
  }

  const sheet = getSheet('Users');
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][U.EMAIL]).toLowerCase().trim() === email.toLowerCase().trim()) {
      sheet.getRange(i + 1, U.STATUS + 1).setValue('Inactive');
      writeLog('Users_Log', 'Users', email, String(data[i][U.NAME]),
        '', '', 'Status', 'Active', 'Inactive', 'Deactivated by ' + actor.email);
      return { success: true };
    }
  }
  return { success: false, error: 'User not found.' };
}
