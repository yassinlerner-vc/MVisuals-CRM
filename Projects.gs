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

    function fmtDate(val) {
      if (!val) return '';
      try { return Utilities.formatDate(new Date(val), tz, 'yyyy-MM-dd'); } catch(e) { return ''; }
    }

    const projects = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[P.PROJECT_ID]) continue;
      try {
        const projectId    = String(row[P.PROJECT_ID]);
        const status       = String(row[P.STATUS] || 'Active');
        const displayStatus = status; // Pending / Active / Completed — set by assignment flow later

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
          createdAt:        fmtDate(row[P.CREATED_AT]),
          completedAt:      fmtDate(row[P.COMPLETED_AT])
        });
      } catch(e) {
        Logger.log('getProjects: skipping row ' + i + ': ' + e);
      }
    }

    const order = { 'Active': 0, 'Pending': 1, 'Completed': 2 };
    projects.sort((a, b) => {
      const ao = order[a.displayStatus] !== undefined ? order[a.displayStatus] : 0;
      const bo = order[b.displayStatus] !== undefined ? order[b.displayStatus] : 0;
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
//  PROJECTS — GET ITEMS FOR A PROJECT
// ============================================================
function getProjectItems(projectId) {
  const piSheet = getSheet('Project_Items');
  if (!piSheet) return { items: [] };

  const piData = piSheet.getDataRange().getValues();
  const tz     = Session.getScriptTimeZone();

  const items = piData.slice(1)
    .filter(row => String(row[PI.PROJECT_ID]) === String(projectId))
    .map(row => ({
      itemId:      String(row[PI.ITEM_ID]),
      projectId:   String(row[PI.PROJECT_ID]),
      quotationId: String(row[PI.QUOTATION_ID]),
      itemName:    String(row[PI.ITEM_NAME]),
      quantity:    Number(row[PI.QUANTITY] || 0),
      description: String(row[PI.DESCRIPTION]   || ''),
      notes:       String(row[PI.NOTES]          || ''),
      internalNotes: String(row[PI.INTERNAL_NOTES] || ''),
      createdAt:   row[PI.CREATED_AT]
        ? Utilities.formatDate(new Date(row[PI.CREATED_AT]), tz, 'yyyy-MM-dd') : ''
    }));

  return { items };
}
