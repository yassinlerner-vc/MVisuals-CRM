// ============================================================
//  PROJECTS — LIST
//  Admins see all projects. Team members call getMyProjects()
//  in Assignments.gs instead.
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
        projects.push({
          projectId:        String(row[P.PROJECT_ID]),
          quotationId:      String(row[P.QUOTATION_ID]      || ''),
          accountId:        String(row[P.ACCOUNT_ID]        || ''),
          accountName:      String(row[P.ACCOUNT_NAME]      || ''),
          projectName:      String(row[P.PROJECT_NAME]      || ''),
          projectDesc:      String(row[P.PROJECT_DESC]      || ''),
          deliveryDeadline: fmtDate(row[P.DELIVERY_DEADLINE]),
          dueDate:          fmtDate(row[P.DUE_DATE]),
          status:           String(row[P.STATUS]            || 'Needs Assignment'),
          totalAmount:      Number(row[P.TOTAL_AMOUNT]      || 0),
          currency:         String(row[P.CURRENCY]          || 'EGP'),
          totalCommission:  Number(row[P.TOTAL_COMMISSION]  || 0),
          remainingAmount:  Number(row[P.REMAINING_AMOUNT]  || 0),
          createdAt:        fmtDate(row[P.CREATED_AT]),
          completedAt:      fmtDate(row[P.COMPLETED_AT])
        });
      } catch(e) {
        Logger.log('getProjects: skipping row ' + i + ': ' + e);
      }
    }

    const order = { 'Needs Assignment': 0, 'Assigned': 1, 'Delivered': 2 };
    projects.sort((a, b) => {
      const ao = order[a.status] !== undefined ? order[a.status] : 1;
      const bo = order[b.status] !== undefined ? order[b.status] : 1;
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
//  PROJECTS — GET ITEMS (lightweight, for non-assignment views)
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
      description: String(row[PI.DESCRIPTION] || ''),
      notes:       String(row[PI.NOTES]        || ''),
      createdAt:   row[PI.CREATED_AT]
        ? Utilities.formatDate(new Date(row[PI.CREATED_AT]), tz, 'yyyy-MM-dd') : ''
    }));

  return { items };
}
