// ============================================================
//  COLUMN INDEXES — Assignments (0-based)
//
//  Sheet headers (row 1):
//  Assignment ID | Project ID | Quotation ID | Account Name |
//  Project Name | Item Name | Assigned To Email | Assigned To Name |
//  Quantity | Status | Created By | Created At
// ============================================================
const ASN = {
  ASSIGNMENT_ID:     0,
  PROJECT_ID:        1,
  QUOTATION_ID:      2,
  ACCOUNT_NAME:      3,
  PROJECT_NAME:      4,
  ITEM_NAME:         5,
  ASSIGNED_TO_EMAIL: 6,
  ASSIGNED_TO_NAME:  7,
  QUANTITY:          8,
  STATUS:            9,   // Active | Superseded
  CREATED_BY:        10,
  CREATED_AT:        11
};

// ============================================================
//  COLUMN INDEXES — Commissions (0-based)
//
//  Sheet headers (row 1):
//  Commission ID | Project ID | Quotation ID | Account Name |
//  Project Name | Currency | Person Email | Person Name |
//  Person Role | Commission % | Commission Amount |
//  Payment Status | Paid At | Notes
// ============================================================
const COM = {
  COMMISSION_ID:    0,
  PROJECT_ID:       1,
  QUOTATION_ID:     2,
  ACCOUNT_NAME:     3,
  PROJECT_NAME:     4,
  CURRENCY:         5,
  PERSON_EMAIL:     6,
  PERSON_NAME:      7,
  PERSON_ROLE:      8,   // Team | Admin | Owner
  COMMISSION_PCT:   9,
  COMMISSION_AMT:   10,
  PAYMENT_STATUS:   11,  // Unpaid | Paid
  PAID_AT:          12,
  NOTES:            13
};

// ============================================================
//  ASSIGNMENTS — GET PROJECT DETAIL
//  Single call that powers the entire project detail view.
//  Returns project row, items with assignment summaries,
//  existing commissions, and team members list.
// ============================================================
function getProjectDetail(projectId) {
  const pSheet  = getSheet('Projects');
  const piSheet = getSheet('Project_Items');
  const aSheet  = getSheet('Assignments');
  const cSheet  = getSheet('Commissions');
  const uSheet  = getSheet('Users');

  if (!pSheet || !piSheet) return null;
  const tz = Session.getScriptTimeZone();

  function fmtDate(val) {
    if (!val) return '';
    try { return Utilities.formatDate(new Date(val), tz, 'yyyy-MM-dd'); } catch(e) { return ''; }
  }

  // ── Project row ───────────────────────────────────────────
  const pData = pSheet.getDataRange().getValues();
  let project = null;
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) {
      const row = pData[i];
      project = {
        projectId:        String(row[P.PROJECT_ID]),
        quotationId:      String(row[P.QUOTATION_ID]   || ''),
        accountId:        String(row[P.ACCOUNT_ID]     || ''),
        accountName:      String(row[P.ACCOUNT_NAME]   || ''),
        projectName:      String(row[P.PROJECT_NAME]   || ''),
        projectDesc:      String(row[P.PROJECT_DESC]   || ''),
        deliveryDeadline: fmtDate(row[P.DELIVERY_DEADLINE]),
        dueDate:          fmtDate(row[P.DUE_DATE]),
        status:           String(row[P.STATUS]         || ''),
        totalAmount:      Number(row[P.TOTAL_AMOUNT]   || 0),
        currency:         String(row[P.CURRENCY]       || 'EGP'),
        totalCommission:  Number(row[P.TOTAL_COMMISSION] || 0),
        remainingAmount:  Number(row[P.REMAINING_AMOUNT] || 0)
      };
      break;
    }
  }
  if (!project) return null;

  // ── Project items ─────────────────────────────────────────
  const piData = piSheet.getDataRange().getValues();
  const items = piData.slice(1)
    .filter(row => String(row[PI.PROJECT_ID]) === String(projectId))
    .map(row => ({
      itemId:      String(row[PI.ITEM_ID]),
      itemName:    String(row[PI.ITEM_NAME]),
      quantity:    Number(row[PI.QUANTITY] || 0),
      description: String(row[PI.DESCRIPTION] || ''),
      notes:       String(row[PI.NOTES]        || '')
    }));

  // ── Active assignments grouped by item ────────────────────
  let assignments = [];
  if (aSheet) {
    const aData = aSheet.getDataRange().getValues();
    assignments = aData.slice(1)
      .filter(row =>
        String(row[ASN.PROJECT_ID]) === String(projectId) &&
        String(row[ASN.STATUS])     === 'Active')
      .map(row => ({
        assignmentId:    String(row[ASN.ASSIGNMENT_ID]),
        itemName:        String(row[ASN.ITEM_NAME]),
        assignedToEmail: String(row[ASN.ASSIGNED_TO_EMAIL]),
        assignedToName:  String(row[ASN.ASSIGNED_TO_NAME]),
        quantity:        Number(row[ASN.QUANTITY] || 0)
      }));
  }

  // Merge assignment summaries into items
  items.forEach(item => {
    const itemAssignments = assignments.filter(a => a.itemName === item.itemName);
    const assignedQty = itemAssignments.reduce((sum, a) => sum + a.quantity, 0);
    item.assignedQty    = assignedQty;
    item.unassignedQty  = Math.max(0, item.quantity - assignedQty);
    item.assignments    = itemAssignments;
  });

  // ── Existing commissions ──────────────────────────────────
  let commissions = [];
  if (cSheet) {
    const cData = cSheet.getDataRange().getValues();
    commissions = cData.slice(1)
      .filter(row => String(row[COM.PROJECT_ID]) === String(projectId))
      .map(row => ({
        commissionId:   String(row[COM.COMMISSION_ID]),
        personEmail:    String(row[COM.PERSON_EMAIL]    || ''),
        personName:     String(row[COM.PERSON_NAME]     || ''),
        personRole:     String(row[COM.PERSON_ROLE]     || ''),
        commissionPct:  Number(row[COM.COMMISSION_PCT]  || 0),
        commissionAmt:  Number(row[COM.COMMISSION_AMT]  || 0),
        paymentStatus:  String(row[COM.PAYMENT_STATUS]  || 'Unpaid')
      }));
  }

  // ── Team members (for assignment dropdowns) ───────────────
  let teamMembers = [];
  if (uSheet) {
    const uData = uSheet.getDataRange().getValues();
    teamMembers = uData.slice(1)
      .filter(row => row[U.EMAIL] && String(row[U.STATUS]) === 'Active')
      .map(row => ({
        email: String(row[U.EMAIL]),
        name:  String(row[U.NAME]),
        role:  String(row[U.ROLE])
      }));
  }

  return { project, items, assignments, commissions, teamMembers };
}

// ============================================================
//  ASSIGNMENTS — SAVE
//  Supersedes all existing Active rows for this project,
//  writes new Active rows, auto-assigns remainder to admin.
//  assignments: [{itemName, assignedToEmail, assignedToName, quantity}]
// ============================================================
function saveAssignments(projectId, assignments) {
  const actor   = getCurrentUser();
  const aSheet  = getSheet('Assignments');
  const piSheet = getSheet('Project_Items');
  const pSheet  = getSheet('Projects');
  if (!aSheet || !piSheet) return { success: false, error: 'Required sheets missing.' };

  const now  = new Date();

  // ── Supersede existing Active rows ────────────────────────
  const aData = aSheet.getDataRange().getValues();
  for (let i = 1; i < aData.length; i++) {
    if (String(aData[i][ASN.PROJECT_ID]) === String(projectId) &&
        String(aData[i][ASN.STATUS])     === 'Active') {
      aSheet.getRange(i + 1, ASN.STATUS + 1).setValue('Superseded');
    }
  }

  // ── Get project meta for denormalisation ──────────────────
  const pData = pSheet.getDataRange().getValues();
  let quotationId = '', accountName = '', projectName = '';
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) {
      quotationId  = String(pData[i][P.QUOTATION_ID]  || '');
      accountName  = String(pData[i][P.ACCOUNT_NAME]  || '');
      projectName  = String(pData[i][P.PROJECT_NAME]  || '');
      break;
    }
  }

  // ── Get all items and their quantities ────────────────────
  const piData = piSheet.getDataRange().getValues();
  const allItems = piData.slice(1)
    .filter(row => String(row[PI.PROJECT_ID]) === String(projectId))
    .map(row => ({
      itemName: String(row[PI.ITEM_NAME]),
      quantity: Number(row[PI.QUANTITY] || 0)
    }));

  // ── Calculate assigned qty per item ───────────────────────
  const assignedMap = {};
  (assignments || []).forEach(a => {
    if (!assignedMap[a.itemName]) assignedMap[a.itemName] = 0;
    assignedMap[a.itemName] += Number(a.quantity) || 0;
  });

  // ── Build final rows: provided assignments + admin remainder ──
  const finalRows = [...(assignments || [])];
  allItems.forEach(item => {
    const assigned  = assignedMap[item.itemName] || 0;
    const remainder = item.quantity - assigned;
    if (remainder > 0) {
      finalRows.push({
        itemName:        item.itemName,
        assignedToEmail: actor.email,
        assignedToName:  actor.name,
        quantity:        remainder
      });
    }
  });

  // ── Write new Active rows ─────────────────────────────────
  finalRows.forEach(a => {
    aSheet.appendRow([
      Utilities.getUuid(),
      projectId,
      quotationId,
      accountName,
      projectName,
      a.itemName,
      a.assignedToEmail,
      a.assignedToName,
      Number(a.quantity) || 0,
      'Active',
      actor.email,
      now
    ]);
  });

  writeLog('Assignments_Log', 'Assignments', projectId,
    projectId + ' — ' + projectName,
    '', accountName,
    'Assignments', '', JSON.stringify(finalRows.map(a => ({
      item: a.itemName, to: a.assignedToName, qty: a.quantity
    }))),
    'Assignments saved by ' + actor.email);

  return { success: true };
}

// ============================================================
//  COMMISSIONS — CONFIRM ASSIGNMENTS + COMMISSIONS
//  Saves commissions, updates project totals and status,
//  sends notification emails to assigned team members.
//  commissions: [{email, name, role, percent, amount}]
//  grossProfit: number
// ============================================================
function confirmAssignmentsAndCommissions(projectId, commissions, grossProfit) {
  const actor  = getCurrentUser();
  if (actor.role !== 'Admin') return { success: false, error: 'Only Admins can confirm assignments.' };

  const cSheet  = getSheet('Commissions');
  const pSheet  = getSheet('Projects');
  const aSheet  = getSheet('Assignments');
  if (!cSheet || !pSheet) return { success: false, error: 'Required sheets missing.' };

  const now = new Date();

  // ── Get project meta ──────────────────────────────────────
  const pData = pSheet.getDataRange().getValues();
  let pRowIndex = -1, quotationId = '', accountName = '', projectName = '',
      totalAmount = 0, currency = 'EGP';
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) {
      pRowIndex   = i + 1;
      quotationId = String(pData[i][P.QUOTATION_ID]  || '');
      accountName = String(pData[i][P.ACCOUNT_NAME]  || '');
      projectName = String(pData[i][P.PROJECT_NAME]  || '');
      totalAmount = Number(pData[i][P.TOTAL_AMOUNT]  || 0);
      currency    = String(pData[i][P.CURRENCY]      || 'EGP');
      break;
    }
  }
  if (pRowIndex === -1) return { success: false, error: 'Project not found.' };

  // ── Delete existing commission rows for this project ──────
  const cData = cSheet.getDataRange().getValues();
  for (let i = cData.length - 1; i >= 1; i--) {
    if (String(cData[i][COM.PROJECT_ID]) === String(projectId)) {
      cSheet.deleteRow(i + 1);
    }
  }

  // ── Write commission rows ─────────────────────────────────
  const totalCommission = (commissions || []).reduce((sum, c) => sum + (Number(c.amount) || 0), 0);

  (commissions || []).forEach(c => {
    cSheet.appendRow([
      Utilities.getUuid(),
      projectId, quotationId, accountName, projectName,
      currency,
      c.email || '',
      c.name  || '',
      c.role  || 'Team',
      Number(c.percent) || 0,
      Number(c.amount)  || 0,
      'Unpaid', '', ''
    ]);
  });

  // Gross profit row
  cSheet.appendRow([
    Utilities.getUuid(),
    projectId, quotationId, accountName, projectName,
    currency,
    '', 'Gross Profit', 'Owner',
    '', Number(grossProfit) || 0,
    'Unpaid', '', ''
  ]);

  // ── Update Projects row ───────────────────────────────────
  pSheet.getRange(pRowIndex, P.STATUS           + 1).setValue('Assigned');
  pSheet.getRange(pRowIndex, P.TOTAL_COMMISSION + 1).setValue(totalCommission);
  pSheet.getRange(pRowIndex, P.REMAINING_AMOUNT + 1).setValue(Number(grossProfit) || 0);

  // ── Send notification emails ──────────────────────────────
  _sendAssignmentEmails(projectId, projectName, accountName, currency, commissions, aSheet);

  writeLog('Assignments_Log', 'Assignments', projectId,
    projectId + ' — ' + projectName, '', accountName,
    'Status', 'Needs Assignment', 'Assigned',
    'Confirmed by ' + actor.email);

  return { success: true };
}

// ============================================================
//  EMAIL HELPER — sends one email per unique team assignee
// ============================================================
function _sendAssignmentEmails(projectId, projectName, accountName, currency, commissions, aSheet) {
  try {
    // Build per-person assignment summary from Assignments sheet
    const aData = aSheet ? aSheet.getDataRange().getValues() : [];
    const byPerson = {};
    aData.slice(1)
      .filter(row =>
        String(row[ASN.PROJECT_ID]) === String(projectId) &&
        String(row[ASN.STATUS])     === 'Active' &&
        String(row[ASN.ASSIGNED_TO_EMAIL]) !== '')
      .forEach(row => {
        const email = String(row[ASN.ASSIGNED_TO_EMAIL]);
        if (!byPerson[email]) byPerson[email] = { name: String(row[ASN.ASSIGNED_TO_NAME]), items: [] };
        byPerson[email].items.push({
          itemName: String(row[ASN.ITEM_NAME]),
          quantity: Number(row[ASN.QUANTITY] || 0)
        });
      });

    // Map commission amounts by email
    const commMap = {};
    (commissions || []).forEach(c => { if (c.email) commMap[c.email] = c.amount; });

    Object.keys(byPerson).forEach(email => {
      const person = byPerson[email];
      const commAmt = commMap[email] || 0;

      const itemLines = person.items.map(i =>
        '  • ' + i.itemName + ' × ' + i.quantity).join('\n');

      const body =
        'Hi ' + person.name + ',\n\n' +
        'You have been assigned work on the following project:\n\n' +
        'Project: ' + projectName + '\n' +
        'Client:  ' + accountName + '\n\n' +
        'Your assigned items:\n' + itemLines + '\n\n' +
        'Your commission: ' + currency + ' ' +
        Number(commAmt).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '\n\n' +
        'Please open the CRM system to view your project details.\n\n' +
        'This is an automated message.';

      MailApp.sendEmail({
        to:      email,
        subject: 'Project Assigned: ' + projectName + ' — ' + accountName,
        body:    body
      });
    });
  } catch(e) {
    Logger.log('Assignment email error: ' + e);
  }
}

// ============================================================
//  PROJECTS — GET MY PROJECTS (Team role)
//  Returns only projects where the user has Active assignments.
// ============================================================
function getMyProjects() {
  const actor  = getCurrentUser();
  const aSheet = getSheet('Assignments');
  const pSheet = getSheet('Projects');
  const cSheet = getSheet('Commissions');
  if (!aSheet || !pSheet) return [];

  const tz = Session.getScriptTimeZone();
  function fmtDate(val) {
    if (!val) return '';
    try { return Utilities.formatDate(new Date(val), tz, 'yyyy-MM-dd'); } catch(e) { return ''; }
  }

  // Projects this user has assignments on
  const aData = aSheet.getDataRange().getValues();
  const myProjectIds = new Set(
    aData.slice(1)
      .filter(row =>
        String(row[ASN.ASSIGNED_TO_EMAIL]) === actor.email &&
        String(row[ASN.STATUS])            === 'Active')
      .map(row => String(row[ASN.PROJECT_ID]))
  );

  if (!myProjectIds.size) return [];

  // My assignment items grouped by project
  const myItems = {};
  aData.slice(1)
    .filter(row =>
      myProjectIds.has(String(row[ASN.PROJECT_ID])) &&
      String(row[ASN.ASSIGNED_TO_EMAIL]) === actor.email &&
      String(row[ASN.STATUS])            === 'Active')
    .forEach(row => {
      const pid = String(row[ASN.PROJECT_ID]);
      if (!myItems[pid]) myItems[pid] = [];
      myItems[pid].push({
        itemName: String(row[ASN.ITEM_NAME]),
        quantity: Number(row[ASN.QUANTITY] || 0)
      });
    });

  // My commission amounts (not %) by project
  const myCommissions = {};
  if (cSheet) {
    const cData = cSheet.getDataRange().getValues();
    cData.slice(1)
      .filter(row =>
        myProjectIds.has(String(row[COM.PROJECT_ID])) &&
        String(row[COM.PERSON_EMAIL]) === actor.email)
      .forEach(row => {
        myCommissions[String(row[COM.PROJECT_ID])] = {
          amount:   Number(row[COM.COMMISSION_AMT]  || 0),
          currency: String(row[COM.CURRENCY]        || 'EGP'),
          status:   String(row[COM.PAYMENT_STATUS]  || 'Unpaid')
        };
      });
  }

  // Build project list
  const pData = pSheet.getDataRange().getValues();
  const projects = [];
  for (let i = 1; i < pData.length; i++) {
    const row = pData[i];
    const pid = String(row[P.PROJECT_ID]);
    if (!myProjectIds.has(pid)) continue;
    projects.push({
      projectId:        pid,
      quotationId:      String(row[P.QUOTATION_ID]   || ''),
      accountName:      String(row[P.ACCOUNT_NAME]   || ''),
      projectName:      String(row[P.PROJECT_NAME]   || ''),
      projectDesc:      String(row[P.PROJECT_DESC]   || ''),
      deliveryDeadline: fmtDate(row[P.DELIVERY_DEADLINE]),
      dueDate:          fmtDate(row[P.DUE_DATE]),
      status:           String(row[P.STATUS]         || ''),
      myItems:          myItems[pid] || [],
      myCommission:     myCommissions[pid] || null
    });
  }

  projects.sort((a, b) => a.projectName.localeCompare(b.projectName));
  return projects;
}
