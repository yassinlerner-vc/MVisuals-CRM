// ============================================================
//  PROJECTS — GENERATE PAYMENT ID
// ============================================================
function generatePaymentId() {
  const sheet = getSheet('Payments');
  const data  = sheet.getDataRange().getValues();
  let maxNum  = 0;
  data.slice(1).forEach(row => {
    const match = String(row[PAY.PAYMENT_ID] || '').match(/^PAY(\d+)$/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  });
  return 'PAY' + String(maxNum + 1).padStart(6, '0');
}

// ============================================================
//  PROJECTS — LIST
//  Payment status is computed live from the Payments sheet,
//  not stored, so it can never drift out of sync.
// ============================================================
function getProjects() {
  try {
    const pSheet    = getSheet('Projects');
    const paySheet  = getSheet('Payments');
    const distSheet = getSheet('Revenue_Distribution');
    if (!pSheet) return [];
    const data = pSheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const tz = Session.getScriptTimeZone();

    function fmtDate(val) {
      if (!val) return '';
      try { return Utilities.formatDate(new Date(val), tz, 'yyyy-MM-dd'); } catch(e) { return ''; }
    }

    // Build paid-amount map: projectId -> sum of payments
    const paidMap = {};
    if (paySheet) {
      paySheet.getDataRange().getValues().slice(1).forEach(row => {
        const pid = String(row[PAY.PROJECT_ID]);
        if (!pid) return;
        paidMap[pid] = (paidMap[pid] || 0) + (Number(row[PAY.AMOUNT]) || 0);
      });
    }

    // Which projects already have a distribution recorded
    const distributedIds = new Set();
    if (distSheet) {
      distSheet.getDataRange().getValues().slice(1).forEach(row => {
        const pid = String(row[DIST.PROJECT_ID]);
        if (pid) distributedIds.add(pid);
      });
    }

    const projects = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[P.PROJECT_ID]) continue;
      const pid         = String(row[P.PROJECT_ID]);
      const totalAmount = Number(row[P.TOTAL_AMOUNT] || 0);
      const amountPaid  = paidMap[pid] || 0;
      const remaining   = Math.max(0, totalAmount - amountPaid);

      let paymentStatus = 'Unpaid';
      if (totalAmount > 0 && amountPaid >= totalAmount) paymentStatus = 'Fully Paid';
      else if (amountPaid > 0) paymentStatus = 'Partially Paid';

      try {
        projects.push({
          projectId:        pid,
          quotationId:      String(row[P.QUOTATION_ID] || ''),
          accountId:        String(row[P.ACCOUNT_ID]   || ''),
          accountName:      String(row[P.ACCOUNT_NAME] || ''),
          projectName:      String(row[P.PROJECT_NAME] || ''),
          projectDesc:      String(row[P.PROJECT_DESC] || ''),
          deliveryDeadline: fmtDate(row[P.DELIVERY_DEADLINE]),
          dueDate:          fmtDate(row[P.DUE_DATE]),
          totalAmount:      totalAmount,
          currency:         String(row[P.CURRENCY] || 'EGP'),
          amountPaid:       amountPaid,
          remainingToPay:   remaining,
          paymentStatus:    paymentStatus,
          isDistributed:    distributedIds.has(pid),
          createdAt:        fmtDate(row[P.CREATED_AT])
        });
      } catch(e) {
        Logger.log('getProjects: skipping row ' + i + ': ' + e);
      }
    }

    const order = { 'Unpaid': 0, 'Partially Paid': 1, 'Fully Paid': 2 };
    projects.sort((a, b) => {
      const ao = order[a.paymentStatus] !== undefined ? order[a.paymentStatus] : 1;
      const bo = order[b.paymentStatus] !== undefined ? order[b.paymentStatus] : 1;
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
//  PROJECTS — GET DETAIL
//  Single call powering the whole detail view: project meta,
//  items (reference only), payments, distribution, team list.
// ============================================================
function getProjectDetail(projectId) {
  const pSheet    = getSheet('Projects');
  const piSheet   = getSheet('Project_Items');
  const paySheet  = getSheet('Payments');
  const distSheet = getSheet('Revenue_Distribution');
  const tSheet    = getSheet('Team');
  if (!pSheet) return null;

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
        quotationId:      String(row[P.QUOTATION_ID] || ''),
        accountId:        String(row[P.ACCOUNT_ID]   || ''),
        accountName:      String(row[P.ACCOUNT_NAME] || ''),
        projectName:      String(row[P.PROJECT_NAME] || ''),
        projectDesc:      String(row[P.PROJECT_DESC] || ''),
        deliveryDeadline: fmtDate(row[P.DELIVERY_DEADLINE]),
        dueDate:          fmtDate(row[P.DUE_DATE]),
        totalAmount:      Number(row[P.TOTAL_AMOUNT] || 0),
        currency:         String(row[P.CURRENCY] || 'EGP')
      };
      break;
    }
  }
  if (!project) return null;

  // ── Items (reference only — no assignment) ─────────────────
  const items = piSheet ? piSheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[PI.PROJECT_ID]) === String(projectId))
    .map(row => ({
      itemId:       String(row[PI.ITEM_ID]),
      itemName:     String(row[PI.ITEM_NAME]),
      displayValue: String(row[PI.DISPLAY_VALUE] || ''),
      quantity:     Number(row[PI.QUANTITY] || 0),
      description:  String(row[PI.DESCRIPTION] || ''),
      notes:        String(row[PI.NOTES] || '')
    })) : [];

  // ── Payments ─────────────────────────────────────────────
  const payments = paySheet ? paySheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[PAY.PROJECT_ID]) === String(projectId))
    .map(row => ({
      paymentId:  String(row[PAY.PAYMENT_ID]),
      amount:     Number(row[PAY.AMOUNT] || 0),
      date:       row[PAY.DATE] ? fmtDate(row[PAY.DATE]) : '',
      method:     String(row[PAY.METHOD] || ''),
      notes:      String(row[PAY.NOTES] || ''),
      recordedBy: String(row[PAY.RECORDED_BY] || ''),
      recordedAt: row[PAY.RECORDED_AT] ? fmtDate(row[PAY.RECORDED_AT]) : ''
    }))
    .sort((a, b) => (a.date || '').localeCompare(b.date || '')) : [];

  const amountPaid    = payments.reduce((sum, p) => sum + p.amount, 0);
  const remainingToPay = Math.max(0, project.totalAmount - amountPaid);
  const isFullyPaid    = project.totalAmount > 0 && amountPaid >= project.totalAmount;

  // ── Distribution ─────────────────────────────────────────
  const distribution = distSheet ? distSheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[DIST.PROJECT_ID]) === String(projectId))
    .map(row => ({
      distributionId: String(row[DIST.DISTRIBUTION_ID]),
      personId:       String(row[DIST.PERSON_ID]   || ''),
      personName:     String(row[DIST.PERSON_NAME] || ''),
      percent:        Number(row[DIST.PERCENT] || 0),
      amount:         Number(row[DIST.AMOUNT]  || 0),
      notes:          String(row[DIST.NOTES]   || '')
    })) : [];

  // ── Team members (for distribution dropdown) ───────────────
  const teamMembers = tSheet ? tSheet.getDataRange().getValues().slice(1)
    .filter(row => row[TM.TEAM_ID])
    .map(row => ({
      id:   String(row[TM.TEAM_ID]),
      name: String(row[TM.NAME]),
      role: String(row[TM.ROLE])
    })) : [];

  return { project, items, payments, distribution, teamMembers, amountPaid, remainingToPay, isFullyPaid };
}

// ============================================================
//  PAYMENTS — RECORD
//  Caps total payments at the project's quotation total.
// ============================================================
function recordPayment(projectId, data) {
  const pSheet   = getSheet('Projects');
  const paySheet = getSheet('Payments');
  if (!pSheet || !paySheet) return { success: false, error: 'Required sheets missing.' };

  const pData = pSheet.getDataRange().getValues();
  let project = null;
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) { project = pData[i]; break; }
  }
  if (!project) return { success: false, error: 'Project not found.' };

  const totalAmount = Number(project[P.TOTAL_AMOUNT] || 0);
  const quotationId = String(project[P.QUOTATION_ID] || '');
  const accountName = String(project[P.ACCOUNT_NAME] || '');
  const projectName = String(project[P.PROJECT_NAME] || '');

  const existingPaid = paySheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[PAY.PROJECT_ID]) === String(projectId))
    .reduce((sum, row) => sum + (Number(row[PAY.AMOUNT]) || 0), 0);

  const amount = Number(data.amount) || 0;
  if (amount <= 0) return { success: false, error: 'Enter a valid payment amount.' };
  if (existingPaid + amount > totalAmount + 0.01) {
    return { success: false, error: 'Payment exceeds the remaining balance of ' +
      (totalAmount - existingPaid).toFixed(2) + '.' };
  }

  const paymentId = generatePaymentId();
  const user = Session.getActiveUser().getEmail();
  const now  = new Date();
  const dateVal = data.date || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  paySheet.appendRow([
    paymentId, projectId, quotationId, accountName, projectName,
    amount, dateVal, data.method || '', data.notes || '', user, now
  ]);

  writeLog('Payments_Log', 'Payments', paymentId,
    paymentId + ' — ' + projectName, '', accountName,
    'Payment', '', amount, 'Payment of ' + amount + ' recorded by ' + user);

  return { success: true, paymentId };
}

// ============================================================
//  PAYMENTS — DELETE
//  Blocked once revenue has been distributed for the project,
//  to keep the books from going inconsistent.
// ============================================================
function deletePayment(paymentId) {
  const paySheet  = getSheet('Payments');
  const distSheet = getSheet('Revenue_Distribution');
  if (!paySheet) return { success: false, error: 'Payments sheet missing.' };

  const data = paySheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][PAY.PAYMENT_ID]) === String(paymentId)) {
      const projectId   = String(data[i][PAY.PROJECT_ID]);
      const projectName = String(data[i][PAY.PROJECT_NAME]);

      if (distSheet) {
        const hasDist = distSheet.getDataRange().getValues().slice(1)
          .some(r => String(r[DIST.PROJECT_ID]) === projectId);
        if (hasDist) {
          return { success: false, error: 'Cannot delete a payment after revenue has been distributed for this project.' };
        }
      }

      paySheet.deleteRow(i + 1);
      writeLog('Payments_Log', 'Payments', paymentId,
        paymentId + ' — ' + projectName, '', '',
        'Status', 'Recorded', 'Deleted', 'Payment deleted');
      return { success: true };
    }
  }
  return { success: false, error: 'Payment not found.' };
}

// ============================================================
//  REVENUE DISTRIBUTION — SAVE
//  Only allowed once the project is fully paid. Replaces any
//  existing distribution rows for this project. Rows carry a
//  PERCENT per person (not an amount) — percentages must sum
//  to 100 (small rounding tolerance). Amount is computed here,
//  server-side, from percent × project total.
//  rows: [{ personId, personName, percent }]
// ============================================================
function saveDistribution(projectId, rows) {
  const pSheet    = getSheet('Projects');
  const paySheet  = getSheet('Payments');
  const distSheet = getSheet('Revenue_Distribution');
  if (!pSheet || !distSheet) return { success: false, error: 'Required sheets missing.' };

  const pData = pSheet.getDataRange().getValues();
  let project = null;
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) { project = pData[i]; break; }
  }
  if (!project) return { success: false, error: 'Project not found.' };

  const totalAmount = Number(project[P.TOTAL_AMOUNT] || 0);
  const quotationId = String(project[P.QUOTATION_ID] || '');
  const accountName = String(project[P.ACCOUNT_NAME] || '');
  const projectName = String(project[P.PROJECT_NAME] || '');
  const currency    = String(project[P.CURRENCY] || 'EGP');

  const amountPaid = paySheet ? paySheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[PAY.PROJECT_ID]) === String(projectId))
    .reduce((sum, row) => sum + (Number(row[PAY.AMOUNT]) || 0), 0) : 0;

  if (totalAmount <= 0 || amountPaid < totalAmount) {
    return { success: false, error: 'Project must be fully paid before distributing revenue.' };
  }

  if (!rows || !rows.length) {
    return { success: false, error: 'Add at least one person to distribute to.' };
  }

  const pctSum = rows.reduce((s, r) => s + (Number(r.percent) || 0), 0);
  if (Math.abs(pctSum - 100) > 0.5) {
    return { success: false, error: 'Distributed percentages (' + pctSum.toFixed(1) +
      '%) must add up to 100%.' };
  }

  // Clear any existing distribution rows for this project
  const distData = distSheet.getDataRange().getValues();
  for (let i = distData.length - 1; i >= 1; i--) {
    if (String(distData[i][DIST.PROJECT_ID]) === String(projectId)) distSheet.deleteRow(i + 1);
  }

  const user = Session.getActiveUser().getEmail();
  const now  = new Date();

  rows.forEach(r => {
    const percent = Number(r.percent) || 0;
    const amount  = totalAmount * percent / 100;
    distSheet.appendRow([
      Utilities.getUuid(), projectId, quotationId, accountName, projectName,
      r.personId || '', r.personName || '', percent, amount,
      currency, r.notes || '', user, now
    ]);
  });

  writeLog('Revenue_Distribution_Log', 'Revenue_Distribution', projectId,
    projectId + ' — ' + projectName, '', accountName,
    'Distribution', '',
    JSON.stringify(rows.map(r => ({ name: r.personName, percent: r.percent }))),
    'Revenue distributed by ' + user);

  return { success: true };
}
