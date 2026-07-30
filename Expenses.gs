// ============================================================
//  EXPENSES — GENERATE IDS
// ============================================================
function generateCatalogId() {
  const sheet = getSheet('Expense_Catalog');
  const data  = sheet.getDataRange().getValues();
  let maxNum  = 0;
  data.slice(1).forEach(row => {
    const match = String(row[EC.CATALOG_ID] || '').match(/^EC(\d+)$/);
    if (match) { const num = parseInt(match[1], 10); if (num > maxNum) maxNum = num; }
  });
  return 'EC' + String(maxNum + 1).padStart(6, '0');
}

function generateExpenseId() {
  const sheet = getSheet('Expenses');
  const data  = sheet.getDataRange().getValues();
  let maxNum  = 0;
  data.slice(1).forEach(row => {
    const match = String(row[EXP.EXPENSE_ID] || '').match(/^EXP(\d+)$/);
    if (match) { const num = parseInt(match[1], 10); if (num > maxNum) maxNum = num; }
  });
  return 'EXP' + String(maxNum + 1).padStart(6, '0');
}

// ============================================================
//  PERIOD KEY HELPER — "2026-07". Recurring instances are
//  matched against this to detect what's already logged.
// ============================================================
function currentPeriodKey() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
}

// ============================================================
//  SPLIT VALIDATION — percentages must add to 100 (small
//  rounding tolerance), same tolerance style as saveDistribution.
// ============================================================
function validateSplits(splits) {
  if (!splits || !splits.length) return 'Add at least one person to the split.';
  const sum = splits.reduce((s, r) => s + (Number(r.percent) || 0), 0);
  if (Math.abs(sum - 100) > 0.5) {
    return 'Split percentages must add up to 100% (currently ' + sum.toFixed(1) + '%).';
  }
  return null;
}

// ============================================================
//  PAID-BY OPTIONS — Company Account + Partners only
// ============================================================
function getPaidByOptions() {
  const options = [{ id: COMPANY_ACCOUNT_ID, name: COMPANY_ACCOUNT_NAME }];
  getPartners().forEach(p => options.push({ id: p.id, name: p.name }));
  return options;
}

// ============================================================
//  CATEGORIES — distinct categories already used in the
//  catalog, for a select-or-type datalist (no separate sheet
//  to maintain, same pattern as Accounts' free-type fields).
// ============================================================
function getExpenseCategories() {
  const sheet = getSheet('Expense_Catalog');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const set = new Set();
  data.slice(1).forEach(row => { if (row[EC.CATEGORY]) set.add(String(row[EC.CATEGORY])); });
  return Array.from(set).sort();
}

// ============================================================
//  EXPENSE CATALOG — LIST
//  For each recurring, Active entry, computes whether this
//  period already has a confirmed instance ("Logged") or not
//  ("Needs Review") — computed live, never stored.
// ============================================================
function getExpenseCatalog() {
  const cSheet = getSheet('Expense_Catalog');
  const eSheet = getSheet('Expenses');
  if (!cSheet) return [];
  const cData = cSheet.getDataRange().getValues();
  const eData = eSheet ? eSheet.getDataRange().getValues() : [];
  const tz     = Session.getScriptTimeZone();
  const period = currentPeriodKey();

  const loggedPeriods    = {}; // catalogId -> Set of periodKeys with a logged instance
  const lastLoggedDate   = {}; // catalogId -> latest instance date (yyyy-MM-dd)
  eData.slice(1).forEach(row => {
    const cid = String(row[EXP.CATALOG_ID]);
    const pk  = String(row[EXP.PERIOD_KEY] || '');
    if (pk) {
      if (!loggedPeriods[cid]) loggedPeriods[cid] = new Set();
      loggedPeriods[cid].add(pk);
    }
    if (row[EXP.DATE]) {
      const ds = Utilities.formatDate(new Date(row[EXP.DATE]), tz, 'yyyy-MM-dd');
      if (!lastLoggedDate[cid] || ds > lastLoggedDate[cid]) lastLoggedDate[cid] = ds;
    }
  });

  return cData.slice(1)
    .filter(row => row[EC.CATALOG_ID])
    .map(row => {
      const catalogId   = String(row[EC.CATALOG_ID]);
      const isRecurring = !!row[EC.IS_RECURRING];
      const status      = String(row[EC.STATUS] || 'Active');
      let reviewStatus  = null; // only meaningful for Active recurring entries
      if (isRecurring && status === 'Active') {
        const logged = !!(loggedPeriods[catalogId] && loggedPeriods[catalogId].has(period));
        reviewStatus = logged ? 'Logged' : 'Needs Review';
      }
      return {
        id:             catalogId,
        name:           row[EC.NAME],
        category:       row[EC.CATEGORY],
        vendor:         row[EC.VENDOR],
        isRecurring:    isRecurring,
        frequency:      row[EC.FREQUENCY],
        lastAmount:     Number(row[EC.LAST_AMOUNT] || 0),
        lastCurrency:   row[EC.LAST_CURRENCY],
        lastPaidById:   row[EC.LAST_PAID_BY],
        lastPaidByName: row[EC.LAST_PAID_BY_NAME],
        lastSplit:      row[EC.LAST_SPLIT_JSON] ? JSON.parse(row[EC.LAST_SPLIT_JSON]) : [],
        status:         status,
        notes:          row[EC.NOTES],
        reviewStatus:   reviewStatus,
        currentPeriod:  period,
        lastLoggedDate: lastLoggedDate[catalogId] || ''
      };
    })
    .sort((a, b) => a.name.localeCompare(b.name));
}

function getExpenseCatalogEntry(catalogId) {
  return getExpenseCatalog().find(c => c.id === catalogId) || null;
}

// ============================================================
//  EXPENSE HISTORY — every logged instance for a catalog entry,
//  newest first. This IS the price-history view.
// ============================================================
function getExpenseHistory(catalogId) {
  const eSheet = getSheet('Expenses');
  const sSheet = getSheet('Expense_Split');
  if (!eSheet) return [];
  const eData = eSheet.getDataRange().getValues();
  const sData = sSheet ? sSheet.getDataRange().getValues() : [];
  const tz    = Session.getScriptTimeZone();

  return eData.slice(1)
    .filter(row => String(row[EXP.CATALOG_ID]) === String(catalogId))
    .map(row => {
      const expenseId = String(row[EXP.EXPENSE_ID]);
      const splits = sData.slice(1)
        .filter(s => String(s[ESPL.EXPENSE_ID]) === expenseId)
        .map(s => ({
          personId:   s[ESPL.PERSON_ID],
          personName: s[ESPL.PERSON_NAME],
          percent:    Number(s[ESPL.PERCENT] || 0),
          amount:     Number(s[ESPL.AMOUNT]  || 0)
        }));
      return {
        expenseId:   expenseId,
        amount:      Number(row[EXP.AMOUNT] || 0),
        currency:    row[EXP.CURRENCY],
        date:        row[EXP.DATE] ? Utilities.formatDate(new Date(row[EXP.DATE]), tz, 'yyyy-MM-dd') : '',
        paidById:    row[EXP.PAID_BY],
        paidByName:  row[EXP.PAID_BY_NAME],
        periodKey:   row[EXP.PERIOD_KEY] || '',
        notes:       row[EXP.NOTES],
        splits:      splits
      };
    })
    .sort((a, b) => b.date.localeCompare(a.date));
}

// ============================================================
//  NEW EXPENSE — creates a catalog entry (empty Last* fields),
//  then logs its first instance via logExpenseInstance, which
//  fills those Last* fields in.
// ============================================================
function createExpenseCatalogEntry(formData) {
  const cSheet = getSheet('Expense_Catalog');
  if (!cSheet) return { success: false, error: 'Expense_Catalog sheet missing.' };
  if (!formData.name || !String(formData.name).trim()) {
    return { success: false, error: 'Please enter an expense name.' };
  }

  const catalogId   = generateCatalogId();
  const now         = new Date();
  const user        = Session.getActiveUser().getEmail();
  const isRecurring = !!formData.isRecurring;

  cSheet.appendRow([
    catalogId, formData.name, formData.category || '', formData.vendor || '',
    isRecurring, isRecurring ? (formData.frequency || 'Monthly') : '',
    '', '', '', '', '', 'Active', formData.notes || '', user, now
  ]);

  writeLog('Expenses_Log', 'Expense_Catalog', catalogId, catalogId + ' — ' + formData.name,
    '', '', 'Status', '', 'Created', 'Expense catalog entry created');

  const instanceResult = logExpenseInstance({
    catalogId:  catalogId,
    amount:     formData.amount,
    currency:   formData.currency,
    date:       formData.date,
    paidById:   formData.paidById,
    paidByName: formData.paidByName,
    splits:     formData.splits,
    notes:      formData.notes,
    periodKey:  isRecurring ? currentPeriodKey() : ''
  });

  if (!instanceResult.success) {
    // Roll back the catalog row so we don't leave an orphaned entry with no instance
    const data = cSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][EC.CATALOG_ID]) === catalogId) { cSheet.deleteRow(i + 1); break; }
    }
    return instanceResult;
  }

  return { success: true, catalogId, expenseId: instanceResult.expenseId };
}

// ============================================================
//  LOG EXPENSE INSTANCE — the shared engine behind "Log Expense",
//  "Review" (recurring), and the first-instance call from
//  createExpenseCatalogEntry. Writes the Expenses row + splits,
//  then rolls the catalog's Last* fields forward.
// ============================================================
function logExpenseInstance(data) {
  const cSheet = getSheet('Expense_Catalog');
  const eSheet = getSheet('Expenses');
  const sSheet = getSheet('Expense_Split');
  if (!cSheet || !eSheet || !sSheet) return { success: false, error: 'Required sheets missing.' };

  const amount = Number(data.amount) || 0;
  if (amount <= 0) return { success: false, error: 'Enter a valid amount.' };
  if (!data.date) return { success: false, error: 'Select a date.' };
  if (!data.paidById) return { success: false, error: 'Select who paid.' };

  const splitError = validateSplits(data.splits);
  if (splitError) return { success: false, error: splitError };

  const cData = cSheet.getDataRange().getValues();
  let catalogRowIndex = -1, catalogRow = null;
  for (let i = 1; i < cData.length; i++) {
    if (String(cData[i][EC.CATALOG_ID]) === String(data.catalogId)) {
      catalogRowIndex = i + 1; catalogRow = cData[i]; break;
    }
  }
  if (!catalogRow) return { success: false, error: 'Expense catalog entry not found.' };

  const catalogName = String(catalogRow[EC.NAME]);
  const category     = String(catalogRow[EC.CATEGORY] || '');
  const isRecurring   = !!catalogRow[EC.IS_RECURRING];
  const periodKey      = isRecurring ? (data.periodKey || currentPeriodKey()) : '';

  // Guard: don't allow two confirmed instances for the same recurring period
  if (isRecurring && periodKey) {
    const eDataCheck = eSheet.getDataRange().getValues();
    const dup = eDataCheck.slice(1).some(row =>
      String(row[EXP.CATALOG_ID]) === String(data.catalogId) &&
      String(row[EXP.PERIOD_KEY]) === periodKey);
    if (dup) return { success: false, error: 'This period has already been logged for this recurring expense.' };
  }

  const expenseId = generateExpenseId();
  const now  = new Date();
  const user = Session.getActiveUser().getEmail();

  eSheet.appendRow([
    expenseId, data.catalogId, catalogName, category,
    amount, data.currency, data.date,
    data.paidById, data.paidByName, periodKey,
    data.notes || '', user, now
  ]);

  data.splits.forEach(s => {
    const pct = Number(s.percent) || 0;
    sSheet.appendRow([
      Utilities.getUuid(), expenseId, data.catalogId,
      s.personId, s.personName, pct, amount * pct / 100,
      data.currency, user, now
    ]);
  });

  // Roll the catalog's "last known" fields forward
  cSheet.getRange(catalogRowIndex, EC.LAST_AMOUNT       + 1).setValue(amount);
  cSheet.getRange(catalogRowIndex, EC.LAST_CURRENCY     + 1).setValue(data.currency);
  cSheet.getRange(catalogRowIndex, EC.LAST_PAID_BY      + 1).setValue(data.paidById);
  cSheet.getRange(catalogRowIndex, EC.LAST_PAID_BY_NAME + 1).setValue(data.paidByName);
  cSheet.getRange(catalogRowIndex, EC.LAST_SPLIT_JSON   + 1).setValue(JSON.stringify(data.splits));

  writeLog('Expenses_Log', 'Expenses', expenseId, data.catalogId + ' — ' + catalogName,
    '', '', 'Amount', '', amount,
    (periodKey ? 'Recurring instance logged for ' + periodKey : 'Expense logged') +
    ', paid by ' + data.paidByName);

  return { success: true, expenseId };
}

// ============================================================
//  DELETE EXPENSE INSTANCE — removes the transaction + its
//  splits. Does not touch the catalog entry itself.
// ============================================================
function deleteExpenseInstance(expenseId) {
  const eSheet = getSheet('Expenses');
  const sSheet = getSheet('Expense_Split');
  if (!eSheet) return { success: false, error: 'Expenses sheet missing.' };

  const eData = eSheet.getDataRange().getValues();
  let found = null;
  for (let i = 1; i < eData.length; i++) {
    if (String(eData[i][EXP.EXPENSE_ID]) === String(expenseId)) {
      found = { row: i + 1, data: eData[i] }; break;
    }
  }
  if (!found) return { success: false, error: 'Expense not found.' };

  eSheet.deleteRow(found.row);

  if (sSheet) {
    const sData = sSheet.getDataRange().getValues();
    for (let i = sData.length - 1; i >= 1; i--) {
      if (String(sData[i][ESPL.EXPENSE_ID]) === String(expenseId)) sSheet.deleteRow(i + 1);
    }
  }

  writeLog('Expenses_Log', 'Expenses', expenseId,
    expenseId + ' — ' + found.data[EXP.CATALOG_NAME], '', '',
    'Status', 'Logged', 'Deleted', 'Expense instance deleted');

  return { success: true };
}

// ============================================================
//  ARCHIVE / PAUSE CATALOG ENTRY
// ============================================================
function setExpenseCatalogStatus(catalogId, status) {
  const cSheet = getSheet('Expense_Catalog');
  if (!cSheet) return { success: false, error: 'Expense_Catalog sheet missing.' };
  const data = cSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][EC.CATALOG_ID]) === String(catalogId)) {
      const oldStatus = data[i][EC.STATUS];
      cSheet.getRange(i + 1, EC.STATUS + 1).setValue(status);
      writeLog('Expenses_Log', 'Expense_Catalog', catalogId,
        catalogId + ' — ' + data[i][EC.NAME], '', '',
        'Status', oldStatus, status, '');
      return { success: true };
    }
  }
  return { success: false, error: 'Expense catalog entry not found.' };
}
