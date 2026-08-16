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
//  logStatus is now computed for EVERY entry (not just Active
//  recurring ones), and is the single source of truth the
//  frontend renders — no more duplicated logic client-side.
//
//    Recurring : 'Logged' if this period already has an
//                instance, else 'Needs Review'.
//    One-time  : 'Logged' if any instance has ever been logged,
//                else 'Not Logged'.
// ============================================================
function getExpenseCatalog() {
  const cSheet = getSheet('Expense_Catalog');
  const eSheet = getSheet('Expenses');
  if (!cSheet) return [];
  const cData = cSheet.getDataRange().getValues();
  const eData = eSheet ? eSheet.getDataRange().getValues() : [];
  const tz     = Session.getScriptTimeZone();
  const period = currentPeriodKey();

  const loggedPeriods  = {}; // catalogId -> Set of periodKeys with a logged instance
  const lastLoggedDate = {}; // catalogId -> latest instance date (yyyy-MM-dd)
  const everLogged     = {}; // catalogId -> true if any instance exists at all
  eData.slice(1).forEach(row => {
    const cid = String(row[EXP.CATALOG_ID]);
    if (!cid) return;
    everLogged[cid] = true;
    const pk = String(row[EXP.PERIOD_KEY] || '');
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

      let reviewStatus;
      if (isRecurring) {
        const logged = !!(loggedPeriods[catalogId] && loggedPeriods[catalogId].has(period));
        reviewStatus = logged ? 'Logged' : 'Needs Review';
      } else {
        reviewStatus = everLogged[catalogId] ? 'Logged' : 'Not Logged';
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
        status:         String(row[EC.STATUS] || 'Active'),
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
//  newest first. This IS the price-history view. Includes the
//  EGP equivalent/exchange rate for foreign-currency instances.
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
      const currency = row[EXP.CURRENCY];
      return {
        expenseId:     expenseId,
        amount:        Number(row[EXP.AMOUNT] || 0),
        currency:      currency,
        egpEquivalent: Number(row[EXP.EGP_EQUIVALENT] || 0),
        exchangeRate:  Number(row[EXP.EXCHANGE_RATE]  || 0),
        isForeign:     String(currency) !== 'EGP',
        date:          row[EXP.DATE] ? Utilities.formatDate(new Date(row[EXP.DATE]), tz, 'yyyy-MM-dd') : '',
        paidById:      row[EXP.PAID_BY],
        paidByName:    row[EXP.PAID_BY_NAME],
        periodKey:     row[EXP.PERIOD_KEY] || '',
        notes:         row[EXP.NOTES],
        splits:        splits
      };
    })
    .sort((a, b) => b.date.localeCompare(a.date));
}

// ============================================================
//  NEW EXPENSE — CREATE CATALOG ENTRY ONLY
//  No amount/date/paidBy required anymore. This just registers
//  the expense type (name, category, vendor, recurring/frequency,
//  notes). Logging the first (or any) paid instance is always
//  done afterwards through logExpenseInstance — one single
//  logging code path for both "first log" and "later log",
//  so there's nothing to roll back and nothing to fall out of
//  sync.
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
    '', '', 'Status', '', 'Created', 'Expense catalog entry created (not yet logged)');

  return { success: true, catalogId };
}

// ============================================================
//  LOG EXPENSE INSTANCE — the single engine behind "Log Expense",
//  "Review" (recurring), and logging the first instance right
//  after creating a new catalog entry. Writes the Expenses row
//  + splits, then rolls the catalog's Last* fields forward.
//
//  EGP equivalent: required whenever currency !== 'EGP' (mirrors
//  the same rule used on Payments). For EGP-currency instances,
//  egpEquivalent defaults to amount and exchangeRate to 1.
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

  const isForeign = String(data.currency) !== 'EGP';
  let egpEquivalent = isForeign ? Number(data.egpEquivalent) || 0 : amount;
  if (isForeign && egpEquivalent <= 0) {
    return { success: false, error: 'Enter the EGP equivalent for this foreign-currency expense.' };
  }
  const exchangeRate = isForeign ? (egpEquivalent / amount) : 1;

  const cData = cSheet.getDataRange().getValues();
  let catalogRowIndex = -1, catalogRow = null;
  for (let i = 1; i < cData.length; i++) {
    if (String(cData[i][EC.CATALOG_ID]) === String(data.catalogId)) {
      catalogRowIndex = i + 1; catalogRow = cData[i]; break;
    }
  }
  if (!catalogRow) return { success: false, error: 'Expense catalog entry not found.' };

  const catalogName  = String(catalogRow[EC.NAME]);
  const category      = String(catalogRow[EC.CATEGORY] || '');
  const isRecurring    = !!catalogRow[EC.IS_RECURRING];
  const periodKey       = isRecurring ? (data.periodKey || currentPeriodKey()) : '';

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
    data.notes || '', user, now,
    egpEquivalent, exchangeRate
  ]);

  data.splits.forEach(s => {
    const pct = Number(s.percent) || 0;
    sSheet.appendRow([
      Utilities.getUuid(), expenseId, data.catalogId,
      s.personId, s.personName, pct, amount * pct / 100,
      data.currency, user, now
    ]);
  });

  // Roll the catalog's "last known" fields forward — used only as a
  // prefill suggestion for next time, never an authority. EGP
  // equivalent is intentionally NOT carried forward here since
  // exchange rates fluctuate month to month; each log should get a
  // fresh EGP entry rather than reusing a stale rate.
  cSheet.getRange(catalogRowIndex, EC.LAST_AMOUNT       + 1).setValue(amount);
  cSheet.getRange(catalogRowIndex, EC.LAST_CURRENCY     + 1).setValue(data.currency);
  cSheet.getRange(catalogRowIndex, EC.LAST_PAID_BY      + 1).setValue(data.paidById);
  cSheet.getRange(catalogRowIndex, EC.LAST_PAID_BY_NAME + 1).setValue(data.paidByName);
  cSheet.getRange(catalogRowIndex, EC.LAST_SPLIT_JSON   + 1).setValue(JSON.stringify(data.splits));

  writeLog('Expenses_Log', 'Expenses', expenseId, data.catalogId + ' — ' + catalogName,
    '', '', 'Amount', '', amount,
    (periodKey ? 'Recurring instance logged for ' + periodKey : 'Expense logged') +
    ', paid by ' + data.paidByName +
    (isForeign ? ' (EGP equiv. ' + egpEquivalent.toFixed(2) + ')' : ''));

  return { success: true, expenseId };
}

// ============================================================
//  DELETE EXPENSE INSTANCE — removes the transaction + its
//  splits. Does not touch the catalog entry itself. Note: this
//  does not roll the catalog's Last* preview fields backward —
//  those are cosmetic prefill hints only; the real history lives
//  in the Expenses sheet and is always accurate.
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

// ============================================================
//  DIAGNOSTIC — run manually from the Apps Script editor
//  (select this function → Run → View → Logs) if expense
//  history still doesn't show up after this update. Prints the
//  live Expenses sheet's actual header row next to what the EXP
//  column map expects, plus a sample of CatalogId values, so any
//  drift between the code's column indices and the real sheet
//  layout is immediately visible.
// ============================================================
function debugExpensesSheet() {
  const sheet = getSheet('Expenses');
  if (!sheet) { Logger.log('Expenses sheet not found.'); return; }
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];

  Logger.log('--- Live sheet headers (by column index) ---');
  headers.forEach((h, i) => Logger.log(i + ': ' + h));

  Logger.log('--- EXP map expects ---');
  Object.keys(EXP).forEach(k => Logger.log(EXP[k] + ': ' + k));

  Logger.log('--- Sample data rows (first 5) ---');
  data.slice(1, 6).forEach((row, i) => {
    Logger.log('Row ' + (i + 2) + ' -> ExpenseId=' + row[EXP.EXPENSE_ID] +
      ' | CatalogId=' + row[EXP.CATALOG_ID] +
      ' | Amount=' + row[EXP.AMOUNT] +
      ' | Date=' + row[EXP.DATE]);
  });

  Logger.log('--- Catalog IDs currently in Expense_Catalog ---');
  const cSheet = getSheet('Expense_Catalog');
  if (cSheet) {
    cSheet.getDataRange().getValues().slice(1).forEach(row => {
      if (row[EC.CATALOG_ID]) Logger.log(row[EC.CATALOG_ID] + ' — ' + row[EC.NAME]);
    });
  }
}
