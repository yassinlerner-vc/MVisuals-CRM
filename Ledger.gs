// ============================================================
//  PARTNER BALANCES — computed live from four sources, never
//  stored, so it can never drift out of sync:
//
//    + their share of every Revenue_Distribution row
//    − their share of every Expense_Split row
//    + the full EGP-equivalent amount of any Expenses row where
//      THEY personally were PaidBy (they fronted real cash —
//      the pool owes it back)
//    − the full EGP-equivalent amount of any Drawings row for
//      them
//
//  Every one of these four sources is now guaranteed to be
//  EGP-denominated at the source:
//    • Revenue_Distribution.AMOUNT is computed from the actual
//      EGP collected on the project (Projects.gs saveDistribution)
//    • Expense_Split.AMOUNT is computed from the expense's EGP
//      equivalent, not its native currency (Expenses.gs
//      logExpenseInstance)
//    • Expenses.EGP_EQUIVALENT / Drawings.EGP_EQUIVALENT are
//      captured at entry time, same pattern as Payments
//
//  So there is nothing left to bucket by currency — every
//  partner gets exactly one net EGP balance. 'currency: EGP' is
//  kept on the returned object purely so TeamForm.html's existing
//  render code (which reads b.currency) needs no changes.
// ============================================================
function getPartnerBalances() {
  const partners = getPartners(); // from Team.gs
  if (!partners.length) return [];

  // balances[personId] = { revenueShare, expenseShare, frontedAmount, drawings }
  const balances = {};
  partners.forEach(p => {
    balances[p.id] = {
      personId: p.id, personName: p.name,
      revenueShare: 0, expenseShare: 0, frontedAmount: 0, drawings: 0
    };
  });

  // Revenue distribution — always EGP (see Projects.gs saveDistribution)
  const distSheet = getSheet('Revenue_Distribution');
  if (distSheet) {
    distSheet.getDataRange().getValues().slice(1).forEach(row => {
      const b = balances[String(row[DIST.PERSON_ID])];
      if (b) b.revenueShare += Number(row[DIST.AMOUNT] || 0);
    });
  }

  // Expense splits — always EGP (see Expenses.gs logExpenseInstance)
  const splitSheet = getSheet('Expense_Split');
  if (splitSheet) {
    splitSheet.getDataRange().getValues().slice(1).forEach(row => {
      const b = balances[String(row[ESPL.PERSON_ID])];
      if (b) b.expenseShare += Number(row[ESPL.AMOUNT] || 0);
    });
  }

  // Fronted amounts — only when PaidBy is a specific partner, not the
  // Company Account. Uses the expense's EGP equivalent, not its native
  // amount, so a partner who fronted cash in a foreign currency is
  // owed back the actual EGP value.
  const expSheet = getSheet('Expenses');
  if (expSheet) {
    expSheet.getDataRange().getValues().slice(1).forEach(row => {
      const paidById = String(row[EXP.PAID_BY]);
      if (paidById === COMPANY_ACCOUNT_ID) return;
      const b = balances[paidById];
      if (b) b.frontedAmount += Number(row[EXP.EGP_EQUIVALENT] || 0);
    });
  }

  // Drawings — uses the drawing's EGP equivalent, not its native amount.
  const drwSheet = getSheet('Drawings');
  if (drwSheet) {
    drwSheet.getDataRange().getValues().slice(1).forEach(row => {
      const b = balances[String(row[DRW.PERSON_ID])];
      if (b) b.drawings += Number(row[DRW.EGP_EQUIVALENT] || 0);
    });
  }

  return Object.values(balances)
    .map(b => ({
      personId:      b.personId,
      personName:    b.personName,
      currency:      'EGP',
      revenueShare:  b.revenueShare,
      expenseShare:  b.expenseShare,
      frontedAmount: b.frontedAmount,
      drawings:      b.drawings,
      net: b.revenueShare - b.expenseShare + b.frontedAmount - b.drawings
    }))
    .sort((a, b) => a.personName.localeCompare(b.personName));
}
