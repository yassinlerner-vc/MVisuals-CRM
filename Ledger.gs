// ============================================================
//  PARTNER BALANCES — computed live from four sources, never
//  stored, so it can never drift out of sync:
//
//    + their share of every Revenue_Distribution row
//    − their share of every Expense_Split row
//    + the full amount of any Expenses row where THEY personally
//      were PaidBy (they fronted real cash — the pool owes it back)
//    − the full amount of any Drawings row for them
//
//  Grouped by currency — EGP and USD (etc.) are never summed
//  together, only shown side by side.
// ============================================================
function getPartnerBalances() {
  const partners = getPartners(); // from Team.gs
  if (!partners.length) return [];

  // balances[personId][currency] = { revenueShare, expenseShare, frontedAmount, drawings }
  const balances = {};
  partners.forEach(p => { balances[p.id] = {}; });

  function bucket(personId, currency) {
    if (!balances[personId]) return null; // not a current partner — ignore
    const cur = currency || '—';
    if (!balances[personId][cur]) {
      balances[personId][cur] = { revenueShare: 0, expenseShare: 0, frontedAmount: 0, drawings: 0 };
    }
    return balances[personId][cur];
  }

  // Revenue distribution
  const distSheet = getSheet('Revenue_Distribution');
  if (distSheet) {
    distSheet.getDataRange().getValues().slice(1).forEach(row => {
      const b = bucket(String(row[DIST.PERSON_ID]), row[DIST.CURRENCY]);
      if (b) b.revenueShare += Number(row[DIST.AMOUNT] || 0);
    });
  }

  // Expense splits (fair share of cost)
  const splitSheet = getSheet('Expense_Split');
  if (splitSheet) {
    splitSheet.getDataRange().getValues().slice(1).forEach(row => {
      const b = bucket(String(row[ESPL.PERSON_ID]), row[ESPL.CURRENCY]);
      if (b) b.expenseShare += Number(row[ESPL.AMOUNT] || 0);
    });
  }

  // Fronted amounts — only when PaidBy is a specific partner, not the Company Account
  const expSheet = getSheet('Expenses');
  if (expSheet) {
    expSheet.getDataRange().getValues().slice(1).forEach(row => {
      const paidById = String(row[EXP.PAID_BY]);
      if (paidById === COMPANY_ACCOUNT_ID) return;
      const b = bucket(paidById, row[EXP.CURRENCY]);
      if (b) b.frontedAmount += Number(row[EXP.AMOUNT] || 0);
    });
  }

  // Drawings
  const drwSheet = getSheet('Drawings');
  if (drwSheet) {
    drwSheet.getDataRange().getValues().slice(1).forEach(row => {
      const b = bucket(String(row[DRW.PERSON_ID]), row[DRW.CURRENCY]);
      if (b) b.drawings += Number(row[DRW.AMOUNT] || 0);
    });
  }

  // Flatten into one row per partner per currency they have any activity in
  const results = [];
  partners.forEach(p => {
    const currencies = Object.keys(balances[p.id]);
    if (!currencies.length) {
      results.push({ personId: p.id, personName: p.name, currency: '—',
        revenueShare: 0, expenseShare: 0, frontedAmount: 0, drawings: 0, net: 0 });
      return;
    }
    currencies.forEach(cur => {
      const b = balances[p.id][cur];
      results.push({
        personId: p.id, personName: p.name, currency: cur,
        revenueShare:  b.revenueShare,
        expenseShare:  b.expenseShare,
        frontedAmount: b.frontedAmount,
        drawings:      b.drawings,
        net: b.revenueShare - b.expenseShare + b.frontedAmount - b.drawings
      });
    });
  });

  return results.sort((a, b) => a.personName.localeCompare(b.personName) || a.currency.localeCompare(b.currency));
}
