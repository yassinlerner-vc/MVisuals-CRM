// ============================================================
//  SHEET INDEX — one-click map of every tab in the workbook.
//  Rebuild anytime via the menu; it never needs manual upkeep
//  because gids/urls are pulled live, not stored.
// ============================================================
const SHEET_INDEX_META = {
  // ── Module data (app-managed — edit via the CRM forms, not the sheet) ──
  'Accounts':              { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Leads/Clients — edit via the Accounts module.' },
  'Team':                  { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Team members & partners — edit via the Team module.' },
  'Drawings':               { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Partner withdrawals — edit via the Drawings module.' },
  'Expense_Catalog':        { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Master expense list — edit via the Expenses module.' },
  'Expenses':                { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Logged expense instances — edit via the Expenses module.' },
  'Expense_Split':          { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Per-partner expense splits — written automatically, do not hand-edit.' },
  'Revenue_Distribution':   { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Per-partner revenue shares — written automatically, do not hand-edit.' },
  'Projects':               { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Confirmed projects — edit via the Projects module.' },
  'Payments':                { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Project payments — edit via the Projects module.' },
  'Project_Items':          { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Scope-of-work snapshot per project — written on quotation confirm.' },
  'Quotations':              { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Quotations — edit via the Quotations module.' },
  'Quote_Items':             { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Quotation line items — edit via the Quotations module.' },

  // ── Editable config / reference — safe & expected to hand-edit ──
  'Quotation Settings':      { category: 'Editable Config & Reference', type: 'Editable', desc: 'Currency list offered on quotations.' },
  'Items':                   { category: 'Editable Config & Reference', type: 'Editable', desc: 'Item catalog (name/description/default price) for quotations.' },
  'Branding':                { category: 'Editable Config & Reference', type: 'Editable', desc: 'Company name, colors, contact info shown on PDFs.' },
  'Bank_Details':            { category: 'Editable Config & Reference', type: 'Editable', desc: 'Bank/Instapay details shown on quotation PDFs.' },
  'Quotation_Terms':         { category: 'Editable Config & Reference', type: 'Editable', desc: 'Terms section printed on every quotation PDF.' },
  'Account Sources':         { category: 'Editable Config & Reference', type: 'Editable', desc: 'Dropdown list — Accounts "Source" field.' },
  'Account Channels':        { category: 'Editable Config & Reference', type: 'Editable', desc: 'Dropdown list — Accounts "Channel" field.' },
  'Countries':               { category: 'Editable Config & Reference', type: 'Editable', desc: 'Dropdown list — Accounts "Country" field.' },
  'Industries':              { category: 'Editable Config & Reference', type: 'Editable', desc: 'Dropdown list — Accounts "Industry" field.' },

  // ── Logs — view only ──
  'Accounts_Log':            { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Accounts.' },
  'Team_Log':                { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Team.' },
  'Drawings_Log':            { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Drawings.' },
  'Expenses_Log':            { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Expenses & Expense_Catalog.' },
  'Payments_Log':            { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Payments.' },
  'Revenue_Distribution_Log':{ category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Revenue_Distribution.' },
  'Quotations_Log':          { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Quotations & Quote_Items.' },
  'Projects_Log':            { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Projects.' }
};

const SHEET_INDEX_CATEGORY_ORDER = [
  'Module Data (app-managed)',
  'Editable Config & Reference',
  'Logs (view-only)',
  '⚠️ Unrecognized / Review'
];

const SHEET_INDEX_CATEGORY_COLOR = {
  'Module Data (app-managed)':   '#e8ecfd',
  'Editable Config & Reference': '#d1f2e1',
  'Logs (view-only)':            '#f0f2fa',
  '⚠️ Unrecognized / Review':    '#fff3cd'
};

const SHEET_INDEX_NAME = 'Sheet Info';

function buildSheetIndex() {
  const ss   = SpreadsheetApp.getActive();
  const ssId = ss.getId();

  let indexSheet = ss.getSheetByName(SHEET_INDEX_NAME);
  if (!indexSheet) {
    indexSheet = ss.insertSheet(SHEET_INDEX_NAME, 0);
  } else {
    indexSheet.clear();
  }

  const headers = ['Category', 'Sheet', 'Type', 'Description', 'Open'];
  indexSheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4361ee')
    .setFontColor('#ffffff');

  const byCategory = {};
  SHEET_INDEX_CATEGORY_ORDER.forEach(c => byCategory[c] = []);

  ss.getSheets()
    .filter(s => s.getName() !== SHEET_INDEX_NAME)
    .forEach(sheet => {
      const name = sheet.getName();
      const meta = SHEET_INDEX_META[name] || {
        category: '⚠️ Unrecognized / Review',
        type: '?',
        desc: 'Not referenced anywhere in Apps Script — confirm it is still needed.'
      };
      const url = `https://docs.google.com/spreadsheets/d/${ssId}/edit#gid=${sheet.getSheetId()}`;
      byCategory[meta.category].push([
        meta.category, name, meta.type, meta.desc, `=HYPERLINK("${url}","Open →")`
      ]);

      // Optional: color-code the actual tab to match its category
      const tabColor = SHEET_INDEX_CATEGORY_COLOR[meta.category];
      if (tabColor) sheet.setTabColor(tabColor);
    });

  const rows = [];
  SHEET_INDEX_CATEGORY_ORDER.forEach(cat => rows.push(...byCategory[cat]));

  if (rows.length) {
    indexSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  indexSheet.setColumnWidth(1, 210);
  indexSheet.setColumnWidth(2, 190);
  indexSheet.setColumnWidth(3, 110);
  indexSheet.setColumnWidth(4, 430);
  indexSheet.setColumnWidth(5, 90);
  indexSheet.setFrozenRows(1);

  ss.setActiveSheet(indexSheet);
  ss.moveActiveSheet(1);

  SpreadsheetApp.getUi().alert('Sheet index rebuilt — ' + rows.length + ' sheets listed.');
}
