// ============================================================
//  HOME — the workbook's landing tab. Always the active sheet
//  on open — refreshHome() runs automatically from onOpen() in
//  Main.gs, no menu item needed. It only rewrites the welcome
//  text range and the sheet-index table range; it never touches
//  images, so the logo (inserted manually, once, via
//  Insert → Image → Image over cells) survives every refresh.
//
//  Logo placement guide (do this once, by hand, in the sheet):
//    Anchor roughly at cell A3, resize to cover rows 3–24,
//    columns A–F (1–6).
// ============================================================
const HOME_SHEET_NAME       = 'Home';
const HOME_WELCOME_ROW      = 1;
const HOME_LOGO_START_ROW   = 3;
const HOME_LOGO_END_ROW     = 24;
const HOME_LOGO_START_COL   = 2;   // column B — logo lives here now
const HOME_LOGO_END_COL     = 7;   // column G
const HOME_WELCOME_COL_SPAN = 6;   // B–G, same width as the logo
const HOME_TABLE_START_COL  = 9;   // column I
const HOME_TABLE_START_ROW  = 2;

const SHEET_INDEX_META = {
  // ── Module data (app-managed — edit via the CRM forms, not the sheet) ──
  'Accounts':                { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Leads/Clients — edit via the Accounts module.' },
  'Team':                    { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Team members & partners — edit via the Team module.' },
  'Drawings':                { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Partner withdrawals — edit via the Drawings module.' },
  'Expense_Catalog':         { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Master expense list — edit via the Expenses module.' },
  'Expenses':                { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Logged expense instances — edit via the Expenses module.' },
  'Expense_Split':           { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Per-partner expense splits — written automatically, do not hand-edit.' },
  'Revenue_Distribution':    { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Per-partner revenue shares — written automatically, do not hand-edit.' },
  'Projects':                { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Confirmed projects — edit via the Projects module.' },
  'Payments':                { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Project payments — edit via the Projects module.' },
  'Project_Items':           { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Scope-of-work snapshot per project — written on quotation confirm.' },
  'Quotations':              { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Quotations — edit via the Quotations module.' },
  'Quote_Items':             { category: 'Module Data (app-managed)', type: 'App-managed', desc: 'Quotation line items — edit via the Quotations module.' },

  // ── Editable config / reference — safe & expected to hand-edit ──
  'Quotation Settings':      { category: 'Editable Config & Reference', type: 'Editable',  desc: 'Currency list offered on quotations.' },
  'Items':                   { category: 'Editable Config & Reference', type: 'Editable',  desc: 'Item catalog (name/description/default price) for quotations.' },
  'Branding':                { category: 'Editable Config & Reference', type: 'Editable',  desc: 'Company name, colors, contact info shown on PDFs.' },
  'Bank_Details':            { category: 'Editable Config & Reference', type: 'Editable',  desc: 'Bank transfer details shown on quotation PDFs.' },
  'Instapay Details':        { category: 'Editable Config & Reference', type: 'Editable',  desc: 'Instapay display value + link shown on quotation PDFs.' },
  'Quotation_Terms':         { category: 'Editable Config & Reference', type: 'Editable',  desc: 'Terms section printed on every quotation PDF.' },
  'Account Sources':         { category: 'Editable Config & Reference', type: 'Editable',  desc: 'Dropdown list — Accounts "Source" field.' },
  'Account Channels':        { category: 'Editable Config & Reference', type: 'Editable',  desc: 'Dropdown list — Accounts "Channel" field.' },
  'Countries':               { category: 'Editable Config & Reference', type: 'Editable',  desc: 'Dropdown list — Accounts "Country" field.' },
  'Industries':              { category: 'Editable Config & Reference', type: 'Editable',  desc: 'Dropdown list — Accounts "Industry" field.' },
  'Account Status':          { category: 'Editable Config & Reference', type: 'Reference', desc: 'Valid Status values for Accounts (Lead/Client/Archived). Set automatically by the app — not a live dropdown source.' },

  // ── Logs — view only ──
  'Accounts_Log':             { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Accounts.' },
  'Team_Log':                 { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Team.' },
  'Drawings_Log':              { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Drawings.' },
  'Expenses_Log':              { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Expenses & Expense_Catalog.' },
  'Payments_Log':              { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Payments.' },
  'Revenue_Distribution_Log':  { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Revenue_Distribution.' },
  'Quotations_Log':            { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Quotations & Quote_Items.' },
  'Projects_Log':              { category: 'Logs (view-only)', type: 'Log', desc: 'Change history for Projects.' }
};

const SHEET_INDEX_CATEGORY_ORDER = [
  'Module Data (app-managed)',
  'Editable Config & Reference',
  'Logs (view-only)',
  '⚠️ Unrecognized / Review'
];

const SHEET_INDEX_CATEGORY_STYLE = {
  'Module Data (app-managed)':   { bg: '#c9daf8', fg: '#1a1a2e' },
  'Editable Config & Reference': { bg: '#fff3cd', fg: '#856404' },
  'Logs (view-only)':            { bg: '#d9d9d9', fg: '#555555' },
  '⚠️ Unrecognized / Review':    { bg: '#f8d7da', fg: '#842029' }
};

const SHEET_INDEX_TAB_COLOR = {
  'Module Data (app-managed)':   '#c9daf8',
  'Editable Config & Reference': '#fff3cd',
  'Logs (view-only)':            '#d9d9d9',
  '⚠️ Unrecognized / Review':    '#f8d7da'
};

// ============================================================
//  REFRESH HOME — called automatically by onOpen() in Main.gs.
//  Can also be run manually from the Apps Script editor
//  (select refreshHome in the function dropdown ▸ Run) if you
//  add/rename/remove a sheet mid-session and want the table to
//  catch up immediately instead of waiting for the next open.
// ============================================================
function refreshHome() {
  const ss   = SpreadsheetApp.getActive();
  const ssId = ss.getId();

  let homeSheet = ss.getSheetByName(HOME_SHEET_NAME);
  if (!homeSheet) {
    homeSheet = ss.insertSheet(HOME_SHEET_NAME, 0);
  }

  writeHomeWelcome(homeSheet);

  const headers = ['Sheet', 'Type', 'Description', 'Open'];

  // Clear only the table's own range — never sheet.clear(), and never
  // touch images, so a manually-placed logo is never wiped out.
  homeSheet.getRange(HOME_TABLE_START_ROW, HOME_TABLE_START_COL, 200, headers.length).clear();

  homeSheet.getRange(HOME_TABLE_START_ROW, HOME_TABLE_START_COL, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1e8a4a')
    .setFontColor('#ffffff');

  const byCategory = {};
  SHEET_INDEX_CATEGORY_ORDER.forEach(c => byCategory[c] = []);

  ss.getSheets()
    .filter(s => s.getName() !== HOME_SHEET_NAME)
    .forEach(sheet => {
      const name = sheet.getName();
      const meta = SHEET_INDEX_META[name] || {
        category: '⚠️ Unrecognized / Review',
        type: '?',
        desc: 'Not referenced anywhere in Apps Script — confirm it is still needed.'
      };
      const url = `https://docs.google.com/spreadsheets/d/${ssId}/edit#gid=${sheet.getSheetId()}`;
      byCategory[meta.category].push([
        name, meta.type, meta.desc, `=HYPERLINK("${url}","Open →")`
      ]);

      const tabColor = SHEET_INDEX_TAB_COLOR[meta.category];
      if (tabColor) sheet.setTabColor(tabColor);
    });

  const TYPE_COL_OFFSET = 1; // "Type" is the 2nd column (0-indexed) in the new 4-col layout

  let row = HOME_TABLE_START_ROW + 1;
  let totalRows = 0;
  SHEET_INDEX_CATEGORY_ORDER.forEach(cat => {
    const rows = byCategory[cat];
    if (!rows.length) return;
    homeSheet.getRange(row, HOME_TABLE_START_COL, rows.length, headers.length).setValues(rows);
    const style = SHEET_INDEX_CATEGORY_STYLE[cat];
    if (style) {
      homeSheet.getRange(row, HOME_TABLE_START_COL + TYPE_COL_OFFSET, rows.length, 1)
        .setBackground(style.bg)
        .setFontColor(style.fg)
        .setFontWeight('bold');
    }
    row += rows.length;
    totalRows += rows.length;
  });

  homeSheet.setColumnWidth(HOME_TABLE_START_COL,     190); // Sheet
  homeSheet.setColumnWidth(HOME_TABLE_START_COL + 1, 150); // Type
  homeSheet.setColumnWidth(HOME_TABLE_START_COL + 2, 430); // Description
  homeSheet.setColumnWidth(HOME_TABLE_START_COL + 3, 90);  // Open
  homeSheet.setFrozenRows(HOME_TABLE_START_ROW);

  ss.setActiveSheet(homeSheet);
  ss.moveActiveSheet(1);

  return totalRows;
}
// ============================================================
//  WELCOME MESSAGE — text only. The logo is a manually-placed
//  image (see placement guide above) and is never touched here.
// ============================================================
function writeHomeWelcome(sheet) {
  // Break apart any old merge shape (e.g. from the previous 8-column
  // version) before re-merging at the new 6-column width.
  sheet.getRange(HOME_WELCOME_ROW, 1, 1, 10).breakApart();
  const range = sheet.getRange(HOME_WELCOME_ROW, HOME_LOGO_START_COL, 1, HOME_WELCOME_COL_SPAN);
  range.merge();
  range.setValue('Welcome')
    .setFontSize(30)
    .setFontStyle('italic')
    .setFontColor('#1A1A1A')
    .setHorizontalAlignment('center');
}
