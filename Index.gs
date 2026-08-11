// ============================================================
//  HOME — the workbook's landing tab. Always the active sheet
//  on open. Ruya logo + welcome message on the left, full sheet
//  index on the right. Rebuild manually (menu) whenever a sheet
//  is added, renamed, or removed.
// ============================================================
const HOME_SHEET_NAME      = 'Home';
const HOME_WELCOME_ROW     = 1;
const HOME_LOGO_ANCHOR_ROW = 3;
const HOME_LOGO_ANCHOR_COL = 1;   // column A
const HOME_LOGO_WIDTH_PX   = 480;
const HOME_LOGO_HEIGHT_PX  = 160; // adjust to match the logo's real aspect ratio
const HOME_TABLE_START_COL = 10;  // column J — clears the logo area (A:H)
const HOME_TABLE_START_ROW = 2;

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
  'Bank_Details':            { category: 'Editable Config & Reference', type: 'Editable',  desc: 'Bank/Instapay details shown on quotation PDFs.' },
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

// Category column highlight — background + text color
const SHEET_INDEX_CATEGORY_STYLE = {
  'Module Data (app-managed)':   { bg: '#c9daf8', fg: '#1a1a2e' }, // light blue
  'Editable Config & Reference': { bg: '#fff3cd', fg: '#856404' }, // yellow
  'Logs (view-only)':            { bg: '#d9d9d9', fg: '#555555' }, // grey
  '⚠️ Unrecognized / Review':    { bg: '#f8d7da', fg: '#842029' }  // red
};

// Physical tab color — same palette, so the tab strip matches the table
const SHEET_INDEX_TAB_COLOR = {
  'Module Data (app-managed)':   '#c9daf8',
  'Editable Config & Reference': '#fff3cd',
  'Logs (view-only)':            '#d9d9d9',
  '⚠️ Unrecognized / Review':    '#f8d7da'
};

function buildSheetIndex() {
  const ss   = SpreadsheetApp.getActive();
  const ssId = ss.getId();

  let homeSheet = ss.getSheetByName(HOME_SHEET_NAME);
  if (!homeSheet) {
    homeSheet = ss.insertSheet(HOME_SHEET_NAME, 0);
  } else {
    homeSheet.getImages().forEach(img => img.remove()); // clear() doesn't remove images
    homeSheet.clear();
  }

  insertHomeWelcomeAndLogo(homeSheet);

  const headers = ['Category', 'Sheet', 'Type', 'Description', 'Open'];
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
        meta.category, name, meta.type, meta.desc, `=HYPERLINK("${url}","Open →")`
      ]);

      const tabColor = SHEET_INDEX_TAB_COLOR[meta.category];
      if (tabColor) sheet.setTabColor(tabColor);
    });

  let row = HOME_TABLE_START_ROW + 1;
  let totalRows = 0;
  SHEET_INDEX_CATEGORY_ORDER.forEach(cat => {
    const rows = byCategory[cat];
    if (!rows.length) return;
    homeSheet.getRange(row, HOME_TABLE_START_COL, rows.length, headers.length).setValues(rows);
    const style = SHEET_INDEX_CATEGORY_STYLE[cat];
    if (style) {
      homeSheet.getRange(row, HOME_TABLE_START_COL, rows.length, 1)
        .setBackground(style.bg)
        .setFontColor(style.fg)
        .setFontWeight('bold');
    }
    row += rows.length;
    totalRows += rows.length;
  });

  homeSheet.setColumnWidth(HOME_TABLE_START_COL,     190); // Category
  homeSheet.setColumnWidth(HOME_TABLE_START_COL + 1, 190); // Sheet
  homeSheet.setColumnWidth(HOME_TABLE_START_COL + 2, 110); // Type
  homeSheet.setColumnWidth(HOME_TABLE_START_COL + 3, 430); // Description
  homeSheet.setColumnWidth(HOME_TABLE_START_COL + 4, 90);  // Open
  homeSheet.setFrozenRows(HOME_TABLE_START_ROW);

  ss.setActiveSheet(homeSheet);
  ss.moveActiveSheet(1);

  SpreadsheetApp.getUi().alert('Home rebuilt — ' + totalRows + ' sheets listed.');
}

// ============================================================
//  WELCOME MESSAGE + LOGO — reuses RUYA_LOGO_BASE64 already
//  defined in Logo.gs, so there's only one copy of the asset
//  in the whole project.
// ============================================================
function insertHomeWelcomeAndLogo(sheet) {
  sheet.getRange(HOME_WELCOME_ROW, HOME_LOGO_ANCHOR_COL, 1, 8).merge();
  sheet.getRange(HOME_WELCOME_ROW, HOME_LOGO_ANCHOR_COL)
    .setValue('Welcome to Ruya')
    .setFontSize(13)
    .setFontStyle('italic')
    .setFontColor('#888888')
    .setHorizontalAlignment('center');

  try {
    const match = RUYA_LOGO_BASE64.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
    if (!match) return;
    const blob  = Utilities.newBlob(Utilities.base64Decode(match[2]), match[1], 'ruya_logo');
    const image = sheet.insertImage(blob, HOME_LOGO_ANCHOR_COL, HOME_LOGO_ANCHOR_ROW);
    image.setWidth(HOME_LOGO_WIDTH_PX).setHeight(HOME_LOGO_HEIGHT_PX);
  } catch(e) {
    Logger.log('Home logo insert error: ' + e);
  }
}
