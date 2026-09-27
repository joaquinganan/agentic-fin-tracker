/**
 * SHEETS WRITER MODULE
 * Handles all Google Sheets database operations
 */

/**
 * v1.1.2: Transactions sheet column layout (0-indexed) — Currency was moved
 * from the end (was index 11) to right after Amount, per request. Every
 * function below that reads by fixed index was updated to match:
 *   0 Date | 1 Bank | 2 Merchant | 3 Amount | 4 Currency | 5 Category |
 *   6 Description | 7 Email Subject | 8 Timestamp | 9 IsCredit |
 *   10 IsCashback | 11 Type | 12 MessageId
 */
const TX_COL = {
  DATE: 0, BANK: 1, MERCHANT: 2, AMOUNT: 3, CURRENCY: 4, CATEGORY: 5,
  DESCRIPTION: 6, SUBJECT: 7, TIMESTAMP: 8, IS_CREDIT: 9, IS_CASHBACK: 10, TYPE: 11,
  MESSAGE_ID: 12,
  TX_REF: 13   // v1.1.23: "<BANK>:<bank's own id>" when the email has one (hidden column N)
};
const TX_NUM_COLS = TX_COL.TX_REF + 1;

/**
 * v1.1.4: canonical sheet order, applied by ensureSheetOrder() below.
 * v1.1.5: swapped Raw_BDI before Raw_POPULAR, per request.
 * v1.1.7: added "Categories" at the end — a new managed reference sheet
 * (see buildOrRefreshCategoriesSheet()) documenting the current category
 * list, since a manually-maintained copy of this (e.g. one carried over
 * from the Excel workbook) has no way to stay in sync when the categorizer
 * changes in code.
 */
const CANONICAL_SHEET_ORDER = [
  "Dashboard", "Holdings", "Transactions", "Bank Transfers",
  "Raw_LAFISE", "Raw_BANESCO", "Raw_BHD", "Raw_BDI", "Raw_POPULAR",
  "Investment Ledger", "Portfolio History", "Custom Rules", "Investment Accounts", "Configuration", "Categories"
];

/**
 * v1.1.5: tab colors — every Raw_<BANK> sheet shares one common color.
 */
const TAB_COLORS = {
  "Dashboard": "#1F3864",
  "Transactions": "#0B8043",
  "Bank Transfers": "#E69138",
  "Custom Rules": "#8E63CE",
  "Configuration": "#B7472A",
  "Categories": "#666666",
  "Holdings": "#0F766E",             // v1.1.29
  "Investment Ledger": "#0F766E",
  "Investment Accounts": "#0F766E",
  "Portfolio History": "#0F766E"      // v1.1.32
};
const RAW_BANK_TAB_COLOR = "#999999";

/**
 * v1.1.4: moves every sheet named in CANONICAL_SHEET_ORDER to that relative
 * position.
 * v1.1.5: BUG FIX — moving sheets via setActiveSheet()+moveActiveSheet()
 * leaves the LAST sheet processed (Configuration) active as a side effect,
 * which is why "Monitor by Date Range" was landing on Configuration instead
 * of Dashboard. Now captures the active sheet first and restores it after —
 * callers that want to land somewhere specific activate it themselves AFTER
 * calling this. Also applies tab colors.
 */
function ensureSheetOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const previousActive = ss.getActiveSheet();

  let pos = 1;
  CANONICAL_SHEET_ORDER.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(pos);
      pos++;
    }
  });

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (TAB_COLORS[name]) {
      sheet.setTabColor(TAB_COLORS[name]);
    } else if (name.startsWith("Raw_")) {
      sheet.setTabColor(RAW_BANK_TAB_COLOR);
    }
  });

  ss.setActiveSheet(previousActive);
}

/**
 * v1.1.1: creates/returns a Raw_<BANK> sheet (Date | Type | Merchant/
 * Description | Category | Amount | Currency | Notes | Email Subject).
 * v1.1.3: Date column gets a date-only number format.
 * v1.1.4: headers translated to English.
 * v1.1.5: only Type=Transaction rows are ever written here now (see
 * saveTransaction() below).
 */
function getOrCreateRawBankSheet(bank) {
  const name = "Raw_" + bank;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(["Date", "Type", "Merchant / Description", "Category", "Amount", "Currency", "Notes", "Email Subject"]);
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold");
    ensureRowCapacity(sheet, 2000); // v1.1.19: formats below cover 2,000 rows
    sheet.getRange(2, 1, 1999).setNumberFormat('yyyy-MM-dd');
    sheet.getRange(2, 5, 1999).setNumberFormat('#,##0.00');
    ensureAutoFilter(sheet, 8, 2000);
    sheet.setTabColor(RAW_BANK_TAB_COLOR);
  }
  return sheet;
}

/**
 * v1.1.28: Bank Transfers now shows each transfer's Category (column D) — the
 * sheet where you review transfers didn't show which ones still had none.
 */
const TRANSFERS_HEADERS = ["Date", "Bank", "Beneficiary / Description", "Category", "Amount", "Currency", "Email Subject"];

/**
 * v1.1.2: dedicated sheet for Type = "Transfer" rows.
 * v1.1.4: renamed "Transferencias" → "Bank Transfers", English headers.
 */
function getOrCreateTransfersSheet() {
  const name = "Bank Transfers";
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
    sheet.appendRow(TRANSFERS_HEADERS);
    sheet.getRange(1, 1, 1, TRANSFERS_HEADERS.length).setFontWeight("bold");
    ensureRowCapacity(sheet, 2000); // v1.1.19: formats below cover 2,000 rows
    sheet.getRange(2, 1, 1999).setNumberFormat('yyyy-MM-dd');
    sheet.getRange(2, 5, 1999).setNumberFormat('#,##0.00');
    ensureAutoFilter(sheet, TRANSFERS_HEADERS.length, 2000);
    sheet.setTabColor(TAB_COLORS["Bank Transfers"]);
  }
  return sheet;
}
/**
 * v1.1.24: widens each visible column to fit its content, never narrower
 * than it is now (so a width you set by hand is kept) and never wider than
 * `max`. autoResizeColumns() fits text tightly, so `pad` leaves room for the
 * filter button in the header. Hidden columns are left alone.
 */
function autoFitColumns(sheet, options) {
  const o = Object.assign({ min: 60, max: 420, pad: 24 }, options || {});
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1 || sheet.getLastRow() < 1) return;
  for (let c = 1; c <= lastCol; c++) {
    if (sheet.isColumnHiddenByUser(c)) continue;
    const before = sheet.getColumnWidth(c);
    sheet.autoResizeColumns(c, 1);
    const fitted = sheet.getColumnWidth(c) + o.pad;
    // the cap only limits automatic growth — a width you set by hand is never reduced
    sheet.setColumnWidth(c, Math.max(before, Math.min(o.max, Math.max(o.min, fitted))));
  }
}

/**
 * v1.1.28: LOOK OF THE DATA SHEETS. Colours come from conditional formatting,
 * not from painting cells, so they follow the data: recategorize a row and its
 * chip changes; add rows and they're styled already. Each run replaces the
 * rules on these tracker-managed sheets, so it's safe to repeat.
 *   · header: navy, white, frozen
 *   · Category: a coloured chip per category (DEFAULT_CATEGORIES colours)
 *   · rows needing attention: transfers with no category (amber), rows whose
 *     merchant couldn't be read or reversals without their purchase (red)
 *   · negative amounts (reversals, refunds): green
 *   · alternate rows: a faint stripe
 */
const SHEET_THEME = {
  headerBg: '#1F3864', headerFg: '#FFFFFF', stripe: '#F7F9FC', attention: '#FFF4D6', problem: '#FDECEC',
  refund: '#2E7D32', keyBg: '#F3F6FB', keyFg: '#374151', example: '#9CA3AF'
};
const TYPE_COLORS = { 'Transfer': ['#E0F2F1', '#00695C'], 'Card Payment': ['#F3F4F6', '#4B5563'] };   // WCAG AA (was 4.39:1)

/** Which column holds what (1-based) in each tracker data sheet; null for other sheets. */
function dataSheetLayout(name) {
  if (name === TRANSACTIONS_SHEET) {
    return { cols: TX_NUM_COLS, merchant: TX_COL.MERCHANT + 1, category: TX_COL.CATEGORY + 1, amount: TX_COL.AMOUNT + 1,
      currency: TX_COL.CURRENCY + 1, type: TX_COL.TYPE + 1, transfersOnlyWhenType: true };
  }
  if (name === 'Bank Transfers') return { cols: TRANSFERS_HEADERS.length, merchant: 3, category: 4, amount: 5, currency: 6, allTransfers: true };
  if (name.indexOf('Raw_') === 0) return { cols: 8, merchant: 3, category: 4, amount: 5, currency: 6, type: 2 };
  return null;
}

const colLetter = n => String.fromCharCode(64 + n);   // data sheets stay within A..Z
const quoteForFormula = text => '"' + String(text).replace(/"/g, '""') + '"';

function styleHeader(sheet, cols) {
  sheet.getRange(1, 1, 1, cols).setBackground(SHEET_THEME.headerBg).setFontColor(SHEET_THEME.headerFg)
    .setFontWeight('bold').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 28);
  sheet.setFrozenRows(1);
}

/** The conditional-format rules of one data sheet, highest priority first (Sheets applies the first match). Pure. */
function dataSheetRules(sheet, L, palette) {
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  const col = c => sheet.getRange(2, c, rows, 1);
  const all = sheet.getRange(2, 1, rows, L.cols);
  const M = '$' + colLetter(L.merchant) + '2', C = '$' + colLetter(L.category) + '2';
  const rules = [];
  const rule = () => SpreadsheetApp.newConditionalFormatRule();
  // rows to fix: unreadable merchant / reversal whose purchase wasn't found
  rules.push(rule().whenFormulaSatisfied('=OR(' + M + '=' + quoteForFormula(GARBLED_PLACEHOLDER) + ',' + M + '=' +
    quoteForFormula(REVERSAL_UNMATCHED) + ')').setBackground(SHEET_THEME.problem).setRanges([all]).build());
  // transfers still without a category (not counted as spending until they get one)
  const noCategory = L.allTransfers ? '=AND($A2<>"",' + C + '="")'
    : L.transfersOnlyWhenType ? '=AND($A2<>"",$' + colLetter(L.type) + '2="Transfer",' + C + '="")' : null;
  if (noCategory) rules.push(rule().whenFormulaSatisfied(noCategory).setBackground(SHEET_THEME.attention).setRanges([all]).build());
  palette.forEach(p => rules.push(rule().whenTextEqualTo(p.name).setBackground(p.bg).setFontColor(p.fg)
    .setRanges([col(L.category)]).build()));
  if (L.type && L.transfersOnlyWhenType) {
    Object.keys(TYPE_COLORS).forEach(t => rules.push(rule().whenTextEqualTo(t).setBackground(TYPE_COLORS[t][0])
      .setFontColor(TYPE_COLORS[t][1]).setRanges([col(L.type)]).build()));
  }
  rules.push(rule().whenNumberLessThan(0).setFontColor(SHEET_THEME.refund).setRanges([col(L.amount)]).build());
  rules.push(rule().whenFormulaSatisfied('=AND($A2<>"",ISEVEN(ROW()))').setBackground(SHEET_THEME.stripe).setRanges([all]).build());
  return rules;
}

function styleDataSheet(sheet, L, palette) {
  styleHeader(sheet, L.cols);
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  // the whole column, so rows typed by hand or saved by older versions look the same as new ones
  sheet.getRange(2, 1, rows, 1).setNumberFormat('yyyy-MM-dd');
  sheet.getRange(2, L.amount, rows, 1).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  sheet.getRange(2, L.currency, rows, 1).setHorizontalAlignment('center');
  if (L.type) sheet.getRange(2, L.type, rows, 1).setHorizontalAlignment('center');
  sheet.setConditionalFormatRules(dataSheetRules(sheet, L, palette));
}

/** Custom Rules: category chips + a dropdown of every category (typing a new name still works). */
function styleCustomRulesSheet(sheet, palette) {
  styleHeader(sheet, 4);
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  const cat = sheet.getRange(2, 2, rows, 1);
  cat.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(palette.map(p => p.name), true).setAllowInvalid(true)
    .setHelpText('Pick a category — or type a new name to create your own. "Exclude" leaves matches out of every total.')
    .build());
  const rules = [SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=LEFT($A2,8)="(example"')
    .setFontColor(SHEET_THEME.example).setItalic(true).setRanges([sheet.getRange(2, 1, rows, 4)]).build()];
  palette.forEach(p => rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(p.name)
    .setBackground(p.bg).setFontColor(p.fg).setRanges([cat]).build()));
  sheet.setConditionalFormatRules(rules);
}

function styleConfigurationSheet(sheet) {
  styleHeader(sheet, 2);
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, 1).setBackground(SHEET_THEME.keyBg).setFontColor(SHEET_THEME.keyFg).setFontWeight('bold');
  if (last > 1) sheet.getRange(2, 2, last - 1, 1).setHorizontalAlignment('left');   // numbers and text line up
  sheet.getRange(1, 1).setNote('Written by 📊 Tracker › Setup Wizard — change your settings there.');
}

/** v1.1.28: styles every tracker sheet except the Dashboard (built by its own code). */
function styleTrackerSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const palette = categoryPalette(config && config.email ? getCustomCategoryNames(config.email) : []);
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    try {
      const layout = dataSheetLayout(name);
      if (layout) styleDataSheet(sheet, layout, palette);
      else if (name === CUSTOM_RULES_SHEET) styleCustomRulesSheet(sheet, palette);
      else if (name === CONFIG_SHEET) styleConfigurationSheet(sheet);
      else if (name === INVESTMENT_LEDGER_SHEET) styleLedgerSheet(sheet);                 // v1.1.29
      else if (name === INVESTMENT_ACCOUNTS_SHEET) styleHeader(sheet, ACCOUNTS_HEADERS.length);
      else if (name === HISTORY_SHEET) styleHeader(sheet, HISTORY_HEADERS.length);
    } catch (error) {
      Logger.log("Could not style " + name + ": " + error);
    }
  });
}

/** v1.1.28: style, then fit the columns — run after each recategorize and each Dashboard build. */
function formatDataSheets() {
  styleTrackerSheets();
  autoFitDataSheets();
}

/** v1.1.24: every data sheet — run after each recategorize and each Dashboard build. */
function autoFitDataSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    const isData = name === TRANSACTIONS_SHEET || name === 'Bank Transfers' || name.indexOf('Raw_') === 0 ||
      name === CUSTOM_RULES_SHEET || name === CONFIG_SHEET || name === INVESTMENT_LEDGER_SHEET || name === INVESTMENT_ACCOUNTS_SHEET ||
      name === HISTORY_SHEET;
    if (!isData) return;
    try { autoFitColumns(sheet); } catch (error) { Logger.log("Could not fit columns on " + name + ": " + error); }
  });
}

/**
 * v1.1.19: grows a sheet so rows 1..`rows` exist — a range past the sheet's
 * last row throws, so every bulk write/format calls this first.
 */
function ensureRowCapacity(sheet, rows) {
  const max = sheet.getMaxRows();
  if (rows > max) sheet.insertRowsAfter(max, rows - max);
}

/**
 * Keeps an auto-filter over the data (v1.1.5).
 * v1.1.19: filters were created once over 2,000 rows and never grew, so rows
 * past 2,000 fell outside the filter. When the data outgrows the filter's
 * range, the filter is recreated over a larger range and every column's
 * current filter criteria is re-applied (nothing the user had set is lost).
 */
function ensureAutoFilter(sheet, numCols, numRows) {
  try {
    const lastRow = sheet.getLastRow();
    const filter = sheet.getFilter();
    if (filter) {
      const r = filter.getRange();
      if (r.getLastRow() >= lastRow && r.getNumColumns() >= numCols) return;
      const criteria = {};
      for (let c = r.getColumn(); c <= r.getLastColumn(); c++) {
        const cr = filter.getColumnFilterCriteria(c);
        if (cr) criteria[c] = typeof cr.copy === 'function' ? cr.copy().build() : cr; // detach from the old filter
      }
      filter.remove();
      const rows = Math.max(numRows || 2000, lastRow + 500);
      const cols = Math.max(numCols, r.getNumColumns());
      ensureRowCapacity(sheet, rows);
      const nf = sheet.getRange(1, 1, rows, cols).createFilter();
      Object.keys(criteria).forEach(c => nf.setColumnFilterCriteria(Number(c), criteria[c]));
      return;
    }
    const rows = Math.max(numRows || 2000, lastRow, 2);
    ensureRowCapacity(sheet, rows);
    sheet.getRange(1, 1, rows, numCols).createFilter();
  } catch (error) {
    Logger.log("Could not set filter on " + sheet.getName() + ": " + error);
  }
}


/**
 * v1.1.5: sorts a data sheet by its Date column, most recent first, leaving
 * the header row untouched.
 */
function sortSheetByDateDesc(sheet, dateCol1Based) {
  try {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 3 || lastCol < 1) return;
    sheet.getRange(2, 1, lastRow - 1, lastCol).sort({ column: dateCol1Based, ascending: false });
  } catch (error) {
    Logger.log("Could not sort " + sheet.getName() + ": " + error);
  }
}

/**
 * v1.1.5: sorts every data sheet the monitor writes to.
 */
function sortAllDataSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tx = ss.getSheetByName(TRANSACTIONS_SHEET);
  if (tx) sortSheetByDateDesc(tx, TX_COL.DATE + 1);
  const bt = ss.getSheetByName("Bank Transfers");
  if (bt) sortSheetByDateDesc(bt, 1);
  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith("Raw_")) sortSheetByDateDesc(sheet, 1);
  });
}
/**
 * Normalizes any stored date shape (Date, legacy "d/m/yyyy" text, ISO text)
 * to 'yyyy-MM-dd' for comparisons (v1.1.5). Unchanged in v1.1.19.
 */
function normalizeDateForCompare(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const str = String(value);
  const legacy = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (legacy) {
    const day = legacy[1].padStart(2, '0');
    const month = legacy[2].padStart(2, '0');
    return `${legacy[3]}-${month}-${day}`;
  }
  const d = new Date(value);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return str;
}

/**
 * SAVE PATH — v1.1.19 (E1) rewrite.
 * Before: for EACH transaction, isDuplicate() re-read the whole
 * Transactions sheet and saveTransaction() made ~7 more calls (appendRow,
 * two formats, a filter check, plus an append to Raw_<BANK>/Bank Transfers
 * that rebuildDerivedSheets() overwrote moments later anyway). A 300-email
 * backlog meant 300 full-sheet reads — the main reason long runs hit the
 * 6-minute ceiling. Now: one read, an in-memory index, one setValues().
 * Duplicate rules are unchanged from v1.1.13:
 *  - same messageId (Gmail id + item index) → duplicate;
 *  - a LEGACY row saved before MessageId existed, with the same date + bank
 *    + amount → duplicate (the only way to recognize those rows).
 * buildExistingIndex / selectNewTransactions / transactionToRow make no
 * Sheets calls, so tests/ can exercise them directly.
 */
function legacyKey(date, bank, amount) {
  return normalizeDateForCompare(date) + '|' + bank + '|' + Math.round(Number(amount) * 100);
}

function buildExistingIndex(values) {
  const messageIds = new Set();
  const legacyKeys = new Set();
  const refs = new Set();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const id = row[TX_COL.MESSAGE_ID];
    if (id) messageIds.add(String(id));
    else legacyKeys.add(legacyKey(row[TX_COL.DATE], row[TX_COL.BANK], row[TX_COL.AMOUNT]));
    if (row[TX_COL.TX_REF]) refs.add(String(row[TX_COL.TX_REF]));
  }
  return { messageIds: messageIds, legacyKeys: legacyKeys, refs: refs };
}

function selectNewTransactions(transactions, index) {
  const fresh = [];
  let duplicates = 0;
  for (const t of transactions) {
    const id = t.messageId ? String(t.messageId) : '';
    const key = legacyKey(t.date, t.bank, t.amount);
    const ref = t.txRef ? String(t.txRef) : '';
    // v1.1.23: the bank's own transaction id wins — the same BANESCO transfer notice
    // arrived twice (identical content, two Gmail messages, so two message ids).
    if ((id && index.messageIds.has(id)) || index.legacyKeys.has(key) || (ref && index.refs.has(ref))) {
      duplicates++;
      continue;
    }
    if (id) index.messageIds.add(id); else index.legacyKeys.add(key); // same-run repeats too
    if (ref) index.refs.add(ref);
    fresh.push(t);
  }
  return { fresh: fresh, duplicates: duplicates };
}

function transactionToRow(t) {
  return [
    t.date, t.bank, t.merchant, t.amount, t.currency || 'DOP', t.category || '',
    t.description, t.subject, t.timestamp, t.isCredit ? 'YES' : 'NO',
    t.isCashback ? 'YES' : 'NO', t.type || 'Transaction', t.messageId || '', t.txRef || ''
  ];
}

/**
 * v1.1.23: pairs each reversal with its original purchase and copies the
 * original's merchant and category onto it, so the NEGATIVE reversal row
 * lands in the same Dashboard category and the two net to zero. Pure (no
 * Sheets calls). Match rules, in order:
 *  1. same bank, same amount, same table date+time (`timeKey` — BHD prints
 *     the original purchase time on the reversal row) within this run;
 *  2. same bank, same amount, same day among rows already saved (the time
 *     isn't stored in the sheet), skipping rows that are reversals themselves.
 * Each original is used at most once. Unmatched reversals keep the
 * REVERSAL_UNMATCHED merchant and are counted in the run summary.
 */
function resolveReversals(transactions, existingValues) {
  const used = new Set();
  const isReversalRow = amount => Number(amount) < 0;   // saved reversals are the negative rows
  let unmatched = 0;
  transactions.forEach((t, i) => {
    if (!t.reversal) return;
    const amount = Math.abs(Number(t.amount));
    const day = t.timeKey ? t.timeKey.replace(/^(\d{2})\/(\d{2})\/(\d{4}).*$/, '$3-$2-$1') : normalizeDateForCompare(t.date);
    let orig = null;
    transactions.forEach((o, j) => {
      if (orig || j === i || o.reversal || used.has('b' + j)) return;
      if (o.bank === t.bank && Math.abs(Number(o.amount) - amount) < 0.005 && t.timeKey && o.timeKey === t.timeKey) {
        orig = { merchant: o.merchant, category: o.category };
        used.add('b' + j);
      }
    });
    for (let r = 1; !orig && r < (existingValues || []).length; r++) {
      const row = existingValues[r];
      if (used.has('s' + r) || isReversalRow(row[TX_COL.AMOUNT])) continue;
      if (row[TX_COL.BANK] === t.bank && Math.abs(Number(row[TX_COL.AMOUNT]) - amount) < 0.005 &&
          normalizeDateForCompare(row[TX_COL.DATE]) === day) {
        orig = { merchant: row[TX_COL.MERCHANT], category: row[TX_COL.CATEGORY] };
        used.add('s' + r);
      }
    }
    if (orig) {
      t.merchant = orig.merchant;
      t.description = orig.merchant + ' (reversal)';
      t.category = orig.category;
    } else {
      unmatched++;
    }
  });
  return unmatched;
}

/** v1.1.23: adds the hidden TxRef header (column N) to a sheet created before it existed. */
function ensureTransactionsSchema(sheet) {
  const header = sheet.getRange(1, TX_COL.TX_REF + 1).getValue();
  if (header !== 'TxRef') {
    sheet.getRange(1, TX_COL.TX_REF + 1).setValue('TxRef').setFontWeight('bold');
    sheet.hideColumns(TX_COL.TX_REF + 1);
  }
}

function saveTransactions(transactions) {
  const results = { success: 0, failed: 0, duplicates: 0, reversalsUnmatched: 0 };
  if (!transactions || transactions.length === 0) return results;

  const sheet = getOrCreateSheet(TRANSACTIONS_SHEET);
  ensureTransactionsSchema(sheet);
  const existing = sheet.getDataRange().getValues();
  const selection = selectNewTransactions(transactions, buildExistingIndex(existing));
  results.duplicates = selection.duplicates;
  if (selection.fresh.length === 0) return results;
  results.reversalsUnmatched = resolveReversals(selection.fresh, existing);

  const rows = selection.fresh.map(transactionToRow);
  try {
    const startRow = sheet.getLastRow() + 1;
    ensureRowCapacity(sheet, startRow + rows.length - 1);
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    sheet.getRange(startRow, TX_COL.DATE + 1, rows.length, 1).setNumberFormat('yyyy-MM-dd');
    sheet.getRange(startRow, TX_COL.AMOUNT + 1, rows.length, 1).setNumberFormat('#,##0.00');
    ensureAutoFilter(sheet, rows[0].length, 2000);
    results.success = rows.length;
    Logger.log("✅ Saved " + rows.length + " transaction(s) in one batch");
  } catch (error) {
    results.failed = rows.length;
    Logger.log("❌ Error saving transactions: " + error);
  }
  return results;
}
const VALID_CURRENCIES = new Set(['DOP', 'USD', 'EUR', 'COP']);

/**
 * Pure per-row rules for "🔁 Recategorize" (no Sheets calls — see tests/).
 * v1.1.19:
 *  - (C3) Type comes from the SUBJECT only (detectTypeFromSubject(),
 *    03_gmailMonitor.gs). If the subject isn't a confirmed one, the saved
 *    Type is KEPT. This used to call detectTransactionType(subject,
 *    MERCHANT) — the merchant fed in as if it were the email body — so
 *    'ACH' inside "CACHAREPA" turned a real purchase into a Transfer on
 *    every run, and any Type decided from the real body got reverted.
 *    A consumo subject pins Type=Transaction, except a Cashback decided
 *    from the body at parse time, which is kept.
 *  - Category: Custom Rules first (any Type), then the defaults for
 *    Type=Transaction only — same rule as at parse time, on the same text.
 *  - IsCredit/IsCashback follow the Type for Card Payment/Cashback.
 */
function computeRecategorization(row, rawCustomRules) {
  const subject = String(row[TX_COL.SUBJECT] || '');
  const merchant = String(row[TX_COL.MERCHANT] || '');
  const description = String(row[TX_COL.DESCRIPTION] || '') || merchant;
  const oldType = row[TX_COL.TYPE] || 'Transaction';

  const subjectType = detectTypeFromSubject(subject);
  let type = subjectType || oldType;
  if (subjectType === 'Transaction' && oldType === 'Cashback') type = 'Cashback';

  let category = findCustomRuleOverride(description, rawCustomRules);
  if (!category && type === 'Transaction') category = categorizeTransaction(description);
  // v1.1.23: paying the card is never spending — explicit "Exclude"
  if (type === 'Card Payment') category = EXCLUDE_CATEGORY;
  // v1.1.23: an unmatched reversal keeps whatever category it has (normally
  // blank) instead of being guessed from the placeholder text
  if (merchant === REVERSAL_UNMATCHED) category = row[TX_COL.CATEGORY] || '';

  const oldCurrency = row[TX_COL.CURRENCY];
  const currency = (oldCurrency && !VALID_CURRENCIES.has(oldCurrency)) ? 'DOP' : oldCurrency;
  const isCredit = (type === 'Card Payment' || type === 'Cashback') ? 'YES' : row[TX_COL.IS_CREDIT];
  const isCashback = type === 'Cashback' ? 'YES' : 'NO';
  return { type: type, category: category || '', currency: currency, isCredit: isCredit, isCashback: isCashback };
}

/**
 * "🔁 Recategorize Saved Transactions" (also runs automatically at the end
 * of every monitor run since v1.1.12).
 * History: v1.1.3 created; v1.1.4 also corrects Type; v1.1.8 fixes invalid
 * currency codes; v1.1.10 added a garbled-merchant cleanup; v1.1.14 rebuilds
 * the derived sheets afterwards.
 * v1.1.19:
 *  - (C5) merchants and descriptions are NEVER overwritten any more. The
 *    garbled-text check ran here on every run and permanently replaced real
 *    names with the placeholder whenever a rule misfired ("HOLA PLAZA LAS
 *    AMERICAS" in v1.1.18; a lowercase-led name like "iTunes" would have
 *    been next — the "bank text is always uppercase" premise is already
 *    false, see the real "PedidosYa*Som Cafe"). Garbled-text detection
 *    stays in the extractors, at extraction time, where nothing is lost.
 *  - (E2) one read, pure per-row computation, and one setValues() per
 *    column that actually changed, instead of a setValue() per cell.
 *  - uses the raw Custom Rules only (getMergedCategoryRules() is gone).
 * Returns the number of cells changed.
 */
function recategorizeAllTransactions(userEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawCustomRules = getUserCustomRules(userEmail);
  let totalChanged = 0;

  const txSheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  if (txSheet && txSheet.getLastRow() > 1) {
    const numRows = txSheet.getLastRow() - 1;
    const numCols = Math.max(txSheet.getLastColumn(), TX_NUM_COLS);
    const values = txSheet.getRange(2, 1, numRows, numCols).getValues();
    const asText = v => (v === null || v === undefined) ? '' : String(v);
    const changedCols = new Set();

    values.forEach(row => {
      const r = computeRecategorization(row, rawCustomRules);
      [[TX_COL.TYPE, r.type], [TX_COL.CATEGORY, r.category], [TX_COL.CURRENCY, r.currency],
       [TX_COL.IS_CREDIT, r.isCredit], [TX_COL.IS_CASHBACK, r.isCashback]].forEach(pair => {
        const c = pair[0], v = pair[1];
        if (asText(row[c]) !== asText(v)) {
          row[c] = v;
          changedCols.add(c);
          totalChanged++;
        }
      });
    });

    changedCols.forEach(c => {
      txSheet.getRange(2, c + 1, numRows, 1).setValues(values.map(row => [row[c]]));
    });
  }

  const rebuildCounts = rebuildDerivedSheets();
  formatDataSheets();   // v1.1.24; v1.1.28: styles too
  Logger.log("✅ Recategorized — " + totalChanged + " cell(s) updated | Rebuilt: " + rebuildCounts);
  return totalChanged;
}

/**
 * v1.1.19 (M7): key used to carry the user's "Notes" (column G) across a
 * rebuild. Raw_<BANK> has no MessageId column, and adding one outside the
 * filter range would desync if the user sorted through the filter menu, so
 * notes are matched by date + merchant + amount + subject instead.
 */
function rawNoteKey(date, merchant, amount, subject) {
  return normalizeDateForCompare(date) + '|' + String(merchant || '') + '|' +
         Math.round(Number(amount) * 100) + '|' + String(subject || '');
}

/**
 * Rebuilds every Raw_<BANK> sheet and Bank Transfers from Transactions (the
 * single source of truth) — v1.1.14, so a row whose Type changed moves to
 * the right sheet instead of being patched in place.
 * v1.1.19:
 *  - (M7) the Notes column is preserved — it used to be rewritten as '' on
 *    every run, so anything typed there vanished the next morning.
 *  - a row with a blank Bank goes to Raw_UNKNOWN instead of creating a sheet
 *    literally named "Raw_".
 *  - row capacity is ensured before bulk writes.
 */
/**
 * v1.1.28: a Bank Transfers sheet from before the Category column gets the new
 * header. Its filter is recreated rather than extended: a filter's criteria are
 * kept by column number, so one set on "Amount" (old column D) would otherwise
 * end up on "Category".
 */
function migrateTransfersLayout(sheet) {
  const width = Math.max(sheet.getLastColumn(), TRANSFERS_HEADERS.length);
  const header = sheet.getRange(1, 1, 1, width).getValues()[0];
  if (header[3] === TRANSFERS_HEADERS[3]) return false;
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.getRange(1, 1, 1, width).clearContent();
  sheet.getRange(1, 1, 1, TRANSFERS_HEADERS.length).setValues([TRANSFERS_HEADERS]);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, width).clearContent();   // rewritten right after, in the new layout
  ensureAutoFilter(sheet, TRANSFERS_HEADERS.length, 2000);
  return true;
}

function rebuildDerivedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  if (!txSheet) return "no Transactions sheet";

  const data = txSheet.getDataRange().getValues();
  const byBank = {};    // bank -> Raw_<BANK> rows (A..H)
  const transfers = []; // Bank Transfers rows (A..F)

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const type = row[TX_COL.TYPE] || 'Transaction';
    const bank = String(row[TX_COL.BANK] || '').trim() || 'UNKNOWN';
    if (type === 'Transaction') {
      if (!byBank[bank]) byBank[bank] = [];
      byBank[bank].push([
        row[TX_COL.DATE], type, row[TX_COL.MERCHANT], row[TX_COL.CATEGORY],
        row[TX_COL.AMOUNT], row[TX_COL.CURRENCY], '', row[TX_COL.SUBJECT]
      ]);
    } else if (type === 'Transfer') {
      transfers.push([
        row[TX_COL.DATE], bank, row[TX_COL.MERCHANT], row[TX_COL.CATEGORY], row[TX_COL.AMOUNT],
        row[TX_COL.CURRENCY], row[TX_COL.SUBJECT]
      ]);
    }
  }

  // Every Raw_<BANK> sheet gets rebuilt — including one whose bank has no
  // rows left (e.g. everything moved to Bank Transfers), so it's emptied.
  const allBanks = new Set(Object.keys(byBank));
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith("Raw_") && name.length > 4) allBanks.add(name.substring(4));
  });

  allBanks.forEach(bank => {
    const sheet = getOrCreateRawBankSheet(bank);
    const lastRow = sheet.getLastRow();
    const notes = {};
    if (lastRow > 1) {
      const existing = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      existing.forEach(r => {
        if (r[6] !== '' && r[6] !== null) notes[rawNoteKey(r[0], r[2], r[4], r[7])] = r[6];
      });
      sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
    }
    const rows = (byBank[bank] || []).map(r => {
      const note = notes[rawNoteKey(r[0], r[2], r[4], r[7])];
      if (note !== undefined) r[6] = note;
      return r;
    });
    if (rows.length > 0) {
      ensureRowCapacity(sheet, rows.length + 1);
      sheet.getRange(2, 1, rows.length, 8).setValues(rows);
      sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-MM-dd');
      sheet.getRange(2, 5, rows.length, 1).setNumberFormat('#,##0.00');
    }
  });

  const transfersSheet = getOrCreateTransfersSheet();
  migrateTransfersLayout(transfersSheet);
  const trLastRow = transfersSheet.getLastRow();
  if (trLastRow > 1) transfersSheet.getRange(2, 1, trLastRow - 1, TRANSFERS_HEADERS.length).clearContent();
  if (transfers.length > 0) {
    ensureRowCapacity(transfersSheet, transfers.length + 1);
    transfersSheet.getRange(2, 1, transfers.length, TRANSFERS_HEADERS.length).setValues(transfers);
    transfersSheet.getRange(2, 1, transfers.length, 1).setNumberFormat('yyyy-MM-dd');
    transfersSheet.getRange(2, 5, transfers.length, 1).setNumberFormat('#,##0.00');
  }

  return Object.keys(byBank).map(b => b + "=" + byBank[b].length).join(", ") + " | Transfers=" + transfers.length;
}
const FIXED_CATEGORY_NAMES = ['Rent', 'Gym + Calisthenics', 'Telecommunications', 'Streaming & Subscriptions'];

/**
 * v1.1.19 (E3): named ranges for the Configuration values the Dashboard
 * reads. The Dashboard used to reference 'Configuration'!B3…B6/B9 by ROW
 * NUMBER — which is why "never insert a key in the middle of Configuration"
 * was a rule (v1.1.6 was exactly that bug). Named ranges are located by KEY
 * (column A), so row positions no longer matter. A missing key is appended
 * with a safe default instead of producing a #REF!.
 */
const CONFIG_NAMED_RANGES = {
  monthlyIncome:  { name: 'CFG_MONTHLY_INCOME',  fallback: 0 },
  ARS:            { name: 'CFG_ARS',             fallback: 0 },
  AFP:            { name: 'CFG_AFP',             fallback: 0 },
  taxRate:        { name: 'CFG_TAX_RATE',        fallback: 0 },
  incomeCurrency: { name: 'CFG_INCOME_CURRENCY', fallback: 'USD' },
  ISR:            { name: 'CFG_ISR',             fallback: 0 },         // v1.1.24
  deductionMode:  { name: 'CFG_DEDUCTION_MODE',  fallback: 'manual' },  // v1.1.24
  otherIncome:         { name: 'CFG_OTHER_INCOME',   fallback: 0 },      // v1.1.27
  otherIncomeCurrency: { name: 'CFG_OTHER_CURRENCY', fallback: 'DOP' }   // v1.1.27
};

/**
 * v1.1.21: (re)points a named range — removes an existing one with the same
 * name first, so re-running a build never depends on how setNamedRange()
 * treats a duplicate name.
 */
function setNamedRangeSafe(ss, name, range) {
  if (ss.getRangeByName(name)) ss.removeNamedRange(name);
  ss.setNamedRange(name, range);
}

function ensureConfigNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet || sheet.getLastRow() < 1) return false;
  const keys = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(r => String(r[0]).trim());
  Object.keys(CONFIG_NAMED_RANGES).forEach(key => {
    let row = keys.indexOf(key) + 1;
    if (row === 0) {
      sheet.appendRow([key, CONFIG_NAMED_RANGES[key].fallback]);
      keys.push(key);
      row = keys.length;
    }
    setNamedRangeSafe(ss, CONFIG_NAMED_RANGES[key].name, sheet.getRange(row, 2));
  });
  return true;
}

/**
 * DASHBOARD — v1.1.21 redesign.
 *
 * Layout (columns B..O; A and P are margins):
 *   1-2   title band (subtitle follows the selected period)
 *   3     hidden helpers: month number (C3) and period start (D3)
 *   4-5   controls — month (by name, or "Current month", which follows
 *         today), year, the three exchange rates, the date they were last
 *         edited, and Google Finance reference values
 *   7-9   KPI cards: net income · spent this month (vs. previous month) ·
 *         remaining · % of income (progress bar) · transfers to review
 *   11+   left: "Where the money went" — categories for the month with an
 *         in-cell bar, share and monthly average;
 *         right: income & deductions · fixed vs. variable · by bank
 *   then  "Month by month" table + column chart
 *   then  "Year at a glance" — 12-month grid as a heat map, selected month
 *         highlighted (this grid is what every other number reads from)
 *   then  credit cards (editable) + "Which card for what" (v1.1.22: the three cashback
 *         programs side by side; v1.1.21 had a LAFISE-only table)
 *
 * What changed vs. v1.1.20 besides the look:
 *  - the month summary shows what was ACTUALLY spent in the month. The old
 *    "Total Expenses Approx." was fixed-this-month + AVERAGE variable, a
 *    projection that didn't match any real month;
 *  - every input lives in a NAMED RANGE (DASH_MONTH, DASH_YEAR, RATE_USD,
 *    RATE_EUR, RATE_COP, RATE_UPDATED, DASH_CARDS), so a future layout
 *    change can move cells without losing what you typed. The first build
 *    of this layout carries over the v1.1.19/20 values from their old
 *    addresses (except the month, which starts on "Current month" — the
 *    old D4 had been a hard-coded 9);
 *  - the LAFISE card note is corrected per the 10/10 terms (the old
 *    "Restaurants, Gas, Groceries" default is replaced only if you never
 *    edited it).
 * Rules kept from earlier versions: grid labels in column B are PLAIN
 * category names (they're SUMIFS criteria — v1.1.7), and columns are never
 * frozen (merged cells — v1.1.6).
 */
const DASH_SHEET = 'Dashboard';
const DASH_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
                     'September', 'October', 'November', 'December'];
const DASH_CURRENT_MONTH = 'Current month';
const DASH_NAMES = {
  month: 'DASH_MONTH', year: 'DASH_YEAR', monthNum: 'DASH_MONTH_NUM', periodStart: 'DASH_PERIOD_START',
  usd: 'RATE_USD', eur: 'RATE_EUR', cop: 'RATE_COP', ratesUpdated: 'RATE_UPDATED', cards: 'DASH_CARDS'
};
// Where v1.1.19/v1.1.20 kept the same inputs — read once, on the first build of this layout.
const DASH_LEGACY = { year: 'G4', usd: 'G5', eur: 'I5', cop: 'K5', ratesUpdated: 'M5', cards: 'I8:L10' };
const DASH_DEFAULT_RATES = { usd: 59.50, eur: 64.00, cop: 0.0150 }; // seeded 23-Sep-2026; EUR/COP approximate
// Card notes (v1.1.22, from each program's terms). A note still equal to one
// of the OLD defaults is upgraded on the next build; anything you typed is kept.
const CARD_NOTES = {
  LAFISE: '10/10: 10% in 10 categories, each only from a minimum spend per billing cycle, ' +
    'with caps per category and RD$7,000/month in total. See "Which card for what" below.',
  BANESCO: 'SuperCashBack: 7% in supermarkets & liquor stores (cap RD$5,000), fast food (cap RD$500), ' +
    'electricity, water & public services (cap RD$2,000), salons, barbershops & spa (cap RD$1,000). No minimum spend.',
  BHD: 'Visa Mi País: 5% in pharmacies, restaurants and fast food (nationwide); 8% Ferretería Listo and ' +
    '6% Tiendas Corripio (each includes 3% off at checkout); BHD Stars everywhere else. Debit card purchases do not count.'
};
const OLD_DEFAULT_CARD_NOTES = [
  'Restaurants, Gas, Groceries', 'Electricity, Services', 'Health, Pharmacy',            // up to v1.1.20
  '10/10 program: 10% only in the categories listed below, each with a minimum spend and a cap per billing cycle. ' +
    'Groceries, pharmacy and utilities earn nothing.',                                  // v1.1.21
  '10/10: 10% in 10 categories, each only from a minimum spend per cycle (15th → 14th), ' +
    'with caps per category and RD$7,000/month in total. See "Which card for what" below.'  // v1.1.22-23
];
/**
 * v1.1.27: card PRODUCTS — public facts about each cashback program (name,
 * headline rate, how it works). Which card you hold, and its statement and
 * payment days, are yours: they're chosen in the Setup Wizard and stored in
 * Configuration ("cards"), never in the code.
 */
const CARD_PRODUCTS = {
  LAFISE_CLASICA:        { bank: 'LAFISE',  name: 'Clásica Mastercard', cashback: 0.10, note: CARD_NOTES.LAFISE },
  BANESCO_SUPERCASHBACK: { bank: 'BANESCO', name: 'Super Cashback',     cashback: 0.07, note: CARD_NOTES.BANESCO },
  BHD_MIPAIS:            { bank: 'BHD',     name: 'Mi País',            cashback: 0.05, note: CARD_NOTES.BHD }
};

/**
 * Configured cards → [{bank, name, cashback, closeDay, dueDay, note}]. A card
 * from the catalogue takes its name, rate and note from there; "other" keeps
 * what was typed. Pure — see tests/.
 */
function resolveCards(cards) {
  return (Array.isArray(cards) ? cards : []).filter(c => c && c.bank && c.product && c.product !== 'none').map(c => {
    const p = CARD_PRODUCTS[c.product];
    return {
      bank: String(c.bank).toUpperCase(),
      name: p ? p.name : String(c.name || '').trim(),
      cashback: p ? p.cashback : (Number(c.cashback) > 0 ? Number(c.cashback) / 100 : 0),
      closeDay: Number(c.closeDay) || null,
      dueDay: Number(c.dueDay) || null,
      note: p ? p.note : ''
    };
  });
}

/**
 * v1.1.27: rows of the Dashboard's credit-card table. From the Setup Wizard
 * when cards are configured; before that, whatever an older Dashboard had in
 * its table (DASH_CARDS), so nothing typed there is lost; otherwise a single
 * row pointing to the wizard. The table used to fall back to bank names with
 * every other cell empty (v1.1.24), which is why it showed no data.
 */
function dashboardCardRows(config, legacyRows) {
  if (config && Array.isArray(config.cards)) {
    const cards = resolveCards(config.cards);
    if (cards.length) {
      return cards.map(c => [c.bank + (c.name ? ' · ' + c.name : ''), c.cashback || '',
        c.closeDay ? 'Day ' + c.closeDay : '—', c.dueDay ? 'Day ' + c.dueDay : '—', c.note]);
    }
  }
  const filled = (legacyRows || []).filter(r => r.some(v => v !== '' && v !== null));
  if (filled.length) return filled.map(r => r.slice(0, 5));
  return [['No credit cards set up yet', '', '', '', 'Add them in 📊 Tracker › Setup Wizard.']];
}

/** Replaces a card note only while it is still one of the old built-in defaults. */
function upgradeCardNotes(cards) {
  return cards.map(row => {
    const card = String(row[0]).trim().toUpperCase();
    const note = String(row[4] === null || row[4] === undefined ? '' : row[4]).trim();
    const upgrade = CARD_NOTES[card] && (note === '' || OLD_DEFAULT_CARD_NOTES.indexOf(note) !== -1);
    const out = upgrade ? row.slice(0, 4).concat([CARD_NOTES[card]]) : row.slice();
    return out;
  });
}

/**
 * v1.1.22: "Which card for what" — the three cashback programs side by side,
 * one row per kind of spending. Sources: LAFISE 10/10 terms (Visa Clásica /
 * Mastercard Standard), BANESCO SuperCashBack reglamento (table 1), BHD Visa
 * Mi País product page (a summary — the full terms may add caps or minimums).
 * Reference only; nothing is computed from it yet. "≥" means the whole
 * category earns nothing that cycle below that amount (LAFISE all-or-nothing).
 * [spending, MCC, LAFISE, BANESCO, BHD, best per the terms]
 */
const CASHBACK_MATRIX = [
  ['Supermarkets', '5411', '—', '7% · cap RD$5,000 (shared with liquor)', 'Stars', 'BANESCO'],
  ['Pharmacies', '—', '—', '—', '5%', 'BHD'],
  ['Electricity, water, public services', '4900, 9399', '— (excluded)', '7% · cap RD$2,000', 'Stars', 'BANESCO'],
  ['Restaurants', '5812', '10% if ≥ RD$5,000 · cap RD$1,000', '—', '5%', 'LAFISE once the cycle reaches RD$5,000; otherwise BHD'],
  ['Fast food', '5814', '10% if ≥ RD$2,000 · cap RD$1,000', '7% · cap RD$500', '5%', 'LAFISE once the cycle reaches RD$2,000; otherwise BANESCO'],
  ['Gas stations', '5541, 5542', '10% if ≥ RD$5,000 · cap RD$1,000', '—', 'Stars', 'LAFISE (pays nothing below RD$5,000)'],
  ['Telecommunications', '4814, 4899', '10% if ≥ RD$3,000 · cap RD$1,000', '—', 'Stars', 'LAFISE (pays nothing below RD$3,000)'],
  ['Taxi apps (Uber, DiDi)', '4121', '10% if ≥ RD$3,000 · cap RD$500', '—', 'Stars', 'LAFISE (pays nothing below RD$3,000)'],
  ['Salons & barbershops', '7230', '10% if ≥ RD$1,500 · cap RD$500', '7% · cap RD$1,000 (with spa)', 'Stars', 'LAFISE once the cycle reaches RD$1,500; otherwise BANESCO'],
  ['Gyms', '7997', '10% if ≥ RD$1,500 · cap RD$500', '—', 'Stars', 'LAFISE (only if the gym takes cards)'],
  ['Streaming', '5815, 5968', '10% if ≥ US$10 · cap RD$500', '—', 'Stars', 'LAFISE'],
  ['Cinema', '7832, 7922', '10% if ≥ RD$1,000 · cap RD$500', '—', 'Stars', 'LAFISE'],
  ['Couriers', '4215', '10% if ≥ RD$1,000 · cap RD$500', '—', 'Stars', 'LAFISE'],
  ['Liquor stores', '5921', '—', '7% (supermarket cap)', 'Stars', 'BANESCO'],
  ['Spa', '7298', '—', '7% (salon cap)', 'Stars', 'BANESCO'],
  ['Ferretería Listo · Tiendas Corripio', '—', '—', '—', '8% · 6% (incl. 3% off at checkout)', 'BHD'],
  ['Everything else (vet, clothing, Amazon…)', '—', '—', '—', 'Stars', 'BHD (Stars — value not in the terms shared)']
];
const CASHBACK_NOTES = [
  'LAFISE 10/10 — all or nothing per category: below the minimum in a billing cycle that category earns RD$0. ' +
    'Total cap RD$7,000/month. Credited on the next statement. Excludes cash advances, taxes & public services, fees, gambling, balance transfers.',
  'BANESCO SuperCashBack — no minimum spend; caps are per month (supermarkets + liquor share one cap, salons + spa share another). ' +
    'Credited at the statement cut; the program closes the day before, so cut-day purchases count next month.',
  'BHD Visa Mi País — from the product page: caps, minimums and the value of BHD Stars are not stated there. ' +
    'Only purchases with the credit card count (the Visa Débito Intl earns nothing).',
  'All three programs pay by the merchant\'s MCC, which bank emails don\'t include — the tracker\'s categories are an approximation. ' +
    'Cashback arrives as a statement credit, never by email.'
];

const DASH_THEME = {
  navy: '#1F3864', blue: '#2E5395', accent: '#4472C4', soft: '#F3F6FB', stripe: '#F7F9FC', line: '#D9E1F2',
  input: '#FFF2CC', inputBorder: '#F4B183', muted: '#7F7F7F', total: '#DDEBF7', heat: '#F4B183',
  highlight: '#FFD966', good: '#548235', warn: '#C55A11', bad: '#C00000', grey: '#A6A6A6'
};

/**
 * Reads what the user typed on the current Dashboard: from the named
 * ranges when they exist, otherwise from the v1.1.19/20 addresses (first
 * build of this layout). Pure apart from reading cells.
 */
function readDashboardInputs(ss, sheet) {
  const inputs = { month: DASH_CURRENT_MONTH, year: null, usd: null, eur: null, cop: null, ratesUpdated: null, cards: null };
  if (!sheet) return inputs;
  const filled = v => v !== '' && v !== null && v !== undefined;
  const named = key => {
    const r = ss.getRangeByName(DASH_NAMES[key]);
    return (r && r.getSheet().getName() === DASH_SHEET) ? r : null;
  };

  if (named('usd')) {
    ['month', 'year', 'usd', 'eur', 'cop', 'ratesUpdated'].forEach(key => {
      const r = named(key);
      if (r && filled(r.getValue())) inputs[key] = r.getValue();
    });
    const cards = named('cards');
    if (cards) {
      const values = cards.getValues();
      if (values.some(row => row.some(filled))) inputs.cards = values;
    }
  } else {
    ['year', 'usd', 'eur', 'cop', 'ratesUpdated'].forEach(key => {
      const v = sheet.getRange(DASH_LEGACY[key]).getValue();
      if (filled(v)) inputs[key] = v;
    });
    const old = sheet.getRange(DASH_LEGACY.cards).getValues(); // [Card, Cashback, Main categories, Payment day]
    if (old.some(row => row.some(filled))) {
      inputs.cards = old.map(row => {
        const isLafise = String(row[0]).trim().toUpperCase() === 'LAFISE';
        return [row[0], row[1], '', row[3], row[2]];   // v1.1.24: no built-in statement dates
      });
    }
  }
  if (inputs.cards) inputs.cards = upgradeCardNotes(inputs.cards);   // v1.1.22: both paths
  if (inputs.month !== DASH_CURRENT_MONTH && DASH_MONTHS.indexOf(inputs.month) === -1) inputs.month = DASH_CURRENT_MONTH;
  return inputs;
}

/** Removes everything a previous build left behind (content alone isn't enough). */
function resetDashboardSheet(sheet) {
  sheet.getCharts().forEach(c => sheet.removeChart(c));
  sheet.clearConditionalFormatRules();
  sheet.setFrozenRows(0);
  const all = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  all.breakApart();
  all.clearDataValidations();
  sheet.clear();
  sheet.clearFormats();
  sheet.showRows(1, sheet.getMaxRows());
  sheet.setRowHeights(1, sheet.getMaxRows(), 21);
}

function buildOrRefreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const userEmail = config ? config.email : '';
  refreshAutoDeductions();   // v1.1.24: DOP automatic mode follows the current payroll parameters
  ensureConfigNamedRanges();

  let sheet = ss.getSheetByName(DASH_SHEET);
  const inputs = readDashboardInputs(ss, sheet);
  if (!sheet) sheet = ss.insertSheet(DASH_SHEET);
  else resetDashboardSheet(sheet);

  const T = DASH_THEME;
  const now = new Date();
  const year = (Number(inputs.year) >= 2000 && Number(inputs.year) <= 2100) ? Number(inputs.year) : now.getFullYear();
  const rate = key => (Number(inputs[key]) > 0 ? Number(inputs[key]) : DASH_DEFAULT_RATES[key]);
  const moneyFmt = '"RD$"#,##0;[Red]-"RD$"#,##0;"–"';
  const pctFmt = '0.0%';
  const colL = c => String.fromCharCode(64 + c);                 // B=2 … O=15
  const MONTHS12 = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const categories = getCategories().concat(userEmail ? getCustomCategoryNames(userEmail) : []);
  const banks = Object.keys(BANK_PATTERNS).filter(b => !(config && config.banksToTrack) || config.banksToTrack[b]);
  const n = categories.length;
  const put = (range, v) => (typeof v === 'string' && v.charAt(0) === '=') ? range.setFormula(v) : range.setValue(v);
  const border = SpreadsheetApp.BorderStyle;

  // ---- row map
  const R = { kpiLabel: 7, kpiValue: 8, kpiNote: 9, sec1: 11, head1: 12, catFirst: 13 };
  R.catLast = R.catFirst + n - 1;
  R.transfers = R.catLast + 1;
  R.total = R.transfers + 1;
  R.income = 12; R.netDop = 18;                                   // right block: 12..18 (v1.1.27: + other income)
  R.fvHead = 20; R.fvFirst = 21;                                  // 21..23
  R.bankHead = 25; R.bankFirst = 26; R.bankLast = R.bankFirst + banks.length - 1;
  R.monthsHead = Math.max(R.total, R.bankLast) + 2;
  R.monthsCols = R.monthsHead + 1; R.monthsFirst = R.monthsCols + 1; R.monthsLast = R.monthsFirst + 11;
  R.gridHead = R.monthsLast + 2; R.gridCols = R.gridHead + 1; R.gridFirst = R.gridCols + 1;
  R.gridLastCat = R.gridFirst + n - 1; R.gridTransfers = R.gridLastCat + 1; R.gridTotal = R.gridTransfers + 1;
  R.gridPct = R.gridTotal + 1; R.gridFixed = R.gridPct + 1; R.gridVar = R.gridFixed + 1;
  const cardRows = dashboardCardRows(config, inputs.cards);       // v1.1.27: from the Setup Wizard
  R.cardsHead = R.gridVar + 2; R.cardsCols = R.cardsHead + 1; R.cardsFirst = R.cardsCols + 1;
  R.cardsLast = R.cardsFirst + cardRows.length - 1; R.cardsHint = R.cardsLast + 1;
  R.mxHead = R.cardsHint + 2; R.mxCols = R.mxHead + 1; R.mxFirst = R.mxCols + 1;
  R.mxLast = R.mxFirst + CASHBACK_MATRIX.length - 1;
  R.notesFirst = R.mxLast + 2; R.notesLast = R.notesFirst + CASHBACK_NOTES.length - 1;
  ensureRowCapacity(sheet, R.notesLast + 2);

  const USD = '$K$4', EUR = '$M$4', COP = '$O$4', MONTH_NUM = '$C$3', PERIOD = '$D$3', NET_DOP = `$N$${R.netDop}`;
  const fxSum = criteria => [['DOP', null], ['USD', USD], ['EUR', EUR], ['COP', COP]]
    .map(([cur, r]) => `SUMIFS(Transactions!$D:$D,${criteria},Transactions!$E:$E,"${cur}")` + (r ? `*${r}` : ''))
    .join('+');
  const periodCrit = `Transactions!$A:$A,">="&${PERIOD},Transactions!$A:$A,"<"&EDATE(${PERIOD},1)`;
  const spend = `Transactions!$L:$L,"<>Card Payment",Transactions!$L:$L,"<>Cashback"`;

  // ---- helpers for the look
  const section = (row, c1, c2, text) => {
    put(sheet.getRange(row, c1, 1, c2 - c1 + 1).merge(), text);
    sheet.getRange(row, c1, 1, c2 - c1 + 1).setFontWeight('bold').setFontSize(11).setFontColor(T.navy)
      .setBorder(null, null, true, null, null, null, T.accent, border.SOLID_MEDIUM);
    sheet.setRowHeight(row, 28);
  };
  const tableHead = (row, c1, labels) => sheet.getRange(row, c1, 1, labels.length).setValues([labels])
    .setFontWeight('bold').setFontSize(8).setFontColor(T.muted).setBackground(T.soft)
    .setBorder(null, null, true, null, null, null, T.line, border.SOLID);
  const stripe = (row, c1, rows, cols) => sheet.getRange(row, c1, rows, cols).setBackgrounds(
    Array.from({ length: rows }, (_, i) => new Array(cols).fill(i % 2 ? T.stripe : '#FFFFFF')));
  const totalRow = (row, c1, cols) => sheet.getRange(row, c1, 1, cols).setFontWeight('bold').setBackground(T.total)
    .setBorder(true, null, null, null, null, null, T.accent, border.SOLID);

  // ---- canvas
  const used = sheet.getRange(1, 1, R.notesLast + 1, 16);
  used.setFontFamily('Roboto').setFontSize(10).setVerticalAlignment('middle');
  sheet.setColumnWidth(1, 18);
  // v1.1.24: column B fits its longest label. autoResizeColumns() can't be used
  // here — merged section titles and notes anchored in B would stretch it.
  const longestB = Math.max.apply(null, categories.map(c => getCategoryWithIcon(c).length).concat([28]));
  sheet.setColumnWidth(2, Math.min(380, Math.max(230, Math.round(longestB * 7.2) + 34)));
  for (let c = 3; c <= 14; c++) sheet.setColumnWidth(c, 86);
  sheet.setColumnWidth(15, 104);
  sheet.setColumnWidth(16, 18);

  // ---- 1-2 title band
  sheet.setRowHeight(1, 46);
  sheet.setRowHeight(2, 24);
  sheet.getRange(1, 1, 2, 16).setBackground(T.navy);
  sheet.getRange('B1:O1').merge().setValue('💰  Financial Tracker').setFontSize(20).setFontWeight('bold').setFontColor('#FFFFFF');
  sheet.getRange('B2:O2').merge()
    .setFormula(`="Dashboard  ·  " & TEXT(${PERIOD},"mmmm yyyy") & "  ·  all amounts in DOP-equivalent"`)
    .setFontColor(T.line);

  // ---- 3 hidden helpers
  const monthList = DASH_MONTHS.map(m => `"${m}"`).join(',');
  sheet.getRange('B3').setValue('helpers (hidden row)');
  sheet.getRange('C3').setFormula(`=IF($C$4="${DASH_CURRENT_MONTH}",MONTH(TODAY()),MATCH($C$4,{${monthList}},0))`);
  sheet.getRange('D3').setFormula('=DATE($F$4,$C$3,1)');
  sheet.hideRows(3);

  // ---- 4-5 controls
  sheet.setRowHeight(4, 30);
  const label = (a1, text) => sheet.getRange(a1).setValue(text).setFontWeight('bold').setHorizontalAlignment('right');
  label('B4', '📅  Month');
  sheet.getRange('C4:D4').merge().setValue(inputs.month).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([DASH_CURRENT_MONTH].concat(DASH_MONTHS), true).setAllowInvalid(false).build());
  label('E4', 'Year');
  sheet.getRange('F4').setValue(year).setNumberFormat('0').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberBetween(2000, 2100).setAllowInvalid(false).build());
  sheet.getRange('H4:I4').merge();
  label('H4', '💱  DOP per 1 unit');
  [['J4', 'USD', 'K4', rate('usd')], ['L4', 'EUR', 'M4', rate('eur')], ['N4', 'COP', 'O4', rate('cop')]]
    .forEach(([labelCell, text, cell, value]) => {
      label(labelCell, text);
      sheet.getRange(cell).setValue(value).setNumberFormat('0.0000').setDataValidation(SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThan(0).setAllowInvalid(false).build());
    });
  ['C4:D4', 'F4', 'K4', 'M4', 'O4'].forEach(a1 => sheet.getRange(a1).setBackground(T.input).setFontColor(T.navy)
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, null, null, T.inputBorder, border.SOLID));

  sheet.getRange('B5:F5').merge().setValue('Yellow cells are yours to edit — everything else updates by itself.')
    .setFontSize(8).setFontStyle('italic').setFontColor(T.muted);
  sheet.getRange('H5').setValue('Rates updated').setFontSize(8).setFontColor(T.muted).setHorizontalAlignment('right');
  sheet.getRange('I5').setValue(inputs.ratesUpdated !== null ? inputs.ratesUpdated : '—').setNumberFormat('yyyy-MM-dd')
    .setFontSize(8).setFontColor(T.muted).setHorizontalAlignment('center')
    .setNote('Stamped automatically when you edit a rate (onEdit, 01_main.gs).');
  sheet.getRange('J5').setValue('Google ref.').setFontSize(8).setFontColor(T.muted).setHorizontalAlignment('right');
  [['K5', 'USDDOP'], ['M5', 'EURDOP'], ['O5', 'COPDOP']].forEach(([a1, pair]) =>
    sheet.getRange(a1).setFormula(`=IFERROR(GOOGLEFINANCE("CURRENCY:${pair}"),"n/a")`).setNumberFormat('0.0000')
      .setFontSize(8).setFontColor(T.muted).setHorizontalAlignment('center'));

  const nameRange = (key, a1) => setNamedRangeSafe(ss, DASH_NAMES[key], sheet.getRange(a1));
  nameRange('month', 'C4'); nameRange('year', 'F4'); nameRange('monthNum', 'C3'); nameRange('periodStart', 'D3');
  nameRange('usd', 'K4'); nameRange('eur', 'M4'); nameRange('cop', 'O4'); nameRange('ratesUpdated', 'I5');
  sheet.setRowHeight(6, 10);

  // ---- 7-9 KPI cards
  const gt = R.gridTotal;
  const kpis = [
    { c1: 2, c2: 2, label: 'NET INCOME', value: `=${NET_DOP}`, fmt: moneyFmt,
      note: '="Gross " & TEXT(CFG_MONTHLY_INCOME,"#,##0") & " " & CFG_INCOME_CURRENCY & IF(CFG_OTHER_INCOME>0," + " & ' +
        'TEXT(CFG_OTHER_INCOME,"#,##0") & " " & CFG_OTHER_CURRENCY & " other","") & " / month"' },
    { c1: 3, c2: 5, label: `="SPENT  ·  " & UPPER(TEXT(${PERIOD},"mmm yyyy"))`, value: `=$C$${R.total}`, fmt: moneyFmt,
      note: `=IF(${MONTH_NUM}=1,"vs. previous month: n/a",IFERROR("vs. previous month: "&TEXT(INDEX($C$${gt}:$N$${gt},1,${MONTH_NUM})/INDEX($C$${gt}:$N$${gt},1,${MONTH_NUM}-1)-1,"+0%;-0%;0%"),"vs. previous month: —"))` },
    { c1: 6, c2: 8, label: 'REMAINING', value: `=${NET_DOP}-$C$${R.total}`, fmt: moneyFmt,
      note: 'Net income − spending this month' },
    { c1: 9, c2: 11, label: '% OF INCOME SPENT', value: `=IFERROR($C$${R.total}/${NET_DOP},0)`, fmt: pctFmt,
      note: `=SPARKLINE(MIN(MAX($I$8,0),1),{"charttype","bar";"max",1;"color1",IF($I$8>1,"${T.bad}",IF($I$8>0.8,"${T.warn}","${T.good}"))})` },
    { c1: 12, c2: 15, label: 'TRANSFERS TO REVIEW', value: `=$C$${R.transfers}`, fmt: moneyFmt,
      note: 'Uncategorized — add a Custom Rule (or Exclude)' }
  ];
  sheet.setRowHeight(R.kpiLabel, 22);
  sheet.setRowHeight(R.kpiValue, 36);
  sheet.setRowHeight(R.kpiNote, 22);
  kpis.forEach(k => {
    const w = k.c2 - k.c1 + 1;
    const block = sheet.getRange(R.kpiLabel, k.c1, 3, w);
    if (w > 1) block.mergeAcross();   // never merge a single column
    block.setBackground(T.soft).setHorizontalAlignment('left')
      .setBorder(null, true, null, true, null, null, '#FFFFFF', border.SOLID_THICK);
    sheet.getRange(R.kpiLabel, k.c1, 1, w).setBorder(true, null, null, null, null, null, T.accent, border.SOLID_THICK);
    put(sheet.getRange(R.kpiLabel, k.c1), k.label);
    sheet.getRange(R.kpiLabel, k.c1).setFontSize(8).setFontWeight('bold').setFontColor(T.muted);
    put(sheet.getRange(R.kpiValue, k.c1), k.value);
    sheet.getRange(R.kpiValue, k.c1).setFontSize(18).setFontWeight('bold').setFontColor(T.navy).setNumberFormat(k.fmt);
    put(sheet.getRange(R.kpiNote, k.c1), k.note);
    sheet.getRange(R.kpiNote, k.c1).setFontSize(8).setFontColor(T.muted);
  });
  sheet.setRowHeight(10, 12);

  // ---- 11+ left: where the money went
  section(R.sec1, 2, 8, `="Where the money went  ·  " & TEXT(${PERIOD},"mmmm yyyy")`);
  tableHead(R.head1, 2, ['Category', 'This month', 'Share', '', '', 'Monthly avg', 'Type']);
  const leftCount = n + 2;                                          // categories + transfers + total
  const leftRows = categories.map((cat, i) => ({ label: getCategoryWithIcon(cat), g: R.gridFirst + i,
    type: FIXED_CATEGORY_NAMES.indexOf(cat) !== -1 ? 'Fixed' : 'Variable', color: T.accent }));
  leftRows.push({ label: '↔️ Transfers (uncategorized)', g: R.gridTransfers, type: '—', color: T.grey });
  const barMax = `MAX(MAX($C$${R.catFirst}:$C$${R.transfers}),1)`;
  sheet.getRange(R.catFirst, 2, leftRows.length, 1).setValues(leftRows.map(x => [x.label]));
  sheet.getRange(R.catFirst, 3, leftRows.length, 2).setFormulas(leftRows.map((x, i) => {
    const r = R.catFirst + i;
    return [`=INDEX($C$${x.g}:$N$${x.g},1,${MONTH_NUM})`, `=IFERROR(C${r}/$C$${R.total},0)`];
  }));
  sheet.getRange(R.catFirst, 5, leftRows.length, 1).setFormulas(leftRows.map((x, i) =>
    [`=SPARKLINE(C${R.catFirst + i},{"charttype","bar";"max",${barMax};"color1","${x.color}"})`]));
  sheet.getRange(R.catFirst, 7, leftRows.length, 1).setFormulas(leftRows.map(x =>
    [`=IFERROR(AVERAGEIF($C$${x.g}:$N$${x.g},">0"),0)`]));
  sheet.getRange(R.catFirst, 8, leftRows.length, 1).setValues(leftRows.map(x => [x.type]));
  sheet.getRange(R.total, 2).setValue('TOTAL');
  sheet.getRange(R.total, 3, 1, 2).setFormulas([[`=SUM(C${R.catFirst}:C${R.transfers})`, `=IF(C${R.total}>0,1,0)`]]);
  sheet.getRange(R.total, 7).setFormula(`=IFERROR(AVERAGEIF($C$${gt}:$N$${gt},">0"),0)`);
  stripe(R.catFirst, 2, leftRows.length, 7);
  sheet.getRange(R.head1, 5, leftCount + 1, 2).mergeAcross();
  sheet.getRange(R.catFirst, 3, leftCount, 1).setNumberFormat(moneyFmt);
  sheet.getRange(R.catFirst, 7, leftCount, 1).setNumberFormat(moneyFmt);
  sheet.getRange(R.catFirst, 4, leftCount, 1).setNumberFormat(pctFmt);
  sheet.getRange(R.catFirst, 8, leftCount, 1).setFontSize(8).setFontColor(T.muted).setHorizontalAlignment('center');
  sheet.getRange(R.transfers, 2, 1, 7).setFontStyle('italic');
  totalRow(R.total, 2, 7);

  // ---- 11+ right: income & deductions
  section(R.sec1, 10, 15, 'Income & deductions');
  const incomeRows = [
    ['="Monthly salary, gross (" & CFG_INCOME_CURRENCY & ")"', '=CFG_MONTHLY_INCOME', '#,##0.00'],
    ['="ARS (" & CFG_INCOME_CURRENCY & ")"', '=CFG_ARS', '#,##0.00'],
    ['="AFP (" & CFG_INCOME_CURRENCY & ")"', '=CFG_AFP', '#,##0.00'],
    // v1.1.24: DOP salaries show ISR as an amount (automatic or entered); other
    // currencies keep a flat rate, shown as the amount it produces
    ['=IF(CFG_INCOME_CURRENCY="DOP","ISR — Impuesto Sobre la Renta" & IF(CFG_DEDUCTION_MODE="auto","  (auto)",""),' +
       '"Income tax (" & TEXT(CFG_TAX_RATE,"0.##") & "%)")',
     `=IF(CFG_INCOME_CURRENCY="DOP",CFG_ISR,N${R.income}*CFG_TAX_RATE/100)`, '#,##0.00'],
    ['="Net salary (" & CFG_INCOME_CURRENCY & ")"', `=N${R.income}-N${R.income + 1}-N${R.income + 2}-N${R.income + 3}`, '#,##0.00'],
    // v1.1.27: other income, added in full (no deductions), in its own currency
    ['="Other income (" & CFG_OTHER_CURRENCY & ", no deductions)"', '=CFG_OTHER_INCOME', '#,##0.00'],
    ['Net income (DOP-equivalent)', `=IF(CFG_INCOME_CURRENCY="DOP",N${R.income + 4},N${R.income + 4}*${USD})+` +
      `IF(CFG_OTHER_CURRENCY="DOP",N${R.income + 5},N${R.income + 5}*${USD})`, moneyFmt]
  ];
  sheet.getRange(R.income, 10, incomeRows.length, 4).mergeAcross();
  sheet.getRange(R.income, 14, incomeRows.length, 2).mergeAcross();
  incomeRows.forEach(([lab, val, fmt], i) => {
    put(sheet.getRange(R.income + i, 10), lab);
    put(sheet.getRange(R.income + i, 14), val);
    sheet.getRange(R.income + i, 14).setNumberFormat(fmt).setHorizontalAlignment('right');
  });
  stripe(R.income, 10, incomeRows.length - 1, 6);
  totalRow(R.netDop, 10, 6);

  // ---- right: fixed vs. variable
  section(R.fvHead, 10, 15, `="Fixed vs. variable  ·  " & TEXT(${PERIOD},"mmm yyyy")`);
  const fvRows = [
    ['Fixed (this month)', `=INDEX($C$${R.gridFixed}:$N$${R.gridFixed},1,${MONTH_NUM})`],
    ['Variable (this month)', `=INDEX($C$${R.gridVar}:$N$${R.gridVar},1,${MONTH_NUM})`],
    ['Variable — monthly average', `=IFERROR(AVERAGEIF($C$${R.gridVar}:$N$${R.gridVar},">0"),0)`]
  ];
  sheet.getRange(R.fvFirst, 10, fvRows.length, 4).mergeAcross();
  sheet.getRange(R.fvFirst, 14, fvRows.length, 2).mergeAcross();
  fvRows.forEach(([lab, val], i) => {
    sheet.getRange(R.fvFirst + i, 10).setValue(lab);
    sheet.getRange(R.fvFirst + i, 14).setFormula(val).setNumberFormat(moneyFmt).setHorizontalAlignment('right');
  });
  stripe(R.fvFirst, 10, fvRows.length, 6);

  // ---- right: by bank
  section(R.bankHead, 10, 15, `="By bank  ·  " & TEXT(${PERIOD},"mmm yyyy")`);
  if (banks.length) {
    const bankBarMax = `MAX(MAX($K$${R.bankFirst}:$K$${R.bankLast}),1)`;
    sheet.getRange(R.bankFirst, 11, banks.length, 2).mergeAcross();
    sheet.getRange(R.bankFirst, 13, banks.length, 3).mergeAcross();
    banks.forEach((bank, i) => {
      const r = R.bankFirst + i;
      sheet.getRange(r, 10).setValue(bank).setFontWeight('bold');
      sheet.getRange(r, 11).setFormula('=' + fxSum(`Transactions!$B:$B,"${bank}",${spend},Transactions!$F:$F,"<>Exclude",${periodCrit}`))
        .setNumberFormat(moneyFmt).setHorizontalAlignment('right');
      sheet.getRange(r, 13).setFormula(`=SPARKLINE(K${r},{"charttype","bar";"max",${bankBarMax};"color1","${T.blue}"})`);
    });
    stripe(R.bankFirst, 10, banks.length, 6);
  }

  // ---- month by month (+ chart)
  section(R.monthsHead, 2, 15, '="Month by month  ·  " & $F$4');
  tableHead(R.monthsCols, 2, ['Month', 'Spent', '% of income', 'vs. average', 'Status']);
  sheet.getRange(R.monthsFirst, 2, 12, 5).setFormulas(MONTHS12.map(m => {
    const r = R.monthsFirst + m;
    return [
      `=TEXT(DATE($F$4,${m + 1},1),"mmm")`,
      `=${colL(3 + m)}$${gt}`,
      `=IFERROR(C${r}/${NET_DOP},0)`,
      `=IF(C${r}=0,"",IFERROR(C${r}/AVERAGEIF($C$${R.monthsFirst}:$C$${R.monthsLast},">0")-1,""))`,
      `=IF(C${r}>0,"Logged",IF(DATE($F$4,${m + 1},1)>TODAY(),"Upcoming","No data"))`
    ];
  }));
  sheet.getRange(R.monthsFirst, 3, 12, 1).setNumberFormat(moneyFmt);
  sheet.getRange(R.monthsFirst, 4, 12, 1).setNumberFormat(pctFmt);
  sheet.getRange(R.monthsFirst, 5, 12, 1).setNumberFormat('+0%;-0%;0%');
  sheet.getRange(R.monthsFirst, 6, 12, 1).setFontSize(9).setHorizontalAlignment('center');
  stripe(R.monthsFirst, 2, 12, 5);
  sheet.insertChart(sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(R.monthsCols, 2, 13, 2))
    .setNumHeaders(1)
    .setPosition(R.monthsCols, 8, 0, 0)
    .setOption('title', 'Spending by month (DOP-equivalent)')
    .setOption('legend', { position: 'none' })
    .setOption('colors', [T.accent])
    .setOption('width', 700)
    .setOption('height', 270)
    .build());

  // ---- year at a glance (every other number reads from this grid)
  section(R.gridHead, 2, 15, '="Year at a glance  ·  " & $F$4 & "  (DOP-equivalent)"');
  sheet.getRange(R.gridCols, 2).setValue('Category');
  sheet.getRange(R.gridCols, 3, 1, 12).setFormulas([MONTHS12.map(m => `=DATE($F$4,${m + 1},1)`)]).setNumberFormat('mmm');
  sheet.getRange(R.gridCols, 15).setValue('Total');
  sheet.getRange(R.gridCols, 2, 1, 14).setFontWeight('bold').setFontSize(8).setFontColor(T.muted).setBackground(T.soft)
    .setBorder(null, null, true, null, null, null, T.line, border.SOLID);
  sheet.getRange(R.gridCols, 3, 1, 13).setHorizontalAlignment('center');
  const monthCrit = L => `Transactions!$A:$A,">="&${L}$${R.gridCols},Transactions!$A:$A,"<"&EDATE(${L}$${R.gridCols},1)`;

  sheet.getRange(R.gridFirst, 2, n, 1).setValues(categories.map(c => [c]));   // PLAIN names — SUMIFS criteria
  sheet.getRange(R.gridFirst, 3, n, 12).setFormulas(categories.map((cat, i) => {
    const base = `Transactions!$F:$F,$B${R.gridFirst + i},${spend}`;
    return MONTHS12.map(m => '=' + fxSum(base + ',' + monthCrit(colL(3 + m))));
  }));
  sheet.getRange(R.gridTransfers, 2).setValue('Transfers (uncategorized)');
  sheet.getRange(R.gridTransfers, 3, 1, 12).setFormulas([MONTHS12.map(m =>
    '=' + fxSum(`Transactions!$L:$L,"Transfer",Transactions!$F:$F,"",` + monthCrit(colL(3 + m))))]);
  sheet.getRange(R.gridFirst, 15, n + 1, 1).setFormulas(Array.from({ length: n + 1 }, (_, i) =>
    [`=SUM(C${R.gridFirst + i}:N${R.gridFirst + i})`]));
  sheet.getRange(R.gridTotal, 2).setValue('TOTAL');
  sheet.getRange(R.gridTotal, 3, 1, 13).setFormulas([Array.from({ length: 13 }, (_, k) =>
    `=SUM(${colL(3 + k)}${R.gridFirst}:${colL(3 + k)}${R.gridTransfers})`)]);
  sheet.getRange(R.gridPct, 2).setValue('% of net income');
  sheet.getRange(R.gridPct, 3, 1, 13).setFormulas([Array.from({ length: 13 }, (_, k) => k < 12
    ? `=IFERROR(${colL(3 + k)}${gt}/${NET_DOP},0)`
    : `=IFERROR(AVERAGEIF(C${R.gridPct}:N${R.gridPct},">0"),0)`)]);
  const flags = `$H$${R.catFirst}:$H$${R.catLast}`;
  sheet.getRange(R.gridFixed, 2, 2, 1).setValues([['Subtotal fixed'], ['Subtotal variable']]);
  sheet.getRange(R.gridFixed, 3, 2, 13).setFormulas(['Fixed', 'Variable'].map((kind, j) =>
    Array.from({ length: 13 }, (_, k) => k < 12
      ? `=SUMPRODUCT((${flags}="${kind}")*(${colL(3 + k)}${R.gridFirst}:${colL(3 + k)}${R.gridLastCat}))`
      : `=SUM(C${R.gridFixed + j}:N${R.gridFixed + j})`)));
  sheet.getRange(R.gridFirst, 3, R.gridTotal - R.gridFirst + 1, 13).setNumberFormat(moneyFmt);
  sheet.getRange(R.gridFixed, 3, 2, 13).setNumberFormat(moneyFmt);
  sheet.getRange(R.gridPct, 3, 1, 13).setNumberFormat(pctFmt);
  sheet.getRange(R.gridFirst, 2, n + 1, 1).setFontSize(9);
  sheet.getRange(R.gridFirst, 3, n + 1, 13).setFontSize(9);
  sheet.getRange(R.gridFirst, 15, n + 1, 1).setFontWeight('bold');
  sheet.getRange(R.gridTransfers, 2, 1, 14).setFontStyle('italic');
  totalRow(R.gridTotal, 2, 14);
  sheet.getRange(R.gridPct, 2, 3, 14).setFontSize(8).setFontColor(T.muted);

  // ---- credit cards (v1.1.27: from the Setup Wizard + the card catalogue)
  section(R.cardsHead, 2, 15, '💳  Credit cards');
  tableHead(R.cardsCols, 2, ['Card', 'Cashback', 'Statement close', 'Payment due', 'Notes']);
  sheet.getRange(R.cardsCols, 6, cardRows.length + 1, 10).mergeAcross();
  sheet.getRange(R.cardsFirst, 2, cardRows.length, 5).setValues(cardRows).setFontColor(T.navy);
  stripe(R.cardsFirst, 2, cardRows.length, 14);
  sheet.getRange(R.cardsFirst, 2, cardRows.length, 1).setFontWeight('bold');
  sheet.getRange(R.cardsFirst, 3, cardRows.length, 1).setNumberFormat('0%').setHorizontalAlignment('center');
  sheet.getRange(R.cardsFirst, 4, cardRows.length, 2).setHorizontalAlignment('center');
  sheet.getRange(R.cardsFirst, 6, cardRows.length, 10).setWrap(true).setFontSize(9);
  sheet.setRowHeights(R.cardsFirst, cardRows.length, 36);
  sheet.getRange(R.cardsHint, 2, 1, 14).merge().setValue('Cards, statement and payment days are set in 📊 Tracker › Setup Wizard.')
    .setFontSize(8).setFontStyle('italic').setFontColor(T.muted);
  setNamedRangeSafe(ss, DASH_NAMES.cards, sheet.getRange(R.cardsFirst, 2, cardRows.length, 5));

  // ---- which card for what (v1.1.22 — replaces the LAFISE-only table of v1.1.21)
  section(R.mxHead, 2, 15, '🧭  Which card for what — cashback by kind of spending (per each program\'s terms)');
  // B spending · C MCC · D:E LAFISE · F:G BANESCO · H:I BHD · J:O best per the terms
  sheet.getRange(R.mxCols, 2, 1, 14).setValues([['Spending', 'MCC', 'LAFISE 10/10', '', 'BANESCO SuperCashBack', '',
    'BHD Visa Mi País', '', 'Best per the terms', '', '', '', '', '']])
    .setFontWeight('bold').setFontSize(8).setFontColor(T.muted).setBackground(T.soft)
    .setBorder(null, null, true, null, null, null, T.line, border.SOLID);
  const mxRows = CASHBACK_MATRIX.length + 1;                       // header + data
  sheet.getRange(R.mxCols, 4, mxRows, 2).mergeAcross();
  sheet.getRange(R.mxCols, 6, mxRows, 2).mergeAcross();
  sheet.getRange(R.mxCols, 8, mxRows, 2).mergeAcross();
  sheet.getRange(R.mxCols, 10, mxRows, 6).mergeAcross();
  CASHBACK_MATRIX.forEach((row, i) => {
    const r = R.mxFirst + i;
    [[2, row[0]], [3, row[1]], [4, row[2]], [6, row[3]], [8, row[4]], [10, row[5]]]
      .forEach(([c, v]) => sheet.getRange(r, c).setValue(v));
  });
  stripe(R.mxFirst, 2, CASHBACK_MATRIX.length, 14);
  sheet.getRange(R.mxFirst, 2, CASHBACK_MATRIX.length, 1).setFontWeight('bold').setFontSize(9);
  sheet.getRange(R.mxFirst, 3, CASHBACK_MATRIX.length, 1).setFontSize(8).setFontColor(T.muted).setHorizontalAlignment('center');
  sheet.getRange(R.mxFirst, 4, CASHBACK_MATRIX.length, 6).setFontSize(9).setWrap(true).setHorizontalAlignment('center');
  sheet.getRange(R.mxFirst, 10, CASHBACK_MATRIX.length, 1).setFontSize(9).setWrap(true).setFontWeight('bold').setFontColor(T.navy);
  sheet.setRowHeights(R.mxFirst, CASHBACK_MATRIX.length, 32);
  sheet.getRange(R.notesFirst, 2, CASHBACK_NOTES.length, 14).mergeAcross();
  sheet.getRange(R.notesFirst, 2, CASHBACK_NOTES.length, 1).setValues(CASHBACK_NOTES.map(t => ['•  ' + t]));
  sheet.getRange(R.notesFirst, 2, CASHBACK_NOTES.length, 14).setWrap(true).setFontSize(8)
    .setFontStyle('italic').setFontColor(T.muted);
  sheet.setRowHeights(R.notesFirst, CASHBACK_NOTES.length, 32);

  // ---- conditional formatting (replaces every rule on the sheet)
  const cf = () => SpreadsheetApp.newConditionalFormatRule();
  const rng = (row, col, rows, cols) => sheet.getRange(row, col, rows, cols);
  sheet.setConditionalFormatRules([
    cf().whenNumberLessThan(0).setFontColor(T.bad).setRanges([rng(R.kpiValue, 6, 1, 1)]).build(),
    cf().whenNumberGreaterThan(0).setFontColor(T.warn).setRanges([rng(R.kpiValue, 12, 1, 1)]).build(),
    cf().whenFormulaSatisfied('=COLUMN()-2=$C$3').setBackground(T.highlight).setBold(true)
      .setRanges([rng(R.gridCols, 3, 1, 12), rng(R.gridTotal, 3, 1, 12)]).build(),
    cf().whenFormulaSatisfied(`=ROW()-${R.monthsFirst - 1}=$C$3`).setBackground(T.input).setBold(true)
      .setRanges([rng(R.monthsFirst, 2, 12, 5)]).build(),
    cf().whenTextEqualTo('Logged').setFontColor(T.good).setRanges([rng(R.monthsFirst, 6, 12, 1)]).build(),
    cf().whenTextEqualTo('No data').setFontColor(T.warn).setRanges([rng(R.monthsFirst, 6, 12, 1)]).build(),
    cf().whenTextEqualTo('Upcoming').setFontColor(T.grey).setRanges([rng(R.monthsFirst, 6, 12, 1)]).build(),
    cf().setGradientMinpoint('#FFFFFF').setGradientMaxpoint(T.heat)
      .setRanges([rng(R.gridFirst, 3, n, 12)]).build(),
    // v1.1.22: in "Which card for what", tint the cell of the card named first in the pick
    cf().whenFormulaSatisfied(`=LEFT($J${R.mxFirst},6)="LAFISE"`).setBackground('#E2EFDA')
      .setRanges([rng(R.mxFirst, 4, CASHBACK_MATRIX.length, 1)]).build(),
    cf().whenFormulaSatisfied(`=LEFT($J${R.mxFirst},7)="BANESCO"`).setBackground('#E2EFDA')
      .setRanges([rng(R.mxFirst, 6, CASHBACK_MATRIX.length, 1)]).build(),
    cf().whenFormulaSatisfied(`=LEFT($J${R.mxFirst},3)="BHD"`).setBackground('#E2EFDA')
      .setRanges([rng(R.mxFirst, 8, CASHBACK_MATRIX.length, 1)]).build()
  ]);

  // ---- finish
  sheet.setFrozenRows(5);                // title + controls stay visible; never freeze columns (v1.1.6)
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(TAB_COLORS[DASH_SHEET]);

  buildOrRefreshCategoriesSheet(userEmail);
  formatDataSheets();                    // v1.1.24; v1.1.28: styles too
  ensureSheetOrder();
  ss.setActiveSheet(sheet);              // LAST — creating the Categories sheet changes the active sheet
  Logger.log("✅ Dashboard built/refreshed (" + n + " categories, " + banks.length + " banks, year " + year + ")");
  return sheet;
}

/**
 * Categories reference sheet — auto-generated, never hand-edited (v1.1.7).
 * v1.1.19: also lists the user's Custom Rules — keywords added to a default
 * category, categories that exist only in Custom Rules, and the reserved
 * "Exclude" — so every rule in effect is visible in one place. Written with
 * one setValues() instead of four calls per row.
 */
function buildOrRefreshCategoriesSheet(userEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Categories");
  if (!sheet) {
    sheet = ss.insertSheet("Categories");
  } else {
    sheet.clear();
    sheet.clearFormats();
  }

  sheet.getRange("A1").setValue("Categories (reference)").setFontWeight("bold").setFontSize(16);
  sheet.getRange("A2").setValue(
    "Auto-generated from DEFAULT_CATEGORIES (02_categorizer.gs) and your Custom Rules every time the Dashboard rebuilds — not hand-edited.")
    .setFontStyle("italic").setFontColor("#808080").setFontSize(9);
  sheet.getRange(4, 1, 1, 4).setValues([["Icon", "Category", "Type", "Keywords"]])
    .setFontWeight("bold").setBackground("#1F3864").setFontColor("#FFFFFF");

  const custom = userEmail ? getUserCustomRules(userEmail) : {};
  const rows = getCategories().map(cat => {
    const data = DEFAULT_CATEGORIES[cat];
    const extra = custom[cat] ? "   |   Custom Rules: " + custom[cat].join(', ') : "";
    return [data.icon, cat, FIXED_CATEGORY_NAMES.includes(cat) ? 'Fixed' : 'Variable',
            (data.keywords || []).join(', ') + extra];
  });
  // v1.1.23: "Exclude" is always listed — Card Payment rows get it automatically
  const excludeKey = Object.keys(custom).find(c => c.toUpperCase() === EXCLUDE_CATEGORY.toUpperCase());
  rows.push(['🚫', EXCLUDE_CATEGORY, 'Excluded from totals', 'Automatic for every Card Payment' +
    (excludeKey ? "   |   Custom Rules: " + custom[excludeKey].join(', ') : '')]);
  Object.keys(custom).filter(c => !DEFAULT_CATEGORIES[c] && c !== excludeKey).forEach(c => {
    rows.push(['📌', c, 'Variable', "Custom Rules: " + custom[c].join(', ')]);
  });

  sheet.getRange(5, 1, rows.length, 4).setValues(rows);
  sheet.getRange(5, 1, rows.length, 1).setHorizontalAlignment("center");
  // v1.1.28: each category in the colours its chip has in the data sheets
  const colours = {};
  categoryPalette(Object.keys(custom)).forEach(p => { colours[p.name] = p; });
  const chip = rows.map(r => colours[r[1]] || { bg: '#FFFFFF', fg: '#1F2937' });
  sheet.getRange(5, 2, rows.length, 1).setBackgrounds(chip.map(c => [c.bg])).setFontColors(chip.map(c => [c.fg])).setFontWeight("bold");
  sheet.getRange(5, 4, rows.length, 1).setFontSize(9).setFontStyle("italic").setWrap(true);
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 600);
  sheet.setFrozenRows(4);
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(TAB_COLORS["Categories"]);
  Logger.log("✅ Categories sheet built/refreshed");
  return sheet;
}
