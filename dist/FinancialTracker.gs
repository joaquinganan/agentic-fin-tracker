/**
 * Financial Tracker v1.1.67 — https://github.com/joaquinganan/agentic-fin-tracker
 *
 * ONE file: in Extensions › Apps Script, this is the only code file of the project.
 * To update: select everything in this file (Ctrl+A), paste the new version, save (Ctrl+S).
 *
 * Generated from src/ by `npm run bundle` — edit src/, not this file.
 */

// ====================================================================================================
// 01_main.gs
// ====================================================================================================

/**
 * FINANCIAL TRACKER SYSTEM
 * Main Orchestrator
 * 
 * Purpose: Track all transactions from Gmail notifications
 * 
 * HOW TO USE:
 * 1. Copy all .gs files into Google Apps Script editor
 * 2. Run: onOpen()
 * 3. Click menu "📊 Tracker" > "Setup Wizard"
 * 4. Authorize Gmail & Sheets access
 * 5. Enter your financial data
 * 6. Monitor runs automatically
 */

// v1.1.12: was stuck at "1.0.0" since the very first version, never updated
// — logged at the start of every monitor/recategorize run (View > Logs) so
// it's possible to tell at a glance whether a specific run used the latest
// deployed code, instead of guessing after the fact. Bump this whenever you
// paste in an update.
const SCRIPT_VERSION = "1.1.67"; // bump on every release (v1.1.19 fixed it being stuck at 1.1.12)
const SHEET_NAME = "Financial Tracker";
// v1.1.4: renamed "Config" → "Configuration" and (below) "CustomRules" →
// "Custom Rules", to match the requested sheet naming/order and keep
// spacing/casing consistent with the rest of the English UI.
const CONFIG_SHEET = "Configuration";
const TRANSACTIONS_SHEET = "Transactions";
const CUSTOM_RULES_SHEET = "Custom Rules";

// ============================================
// MAIN ENTRY POINTS
// ============================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 Tracker")
    .addItem("📘 Start here", "openStartHere")   // v1.1.47
    .addItem("🔧 Setup Wizard", "openSetupWizard")
    .addItem("🔄 Monitor Gmail Now", "runGmailMonitor")
    .addItem("📅 Monitor by Date Range", "openDateRangeDialog")
    .addItem("📊 Dashboard", "openDashboard")
    .addItem("🔁 Recategorize Saved Transactions", "recategorizeAllTransactionsPrompt")
    .addItem("📬 Send Daily Summary Now", "sendDailySummaryNow")
    .addItem("🗓️ Send Monthly Summary Now", "sendMonthlySummaryNow")
    .addItem("📈 Refresh Investments", "refreshInvestmentsNow")
    .addItem("➕ Add Balance or Deposit", "openValuationDialog")
    .addItem("📋 Paste Broker Positions", "openPortfolioPasteDialog")   // v1.1.48
    .addItem("🙈 Show / Hide Settings Tabs", "toggleSettingsTabs")
    .addItem("⚙️ View Config", "viewConfig")
    .addSeparator()
    .addItem("🗑️ Reset System", "resetSystem")
    .addToUi();
}

/**
 * Simple trigger — stamps the Dashboard's "Rates updated" cell whenever one
 * of the three exchange rates is edited (v1.1.19). v1.1.21: finds the cells
 * through their named ranges (RATE_USD / RATE_EUR / RATE_COP /
 * RATE_UPDATED) instead of fixed addresses, so a layout change can't break
 * it. Simple triggers run on their own; nothing to install.
 */
function onEdit(e) {
  // v1.1.54: a category typed in Bank Transfers or Incoming Transfers reaches Transactions (and so the Dashboard) at once
  try { onEditTransferCategory(e); } catch (err) { /* never surface an error from a convenience trigger */ }
  try {
    const range = e && e.range;
    if (!range || range.getSheet().getName() !== "Dashboard") return;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const overlaps = r => r && r.getSheet().getName() === "Dashboard" &&
      r.getRow() <= range.getLastRow() && range.getRow() <= r.getLastRow() &&
      r.getColumn() <= range.getLastColumn() && range.getColumn() <= r.getLastColumn();
    if (!["RATE_USD", "RATE_EUR", "RATE_COP"].some(name => overlaps(ss.getRangeByName(name)))) return;
    const stamp = ss.getRangeByName("RATE_UPDATED");
    if (stamp) stamp.setValue(new Date()).setNumberFormat("yyyy-MM-dd");
  } catch (err) {
    // never let a convenience trigger surface an error to the user
  }
}

/**
 * v1.1.19 (C2): runs `fn` only if no other run of this system is in
 * progress. Nothing used to prevent two runs at once (the 6 AM trigger and
 * a manual click, or a double click on Run now that the date-range dialog
 * closes immediately). Two runs both read the sheet before either wrote,
 * so both saved the same transactions, and two rebuildDerivedSheets()
 * interleaved their clear/write calls and corrupted Raw_<BANK> / Bank
 * Transfers. Apps Script releases the lock automatically when an execution
 * ends — including a hard timeout — so a crashed run can't leave it stuck.
 */
/**
 * v1.1.36: Apps Script kills a run at 6 minutes — mid-step, no finally, the progress toast left on screen and the
 * last steps never run. Reading emails (the slow part: one getPlainBody() per email) now stops RUN_READ_BUDGET_MS
 * after the action started, which leaves time to save, mark and run every cleanup step; the next run continues.
 */
const RUN_READ_BUDGET_MS = 3.5 * 60 * 1000;
// v1.1.43: no step starts after this — reported: a stopped run still ran every heavy step and hit the 6-minute limit
const RUN_STEPS_DEADLINE_MS = 5 * 60 * 1000;
// steps a stopped run leaves to the run that completes the range (the data is re-sorted, recategorized and refreshed there)
const DEFERRABLE_STEPS = ['sort', 'recategorize', 'investments'];
let RUN_STARTED_AT = 0;
let runClock = () => Date.now();   // replaceable in tests

function withRunLock(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    Logger.log("Another run is in progress — skipped.");
    safeAlert("⏳ Another run is already in progress. Wait for its summary before starting a new one.");
    return undefined;
  }
  RUN_STARTED_AT = runClock();   // v1.1.36
  try {
    return fn();
  } finally {
    if (PROGRESS_TOAST_OPEN) safeToast("Finished.", "📊 Financial Tracker", 3);   // v1.1.33
    lock.releaseLock();
  }
}

/**
 * v1.1.19 (M11): config for an entry point, or null after telling the user
 * why. A broken banksToTrack value used to surface as the misleading "Setup
 * not completed".
 */
function requireConfig() {
  const config = getConfig();
  if (!config) {
    safeAlert("❌ Setup not completed. Please run Setup Wizard first.");
    return null;
  }
  if (config.configError) {
    safeAlert("❌ " + config.configError);
    return null;
  }
  return config;
}

function openSetupWizard() {
  // v1.1.19 (M11): opens pre-filled with the current Configuration. Values are
  // passed as JSON ('<' escaped) and set from script, never pasted into the HTML.
  // v1.1.24: no personal values are written in this file any more — a first run
  // suggests the email of the Google account opening the wizard and leaves every
  // other field empty (placeholders are neutral). New: DOP salaries can have
  // ARS/SFS, AFP and ISR calculated automatically from the DR payroll rules
  // (see computeDrPayroll()), with a live preview; and the daily summary email
  // is configured here.
  const existing = getConfig();
  let defaultEmail = '';
  try { defaultEmail = Session.getActiveUser().getEmail() || ''; } catch (e) { defaultEmail = ''; }
  const prefill = existing ? {
    email: existing.email, incomeCurrency: existing.incomeCurrency, monthlyIncome: existing.monthlyIncome,
    deductionMode: existing.deductionMode, ARS: existing.deductions.ARS, AFP: existing.deductions.AFP,
    taxRate: existing.deductions.taxRate, ISR: existing.deductions.ISR,
    banksToTrack: existing.banksToTrack || {},
    notifyEnabled: existing.notifyEnabled, notifyMonthly: existing.notifyMonthly, notifyEmail: existing.notifyEmail,
    notifyHour: existing.notifyHour, notifySections: existing.notifySections,
    otherIncomes: existing.otherIncomes, otherDeductions: existing.otherDeductions, cards: existing.cards   // v1.1.59
  } : { email: defaultEmail };
  // v1.1.27: card catalogue for the "Credit Cards" section (public product facts only)
  // v1.1.60: one checkbox per bank the tracker reads, from the one list of banks
  const bankBoxesHtml = BANK_ORDER.map(b => '<div class="checkbox-item"><input type="checkbox" id="bank_' + b.toLowerCase() +
    '" checked onchange="renderCards()"><label for="bank_' + b.toLowerCase() + '">' + BANK_PATTERNS[b].name + '</label></div>').join('\n            ');
  const catalogJson = JSON.stringify(Object.keys(CARD_PRODUCTS).map(id =>
    ({ id: id, bank: CARD_PRODUCTS[id].bank, name: CARD_PRODUCTS[id].name }))).replace(/</g, '\\u003c');
  const prefillJson = JSON.stringify(prefill).replace(/</g, '\\u003c');
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        /* v1.1.59: other incomes / deductions as lines */
        .line { display: flex; gap: 6px; margin-bottom: 6px; align-items: center; }
        .line .lineLabel { flex: 2; } .line .lineAmount { flex: 1.2; } .line .lineCur { flex: 0 0 76px; }
        .lineRemove { flex: 0 0 34px; align-self: stretch; border: 1px solid #D1D5DB; background: #fff; border-radius: 6px; cursor: pointer; color: #6B7280; }
        .lineRemove:hover { color: #B91C1C; border-color: #B91C1C; }
        .lineAdd { background: none; border: none; color: #1D4ED8; cursor: pointer; padding: 2px 0; font-size: 13px; }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 20px;
        }
        .container { background: white; border-radius: 12px; box-shadow: 0 20px 60px rgba(0,0,0,0.3);
          max-width: 600px; width: 100%; padding: 36px; }
        .header { text-align: center; margin-bottom: 24px; }
        .header h1 { font-size: 26px; color: #333; margin-bottom: 8px; }
        .header p { color: #666; font-size: 14px; }
        .form-group { margin-bottom: 16px; }
        label { display: block; margin-bottom: 6px; color: #333; font-weight: 600; font-size: 14px; }
        input, select { width: 100%; padding: 11px; border: 2px solid #e0e0e0; border-radius: 8px;
          font-size: 14px; transition: border-color 0.3s; }
        input:focus, select:focus { outline: none; border-color: #667eea; }
        .row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
        .section-title { font-size: 16px; font-weight: 700; color: #333; margin-top: 22px; margin-bottom: 14px;
          border-bottom: 2px solid #667eea; padding-bottom: 8px; }
        .checkbox-group { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 12px; }
        .checkbox-item { display: flex; align-items: center; gap: 8px; }
        .checkbox-item input { width: auto; margin: 0; }
        .checkbox-item label { margin-bottom: 0; font-weight: 500; }
        .radio-group { display: flex; flex-direction: column; gap: 8px; margin-bottom: 14px; }
        .hint { color: #777; font-size: 12px; margin-top: 6px; line-height: 1.4; }
        .calc { width: 100%; border-collapse: collapse; font-size: 14px; margin-top: 4px; }
        .calc td { padding: 6px 4px; border-bottom: 1px solid #eee; }
        .calc td:last-child { text-align: right; font-variant-numeric: tabular-nums; }
        .box { background: #f6f7fb; border-radius: 8px; padding: 12px 14px; margin-bottom: 14px; }
        .buttons { display: flex; gap: 10px; margin-top: 26px; }
        button { flex: 1; padding: 14px; border: none; border-radius: 8px; font-size: 15px; font-weight: 600;
          cursor: pointer; transition: all 0.3s; }
        .btn-primary { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; }
        .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3); }
        .btn-secondary { background: #f0f0f0; color: #333; }
        .btn-secondary:hover { background: #e0e0e0; }
        .status { text-align: center; margin-top: 20px; padding: 15px; border-radius: 8px; display: none; }
        .status.success { background: #d4edda; color: #155724; display: block; }
        .status.error { background: #f8d7da; color: #721c24; display: block; }
        .status.info { background: #eaf1fe; color: #1d4ed8; display: block; }
        .card-row { display: grid; grid-template-columns: 84px 1fr 96px 96px; gap: 8px; align-items: center; margin-bottom: 8px; }
        .card-row .card-bank { font-weight: 700; font-size: 13px; color: #333; }
        .card-row .c-other { grid-column: 1 / -1; display: none; grid-template-columns: 84px 1fr 120px; gap: 8px; align-items: center; }
        .card-head { display: grid; grid-template-columns: 84px 1fr 96px 96px; gap: 8px; margin-bottom: 4px;
          font-size: 11px; font-weight: 600; color: #777; text-transform: uppercase; letter-spacing: .4px; }
        .c-other .c-label { font-size: 12px; color: #777; }
        .card-row input, .card-row select { padding: 9px; }
        button:disabled { opacity: 0.6; cursor: default; transform: none; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>💰 Financial Tracker Setup</h1>
          <p>Configure your transaction tracking system</p>
        </div>
        <form id="setupForm">
          <div class="form-group">
            <label>📧 Email Address</label>
            <input type="email" id="email" placeholder="you@example.com" required>
          </div>
          <div class="row">
            <div class="form-group">
              <label>💱 Income Currency</label>
              <select id="incomeCurrency" onchange="refresh()">
                <option value="USD">USD</option>
                <option value="DOP">DOP</option>
              </select>
            </div>
            <div class="form-group">
              <label>💵 Monthly Gross Salary (<span class="curLabel">USD</span>)</label>
              <input type="number" id="monthlyIncome" placeholder="0.00" step="0.01" min="0" required oninput="refresh()">
            </div>
          </div>

          <div class="form-group">
            <label>➕ Other Monthly Income</label>
            <div id="incomeLines"></div>
            <button type="button" class="lineAdd" onclick="addLine('income')">+ Add another income</button>
          </div>
          <div class="hint" style="margin:-6px 0 4px">Freelance work, rent you receive, etc. — each in its own currency. Added in full to your net income, no deductions.</div>

          <div class="section-title">📊 Monthly Deductions</div>
          <div id="dopModeGroup" class="radio-group">
            <div class="checkbox-item">
              <input type="radio" name="dedMode" id="dedAuto" value="auto" checked onchange="refresh()">
              <label for="dedAuto">Calculate ARS, AFP and ISR automatically (Dominican payroll rules)</label>
            </div>
            <div class="checkbox-item">
              <input type="radio" name="dedMode" id="dedManual" value="manual" onchange="refresh()">
              <label for="dedManual">Enter them myself</label>
            </div>
          </div>
          <div id="autoBox" class="box">
            <div id="autoPreview" class="hint">Enter your monthly gross salary to see the calculation.</div>
          </div>
          <div id="manualBox">
            <div class="row">
              <div class="form-group">
                <label>ARS (<span class="curLabel">USD</span> / month)</label>
                <input type="number" id="arsAmount" placeholder="0.00" step="0.01" min="0">
              </div>
              <div class="form-group">
                <label>AFP (<span class="curLabel">USD</span> / month)</label>
                <input type="number" id="afpAmount" placeholder="0.00" step="0.01" min="0">
              </div>
            </div>
            <div class="form-group" id="taxRateGroup">
              <label>Income Tax Rate (%)</label>
              <input type="number" id="taxRate" placeholder="0" step="0.01" min="0" max="100">
            </div>
            <div class="form-group" id="isrGroup">
              <label>ISR — Impuesto Sobre la Renta (RD$ / month)</label>
              <input type="number" id="isrAmount" placeholder="0.00" step="0.01" min="0">
            </div>
          </div>
          <div class="form-group" style="margin-top:10px">
            <label>➖ Other Monthly Deductions</label>
            <div id="deductionLines"></div>
            <button type="button" class="lineAdd" onclick="addLine('deduction')">+ Add a deduction</button>
          </div>
          <div class="hint" style="margin:-6px 0 4px">Anything else taken from your pay — a loan installment, a cooperative, insurance — each in its own currency. Subtracted from your net income.</div>

          <div class="section-title">🏦 Banks to Track</div>
          <div class="checkbox-group">
            ${bankBoxesHtml}
          </div>

          <div class="section-title">💳 Credit Cards</div>
          <div class="hint" style="margin:-4px 0 10px">The credit card you have with each bank (if any), and the day its statement
            closes and the day payment is due. They feed the Dashboard's card table and the cashback tips.</div>
          <div id="cardsBox"></div>

          <div class="section-title">📬 Summary Emails</div>
          <div class="checkbox-item" style="margin-bottom:8px">
            <input type="checkbox" id="notifyEnabled" onchange="refresh()">
            <label for="notifyEnabled">Daily — a summary of the previous day, every morning</label>
          </div>
          <div class="checkbox-item" style="margin-bottom:12px">
            <input type="checkbox" id="notifyMonthly" onchange="refresh()">
            <label for="notifyMonthly">Monthly — the previous month, on the 1st</label>
          </div>
          <div id="notifyBox">
            <div class="row">
              <div class="form-group">
                <label>Send to</label>
                <input type="email" id="notifyEmail" placeholder="Same as the email above">
              </div>
              <div class="form-group">
                <label>Around</label>
                <select id="notifyHour">
                  <option value="7">7:00 AM</option>
                  <option value="8" selected>8:00 AM</option>
                  <option value="9">9:00 AM</option>
                  <option value="10">10:00 AM</option>
                </select>
              </div>
            </div>
            <div id="dailySections">
            <label>The daily email includes</label>
            <div class="checkbox-group">
              <div class="checkbox-item"><input type="checkbox" id="sec_totals" checked><label for="sec_totals">Yesterday's total &amp; top purchases</label></div>
              <div class="checkbox-item"><input type="checkbox" id="sec_vsAverage" checked><label for="sec_vsAverage">Comparison with a typical day and your usual month</label></div>
              <div class="checkbox-item"><input type="checkbox" id="sec_transfers" checked><label for="sec_transfers">Transfers to review</label></div>
              <div class="checkbox-item"><input type="checkbox" id="sec_recommendations" checked><label for="sec_recommendations">Recommendations &amp; feedback</label></div>
              <div class="checkbox-item"><input type="checkbox" id="sec_cashback" checked><label for="sec_cashback">Card &amp; cashback tips</label></div>
              <div class="checkbox-item"><input type="checkbox" id="sec_investments" checked><label for="sec_investments">Investments</label></div>
            </div>
            </div>
            <div class="hint">Sent from your own Gmail account after the morning update, with links to this spreadsheet.
              The monthly email covers the whole previous calendar month.</div>
          </div>

          <div class="buttons">
            <button type="submit" class="btn-primary" id="saveBtn">✅ Save Configuration</button>
            <button type="button" class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
          </div>
          <div id="status" class="status"></div>
        </form>
      </div>
      <script>
        const SECTIONS = ['totals', 'vsAverage', 'transfers', 'recommendations', 'cashback', 'investments'];
        function el(id) { return document.getElementById(id); }
        function show(id, on) { el(id).style.display = on ? '' : 'none'; }
        function num(id) { return parseFloat(el(id).value) || 0; }
        function mode() {
          if (el('incomeCurrency').value !== 'DOP') return 'manual';
          return el('dedManual').checked ? 'manual' : 'auto';
        }
        function esc(t) {
          return String(t).replace(/[&<>"]/g, function(c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
        }
        function money(n) {
          return 'RD$' + Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }
        let previewTimer = null;
        function refresh() {
          const cur = el('incomeCurrency').value;
          const dop = cur === 'DOP';
          document.querySelectorAll('.curLabel').forEach(function(x) { x.textContent = cur; });
          show('dopModeGroup', dop);
          show('autoBox', dop && mode() === 'auto');
          show('manualBox', mode() === 'manual');
          show('taxRateGroup', !dop);
          show('isrGroup', dop);
          show('notifyBox', el('notifyEnabled').checked || el('notifyMonthly').checked);
          show('dailySections', el('notifyEnabled').checked);
          if (dop && mode() === 'auto') { clearTimeout(previewTimer); previewTimer = setTimeout(preview, 350); }
        }
        function preview() {
          const gross = num('monthlyIncome');
          if (!(gross > 0)) { el('autoPreview').textContent = 'Enter your monthly gross salary to see the calculation.'; return; }
          google.script.run
            .withSuccessHandler(function(r) {
              const line = function(a, b) { return '<tr><td>' + a + '</td><td>' + b + '</td></tr>'; };
              el('autoPreview').innerHTML = '<table class="calc">' +
                line('ARS / SFS (3.04%)', money(r.sfs)) + line('AFP (2.87%)', money(r.afp)) +
                line('ISR (DGII scale ' + r.year + ')', money(r.isr)) +
                line('<b>Net monthly salary</b>', '<b>' + money(r.net) + '</b>') + '</table>' +
                '<div class="hint">' + (r.stale ? '⚠️ Using the ' + r.year + ' scale — update it when DGII publishes the new one. ' : '') +
                'Sources: ' + esc(r.sources) + '. Not included: dependents or other deductions.</div>';
            })
            .withFailureHandler(function(e) { el('autoPreview').textContent = '❌ ' + e; })
            .previewDrPayroll(gross);
        }

        const CATALOG = ${catalogJson};
        const BANKS = ${JSON.stringify(BANK_ORDER)};   // v1.1.60
        function cardRows() { return Array.prototype.slice.call(document.querySelectorAll('.card-row')); }
        function readCards() {
          return cardRows().map(function(r) {
            const q = function(c) { return r.querySelector(c); };
            return { bank: r.getAttribute('data-bank'), product: q('.c-product').value,
              closeDay: parseInt(q('.c-close').value, 10) || null, dueDay: parseInt(q('.c-due').value, 10) || null,
              name: q('.c-name').value.trim(), cashback: parseFloat(q('.c-cashback').value) || 0 };
          });
        }
        function applyCards(list) {
          (list || []).forEach(function(c) {
            const r = document.querySelector('.card-row[data-bank="' + c.bank + '"]');
            if (!r) return;
            const q = function(x) { return r.querySelector(x); };
            if (c.product) q('.c-product').value = c.product;
            if (c.closeDay) q('.c-close').value = c.closeDay;
            if (c.dueDay) q('.c-due').value = c.dueDay;
            if (c.name) q('.c-name').value = c.name;
            if (c.cashback) q('.c-cashback').value = c.cashback;
            refreshCardRow(q('.c-product'));
          });
        }
        function refreshCardRow(sel) {
          const r = sel.parentNode;
          const has = sel.value !== 'none';
          r.querySelector('.c-close').style.visibility = has ? 'visible' : 'hidden';
          r.querySelector('.c-due').style.visibility = has ? 'visible' : 'hidden';
          r.querySelector('.c-other').style.display = sel.value === 'other' ? 'grid' : 'none';
        }
        function renderCards() {
          const keep = readCards();
          let html = '';
          BANKS.forEach(function(b) {
            if (!el('bank_' + b.toLowerCase()).checked) return;
            const opts = ['<option value="none">No credit card</option>']
              .concat(CATALOG.filter(function(c) { return c.bank === b; })
                .map(function(c) { return '<option value="' + c.id + '">' + esc(c.name) + '</option>'; }))
              .concat(['<option value="other">Other card…</option>']).join('');
            html += '<div class="card-row" data-bank="' + b + '"><div class="card-bank">' + b + '</div>' +
              '<select class="c-product" onchange="refreshCardRow(this)">' + opts + '</select>' +
              '<input class="c-close" type="number" min="1" max="31" placeholder="Closes (day)" title="Statement closes (day of month)">' +
              '<input class="c-due" type="number" min="1" max="31" placeholder="Due (day)" title="Payment due (day of month)">' +
              '<div class="c-other"><span class="c-label">↳ name / %</span><input class="c-name" placeholder="Card name" title="Card name">' +
              '<input class="c-cashback" type="number" min="0" max="100" step="0.1" placeholder="% cashback" title="% cashback"></div></div>';
          });
          el('cardsBox').innerHTML = html
            ? '<div class="card-head"><div>Bank</div><div>Card</div><div>Closes (day)</div><div>Due (day)</div></div>' + html
            : '<div class="hint">Select a bank above to add its card.</div>';
          cardRows().forEach(function(r) { refreshCardRow(r.querySelector('.c-product')); });
          applyCards(keep);
        }

        // v1.1.59: other incomes / deductions as lines
        function addLine(kind, item) {
          const box = el(kind === 'income' ? 'incomeLines' : 'deductionLines');
          if (box.children.length >= 10) return;
          const row = document.createElement('div');
          row.className = 'line';
          row.innerHTML = '<input class="lineLabel" maxlength="40">' +
            '<input type="number" class="lineAmount" placeholder="0.00" step="0.01" min="0">' +
            '<select class="lineCur"><option value="DOP">DOP</option><option value="USD">USD</option><option value="EUR">EUR</option></select>' +
            '<button type="button" class="lineRemove" title="Remove" onclick="removeLine(this)">✕</button>';
          box.appendChild(row);
          row.querySelector('.lineLabel').placeholder = kind === 'income' ? 'Source (optional)' : 'What (optional)';
          if (item) {
            row.querySelector('.lineLabel').value = item.label || '';
            row.querySelector('.lineAmount').value = item.amount;
            row.querySelector('.lineCur').value = item.currency || 'DOP';
          }
        }
        function removeLine(btn) {
          const row = btn.parentNode, box = row.parentNode;
          box.removeChild(row);
          if (!box.children.length) addLine(box.id === 'incomeLines' ? 'income' : 'deduction');
        }
        function lines(kind) {
          const rows = el(kind === 'income' ? 'incomeLines' : 'deductionLines').children;
          return Array.prototype.map.call(rows, function (row) {
            return { label: row.querySelector('.lineLabel').value.trim(), amount: parseFloat(row.querySelector('.lineAmount').value),
              currency: row.querySelector('.lineCur').value };
          }).filter(function (x) { return x.amount > 0; });
        }
        const EXISTING = ${prefillJson};
        (function prefill() {
          const setVal = function(id, v) { if (v !== undefined && v !== null && v !== '') el(id).value = v; };
          setVal('email', EXISTING.email);
          setVal('incomeCurrency', EXISTING.incomeCurrency);
          setVal('monthlyIncome', EXISTING.monthlyIncome);
          (EXISTING.otherIncomes || []).forEach(function (x) { addLine('income', x); });
          (EXISTING.otherDeductions || []).forEach(function (x) { addLine('deduction', x); });
          if (!el('incomeLines').children.length) addLine('income');
          if (!el('deductionLines').children.length) addLine('deduction');
          if (EXISTING.deductionMode === 'manual') el('dedManual').checked = true;
          if (EXISTING.deductionMode !== 'auto' || EXISTING.incomeCurrency !== 'DOP') {
            setVal('arsAmount', EXISTING.ARS);
            setVal('afpAmount', EXISTING.AFP);
            setVal('taxRate', EXISTING.taxRate);
            setVal('isrAmount', EXISTING.ISR);
          }
          if (EXISTING.banksToTrack) {
            // v1.1.60: a bank added since the setup was saved starts unticked — saving doesn't turn on banks you don't use
            BANKS.forEach(function(b) { el('bank_' + b.toLowerCase()).checked = !!EXISTING.banksToTrack[b]; });
          }
          el('notifyEnabled').checked = !!EXISTING.notifyEnabled;
          el('notifyMonthly').checked = !!EXISTING.notifyMonthly;
          if (EXISTING.notifyEmail && EXISTING.notifyEmail !== EXISTING.email) setVal('notifyEmail', EXISTING.notifyEmail);
          setVal('notifyHour', EXISTING.notifyHour);
          if (EXISTING.notifySections) {
            SECTIONS.forEach(function(k) { if (k in EXISTING.notifySections) el('sec_' + k).checked = !!EXISTING.notifySections[k]; });
          }
          renderCards();
          applyCards(EXISTING.cards);
          refresh();
        })();

        el('setupForm').addEventListener('submit', function(e) {
          e.preventDefault();
          const statusEl = el('status');
          const fail = function(msg) { statusEl.className = 'status error'; statusEl.textContent = '❌ ' + msg; };
          const cur = el('incomeCurrency').value;
          const income = num('monthlyIncome');
          if (!(income > 0)) return fail('Monthly income must be a number greater than 0.');
          if (cur !== 'DOP' && (num('taxRate') < 0 || num('taxRate') > 100)) return fail('Tax rate must be between 0 and 100.');
          const sections = {};
          SECTIONS.forEach(function(k) { sections[k] = el('sec_' + k).checked; });
          if (el('notifyEnabled').checked && !SECTIONS.some(function(k) { return sections[k]; })) {
            return fail('Pick at least one thing to include in the daily summary.');
          }
          const cards = readCards().filter(function(c) { return c.product !== 'none'; });
          const badDay = cards.find(function(c) { return [c.closeDay, c.dueDay].some(function(d) { return d !== null && (d < 1 || d > 31); }); });
          if (badDay) return fail(badDay.bank + ': statement and payment days must be between 1 and 31.');
          const noName = cards.find(function(c) { return c.product === 'other' && !c.name; });
          if (noName) return fail(noName.bank + ': give the other card a name.');
          const config = {
            email: el('email').value.trim(),
            incomeCurrency: cur,
            monthlyIncome: income,
            otherIncomes: lines('income'),         // v1.1.59: as many as needed, each in its currency
            otherDeductions: lines('deduction'),
            cards: cards,
            deductionMode: mode(),
            deductions: {
              ARS: num('arsAmount'), AFP: num('afpAmount'),
              taxRate: cur === 'DOP' ? 0 : num('taxRate'),
              ISR: cur === 'DOP' ? num('isrAmount') : 0
            },
            banksToTrack: BANKS.reduce(function (o, b) { o[b] = el('bank_' + b.toLowerCase()).checked; return o; }, {}),
            notify: {
              enabled: el('notifyEnabled').checked, monthly: el('notifyMonthly').checked, email: el('notifyEmail').value.trim(),
              hour: parseInt(el('notifyHour').value, 10), sections: sections
            },
            timestamp: new Date().toISOString()
          };
          // v1.1.27: like the other long actions, the window closes right away and the
          // sheet shows the progress (toasts) and a summary when it finishes.
          el('saveBtn').disabled = true;
          statusEl.className = 'status info';
          statusEl.textContent = '⏳ Saving and rebuilding the Dashboard in the background — follow the progress at ' +
            'the bottom-right of the sheet; a summary pops up when it finishes. This window will close.';
          google.script.run
            .withFailureHandler(function(error) { fail('Error: ' + error); el('saveBtn').disabled = false; })
            .saveSetupConfig(config);
          setTimeout(function() { google.script.host.close(); }, 1500);
        });
      </script>
    </body>
    </html>
  `);
  html.setWidth(620).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, "Setup Wizard");
}

/**
 * v1.1.24: Dominican payroll deductions for a DOP salary — the employee's
 * SFS (shown as ARS) and AFP contributions, each capped at its maximum
 * contributable salary, and ISR withheld per the DGII scale on the salary
 * net of those contributions (annualized, scale applied, divided by 12).
 * Parameters live per year, each from a published source — update this table
 * when TSS / DGII publish new values. Ley 30-26 changes the ISR scale from
 * 2027 (exempt up to RD$480,000, a new 27% top bracket) but DGII hadn't
 * published the full table yet (Sept 2026): until a 2027 entry is added, a
 * later year uses the latest one available and comes back flagged `stale`.
 * Not covered: dependents, deductible education expenses, other withholdings.
 */
const DR_PAYROLL_PARAMS = {
  2026: {
    sources: 'TSS Res. 01-2025 (caps from Feb 2026, minimum wage for caps RD$23,223) · DGII Res. DDG-AR1-2026-00001',
    sfsRate: 0.0304, afpRate: 0.0287,
    sfsCapMonthly: 232230,        // 10 minimum wages
    afpCapMonthly: 464460,        // 20 minimum wages
    isrAnnualScale: [             // annual taxable income; tax = base + (income - over) * rate
      { over: 0, base: 0, rate: 0 },
      { over: 416220.01, base: 0, rate: 0.15 },
      { over: 624329.01, base: 31216, rate: 0.20 },
      { over: 867123.01, base: 79776, rate: 0.25 }
    ]
  }
};

function drPayrollParams(year) {
  const years = Object.keys(DR_PAYROLL_PARAMS).map(Number).sort((a, b) => a - b);
  const usable = years.filter(y => y <= year);
  const chosen = usable.length ? usable[usable.length - 1] : years[0];
  return Object.assign({ year: chosen, stale: year > years[years.length - 1] }, DR_PAYROLL_PARAMS[chosen]);
}

/** Monthly amounts for a gross monthly DOP salary. Pure — see tests/. */
function computeDrPayroll(grossMonthly, year) {
  const p = drPayrollParams(year || new Date().getFullYear());
  const gross = Number(grossMonthly) || 0;
  const r2 = n => Math.round(n * 100) / 100;
  const sfs = r2(Math.min(gross, p.sfsCapMonthly) * p.sfsRate);
  const afp = r2(Math.min(gross, p.afpCapMonthly) * p.afpRate);
  const annualTaxable = (gross - sfs - afp) * 12;
  let bracket = p.isrAnnualScale[0];
  p.isrAnnualScale.forEach(b => { if (annualTaxable >= b.over) bracket = b; });
  const isrAnnual = bracket.rate ? bracket.base + (annualTaxable - bracket.over) * bracket.rate : 0;
  const isr = r2(Math.max(0, isrAnnual) / 12);
  return { gross: gross, sfs: sfs, afp: afp, isr: isr, net: r2(gross - sfs - afp - isr),
           year: p.year, stale: p.stale, sources: p.sources };
}

/** Called by the Setup Wizard for its live preview. */
function previewDrPayroll(grossMonthly) {
  return computeDrPayroll(grossMonthly);
}

/** Writes key → value pairs into Configuration (by key; appends a missing key). */
function setConfigValues(values) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  if (!sheet) return;
  const keys = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues().map(r => String(r[0]).trim());
  Object.keys(values).forEach(key => {
    const row = keys.indexOf(key) + 1;
    if (row > 0) sheet.getRange(row, 2).setValue(values[key]);
    else { sheet.appendRow([key, values[key]]); keys.push(key); }
  });
}

/**
 * v1.1.24: in automatic mode the stored ARS / AFP / ISR are recomputed from
 * the current parameters (called when the Dashboard is built and before each
 * daily summary), so a parameter update in code reaches the numbers without
 * re-running the wizard. Returns the calculation, or null when not in auto mode.
 */
function refreshAutoDeductions() {
  const config = getConfig();
  if (!config || config.incomeCurrency !== 'DOP' || config.deductionMode !== 'auto') return null;
  const calc = computeDrPayroll(config.monthlyIncome);
  setConfigValues({ ARS: calc.sfs, AFP: calc.afp, ISR: calc.isr, taxRate: 0 });
  return calc;
}

const EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

/**
 * v1.1.19 (M11): server-side check of what the wizard sends — the browser
 * checks are a convenience, this is the guard.
 */
function validateSetupInput(config) {
  // v1.1.24: deduction mode (auto only for DOP), ISR amount, daily summary options
  if (!config || typeof config !== 'object') return "No configuration received.";
  if (!EMAIL_RE.test(String(config.email || '').trim())) return "Enter a valid email address.";
  if (['USD', 'DOP'].indexOf(config.incomeCurrency) === -1) return "Income currency must be USD or DOP.";
  if (!(Number(config.monthlyIncome) > 0)) return "Monthly income must be a number greater than 0.";
  const mode = config.deductionMode || 'manual';
  if (['auto', 'manual'].indexOf(mode) === -1) return "Deduction mode must be automatic or manual.";
  if (mode === 'auto' && config.incomeCurrency !== 'DOP') return "Automatic deductions are only available for a DOP salary.";
  const d = config.deductions || {};
  if (Number(d.ARS) < 0 || Number(d.AFP) < 0 || Number(d.ISR) < 0) return "Deductions can't be negative.";
  const tax = Number(d.taxRate || 0);
  if (!(tax >= 0 && tax <= 100)) return "Tax rate must be between 0 and 100.";
  if (!config.banksToTrack || typeof config.banksToTrack !== 'object') return "Select at least one bank.";
  if (!Object.keys(config.banksToTrack).some(b => config.banksToTrack[b])) return "Select at least one bank.";
  // v1.1.27: other income and cards
  if (Number(config.otherIncome || 0) < 0) return "Other income can't be negative.";   // the older single field
  // v1.1.59: other incomes and other deductions — lists, each line in its own currency
  for (const [list, what] of [[config.otherIncomes, 'Other income'], [config.otherDeductions, 'Other deduction']]) {
    if (list === undefined || list === null) continue;
    if (!Array.isArray(list)) return what + 's must be a list.';
    if (list.length > MONEY_LINES_MAX) return 'Up to ' + MONEY_LINES_MAX + ' ' + what.toLowerCase() + 's.';
    for (const x of list) {
      if (!x || !(Number(x.amount) >= 0) || !isFinite(Number(x.amount))) return what + ' amounts must be numbers of 0 or more.';
      if (OTHER_MONEY_CURRENCIES.indexOf(x.currency) === -1) return what + ' currency must be ' + OTHER_MONEY_CURRENCIES.join(', ') + '.';
    }
  };
  const cards = config.cards || [];
  if (!Array.isArray(cards)) return "Cards must be a list.";
  for (const c of cards) {
    if (!c || c.product === 'none') continue;          // "No credit card" — nothing to check
    if (c.product !== 'other' && !CARD_PRODUCTS[c.product]) return c.bank + ": unknown card.";
    if (c.product === 'other' && !String(c.name || '').trim()) return c.bank + ": give the other card a name.";
    for (const d of [c.closeDay, c.dueDay]) {
      if (d !== null && d !== undefined && d !== '' && !(Number(d) >= 1 && Number(d) <= 31)) {
        return c.bank + ": statement and payment days must be between 1 and 31.";
      }
    }
  }
  const n = config.notify || {};
  if (n.enabled || n.monthly) {
    if (n.email && !EMAIL_RE.test(String(n.email).trim())) {
      return "Enter a valid email for the summaries, or leave it empty to use the main one.";
    }
    const h = Number(n.hour);
    if (!(h >= 0 && h <= 23 && Math.floor(h) === h)) return "Pick an hour for the summary emails.";
  }
  if (n.enabled && (!n.sections || !Object.keys(n.sections).some(k => n.sections[k]))) {
    return "Pick at least one thing to include in the daily summary.";
  }
  return null;
}

function saveSetupConfig(config) {
  // v1.1.27: the wizard window closes right away (like Monitor by Date Range), so
  // everything the user needs to know is reported here: toasts while it works, and
  // a summary alert (or the error) at the end. Runs under the run lock and rebuilds
  // the Dashboard so new income, deductions and cards show at once.
  const problem = validateSetupInput(config);
  if (problem) {
    safeAlert("❌ Setup not saved: " + problem);
    return false;
  }
  return withRunLock(() => {
    try {
      safeToast("Saving your configuration...", "📊 Financial Tracker", -1);
      getOrCreateSheet(CONFIG_SHEET);
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(CONFIG_SHEET);
      sheet.clear();

      if (!Array.isArray(config.otherIncomes) && Number(config.otherIncome) > 0) {   // v1.1.59: the older single field
        config.otherIncomes = [{ label: '', amount: Number(config.otherIncome), currency: config.otherIncomeCurrency === 'USD' ? 'USD' : 'DOP' }];
      }
      const dop = config.incomeCurrency === 'DOP';
      const auto = dop && config.deductionMode === 'auto';
      const calc = auto ? computeDrPayroll(config.monthlyIncome) : null;
      const d = config.deductions || {};
      const n = config.notify || {};
      const rows = [
        ["Key", "Value"],
        ["email", String(config.email).trim()],
        ["incomeCurrency", config.incomeCurrency],
        ["monthlyIncome", Number(config.monthlyIncome)],
        // v1.1.59: the lines, and their totals per currency (what the Dashboard converts at its live rates)
        ["otherIncomes", JSON.stringify(normalizeMoneyLines(config.otherIncomes))],
        ["otherDeductions", JSON.stringify(normalizeMoneyLines(config.otherDeductions))],
      ].concat(moneyTotalsRows(config)).concat([
        ["deductionMode", auto ? 'auto' : 'manual'],
        ["ARS", auto ? calc.sfs : Number(d.ARS) || 0],
        ["AFP", auto ? calc.afp : Number(d.AFP) || 0],
        ["ISR", auto ? calc.isr : (dop ? Number(d.ISR) || 0 : 0)],
        ["taxRate", dop ? 0 : Number(d.taxRate) || 0],
        ["banksToTrack", JSON.stringify(config.banksToTrack)],
        ["cards", JSON.stringify((config.cards || []).filter(c => c && c.product && c.product !== 'none').map(c => ({ bank: c.bank, product: c.product,
          closeDay: Number(c.closeDay) || null, dueDay: Number(c.dueDay) || null,
          name: c.product === 'other' ? String(c.name || '').trim() : '', cashback: c.product === 'other' ? Number(c.cashback) || 0 : null })))],
        ["setupDate", config.timestamp || new Date().toISOString()],
        ["notifyEnabled", !!n.enabled],
        ["notifyMonthly", !!n.monthly],
        ["notifyEmail", String(n.email || '').trim()],
        ["notifyHour", Number(n.hour) || 8],
        ["notifySections", JSON.stringify(Object.assign({}, SUMMARY_SECTIONS_DEFAULT, n.sections || {}))]
      ]);
      sheet.getRange(1, 1, rows.length, 2).setValues(rows);
      ensureConfigNamedRanges();
      Logger.log("✅ Config saved successfully");

      safeToast("Preparing sheets and daily triggers...", "📊 Financial Tracker", -1);
      getOrCreateSheet(TRANSACTIONS_SHEET);
      initializeTransactionsSheet();
      // Custom Rules starts with one example row: Category "Exclude" leaves a
      // matching transfer (e.g. between your own accounts) out of every total.
      const customRulesSheet = getOrCreateSheet(CUSTOM_RULES_SHEET);
      if (customRulesSheet.getLastRow() < 1) {
        customRulesSheet.appendRow(["UserEmail", "Category", "Keyword", "Timestamp"]);
        customRulesSheet.appendRow(["(example — delete this row)", "Exclude", "YOUR OWN NAME HERE",
          "Category=\"Exclude\" leaves a matching Transfer out of every Dashboard total (self-transfers)"]);
      }
      createTrigger();
      ensureSheetOrder();

      safeToast("Rebuilding the Dashboard...", "📊 Financial Tracker", -1);
      buildOrRefreshDashboard();          // also lands on the Dashboard
      safeToast("Done.", "📊 Financial Tracker", 3);
      safeAlert(setupSavedMessage(getConfig(), ss));
      return true;
    } catch (error) {
      Logger.log("❌ Error saving config: " + error);
      safeAlert("❌ Setup could not finish: " + error + "\nWhatever was saved before the error is kept — run the Setup Wizard again.");
      return false;
    }
  });
}

/** v1.1.27: what the user sees after saving the Setup Wizard. */
function setupSavedMessage(config, ss) {
  const rates = readDashboardRates(ss);
  const lines = ["✅ Configuration saved and Dashboard updated.", "",
    "💵 Net income: " + summaryMoney(computeNetIncomeDop(config, rates)) + " / month (DOP-equivalent)" +
      ((config.otherIncomes || []).length ? ", with " + config.otherIncomes.length + " other income(s)" : "") +
      ((config.otherDeductions || []).length ? " and " + config.otherDeductions.length + " other deduction(s)" : ""),
    "🔄 Bank emails are read every morning around 6 AM."];
  if (config.notifyEnabled) lines.push("📬 Daily summary around " + config.notifyHour + ":00.");
  if (config.notifyMonthly) lines.push("🗓️ Monthly summary on the 1st around " + config.notifyHour + ":00.");
  const cards = resolveCards(config.cards);
  lines.push("💳 Credit cards: " + (cards.length ? cards.map(c => c.bank + (c.name ? ' ' + c.name : '')).join(', ') : 'none set'));
  lines.push("", "Next: 📊 Tracker › 📘 Start here — a checklist of what's left, checked against your sheet.");   // v1.1.47
  return lines.join("\n");
}

/**
 * v1.0.3: Added the "Type" column (Transaction / Transfer / Card Payment / Cashback).
 * v1.1.2: Moved "Currency" from the end to right after "Amount" (was column L,
 * now column E) per request — see TX_COL in 04_sheetsWriter.gs for the full
 * layout every reader function now uses. If you deployed before v1.1.2, your
 * sheet's headers won't match new rows until you delete and let this
 * function recreate it (you'll lose logged rows), or manually reorder the
 * header row to: Date, Bank, Merchant, Amount, Currency, Category,
 * Description, Email Subject, Timestamp, IsCredit, IsCashback, Type.
 * v1.1.4: Type values are now English ('Transaction'/'Transfer'/etc,
 * were Spanish) — see 03_gmailMonitor.gs TYPE_KEYWORDS for why that mattered
 * beyond naming (a real bug: bare 'CASHBACK' matched inside "SUPERCASHBACK",
 * a real Banesco card product name, mistyping every purchase on that card).
 */

function initializeTransactionsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSACTIONS_SHEET);
  
  if (sheet.getLastRow() < 2) {
    const headers = [
      "Date",
      "Bank",
      "Merchant",
      "Amount",
      "Currency",
      "Category",
      "Description",
      "Email Subject",
      "Timestamp",
      "IsCredit",
      "IsCashback",
      "Type",
      "MessageId",
      "TxRef",     // v1.1.23: bank's own transaction id, hidden
      "Auto Category"   // v1.1.39: the category the tracker last set, hidden — a different one was set by hand
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    // v1.1.5: Timestamp (column I) hidden per request — "no es necesario
    // mostrar el Timestamp". Kept in the data model (still written on every
    // save) in case it's useful later; just not shown by default. Unhide via
    // Format > Hidden sheets/columns if you ever want it back.
    sheet.hideColumns(TX_COL.TIMESTAMP + 1);
    // v1.1.6: IsCashback (column K) hidden too, same reasoning — "no
    // necesito mostrar si es cashback o no."
    sheet.hideColumns(TX_COL.IS_CASHBACK + 1);
    // v1.1.13: MessageId (column M) hidden — internal-only, used by
    // isDuplicate() to tell apart two real transactions that happen to
    // share the same date/bank/amount. Not meant to be read by a person.
    sheet.hideColumns(TX_COL.MESSAGE_ID + 1);
    sheet.hideColumns(TX_COL.TX_REF + 1);
    sheet.hideColumns(TX_COL.AUTO_CATEGORY + 1);   // v1.1.39
  }
}

/**
 * v1.1.12: BUG FIX — the 6:00 AM daily trigger has been crashing on every
 * run since it was introduced. A time-triggered execution has no UI session
 * at all (no one is looking at the sheet), so SpreadsheetApp.getUi() throws
 * immediately — "Cannot call SpreadsheetApp.getUi() from this context."
 * Worse, the original error-handling in runGmailMonitor() tried to report
 * that same failure via ANOTHER getUi().alert() call inside the catch
 * block, which threw too, uncaught — that second crash is what actually
 * showed up in Executions. The real work (search/parse/save) likely
 * completed fine before hitting the first alert; only the user-facing
 * notification (meaningless with no one watching anyway) crashed.
 * safeAlert()/safeToast() degrade to Logger.log() when there's no UI,
 * instead of throwing — used everywhere in the monitor pipeline the
 * trigger can reach, since a menu click always has a UI and doesn't need
 * the fallback, but a triggered run never does.
 */
function safeAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log("[No UI available — trigger-based run] " + message);
  }
}

// v1.1.33: a progress toast shown with no timeout (-1) stays on screen until another toast replaces it —
// "Refreshing investments..." did, and so did every progress toast on an error path. withRunLock() now
// closes whichever is still open when its action ends.
let PROGRESS_TOAST_OPEN = false;
function safeToast(message, title, timeout) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeout);
    PROGRESS_TOAST_OPEN = timeout === -1;
  } catch (e) {
    Logger.log("[Toast unavailable] " + title + ": " + message);
  }
}
/**
 * Core pipeline shared by "Monitor Gmail Now", the 6 AM trigger and
 * "Monitor by Date Range": parse → save → mark processed, then (always)
 * sort → recategorize (+ rebuild derived sheets) → sheet order → Dashboard.
 * History: v1.1.12 auto-recategorize at the end and UI-safe alerts;
 * v1.1.13 try/finally so cleanup still runs after a catchable error (a hard
 * timeout still can't be caught — keep each run's range small).
 * v1.1.19:
 *  - takes the search result { threads, range, capped } so messages outside
 *    the range are skipped (M6) and a capped search is reported;
 *  - each cleanup step in `finally` has its own try/catch (M10) — an error
 *    in recategorize used to escape the finally block with no alert at all;
 *  - threads that failed to parse stay unread (M4);
 *  - the summary reports every filter/failure counter (E7).
 */
function runGmailMonitorCore(search, config) {
  const threads = search.threads || [];
  Logger.log("Found " + threads.length + " threads to process");
  safeToast("Parsing " + threads.length + " email thread(s)... this can take a minute for a large batch.",
    "📊 Financial Tracker", -1);

  const rawCustomRules = getUserCustomRules(config.email);
  let transactions = [];
  let stats = newParseStats();
  let results = { success: 0, failed: 0, duplicates: 0 };
  let marked = null;
  let recatChanged = 0;
  let investments = null;   // v1.1.29
  let unrecognizedOpen = null;   // v1.1.35
  let stopped = null;            // v1.1.36
  let rides = null;              // v1.1.67
  const deferred = [];           // v1.1.43
  const errors = [];

  try {
    // v1.1.36: emails already saved are not read again, and reading stops in time (see RUN_READ_BUDGET_MS)
    const extraction = extractTransactionsFromThreads(threads, rawCustomRules, search.range, {
      skipIds: alreadyReadIds(), deadline: (RUN_STARTED_AT || runClock()) + RUN_READ_BUDGET_MS, clock: runClock });
    transactions = extraction.transactions;
    stats = extraction.stats;
    stopped = extraction.stopped;
    Logger.log("Parsed " + transactions.length + " transactions");
    safeToast("Saving " + transactions.length + " transaction(s)...", "📊 Financial Tracker", -1);
    const timed = (label, fn) => { const t0 = Date.now(); const out = fn(); Logger.log('⏱ ' + label + ': ' + Math.round((Date.now() - t0) / 1000) + ' s'); return out; };
    results = timed('save', () => saveTransactions(transactions));
    marked = timed('mark', () => markEmailsAsProcessed(extraction.processedThreads, extraction.failedThreadIds,   // unread ones stay for the next run
      { query: search.query, deadline: (RUN_STARTED_AT || runClock()) + RUN_STEPS_DEADLINE_MS, clock: runClock }));
    // v1.1.43: statements, promotions, declined… aren't read again. v1.1.49: ONLY those — saved emails are recognized by
    // Transactions itself, so a row you delete is read again (logging them too meant it never came back)
    timed('read log', () => logReadEmails(stats.filteredIds));
  } catch (error) {
    errors.push("parse/save — " + error);
    Logger.log("❌ Error during parse/save: " + error);
  } finally {
    const steps = [
      ["sort", () => sortAllDataSheets()],
      // v1.1.67: Uber ride alerts checked against Uber's trip receipts — before Recategorize rebuilds the sheets
      ["uber receipts", () => { rides = reconcileRideReceipts(new Date(), search.range ? search.range.start : null); }],
      ["recategorize", () => {
        safeToast("Recategorizing...", "📊 Financial Tracker", -1);
        recatChanged = recategorizeAllTransactions(config.email);
        Logger.log("Auto-recategorize: " + recatChanged + " cell(s) updated");
      }],
      // v1.1.29: broker emails → Investment Ledger, bank deposits → ledger, Holdings rebuilt.
      // Isolated like every step here: an investments problem never affects the bank side.
      ["investments", () => {
        safeToast("Updating investments...", "📊 Financial Tracker", -1);
        investments = runInvestmentsStep(search.range);
        styleInvestmentSheets();   // v1.1.43: only these — the full styling already ran with recategorize
      }],
      // v1.1.35: emails that couldn't be read (bank and broker) → Unrecognized; ones read cleanly now drop off it
      ["unrecognized", () => {
        unrecognizedOpen = recordUnrecognized(stats.unrecognized.concat(investments ? investments.unrecognized : []),
          stats.readIds.concat(investments ? investments.readIds : []), new Date());
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNRECOGNIZED_SHEET);
        if (sheet) styleUnrecognizedSheet(sheet);
      }],
      ["sheet order", () => ensureSheetOrder()],
      ["open Dashboard", () => {
        const dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
        if (dashboardSheet) SpreadsheetApp.setActiveSheet(dashboardSheet);
      }]
    ];
    steps.forEach(([name, fn]) => {
      // v1.1.43: a stopped run leaves the heavy steps to the run that completes the range; and no step starts late
      if (stopped && DEFERRABLE_STEPS.indexOf(name) !== -1) { deferred.push(name); return; }
      if (runClock() - (RUN_STARTED_AT || runClock()) > RUN_STEPS_DEADLINE_MS) { deferred.push(name); return; }
      const started = Date.now();
      try { fn(); Logger.log('⏱ ' + name + ': ' + Math.round((Date.now() - started) / 1000) + ' s'); } catch (error) {   // v1.1.44: timings
        errors.push(name + " — " + error);
        Logger.log("❌ Error during " + name + ": " + error);
      }
    });
  }

  // v1.1.26: the summary emails report whether the morning update ran and how it went
  recordLastRun({ at: new Date().toISOString(), saved: results.success, duplicates: results.duplicates,
    failed: results.failed, unparsed: stats.amountNotFound + stats.parseErrors + (investments ? investments.unparsed : 0),
    errors: errors.length, unrecognized: unrecognizedOpen || 0, partial: !!stopped });
  safeToast("Done.", "📊 Financial Tracker", 3);
  safeAlert(buildRunSummary({ search: search, threads: threads, transactions: transactions, stats: stats,
                              results: results, marked: marked, recatChanged: recatChanged, errors: errors,
                              investments: investments, unrecognizedOpen: unrecognizedOpen, stopped: stopped, deferred: deferred,
                              rides: rides }));
}

/** v1.1.19 (E7): run summary text — pure, so tests/ can check it. */
function buildRunSummary(r) {
  const s = r.stats;
  const lines = [];
  if (r.stopped) {   // v1.1.36
    lines.push("⏸ Stopped reading early to stay within Google's 6-minute limit — " + r.stopped.remaining +
      " of " + r.threads.length + " thread(s) left. Run it again (same range) to continue; emails already read are skipped.", "");
  }
  if (r.deferred && r.deferred.length) {   // v1.1.43
    lines.push("⏭ Left for the next run: " + r.deferred.join(", ") + " — run it again (same range) until this line is gone.", "");
  }
  lines.push(
    "✅ Email threads found: " + r.threads.length + " (" + s.messagesSeen + " message(s) in range)",
    "📄 Transactions parsed: " + r.transactions.length,
    "💾 Saved: " + r.results.success + " | Duplicates: " + r.results.duplicates + " | Failed: " + r.results.failed,
    "🔁 Recategorized: " + r.recatChanged + " cell(s)"
  );
  if (s.alreadySaved) lines.push("⏭ Read before, skipped: " + s.alreadySaved + " email(s)");   // v1.1.36; v1.1.43: saved or read-only
  const skipped = [];
  if (s.promotional) skipped.push("promotional " + s.promotional);
  if (s.nonTransactional) skipped.push("non-transactional " + s.nonTransactional);
  if (s.declined) skipped.push("declined " + s.declined);
  if (s.notOwnBank) skipped.push("other senders " + s.notOwnBank);
  if (skipped.length) lines.push("🧹 Skipped on purpose: " + skipped.join(" · "));
  const unparsed = s.amountNotFound + s.parseErrors;
  if (unparsed) {
    lines.push("⚠️ Could not parse: " + unparsed + " email(s) — left UNREAD in Gmail; details in View > Executions");
  }
  if (s.placeholders) lines.push("⚠️ Saved with placeholder merchant: " + s.placeholders);
  if (r.unrecognizedOpen) {   // v1.1.35
    lines.push("🔎 Unrecognized: " + r.unrecognizedOpen + " email(s) need a look — see the Unrecognized sheet");
  }
  const inv = r.investments;   // v1.1.29
  if (inv && (inv.parsed || inv.deposits || inv.unparsed || inv.saved)) {
    lines.push("📈 Investments: " + inv.saved + " new ledger row(s) (" + inv.parsed + " from broker emails, " +
      inv.deposits + " deposit(s) from bank transfers) · duplicates " + inv.duplicates);
    if (inv.unparsed) lines.push("⚠️ Broker emails not read: " + inv.unparsed + " — left UNREAD; details in View > Executions");
  }
  if (inv && inv.warnings && inv.warnings.length) lines.push("⚠️ Holdings: " + inv.warnings.join(" · "));
  const rd = r.rides;   // v1.1.67
  if (rd && (rd.holds || rd.adjusted || rd.charged || rd.unmatched)) {
    const parts = [rd.receipts + " trip receipt(s) checked"];
    if (rd.holds) parts.push(rd.holds + " authorization(s) never charged left out (" + rd.currency + " " + rd.holdsTotal.toFixed(2) + ")");
    if (rd.adjusted) parts.push(rd.adjusted + " amount(s) corrected");
    if (rd.unmatched) parts.push(rd.unmatched + " ride alert(s) without a receipt kept");
    lines.push("🚕 Uber: " + parts.join(" · "));
  }
  if (s.reversals) lines.push("↩️ Reversals: " + s.reversals + " — saved as negative rows that cancel the original purchase");
  if (r.results.reversalsUnmatched) {
    lines.push("⚠️ Reversal without its original purchase: " + r.results.reversalsUnmatched +
      " — look for \"" + REVERSAL_UNMATCHED + "\" in Transactions");
  }
  if (r.search && r.search.capped) {
    lines.push("⚠️ Hit the " + MAX_THREADS_PER_RUN + "-thread cap — split this date range into smaller pieces.");
  }
  if (r.search && r.search.untracked && r.search.untracked.length) {   // v1.1.61
    lines.push("⚠️ Emails found from banks NOT ticked in the Setup Wizard: " + r.search.untracked.join(", ") +
      ". They were not read: tick them in 🔧 Setup Wizard and run this range again.");
  }
  const byBank = {};   // v1.1.61: what each bank gave, so a bank that gives nothing is easy to spot
  (r.transactions || []).forEach(t => { byBank[t.bank] = (byBank[t.bank] || 0) + 1; });
  if (Object.keys(byBank).length) {
    lines.push("🏦 Parsed by bank: " + BANK_ORDER.filter(b => byBank[b]).map(b => b + " " + byBank[b]).join(" · "));
  }
  if (r.threads.length === 0) {
    lines.push("", "⚠️ 0 emails found — check that the selected banks match your real emails. " +
      "Run debugBankEmailSample() from the Apps Script editor (View > Logs) to see them.");
  }
  if (r.errors.length) {
    lines.push("", "❌ Errors (whatever was saved before them is still there):", ...r.errors.map(e => "  • " + e));
  }
  return lines.join("\n");
}

/**
 * v1.1.19 (C1): window for the daily run — from the 1st of YESTERDAY's
 * month through today, both included. On the 1st of a month that still
 * covers the whole previous month, so its last day (and anything that
 * arrived after the 6 AM run on that day) can't fall through the gap
 * between months. Re-reading already-saved emails is safe (dedup).
 */
function computeDailyWindow(today) {
  const tz = Session.getScriptTimeZone();
  const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1, 12);
  return {
    startStr: Utilities.formatDate(new Date(yesterday.getFullYear(), yesterday.getMonth(), 1, 12), tz, 'yyyy-MM-dd'),
    endStr: Utilities.formatDate(today, tz, 'yyyy-MM-dd')
  };
}

/**
 * "🔄 Monitor Gmail Now" and the 6 AM trigger.
 * v1.1.12: narrowed to the current month (a backlog is what the date-range
 * run is for). v1.1.19: window from computeDailyWindow() (C1) and the run
 * lock (C2).
 */
function runGmailMonitor() {
  withRunLock(() => {
    try {
      const config = requireConfig();
      if (!config) return;
      Logger.log("🔄 Starting Gmail Monitor... (" + SCRIPT_VERSION + ")");
      Logger.log("Tracking banks: " + JSON.stringify(config.banksToTrack));
      safeToast("Searching Gmail...", "📊 Financial Tracker", -1);
      const win = computeDailyWindow(new Date());
      const search = searchTransactionEmailsByDateRange(win.startStr, win.endStr, config.banksToTrack);
      runGmailMonitorCore(search, config);
    } catch (error) {
      Logger.log("❌ Error in Gmail Monitor: " + error);
      safeAlert("❌ Error: " + error.toString());
    }
  });
}


/**
 * v1.1.2: prompts for a start/end date and runs the same pipeline scoped to
 * that window (via searchTransactionEmailsByDateRange, 03_gmailMonitor.gs)
 * instead of the rolling `daysBack` default. Useful for a backlog bigger
 * than the search's per-run cap (500 threads, paginated) — split the whole
 * year into a few date-range runs (e.g. one per quarter or month) instead of
 * one huge run. isDuplicate() (04_sheetsWriter.gs) still guards every save,
 * so overlapping ranges across runs are always safe — nothing gets
 * double-counted.
 * v1.1.3: replaces the old two-sequential-ui.prompt() version with a proper
 * HTML dialog — both dates are entered together, each as a native
 * <input type="date"> (renders as the browser's own calendar picker; no
 * custom calendar widget needed). Pre-fills with last calendar month as a
 * convenient default (covers the common "catch up on last month" case in
 * one click) while staying fully editable for any custom range — e.g. one
 * quarter at a time to work through a large backlog without hitting Apps
 * Script's 6-minute execution ceiling.
 */
function openDateRangeDialog() {
  const config = getConfig();
  if (!config) {
    SpreadsheetApp.getUi().alert("❌ Setup not completed. Please run Setup Wizard first.");
    return;
  }

  const today = new Date();
  const firstOfThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const lastMonthEnd = new Date(firstOfThisMonth.getTime() - 1);
  const lastMonthStart = new Date(lastMonthEnd.getFullYear(), lastMonthEnd.getMonth(), 1);
  const fmt = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 24px; }
        h2 { font-size: 16px; margin-bottom: 16px; }
        label { display: block; font-size: 13px; font-weight: 600; margin: 12px 0 4px; }
        input[type="date"] { width: 100%; padding: 8px; font-size: 14px; border: 1px solid #ccc; border-radius: 6px; }
        .hint { font-size: 11px; color: #888; margin-top: 4px; }
        .buttons { margin-top: 20px; display: flex; gap: 8px; justify-content: flex-end; }
        button { padding: 8px 16px; border-radius: 6px; border: none; font-size: 13px; cursor: pointer; }
        .btn-primary { background: #1F3864; color: white; }
        .btn-secondary { background: #eee; color: #333; }
        .status { margin-top: 12px; font-size: 13px; }
        .status.error { color: #c00; }
      </style>
    </head>
    <body>
      <h2>📅 Monitor by Date Range</h2>
      <label for="startDate">Start date</label>
      <input type="date" id="startDate" value="${fmt(lastMonthStart)}">
      <label for="endDate">End date</label>
      <input type="date" id="endDate" value="${fmt(lastMonthEnd)}">
      <div class="hint">Defaults to all of last month. Change it to any range — for example, one quarter at a time to work through a full year of emails without hitting the 6-minute execution limit.</div>
      <div id="status" class="status"></div>
      <div class="buttons">
        <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
        <button class="btn-primary" id="runBtn" onclick="run()">Run</button>
      </div>
      <script>
        function run() {
          const startDate = document.getElementById('startDate').value;
          const endDate = document.getElementById('endDate').value;
          const status = document.getElementById('status');
          if (!startDate || !endDate) {
            status.className = 'status error';
            status.textContent = 'Select both dates.';
            return;
          }
          // v1.1.19 (M10): 'yyyy-mm-dd' strings compare correctly as text.
          if (startDate > endDate) {
            status.className = 'status error';
            status.textContent = 'The start date must be on or before the end date.';
            return;
          }
          document.getElementById('runBtn').disabled = true;
          status.className = 'status';
          status.textContent = '⏳ Processing in the background — progress shows at the bottom-right of the ' +
            'sheet and a summary pops up when it finishes. This window will close.';
          // v1.1.18 closed the window immediately, per request. v1.1.19
          // (M10): the close now waits 1.5 s so the server call is surely
          // dispatched first, and the failure handler is back for anything
          // that fails fast. Errors after the window closes are shown by
          // the server itself (safeAlert).
          google.script.run
            .withFailureHandler(function(error) {
              status.className = 'status error';
              status.textContent = '❌ Error: ' + error;
              document.getElementById('runBtn').disabled = false;
            })
            .runGmailMonitorForDateRange(startDate, endDate);
          setTimeout(function() { google.script.host.close(); }, 1500);
        }
      </script>
    </body>
    </html>
  `);
  html.setWidth(420).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, "Monitor by Date Range");
}
/**
 * Server side of the date-range dialog. v1.1.19: run lock (C2), end date
 * actually included (C1), and every error reported (M10) — the dialog may
 * already be closed, so the server shows it.
 */
function runGmailMonitorForDateRange(startDate, endDate) {
  withRunLock(() => {
    try {
      const config = requireConfig();
      if (!config) return;
      Logger.log("🔄 Starting Gmail Monitor for range " + startDate + " to " + endDate + "... (" + SCRIPT_VERSION + ")");
      safeToast("Searching Gmail for " + startDate + " to " + endDate + "...", "📊 Financial Tracker", -1);
      const search = searchTransactionEmailsByDateRange(startDate, endDate, config.banksToTrack);
      runGmailMonitorCore(search, config);
    } catch (error) {
      Logger.log("❌ Error in date-range monitor: " + error);
      safeAlert("❌ Error: " + error.toString());
    }
  });
}


/**
 * v1.1.3: confirmation wrapper for recategorizeAllTransactions()
 * (04_sheetsWriter.gs) — menu item "🔁 Recategorize Saved Transactions".
 * v1.1.4: recategorizeAllTransactions() now also corrects Type (not just
 * Category) — updated the confirmation copy to say so.
 */
function recategorizeAllTransactionsPrompt() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();
  if (!config) {
    ui.alert("❌ Setup not completed. Please run Setup Wizard first.");
    return;
  }
  Logger.log("🔁 Starting manual recategorize... (" + SCRIPT_VERSION + ")");
  const response = ui.alert(
    "This will review every already-saved transaction (Transactions + each Raw_<BANK> sheet) " +
    "and update its Type and Category using the current rules (including your Custom Rules). " +
    "It doesn't change amounts, dates, or merchants. Continue?",
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  withRunLock(() => { // v1.1.19 (C2)
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast("Recategorizing...", "📊 Financial Tracker", -1);
      const changed = recategorizeAllTransactions(config.email);
      ensureSheetOrder();   // v1.1.38: this menu item didn't reorder tabs
      SpreadsheetApp.getActiveSpreadsheet().toast("Done.", "📊 Financial Tracker", 3);
      ui.alert("✅ Done — " + changed + " cell(s) updated.");
    } catch (error) {
      ui.alert("❌ Error: " + error.toString());
    }
  });
}

function viewConfig() {
  const config = getConfig();
  if (!config) {
    SpreadsheetApp.getUi().alert("❌ No configuration found. Please run Setup Wizard first.");
    return;
  }
  if (!config.banksToTrack) {
    SpreadsheetApp.getUi().alert("❌ " + (config.configError || "banksToTrack is missing."));
    return;
  }
  
  const configText = `
📊 Current Configuration
═══════════════════════
Email: ${config.email}
Income currency: ${config.incomeCurrency}
Monthly Income: ${config.monthlyIncome} ${config.incomeCurrency}

📉 Deductions (${config.deductionMode === 'auto' ? 'automatic — DR payroll rules' : 'entered manually'}):
  ARS: ${config.deductions.ARS} ${config.incomeCurrency}
  AFP: ${config.deductions.AFP} ${config.incomeCurrency}
  ${config.incomeCurrency === 'DOP' ? 'ISR: ' + config.deductions.ISR + ' DOP' : 'Tax rate: ' + config.deductions.taxRate + '%'}

➕ Other income: ${describeMoneyLines(config.otherIncomes)} (no deductions)
➖ Other deductions: ${describeMoneyLines(config.otherDeductions)}
💳 Cards: ${resolveCards(config.cards).map(c => c.bank + (c.name ? ' ' + c.name : '') + (c.closeDay ? ' · closes day ' + c.closeDay : '') + (c.dueDay ? ' · due day ' + c.dueDay : '')).join('; ') || 'none set'}

📬 Daily summary: ${config.notifyEnabled ? 'on — around ' + config.notifyHour + ':00 to ' + (config.notifyEmail || config.email) : 'off'}
🗓️ Monthly summary: ${config.notifyMonthly ? 'on — the 1st, around ' + config.notifyHour + ':00' : 'off'}

🏦 Banks Tracked:
${BANK_ORDER.map(b => '  ' + b + ': ' + (config.banksToTrack[b] ? '✅' : '❌')).join('\n')}

Version: ${SCRIPT_VERSION}
  `;
  
  SpreadsheetApp.getUi().alert(configText);
}

/**
 * v1.1.3: BUG FIX — used to only delete Config and Transactions, leaving
 * every Raw_<BANK> sheet, Bank Transfers, Dashboard, the Pivot sheet, any
 * monthly Summary sheets, Custom Rules, and the daily trigger all in place.
 * Now discovers and removes everything the system creates.
 * v1.1.4: updated sheet names (Configuration, Custom Rules, Bank Transfers)
 * to match the v1.1.4 rename.
 */
function resetSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ This will delete everything the system created: Configuration, Transactions, every ' +
    'Raw_<BANK> sheet, Bank Transfers, Incoming Transfers, Dashboard, Pivot, monthly summaries, Custom Rules, Unrecognized, the ' +
    'record of emails already read, and the daily triggers (update and summary email). It cannot be undone.\n\n' +
    'Kept: the investment tabs (Investment Ledger, Holdings, Portfolio History, Investment Accounts) — ' +
    'they hold what you typed, which no email can bring back. Continue?',
    ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const fixedNames = new Set([
        CONFIG_SHEET, TRANSACTIONS_SHEET, "Bank Transfers", INCOMING_SHEET, "Dashboard",   // v1.1.53: Incoming
        "Pivot - Category x Bank", CUSTOM_RULES_SHEET,
        UNRECOGNIZED_SHEET, READ_LOG_SHEET   // v1.1.49: without these, re-reading after a reset skipped every email
      ]);
      const toDelete = ss.getSheets().filter(s => {
        const name = s.getName();
        return fixedNames.has(name) || name.startsWith("Raw_") || name.startsWith("Summary-");
      });
      toDelete.forEach(s => {
        // Apps Script requires at least one sheet to remain in the workbook.
        if (ss.getSheets().length > 1) ss.deleteSheet(s);
      });
      deleteSystemTriggers(); // v1.1.19 (M12) / v1.1.24: only this system's triggers
      PropertiesService.getScriptProperties().deleteProperty(LAST_RUN_PROPERTY);   // v1.1.49: no "last update" from before the reset
      ui.alert("✅ System reset (" + toDelete.length + " sheet(s) deleted). Run Setup Wizard to start again.");
    } catch (error) {
      ui.alert("❌ Error: " + error);
    }
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function getConfig() {
  // v1.1.19 (M11): a malformed banksToTrack sets `configError` (shown by
  // requireConfig()) instead of pretending setup never ran.
  // v1.1.24: deductionMode, ISR and the daily summary settings; defaults keep
  // a Configuration saved by an older version working unchanged.
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const config = { deductions: { ARS: 0, AFP: 0, taxRate: 0, ISR: 0 } };
    const asBool = v => v === true || String(v).trim().toUpperCase() === 'TRUE';

    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0]).trim();
      const value = data[i][1];
      if (key === 'banksToTrack') {
        try {
          config.banksToTrack = JSON.parse(value);
        } catch (parseError) {
          config.banksToTrack = null;
          config.configError = "Configuration → banksToTrack is not valid (" + value +
            "). Run the Setup Wizard again to fix it.";
        }
      } else if (['ARS', 'AFP', 'taxRate', 'ISR'].includes(key)) {
        config.deductions[key] = parseFloat(value) || 0;
      } else if (key === 'monthlyIncome') {
        const n = parseFloat(value);
        config.monthlyIncome = isNaN(n) ? 0 : n;
      } else if (key === 'otherIncome') {
        config.otherIncome = parseFloat(value) || 0;
      } else if (key === 'otherIncomes' || key === 'otherDeductions') {   // v1.1.59
        try { config[key] = JSON.parse(value); } catch (e) { config[key] = []; }
      } else if (key === 'cards') {
        try { config.cards = JSON.parse(value); } catch (e) { config.cards = []; }
        if (!Array.isArray(config.cards)) config.cards = [];
      } else if (key === 'notifyEnabled') {
        config.notifyEnabled = asBool(value);
      } else if (key === 'notifyMonthly') {
        config.notifyMonthly = asBool(value);
      } else if (key === 'notifyHour') {
        config.notifyHour = parseInt(value, 10);
      } else if (key === 'notifySections') {
        try { config.notifySections = JSON.parse(value); } catch (e) { config.notifySections = null; }
      } else if (key) {
        config[key] = value;
      }
    }
    if (!config.incomeCurrency) config.incomeCurrency = 'USD';
    config.otherIncome = Number(config.otherIncome) || 0;                  // v1.1.27
    if (config.otherIncomeCurrency !== 'USD') config.otherIncomeCurrency = 'DOP';
    // v1.1.59: one "other income" became a list — an older setup's single one becomes its first line
    if (!Array.isArray(config.otherIncomes)) {
      config.otherIncomes = config.otherIncome > 0 ? [{ label: '', amount: config.otherIncome, currency: config.otherIncomeCurrency }] : [];
    }
    config.otherIncomes = normalizeMoneyLines(config.otherIncomes);
    config.otherDeductions = normalizeMoneyLines(config.otherDeductions);
    if (config.deductionMode !== 'auto') config.deductionMode = 'manual';
    config.notifyEnabled = !!config.notifyEnabled;
    config.notifyMonthly = !!config.notifyMonthly;
    if (!(config.notifyHour >= 0 && config.notifyHour <= 23)) config.notifyHour = 8;
    config.notifySections = Object.assign({}, SUMMARY_SECTIONS_DEFAULT, config.notifySections || {});
    config.notifyEmail = String(config.notifyEmail || '').trim();
    return config.email ? config : null;
  } catch (error) {
    Logger.log("Error reading config: " + error);
    return null;
  }
}

// NOTE: parseTransaction/extractMerchant/categorizeTransaction/saveTransaction
// used to be duplicated here with older, incompatible signatures. The real
// implementations now live in:
//   - parseEmailMessage(), extractAllAmounts(), extractMerchantFromContext() -> 03_gmailMonitor.gs
//   - categorizeTransaction(text, customRules) -> 02_categorizer.gs
//   - saveTransaction(transaction), saveTransactions(transactions) -> 04_sheetsWriter.gs
// Keeping a single definition per function avoids Apps Script's global
// namespace silently overwriting one file's function with another's.

function getOrCreateSheet(sheetName) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  }
  return sheet;
}

/**
 * v1.1.2: BUG FIX — this used to just show a "coming in Phase 2" alert and
 * never built anything, which is why the dashboard "wasn't updating": there
 * was no dashboard. Now builds (or refreshes/re-links) a real Dashboard
 * sheet with live SUMIFS formulas (see buildOrRefreshDashboard(),
 * 04_sheetsWriter.gs) and activates it so the user lands on it directly.
 */
function openDashboard() {
  const sheet = buildOrRefreshDashboard();
  SpreadsheetApp.setActiveSheet(sheet);
}

// Trigger to run Gmail Monitor daily
// v1.1.5: changed from 8:00 AM to 6:00 AM, per request. If you set up your
// system before this change, the trigger created back then still fires at
// 8:00 AM — run Setup Wizard again (or just createTrigger() from the Apps
// Script editor's function dropdown) to pick up the new time; saveSetupConfig()
// always deletes and recreates the trigger, so this is safe to repeat.
function createTrigger() {
  deleteSystemTriggers();
  ScriptApp.newTrigger("runGmailMonitor").timeBased().atHour(6).everyDays(1).create();
  Logger.log("✅ Daily trigger created for 6:00 AM");
  // v1.1.24: the daily summary gets its own trigger, only when it's enabled
  const config = getConfig();
  if (config && config.notifyEnabled) {
    ScriptApp.newTrigger("sendDailySummary").timeBased().atHour(config.notifyHour).everyDays(1).create();
    Logger.log("✅ Daily summary trigger created for " + config.notifyHour + ":00");
  }
  // v1.1.26: monthly summary on the 1st, same hour (after that morning's update)
  if (config && config.notifyMonthly) {
    ScriptApp.newTrigger("sendMonthlySummary").timeBased().onMonthDay(1).atHour(config.notifyHour).create();
    Logger.log("✅ Monthly summary trigger created for the 1st at " + config.notifyHour + ":00");
  }
}

const SYSTEM_TRIGGER_HANDLERS = ["runGmailMonitor", "sendDailySummary", "sendMonthlySummary"];

/** v1.1.19 (M12) / v1.1.24: removes only this system's triggers (monitor + daily summary). */
function deleteSystemTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => SYSTEM_TRIGGER_HANDLERS.indexOf(t.getHandlerFunction()) !== -1)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/* ======================================================================
 * OTHER INCOMES AND DEDUCTIONS — v1.1.59
 * Lines of { label, amount, currency }, as many as needed (up to MONEY_LINES_MAX), in DOP, USD or EUR — the currencies
 * the Dashboard has live rates for. Configuration keeps the lines and their totals per currency; the Dashboard and the
 * summaries convert those totals at their rates.
 * ====================================================================== */
const OTHER_MONEY_CURRENCIES = ['DOP', 'USD', 'EUR'];
const MONEY_LINES_MAX = 10;

function normalizeMoneyLines(list) {
  return (Array.isArray(list) ? list : []).filter(x => x && Number(x.amount) > 0)
    .map(x => ({ label: String(x.label || '').trim().slice(0, 40), amount: +Number(x.amount).toFixed(2),
      currency: OTHER_MONEY_CURRENCIES.indexOf(x.currency) !== -1 ? x.currency : 'DOP' }))
    .slice(0, MONEY_LINES_MAX);
}

function moneyLineTotals(list) {
  const t = { DOP: 0, USD: 0, EUR: 0 };
  normalizeMoneyLines(list).forEach(x => { t[x.currency] += x.amount; });
  Object.keys(t).forEach(c => { t[c] = +t[c].toFixed(2); });
  return t;
}

/** DOP-equivalent of a list, at the given rates ({ USD, EUR } in DOP). */
function moneyLinesDop(list, rates) {
  const t = moneyLineTotals(list);
  return t.DOP + t.USD * (Number(rates && rates.USD) || 0) + t.EUR * (Number(rates && rates.EUR) || 0);
}

/** Configuration rows with each list's totals per currency (the Dashboard's named ranges). */
function moneyTotalsRows(config) {
  const inc = moneyLineTotals(config.otherIncomes), ded = moneyLineTotals(config.otherDeductions);
  return OTHER_MONEY_CURRENCIES.map(c => ['otherIncome' + c, inc[c]]).concat(OTHER_MONEY_CURRENCIES.map(c => ['otherDeduction' + c, ded[c]]));
}

function describeMoneyLines(list) {
  const lines = normalizeMoneyLines(list);
  return lines.length ? lines.map(x => (x.label ? x.label + ' ' : '') + x.amount + ' ' + x.currency).join('; ') : 'none';
}

// ====================================================================================================
// 02_categorizer.gs
// ====================================================================================================

/**
 * CATEGORIZER MODULE
 * Handles transaction categorization with extensible rules
 */

/**
 * v1.1.5: further category consolidation, per direct request:
 * - 'Housing' → 'Rent' (clearer/shorter).
 * - 'Internet' + 'Mobile Data' merged into one 'Telecommunications' category
 *   (both keyword lists combined).
 * - 'Clothing' and 'Other' removed as separate categories — everything that
 *   would have landed in either now falls into
 *   'Dining/Delivery + Entertainment + Other' instead (Clothing's keywords
 *   were folded in, and the categorizeTransaction() fallback below now
 *   returns that label instead of 'Other'). Note the side effect: a
 *   genuinely unmatched merchant will now show as Dining rather than a
 *   distinct "uncategorized" bucket — that's the tradeoff of this merge.
 * - Added real merchants seen in live data that were falling through to the
 *   old 'Other' bucket: PedidosYa (a delivery app — every PedidosYa* order
 *   is food), Krispy Kreme, IKEA Restaurant.
 * Now 11 categories total (was 13).
 */

/**
 * v1.1.4: category names translated to English (were the only Spanish-named
 * part of an otherwise-English system, e.g. "Vivienda" next to an English
 * "Category" header). Keywords themselves stay in Spanish — they're
 * matching against Spanish bank-email text, which didn't change.
 */
const DEFAULT_CATEGORIES = {
  'Rent': {
    // v1.1.19: removed bare 'RENT' and 'RENTA' — substring matches sent
    // "AUTO RENT A CAR" (verified) and similar to Rent, a FIXED category.
    // A landlord is matched by a Custom Rule (e.g. the landlord's name),
    // which is how real rent payments have been categorized all along.
    keywords: ['ALQUILER', 'ARRENDAMIENTO'],
    icon: '🏠',
    color: ['#EDE9E4', '#57534E']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Gym + Calisthenics': {
    keywords: ['GYM', 'GIMNASIO', 'CALISTENIA', 'FITNESS', 'CROSSFIT', 'PILATES', 'YOGA', 'ENTRENADOR'],
    icon: '💪',
    color: ['#CCFBF1', '#0F766E']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Telecommunications': {
    // v1.1.5: merged Internet + Mobile Data. 'CLARO' bare (not just
    // 'CLARO MOVIL') so "CLARO P REC" (Claro Prepago Recarga, a real
    // merchant descriptor) matches too.
    keywords: ['INTERNET', 'WIFI', 'BANDA ANCHA', 'CLARO INTERNET',
               'DATOS MOVILES', 'RECARGA', 'PLAN MOVIL', 'CLARO MOVIL', 'CLARO', 'PREPAGO'],
    icon: '📶',
    color: ['#F3E8FF', '#7E22CE']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Streaming & Subscriptions': {
    // v1.1.15: BUG FIX — bare 'AMAZON' here caught EVERY Amazon purchase
    // (Amazon.com retail orders, Amazon Marketplace), not just the Prime
    // subscription itself. Narrowed to 'AMAZON PRIME' only; general Amazon
    // purchases now fall through to Dining's bare 'AMAZON' keyword instead
    // (Streaming is checked first, so 'AMAZON PRIME' still correctly
    // intercepts the subscription case before Dining's broader keyword
    // would otherwise catch it too).
    // v1.1.23: 'GOOGLE' — real POPULAR USD rows "Google" / "GOOGLE *Google" (Google One/Play)
    keywords: ['NETFLIX', 'SPOTIFY', 'DISNEY', 'HBO', 'AMAZON PRIME', 'SUSCRIPCION',
               'YOUTUBE PREMIUM', 'OPENAI', 'GOOGLE'],
    icon: '🎬',
    color: ['#E0E7FF', '#3730A3']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Electricity': {
    // v1.1.19: removed bare 'LUZ' — "LUZ DE LUNA CAFE" landed here
    // (verified). The three distributors below cover every real bill.
    keywords: ['EDESUR', 'EDENORTE', 'EDEESTE', 'ELECTRICIDAD'],
    icon: '💡',
    color: ['#FEF9C3', '#854D0E']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Groceries + Barbershop': {
    // v1.1.23: 'POLA' — Supermercados Pola ("SM POLA INDEPENDENCIA"); whole word
    keywords: ['SUPER', 'CARREFOUR', 'JUMBO', 'BRAVO', 'MERCADO', 'FRUTAS', 'BARBERIA', 'SALON',
               'SUPERMERCADOS NACIONAL', 'SM NACIONAL', 'LA SIRENA', 'PLAZA LAMA', 'IBERIA',
               'APREZIO', 'LA FUENTE', 'PRICESMART', 'HIPER OLE', '365 EL CACIQUE', 'HOLA PLAZA', 'POLA'],
    icon: '🛒',
    color: ['#FFF4E5', '#B45309']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Dining/Delivery + Entertainment + Other': {
    // v1.1.5: absorbed Clothing's keywords and the old 'Other' fallback role
    // (see header comment), plus PedidosYa/Krispy Kreme/IKEA Restaurant from
    // real transaction data.
    // v1.1.17: added bare 'UBER EATS' (space, no asterisk) — a real
    // merchant descriptor "UBER EATS-W*UBER EATS- SANTO DOMINGO DOM" didn't
    // match 'UBER*EATS' or 'UBEREATS' and fell through all the way to
    // Transportation's bare 'UBER' keyword instead.
    // v1.1.15: added bare 'AMAZON' (see Streaming & Subscriptions note
    // above — general Amazon purchases live here now, only 'AMAZON PRIME'
    // stays under Streaming).
    // v1.1.15: this category is now checked BEFORE Vehicle Gas (see the
    // reordering note on Vehicle Gas below) specifically so 'BONJOUR' here
    // intercepts "BONJOUR TOTAL ARENOSO" before Vehicle Gas's bare 'TOTAL'
    // keyword ever gets a chance to wrongly match it.
    // v1.1.23: ALISS and the Metro Plaza mall ("ALISS METRO PZA") — 'METRO' alone is Transportation, checked later
    keywords: ['RESTAURANTE', 'PIZZERIA', 'BURGER', 'COMIDA', 'FOOD', 'CAFE', 'DELIVERY',
               'UBER*EATS', 'UBEREATS', 'UBER EATS', 'CINE', 'CINEMA', 'PELICULAS', 'CONCIERTO',
               'ENTRETENIMIENTO', 'BAR', 'SUSHI', 'POLLO', 'HELADOS BON', 'BONJOUR',
               'MCDONALD', 'BURGER KING', 'KFC', 'PIZZA HUT', 'SUBWAY', 'POLLOS VICTORINA',
               'COMEDOR', 'WENDYS', 'JADE TERIYAKI', 'SWEETFROG', 'SWEET FROG',
               'PEDIDOSYA', 'KRISPY KREME', 'IKEA REST', 'AMAZON',
               'ROPA', 'ZAPATOS', 'ADIDAS', 'NIKE', 'PUMA STORE', 'ZARA', 'FASHION', 'BOUTIQUE', 'TIENDA', 'ALISS', 'METRO PZA', 'METRO PLZA', 'METRO PLAZA'],
    icon: '🍽️',
    color: ['#FFE4E6', '#BE123C']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Vehicle Gas': {
    // v1.1.15: BUG FIX/CHANGE — added bare 'TOTAL' and 'NEXT LINCOLN' (a
    // real gas-station brand, confirmed as "NEXT LINCOLN GASOPOLIS" from
    // LAFISE and bare "NEXT LINCOLN" from POPULAR — same business, two
    // different card-descriptor truncations). Bare 'TOTAL' was deliberately
    // left OUT before specifically because "BONJOUR TOTAL ARENOSO" (the
    // convenience-store purchase, → Dining) also contains "TOTAL" as a
    // separate word — now safe to add because Dining/Delivery + ... (with
    // 'BONJOUR') is checked BEFORE this category (see its reordering note),
    // so the Bonjour case is intercepted there first, before ever reaching
    // this keyword. 'TOTALENERGIES' kept too, redundant with bare 'TOTAL'
    // now but harmless.
    keywords: ['SHELL', 'HESS', 'PUMA', 'TEXACO', 'GASOLINA', 'COMBUSTIBLE', 'FUEL',
               'TOTALENERGIES', 'TOTAL', 'NEXT LINCOLN'],
    icon: '⛽',
    color: ['#E2E8F0', '#334155']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Health + Vet + Pharmacy': {
    // v1.1.18: was 'LA CASITA DE BOB' — narrowed to 'LA CASITA DE' (the
    // vet clinic) because POPULAR truncates the merchant name to just "LA
    // CASITA DE" (no "BOB-EP" suffix), which didn't match the longer
    // keyword and fell through to the Dining fallback. The shorter keyword
    // is a substring of the longer one, so it still matches LAFISE/BHD's
    // full "LA CASITA DE BOB-EP..." too — one keyword now covers both.
    // v1.1.23: 'FARMA' (FARMA VALUE…), 'FARM' whole word (FARM CAROL…), labs and clinics (ANALISA, CTRO ESPEC MED)
    keywords: ['FARMACIA', 'FCIA', 'DOCTOR', 'MEDICO', 'HOSPITAL', 'CLINICA', 'SALUD',
               'MEDICINAS', 'VETERINARIA', 'VET', 'MEDICAR', 'GBC', 'LOS HIDALGOS',
               'FARMAX', 'FARMAXTRA', 'CRUZ VERDE', 'FARMATODO', 'FARMAVALUE', 'LA CASITA DE', 'FARMA', 'FARM', 'ANALISA', 'LABORATORIO', 'ESPEC MED'],
    icon: '🏥',
    color: ['#E8F5E9', '#2E7D32']   // v1.1.28: chip [background, text] in the data sheets
  },
  // v1.1.24: 'Education' removed — not in use. Those purchases now fall to the
  // Dining/Other fallback; a Custom Rule with Category "Education" brings it
  // back as its own Dashboard row (custom-only categories get one automatically).
  'Transportation': {
    // v1.1.20: added 'DIDI' — real POPULAR rows in COP ("DiDi CO Ride",
    // "DL*DIDI RIDES", "DLO*Didi") were falling through to the Dining
    // fallback. Whole-word match (4 letters), so the '*' variants match;
    // "DIDI FOOD" still goes to Dining via 'FOOD', which is checked first.
    // v1.1.23: Uber rides billed through PayPal ("PAYPAL *UBERBV", "PAYPAL *UBER BV")
    keywords: ['UBER', 'UBER*RIDES', 'UBER*TRIP', 'DIDI', 'TAXI', 'BUS', 'METRO', 'PARKING', 'GUAGUA', 'MOTOCONCHO', 'UBERBV', 'UBER BV'],
    icon: '🚗',
    color: ['#EAF1FE', '#1D4ED8']   // v1.1.28: chip [background, text] in the data sheets
  }
};

const FALLBACK_CATEGORY = 'Dining/Delivery + Entertainment + Other';

/**
 * v1.1.19: SHARED KEYWORD ENGINE — used by categorization, Custom Rules,
 * type detection and the promotional/non-transactional/declined filters
 * (03_gmailMonitor.gs), so every keyword list follows one matching rule.
 *
 * Why: plain substring matching caused a whole class of bugs — 'ACH' matched
 * inside "CACHAREPA", 'RENT' inside "AUTO RENT A CAR", 'ROPA' inside
 * "EUROPA", 'BUS' inside "BUSINESS". Rule now:
 *   - keywords of 4 or fewer letters/digits (BAR, BUS, VET, ACH, OTP, UBER…)
 *     must match as a WHOLE WORD — the characters on each side must not be
 *     a letter or digit ('*', '-', space, start/end of text all count as
 *     boundaries, so "UBER*RIDES" still matches 'UBER');
 *   - longer keywords keep SUBSTRING matching, which several rely on on
 *     purpose ('SUPER' → SUPERMERCADO, 'FARMAX' → FARMAXTRA, 'TOTAL' →
 *     "TOTALENERGIES"-style descriptors).
 * Regexes are compiled once and cached per keyword.
 */
const KEYWORD_WORD_CHARS = 'A-Z0-9ÁÉÍÓÚÑÜ';
const KEYWORD_WHOLE_WORD_MAX_LEN = 4;
const _keywordRegexCache = {};

/**
 * Uppercase + v1.1.23 spacing rules, applied to BOTH the text and every
 * keyword: spaces around '*' are removed ("UBER * EATS" → "UBER*EATS" — real
 * LAFISE rows fell to Transportation because of the spaces) and runs of
 * whitespace collapse to one space.
 */
function normalizeKeyword(keyword) {
  return String(keyword === null || keyword === undefined ? '' : keyword)
    .toUpperCase().replace(/\s*\*\s*/g, '*').replace(/\s+/g, ' ').trim();
}

function keywordMatches(upperText, keyword) {
  const kw = normalizeKeyword(keyword);
  if (!kw || !upperText) return false;       // an empty keyword never matches (see C4 fix below)
  const coreLen = kw.replace(new RegExp('[^' + KEYWORD_WORD_CHARS + ']', 'g'), '').length;
  if (coreLen > KEYWORD_WHOLE_WORD_MAX_LEN) return upperText.includes(kw);
  let re = _keywordRegexCache[kw];
  if (!re) {
    const esc = kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    re = new RegExp('(^|[^' + KEYWORD_WORD_CHARS + '])' + esc + '(?=$|[^' + KEYWORD_WORD_CHARS + '])');
    _keywordRegexCache[kw] = re;
  }
  return re.test(upperText);
}

/** Returns the first keyword in `keywords` that matches, or null. */
function findMatchingKeyword(upperText, keywords) {
  for (const kw of keywords || []) {
    if (keywordMatches(upperText, kw)) return kw;
  }
  return null;
}

/**
 * Main categorization function — DEFAULT_CATEGORIES only, evaluated in
 * declaration order (order matters, see the category notes above).
 * v1.1.5: fallback changed from 'Other' to FALLBACK_CATEGORY.
 * v1.1.19: the old `customRules` (merged defaults + Custom Rules) parameter
 * is gone. Every caller already runs findCustomRuleOverride() FIRST, so the
 * merged rules could only ever be reached when no custom rule matched — at
 * which point they were identical to the defaults. Also uses the shared
 * keyword engine above instead of raw substring matching.
 */
function categorizeTransaction(text) {
  if (!text) return FALLBACK_CATEGORY;
  const upperText = normalizeKeyword(text);
  for (let category in DEFAULT_CATEGORIES) {
    if (findMatchingKeyword(upperText, DEFAULT_CATEGORIES[category].keywords)) {
      return category;
    }
  }
  return FALLBACK_CATEGORY;
}

/**
 * v1.1.5: checks ONLY the user's Custom Rules (not the defaults) against
 * `text`, regardless of transaction Type. Used so a personal rule — e.g.
 * "<LANDLORD NAME>" → Rent for a landlord's transfer — can override the normal
 * "only Type=Transaction gets a category" rule (03_gmailMonitor.gs /
 * 04_sheetsWriter.gs both call this BEFORE deciding whether to categorize
 * at all). Returns the matched category, or '' if no custom rule matches.
 */
function findCustomRuleOverride(text, rawCustomRules) {
  // v1.1.19: String() + shared engine — a numeric keyword (e.g. 1111, an
  // account's last 4 digits, which Sheets stores as a number) used to throw
  // "toUpperCase is not a function" here.
  if (!text || !rawCustomRules) return '';
  const upperText = normalizeKeyword(text);
  for (let category in rawCustomRules) {
    if (findMatchingKeyword(upperText, rawCustomRules[category])) return category;
  }
  return '';
}

/**
 * Get all available categories
 */
function getCategories() {
  return Object.keys(DEFAULT_CATEGORIES);
}

/**
 * Get category with icon
 */
function getCategoryWithIcon(category) {
  if (DEFAULT_CATEGORIES[category]) {
    return DEFAULT_CATEGORIES[category].icon + ' ' + category;
  }
  return '📌 ' + category;
}

/**
 * Add custom rule to user config
 */
function addCustomCategoryRule(userEmail, category, keyword) {
  try {
    const sheet = getOrCreateSheet(CUSTOM_RULES_SHEET);
    sheet.appendRow([
      userEmail,
      category,
      normalizeKeyword(keyword),
      new Date().toISOString()
    ]);
    return true;
  } catch (error) {
    Logger.log("Error adding custom rule: " + error);
    return false;
  }
}

/**
 * Get custom rules for user
 * v1.1.6: email comparison now trims whitespace and ignores case — a
 * mismatched email (extra space, different case) between what's typed into
 * the Custom Rules sheet and what Configuration stores would otherwise
 * silently return zero rules, which is one of the ways the landlord/Rent
 * override could fail to apply even with the right keyword.
 */
function getUserCustomRules(userEmail) {
  try {
    const sheet = getOrCreateSheet(CUSTOM_RULES_SHEET);
    const data = sheet.getDataRange().getValues();
    const customRules = {};
    const targetEmail = String(userEmail).trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][0] || '').trim().toLowerCase();
      if (rowEmail === targetEmail) {
        // v1.1.19: BUG FIX (C4) — values used to be pushed raw. A blank
        // keyword cell became '' and `text.includes('')` is ALWAYS true, so
        // that one rule captured every transaction; a numeric keyword (1111)
        // threw inside every parse and every recategorize, silently dropping
        // every email of the run. Incomplete rows are now skipped.
        const category = String(data[i][1] === null || data[i][1] === undefined ? '' : data[i][1]).trim();
        const keyword = normalizeKeyword(data[i][2]);
        if (!category || !keyword) continue;

        if (!customRules[category]) {
          customRules[category] = [];
        }
        customRules[category].push(keyword);
      }
    }
    
    return customRules;
  } catch (error) {
    Logger.log("Error reading custom rules: " + error);
    return {};
  }
}

const EXCLUDE_CATEGORY = 'Exclude';

/**
 * v1.1.28: chip colours for categories that aren't in DEFAULT_CATEGORIES —
 * "Exclude" (muted: it's left out of every total) and your own Custom Rules
 * categories, which take the next colour of this rotation in order.
 */
const EXCLUDE_COLOR = ['#F3F4F6', '#4B5563'];   // #6B7280 was 4.39:1 — below WCAG AA (caught by v128 test)
const CUSTOM_CATEGORY_COLORS = [
  ['#CFFAFE', '#0E7490'], ['#FAE8FF', '#A21CAF'], ['#D1FAE5', '#047857'], ['#FEF3C7', '#92400E'], ['#FCE7F3', '#9D174D']
];

/** Every category with its chip colours, in display order: defaults, Exclude, then custom ones. */
function categoryPalette(customNames) {
  const list = getCategories().map(name => ({ name: name, bg: DEFAULT_CATEGORIES[name].color[0], fg: DEFAULT_CATEGORIES[name].color[1] }));
  list.push({ name: EXCLUDE_CATEGORY, bg: EXCLUDE_COLOR[0], fg: EXCLUDE_COLOR[1] });
  (customNames || []).filter(n => n && !DEFAULT_CATEGORIES[n] && n.toUpperCase() !== EXCLUDE_CATEGORY.toUpperCase())
    .forEach((name, i) => {
      const c = CUSTOM_CATEGORY_COLORS[i % CUSTOM_CATEGORY_COLORS.length];
      list.push({ name: name, bg: c[0], fg: c[1] });
    });
  return list;
}

/**
 * v1.1.19: categories that exist ONLY in Custom Rules (not in
 * DEFAULT_CATEGORIES, and not the reserved "Exclude"). The Dashboard grid
 * used to be built from the defaults alone, so spend in a custom category
 * (say "Pets") never reached any row or the TOTAL — see
 * buildOrRefreshDashboard(). getMergedCategoryRules() was removed in this
 * version (see categorizeTransaction()).
 */
function getCustomCategoryNames(userEmail) {
  const custom = getUserCustomRules(userEmail);
  return Object.keys(custom).filter(c =>
    !DEFAULT_CATEGORIES[c] && c.toUpperCase() !== EXCLUDE_CATEGORY.toUpperCase());
}

// ====================================================================================================
// 03_gmailMonitor.gs
// ====================================================================================================

/**
 * GMAIL MONITOR MODULE
 * Searches and parses transaction emails from different banks
 */

/**
 * v1.1.0: senders, subjects, AND real body formats are now all confirmed —
 * the user shared actual .eml files for 6 real transaction emails
 * (Sept 23, 2026), so extraction below (see the bank-specific extractXxx()
 * functions further down) is written and tested directly against real text,
 * not guessed. Summary per bank:
 *   LAFISE   — notificaciones@bancolafise.com
 *              Consumo ("Servicio de Alerta - Nuevo Consumo"): labeled fields,
 *              "Comercio/Ciudad/País:" (merchant) BEFORE "Monto:" (amount,
 *              format "DOP 412.86" — no $ sign, uses the DOP code instead).
 *              Transferencia ("<Nombre>, ¡Transferencia exitosa!"): HTML only
 *              (empty plain part); real sample seen in v1.1.66, see
 *              extractLAFISETransferTransactions().
 *   BANESCO  — notificaciones@banesco.com.do
 *              Consumo, approved ("Alerta de Consumo Banesco RD"): one
 *              sentence — "...consumo de RD$ 2,640.00, en MERCHANT y su
 *              estado es aprobada." Merchant comes AFTER the amount here,
 *              opposite of LAFISE. A declined variant exists with different
 *              wording ("...ha sido rechazada...") — skipped per-match via a
 *              status check, not by rejecting the whole email.
 *              Transferencia ("Notificación de Transferencia Realizada"):
 *              "Monto: DOP545.00" (no $, no space) ... later ... "Nombre del
 *              Beneficiario: X".
 *   BHD      — Alertas@bhd.com.do
 *              Consumo ("BHD Notificación de Transacciones"): an actual
 *              table — header row (Fecha/Moneda/Monto/Comercio/Estado/Tipo)
 *              then one repeating 6-line block per transaction. This is
 *              where same-day bundling is most likely to actually show up
 *              (multiple row-blocks) — the extractor uses a global regex so
 *              it naturally picks up as many rows as exist.
 *              Transferencia ("Transacciones entre mis productos"): labeled
 *              fields, "Monto:" then later "Beneficiario:".
 *   POPULAR  — notificaciones@popularenlinea.com
 *              Consumo ("Notificación de Consumo"): a compact table (Monto /
 *              Moneda / Fecha / Comercio / Estatus) on one visual row, amount
 *              format "US$14.27".
 *   BDI      — bdi.com.do confirmed, but still no consumption-specific
 *              sample — kept on the old generic extractor (extractAllAmounts)
 *              until a real one shows up.
 *
 * Currency formats actually observed across these 6 samples: "RD$ 2,640.00",
 * "RD 8,315.20" (no $), "DOP 412.86", "DOP545.00" (no space), "RD$ 91.37",
 * "RD" + "$118.40" split across a table cell/newline, "US$14.27". All of
 * every extractXxx() function (and the generic extractAllAmounts() fallback) use \b(?:RD|DOP|USD|US\$|EUR|COP) then an
 * independently-optional \$? — that's what makes both the attached
 * ("RD$91.37") and detached ("RD" ... "$118.40") cases match correctly.
 */
// v1.1.19 (E4): each bank declares its own extractors ({consumo, transfer}),
// so extractTransactionItems() is a lookup instead of an if/else chain and
// adding BDI later means filling in two functions here. The unused
// `amountPattern` field was removed (M13) — nothing had read it since the
// bank-specific extractors replaced it in v1.1.0. Extractor functions are
// plain function declarations further down (hoisted, so referencing them
// here is safe).
const BANK_PATTERNS = {
  LAFISE: {
    name: 'LAFISE',
    // v1.1.5: LAFISE actually sends from TWO different domains — consumption
    // alerts from bancolafise.com, but transfers ("<Name>, ¡Transferencia
    // exitosa!") from a separate digital@notificaciones.lafise.com address,
    // confirmed directly by the user. Neither domain is a substring of the
    // other, so fromDomain is broadened to the shared suffix 'lafise.com'
    // (matches both) and searchQuery now ORs both exact addresses. This is
    // why LAFISE transfers were invisible before — the old search only ever
    // looked for the consumption-alert sender.
    fromDomain: 'lafise.com',
    // v1.1.51: incoming "Pagos al Instante" transfers come from a third address
    // v1.1.61: two more, from real samples: card payments (bancanet@notificaciones.lafise.com) and a second card-alert
    // template (notificacioneslafisedo@lafise.com.do). Neither was searched, so neither ever reached the sheet.
    searchQuery: 'from:notificaciones@bancolafise.com OR from:digital@notificaciones.lafise.com OR from:PagosAlInstanteMT103@lafise.com' +
      ' OR from:bancanet@notificaciones.lafise.com OR from:notificacioneslafisedo@lafise.com.do',
    extractors: { consumo: extractLAFISEAnyConsumo, transfer: extractLAFISEAnyTransfer, incoming: extractLAFISEIncomingTransactions,
      cardPayment: extractLAFISECardPaymentTransactions },
    keywords: ['LAFISE', 'LAFISE BANCO'],
    merchantPattern: /(?:en|en el|en\s+)([^\n]{10,50})/i
  },
  BANESCO: {
    name: 'BANESCO',
    fromDomain: 'banesco.com.do',
    // v1.1.51: the monthly savings statement (a PDF) — Banesco doesn't notify most incoming transfers
    searchQuery: 'from:notificaciones@banesco.com.do OR from:estadodecuenta@banesco.com.do',
    extractors: { consumo: extractBANESCOConsumoTransactions, transfer: extractBANESCOTransferTransactions,
      incoming: extractBANESCOIncomingTransactions },   // v1.1.62: "Notificación de Transferencia Recibida"
    statement: { subject: /ESTADO DE CUENTA DE AHORROS/i, parse: parseBanescoSavingsStatement },
    keywords: ['BANESCO', 'BANESCO RD'],
    merchantPattern: /realizada en el\s+([^\n]+?)\s+por\s/i
  },
  BHD: {
    name: 'BHD',
    fromDomain: 'bhd.com.do',
    searchQuery: 'from:Alertas@bhd.com.do',
    extractors: { consumo: extractBHDConsumoTransactions, transfer: extractBHDTransferTransactions },
    keywords: ['BHD', 'BANCO HOTELES DOMINICANA'],
    merchantPattern: /(?:en|en el)([^\n]{10,50})/i
  },
  POPULAR: {
    name: 'POPULAR',
    fromDomain: 'popularenlinea.com',
    searchQuery: 'from:notificaciones@popularenlinea.com',
    extractors: { consumo: extractPOPULARTransactions, transfer: extractPOPULARTransactions },
    keywords: ['POPULAR', 'BANCO POPULAR'],
    merchantPattern: /(?:en|en el)([^\n]{10,50})/i
  },
  BDI: {
    name: 'BDI',
    fromDomain: 'bdi.com.do',
    searchQuery: 'from:bdi.com.do',
    // v1.1.60: real samples — card purchases ("Notificacion de Consumos") and interbank transfers sent
    // v1.1.61: interbank transfers RECEIVED (same subject as the ones sent; the body says "Recibida"), and strict: an
    // email none of these recognizes goes to Unrecognized instead of being guessed (see STRICT_BANK_NOTE)
    extractors: { consumo: extractBDIConsumoTransactions, transfer: extractBDITransferTransactions, incoming: extractBDIIncomingTransactions,
      cardPayment: extractBDICardPaymentTransactions },
    detectType: (subject, text) => /Interbancaria\s+Recibida/i.test(text) ? 'Incoming' :
      /Tipo de Transacci[óo]n\s+Pago Tarjetas? de Cr[ée]dito/i.test(text) ? 'Card Payment' : null,   // v1.1.62
    // v1.1.62: "… Interbancaria - Completada" confirms a transfer already read from its "Comprobante" (same time)
    ignore: [/Transacci[óo]n Interbancaria\s*-\s*Completada/i],
    strict: true,
    keywords: ['BDI', 'BDI DIGITAL', 'BANCO BDI'],
    merchantPattern: /(?:en|en el|hacia|a)([^\n]{10,50})/i
  },
  // v1.1.60: two more banks, from real samples
  SCOTIABANK: {
    name: 'SCOTIABANK',
    fromDomain: 'scotiabank.com',                        // alertas@scotiabank.com, and *.scotiabank.com.do
    searchQuery: 'from:alertas@scotiabank.com OR from:scotiabank.com.do',
    extractors: { consumo: extractSCOTIABANKConsumoTransactions, incoming: extractSCOTIABANKIncomingTransactions },   // v1.1.61: incoming
    detectType: (subject, text) => /recibido un cr[ée]dito a su cuenta/i.test(text) ? 'Incoming' : null,
    strict: true,                                        // v1.1.61
    keywords: ['SCOTIABANK', 'SCOTIA'],
    merchantPattern: /\ben\s+(.{3,50}?)\s+con su/i
  },
  QIK: {
    name: 'QIK',
    fromDomain: '@qik.',                                 // notificaciones@qik.do, …@qik.com.do
    searchQuery: 'from:qik.do OR from:qik.com.do',
    extractors: { consumo: extractQIKConsumoTransactions },   // v1.1.61: also Código CASH withdrawals
    // v1.1.62: a Código CASH CREATED is not money out yet; the withdrawal comes in its own email ("… utilizado")
    ignore: [/C[óo]digo CASH para ti creado|C[óo]digo CASH para ti ha sido creado/i],
    strict: true,                                        // v1.1.61
    keywords: ['QIK'],
    merchantPattern: /\ben\s+(.{3,50}?)\s+con tu tarjeta/i
  },
  // v1.1.61: from real samples. Two senders: the app's receipts ("Recibo de la transacción": transfers sent and
  // TuEfectivo withdrawals) and notificaciones@ ("Notificaciones Banreservas": transfers received). Neither subject says
  // what happened, so the type is read from the body (detectType).
  BANRESERVAS: {
    name: 'BANRESERVAS',
    fromDomain: 'banreservas.com',
    searchQuery: 'from:notificaciones@banreservas.com OR from:NotificacionesTuBancoApp@banreservas.com',
    extractors: { consumo: extractBANRESERVASWithdrawalTransactions, transfer: extractBANRESERVASTransferTransactions,
      incoming: extractBANRESERVASIncomingTransactions },
    detectType: detectBANRESERVASType,
    strict: true,
    keywords: ['BANRESERVAS'],
    merchantPattern: /Destino:\s*([^,\n]{3,50})/i
  }
};

// v1.1.60: THE list of banks, in tab order (Raw_BDI before Raw_POPULAR, as asked in v1.1.5) — the Setup Wizard, View
// Config and the sheet order all come from it, so a bank is added in one place
const BANK_ORDER = ['LAFISE', 'BANESCO', 'BHD', 'BDI', 'POPULAR', 'SCOTIABANK', 'QIK', 'BANRESERVAS'];

/**
 * v1.1.61: a bank marked `strict` never goes to the generic amount guesser (extractAllAmounts). An email of its own that
 * none of its extractors recognizes (a welcome email, an account-opening notice, a template never seen) goes to
 * Unrecognized, where it can be looked at, instead of being saved as a purchase. Reported on BDI: transfers received
 * and a welcome email listing deposit limits were saved as spending.
 */
const STRICT_BANK_NOTE = 'Format not known yet: nothing guessed. Share its original email (.eml) to add it';

/**
 * v1.0.3: Emails to reject outright — marketing/promo sends from the bank,
 * never real consumption notifications. Checked against subject + body.
 * v1.0.8: removed the standalone 'GANA' entry — a real run showed it
 * matching randomly inside long tracking-parameter strings in HTML email
 * boilerplate (e.g. "z=1AvGy9akdi2..."), not real promotional text. Lesson:
 * keep entries here as full words/phrases of 6+ characters where possible;
 * short 4-letter fragments have real collision risk against URL noise.
 * 'PARTICIPA Y GANA' (the full phrase) is unaffected and stays.
 */
const PROMOTIONAL_KEYWORDS = [
  'PROMOCIÓN', 'PROMOCION', 'PROMO', 'OFERTA', 'DESCUENTO', 'NEWSLETTER',
  'BOLETÍN', 'BOLETIN', 'ENTÉRATE', 'ENTERATE', 'NUEVO BENEFICIO',
  'SORTEO', 'PARTICIPA Y GANA', 'DISFRUTA DE', 'CONOCE NUESTRO'
];

/**
 * v1.0.6: real bank notifications that aren't a single transaction — account
 * security (OTP, login, beneficiaries), periodic statements, and surveys/
 * service announcements. v1.0.8: widened past just "account notices" once
 * real samples showed "Estado de cuenta...", "Queremos conocer tu opinión",
 * and event-invite mail ("Acompáñanos en...") all needed the same treatment
 * — none of these represent one transaction, so none should reach
 * extractAmount(). Checked against subject + body, same as
 * PROMOTIONAL_KEYWORDS, for the same reason: a stray RD$ figure inside a
 * statement or survey email would otherwise create a phantom transaction.
 * v1.1.9: widened again — a real POPULAR credit-LIMIT-INCREASE notification
 * ("Actualización de Límite", subject confirmed from a live run) mentions
 * the new/old limit as a dollar figure and was getting parsed as if it were
 * a real purchase (saved rows showed amounts like 50,000/70,000 with
 * merchant "Unknown Merchant" or garbled limit-notice text) — same failure
 * shape as every other non-transactional email that slipped through before
 * its category was covered here.
 */
const NON_TRANSACTIONAL_KEYWORDS = [
  // Account / security
  'OTP', 'CÓDIGO DE VERIFICACIÓN', 'CODIGO DE VERIFICACION', 'CÓDIGO DE SEGURIDAD',
  'CODIGO DE SEGURIDAD', 'ALERTA DE ACCESO', 'ALERTA ACCESO', 'INICIO DE SESIÓN',
  'INICIO DE SESION', 'NUEVO DISPOSITIVO', 'BENEFICIARIO AGREGADO', 'BENEFICIARIO ELIMINADO',
  'CAMBIO DE CONTRASEÑA', 'CAMBIO DE CLAVE', 'RESTABLECER CONTRASEÑA', 'ACTUALIZACIÓN DE DATOS',
  'ACTUALIZACION DE DATOS', 'VERIFICACIÓN DE IDENTIDAD', 'VERIFICACION DE IDENTIDAD',
  'SOLICITUD OPCIONES DE MENÚ', 'SOLICITUD OPCIONES DE MENU',
  // v1.1.63: account notices with no money in them — LAFISE's Bancanet (user blocked or about to be, temporary password,
  // alias or password changed, duplicate session, unblocked) and BDI's online registration
  'USUARIO DE BANCANET', 'TU USUARIO SE BLOQUEE', 'CONTRASEÑA TEMPORAL', 'CONTRASENA TEMPORAL', 'CONTRASEÑA EN BANCANET',
  'ALIAS DEL USUARIO', 'SESIÓN DUPLICADA', 'SESION DUPLICADA', 'DESBLOQUEO DE USUARIO', 'HACE TIEMPO QUE NO TE VEMOS',
  'SOLICITUD DE REGISTRO', 'CÓDIGO DE ACTIVACIÓN', 'CODIGO DE ACTIVACION', 'FIRMA DE CONTRATO', 'CLICK & SIGN',
  // v1.1.52: bank announcements (LAFISE's Pagos al Instante schedule comes from the incoming-transfers sender)
  'HORARIO TRANSFERENCIAS', 'HORARIOS TRANSFERENCIAS', 'HORARIO DE TRANSFERENCIAS', 'NOTA INFORMATIVA',
  // v1.1.23: incoming payroll deposit notices (POPULAR "Notificación Depósito de
  // Nómina") are income, not spending — they were reported as "Could not parse".
  'DEPÓSITO DE NÓMINA', 'DEPOSITO DE NOMINA',
  // v1.1.62: Banreservas' payroll notice (the salary is set in the Setup Wizard, as with POPULAR's above), its loan
  // reminder, and BHD's purchase-validation code
  'PAGO NÓMINA', 'PAGO NOMINA', 'PAGO DE NÓMINA', 'PAGO DE NOMINA', 'NOTIFICACIÓN DE BALANCES', 'NOTIFICACION DE BALANCES',
  'CÓDIGO DE VALIDACIÓN', 'CODIGO DE VALIDACION',
  // Statements / periodic summaries — not a single transaction
  'ESTADO DE CUENTA', 'FONDO DE INVERSIÓN', 'FONDO DE INVERSION',
  // Surveys / service announcements / events
  'QUEREMOS CONOCER TU OPINIÓN', 'QUEREMOS CONOCER TU OPINION', 'ENCUESTA',
  'ACOMPÁÑANOS', 'ACOMPAÑANOS', 'FORMA MÁS CONVENIENTE', 'FORMA MAS CONVENIENTE',
  // v1.1.61: welcome / account-opening emails (BDI's lists the first deposit and the account's limits as amounts). A real
  // transaction that happens to say "bienvenido" is still read: its bank's extractor finds it (see parseEmailMessage)
  'BIENVENIDO', 'BIENVENIDA', 'PRIMER DEPÓSITO', 'PRIMER DEPOSITO', 'APERTURA DE CUENTA', 'APERTURA DE TU CUENTA',
  // Credit limit change notices — mention a dollar figure, but it's a limit, not a purchase
  'ACTUALIZACIÓN DE LÍMITE', 'ACTUALIZACION DE LIMITE', 'AUMENTO DE LÍMITE', 'AUMENTO DE LIMITE',
  'INCREMENTO DE LÍMITE', 'INCREMENTO DE LIMITE', 'NUEVO LÍMITE', 'NUEVO LIMITE',
  'LÍMITE ANTERIOR', 'LIMITE ANTERIOR', 'LÍMITE DE TU TARJETA', 'LIMITE DE TU TARJETA'
];

/**
 * v1.0.8: a real transaction alert can still describe a DECLINED charge —
 * confirmed from a real Banesco sample ("...ha sido rechazada. Motivo de la
 * declinación: TARJETA VENCIDA"). No money moved, so this must not be logged
 * as an expense even though it has a real RD$ amount and a real merchant.
 */
const DECLINED_KEYWORDS = [
  'HA SIDO RECHAZADA', 'FUE RECHAZADA', 'FUE RECHAZADO', 'DECLINADA', 'DECLINACIÓN',
  'DECLINACION', 'TRANSACCIÓN FALLIDA', 'TRANSACCION FALLIDA', 'NO APROBADA', 'NO FUE APROBADA'
];

/**
 * v1.1.4: Type values translated to English throughout the system (sheet
 * names, headers, category names, and these values were a mix of English UI
 * with Spanish data — now consistently English end to end).
 * Also fixes a real bug: bare 'CASHBACK' matched inside "VISA CLASICA
 * SUPERCASHBACK" — a real Banesco CARD PRODUCT NAME (confirmed from a live
 * email), not an actual cashback credit — which mistyped every purchase made
 * with that card as Type=Cashback instead of Transaction, and since Category
 * is only assigned when Type=Transaction, those rows silently got no
 * category at all (and stayed that way even after recategorizing, since
 * recategorization also respects the — wrong — stored Type). Removed the
 * bare keyword; kept only the more specific reward-language ones. No real
 * "you received cashback" sample has been seen yet, so this Type is
 * best-effort until one surfaces — it errs toward under-detecting Cashback
 * rather than risk another silent false-positive like this one.
 */
const TYPE_KEYWORDS = {
  'Card Payment': ['PAGO DE TARJETA', 'PAGO TC', 'PAGO REALIZADO A SU TARJETA', 'PAGO A TARJETA DE CREDITO'],
  'Cashback': ['DEVOLUCION', 'REEMBOLSO', 'REBATE', 'CASHBACK ACREDITADO', 'CASHBACK RECIBIDO'],
  'Transfer': ['TRANSFERENCIA ENVIADA', 'TRANSFERENCIA RECIBIDA', 'PAGOS AL INSTANTE', 'ACH', 'SWIFT'],
};

/**
 * v1.0.9: confirmed real subject lines, checked first (subject alone, exact
 * substring) since they're a much stronger signal than guessing body
 * phrasing. LAFISE's transfer subject includes the customer's name as a
 * prefix ("<Nombre>, ¡Transferencia exitosa!") so this matches on the fixed
 * suffix only. Keys are the English Type values (v1.1.4) — the keywords
 * themselves stay in Spanish since that's the actual language of the bank
 * emails being matched.
 */
const TYPE_SUBJECT_KEYWORDS = {
  // v1.1.51: money received. First: its body says "TRANSFERENCIA ... RECIBIDA" and "PAGOS AL INSTANTE", which are
  // Transfer keywords — without this, a received transfer would be read as one sent.
  'Incoming': ['TRANSFERENCIA ENTRANTE', 'PAGO AL INSTANTE RECIBIDO',   // v1.1.61: Scotiabank
    'NOTIFICACIÓN DE TRANSFERENCIA RECIBIDA', 'NOTIFICACION DE TRANSFERENCIA RECIBIDA'],   // v1.1.62: BANESCO
  'Transfer': [
    '¡TRANSFERENCIA EXITOSA!',              // LAFISE
    'NOTIFICACIÓN DE TRANSFERENCIA REALIZADA', 'NOTIFICACION DE TRANSFERENCIA REALIZADA', // BANESCO
    'TRANSACCIONES ENTRE MIS PRODUCTOS',    // BHD
    'TRANSACCIÓN INTERBANCARIA', 'TRANSACCION INTERBANCARIA',   // BDI (v1.1.60)
    'AVISO DE TRANSFERENCIA',                 // LAFISE online banking, to another bank (v1.1.63)
  ],
  // v1.1.17: BUG FIX — real confirmed LAFISE subject is "¡Realizaste un
  // pago a tu tarjeta LAFISE!" (active/informal phrasing — "you made a
  // payment"), which the old TYPE_KEYWORDS body-guesses ('PAGO REALIZADO A
  // SU TARJETA', passive/formal — "was made to your card") never matched.
  // These rows were getting Type=Transaction with Category defaulting to
  // Dining (the fallback), so paying off the LAFISE card bill was being
  // double-counted as a new expense in the Dashboard on top of the
  // original purchases that made up the balance.
  'Card Payment': [
    '¡REALIZASTE UN PAGO A TU TARJETA',     // LAFISE
    'NOTIFICACIÓN DE PAGO DE TARJETA', 'NOTIFICACION DE PAGO DE TARJETA',   // LAFISE, from bancanet@ (v1.1.61)
  ],
  // v1.1.19 (C3): confirmed CONSUMO subjects. Matching one of these pins the
  // email to Transaction (only a Cashback refinement from the body is still
  // allowed — see detectTransactionType()), so body-keyword guesses like
  // 'ACH' can never re-type a purchase again. This is also what lets
  // "🔁 Recategorize" repair rows that were already mis-typed (the real
  // CACHAREPA CHURCHILL row) — recategorize has the subject, not the body.
  'Transaction': [
    'SERVICIO DE ALERTA - NUEVO CONSUMO',                                   // LAFISE
    'ALERTA DE CONSUMO BANESCO',                                            // BANESCO
    'BHD NOTIFICACIÓN DE TRANSACCIONES', 'BHD NOTIFICACION DE TRANSACCIONES', // BHD
    'NOTIFICACIÓN DE CONSUMO', 'NOTIFICACION DE CONSUMO',                   // POPULAR
    'DETALLE DE TRANSACCION TARJETA', 'DETALLE DE TRANSACCIÓN TARJETA',     // LAFISE, second card template (v1.1.61)
    'RETIRO CON CÓDIGO CASH', 'RETIRO CON CODIGO CASH',                     // QIK (v1.1.61)
  ]
};

/**
 * Returns true if this email looks promotional rather than a real
 * consumption/transaction notification.
 */
function isPromotionalEmail(subject, bodyText) {
  const upper = (subject + ' ' + bodyText).toUpperCase();
  const matched = findMatchingKeyword(upper, PROMOTIONAL_KEYWORDS); // v1.1.19: shared engine (02_categorizer.gs)
  if (matched) {
    Logger.log("  ↳ matched promotional keyword: \"" + matched + "\"");
    return true;
  }
  return false;
}

/**
 * v1.0.8: renamed from isAccountNoticeEmail — scope widened past just
 * account/security notices (see NON_TRANSACTIONAL_KEYWORDS above).
 */
function isNonTransactionalEmail(subject, bodyText) {
  // v1.1.63: a reply or a forward is a conversation (with the bank's staff, or your own), never an alert
  if (/^\s*(RE|RV|FW|FWD|TR)\s*:/i.test(String(subject || ''))) {
    Logger.log('  ↳ a reply or forward — a conversation, not an alert');
    return true;
  }
  const upper = (subject + ' ' + bodyText).toUpperCase();
  const matched = findMatchingKeyword(upper, NON_TRANSACTIONAL_KEYWORDS); // v1.1.19: shared engine (02_categorizer.gs)
  if (matched) {
    Logger.log("  ↳ matched non-transactional keyword: \"" + matched + "\"");
    return true;
  }
  return false;
}

/**
 * v1.0.8: true if this is a real transaction alert describing a DECLINED /
 * failed charge — no money moved, so it must not be logged as an expense.
 */
function isDeclinedTransactionEmail(subject, bodyText) {
  // v1.1.63: LAFISE's "Aviso de transferencia" reports its result — "Estado: Error" means no money moved (reported: two
  // failed rent payments would have been saved as transfers)
  const failed = flatText(bodyText).match(/\bEstado:\s*(Error|Rechazad[ao]|Fallid[ao]|Cancelad[ao]|Denegad[ao])\b/i);
  if (failed) {
    Logger.log('  ↳ the transfer failed: "Estado: ' + failed[1] + '"');
    return true;
  }
  const upper = (subject + ' ' + bodyText).toUpperCase();
  const matched = findMatchingKeyword(upper, DECLINED_KEYWORDS); // v1.1.19: shared engine (02_categorizer.gs)
  if (matched) {
    Logger.log("  ↳ matched declined-transaction keyword: \"" + matched + "\"");
    return true;
  }
  return false;
}

/**
 * Classifies a transaction email into Transaction / Transfer / Card Payment / Cashback.
 * v1.0.9: checks confirmed subject lines first (TYPE_SUBJECT_KEYWORDS), then
 * falls back to body-keyword guessing (TYPE_KEYWORDS). Defaults to
 * "Transaction" (a card consumption) when nothing else matches.
 */
/**
 * v1.1.19 (C3): subject-only type detection. Returns null when the subject
 * isn't one of the confirmed ones. Used on its own by
 * recategorizeAllTransactions() (04_sheetsWriter.gs), which only has the
 * subject — it used to call detectTransactionType(subject, MERCHANT), i.e.
 * it fed the merchant name in as if it were the email body, so 'ACH' inside
 * "CACHAREPA" re-typed a real purchase as Transfer on every run (verified),
 * and any type that had been decided from the real body got reverted.
 */
function detectTypeFromSubject(subject) {
  const upper = String(subject || '').toUpperCase();
  for (let type in TYPE_SUBJECT_KEYWORDS) {
    if (findMatchingKeyword(upper, TYPE_SUBJECT_KEYWORDS[type])) return type;
  }
  return null;
}

function detectTransactionType(subject, bodyText) {
  const bySubject = detectTypeFromSubject(subject);
  if (bySubject === 'Transfer' || bySubject === 'Card Payment' || bySubject === 'Incoming') return bySubject;
  const upper = (String(subject || '') + ' ' + String(bodyText || '')).toUpperCase();
  if (bySubject === 'Transaction') {
    // Confirmed consumo template: the body may only refine it to Cashback
    // (a refund row in the same template) — never to Transfer/Card Payment.
    return findMatchingKeyword(upper, TYPE_KEYWORDS['Cashback']) ? 'Cashback' : 'Transaction';
  }
  // Unknown subject (BDI, new templates): body keywords, whole-word for
  // short ones like 'ACH' (see keywordMatches(), 02_categorizer.gs).
  for (let type in TYPE_KEYWORDS) {
    if (findMatchingKeyword(upper, TYPE_KEYWORDS[type])) return type;
  }
  return 'Transaction';
}


/**
 * v1.1.19 (C1): turns a user-facing inclusive date range into exact
 * [start, endExclusive) bounds in the script's time zone.
 * Why: Gmail's `before:` operator EXCLUDES the date it's given (Google's own
 * example `after:2004/04/16 before:2004/04/18` returns the 16th and 17th).
 * The old query used `before:<end date>`, so "Monitor Gmail Now" never saw
 * today's emails, every date-range run silently dropped its own end date,
 * and the 6 AM run lost the LAST DAY OF EVERY MONTH (the run on the 30th
 * stopped at the 29th; the run on the 1st searched `after:X before:X`, an
 * empty window). Bounds are sent to Gmail as epoch seconds, which Gmail
 * accepts and which also removes any ambiguity about which time zone Gmail
 * uses for plain yyyy/MM/dd dates. Accepts 'yyyy-MM-dd' or 'yyyy/MM/dd'.
 */
function buildRangeBounds(startStr, endStr) {
  const tz = Session.getScriptTimeZone();
  const norm = v => String(v).trim().replace(/\//g, '-');
  const start = Utilities.parseDate(norm(startStr), tz, 'yyyy-MM-dd');
  const endExclusive = Utilities.parseDate(norm(endStr), tz, 'yyyy-MM-dd');
  if (isNaN(start.getTime()) || isNaN(endExclusive.getTime())) {
    throw new Error('Invalid date range: ' + startStr + ' – ' + endStr);
  }
  if (start.getTime() > endExclusive.getTime()) {
    throw new Error('Start date (' + startStr + ') is after end date (' + endStr + ').');
  }
  endExclusive.setDate(endExclusive.getDate() + 1); // include the whole end day
  return { start: start, endExclusive: endExclusive };
}

function toEpochSeconds(date) {
  return Math.floor(date.getTime() / 1000);
}

/**
 * v1.1.19 (M5): paginated Gmail search. The date-range search used to make
 * ONE GmailApp.search(query, 0, 500) call, so a big quarterly backlog lost
 * everything past thread 500 without a word. Pages through results up to
 * `limit` and reports whether the cap was reached so the run summary can say
 * so. (The old rolling-window searchTransactionEmails() was removed here —
 * nothing had called it since v1.1.12.)
 */
const MAX_THREADS_PER_RUN = 2000;

function gmailSearchAll(query, limit) {
  const PAGE = 100;
  let threads = [];
  let start = 0;
  while (threads.length < limit) {
    const want = Math.min(PAGE, limit - threads.length);
    const batch = GmailApp.search(query, start, want);
    threads = threads.concat(batch);
    start += batch.length;
    if (batch.length < want) return { threads: threads, capped: false };
  }
  return { threads: threads, capped: true };
}

/**
 * Diagnostic helper (added v1.0.4, sender logging added v1.0.5). Run this
 * directly from the Apps Script editor — pick debugBankEmailSample in the
 * function dropdown, run it, then View > Logs — any time the monitor comes
 * back with 0 results, or to confirm BHD's/POPULAR's real sender (still
 * unconfirmed as of v1.0.5 — see the BANK_PATTERNS comment above). It
 * searches by bank name ONLY, no domain/keyword restriction, so you see
 * everything, including false positives.
 */
function debugBankEmailSample(daysBack = 30, perBank = 5) {
  const afterDate = Utilities.formatDate(
    new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  for (let bankName in BANK_PATTERNS) {
    const keyword = BANK_PATTERNS[bankName].keywords[0];
    const query = `${keyword} after:${afterDate}`;
    const threads = GmailApp.search(query, 0, perBank);
    Logger.log(`--- ${bankName} (query: "${query}") — ${threads.length} thread(s) ---`);
    threads.forEach(t => {
      const msg = t.getMessages()[0];
      Logger.log(`  [${msg.getDate()}] isUnread=${msg.isUnread()} | From: ${msg.getFrom()} | ${msg.getSubject()}`);
    });
  }
  Logger.log("Done. All 5 banks now have a confirmed sender (see BANK_PATTERNS above); " +
             "re-run this if a new promotional/marketing address starts showing up " +
             "under a bank you track.");
}

/**
 * v1.1.19 (E7): per-run counters, surfaced in the final summary alert —
 * before this, a filtered or unparseable email only ever showed up in the
 * execution log.
 */
function newParseStats() {
  return {
    messagesSeen: 0, outOfRange: 0, notOwnBank: 0, promotional: 0,
    nonTransactional: 0, declined: 0, amountNotFound: 0, parseErrors: 0, placeholders: 0, reversals: 0,
    unrecognized: [], readIds: [],  // v1.1.35: emails for the Unrecognized sheet, and emails read cleanly
    alreadySaved: 0,                 // v1.1.36: skipped without reading — already in Transactions
    filteredIds: []                  // v1.1.49: read, nothing to save (statements, promotions…) — for the read log
  };
}

/**
 * Extract transactions from email threads.
 * History: v1.0.4 stopped skipping read messages (dedup is the real guard);
 * v1.0.9 made parseEmailMessage() return several items per email; v1.1.5
 * added rawCustomRules so a personal rule can categorize any Type.
 * v1.1.19:
 *  - (M6) takes the run's `range` and skips messages outside it. Gmail
 *    matches THREADS, and it groups every BHD alert with the same subject
 *    into one thread, so one thread can hold months of messages — all of
 *    them were re-parsed on every run, eating into the 6-minute limit.
 *  - returns { transactions, stats, failedThreadIds } — failedThreadIds are
 *    threads with a message that couldn't be parsed, which
 *    markEmailsAsProcessed() now leaves UNREAD so they stay visible (M4).
 *  - the merged `customRules` argument is gone (see categorizeTransaction()).
 */
function extractTransactionsFromThreads(threads, rawCustomRules, range, opts) {
  opts = opts || {};
  const transactions = [];
  const stats = newParseStats();
  const failedThreadIds = new Set();
  // v1.1.36: an email already in Transactions isn't read again (reading is what takes time), and reading stops at
  // opts.deadline — the threads not reached stay unread and are picked up by the next run.
  const skipIds = opts.skipIds || new Set();
  const clock = opts.clock || (() => Date.now());
  const processedThreads = [];
  let stopped = null;

  for (let t = 0; t < threads.length; t++) {
    const thread = threads[t];
    if (opts.deadline && clock() > opts.deadline) {
      stopped = { remaining: threads.length - t, processed: t };
      Logger.log("⏸ Stopped reading at the time budget: " + stopped.remaining + " thread(s) left for the next run");
      break;
    }
    for (let message of thread.getMessages()) {
      if (range) {
        const d = message.getDate();
        if (d < range.start || d >= range.endExclusive) { stats.outOfRange++; continue; }
      }
      stats.messagesSeen++;
      // v1.1.55: statements are read again — each row is recognized by its reference, and a newer version may take more rows
      if (skipIds.has(message.getId()) && !isStatementEmail(message)) { stats.alreadySaved++; continue; }
      const result = parseEmailMessage(message, rawCustomRules, stats);
      if (result.status === 'failed') failedThreadIds.add(thread.getId());
      transactions.push(...result.items);
      // v1.1.35: what the Unrecognized sheet lists — not saved, or saved with an unreadable merchant
      const unreadable = result.status === 'ok' && result.items.some(it => it.merchant === GARBLED_PLACEHOLDER);
      if (result.status === 'failed' || unreadable) {
        stats.unrecognized.push({ id: message.getId(), date: message.getDate(), bank: result.bank || (result.items[0] || {}).bank || '',
          subject: message.getSubject() || '', reason: unreadable ? 'Saved — merchant unreadable (red row in Transactions)' : result.reason,
          snippet: result.snippet || (unreadable ? (result.items[0].description || '') : '') });
      } else if (result.status === 'ok' || result.status === 'filtered') {
        stats.readIds.push(message.getId());
        if (result.status === 'filtered') stats.filteredIds.push(message.getId());
      }
    }
    processedThreads.push(thread);
  }
  return { transactions: transactions, stats: stats, failedThreadIds: failedThreadIds, processedThreads: processedThreads, stopped: stopped };
}

/**
 * v1.1.19 (M1): true if the email has at least one APPROVED row. The
 * whole-email "declined" filter used to discard a multi-row BHD table
 * entirely if ANY row said "Declinada" — including the approved rows, which
 * the extractor already knows how to keep. Now the email-level filter only
 * fires when nothing in it was approved; row-level status is handled by the
 * extractors (they flag declined rows with `declined: true`).
 * ("NO FUE APROBADA"/"NO APROBADA" are removed first so they don't count.)
 */
function hasApprovedRow(text) {
  const cleaned = String(text || '').toUpperCase().replace(/\bNO\s+(FUE\s+)?APROBAD[AO]\b/g, '');
  return /\bAPROBAD[AO]\b/.test(cleaned);
}

/**
 * v1.1.19 (M3): IsCredit = money coming IN (an abono). It used to be YES
 * whenever the WHOLE email body mentioned "crédito" or "depósito" anywhere,
 * including boilerplate like "tu tarjeta de crédito". Now: Card Payment and
 * Cashback are credits by definition; otherwise only an explicit CRÉDITO /
 * DEPÓSITO inside THIS item's own context counts, and the product name
 * "TARJETA DE CRÉDITO" never does.
 */
function computeIsCredit(type, context) {
  if (type === 'Card Payment' || type === 'Cashback') return true;
  const upper = String(context || '').toUpperCase().replace(/TARJETA\s+DE\s+CR[ÉE]DITO/g, '');
  return /\b(CR[ÉE]DITO|DEP[ÓO]SITO)\b/.test(upper);
}

/**
 * v1.1.66: the email's text. Some emails carry an EMPTY plain-text part and everything in the HTML (LAFISE's
 * "¡Transferencia exitosa!", seen in a real .eml): getPlainBody() alone gave nothing and the email went to
 * Unrecognized as "Amount not found". Then the HTML is read as text. A plain part with any text is used as before.
 */
function emailPlainText(message) {
  let text = '';
  try { text = String(message.getPlainBody() || ''); } catch (e) { text = ''; }
  if (/\S/.test(text) || typeof message.getBody !== 'function') return text;
  try { return htmlToPlainText(message.getBody()) || text; } catch (e) { return text; }
}

const HTML_ENTITIES = { nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", aacute: 'á', eacute: 'é', iacute: 'í',
  oacute: 'ó', uacute: 'ú', ntilde: 'ñ', uuml: 'ü', Aacute: 'Á', Eacute: 'É', Iacute: 'Í', Oacute: 'Ó', Uacute: 'Ú', Ntilde: 'Ñ',
  Uuml: 'Ü', iexcl: '¡', iquest: '¿', ordm: 'º', ordf: 'ª', deg: '°', copy: '©', reg: '®', euro: '€', middot: '·', ndash: '–',
  mdash: '—', hellip: '…', laquo: '«', raquo: '»' };

/**
 * v1.1.66: HTML → text the way the extractors expect it from Gmail: one line per block or table cell, bold as *text*
 * (as getPlainBody() writes it), entities decoded; styles, scripts and comments dropped; no empty lines. Pure.
 */
function htmlToPlainText(html) {
  const code = n => (n > 0 && n <= 0x10FFFF ? String.fromCodePoint(n) : '');
  return String(html || '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<(style|script|head|title)\b[\s\S]*?<\/\1\s*>/gi, '')
    .replace(/<(b|strong)\b[^>]*>([\s\S]*?)<\/\1\s*>/gi, (m, tag, inner) => {
      const x = inner.replace(/<[^>]*>/g, '').trim();
      return x ? '*' + x + '*' : '';
    })
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/?(p|div|tr|td|th|li|ul|ol|table|tbody|thead|tfoot|h[1-6]|blockquote|center)\b[^>]*>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&#(\d+);/g, (m, n) => code(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (m, n) => code(parseInt(n, 16)))
    .replace(/&([a-z]+);/gi, (m, n) => (HTML_ENTITIES[n] !== undefined ? HTML_ENTITIES[n] : m))
    .split('\n').map(line => line.replace(/[ \t\r\f\v ]+/g, ' ').trim()).filter(Boolean).join('\n');
}

/**
 * Parse one email into zero or more transaction items.
 * History: v1.0.5 bank by sender domain; v1.0.3/v1.0.8 promotional,
 * non-transactional and declined filters; v1.0.9 array of items; v1.1.0
 * bank+type-specific extractors with the generic fallback; v1.1.5 Custom
 * Rules checked first for every Type; v1.1.13 messageId = Gmail id + index.
 * v1.1.19:
 *  - returns { items, status } — status is 'ok', 'filtered', 'skipped' or
 *    'failed', and `stats` (see newParseStats()) is updated in place.
 *  - (E2) categorizes on the MERCHANT, the same text "🔁 Recategorize" uses.
 *    It used to categorize on the raw match context here and on the
 *    Description there, so the two could disagree and every run rewrote
 *    categories back and forth.
 *  - (M1) declined rows flagged by an extractor are dropped individually;
 *    if every row was declined, the email is skipped WITHOUT falling back to
 *    extractAllAmounts() — that fallback used to save a declined BANESCO
 *    charge anyway, plus its "Dispones de RD$ …" balance as a second row.
 *  - (M3) IsCredit via computeIsCredit().
 */
function parseEmailMessage(message, rawCustomRules, stats) {
  stats = stats || newParseStats();
  try {
    const subject = message.getSubject() || '';
    const plainText = emailPlainText(message);   // v1.1.66: the HTML when the plain-text part is empty
    const fromAddress = message.getFrom() || ''; // e.g. "LAFISE CF <notificaciones@bancolafise.com>"

    // v1.0.5: bank by the sender's DOMAIN, not by the bank name appearing in
    // the text (Uber receipts mention "LAFISE" as the payment method).
    let bank = null;
    let bankPattern = null;
    for (let bankName in BANK_PATTERNS) {
      const pattern = BANK_PATTERNS[bankName];
      if (pattern.fromDomain && fromAddress.toLowerCase().includes(pattern.fromDomain.toLowerCase())) {
        bank = bankName;
        bankPattern = pattern;
        break;
      }
    }
    if (!bank) {
      stats.notOwnBank++;
      Logger.log("Bank not detected (sender domain not recognized): " + fromAddress + " | " + subject);
      return { items: [], status: 'skipped' };
    }

    // v1.1.51: a bank statement (PDF): the incoming transfers the bank didn't notify
    if (bankPattern.statement && bankPattern.statement.subject.test(subject)) {
      return parseStatementEmail(message, bank, bankPattern.statement, rawCustomRules, stats);
    }

    // v1.1.62: emails a bank sends that are not a movement of their own (a confirmation of one already notified, a
    // code created but not yet used). Declared per bank; filtered before anything else, so a rescue can't revive them.
    const ignored = (bankPattern.ignore || []).find(rule => rule.test(subject + ' ' + flatText(plainText)));
    if (ignored) {
      stats.nonTransactional++;
      Logger.log("Skipped (" + bank + ": not a movement of its own): " + subject);
      return { items: [], status: 'filtered' };
    }

    if (isPromotionalEmail(subject, plainText)) {
      stats.promotional++;
      Logger.log("Skipped promotional email: " + subject);
      return { items: [], status: 'filtered' };
    }
    if (isNonTransactionalEmail(subject, plainText)) {
      // v1.1.60: a footer like "we'll never ask for your card's security code" isn't a security-code email. When the
      // bank's own extractor finds a real transaction (an amount), the email is read; otherwise it's filtered.
      const bankType = bankTransactionType(bankPattern, subject, plainText);   // v1.1.61: the bank's own reading first
      const real = bankPattern.extractors && Object.keys(bankPattern.extractors).length &&
        extractTransactionItems(bank, bankType, plainText).some(x => x && x.amount > 0);
      if (!real) {
        stats.nonTransactional++;
        Logger.log("Skipped non-transactional email: " + subject);
        return { items: [], status: 'filtered' };
      }
      Logger.log("Kept: a real transaction whose footer only mentions a non-transactional phrase — " + subject);
    }
    if (isDeclinedTransactionEmail(subject, plainText) && !hasApprovedRow(plainText)) {
      stats.declined++;
      Logger.log("Skipped declined/failed transaction: " + subject);
      return { items: [], status: 'filtered' };
    }

    const date = formatDate(message.getDate());
    const type = bankTransactionType(bankPattern, subject, plainText);   // v1.1.61
    const emailType = type;   // v1.1.62: items may override it, one by one

    let items = extractTransactionItems(bank, type, plainText);
    items.forEach(it => { it.merchant = fixStatusAsMerchant(it.merchant, subject, plainText); });   // v1.1.34
    const declinedRows = items.filter(it => it.declined).length;
    items = items.filter(it => !it.declined);
    if (items.length === 0 && declinedRows > 0) {
      stats.declined++;
      Logger.log("Skipped declined transaction (row-level status): " + subject);
      return { items: [], status: 'filtered' };
    }
    if (items.length === 0 && !bankPattern.strict) {   // v1.1.61: a strict bank's unknown email is never guessed
      items = extractAllAmounts(plainText, bankPattern);
    }
    if (items.length === 0) {
      stats.amountNotFound++;
      const snippet = plainText.substring(0, 700).replace(/\s+/g, ' ').trim();
      Logger.log("Amount not found | " + bank + " | " + subject + " | Body: \"" + snippet + "\"");
      return { items: [], status: 'failed', bank: bank, reason: bankPattern.strict ? STRICT_BANK_NOTE : 'Amount not found', snippet: snippet };
    }

    const results = [];
    const baseMessageId = message.getId(); // v1.1.13: unique per email
    let itemIndex = 0;
    for (let item of items) {
      // v1.0.3: skip zero-amount lines (v1.1.19: and NaN, which used to slip through)
      if (!(item.amount > 0)) {
        Logger.log("Skipped $0.00 line within: " + subject);
        continue;
      }
      // v1.1.23: a reversal is saved as a NEGATIVE amount; its merchant and
      // category are filled in from the original purchase at save time
      // (resolveReversals(), 04_sheetsWriter.gs), so the two net to zero.
      const isReversal = !!item.reversal;
      if (isReversal) stats.reversals++;
      const merchant = isReversal ? REVERSAL_UNMATCHED : (String(item.merchant || '').trim() || 'Unknown Merchant');
      // v1.1.62: an item can carry its own type (the tax row of a card payment is a Transfer, not the payment itself)
      const type = item.type || emailType;
      let category = findCustomRuleOverride(merchant, rawCustomRules);
      if (!category && type === 'Transaction' && !isReversal) {   // a reversal takes its original's category
        category = categorizeTransaction(merchant);
      }
      // v1.1.51: money received is saved NEGATIVE, like a reversal — given a category (a Custom Rule, or by hand) it
      // reduces what you spent there; money from your own account is Exclude
      const isIncoming = type === 'Incoming';
      if (item.own && (isIncoming || type === 'Transfer')) category = EXCLUDE_CATEGORY;   // v1.1.63: sent, too
      // v1.1.23: paying the card is never spending — explicit "Exclude"
      if (type === 'Card Payment') category = EXCLUDE_CATEGORY;
      if (merchant === GARBLED_PLACEHOLDER) stats.placeholders++;
      const currency = item.currency || 'DOP';

      Logger.log("Parsed OK | " + bank + " | " + type + " | " + (category || '-') + " | " + currency + " " +
                 item.amount + " | merchant=\"" + merchant + "\" | Context: \"" + (item.context || '') + "\"");

      results.push({
        date: date,
        bank: bank,
        merchant: merchant,
        amount: isReversal || isIncoming ? -item.amount : item.amount,
        currency: currency,
        category: category,
        type: type,
        description: item.own && (isIncoming || type === 'Transfer') ? merchant + OWN_ACCOUNT_SUFFIX : merchant,
        reversal: isReversal,
        timeKey: item.timeKey || '',
        txRef: item.ref ? bank + ':' + item.ref : '',
        subject: subject,
        timestamp: new Date().toISOString(),
        messageId: baseMessageId + '_' + (itemIndex++),
        isCredit: isReversal || isIncoming || computeIsCredit(type, item.context),
        isCashback: type === 'Cashback'
      });
    }
    return { items: results, status: 'ok' };
  } catch (error) {
    stats.parseErrors++;
    Logger.log("Error parsing email: " + error + (error && error.stack ? " | " + error.stack : ""));
    let snippet = '';
    try { snippet = String(emailPlainText(message)).substring(0, 700).replace(/\s+/g, ' ').trim(); } catch (e) { snippet = ''; }
    return { items: [], status: 'failed', bank: '', reason: 'Could not read: ' + error, snippet: snippet };
  }
}

/**
 * v1.1.0: routes to the right bank+type-specific extractor, all written and
 * tested against real .eml samples (Sept 23, 2026). Each returns an array of
 * {amount, merchant, context}. BDI has no real sample yet, so it goes
 * straight to the generic fallback (extractAllAmounts, further below).
 */
/**
 * v1.1.1: detects USD vs DOP from the matched text itself.
 * v1.1.5: broadened to any 2-3 letter code, to catch a real COP transaction
 * — but that was TOO broad: a live run showed it capturing garbage like
 * "FKR", "WEC", "QAH", "JDH", "UCO", "BGJ" as "currency" — random uppercase
 * fragments from authorization codes, reference numbers, and other noise in
 * the email body that happened to precede a digit somewhere in the matched
 * context window.
 * v1.1.6: BUG FIX — reverted to a fixed, explicit list of the 4 currencies
 * actually relevant here (DOP, USD, EUR, COP), per request, rather than a
 * generic [A-Z]{3} wildcard. Extend this list deliberately if a genuine new
 * currency shows up in real data — don't widen it back to a wildcard.
 */
function detectCurrencyFromMatch(matchText) {
  // v1.1.63: "US" is dollars only as "US$" (or USD) — a LAFISE authorization code like "US34K7" read as US 12 made a
  // DOP purchase USD (reported, Uber Eats DOP 390)
  const m = matchText.match(/\b(RD|US(?=\s*\$)|DOP|USD|EUR|COP)\s*\$?\s*[\d,]/);
  if (!m) return 'DOP';
  if (m[1] === 'RD') return 'DOP';
  if (m[1] === 'US') return 'USD';
  return m[1];
}

/**
 * Dispatches to the bank+type extractor declared in BANK_PATTERNS.
 * History: v1.1.10 fixed LAFISE transfers always using the consumo
 * extractor (they had no extractor of their own, so they fell through to the
 * generic fallback and produced garbled merchants).
 * v1.1.19 (E4): table lookup. Anything that isn't a Transfer (Transaction,
 * Card Payment, Cashback) uses the bank's consumo extractor, as before; a
 * bank with no extractor (BDI) returns [] and parseEmailMessage() falls back
 * to extractAllAmounts().
 */
function extractTransactionItems(bank, type, text) {
  const extractors = (BANK_PATTERNS[bank] && BANK_PATTERNS[bank].extractors) || {};
  const fn = type === 'Transfer' ? extractors.transfer : type === 'Incoming' ? extractors.incoming :
    (type === 'Card Payment' && extractors.cardPayment) || extractors.consumo;   // v1.1.61: a card payment's own template
  return fn ? fn(text) : [];
}

/**
 * v1.1.61: the type of an email, read by its bank first. Some banks use one subject for different things (Banreservas'
 * "Recibo de la transacción" is a transfer or a withdrawal; BDI's "Comprobante transacción Interbancaria" is sent or
 * received), so a bank can declare detectType(subject, text); when it returns nothing, the shared rules decide.
 */
function bankTransactionType(bankPattern, subject, text) {
  const own = bankPattern && bankPattern.detectType ? bankPattern.detectType(subject, flatText(text)) : null;
  return own || detectTransactionType(subject, text);
}

/**
 * LAFISE consumo ("Servicio de Alerta - Nuevo Consumo"). Real sample:
 *   Comercio/Ciudad/País:
 *   UBER*EATS SANTO DOMINGO DOM
 *   ...
 *   Monto:
 *   DOP 412.86
 * Merchant comes BEFORE the amount here — the opposite of every other bank.
 * `[\s\S]*?` (not `.*?`) bridges the Fecha/Marca/Tarjeta/Autorización/
 * Referencia/Tipo fields in between regardless of the newlines they contain.
 */
function extractLAFISETransactions(text) {
  const re = /Comercio\/Ciudad\/País:\s*([^\n]+?)\s*\n[\s\S]*?Monto:\s*\**\s*(RD|DOP|USD|US\$|EUR|COP)?\s*\$?\s*([\d,]+\.?\d*)/gi;
  const results = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    // v1.1.62: the currency written next to the amount wins; the whole context (every field between the merchant and
    // the amount) is only the fallback when the amount has none
    const currency = m[2] ? detectCurrencyFromMatch(m[2] + ' 0') : detectCurrencyFromMatch(m[0]);
    results.push({ merchant: m[1].replace(/^\*+|\*+$/g, '').trim().substring(0, 50), amount: parseFloat(m[3].replace(/,/g, '')), currency: currency, context: m[0].replace(/\s+/g, ' ').trim() });
  }
  return results;
}

/**
 * LAFISE transferencia ("<Nombre>, ¡Transferencia exitosa!").
 * v1.1.10: STOPGAP — no real LAFISE transfer .eml has been seen yet, only
 * the confirmed subject and sender (digital@notificaciones.lafise.com). A
 * live run showed the OLD behavior (silently falling through to the
 * generic backward-looking fallback, since this Type was never dispatched
 * to a dedicated extractor at all — see extractTransactionItems() above)
 * producing garbled mid-sentence merchant text. This deliberately extracts
 * the FIRST currency-prefixed amount in the body ONLY (not every one —
 * safer against accidentally grabbing an unrelated balance/limit figure
 * mentioned elsewhere) and labels the merchant generically rather than
 * guessing at a beneficiary field whose label hasn't been confirmed to
 * exist in this template. Known limitation: won't split a same-day bundle
 * of multiple LAFISE transfers into separate rows (unconfirmed whether
 * LAFISE even bundles transfers the way BHD does). Share a real LAFISE
 * transfer .eml to replace this with a proper extractor, the same way
 * BANESCO/BHD/POPULAR were built from real samples.
 * v1.1.12: tries a handful of common Spanish reference-number labels
 * (Referencia / No. Referencia / Número de confirmación / Confirmación) and
 * appends whatever it finds to the placeholder — "LAFISE Transfer (ref:
 * XXXXX)" instead of the fully generic label, per request, so the row is at
 * least traceable back to the real transfer even without a confirmed
 * beneficiary field. Silently keeps the plain placeholder if none of these
 * labels match (still a guess at the real template).
 * v1.1.66: a real .eml is in hand: "Acabas de realizar una transferencia de USD 75.50 entre tus cuentas. … De paso te
 * dejamos tu número de referencia: <n>", all in the HTML (the plain-text part is empty — see emailPlainText()). The
 * amount and reference rules above read it; "entre tus cuentas" makes it yours (Exclude). Only a transfer between
 * your own accounts has been seen: one to someone else is read the same way, without a beneficiary.
 */
function extractLAFISETransferTransactions(text) {
  // v1.1.19: first amount that is NOT an available-balance figure.
  const amtRe = /\b(?:RD|DOP|USD|US\$|EUR|COP)\s*\$?\s*([\d,]+\.?\d*)/gi;
  let m = null, cand;
  while ((cand = amtRe.exec(text)) !== null) {
    if (!isBalanceAmount(text, cand.index)) { m = cand; break; }
  }
  if (!m) return [];
  const amount = parseFloat(m[1].replace(/,/g, ''));
  if (amount === 0) return [];

  let merchant = 'LAFISE Transfer (details unconfirmed)';
  const refMatch = text.match(/(?:No\.?\s*Referencia|Referencia|Número de [Cc]onfirmaci[oó]n|Confirmaci[oó]n)\s*[:\-]?\s*([A-Za-z0-9\-\.]{4,30})/i);
  if (refMatch) {
    merchant = `LAFISE Transfer (ref: ${refMatch[1]})`;
  }

  return [{
    merchant: merchant,
    amount: amount,
    currency: detectCurrencyFromMatch(m[0]),
    context: text.substring(0, 400).replace(/\s+/g, ' ').trim(),
    ref: refMatch ? refMatch[1] : '',   // v1.1.23: bank's unique id → duplicate guard
    own: /\bentre tus cuentas\b/i.test(flatText(text))   // v1.1.66: "…una transferencia de USD 75.50 entre tus cuentas."
  }];
}

/**
 * BANESCO consumo, approved ("Alerta de Consumo Banesco RD"). Real sample:
 *   "...presenta un consumo de RD$ 2,640.00, en SM BRAVO LA ESPERILLA y su
 *    estado es aprobada."
 * Amount comes BEFORE the merchant here. A declined variant exists with
 * different wording ("...ha sido rechazada...") that this pattern doesn't
 * match at all (no "y su estado es APROBADA/status" phrase in that
 * template) — declined ones are also caught by the whole-email
 * isDeclinedTransactionEmail() check upstream, this is defense-in-depth.
 */
function extractBANESCOConsumoTransactions(text) {
  const re = /consumo de\s*(?:RD|DOP|USD|US\$|EUR|COP)?\s*\$?\s*([\d,]+\.?\d*),?\s*en\s+([^\n]+?)\s+y su estado es\s+(\w+)/gi;
  const results = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    const declined = /RECHAZ|DECLIN/i.test(m[3]); // v1.1.19: flagged, not dropped (M1)
    results.push({ declined: declined, merchant: m[2].trim().substring(0, 50), amount: parseFloat(m[1].replace(/,/g, '')), currency: detectCurrencyFromMatch(m[0]), context: m[0].replace(/\s+/g, ' ').trim() });
  }
  return results;
}

/**
 * BANESCO transferencia ("Notificación de Transferencia Realizada"). Real
 * sample: "Monto: DOP545.00" ... later ... "Nombre del Beneficiario: AURORA
 * BEATRIZ PEREZ FERNANDEZ". Uses the beneficiary's name as "merchant" —
 * Category stays blank for Type=Transferencia regardless (Regla 3), this is
 * just for the description/record.
 */
function extractBANESCOTransferTransactions(text) {
  const re = /Monto:\s*(?:RD|DOP|USD|US\$|EUR|COP)?\s*\$?\s*([\d,]+\.?\d*)[\s\S]*?Nombre del Beneficiario:\s*([^\n]+)/gi;
  const results = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    // v1.1.23: the bank's own unique id ("No. Referencia: E000073.A1195451") — the same
    // transfer notice arrived twice in one thread, with identical content.
    const refM = text.substring(m.index, m.index + 800).match(/No\.?\s*Referencia:\s*([A-Za-z0-9.\-]{4,40})/i);
    results.push({ merchant: m[2].trim().substring(0, 50), amount: parseFloat(m[1].replace(/,/g, '')), currency: detectCurrencyFromMatch(m[0]), context: m[0].replace(/\s+/g, ' ').trim(), ref: refM ? refM[1] : '' });
  }
  return results;
}

/**
 * BHD consumo ("BHD Notificación de Transacciones") — a real table:
 *   Fecha | Moneda | Monto | Comercio | Estado | Tipo   (header, once)
 *   01/09/2026 08:38 am / RD / $118.40 / FCIA MEDICAR... / Aprobada / Compra  (one row per transaction)
 * The `g` flag naturally picks up every repeating 6-line row block, so this
 * is the one most likely to correctly handle the same-day-bundling case the
 * user described, once a real multi-row sample confirms the row layout
 * repeats the same way for row 2, 3, etc.
 */
/**
 * BHD consumo ("BHD Notificación de Transacciones") — a table (Fecha /
 * Moneda / Monto / Comercio / Estado / Tipo), one row per transaction.
 * v1.1.11: BUG FIX — required a literal `\n` between every field, which
 * broke on a real failing sample: "08/04/2026 10:20 am US $7.85 LIBERTARIO
 * COFFEE ROASTER Aprobada Compra" — fields separated by spaces, not
 * newlines (also the first confirmed sample with currency "US" rather than
 * "RD"). Switched every field separator from `\n` to `\s+` (matches either),
 * and anchored the boundary between the merchant and the following fields
 * to the known status words ("Aprobada"/"Rechazada"/"Declinada") instead of
 * a bare `\w+` — needed because a non-greedy merchant capture bordered by
 * `\s+` alone (ambiguous with the spaces INSIDE a multi-word merchant name
 * like "LIBERTARIO COFFEE ROASTER") would otherwise stop too early and
 * misattribute merchant words to Status/Type. Re-verified against both
 * previously-working real samples plus this new one — all three still
 * extract correctly.
 */
const BHD_ROW_STATUSES = 'Aprobada|Rechazada|Declinada|Reversada|Reversado|Reverso';

/**
 * v1.1.23: a REVERSAL row (real .eml, 06/05/2026: the same table row a few
 * seconds later with Estado "Reversada" and an EMPTY Comercio cell) used to
 * match nothing here — the status list didn't include it and the merchant
 * was mandatory — so the generic fallback saved it as a brand-new RD$488
 * "Unknown Merchant" expense. Now: the merchant is optional (and can never
 * swallow a status word, so a one-line multi-row body can't merge rows), and
 * a reversed row comes back with `reversal: true` plus `timeKey` — the
 * table's own date+time, identical on the purchase and its reversal — so
 * saveTransactions() can pair them (see resolveReversals()).
 */
function extractBHDConsumoTransactions(text) {
  const merchant = `((?:(?!\\b(?:${BHD_ROW_STATUSES})\\b)[^\\n])+?)`;
  const re = new RegExp(`(\\d{2}\\/\\d{2}\\/\\d{4}[^\\n]*?)\\s+(RD|DOP|USD?|EUR|COP)\\s+\\$?\\s*([\\d,]+\\.?\\d*)` +
                        `(?:\\s+${merchant})?\\s+(${BHD_ROW_STATUSES})\\s+(\\w+)`, 'gi');
  const results = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    const status = m[5];
    results.push({
      declined: /RECHAZ|DECLIN/i.test(status), // v1.1.19: flagged, not dropped (M1)
      reversal: /REVERS/i.test(status),
      timeKey: m[1].replace(/\s+/g, ' ').trim(),
      merchant: (m[4] || '').trim().substring(0, 50),
      amount: parseFloat(m[3].replace(/,/g, '')),
      currency: detectCurrencyFromMatch(m[2] + ' ' + m[0]),
      context: m[0].replace(/\s+/g, ' ').trim()
    });
  }
  return results;
}

/**
 * BHD transferencia ("Transacciones entre mis productos"). Real single-
 * transaction sample:
 *   Monto:
 *   RD$ 91.37
 *   Beneficiario:
 *   <ACCOUNT HOLDER NAME>
 * v1.1.5: a real multi-transaction BHD transfer email (several transfers
 * bundled same-day, confirmed from a live screenshot) broke this — the
 * garbled merchant values saved (": *DO60BCBH0000...", "ta desde. EC:MQ...")
 * show the [\s\S]*? bridge jumping to the WRONG "Beneficiario:" across
 * transaction boundaries when "Producto origen/destino" fields repeat
 * between entries. Added a sanity check that rejects an obviously-broken
 * capture (looks like an account number / "Producto" label fragment rather
 * than a name) instead of saving garbage — this loses those rows rather
 * than mislabeling them, which is the safer failure mode until a real
 * multi-transaction sample is available to fix the pattern properly.
 */
/**
 * v1.1.10: shared detector for an obviously-broken merchant/beneficiary
 * capture — an account-number fragment, a field label like "Producto
 * destino:", or a stray mid-sentence continuation — rather than a real
 * name. Real merchant/beneficiary text from every bank seen so far is
 * always UPPERCASE (UBER*EATS, SM BRAVO, FCIA MEDICAR, AURORA BEATRIZ
 * PEREZ...), so a capture starting lowercase is itself a strong signal of
 * a broken capture, not just the bank-specific fragments already known.
 * Used both to reject a capture live (BHD/LAFISE transfer extractors) and
 * to clean up already-saved rows (see cleanGarbledMerchants() below).
 * v1.1.18: BUG FIX — removed the 'HOLA ' substring check. It was added to
 * catch the fragment "Hola <name>, Acabas de realiza..." (a greeting
 * leaking into a broken capture), but it was a false-positive magnet: a
 * real BANESCO merchant "HOLA PLAZA LAS AMERICAS" (a real business inside
 * Plaza Las Américas mall) was correctly extracted by
 * extractBANESCOConsumoTransactions(), then NUKED into the placeholder by
 * this check during the very next auto-recategorize pass — which runs at
 * the end of EVERY monitor run, so this silently corrupted the row
 * immediately after it was saved correctly, not just once. Confirmed the
 * original fragment this rule was for is still caught without it — it
 * starts lowercase (`^[a-z]`) AND contains a bracket, both independently
 * sufficient. 'ACABAS DE' alone was already redundant with those too but
 * is left as-is (harmless, no known false-positive risk).
 */
const GARBLED_PLACEHOLDER = '(unparsed — see Email Subject)'; // v1.1.19: moved here from 04_sheetsWriter.gs

const REVERSAL_UNMATCHED = 'Reversal (original purchase not found)'; // v1.1.23

/**
 * v1.1.34: a POPULAR "Código Cash" withdrawal has no merchant column (Monto | Moneda | Fecha | Estatus), so the
 * status word after the date — "Aprobada" — was saved as the merchant (seen in a live Raw_POPULAR). A status word is
 * never a merchant: a Código Cash email gets a clear label, anything else the unreadable-merchant placeholder.
 * The category stays the fallback, Dining/Delivery + Entertainment + Other.
 */
const STATUS_WORD_MERCHANT = /^(APROBAD[AO]|RECHAZAD[AO]|DECLINAD[AO]|REVERSAD[AO]|REVERSO)$/i;
const CODIGO_CASH_MERCHANT = 'Código Cash (cash withdrawal)';
function fixStatusAsMerchant(merchant, subject, text) {
  if (!STATUS_WORD_MERCHANT.test(String(merchant || '').trim())) return merchant;
  return /C[OÓ]DIGO\s*CASH/i.test(String(subject || '') + ' ' + String(text || '')) ? CODIGO_CASH_MERCHANT : GARBLED_PLACEHOLDER;
}

function looksGarbled(text) {
  if (!text) return false;
  const t = String(text).trim();
  if (t.length === 0) return false;
  if (/^[a-z]/.test(t)) return true;       // real bank text is never lowercase-led
  if (/[\[\]]/.test(t)) return true;       // stray bracket from a truncated annotation
  if (/^[:*]/.test(t)) return true;
  const upper = t.toUpperCase();
  if (upper.includes('PRODUCTO') || upper.startsWith('TA DESDE') || /\*DO\d/.test(t) ||
      upper.includes('ACABAS DE')) return true;
  return false;
}

/**
 * BHD transferencia ("Transacciones entre mis productos"). Real confirmed
 * per-transaction field order:
 *   Producto origen: <acct> / Producto destino: <acct> / Descripción: /
 *   Monto: RD$ X / Beneficiario: <name> / Número de confirmación: ... /
 *   Fecha: ... / Tipo de transacción: ...
 * v1.1.15: BUG FIX — the garble rejection used `continue`, silently
 * dropping the item; if that was the only match, `results` came back
 * EMPTY, and parseEmailMessage() falls back to the generic
 * extractAllAmounts() (no garble protection), which re-extracted equally-
 * broken text. Fixed by never discarding — always keep the transaction.
 * v1.1.16: BUG FIX/CHANGE — a live log showed "Monto: *RD$ 2,950.00" with
 * an asterisk directly before the currency (Markdown-style bold from an
 * HTML→text conversion) that the amount regex didn't account for at all,
 * on top of the known multi-transaction-bundling issue. Also, garbled
 * captures strongly suggested some bundled transactions might be INTERNAL
 * transfers between the user's own BHD accounts, with no "Beneficiario:"
 * field at all. Rewritten to: (1) find every "Monto:" occurrence with its
 * own position, tolerating the asterisk; (2) for each, look FORWARD
 * (bounded to before the NEXT "Monto:") for its own "Beneficiario:"; (3) if
 * none is found there, look BACKWARD (bounded to after the PREVIOUS
 * "Monto:") for its own "Producto destino:" account — which structurally
 * comes before "Monto:", not after — and use its last 4 digits as a
 * traceable identifier instead of a name.
 * v1.1.18: CORRECTION from a real single-transaction .eml (self-transfer
 * between the user's own BHD accounts) — the v1.1.16 theory that an
 * internal/self transfer has NO "Beneficiario:" field was wrong. It DOES
 * have one; BHD just fills it with the account holder's own name
 * ("Beneficiario: <ACCOUNT HOLDER NAME>") rather than omitting the field. Verified
 * this real sample extracts correctly as-is (forward-look finds
 * "Beneficiario:" immediately, backward-look branch never triggers) — no
 * code change needed here, just correcting the inaccurate assumption in the
 * comment above so it doesn't mislead future debugging. The backward
 * "Producto destino:" fallback branch stays as defense-in-depth for
 * whatever templates it WAS built from, but its premise is unconfirmed —
 * still no real bundled multi-transaction .eml seen to verify that branch
 * or the field-order-repeats-per-row assumption against.
 */
function extractBHDTransferTransactions(text) {
  const results = [];
  const montoRe = /Monto:\s*\*?\s*(?:RD|DOP|USD|US\$|EUR|COP)?\s*\$?\s*([\d,]+\.?\d*)\*?/gi;
  let m;
  const matches = [];
  while ((m = montoRe.exec(text)) !== null) {
    matches.push({
      amount: parseFloat(m[1].replace(/,/g, '')),
      currency: detectCurrencyFromMatch(m[0]),
      index: m.index,
      end: m.index + m[0].length
    });
  }
  for (let i = 0; i < matches.length; i++) {
    const forwardEnd = (i + 1 < matches.length) ? matches[i + 1].index : text.length;
    const backwardStart = (i > 0) ? matches[i - 1].index : 0;
    const forwardBlock = text.substring(matches[i].end, forwardEnd);
    const backwardBlock = text.substring(backwardStart, matches[i].index);

    let merchant;
    const benefM = forwardBlock.match(/Beneficiario:\s*([^\n]+)/i);
    if (benefM && !looksGarbled(benefM[1].trim())) {
      merchant = benefM[1].trim().substring(0, 50);
    } else {
      // "Producto destino:" comes BEFORE "Monto:" for the same transaction
      // — take the LAST occurrence in the backward block (closest to this
      // transaction's own "Monto:"), not the first, so a bundle doesn't
      // pick up an earlier transaction's account instead of this one's.
      const destMatches = [...backwardBlock.matchAll(/Producto destino:\s*\*?[X\d]*?(\d{4})\D*?(?=\n|Descripci|Monto|$)/gi)];
      const destM = destMatches.length > 0 ? destMatches[destMatches.length - 1] : null;
      merchant = destM ? `BHD Transfer (to account ...${destM[1]})` : GARBLED_PLACEHOLDER;
    }
    const confM = forwardBlock.match(/N[úu]mero de confirmaci[óo]n:\s*([A-Za-z0-9\-]{4,40})/i); // v1.1.23
    results.push({
      merchant: merchant,
      amount: matches[i].amount,
      currency: matches[i].currency,
      context: text.substring(Math.max(0, matches[i].index - 100), Math.min(text.length, matches[i].end + 100)).replace(/\s+/g, ' ').trim(),
      ref: confM ? confM[1] : ''
    });
  }
  return results;
}

/**
 * POPULAR consumo ("Notificación de Consumo") — a compact table on one
 * visual row: Monto | Moneda | Fecha | Comercio | Estatus, e.g.
 *   US$14.27 [tab] Dólar estadounidense [tab] 22/05/2026 [tab] ELDORADO [tab] Aprobada
 * Lower confidence than the others — the real column separators (tabs vs.
 * spaces vs. newlines) may render slightly differently than in the one
 * sample seen; every parse is still logged with its context so this can be
 * tightened from a live run if needed.
 */
/**
 * POPULAR consumo ("Notificación de Consumo") — a compact table on one
 * visual row: Monto | Moneda | Fecha | Comercio | Estatus.
 * v1.1.9: BUG FIX — the currency prefix only accepted RD/DOP/USD/US$; a
 * real POPULAR sample in Colombian Pesos ("COP$33,000.00") never matched at
 * all, so the whole transaction silently disappeared (not just mis-tagged —
 * genuinely never extracted). This is what looked like "transactions within
 * a thread not being returned" — the COP-denominated messages in a thread
 * were failing outright while DOP/USD ones in the same thread succeeded.
 * Added EUR/COP to match the same 4-currency set used everywhere else.
 */
function extractPOPULARTransactions(text) {
  const re = /(?:RD|DOP|USD|US\$|EUR|COP)\s*\$?\s*([\d,]+\.?\d*)[^\n]*?\d{2}\/\d{2}\/\d{4}[^\n]*?[\t ]([A-ZÁÉÍÓÚÑ][^\t\n]{2,40}?)\s*[\t\n]+\s*(\w+)/gi;
  const results = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    const declined = /RECHAZ|DECLIN/i.test(m[3]); // v1.1.19: flagged, not dropped (M1)
    results.push({ declined: declined, merchant: m[2].trim().substring(0, 50), amount: parseFloat(m[1].replace(/,/g, '')), currency: detectCurrencyFromMatch(m[0]), context: m[0].replace(/\s+/g, ' ').trim() });
  }
  return results;
}

/**
 * v1.0.9: generic fallback — finds EVERY currency amount in the body and
 * pairs each with its best-effort PRECEDING context for merchant extraction.
 * Used only for BDI (no real sample yet) and as a safety net when a
 * bank-specific extractor above finds nothing. Note this only looks
 * BACKWARD from each amount, which is wrong for BANESCO/BHD/POPULAR (where
 * merchant comes AFTER the amount) — that's exactly why those three now have
 * dedicated extractors instead of relying on this.
 */
function extractAllAmounts(text, bankPattern) {
  const amountRegex = /\b(?:RD|DOP|USD|US\$|EUR|COP)\s*\$?\s*([\d,]+\.?\d*)/gi;
  const rawMatches = [];
  let m;
  while ((m = amountRegex.exec(text)) !== null) {
    if (isBalanceAmount(text, m.index)) continue; // v1.1.19: "Dispones de RD$ …" is a balance, not a charge
    rawMatches.push({ amount: parseFloat(m[1].replace(/,/g, '')), start: m.index, end: amountRegex.lastIndex });
  }

  const items = [];
  for (let i = 0; i < rawMatches.length; i++) {
    const cur = rawMatches[i];
    const windowStart = i === 0 ? Math.max(0, cur.start - 150) : rawMatches[i - 1].end;
    const context = text.substring(windowStart, cur.end).replace(/\s+/g, ' ').trim();
    const merchant = extractMerchantFromContext(context, bankPattern);
    items.push({ amount: cur.amount, merchant: merchant, currency: detectCurrencyFromMatch(context), context: context });
  }
  return items;
}

/**
 * v1.0.9: merchant extraction scoped to a single transaction's context
 * window (see extractAllAmounts) instead of the whole email body — this
 * also fixes part of the merchant="<first name>" bug, since the customer's name
 * in the greeting is now outside most items' windows (still inside the
 * FIRST item's window if the amount appears early in the body; unconfirmed
 * until a real sample shows exactly where).
 */
function extractMerchantFromContext(context, bankPattern) {
  try {
    const match = context.match(bankPattern.merchantPattern);
    if (match) {
      return match[1].trim().substring(0, 50);
    }
    const words = context.match(/\b[A-Z]{3,}\b/g);
    if (words && words.length > 0) {
      return words[0];
    }
    return "Unknown Merchant";
  } catch (error) {
    Logger.log("Error extracting merchant: " + error);
    return "Unknown";
  }
}

/**
 * Format date consistently
 */
/**
 * v1.1.2: BUG FIX — used to return date.toLocaleDateString('es-DO'), a plain
 * TEXT string. Google Sheets sometimes auto-detects that as a date and
 * sometimes doesn't (depends on the exact string + sheet locale), which is
 * exactly the kind of inconsistency that breaks SUMIFS/QUERY date-range
 * formulas — the Dashboard sheet depends on Date being a real date value in
 * every row. Now returns an actual Date object.
 * v1.1.5: BUG FIX — that Date object was built at MIDNIGHT
 * (new Date(y,m,d)), which is timezone-fragile: the Apps Script PROJECT's
 * time zone (Project Settings) and the SPREADSHEET's own time zone (File >
 * Settings) are two separate settings, and if they don't match, midnight in
 * one can land on the PREVIOUS calendar day when displayed in the other —
 * exactly what caused the Dashboard's month headers to read Dec-25...Nov-26
 * instead of Jan-26...Dec-26 (every date effectively shifted back). Building
 * at NOON instead absorbs any reasonable timezone mismatch (safe up to a
 * 12-hour gap) without crossing a day boundary. The real, permanent fix is
 * aligning the two timezone settings so they match (Project Settings > Time
 * zone = the same zone as the spreadsheet, e.g. America/Santo_Domingo) —
 * this is just a defensive floor under that.
 * NOTE: rows saved before this fix may still have a text date or a
 * midnight-based date — isDuplicate() normalizes both sides through
 * Utilities.formatDate() specifically so old and new rows still compare
 * correctly for duplicate-checking, but a date-range SUMIFS on the
 * Dashboard may still misplace a pre-fix row by one day if the two
 * timezones actually disagree. Re-entering that row's Date cell (or running
 * "🔁 Recategorize Saved Transactions", which does not touch dates, so this
 * still requires manual correction) can fix a specific bad row if noticed.
 */
function formatDate(date) {
  if (typeof date === 'string') {
    date = new Date(date);
  }
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0);
}

/**
 * v1.1.19: true when the amount at `index` is an available balance or limit
 * ("Dispones de RD$ 25,410.08", "Balance disponible", "Límite …"), not a
 * charge. The generic fallback and the LAFISE transfer extractor used to
 * pick these up as transactions.
 */
function isBalanceAmount(text, index) {
  const before = String(text).substring(Math.max(0, index - 40), index).toUpperCase();
  return /(DISPON|SALDO|BALANCE|L[ÍI]MITE)[^\n]{0,25}$/.test(before);
}

/**
 * Mark processed threads as read and label them.
 * v1.1.19 (M4):
 *  - threads with a message that could NOT be parsed (failedThreadIds) are
 *    left UNREAD — marking them read used to hide exactly the emails that
 *    need a look;
 *  - the "Procesado" label is fetched once (it used to be looked up inside
 *    the loop) and created if missing (it was never created, so it was never
 *    applied);
 *  - batched: GmailApp.markThreadsRead() / label.addToThreads() take up to
 *    100 threads per call instead of 2 calls per thread.
 */
const PROCESSED_LABEL = "Procesado";

function markEmailsAsProcessed(threads, failedThreadIds, opts) {
  opts = opts || {};
  const failed = failedThreadIds || new Set();
  let done = threads.filter(t => !failed.has(t.getId()));
  // v1.1.44: every thread in the range used to be marked again — read and labelled — even those marked by earlier
  // runs; over a year-long range that was minutes of Gmail calls (reported: a run killed right after "Parsed 1
  // transactions"). Gmail lists the ones still unread or without the label; only those are marked.
  let alreadyMarked = 0;
  if (opts.query && done.length) {
    try {
      const pending = new Set(gmailSearchAll(opts.query + ' {is:unread -label:' + PROCESSED_LABEL + '}', MAX_THREADS_PER_RUN)
        .threads.map(t => t.getId()));
      const before = done.length;
      done = done.filter(t => pending.has(t.getId()));
      alreadyMarked = before - done.length;
    } catch (error) {
      Logger.log("Could not list the threads still to mark (marking them all): " + error);
    }
  }
  const clock = opts.clock || (() => Date.now());
  let label = null;
  try {
    label = GmailApp.getUserLabelByName(PROCESSED_LABEL) || GmailApp.createLabel(PROCESSED_LABEL);
  } catch (error) {
    Logger.log("Could not get/create the \"" + PROCESSED_LABEL + "\" label: " + error);
  }
  let marked = 0;
  for (let i = 0; i < done.length; i += 100) {
    if (opts.deadline && clock() > opts.deadline) {   // v1.1.44: the rest stays for the next run
      Logger.log("⏸ Stopped marking at the time budget: " + (done.length - i) + " thread(s) left to mark");
      break;
    }
    const chunk = done.slice(i, i + 100);
    marked += chunk.length;
    try {
      GmailApp.markThreadsRead(chunk);
      if (label) label.addToThreads(chunk);
    } catch (error) {
      Logger.log("Error marking emails: " + error);
    }
  }
  const keptUnread = threads.filter(t => failed.has(t.getId())).length;
  Logger.log("Marked " + marked + " thread(s) as processed (" + alreadyMarked + " already were); left " + keptUnread +
             " unread because something in them failed to parse.");
  return { marked: marked, alreadyMarked: alreadyMarked, keptUnread: keptUnread };
}

/**
 * Search bank emails in an INCLUSIVE date range ('yyyy-MM-dd' or 'yyyy/MM/dd').
 * v1.1.19 (C1, M5): exact epoch bounds from buildRangeBounds() (the end date
 * is now actually included) and paginated via gmailSearchAll(). Returns
 * { threads, range, capped } — `range` is passed on so messages outside it
 * can be skipped (see extractTransactionsFromThreads()). Search errors are no
 * longer swallowed into an empty result — they propagate to the entry point,
 * which shows them.
 * getUnreadTransactionCount() was removed (M13 — never called).
 */
function searchTransactionEmailsByDateRange(startDate, endDate, banksToTrack) {
  const range = buildRangeBounds(startDate, endDate);
  const selectedBanks = Object.keys(banksToTrack || {}).filter(b => banksToTrack[b] && BANK_PATTERNS[b]);
  const fromClauses = selectedBanks.map(b => BANK_PATTERNS[b].searchQuery).filter(Boolean);
  if (fromClauses.length === 0) {
    Logger.log("No bank in banksToTrack has a confirmed searchQuery — nothing to search.");
    return { threads: [], range: range, capped: false };
  }
  const query = `(${fromClauses.join(" OR ")}) after:${toEpochSeconds(range.start)} before:${toEpochSeconds(range.endExclusive)}`;
  Logger.log("Search query: " + query + "  [" + startDate + " … " + endDate + ", both days included]");
  const result = gmailSearchAll(query, MAX_THREADS_PER_RUN);
  Logger.log("Found " + result.threads.length + " thread(s)" +
             (result.capped ? " — hit the " + MAX_THREADS_PER_RUN + "-thread cap, split the range" : ""));
  return { threads: result.threads, range: range, capped: result.capped, query: query,   // v1.1.44: query, to mark only what's pending
           untracked: findUntrackedBanks(selectedBanks, range) };                        // v1.1.61
}

/**
 * v1.1.61: banks NOT ticked in the Setup Wizard that do have emails in this range (one quick search each, one result
 * at most). Reported: Scotiabank purchases never showed up while the same email was read correctly in memory. A bank
 * added in a new version starts unticked in a setup saved before (v1.1.60), so its emails are never searched; the run
 * summary now says so instead of staying silent. Never stops the run.
 */
function findUntrackedBanks(selectedBanks, range) {
  const out = [];
  BANK_ORDER.filter(b => selectedBanks.indexOf(b) === -1 && BANK_PATTERNS[b] && BANK_PATTERNS[b].searchQuery).forEach(b => {
    try {
      const q = '(' + BANK_PATTERNS[b].searchQuery + ') after:' + toEpochSeconds(range.start) + ' before:' + toEpochSeconds(range.endExclusive);
      if (GmailApp.search(q, 0, 1).length) out.push(b);
    } catch (error) {
      Logger.log('Could not check ' + b + ' (not ticked): ' + error);
    }
  });
  if (out.length) Logger.log('⚠️ Emails found from banks not ticked in the Setup Wizard: ' + out.join(', '));
  return out;
}

/* ======================================================================
 * INCOMING TRANSFERS — v1.1.51
 * Money received: saved as NEGATIVE rows of Type "Incoming". With a category (a Custom Rule on the sender's name, or
 * typed in Incoming Transfers) it reduces what you spent in that category — a roommate's share of the rent lowers
 * Rent. Money from your own account is Exclude. Nothing else is decided for you.
 * ====================================================================== */
const OWN_ACCOUNT_SUFFIX = ' (own account)';

/** Two person names are the same holder when their first two names match (accents, case and signs aside — a
 *  statement writes "Ñ" as a space and cuts long names). Pure. */
function normalizedWords(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/[^A-Z ]/g, ' ').split(/\s+/).filter(Boolean);
}

function sameHolder(a, b) {
  const x = normalizedWords(a), y = normalizedWords(b);
  return x.length >= 2 && y.length >= 2 && x[0] === y[0] && x[1] === y[1];
}

/**
 * LAFISE "TRANSFERENCIA ENTRANTE APLICADA EXITOSAMENTE." (Pagos al Instante). Real sample (values changed):
 *   Nombre del cliente: / <HOLDER> / Número de cuenta: / … / Ordenante: / <SENDER> <cédula> / Monto total: / DOP 4500.00
 */
function extractLAFISEIncomingTransactions(text) {
  const flat = String(text || '').replace(/\s+/g, ' ');
  const m = flat.match(/Monto total:\s*(DOP|USD|RD\$|US\$)?\s*([\d,]+\.\d{2})/i);
  if (!m) return [];
  const sender = ((flat.match(/Ordenante:\s*(.+?)\s*(?:\d{6,}\s*)?Monto total:/i) || [])[1] || '').replace(/\s+\d+$/, '').trim();
  const holder = ((flat.match(/Nombre del cliente:\s*(.+?)\s*N[úu]mero de cuenta:/i) || [])[1] || '').trim();
  const currency = /USD|US\$/i.test(m[1] || '') ? 'USD' : 'DOP';
  return [{ amount: Number(m[2].replace(/,/g, '')), currency: currency, merchant: sender || 'Incoming transfer',
    own: !!(sender && holder && sameHolder(sender, holder)), context: 'incoming transfer' }];
}

/** A statement PDF (an attachment's blob) as text — read by the tracker itself (09_pdfText.gs), no Drive API.
 *  Apps Script gives the bytes signed (-128…127). */
function statementPdfText(blob) {
  const raw = blob.getBytes(), bytes = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) bytes[i] = raw[i] & 255;
  return pdfToText(bytes);
}

/**
 * Banesco savings statement (text of its PDF) → { holder, currency, rows, incoming, problems }. Each row is
 * "dd/mm/yyyy <description> <amount> <balance>"; whether it's a debit or a credit is read from the balance (the column
 * isn't in the text). Everything must add up to the statement's own totals, or nothing is taken from it. Pure.
 */
function parseBanescoSavingsStatement(text) {
  const flat = String(text || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ');
  const num = s => Number(String(s).replace(/,/g, ''));
  const total = re => { const m = flat.match(re); return m ? num(m[1]) : null; };
  const out = { holder: '', currency: /US\$|D[óo]lares|Moneda:\s*USD/i.test(flat) && !/RD\$/.test(flat) ? 'USD' : 'DOP', rows: [], incoming: [], problems: [] };
  const opening = total(/Balance mes anterior:?\s*(-?[\d,]+\.\d{2})/i), closing = total(/Balance al corte:?\s*(-?[\d,]+\.\d{2})/i);
  const credits = total(/Cr[ée]ditos del mes:?\s*([\d,]+\.\d{2})/i), debits = total(/D[ée]bitos del mes:?\s*([\d,]+\.\d{2})/i);
  const holder = flat.match(/(?:Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+\d{4}\s+([A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ .]{4,}?)\s+Cuenta/i)
    || flat.match(/Detalle de tus transacciones\s+([A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ .]{4,}?)\s+Transacciones en/i);
  out.holder = holder ? holder[1].trim() : '';
  if (opening === null || closing === null || credits === null || debits === null) {
    out.problems.push('the statement\'s totals (balances, credits, debits) were not found');
    return out;
  }
  const rowRe = /(\d{2})\/(\d{2})\/(\d{4}) (.+?) (-?[\d,]+\.\d{2}) (-?[\d,]+\.\d{2})(?= \d{2}\/\d{2}\/\d{4}| |$)/g;
  let prev = opening, sumC = 0, sumD = 0, m;
  while ((m = rowRe.exec(flat)) !== null) {
    const amount = num(m[5]), balance = num(m[6]), delta = +(balance - prev).toFixed(2);
    let direction = null;
    if (Math.abs(delta - amount) < 0.01) direction = 'credit';
    else if (Math.abs(delta + amount) < 0.01) direction = 'debit';
    if (!direction) { out.problems.push('row ' + m[1] + '/' + m[2] + ' "' + m[4] + '" does not follow the balance'); prev = balance; continue; }
    if (direction === 'credit') sumC += amount; else sumD += amount;
    out.rows.push({ day: m[3] + '-' + m[2] + '-' + m[1], description: m[4].trim(), amount: amount, balance: balance, direction: direction });
    prev = balance;
  }
  if (!out.rows.length) out.problems.push('no transactions found');
  if (Math.abs(sumC - credits) > 0.01) out.problems.push('credits add up to ' + sumC.toFixed(2) + ', the statement says ' + credits.toFixed(2));
  if (Math.abs(sumD - debits) > 0.01) out.problems.push('debits add up to ' + sumD.toFixed(2) + ', the statement says ' + debits.toFixed(2));
  if (out.rows.length && Math.abs(prev - closing) > 0.01) out.problems.push('the last balance is ' + prev.toFixed(2) + ', the statement says ' + closing.toFixed(2));
  if (out.problems.length) return out;
  // v1.1.55: EVERY credit is money received — descriptions vary ("Ach Ibanking", "Lbtr <name>", deposits, interest…);
  // what each one pays back is the user's call (a Custom Rule on its description, or typed in Incoming Transfers).
  // Yours when it's an LBTR from the holder or its description names the holder.
  const holderKey = out.holder ? normalizedWords(out.holder).slice(0, 2).join(' ') : '';
  out.rows.filter(r => r.direction === 'credit').forEach(r => {
    const lbtr = /^LBTR\b/i.test(r.description), ach = /^ACH\b/i.test(r.description);
    const name = lbtr ? r.description.replace(/^LBTR\s+/i, '').trim() : '';
    const own = !!out.holder && ((lbtr && sameHolder(name, out.holder)) || (holderKey.length > 3 && (' ' + normalizedWords(r.description).join(' ') + ' ').indexOf(' ' + holderKey + ' ') !== -1));
    out.incoming.push({ day: r.day, amount: r.amount, balance: r.balance, currency: out.currency, description: r.description,
      merchant: name || (ach ? 'ACH transfer (sender not in the statement)' : r.description), own: own });
  });
  return out;
}

/** A statement email: its PDF's incoming transfers as Incoming items (one per row, deduplicated per statement row). */
function parseStatementEmail(message, bank, statement, rawCustomRules, stats) {
  const subject = message.getSubject() || '';
  const pdf = (message.getAttachments() || []).find(a => /pdf/i.test(a.getContentType() || '') || /\.pdf$/i.test(a.getName() || ''));
  if (!pdf) return { items: [], status: 'failed', bank: bank, reason: 'Statement without a PDF', snippet: subject };
  let parsed;
  try {
    parsed = statement.parse(statementPdfText(pdf.copyBlob()));
  } catch (error) {
    stats.parseErrors++;
    return { items: [], status: 'failed', bank: bank, reason: 'Statement not read: ' + error, snippet: subject };
  }
  if (parsed.problems.length) {
    stats.parseErrors++;
    return { items: [], status: 'failed', bank: bank, reason: "Statement doesn't add up — nothing taken from it: " + parsed.problems.slice(0, 2).join('; '),
      snippet: subject + ' · ' + parsed.rows.length + ' row(s) read' };
  }
  if (!parsed.incoming.length) return { items: [], status: 'filtered' };   // nothing received by transfer this month
  const baseMessageId = message.getId();
  const items = parsed.incoming.map(x => {
    let category = findCustomRuleOverride(x.merchant, rawCustomRules);
    if (x.own) category = EXCLUDE_CATEGORY;
    const p = x.day.split('-').map(Number);
    return { date: new Date(p[0], p[1] - 1, p[2], 12), bank: bank, merchant: x.merchant, amount: -x.amount, currency: x.currency,
      category: category || '', type: 'Incoming', description: x.own ? x.merchant + OWN_ACCOUNT_SUFFIX : x.merchant,
      reversal: false, timeKey: '', txRef: bank + ':STMT:' + x.day + ':' + x.amount.toFixed(2) + ':' + x.balance.toFixed(2),
      subject: subject, timestamp: new Date().toISOString(), isCredit: true, isCashback: false,
      // v1.1.55: the row itself, not its position — positions moved once every credit was taken
      messageId: baseMessageId + '_' + x.day.replace(/-/g, '') + '-' + Math.round(x.amount * 100) + '-' + Math.round(x.balance * 100) };
  });
  Logger.log('Statement read | ' + bank + ' | ' + parsed.rows.length + ' rows, ' + items.length + ' incoming transfer(s)');
  return { items: items, status: 'ok' };
}

/** v1.1.55: a bank statement email (one this tracker reads: its bank has a statement reader and the subject matches). */
function isStatementEmail(message) {
  const subject = message.getSubject() || '';
  return Object.keys(BANK_PATTERNS).some(b => BANK_PATTERNS[b].statement && BANK_PATTERNS[b].statement.subject.test(subject));
}

/* ======================================================================
 * BDI, SCOTIABANK, QIK — v1.1.60 (real samples; fixtures with invented data)
 * Each reads the text flattened to single spaces, so it doesn't matter how Gmail's plain-text version splits the
 * email's table cells into lines.
 * ====================================================================== */
/**
 * v1.1.62: the text flattened to single spaces, WITHOUT Gmail's bold markers. getPlainBody() writes bold as *text*
 * ("*COMERCIO: *UBER*EATS …", "*RD$ 20.00*"), and that broke every extractor below on real emails while their
 * fixtures (written without the markers) passed: QIK purchases were filtered silently and LAFISE's second template and
 * BDI's transfers received went to Unrecognized. An asterisk that touches a space or the edge of the text is a marker
 * and goes; one inside a word stays (UBER*EATS, PedidosYa*Market, a masked 53****1234).
 */
function flatText(text) {
  return String(text || '').replace(/\s+/g, ' ').replace(/(^|\s)\*+/g, '$1').replace(/\*+(?=\s|$)/g, '').replace(/\s+/g, ' ').trim();
}
function moneyNumber(s) { return parseFloat(String(s).replace(/,/g, '')); }

function moneyCurrency(token, fallback) {
  const t = String(token || '').toUpperCase().replace(/\s+/g, '');
  if (t === 'RD$' || t === 'DOP') return 'DOP';
  if (t === 'US$' || t === 'USD') return 'USD';
  if (t === 'EUR' || t === '€') return 'EUR';
  return fallback || 'DOP';
}

/** BDI "Notificacion de Consumos": a table — Fecha | Moneda | Monto | Comercio | Estado — with one or more rows. */
function extractBDIConsumoTransactions(text) {
  const flat = flatText(text);   // v1.1.62: without Gmail's bold markers
  const re = /(\d{2}\/\d{2}\/\d{2,4}\s+\d{1,2}:\d{2})\s+(RD\$|US\$|DOP|USD|EUR)\s*([\d,]+\.\d{2})\s+(.+?)\s+(APROBADA|RECHAZADA|DECLINADA|REVERSADA|ANULADA)\b/gi;
  const out = [];
  let m;
  while ((m = re.exec(flat)) !== null) {
    if (m[5].toUpperCase() !== 'APROBADA') continue;   // only charges that went through
    out.push({ amount: parseFloat(m[3].replace(/,/g, '')), currency: moneyCurrency(m[2]), merchant: m[4].trim().substring(0, 50),
      timeKey: m[1], context: m[0] });
  }
  return out;
}

/**
 * BDI "Comprobante transacción Interbancaria": a transfer SENT ("[Salida]") to another bank — the amount, the
 * beneficiary's name and, as its own row, the tax and commission when there are any. Anything else (an incoming one
 * would say "[Entrada]") is left for Unrecognized rather than guessed.
 */
function extractBDITransferTransactions(text) {
  const flat = flatText(text);   // v1.1.62: without Gmail's bold markers
  if (!/\[\s*Salida\s*\]/i.test(flat)) return [];
  const amt = flat.match(/Monto\s+(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i);
  if (!amt) return [];
  const currency = moneyCurrency(amt[1]);
  const who = (flat.match(/Beneficiario\s+\**\d*\s*(.+?)\s+Banco Destino/i) || [])[1];
  const bank = (flat.match(/Banco Destino(?:\s*\/\s*C[óo]digo Swift)?\s+(.+?)\s+[A-Z0-9]{8,11}\s+Monto/i) || [])[1];
  const fee = s => { const f = flat.match(s); return f ? parseFloat(f[2].replace(/,/g, '')) : 0; };
  const fees = fee(/Impuesto[^$]*?(RD|US)\$\s*([\d,]+\.\d{2})/i) + fee(/Comisi[óo]n\s+(RD|US)\$\s*([\d,]+\.\d{2})/i);
  const out = [{ amount: parseFloat(amt[2].replace(/,/g, '')), currency: currency,
    merchant: (who || 'Transferencia interbancaria').trim().substring(0, 50), context: (bank ? 'to ' + bank.trim() : 'interbank transfer') }];
  if (fees > 0) out.push({ amount: +fees.toFixed(2), currency: currency, merchant: 'BDI — impuesto y comisión de transferencia', context: 'transfer fees' });
  return out;
}

/** Scotiabank "Autorización …": "por un monto de $25.50 USD en <merchant> con su Tarjeta de Crédito Scotiabank ***1234". */
function extractSCOTIABANKConsumoTransactions(text) {
  const flat = flatText(text);   // v1.1.62: without Gmail's bold markers
  const m = flat.match(/por un monto de\s+(RD\$|US\$|\$)?\s*([\d,]+\.\d{2})\s*(USD|DOP|EUR)?\s+en\s+(.+?)\s+con su\s+Tarjeta/i);
  if (!m) return [];
  const currency = m[3] ? moneyCurrency(m[3]) : (m[1] === 'US$' ? 'USD' : 'DOP');
  return [{ amount: parseFloat(m[2].replace(/,/g, '')), currency: currency, merchant: m[4].trim().substring(0, 50), context: m[0] }];
}


/* ======================================================================
 * v1.1.61: more formats, from real samples shared by another user (the fixtures keep their structure with invented
 * data). Same approach as above: each reads the text flattened to single spaces.
 * ====================================================================== */

/** LAFISE card purchases: the original template ("Comercio/Ciudad/País:"), or the second one below. */
function extractLAFISEAnyConsumo(text) {
  const first = extractLAFISETransactions(text);
  return first.length ? first : extractLAFISECardDetailTransactions(text);
}

/**
 * LAFISE "Detalle de Transaccion Tarjeta de Crédito" (notificacioneslafisedo@lafise.com.do):
 *   COMERCIO: <merchant> MONTO: 1,234.56 PESOS DOMI FECHA: dd/mm/yyyy
 * The currency comes as a word after the amount, cut short ("PESOS DOMI"). A currency word that isn't recognized is not
 * guessed: nothing is returned and the email goes to Unrecognized.
 */
function extractLAFISECardDetailTransactions(text) {
  const flat = flatText(text);
  const m = flat.match(/COMERCIO:\s*(.+?)\s*MONTO:\s*(RD\$|US\$|DOP|USD|EUR)?\s*([\d,]+\.\d{2})\s*(.*?)\s*FECHA:/i);
  if (!m) return [];
  let currency = m[2] ? moneyCurrency(m[2]) : null;
  if (!currency) {
    const word = m[4].toUpperCase();
    if (/^PESOS?\s*DOM|^RD|^DOP/.test(word)) currency = 'DOP';
    else if (/^D[OÓ]LAR|^US/.test(word)) currency = 'USD';
    else if (/^EURO/.test(word)) currency = 'EUR';
    else return [];
  }
  return [{ amount: moneyNumber(m[3]), currency: currency, merchant: m[1].trim().substring(0, 50), context: m[0] }];
}

/**
 * LAFISE "Notificación de pago de tarjeta de crédito" (bancanet@notificaciones.lafise.com): Concepto, Monto, Estado and
 * the card's last digits. Only a payment marked Exitoso counts; any other state is treated as declined.
 */
function extractLAFISECardPaymentTransactions(text) {
  const flat = flatText(text);
  const m = flat.match(/Concepto:\s*(.+?)\s*Monto:\s*(RD\$|US\$|DOP|USD|EUR)\s*([\d,]+\.\d{2})/i);
  if (!m) return [];
  const card = (flat.match(/N[úu]mero de tarjeta de cr[ée]dito:\s*(\S+)/i) || [])[1] || '';
  const state = (flat.match(/Estado:\s*(\S+)/i) || [])[1] || '';
  const ref = (flat.match(/Referencia:\s*(\d{4,})/i) || [])[1] || '';
  return [{ amount: moneyNumber(m[3]), currency: moneyCurrency(m[2]), merchant: (m[1].trim() + (card ? ' ' + card : '')).substring(0, 50),
    declined: !!state && !/^EXITOS[AO]$/i.test(state), ref: ref ? 'PAY:' + ref : '', context: 'card payment' }];
}

/**
 * BDI "Comprobante Transacción Interbancaria Recibida": who paid (Pagado Por), the amount. Yours when the payer and the
 * beneficiary are the same person.
 */
function extractBDIIncomingTransactions(text) {
  const flat = flatText(text);
  if (!/Interbancaria\s+Recibida/i.test(flat)) return [];
  const amt = flat.match(/Monto\s+(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i);
  if (!amt) return [];
  const payer = ((flat.match(/Pagado Por\s+(.+?)\s+Fecha Transacci[óo]n/i) || [])[1] || '').trim();
  const beneficiary = ((flat.match(/Nombre Beneficiario\s+(.+?)\s+No\. Autorizaci[óo]n/i) || [])[1] || '').replace(/[\s.]+$/, '').trim();
  const ref = (flat.match(/No\. Referencia\s+(\d+)/i) || [])[1] || '';
  return [{ amount: moneyNumber(amt[2]), currency: moneyCurrency(amt[1]), merchant: (payer || 'Interbank transfer received').substring(0, 50),
    own: !!(payer && beneficiary && sameHolder(payer, beneficiary)), ref: ref ? 'IN:' + ref : '', context: 'incoming transfer' }];
}

/**
 * Scotiabank "Pago al Instante recibido": "Ha recibido un crédito a su cuenta ***1234 por un Pago al Instante realizado
 * desde Banco X por valor de $1,250.00 DOP." The email names the sending bank, not the person.
 */
function extractSCOTIABANKIncomingTransactions(text) {
  const flat = flatText(text);
  const m = flat.match(/recibido un cr[ée]dito a su cuenta.*?(?:desde\s+(.+?)\s+)?por valor de\s+(RD\$|US\$|\$)?\s*([\d,]+\.\d{2})\s*(USD|DOP|EUR)?/i);
  if (!m) return [];
  const currency = m[4] ? moneyCurrency(m[4]) : (m[2] === 'US$' ? 'USD' : 'DOP');
  const how = /Pago al Instante/i.test(m[0]) ? 'Pago al Instante' : 'Crédito recibido';
  return [{ amount: moneyNumber(m[3]), currency: currency, merchant: (how + (m[1] ? ' desde ' + m[1].trim() : '')).substring(0, 50),
    own: false, context: 'incoming transfer' }];
}

/**
 * QIK. A card purchase ("Se hizo una transacción de RD$ 640.00 en <merchant> con tu tarjeta"), or a Código CASH
 * withdrawal ("El Código CASH … ha sido utilizado con éxito. Monto RD$ 2,000.00 Estatus Exitoso"). The email also shows
 * the card's available balance, never taken for the amount.
 */
function extractQIKConsumoTransactions(text) {
  const flat = flatText(text);
  const m = flat.match(/transacci[óo]n de\s+(RD\$|US\$|DOP|USD|EUR)\s*([\d,]+\.\d{2})\s+en\s+(.+?)\s+con tu tarjeta/i);
  if (m) {
    const when = (flat.match(/Fecha y hora\s+(\d{2}-\d{2}-\d{4}\s+\d{1,2}:\d{2}\s*[AP]M)/i) || [])[1] || '';
    return [{ amount: moneyNumber(m[2]), currency: moneyCurrency(m[1]), merchant: m[3].trim().substring(0, 50), timeKey: when, context: m[0] }];
  }
  if (/C[óo]digo CASH\b.*?utilizado/i.test(flat)) {   // v1.1.61
    const cash = flat.match(/Monto\s+(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i);
    if (!cash) return [];
    const state = (flat.match(/Estatus\s+(\S+)/i) || [])[1] || '';
    return [{ amount: moneyNumber(cash[2]), currency: moneyCurrency(cash[1]), merchant: CODIGO_CASH_MERCHANT,
      declined: !!state && !/^EXITOS[AO]$/i.test(state), context: 'cash withdrawal' }];
  }
  return [];
}

/* ---------- Banreservas (v1.1.61) ---------- */
/** The type from the body: "Transferencia Recibida" is money received; "Transacción: Transferencia …" one sent;
 *  "Transacción: Retiro …" a cash withdrawal. Anything else: not decided here (and, strict, not guessed). */
function detectBANRESERVASType(subject, flat) {
  if (/Transferencia Recibida/i.test(flat)) return 'Incoming';
  if (/Transacci[óo]n:\s*Transferencia/i.test(flat)) return 'Transfer';
  if (/Transacci[óo]n:\s*Retiro/i.test(flat)) return 'Transaction';
  return null;
}

function banreservasAmount(flat) {
  const m = flat.match(/Monto:\s*(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i);
  return m ? { amount: moneyNumber(m[2]), currency: moneyCurrency(m[1]) } : null;
}

/** Tax and commission, when the receipt has them (an empty "Comisión:" is nothing). */
function banreservasFees(flat) {
  const fee = re => { const f = flat.match(re); return f ? moneyNumber(f[1]) : 0; };
  return +(fee(/Comisi[óo]n:\s*(?:RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i) + fee(/Impuestos?:\s*(?:RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i)).toFixed(2);
}

function banreservasRef(flat) {
  return (flat.match(/N[úu]mero de (?:transacci[óo]n|referencia):\s*(\d{6,})/i) || [])[1] || '';
}

/** "Recibo de la transacción" with "Transacción: Transferencia a Tercero": the beneficiary (Destino, before the comma,
 *  without the "SR"/"SRA" the bank adds), and the tax and commission as their own row when not zero. */
function extractBANRESERVASTransferTransactions(text) {
  const flat = flatText(text);
  if (!/Transacci[óo]n:\s*Transferencia/i.test(flat)) return [];
  const amt = banreservasAmount(flat);
  if (!amt) return [];
  const who = ((flat.match(/Destino:\s*([^,]+?)\s*,/i) || [])[1] || '').replace(/^(SR|SRA|SRTA)\.?\s+/i, '').trim();
  const kind = ((flat.match(/Transacci[óo]n:\s*(.+?)\s+Origen:/i) || [])[1] || 'Transferencia').trim();
  const ref = banreservasRef(flat);
  const out = [{ amount: amt.amount, currency: amt.currency, merchant: (who || kind).substring(0, 50), ref: ref, context: kind }];
  const fees = banreservasFees(flat);
  if (fees > 0) out.push({ amount: fees, currency: amt.currency, merchant: 'Banreservas: impuesto y comisión de transferencia', context: 'transfer fees' });
  return out;
}

/** "Recibo de la transacción" with "Transacción: Retiro TuEfectivo" (cash sent to a phone): a cash withdrawal. The
 *  phone number is not kept. */
function extractBANRESERVASWithdrawalTransactions(text) {
  const flat = flatText(text);
  const kind = ((flat.match(/Transacci[óo]n:\s*(Retiro.*?)\s+Origen:/i) || [])[1] || '').trim();
  if (!kind) return [];
  const amt = banreservasAmount(flat);
  if (!amt) return [];
  const out = [{ amount: amt.amount, currency: amt.currency, merchant: (kind + ' (cash withdrawal)').substring(0, 50), ref: banreservasRef(flat), context: kind }];
  const fees = banreservasFees(flat);
  if (fees > 0) out.push({ amount: fees, currency: amt.currency, merchant: 'Banreservas: impuesto y comisión de retiro', context: 'withdrawal fees' });
  return out;
}

/** "Notificaciones Banreservas" / "Transferencia Recibida": who sent it (Origen). The email doesn't name the account
 *  holder, so whether it's your own money is left to you (a Custom Rule with your name → Exclude). */
function extractBANRESERVASIncomingTransactions(text) {
  const flat = flatText(text);
  if (!/Transferencia Recibida/i.test(flat)) return [];
  const amt = banreservasAmount(flat);
  if (!amt) return [];
  const sender = ((flat.match(/Origen:\s*(.+?)\s+Banco Origen:/i) || [])[1] || '').trim();
  return [{ amount: amt.amount, currency: amt.currency, merchant: (sender || 'Transferencia recibida').substring(0, 50), own: false,
    context: 'incoming transfer' }];
}

/* ======================================================================
 * v1.1.62: BDI card payment from the account; BANESCO transfer received.
 * ====================================================================== */
/**
 * BDI "Comprobante de Transacción" with "Tipo de Transacción Pago Tarjetas de Crédito" (real sample, Unrecognized):
 *   Monto Transferido RD$4,000.00 … Impuesto 0.20% RD$8.00 Monto RD$4,008.00
 * The amount paid to the card is a Card Payment (Exclude: it's the purchases already counted, being paid); the tax is
 * a cost of its own, a separate row typed Transfer like the other banks' transfer taxes.
 */
function extractBDICardPaymentTransactions(text) {
  const flat = flatText(text);
  if (!/Pago Tarjetas? de Cr[ée]dito/i.test(flat)) return [];
  const paid = flat.match(/Monto Transferido\s+(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i) || flat.match(/\bMonto\s+(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i);
  if (!paid) return [];
  const currency = moneyCurrency(paid[1]);
  const fee = re => { const f = flat.match(re); return f ? moneyNumber(f[1]) : 0; };
  const fees = +(fee(/Impuesto[^$]*?(?:RD|US)\$\s*([\d,]+\.\d{2})/i) + fee(/Comisi[óo]n\s+(?:RD|US)\$\s*([\d,]+\.\d{2})/i)).toFixed(2);
  const ref = (flat.match(/No\.\s*Ref(?:erencia)?\.?\s+(\d{4,})/i) || [])[1] || '';
  const out = [{ amount: moneyNumber(paid[2]), currency: currency, merchant: 'Pago Tarjetas de Crédito (BDI)', ref: ref ? 'PAY:' + ref : '',
    context: 'card payment' }];
  if (fees > 0) out.push({ amount: fees, currency: currency, merchant: 'BDI: impuesto y comisión de pago de tarjeta', type: 'Transfer', context: 'payment fees' });
  return out;
}

/**
 * BANESCO "Notificación de Transferencia Recibida" (reported: saved as a transfer sent, i.e. spending). No sample of
 * its own yet; the sent one ("… Realizada") is labeled fields ("Monto: DOP1,000.00 … Nombre del Beneficiario: X
 * Concepto: LBTR <sender>"), and the received one was read by that same extractor (the amount and the beneficiary
 * came out right), so the same labels are read here. Who sent it: a sender field if there is one, else the name in
 * "Concepto: LBTR <name>"; yours when it's the beneficiary. Without either, the sender is not guessed.
 */
function extractBANESCOIncomingTransactions(text) {
  const raw = String(text || '');
  const flat = flatText(raw);
  const amt = flat.match(/Monto:\s*(RD\$|US\$|DOP|USD|EUR)?\s*([\d,]+\.\d{2})/i);
  if (!amt) return [];
  // each field is on its own line, as in the email sent ("Nombre del Beneficiario: X" ⏎ "Concepto: …")
  const field = re => { const m = raw.match(re); return m ? flatText(m[1]) : ''; };
  const beneficiary = field(/Nombre del Beneficiario:[ \t*]*([^\n]+)/i);
  let sender = field(/(?:^|\n)[ \t*]*(?:Nombre del (?:Ordenante|Originante|Originador|Remitente|Emisor)|Ordenante|Remitente|Originador|Enviad[oa] por):[ \t*]*([^\n]+)/i);
  if (!sender) sender = field(/Concepto:[ \t*]*LBTR\s+([^\n]+)/i);
  const fromBank = field(/Banco (?:Ordenante|Origen|Emisor|Remitente):[ \t*]*([^\n]+)/i);
  const ref = (flat.match(/No\.?\s*Referencia:\s*([A-Za-z0-9.\-]{4,40})/i) || [])[1] || '';
  const merchant = sender || ('Transferencia recibida' + (fromBank ? ' desde ' + fromBank : ''));
  return [{ amount: moneyNumber(amt[2]), currency: amt[1] ? moneyCurrency(amt[1]) : 'DOP', merchant: merchant.substring(0, 50),
    own: !!(sender && beneficiary && sameHolder(sender, beneficiary)), ref: ref ? 'IN:' + ref : '', context: 'incoming transfer' }];
}

/* ======================================================================
 * LAFISE online banking "Aviso de transferencia en banco local" — v1.1.63 (real sample; fixture with invented data)
 *   Cuenta de origen Titular: <HOLDER> … Cuenta destino Titular: <NAME> … Concepto: <text> Monto: 12,500.00 DOP
 *   … Resultado Estado: <status> Referencia: <n>
 * A failed one ("Estado: Error") never gets here: isDeclinedTransactionEmail filters it. To your own name, it's Exclude.
 * ====================================================================== */
function extractLAFISELocalBankTransfer(text) {
  const flat = flatText(text);
  if (!/Aviso de transferencia/i.test(flat)) return [];
  const amt = flat.match(/Monto:\s*([\d,]+\.\d{2})\s*(DOP|USD|EUR)\b/i);
  if (!amt) return [];
  const origin = ((flat.match(/Cuenta de origen\s+Titular:\s*(.+?)\s+N[úu]mero de cuenta:/i) || [])[1] || '').trim();
  const dest = ((flat.match(/Cuenta destino\s+Titular:\s*(.+?)\s+N[úu]mero de cuenta:/i) || [])[1] || '').trim();
  const concept = ((flat.match(/Concepto:\s*(.+?)\s+Monto:/i) || [])[1] || '').trim();
  const ref = (flat.match(/Referencia:\s*(\d{4,})/i) || [])[1];
  return [{ amount: parseFloat(amt[1].replace(/,/g, '')), currency: amt[2].toUpperCase(), merchant: (dest || 'Transferencia').substring(0, 50),
    context: concept ? 'Concepto: ' + concept : 'local bank transfer', own: !!(origin && dest && sameHolder(origin, dest)),
    ref: ref ? 'OUT:' + ref : '' }];
}

/** LAFISE transfers sent: the online-banking notice, or the app's "¡Transferencia exitosa!". */
function extractLAFISEAnyTransfer(text) {
  const notice = extractLAFISELocalBankTransfer(text);
  return notice.length ? notice : extractLAFISETransferTransactions(text);
}

/* ======================================================================
 * UBER RIDES — v1.1.67 (real LAFISE alerts and Uber receipts, Sept–Oct 2026; fixtures with invented data)
 * Uber authorizes an estimate on the card when a ride is requested. When the fare changes it authorizes the final
 * amount and the estimate is released; when the final fare is lower it may charge it with no new alert. A request
 * that never becomes a trip leaves an authorization too. The bank emails every authorization as a purchase, so a ride
 * was saved twice (estimate + final), or at the estimate, and the card statement shows only what Uber charged.
 * Uber's trip receipt ("Your <day> <time> trip with Uber") says what each card was charged and when the ride was
 * requested: the ride alerts around it are checked against it (matchRideReceipts) and the sheet is corrected
 * (reconcileRideReceipts, 04_sheetsWriter.gs). Uber Eats is charged once and is not touched.
 * ====================================================================== */
const RIDE_RECEIPT_QUERY = 'from:noreply@uber.com subject:"trip with Uber"';
const RIDE_RECEIPT_SUBJECT = /trip with Uber/i;
const RIDE_MERCHANT_RE = /\bUBER\s*\*?\s*(?:RIDES?|TRIP)\b|\bUBR\s*\*|PENDING\.UBER/i;
const RECEIPT_MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };

/**
 * The receipt's text → { requested: Date (the time printed at its top, the rider's local time), payments: [{ last4,
 * bank, currency, amount }] }, or null when it isn't a trip receipt. Only card payments whose bank (in parentheses,
 * "Mastercard ••••1234 (LAFISE)") is one the tracker reads; Uber Cash, cash and other cards are left out. Works on the
 * text by lines or flattened to one line. Pure — see tests/.
 */
function parseUberTripReceipt(text) {
  const flat = String(text || '').replace(/[\u00a0\u202f]/g, ' ').replace(/\s+/g, ' ');
  const when = /\b([A-Z][a-z]{2})[a-z]*\.? (\d{1,2}), (\d{4}) ,? ?(\d{1,2}):(\d{2}) ?([AP])\.? ?M\b/i.exec(flat);
  if (!when || RECEIPT_MONTHS[when[1].toLowerCase()] === undefined) return null;
  const hour = Number(when[4]) % 12 + (/p/i.test(when[6]) ? 12 : 0);
  const requested = new Date(Number(when[3]), RECEIPT_MONTHS[when[1].toLowerCase()], Number(when[2]), hour, Number(when[5]));
  const payments = [];
  const re = /(?:•|\*){2,} ?(\d{4}) ?\(([^)]+)\) ?(RD\$|DOP|US\$|USD) ?([\d,]+\.\d{2})/gi;
  let m;
  while ((m = re.exec(flat)) !== null) {
    const label = m[2].toUpperCase().replace(/\s+/g, '');
    const bank = BANK_ORDER.find(b => label.indexOf(b) !== -1);
    if (bank) payments.push({ last4: m[1], bank: bank, currency: moneyCurrency(m[3]), amount: moneyNumber(m[4]) });
  }
  return { requested: requested, payments: payments };
}

/**
 * Ride alerts vs. trip receipts. alerts: [{ id, bank, currency, amount, at (ms, the alert email's time), state: ''
 * | 'charged' | 'hold' (settled by an earlier run) }]; receipts: [{ bank, currency, amount (that card's charge),
 * requested, sent (ms) }]. Returns { decisions: { id: { kind: 'charge' } | { kind: 'hold' } | { kind: 'adjust', amount,
 * was } }, unmatched: [ids] } — decisions only for alerts not settled yet. Pure — see tests/.
 *  1. the charge: an alert of the receipt's amount from the request (−2 min) to the receipt (+15 min), the one
 *     closest to the receipt (the final authorization comes with it);
 *  2. no alert of that amount: the ride's own authorization (the alert closest to the request) was charged at the
 *     receipt's amount — corrected;
 *  3. every other alert from 30 min before the request up to the charge: the estimate, or requests that never
 *     became a trip — holds.
 * An alert near no receipt is left as it is (unmatched): a missing receipt never removes a real charge.
 */
const RIDE_WINDOW = { beforeRequest: 2, afterReceipt: 15, attemptsBefore: 30 };   // minutes
function matchRideReceipts(alerts, receipts) {
  const M = 60000, used = {}, decisions = {};
  const seen = {};
  const list = receipts.filter(r => {   // Uber sends some receipts twice
    const k = r.bank + '|' + r.currency + '|' + Math.round(r.amount * 100) + '|' + r.requested;
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  }).sort((a, b) => a.requested - b.requested);
  const fits = (a, r) => a.bank === r.bank && a.currency === r.currency &&
    a.at >= r.requested - RIDE_WINDOW.beforeRequest * M && a.at <= r.sent + RIDE_WINDOW.afterReceipt * M;
  const charge = [];
  list.forEach((r, i) => {
    let best = null;
    alerts.forEach(a => {
      if (used[a.id] || !fits(a, r) || Math.abs(a.amount - r.amount) >= 0.005) return;
      if (!best || Math.abs(a.at - r.sent) < Math.abs(best.at - r.sent)) best = a;
    });
    if (!best) return;
    used[best.id] = true;
    charge[i] = best;
    if (best.state !== 'charged') decisions[best.id] = { kind: 'charge' };
  });
  list.forEach((r, i) => {
    if (charge[i]) return;
    let best = null;
    alerts.forEach(a => {
      if (used[a.id] || a.state || !fits(a, r)) return;
      if (!best || Math.abs(a.at - r.requested) < Math.abs(best.at - r.requested)) best = a;
    });
    if (!best) return;
    used[best.id] = true;
    charge[i] = best;
    decisions[best.id] = { kind: 'adjust', amount: r.amount, was: best.amount };
  });
  list.forEach((r, i) => {
    if (!charge[i]) return;
    alerts.forEach(a => {
      if (used[a.id] || a.bank !== r.bank || a.currency !== r.currency ||
          a.at < r.requested - RIDE_WINDOW.attemptsBefore * M || a.at > charge[i].at) return;
      used[a.id] = true;
      if (a.state !== 'hold') decisions[a.id] = { kind: 'hold' };
    });
  });
  return { decisions: decisions, unmatched: alerts.filter(a => !a.state && !used[a.id]).map(a => a.id) };
}

// ====================================================================================================
// 04_sheetsWriter.gs
// ====================================================================================================

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
  TX_REF: 13,  // v1.1.23: "<BANK>:<bank's own id>" when the email has one (hidden column N)
  AUTO_CATEGORY: 14   // v1.1.39: the category the tracker last set (hidden column O) — see recategorizeAllTransactions
};
const TX_NUM_COLS = TX_COL.AUTO_CATEGORY + 1;
// v1.1.39: stored when the tracker's own category is empty (a transfer with no rule), so "set to nothing" and
// "not recorded yet" (rows saved before v1.1.39) can be told apart
const AUTO_NONE = '(none)';

/**
 * v1.1.4: canonical sheet order, applied by ensureSheetOrder() below.
 * v1.1.5: swapped Raw_BDI before Raw_POPULAR, per request.
 * v1.1.7: added "Categories" at the end — a new managed reference sheet
 * (see buildOrRefreshCategoriesSheet()) documenting the current category
 * list, since a manually-maintained copy of this (e.g. one carried over
 * from the Excel workbook) has no way to stay in sync when the categorizer
 * changes in code.
 */
// v1.1.37: spending, then everything about investments together, then settings
const CANONICAL_SHEET_ORDER = [
  "Dashboard", "Transactions", "Bank Transfers", "Incoming Transfers",
].concat(BANK_ORDER.map(b => 'Raw_' + b)).concat([   // v1.1.60: from the one list of banks
  "Unrecognized",
  "Holdings", "Investment Ledger", "Portfolio History",
  "Custom Rules", "Investment Accounts", "Configuration", "Categories"
]);
// Settings tabs you rarely open; 📊 Tracker › Show / Hide Settings Tabs toggles them. Custom Rules isn't one of them:
// it's edited often, and the summary emails link to it (a hidden sheet can't be opened from a link).
const SETTINGS_TABS = ["Investment Accounts", "Configuration", "Categories"];

/** v1.1.37: hides the settings tabs, or shows them again if any is hidden. Never done automatically. */
function toggleSettingsTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabs = SETTINGS_TABS.map(n => ss.getSheetByName(n)).filter(Boolean);
  if (!tabs.length) return;
  const show = tabs.some(s => s.isSheetHidden());
  tabs.forEach(s => { if (show) s.showSheet(); else s.hideSheet(); });
  safeAlert(show ? "👁 Settings tabs shown again: " + tabs.map(s => s.getName()).join(", ")
    : "🙈 Settings tabs hidden: " + tabs.map(s => s.getName()).join(", ") + ".\nRun the same menu item to show them again.");
}

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
  "Portfolio History": "#0F766E",     // v1.1.32
  "Unrecognized": "#C00000",         // v1.1.35
  "Incoming Transfers": "#38761D"    // v1.1.51
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
  // v1.1.38: moving a tab means activating it, which can show a hidden one — remember which were hidden and hide them
  // again afterwards; and when the order is already right, move nothing
  const wanted = CANONICAL_SHEET_ORDER.filter(name => ss.getSheetByName(name));
  const current = ss.getSheets().map(s => s.getName()).filter(n => wanted.indexOf(n) !== -1);
  if (current.join('|') !== wanted.join('|')) {
    const hidden = ss.getSheets().filter(s => s.isSheetHidden());
    let pos = 1;
    wanted.forEach(name => {
      ss.setActiveSheet(ss.getSheetByName(name));
      ss.moveActiveSheet(pos);
      pos++;
    });
    hidden.forEach(s => { if (!s.isSheetHidden()) s.hideSheet(); });
  }

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (TAB_COLORS[name]) {
      sheet.setTabColor(TAB_COLORS[name]);
    } else if (name.startsWith("Raw_")) {
      sheet.setTabColor(RAW_BANK_TAB_COLOR);
    }
  });

  if (previousActive && !previousActive.isSheetHidden()) ss.setActiveSheet(previousActive);
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
const TRANSFERS_HEADERS = ["Date", "Bank", "Beneficiary / Description", "Category", "Amount", "Currency", "Email Subject", "Id"];   // v1.1.39: Id (hidden)

/**
 * v1.1.2: dedicated sheet for Type = "Transfer" rows.
 * v1.1.4: renamed "Transferencias" → "Bank Transfers", English headers.
 */
/** v1.1.51: money received — same layout as Bank Transfers (amounts shown as received, positive). */
const INCOMING_SHEET = 'Incoming Transfers';

function getOrCreateTransfersSheet(name) {
  name = name || "Bank Transfers";
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
    sheet.appendRow(TRANSFERS_HEADERS);
    sheet.getRange(1, 1, 1, TRANSFERS_HEADERS.length).setFontWeight("bold");
    sheet.hideColumns(TRANSFERS_HEADERS.length);   // v1.1.39: Id — internal, matches your edits back to Transactions
    ensureRowCapacity(sheet, 2000); // v1.1.19: formats below cover 2,000 rows
    sheet.getRange(2, 1, 1999).setNumberFormat('yyyy-MM-dd');
    sheet.getRange(2, 5, 1999).setNumberFormat('#,##0.00');
    ensureAutoFilter(sheet, TRANSFERS_HEADERS.length, 2000);
    sheet.setTabColor(TAB_COLORS[name] || TAB_COLORS["Bank Transfers"]);
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
  headerBg: '#1F3864', headerFg: '#FFFFFF', stripe: '#F7F9FC', attention: '#FDBA74', problem: '#FCA5A5',   // v1.1.37: were #FFF4D6 and #FDECEC, ΔE 7 and 6 from chips (tested)
  refund: '#2E7D32', keyBg: '#F3F6FB', keyFg: '#374151', example: '#9CA3AF'
};
const TYPE_COLORS = { 'Transfer': ['#E0F2F1', '#00695C'], 'Card Payment': ['#F3F4F6', '#4B5563'] };   // WCAG AA (was 4.39:1)

/** Which column holds what (1-based) in each tracker data sheet; null for other sheets. */
function dataSheetLayout(name) {
  if (name === TRANSACTIONS_SHEET) {
    return { cols: TX_NUM_COLS, merchant: TX_COL.MERCHANT + 1, category: TX_COL.CATEGORY + 1, amount: TX_COL.AMOUNT + 1,
      currency: TX_COL.CURRENCY + 1, type: TX_COL.TYPE + 1, transfersOnlyWhenType: true };
  }
  if (name === 'Bank Transfers' || name === INCOMING_SHEET) return { cols: TRANSFERS_HEADERS.length, merchant: 3, category: 4, amount: 5, currency: 6, allTransfers: true };
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
      else if (name === INVESTMENT_ACCOUNTS_SHEET) styleAccountsSheet(sheet);   // v1.1.34
      else if (name === HISTORY_SHEET) styleHistorySheet(sheet);
      else if (name === UNRECOGNIZED_SHEET) styleUnrecognizedSheet(sheet);           // v1.1.35
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
    const isData = name === TRANSACTIONS_SHEET || name === 'Bank Transfers' || name === INCOMING_SHEET || name.indexOf('Raw_') === 0 ||   // v1.1.53: Incoming
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
  const inc = ss.getSheetByName(INCOMING_SHEET);   // v1.1.53
  if (inc) sortSheetByDateDesc(inc, 1);
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
  // v1.1.58: a yyyy-mm-dd text IS the day. new Date('2026-08-28') is midnight UTC — the 27th in Santo Domingo (UTC−4) —
  // so a date typed as text came out a day early
  const iso = str.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
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

/**
 * v1.1.66: a transfer received can arrive twice — its email (BANESCO since v1.1.62) and the monthly statement, which
 * lists every credit (v1.1.55). Their references differ, so both were saved and the money counted twice. A statement
 * credit and a notified one (an email, or a row typed by hand) of the same bank, currency and amount, at most
 * STATEMENT_MATCH_DAYS apart, are the same money: the one saved first stays. The days differ: the email is dated when
 * it arrives, the statement on the day the bank posts it (a Friday-evening transfer posts on Monday). Each credit
 * pairs once, so two equal transfers with only one notified still save the other.
 */
const STATEMENT_MATCH_DAYS = 5;
function isStatementRef(ref) { return /:STMT:/.test(String(ref || '')); }
function incomingKey(bank, currency, amount) {
  return String(bank) + '|' + (currency || 'DOP') + '|' + Math.round(Math.abs(Number(amount)) * 100);
}
function dayNumber(date) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(normalizeDateForCompare(date));
  return m ? Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) / 86400000 : NaN;
}
/** The closest unpaired credit of the other kind (statement ↔ notified) for this one, or null. */
function findIncomingPair(index, key, day, statement) {
  let best = null;
  (index.incoming[key] || []).forEach(e => {
    if (e.paired || e.statement === statement || !(Math.abs(e.day - day) <= STATEMENT_MATCH_DAYS)) return;
    if (!best || Math.abs(e.day - day) < Math.abs(best.day - day)) best = e;
  });
  return best;
}

function buildExistingIndex(values) {
  const messageIds = new Set();
  const legacyKeys = new Set();
  const refs = new Set();
  const incoming = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const id = row[TX_COL.MESSAGE_ID];
    if (id) messageIds.add(String(id));
    else legacyKeys.add(legacyKey(row[TX_COL.DATE], row[TX_COL.BANK], row[TX_COL.AMOUNT]));
    if (row[TX_COL.TX_REF]) refs.add(String(row[TX_COL.TX_REF]));
    if (row[TX_COL.TYPE] === 'Incoming' && row[TX_COL.DATE]) {
      const key = incomingKey(row[TX_COL.BANK], row[TX_COL.CURRENCY], row[TX_COL.AMOUNT]);
      (incoming[key] = incoming[key] || []).push({ day: dayNumber(row[TX_COL.DATE]), statement: isStatementRef(row[TX_COL.TX_REF]), paired: false });
    }
  }
  const index = { messageIds: messageIds, legacyKeys: legacyKeys, refs: refs, incoming: incoming };
  // both copies saved before v1.1.66: paired now, so neither is taken as the copy of a later credit
  Object.keys(incoming).forEach(key => incoming[key].forEach(e => {
    if (!e.statement || e.paired) return;
    const other = findIncomingPair(index, key, e.day, true);
    if (other) { e.paired = true; other.paired = true; }
  }));
  return index;
}

function selectNewTransactions(transactions, index) {
  const fresh = [];
  let duplicates = 0;
  if (!index.incoming) index.incoming = {};
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
    if (t.type === 'Incoming') {   // v1.1.66: the same transfer received, from its email and from the statement
      const inKey = incomingKey(t.bank, t.currency, t.amount), day = dayNumber(t.date), statement = isStatementRef(ref);
      const pair = findIncomingPair(index, inKey, day, statement);
      if (pair) {
        pair.paired = true;
        duplicates++;
        Logger.log('Already saved from ' + (statement ? 'its notice' : 'the statement') + ': ' + t.bank + ' ' + (t.currency || 'DOP') + ' ' +
          Math.abs(Number(t.amount)).toFixed(2) + ' received ' + normalizeDateForCompare(t.date) + ' — not saved twice');
        continue;
      }
      (index.incoming[inKey] = index.incoming[inKey] || []).push({ day: day, statement: statement, paired: false });
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
    t.isCashback ? 'YES' : 'NO', t.type || 'Transaction', t.messageId || '', t.txRef || '',
    t.autoCategory !== undefined ? t.autoCategory : (t.category || AUTO_NONE)   // v1.1.54: rows typed by hand: none
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

/** v1.1.23: adds the hidden TxRef header (column N) to a sheet created before it existed; v1.1.39: Auto Category (O). */
function ensureTransactionsSchema(sheet) {
  [[TX_COL.TX_REF, 'TxRef'], [TX_COL.AUTO_CATEGORY, 'Auto Category']].forEach(pair => {
    if (sheet.getRange(1, pair[0] + 1).getValue() !== pair[1]) {
      sheet.getRange(1, pair[0] + 1).setValue(pair[1]).setFontWeight('bold');
      sheet.hideColumns(pair[0] + 1);
    }
  });
}

/**
 * v1.1.39: a category you set yourself is kept. The tracker stores the category IT set (Auto Category); a category
 * that differs from it was changed by hand. Rows saved before v1.1.39 have no Auto Category: there, a transfer with a
 * category no rule gives is taken as yours. An empty category is never "yours" — clear a cell to get the automatic one
 * back. Pure — see tests/.
 */
function isManualCategory(row, computedCategory) {
  const current = String(row[TX_COL.CATEGORY] || '').trim();
  if (!current) return false;
  const stored = row[TX_COL.AUTO_CATEGORY];
  if (stored !== undefined && stored !== null && String(stored) !== '') {
    return current !== (String(stored) === AUTO_NONE ? '' : String(stored));
  }
  return (row[TX_COL.TYPE] || '') === 'Transfer' && !computedCategory;   // before v1.1.39
}

/**
 * v1.1.39: Bank Transfers is rebuilt from Transactions on every run, so a category typed THERE was lost. Before
 * recategorizing, a Bank Transfers category that differs from its Transactions row is copied back as your own.
 * Rows are matched by the hidden Id (Gmail message id), or — on a sheet from before it — by date, beneficiary and amount.
 */
function syncTransferCategoryEdits(txValues) {
  // v1.1.51: Bank Transfers and Incoming Transfers alike
  return syncCategoryEditsFrom('Bank Transfers', 'Transfer', txValues).concat(syncCategoryEditsFrom(INCOMING_SHEET, 'Incoming', txValues));
}

function syncCategoryEditsFrom(sheetName, rowType, txValues) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const width = Math.max(sheet.getLastColumn(), TRANSFERS_HEADERS.length);
  const header = sheet.getRange(1, 1, 1, width).getValues()[0];
  if (header[3] !== 'Category') return [];
  const idCol = header.indexOf('Id');
  const keyOf = (date, merchant, amount) => normalizeDateForCompare(date) + '|' + String(merchant).trim() + '|' + Number(amount);
  const byId = {}, byKey = {};
  txValues.forEach((r, i) => {
    if (r[TX_COL.TYPE] !== rowType) return;
    if (r[TX_COL.MESSAGE_ID]) byId[String(r[TX_COL.MESSAGE_ID])] = i;
    byKey[keyOf(r[TX_COL.DATE], r[TX_COL.MERCHANT], Math.abs(Number(r[TX_COL.AMOUNT]) || 0))] = i;
  });
  const changed = [];
  sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues().forEach(b => {
    if (!b[0]) return;
    const i = idCol !== -1 && b[idCol] && byId[String(b[idCol])] !== undefined ? byId[String(b[idCol])] : byKey[keyOf(b[0], b[2], Math.abs(Number(b[4]) || 0))];
    if (i === undefined) return;
    const edited = String(b[3] || '').trim();
    if (edited && edited !== String(txValues[i][TX_COL.CATEGORY] || '').trim()) {
      txValues[i][TX_COL.CATEGORY] = edited;
      changed.push(i);
    }
  });
  return changed;
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
  if (oldType === 'Incoming') type = 'Incoming';   // v1.1.51: set when read (a statement's subject says nothing)

  let category = findCustomRuleOverride(description, rawCustomRules);
  if (!category && type === 'Transaction') category = categorizeTransaction(description);
  // v1.1.23: paying the card is never spending — explicit "Exclude"
  if (type === 'Card Payment') category = EXCLUDE_CATEGORY;
  // v1.1.51: your own money. v1.1.66: sent, too — a transfer to your own account was Exclude when read (v1.1.63) and
  // lost it here at the end of the same run
  if ((type === 'Incoming' || type === 'Transfer') && description.endsWith(OWN_ACCOUNT_SUFFIX)) category = EXCLUDE_CATEGORY;
  if (type === 'Transaction' && description.endsWith(RIDE_HOLD_SUFFIX)) category = EXCLUDE_CATEGORY;   // v1.1.67: never charged
  // v1.1.23: an unmatched reversal keeps whatever category it has (normally
  // blank) instead of being guessed from the placeholder text
  if (merchant === REVERSAL_UNMATCHED) category = row[TX_COL.CATEGORY] || '';

  const oldCurrency = row[TX_COL.CURRENCY];
  const currency = (oldCurrency && !VALID_CURRENCIES.has(oldCurrency)) ? 'DOP' : oldCurrency;
  const isCredit = (type === 'Card Payment' || type === 'Cashback' || type === 'Incoming') ? 'YES' : row[TX_COL.IS_CREDIT];
  const isCashback = type === 'Cashback' ? 'YES' : 'NO';
  // v1.1.62: money received is saved negative. A row that becomes Incoming here (BANESCO's "Transferencia Recibida",
  // saved as a positive transfer before this version) turns negative, so it stops counting as spending.
  const oldAmount = row[TX_COL.AMOUNT];
  const amount = (type === 'Incoming' && typeof oldAmount === 'number' && oldAmount > 0) ? -oldAmount : oldAmount;
  return { type: type, category: category || '', currency: currency, isCredit: isCredit, isCashback: isCashback, amount: amount };
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
  // v1.1.54: rows typed by hand in Incoming Transfers are saved before that sheet is rebuilt below
  try { importTypedIncomingRows(); } catch (error) { Logger.log('Could not save the rows typed in Incoming Transfers: ' + error); }
  try { importTypedTransferRows('Bank Transfers', 'Transfer'); } catch (error) { Logger.log('Could not save the rows typed in Bank Transfers: ' + error); }   // v1.1.57

  const txSheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  if (txSheet && txSheet.getLastRow() > 1) {
    const numRows = txSheet.getLastRow() - 1;
    const numCols = Math.max(txSheet.getLastColumn(), TX_NUM_COLS);
    const values = txSheet.getRange(2, 1, numRows, numCols).getValues();
    const asText = v => (v === null || v === undefined) ? '' : String(v);
    const changedCols = new Set();
    ensureTransactionsSchema(txSheet);
    const fromTransfers = syncTransferCategoryEdits(values);   // v1.1.39: categories typed in Bank Transfers
    if (fromTransfers.length) {
      changedCols.add(TX_COL.CATEGORY);
      totalChanged += fromTransfers.length;
      Logger.log('✍️ Categories set in Bank Transfers kept: ' + fromTransfers.length);
    }
    let keptManual = 0;

    values.forEach(row => {
      // v1.1.34: the ONE merchant repair recategorize makes, kept outside computeRecategorization (which must never
      // rewrite merchants — C5): a merchant that is exactly a status word ("Aprobada") was never a merchant.
      const fixed = fixStatusAsMerchant(row[TX_COL.MERCHANT], row[TX_COL.SUBJECT], '');
      if (fixed !== row[TX_COL.MERCHANT]) {
        if (!row[TX_COL.DESCRIPTION] || row[TX_COL.DESCRIPTION] === row[TX_COL.MERCHANT]) {
          row[TX_COL.DESCRIPTION] = fixed;
          changedCols.add(TX_COL.DESCRIPTION);
        }
        row[TX_COL.MERCHANT] = fixed;
        changedCols.add(TX_COL.MERCHANT);
        totalChanged++;
      }
      const r = computeRecategorization(row, rawCustomRules);
      // v1.1.39: a category you set by hand is kept; the tracker's own category is recorded either way
      const manual = isManualCategory(row, r.category);
      if (manual) keptManual++;
      const auto = r.category || AUTO_NONE;
      if (asText(row[TX_COL.AUTO_CATEGORY]) !== auto) { row[TX_COL.AUTO_CATEGORY] = auto; changedCols.add(TX_COL.AUTO_CATEGORY); }
      [[TX_COL.TYPE, r.type], [TX_COL.CATEGORY, manual ? row[TX_COL.CATEGORY] : r.category], [TX_COL.CURRENCY, r.currency],
       [TX_COL.IS_CREDIT, r.isCredit], [TX_COL.IS_CASHBACK, r.isCashback], [TX_COL.AMOUNT, r.amount]].forEach(pair => {
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
    if (keptManual) Logger.log('✍️ Categories you set by hand, kept: ' + keptManual);
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
  if (header[3] === TRANSFERS_HEADERS[3]) {
    if (header[7] !== 'Id') {   // v1.1.39: the hidden Id column
      sheet.getRange(1, 8).setValue('Id');
      sheet.hideColumns(8);
      ensureAutoFilter(sheet, TRANSFERS_HEADERS.length, 2000);
    }
    return false;
  }
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
  const incoming = [];  // v1.1.51: Incoming Transfers rows, same layout

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
        row[TX_COL.CURRENCY], row[TX_COL.SUBJECT], row[TX_COL.MESSAGE_ID] || ''
      ]);
    } else if (type === 'Incoming') {
      incoming.push([
        row[TX_COL.DATE], bank, row[TX_COL.MERCHANT], row[TX_COL.CATEGORY], Math.abs(Number(row[TX_COL.AMOUNT]) || 0),
        row[TX_COL.CURRENCY], row[TX_COL.SUBJECT], row[TX_COL.MESSAGE_ID] || ''
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

  // v1.1.51: Incoming Transfers — created the first time there's money received
  let inSheet = ss.getSheetByName(INCOMING_SHEET);
  if (incoming.length && !inSheet) inSheet = getOrCreateTransfersSheet(INCOMING_SHEET);
  if (inSheet) {
    const inLast = inSheet.getLastRow();
    if (inLast > 1) inSheet.getRange(2, 1, inLast - 1, TRANSFERS_HEADERS.length).clearContent();
    if (incoming.length) {
      ensureRowCapacity(inSheet, incoming.length + 1);
      inSheet.getRange(2, 1, incoming.length, TRANSFERS_HEADERS.length).setValues(incoming);
      inSheet.getRange(2, 1, incoming.length, 1).setNumberFormat('yyyy-MM-dd');
      inSheet.getRange(2, 5, incoming.length, 1).setNumberFormat('#,##0.00');
    }
    inSheet.getRange(1, 3).setValue('From');
  }

  return Object.keys(byBank).map(b => b + "=" + byBank[b].length).join(", ") + " | Transfers=" + transfers.length + " | Incoming=" + incoming.length;
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
  // v1.1.59: other incomes and deductions — totals per currency; a sheet from before gets them from its setup
  otherIncomeDOP:    { name: 'CFG_OTHER_INCOME_DOP', fallback: 0, derive: c => moneyLineTotals(c.otherIncomes).DOP },
  otherIncomeUSD:    { name: 'CFG_OTHER_INCOME_USD', fallback: 0, derive: c => moneyLineTotals(c.otherIncomes).USD },
  otherIncomeEUR:    { name: 'CFG_OTHER_INCOME_EUR', fallback: 0, derive: c => moneyLineTotals(c.otherIncomes).EUR },
  otherDeductionDOP: { name: 'CFG_OTHER_DED_DOP',    fallback: 0, derive: c => moneyLineTotals(c.otherDeductions).DOP },
  otherDeductionUSD: { name: 'CFG_OTHER_DED_USD',    fallback: 0, derive: c => moneyLineTotals(c.otherDeductions).USD },
  otherDeductionEUR: { name: 'CFG_OTHER_DED_EUR',    fallback: 0, derive: c => moneyLineTotals(c.otherDeductions).EUR }
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
  let config = null;
  Object.keys(CONFIG_NAMED_RANGES).forEach(key => {
    let row = keys.indexOf(key) + 1;
    if (row === 0) {
      const def = CONFIG_NAMED_RANGES[key];
      let value = def.fallback;
      if (def.derive) {   // v1.1.59: e.g. an older setup's single "other income" keeps showing on the Dashboard
        try { config = config || getConfig() || {}; value = def.derive(config); } catch (error) { value = def.fallback; }
      }
      sheet.appendRow([key, value]);
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
  BHD_MIPAIS:            { bank: 'BHD',     name: 'Mi País',            cashback: 0.05, note: CARD_NOTES.BHD },
  // v1.1.60: public product facts (each bank's site, Sept 2026). The rate is the BASE cashback every purchase earns —
  // what the Dashboard uses; points, miles and day- or merchant-specific promotions are in the note, not in the rate.
  // Nothing about a particular holder (credit limit, statement and payment days) belongs here.
  BHD_LIFEMILES:         { bank: 'BHD',     name: 'Visa LifeMiles',     cashback: 0,
    note: 'Miles, not cashback: 2 LifeMiles per US$1 at Avianca and Star Alliance airlines, 1 per US$1 elsewhere. ' +
      'Installment purchases and cash advances earn none. 4 LoungeKey visits a year.' },
  BDI_VISA_CLASICA:      { bank: 'BDI',     name: 'Visa Clásica',       cashback: 0,
    note: 'No base cashback: returns of 15-20% on set days at specific merchants (a yearly program, with minimums ' +
      'per purchase and caps). Check the current program on the bank\'s site.' },
  SCOTIABANK_AMEX_GOLD:  { bank: 'SCOTIABANK', name: 'American Express Gold', cashback: 0,
    note: 'Membership Rewards points, not cashback: 3x in restaurants, bars and entertainment; 2x in clothing, shoes, ' +
      'gyms, spas and salons; 1x (1 point per US$1) elsewhere. Points don\'t expire. Cashback promotions vary by year.' },
  QIK_MASTERCARD:        { bank: 'QIK',     name: 'Mastercard',         cashback: 0.01,
    note: '1% cashback on every purchase, no cap, credited weekly. No issuance, renewal or handling fees.' }
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

/**
 * v1.1.62: categories used in Transactions that are neither a default one nor a Custom Rule's (typed by hand in
 * Transactions, Bank Transfers or Incoming Transfers). The Dashboard only had rows for the named ones, so amounts in a
 * category typed by hand were in no row and no total. Exclude is never a row. Case and spaces are ignored, as SUMIFS
 * ignores case.
 */
function categoriesOnlyInTransactions(named) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSACTIONS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const known = new Set(named.map(c => String(c).trim().toUpperCase()).concat([EXCLUDE_CATEGORY.toUpperCase()]));
  const out = [];
  sheet.getDataRange().getValues().slice(1).forEach(r => {
    const c = String(r[TX_COL.CATEGORY] === null || r[TX_COL.CATEGORY] === undefined ? '' : r[TX_COL.CATEGORY]).trim();
    if (c && !known.has(c.toUpperCase())) { known.add(c.toUpperCase()); out.push(c); }
  });
  return out.sort();
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
  const named = getCategories().concat(userEmail ? getCustomCategoryNames(userEmail) : []);
  const categories = named.concat(categoriesOnlyInTransactions(named));   // v1.1.62
  const banks = Object.keys(BANK_PATTERNS).filter(b => !(config && config.banksToTrack) || config.banksToTrack[b]);
  const n = categories.length;
  const put = (range, v) => (typeof v === 'string' && v.charAt(0) === '=') ? range.setFormula(v) : range.setValue(v);
  const border = SpreadsheetApp.BorderStyle;

  // ---- row map
  const R = { kpiLabel: 7, kpiValue: 8, kpiNote: 9, sec1: 11, head1: 12, catFirst: 13 };
  R.catLast = R.catFirst + n - 1;
  R.transfers = R.catLast + 1;
  R.total = R.transfers + 1;
  R.income = 12; R.netDop = 19;                                   // right block: 12..19 (v1.1.59: + other deductions)
  R.fvHead = 21; R.fvFirst = 22;                                  // 22..24
  R.bankHead = 26; R.bankFirst = 27; R.bankLast = R.bankFirst + banks.length - 1;
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
      note: '="Gross " & TEXT(CFG_MONTHLY_INCOME,"#,##0") & " " & CFG_INCOME_CURRENCY & IF(N' + (R.income + 5) + '>0," + other income","") & ' +
        'IF(N' + (R.income + 6) + '>0," − other deductions","") & " / month"' },
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
    // v1.1.59: other incomes (in full, no deductions) and other deductions — every line converted at the live rates
    ['Other income (DOP-equivalent, no deductions)', `=CFG_OTHER_INCOME_DOP+CFG_OTHER_INCOME_USD*${USD}+CFG_OTHER_INCOME_EUR*${EUR}`, moneyFmt],
    ['Other deductions (DOP-equivalent)', `=CFG_OTHER_DED_DOP+CFG_OTHER_DED_USD*${USD}+CFG_OTHER_DED_EUR*${EUR}`, moneyFmt],
    ['Net income (DOP-equivalent)', `=IF(CFG_INCOME_CURRENCY="DOP",N${R.income + 4},N${R.income + 4}*${USD})+N${R.income + 5}-N${R.income + 6}`, moneyFmt]
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

/* ======================================================================
 * UNRECOGNIZED — v1.1.35
 * Every bank or broker email the tracker couldn't read (or saved with an unreadable merchant), so failures show up by
 * themselves instead of being found by spotting odd rows. One row per email (a later run updates it, keeping your
 * Status); a row disappears once a later version reads that email cleanly.
 * ====================================================================== */
const UNRECOGNIZED_SHEET = 'Unrecognized';
const UNRECOGNIZED_HEADERS = ['Date', 'Bank', 'Subject', 'Reason', 'What the email says', 'Gmail', 'Status', 'First seen', 'Last seen', 'Id'];
const UNRECOGNIZED_STATUSES = ['New', 'Ignore'];

/** Merges this run's entries into the sheet; removes emails now read cleanly. Returns how many are still New. */
function recordUnrecognized(entries, readIds, now) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(UNRECOGNIZED_SHEET);
  if (!sheet && !(entries || []).length) return 0;
  if (!sheet) {
    sheet = ss.insertSheet(UNRECOGNIZED_SHEET);
    sheet.getRange(1, 1, 1, UNRECOGNIZED_HEADERS.length).setValues([UNRECOGNIZED_HEADERS]);
    sheet.hideColumns(UNRECOGNIZED_HEADERS.length);
  }
  const width = UNRECOGNIZED_HEADERS.length, idCol = width - 1;
  const last = sheet.getLastRow();
  let rows = last > 1 ? sheet.getRange(2, 1, last - 1, width).getValues().filter(r => r[idCol]) : [];
  const read = new Set(readIds || []);
  rows = rows.filter(r => !read.has(String(r[idCol])));                       // read cleanly now → resolved
  const byId = {};
  rows.forEach(r => { byId[String(r[idCol])] = r; });
  (entries || []).forEach(e => {
    const link = !/^(typed|notice):/.test(String(e.id)) ? '=HYPERLINK("https://mail.google.com/mail/u/0/#all/' + e.id + '","Open")' : '';   // v1.1.54: not an email → no link
    const snippet = String(e.snippet || '').replace(/<https?:[^>]*>/g, ' ').replace(/https?:\/\/\S+/g, ' ')   // v1.1.37: no link noise
      .replace(/\s+/g, ' ').trim().substring(0, 300);
    const known = byId[e.id];
    if (known) {                                                               // seen before: refresh, keep the Status
      known[3] = e.reason; known[4] = snippet; known[5] = link; known[8] = now;
    } else {
      const row = [e.date, e.bank, e.subject, e.reason, snippet, link, 'New', now, now, e.id];
      rows.push(row); byId[e.id] = row;
    }
  });
  rows.sort((a, b) => (b[0] instanceof Date ? b[0].getTime() : 0) - (a[0] instanceof Date ? a[0].getTime() : 0));
  if (last > 1) sheet.getRange(2, 1, last - 1, width).clearContent();
  if (rows.length) {
    ensureRowCapacity(sheet, rows.length + 1);
    sheet.getRange(2, 1, rows.length, width).setValues(rows);
  }
  return rows.filter(r => r[6] !== 'Ignore').length;
}

function styleUnrecognizedSheet(sheet) {
  styleHeader(sheet, UNRECOGNIZED_HEADERS.length);
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 1, rows, 1).setNumberFormat('yyyy-MM-dd HH:mm').setHorizontalAlignment('center');
  sheet.getRange(2, 2, rows, 1).setFontWeight('bold').setFontColor(DASH_THEME.navy);
  sheet.getRange(2, 4, rows, 1).setWrap(true);   // long reasons wrap instead of being cut
  sheet.getRange(2, 5, rows, 1).setWrap(true).setFontSize(9).setFontColor('#4B5563');
  sheet.getRange(2, 6, rows, 2).setHorizontalAlignment('center');
  sheet.getRange(2, 8, rows, 2).setNumberFormat('yyyy-MM-dd').setFontSize(9).setFontColor('#6B7280').setHorizontalAlignment('center');
  sheet.getRange(2, 7, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(UNRECOGNIZED_STATUSES, true).setAllowInvalid(false)
    .setHelpText('"Ignore" keeps it here but out of the counts. Rows disappear when a later version reads the email.').build());
  const all = sheet.getRange(2, 1, rows, UNRECOGNIZED_HEADERS.length - 1), reason = sheet.getRange(2, 4, rows, 1), status = sheet.getRange(2, 7, rows, 1);
  const rule = () => SpreadsheetApp.newConditionalFormatRule();
  sheet.setConditionalFormatRules([
    rule().whenFormulaSatisfied('=$G2="Ignore"').setFontColor('#9CA3AF').setRanges([all]).build(),       // ignored: greyed out
    rule().whenTextEqualTo('New').setBackground('#FDECEC').setFontColor('#B91C1C').setRanges([status]).build(),
    rule().whenTextEqualTo('Ignore').setBackground('#F3F4F6').setFontColor('#4B5563').setRanges([status]).build(),
    rule().whenTextStartsWith('Amount not found').setBackground('#FFF4D6').setFontColor('#92400E').setRanges([reason]).build(),
    rule().whenTextStartsWith('Could not read').setBackground('#FDECEC').setFontColor('#B91C1C').setRanges([reason]).build(),
    rule().whenTextStartsWith('Broker email').setBackground('#E0F2F1').setFontColor('#00695C').setRanges([reason]).build(),
    rule().whenTextStartsWith('Saved').setBackground('#F3E8FF').setFontColor('#7E22CE').setRanges([reason]).build(),
    rule().whenTextStartsWith('Deposit without amount').setBackground('#EAF1FE').setFontColor('#1D4ED8').setRanges([reason]).build(),   // v1.1.42
    rule().whenFormulaSatisfied('=AND($A2<>"",ISEVEN(ROW()))').setBackground(SHEET_THEME.stripe).setRanges([all]).build()
  ]);
  [125, 85, 260, 250, 420, 60, 70, 85, 85].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

/** v1.1.36: Gmail ids of every email already in Transactions (MessageId is "<gmail id>_<n>"), to skip re-reading them. */
function savedMessageIds() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSACTIONS_SHEET);
  const ids = new Set();
  if (!sheet || sheet.getLastRow() < 2) return ids;
  sheet.getRange(2, TX_COL.MESSAGE_ID + 1, sheet.getLastRow() - 1, 1).getValues().forEach(r => {
    const v = String(r[0] || '');
    if (v) ids.add(v.replace(/_\d+$/, ''));
  });
  return ids;
}

/* ======================================================================
 * READ LOG — v1.1.43
 * Emails read without saving anything (promotions, statements, notices, declined…) used to be read again on every
 * run — only saved transactions were skipped — so a year-long date range could re-read the same emails and never
 * finish. Their ids go to a hidden sheet, tagged with the tracker version: after an update they are read once more
 * (the new version may read them differently), then skipped again. Emails that failed aren't logged: they're retried.
 * ====================================================================== */
const READ_LOG_SHEET = 'Read Emails';

function readLogSheet(create) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(READ_LOG_SHEET);
  if (!sheet && create) {
    sheet = ss.insertSheet(READ_LOG_SHEET);
    sheet.getRange(1, 1, 1, 2).setValues([['Id', 'Version']]);
    sheet.hideSheet();
  }
  return sheet;
}

/** Gmail ids not to read again: saved transactions, and emails read by this version without saving anything. */
function alreadyReadIds() {
  const ids = savedMessageIds();
  const sheet = readLogSheet(false);
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(r => {
      if (r[0] && String(r[1]) === SCRIPT_VERSION) ids.add(String(r[0]));
    });
  }
  return ids;
}

function logReadEmails(ids) {
  if (!ids || !ids.length) return;
  const sheet = readLogSheet(true);
  const byId = {};
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) byId[String(r[0])] = r[1]; });
  ids.forEach(id => { byId[String(id)] = SCRIPT_VERSION; });
  const rows = Object.keys(byId).map(id => [id, byId[id]]);
  ensureRowCapacity(sheet, rows.length + 1);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

/* ======================================================================
 * TYPED BY HAND IN THE TRANSFER SHEETS — v1.1.54
 * Bank Transfers and Incoming Transfers are rebuilt from Transactions on every update, so:
 *  - a category typed in either is copied to its Transactions row the moment it's typed (onEdit) — the Dashboard reads
 *    Transactions, and used to show the change only after the next update (reported: "incoming transfers don't change
 *    the Dashboard");
 *  - a row typed by hand in Incoming Transfers (no Id) is saved into Transactions before the rebuild, as money received
 *    — or, missing a date or an amount, listed in Unrecognized with what was typed, never silently dropped.
 * ====================================================================== */
function onEditTransferCategory(e) {
  const range = e && e.range;
  if (!range) return;
  const sheet = range.getSheet(), name = sheet.getName();
  if (name !== 'Bank Transfers' && name !== INCOMING_SHEET) return;
  const CAT = 4, ID = TRANSFERS_HEADERS.length;
  if (range.getColumn() > CAT || range.getLastColumn() < CAT || range.getLastRow() < 2) return;
  const first = Math.max(2, range.getRow()), n = range.getLastRow() - first + 1;
  const rows = sheet.getRange(first, 1, n, ID).getValues();
  const tx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSACTIONS_SHEET);
  if (!tx || tx.getLastRow() < 2) return;
  const ids = tx.getRange(2, TX_COL.MESSAGE_ID + 1, tx.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
  rows.forEach(r => {
    const id = String(r[ID - 1] || '').trim();
    if (!id) return;                                   // typed by hand: saved at the next update
    const at = ids.indexOf(id);
    if (at !== -1) tx.getRange(at + 2, TX_COL.CATEGORY + 1).setValue(String(r[CAT - 1] || '').trim());
  });
}

/** Rows typed by hand in Incoming Transfers → Transactions (Type Incoming, negative, your category kept). */
function importTypedIncomingRows(now) {
  return importTypedTransferRows(INCOMING_SHEET, 'Incoming', now);
}

/**
 * v1.1.66: an amount typed by hand → a positive number, or NaN when it can't be read safely. Numbers as they are; text
 * like "1,234.56", "4800", "RD$ 1,500" or "-300" (commas only between thousands, at most two decimals). "1.234,56" used
 * to become 1.23456 (every comma dropped) and was saved as RD$1.23 without a word; it, and "1.500", now go to
 * Unrecognized with what was typed. Pure — see tests/.
 */
function typedAmount(value) {
  if (typeof value === 'number') return Math.abs(value);
  const s = String(value === null || value === undefined ? '' : value).replace(/RD\$|US\$|DOP|USD|EUR|COP|\$|\s/gi, '');
  if (!/^-?(\d{1,3}(,\d{3})+|\d+)(\.\d{1,2})?$/.test(s)) return NaN;
  return Math.abs(Number(s.replace(/,/g, '')));
}

/**
 * v1.1.57: rows typed by hand in a transfer sheet → Transactions. Incoming Transfers: money received (Type Incoming,
 * saved negative). Bank Transfers: money sent (Type Transfer, positive). Your category is kept as yours; the row gets its
 * Id at once, so it's never imported twice; a row without a date or an amount goes to Unrecognized with what was typed.
 */
function importTypedTransferRows(sheetName, type, now) {
  now = now || new Date();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return { added: 0, rejected: 0 };
  const W = TRANSFERS_HEADERS.length, incoming = type === 'Incoming';
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, W).getValues();
  // a row with no Id that matches a saved row (date, name, amount — the key the category sync uses for sheets from
  // before v1.1.39, which had no Ids) is one of those, not typed by hand: importing it would duplicate every transfer
  const keyOf = (date, merchant, amount) => normalizeDateForCompare(date) + '|' + String(merchant).trim() + '|' + Number(amount);
  const saved = new Set();
  // v1.1.58: rows typed by hand that were saved WITHOUT a date (reported, v1.1.57) — typed again, they get their date
  // back instead of being saved a second time
  const undated = {};
  const subject = 'Added by hand in ' + sheetName;
  const tx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSACTIONS_SHEET);
  if (tx && tx.getLastRow() > 1) {
    tx.getRange(2, 1, tx.getLastRow() - 1, TX_NUM_COLS).getValues().forEach((r, i) => {
      if (r[TX_COL.TYPE] !== type) return;
      if (r[TX_COL.DATE]) saved.add(keyOf(r[TX_COL.DATE], r[TX_COL.MERCHANT], Math.abs(Number(r[TX_COL.AMOUNT]) || 0)));
      else if (r[TX_COL.SUBJECT] === subject) {
        const k = String(r[TX_COL.MERCHANT]).trim() + '|' + Math.abs(Number(r[TX_COL.AMOUNT]) || 0).toFixed(2);
        (undated[k] = undated[k] || []).push({ row: i + 2, id: String(r[TX_COL.MESSAGE_ID] || '') });
      }
    });
  }
  const add = [], bad = [];
  let repaired = 0;
  values.forEach((r, i) => {
    if (String(r[W - 1] || '').trim()) return;                                    // has an Id: it comes from Transactions
    if (r.slice(0, W - 1).every(x => x === '' || x === null)) return;             // empty row
    if (r[0] && saved.has(keyOf(r[0], r[2], typedAmount(r[4]) || 0))) return;   // an older sheet's row
    // v1.1.58: the day as yyyy-mm-dd text first, then a NEW date at noon. v1.1.57 passed the cell's own Date on, and
    // rows pasted from Excel (real dates) were saved with none, while rows pasted as text kept theirs (reported)
    const dayKey = r[0] === '' || r[0] === null ? '' : String(normalizeDateForCompare(r[0]) || '').trim();
    const dm = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dayKey);
    const d = dm ? new Date(Number(dm[1]), Number(dm[2]) - 1, Number(dm[3]), 12, 0, 0) : null;
    const amount = typedAmount(r[4]);   // v1.1.66: NaN for a format it can't read safely
    if (!d || isNaN(d.getTime()) || !(amount > 0)) { bad.push(r); return; }
    const uk = (String(r[2] || '').trim() || 'Added by hand') + '|' + amount.toFixed(2);
    if (undated[uk] && undated[uk].length) {        // the same row, saved without its date: date it, don't add it again
      const found = undated[uk].shift();
      tx.getRange(found.row, TX_COL.DATE + 1).setValue(d).setNumberFormat('yyyy-MM-dd');
      sheet.getRange(i + 2, W).setValue(found.id);
      repaired++;
      return;
    }
    const cur = String(r[5] || '').trim().toUpperCase();
    const merchant = String(r[2] || '').trim() || 'Added by hand';
    const id = 'manual:' + now.getTime() + ':' + (incoming ? 'in' : 'out') + ':' + (i + 2);
    add.push({ date: d, bank: String(r[1] || '').trim().toUpperCase() || 'MANUAL', merchant: merchant,
      amount: incoming ? -amount : amount, currency: VALID_CURRENCIES.has(cur) ? cur : 'DOP', category: String(r[3] || '').trim(),
      autoCategory: AUTO_NONE, type: type, description: merchant, reversal: false, timeKey: '', txRef: '',
      subject: 'Added by hand in ' + sheetName, timestamp: now.toISOString(), messageId: id, isCredit: incoming, isCashback: false });
    sheet.getRange(i + 2, W).setValue(id);             // never imported twice, even if the rebuild doesn't follow
  });
  if (add.length) {
    saveTransactions(add);
    // v1.1.58: read back what was saved — a row without its date is dated now, and the log says so
    const want = {};
    add.forEach(t => { want[t.messageId] = t.date; });
    const txNow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSACTIONS_SHEET);
    txNow.getRange(2, 1, txNow.getLastRow() - 1, TX_NUM_COLS).getValues().forEach((r, i) => {
      const id = String(r[TX_COL.MESSAGE_ID] || '');
      if (want[id] === undefined || r[TX_COL.DATE]) return;
      txNow.getRange(i + 2, TX_COL.DATE + 1).setValue(want[id]).setNumberFormat('yyyy-MM-dd');
      Logger.log('⚠️ ' + sheetName + ': a typed row was saved without its date and was dated again (' + id + ', ' +
        normalizeDateForCompare(want[id]) + ') — please report this line');
    });
  }
  if (repaired) Logger.log(sheetName + ': ' + repaired + ' row(s) typed again gave their date back to rows saved without one');
  if (bad.length) {
    recordUnrecognized(bad.map(r => {
      const typed = r.slice(0, 6).map(x => x instanceof Date ? normalizeDateForCompare(x) : String(x)).join(' · ');
      return { id: 'typed:' + typed.replace(/\s+/g, ''), date: now, bank: sheetName, subject: 'Typed by hand',
        reason: 'Not saved: a transfer needs a date (yyyy-mm-dd) and an amount (like 1,234.56) — type it again in ' + sheetName, snippet: typed };
    }), [], now);
  }
  if (add.length || bad.length) Logger.log(sheetName + ' typed by hand: ' + add.length + ' saved, ' + bad.length + ' without a date or amount');
  return { added: add.length, rejected: bad.length, repaired: repaired };
}

/**
 * v1.1.67: Uber ride alerts checked against Uber's trip receipts (matchRideReceipts, 03_gmailMonitor.gs). A hold — the
 * estimate of a ride whose fare changed, or a request that never became a trip — is Exclude and says so in its
 * Description; a ride charged below its alert gets the amount charged; a confirmed charge says "(Uber receipt)". Rides
 * of the last RIDE_LOOKBACK_DAYS (or since `since`, a run's range) not settled yet; their alert times come from Gmail.
 * Without receipts in Gmail nothing changes. A category you set on a hold is yours (Recategorize keeps it).
 */
const RIDE_LOOKBACK_DAYS = 35;
const RIDE_HOLD_SUFFIX = ' (Uber hold, not charged)';
const RIDE_RECEIPT_MARK = ' (Uber receipt';
function rideRowState(description) {
  const d = String(description || '');
  if (d.endsWith(RIDE_HOLD_SUFFIX)) return 'hold';
  return d.indexOf(RIDE_RECEIPT_MARK) !== -1 ? 'charged' : '';
}

function reconcileRideReceipts(now, since) {
  now = now || new Date();
  const out = { receipts: 0, charged: 0, holds: 0, adjusted: 0, unmatched: 0, holdsTotal: 0, currency: 'DOP' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSACTIONS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return out;
  const lookback = new Date(now.getFullYear(), now.getMonth(), now.getDate() - RIDE_LOOKBACK_DAYS, 12);
  const from = since instanceof Date && since < lookback ? since : lookback;
  const fromKey = normalizeDateForCompare(from);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, TX_NUM_COLS).getValues();
  const rides = [];
  values.forEach((r, i) => {
    if ((r[TX_COL.TYPE] || 'Transaction') !== 'Transaction' || !RIDE_MERCHANT_RE.test(String(r[TX_COL.MERCHANT] || ''))) return;
    if (!r[TX_COL.DATE] || normalizeDateForCompare(r[TX_COL.DATE]) < fromKey) return;
    const id = String(r[TX_COL.MESSAGE_ID] || '');
    if (!id || /^manual:/.test(id)) return;   // typed by hand: no alert email to look up
    rides.push({ i: i, gmailId: id.replace(/_\d+$/, ''), day: dayNumber(r[TX_COL.DATE]), state: rideRowState(r[TX_COL.DESCRIPTION]) });
  });
  const open = rides.filter(x => !x.state);
  if (!open.length) return out;

  const receipts = [];
  const after = Utilities.formatDate(new Date(from.getTime() - 86400000), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  GmailApp.search(RIDE_RECEIPT_QUERY + ' after:' + after, 0, 200).forEach(thread => thread.getMessages().forEach(m => {
    if (!/uber\.com/i.test(String(m.getFrom() || '')) || !RIDE_RECEIPT_SUBJECT.test(String(m.getSubject() || ''))) return;
    let text = '';
    try { text = typeof m.getBody === 'function' ? htmlToPlainText(m.getBody()) : ''; } catch (e) { text = ''; }
    const parsed = parseUberTripReceipt(text || emailPlainText(m));
    if (!parsed || !parsed.payments.length) return;
    out.receipts++;
    parsed.payments.forEach(p => receipts.push({ bank: p.bank, currency: p.currency, amount: p.amount,
      requested: parsed.requested.getTime(), sent: m.getDate().getTime() }));
  }));
  if (!receipts.length) return out;

  // the alert's own time (the sheet keeps only the day): rides on the days around an open one
  const alerts = [];
  rides.forEach(x => {
    if (!open.some(o => Math.abs(o.day - x.day) <= 1)) return;
    let at = null;
    try { at = GmailApp.getMessageById(x.gmailId).getDate().getTime(); } catch (e) { at = null; }
    if (at === null) return;
    const r = values[x.i];
    alerts.push({ id: String(x.i), bank: String(r[TX_COL.BANK] || ''), currency: String(r[TX_COL.CURRENCY] || 'DOP'),
      amount: Math.abs(Number(r[TX_COL.AMOUNT]) || 0), at: at, state: x.state });
  });
  const result = matchRideReceipts(alerts, receipts);
  Object.keys(result.decisions).forEach(k => {
    const d = result.decisions[k], i = Number(k), r = values[i];
    const desc = String(r[TX_COL.DESCRIPTION] || r[TX_COL.MERCHANT] || '');
    const cur = String(r[TX_COL.CURRENCY] || 'DOP');
    if (d.kind === 'hold') {
      r[TX_COL.DESCRIPTION] = desc + RIDE_HOLD_SUFFIX;
      r[TX_COL.CATEGORY] = EXCLUDE_CATEGORY;
      r[TX_COL.AUTO_CATEGORY] = EXCLUDE_CATEGORY;
      out.holds++;
      out.holdsTotal += Math.abs(Number(r[TX_COL.AMOUNT]) || 0);
      out.currency = cur;
    } else if (d.kind === 'adjust') {
      r[TX_COL.AMOUNT] = d.amount;
      r[TX_COL.DESCRIPTION] = desc + RIDE_RECEIPT_MARK + ': ' + cur + ' ' + d.amount.toFixed(2) + ', the alert said ' + d.was.toFixed(2) + ')';
      out.adjusted++;
    } else {
      r[TX_COL.DESCRIPTION] = desc + RIDE_RECEIPT_MARK + ')';
      out.charged++;
    }
    sheet.getRange(i + 2, 1, 1, TX_NUM_COLS).setValues([r]);
    Logger.log('Uber receipt | ' + d.kind + ' | row ' + (i + 2) + ' | ' + r[TX_COL.BANK] + ' ' + cur + ' ' + (d.was || r[TX_COL.AMOUNT]) +
      (d.kind === 'adjust' ? ' → ' + d.amount : ''));
  });
  out.unmatched = result.unmatched.length;
  out.holdsTotal = Math.round(out.holdsTotal * 100) / 100;
  if (out.unmatched) Logger.log('Uber: ' + out.unmatched + ' ride alert(s) with no trip receipt near them — kept as charges');
  return out;
}

// ====================================================================================================
// 05_dailySummary.gs
// ====================================================================================================

/**
 * DAILY SUMMARY EMAIL — v1.1.24
 *
 * A short email about the previous day, sent by its own daily trigger at the
 * hour chosen in the Setup Wizard, from the user's own Gmail account, with a
 * link to the spreadsheet. Each section can be turned on or off in the wizard.
 *
 * Numbers follow the Dashboard's rules exactly: spending = Transaction and
 * Transfer rows that have a category other than "Exclude" (Card Payment and
 * Cashback never count; reversals are negative rows and subtract), converted
 * to DOP-equivalent at the Dashboard's exchange rates. Transfers without a
 * category are listed on their own. Recommendations are fixed rules — no
 * guessing — and card tips only mention banks that have a cashback rate in the
 * Dashboard's credit-card table.
 *
 * Menu "📬 Send Daily Summary Now" sends the same email immediately (for a
 * test, or when the trigger is off).
 */
const SUMMARY_SECTIONS_DEFAULT = { totals: true, vsAverage: true, transfers: true, recommendations: true, cashback: true,
  investments: true };   // v1.1.32
const SUMMARY_MIN_HISTORY_DAYS = 7;
// v1.1.26: bills and fixed costs are left out of "day-to-day" comparisons — a rent
// or electricity payment would otherwise make any day look like an alarming spike.
const SUMMARY_BILL_CATEGORIES = ['Rent', 'Gym + Calisthenics', 'Telecommunications', 'Streaming & Subscriptions', 'Electricity'];
function isBillCategory(cat) { return SUMMARY_BILL_CATEGORIES.indexOf(cat) !== -1; }
const SUMMARY_ABOVE_USUAL = 1.05;   // v1.1.26: only flag a category once it's 5%+ over its usual month
// v1.1.64: yesterday is compared with a "typical day" — the median of the day-to-day totals of the last 60 days (days
// with nothing spent count as RD$0) — instead of the 30-day average, which one tank of fuel or one big purchase turned
// into "+730%". A day is flagged only when it ranks among the highest of those days (above the 90th percentile), and
// a purchase of several typical days is named apart so the rest of the day is compared on its own.
const SUMMARY_TYPICAL_WINDOW_DAYS = 60;
const SUMMARY_HIGH_DAY_PERCENTILE = 0.9;
const SUMMARY_SPORADIC_MULTIPLE = 3;

// Kinds of purchase with an unambiguous best card (see "Which card for what").
const CARD_TIP_RULES = [
  { label: 'Pharmacies', bank: 'BHD', test: /\b(FARMACIAS?|FCIA|FARMA\w*|FARM|MEDICAR|HIDALGOS)\b/ },
  { label: 'Supermarkets', bank: 'BANESCO', test: /\b(SM|SUPERMERCADOS?|SUPER|JUMBO|BRAVO|POLA|PLAZA LAMA|NACIONAL)\b/ },
  { label: 'Electricity and water', bank: 'BANESCO', test: /\b(EDESUR|EDENORTE|EDEESTE|CAASD)\b/ }
];

const SUMMARY_DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const SUMMARY_MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function summaryMoney(n) {
  const v = Math.round(Math.abs(Number(n) || 0)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '−' : '') + 'RD$' + v;
}
function summaryPct(x) { return Math.round(Math.abs(x) * 100) + '%'; }
function summaryEscape(t) {
  return String(t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
/**
 * v1.1.25: GmailApp.sendEmail() breaks characters outside Unicode's Basic
 * Multilingual Plane — every emoji arrived as six "�" (its surrogate pair
 * encoded as two separate 3-byte sequences), while BMP characters such as
 * — → − came through fine. So the HTML body is sent as pure ASCII (each
 * non-ASCII character as a numeric entity, which every client decodes), and
 * the subject and plain-text body, where entities don't exist, carry no
 * emoji at all.
 */
function toAsciiHtml(html) {
  let out = '';
  for (const ch of String(html)) {            // for…of walks code points, keeping surrogate pairs together
    const cp = ch.codePointAt(0);
    out += cp < 128 ? ch : '&#' + cp + ';';
  }
  return out;
}
function stripAstral(text) {
  return String(text).replace(/[\u{10000}-\u{10FFFF}]\uFE0F?[ ]?/gu, '');
}

/** v1.1.64: median of a list of numbers (the mean of the middle two for an even count); null when empty. */
function summaryMedian(list) {
  if (!list.length) return null;
  const a = list.slice().sort((x, y) => x - y), mid = Math.floor(a.length / 2);
  return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}
/** v1.1.64: nearest-rank percentile (p in 0..1) — always a value that occurred; null when empty. */
function summaryPercentile(list, p) {
  if (!list.length) return null;
  const a = list.slice().sort((x, y) => x - y);
  return a[Math.min(a.length, Math.max(1, Math.ceil(p * a.length))) - 1];
}

function keyToDate(key) {
  const m = String(key).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12) : null;
}

/** Net monthly income in DOP-equivalent — same arithmetic as the Dashboard's income box. */
function computeNetIncomeDop(config, rates) {
  const d = config.deductions || {};
  const gross = Number(config.monthlyIncome) || 0;
  const salary = config.incomeCurrency === 'DOP'
    ? gross - (d.ARS || 0) - (d.AFP || 0) - (d.ISR || 0)
    : (gross - (d.ARS || 0) - (d.AFP || 0) - gross * (d.taxRate || 0) / 100) * rates.USD;
  // v1.1.27: other income is added in full, in its own currency
  // v1.1.59: every other income line in, every other deduction line out, each at its currency's rate. A config object
  // from before (a single otherIncome) still counts it.
  const incomes = Array.isArray(config.otherIncomes) ? config.otherIncomes
    : (Number(config.otherIncome) > 0 ? [{ amount: Number(config.otherIncome), currency: config.otherIncomeCurrency === 'USD' ? 'USD' : 'DOP' }] : []);
  return salary + moneyLinesDop(incomes, rates) - moneyLinesDop(config.otherDeductions || [], rates);
}

/**
 * Everything the email needs about the day before `opts.today`. Pure (no
 * Gmail/Sheets calls) — see tests/.
 * opts: { today: Date, rates: {USD, EUR, COP}, netIncomeDop, cards: [{bank, rate}] }
 */
function computeDailySummary(values, opts) {
  const rates = opts.rates;
  const today = opts.today;
  const day = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1, 12);
  const dayKey = normalizeDateForCompare(day);
  const shiftKey = n => normalizeDateForCompare(new Date(day.getFullYear(), day.getMonth(), day.getDate() + n, 12));
  const monthKey = dayKey.slice(0, 7);
  const prevMonths = [1, 2, 3].map(k => normalizeDateForCompare(new Date(day.getFullYear(), day.getMonth() - k, 1, 12)).slice(0, 7));
  const toDop = (amount, cur) => (Number(amount) || 0) * ({ USD: rates.USD, EUR: rates.EUR, COP: rates.COP }[cur] || 1);

  const s = {
    day: day, dayKey: dayKey, spent: 0, spentDayToDay: 0, purchases: 0, items: [], openTransfers: [], openTransfersTotal: 0,
    mtd: 0, mtdFixed: 0, mtdByCategory: {}, prevByCategory: {}, prevMonthsWithData: 0,
    mtdOpenCount: 0, mtdOpenTotal: 0, firstKey: null,
    dayOfMonth: day.getDate(), daysInMonth: new Date(day.getFullYear(), day.getMonth() + 1, 0).getDate(),
    netIncomeDop: Number(opts.netIncomeDop) || 0
  };
  const monthsSeen = {};
  const windowStartKey = shiftKey(-SUMMARY_TYPICAL_WINDOW_DAYS);
  const weekStartKey = shiftKey(-6);          // v1.1.26: last 7 days incl. yesterday
  const dayTotals = {};
  const dayToDayTotals = {};                  // v1.1.64: day-to-day spending per day, for the typical day

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[TX_COL.DATE]) continue;
    const key = normalizeDateForCompare(r[TX_COL.DATE]);
    const type = r[TX_COL.TYPE] || 'Transaction';
    const cat = String(r[TX_COL.CATEGORY] || '');
    const amt = toDop(r[TX_COL.AMOUNT], r[TX_COL.CURRENCY]);
    // v1.1.51: money received with a category reduces that category (its amount is negative)
    const spend = (type === 'Transaction' || type === 'Transfer' || type === 'Incoming') && cat !== '' && cat !== EXCLUDE_CATEGORY;
    const open = type === 'Transfer' && cat === '';
    const month = key.slice(0, 7);
    const item = { merchant: String(r[TX_COL.MERCHANT] || ''), bank: String(r[TX_COL.BANK] || ''), category: cat, amount: amt, type: type };

    if (spend && (s.firstKey === null || key < s.firstKey)) s.firstKey = key;
    if (key === dayKey) {
      if (spend) {
        s.spent += amt; s.items.push(item); if (amt > 0) s.purchases++;
        if (!isBillCategory(cat)) s.spentDayToDay += amt;
      }
      if (open) { s.openTransfers.push(item); s.openTransfersTotal += amt; }
    }
    if (spend && !isBillCategory(cat) && key >= windowStartKey && key < dayKey) dayToDayTotals[key] = (dayToDayTotals[key] || 0) + amt;
    if (spend && key >= weekStartKey && key <= dayKey) dayTotals[key] = (dayTotals[key] || 0) + amt;
    if (month === monthKey && key <= dayKey) {
      if (spend) {
        s.mtd += amt; s.mtdByCategory[cat] = (s.mtdByCategory[cat] || 0) + amt;
        if (isBillCategory(cat)) s.mtdFixed += amt;
      }
      if (open) { s.mtdOpenCount++; s.mtdOpenTotal += amt; }
    }
    if (spend && prevMonths.indexOf(month) !== -1) {
      s.prevByCategory[cat] = (s.prevByCategory[cat] || 0) + amt;
      monthsSeen[month] = true;
    }
  }

  // v1.1.64: typical day over the 60 days before yesterday — or fewer, if tracking started later. Every day counts,
  // RD$0 included. Day-to-day spending only (bills and fixed costs excluded on both sides, as since v1.1.26).
  const windowStart = s.firstKey && s.firstKey > windowStartKey ? s.firstKey : windowStartKey;
  s.historyDays = Math.max(0, Math.round((day - keyToDate(windowStart)) / 86400000));
  const history = [];
  for (let n = 1; n <= s.historyDays; n++) history.push(dayToDayTotals[shiftKey(-n)] || 0);
  const enough = s.historyDays >= SUMMARY_MIN_HISTORY_DAYS;
  s.typicalDay = enough ? summaryMedian(history) : null;
  s.highDayThreshold = enough ? summaryPercentile(history, SUMMARY_HIGH_DAY_PERCENTILE) : null;
  const comparable = s.typicalDay !== null && s.typicalDay > 0;
  // A threshold of RD$0 (most days with nothing spent) would make every purchase "high": no warning then.
  s.highDay = comparable && s.highDayThreshold > 0 && s.spentDayToDay > s.highDayThreshold;
  s.vsTypical = comparable && s.spentDayToDay > 0 ? s.spentDayToDay / s.typicalDay - 1 : null;
  // Purchases of several typical days at once (fuel, an appliance): named apart, the rest compared on its own
  s.sporadic = comparable ? s.items.filter(it => !isBillCategory(it.category) && it.amount >= s.typicalDay * SUMMARY_SPORADIC_MULTIPLE)
    .sort((a, b) => b.amount - a.amount).map(it => ({ merchant: it.merchant, amount: it.amount })) : [];
  s.restDayToDay = s.spentDayToDay - s.sporadic.reduce((sum, it) => sum + it.amount, 0);
  s.vsTypicalRest = comparable ? s.restDayToDay / s.typicalDay - 1 : null;
  s.prevMonthsWithData = Object.keys(monthsSeen).length;
  s.usualByCategory = {};
  if (s.prevMonthsWithData) {
    Object.keys(s.prevByCategory).forEach(c => { s.usualByCategory[c] = s.prevByCategory[c] / s.prevMonthsWithData; });
  }
  // v1.1.65: bills and fixed costs are paid once a month, mostly at its start: they count once, not times the days
  // left (a rent paid on day 1 made day 5 project the month at 6.2 rents). Bills of a usual month not paid yet are
  // added, so the pace is not low before rent day. Only day-to-day spending is extrapolated.
  s.mtdDayToDay = s.mtd - s.mtdFixed;
  s.pendingFixed = SUMMARY_BILL_CATEGORIES
    .reduce((sum, c) => sum + Math.max(0, (s.usualByCategory[c] || 0) - (s.mtdByCategory[c] || 0)), 0);
  s.projected = s.dayOfMonth ? s.mtdFixed + s.pendingFixed + s.mtdDayToDay / s.dayOfMonth * s.daysInMonth : 0;
  s.mtdPct = s.netIncomeDop > 0 ? s.mtd / s.netIncomeDop : null;
  s.projectedPct = s.netIncomeDop > 0 ? s.projected / s.netIncomeDop : null;
  s.items.sort((a, b) => b.amount - a.amount);
  // v1.1.26: extra material for the redesigned email
  s.elapsedPct = s.dayOfMonth / s.daysInMonth;
  s.last7 = [-6, -5, -4, -3, -2, -1, 0].map(n => {
    const date = new Date(day.getFullYear(), day.getMonth(), day.getDate() + n, 12);
    const key = normalizeDateForCompare(date);
    return { key: key, date: date, amount: dayTotals[key] || 0 };
  });
  s.mtdTop = Object.keys(s.mtdByCategory)
    .filter(c => !isBillCategory(c))
    .map(c => ({ cat: c, mtd: s.mtdByCategory[c], usual: s.usualByCategory[c] || 0 }))
    .filter(c => c.mtd > 0).sort((a, b) => b.mtd - a.mtd).slice(0, 5);
  s.mtdBills = Object.keys(s.mtdByCategory).filter(isBillCategory)
    .reduce((sum, c) => sum + s.mtdByCategory[c], 0);
  s.lastRun = opts.lastRun || null;
  s.now = opts.now || today;
  s.investments = opts.investments || null;   // v1.1.32
  s.recommendations = summaryRecommendations(s);
  s.tips = summaryCardTips(s.items, opts.cards || []);
  return s;
}

/** Fixed rules → [{icon, text}]. */
function summaryRecommendations(s) {
  const out = [];
  if (s.spent <= 0) {
    out.push({ icon: '🟢', text: 'No spending logged yesterday.' });
  } else if (s.highDay) {
    // v1.1.64: from the day's rank among the last 60 days, not from a percentage over the average
    const top = s.items.filter(it => !isBillCategory(it.category))[0];
    out.push({ icon: '🟠', text: 'Day-to-day spending (' + summaryMoney(s.spentDayToDay) + ') was one of your highest days in the last ' +
      s.historyDays + ' days (the top ' + summaryPct(1 - SUMMARY_HIGH_DAY_PERCENTILE) + ' start above ' + summaryMoney(s.highDayThreshold) +
      '). A typical day is ' + summaryMoney(s.typicalDay) + '.' + (s.sporadic.length ? ' It includes ' + summarySporadicText(s.sporadic) + '.'
      : ' Biggest item: ' + top.merchant + ' (' + summaryMoney(top.amount) + ').') });
  } else if (s.vsTypical !== null && s.vsTypical < -0.3) {
    out.push({ icon: '🟢', text: 'Nice — day-to-day spending was ' + summaryPct(s.vsTypical) + ' below a typical day (' +
      summaryMoney(s.typicalDay) + ').' });
  }
  if (s.projectedPct !== null && s.dayOfMonth >= 5) {
    if (s.projectedPct > 1) {
      out.push({ icon: '🔴', text: 'At this pace the month ends around ' + summaryMoney(s.projected) + ' — ' +
        summaryPct(s.projectedPct) + ' of your net income.' });
    } else if (s.projectedPct > 0.85) {
      out.push({ icon: '🟠', text: 'Month pace: about ' + summaryPct(s.projectedPct) + ' of your net income by month-end.' });
    } else {
      out.push({ icon: '🟢', text: 'On track: at this pace the month ends at about ' + summaryPct(s.projectedPct) +
        ' of your net income.' });
    }
  }
  Object.keys(s.mtdByCategory)
    .filter(c => !isBillCategory(c))
    .map(c => ({ cat: c, mtd: s.mtdByCategory[c], usual: s.usualByCategory[c] || 0 }))
    .filter(x => x.usual > 0 && x.mtd > x.usual * SUMMARY_ABOVE_USUAL)
    .sort((a, b) => b.mtd / b.usual - a.mtd / a.usual)
    .slice(0, 2)
    .forEach(x => out.push({ icon: '🟠', text: x.cat + ': ' + summaryMoney(x.mtd) +
      ' so far this month, already above your usual month (' + summaryMoney(x.usual) + ').' }));
  if (s.mtdOpenCount > 0) {
    out.push({ icon: '📝', text: s.mtdOpenCount + ' transfer(s) this month have no category (' + summaryMoney(s.mtdOpenTotal) +
      ') — add a Custom Rule, or use "Exclude" for transfers between your own accounts.' });
  }
  return out;
}

/** v1.1.64: "SHELL, RD$2,500; FERRETERIA, RD$1,500" — the first three, then "and N more". */
function summarySporadicText(list) {
  const shown = list.slice(0, 3).map(it => it.merchant + ', ' + summaryMoney(it.amount)).join('; ');
  return shown + (list.length > 3 ? ' and ' + (list.length - 3) + ' more' : '');
}

/**
 * v1.1.64: how yesterday compares with a typical day, for the hero chip, the line under it and the subject.
 * No percentage when there is too little history or the typical day is RD$0; big one-off purchases are named and the
 * rest of the day is compared on its own. Returns {chip: {text, tone}, line, subject}.
 */
function summaryTypicalView(s) {
  const typical = summaryMoney(s.typicalDay);
  if (s.spent <= 0) return { chip: { text: 'No spending logged', tone: 'good' }, line: '', subject: '' };
  if (s.spentDayToDay <= 0) return { chip: { text: 'Only bills and fixed costs', tone: 'info' }, line: '', subject: '' };
  if (s.typicalDay === null) {
    return { chip: { text: 'Typical day available after ' + SUMMARY_MIN_HISTORY_DAYS + ' days of history', tone: 'note' }, line: '', subject: '' };
  }
  if (!(s.typicalDay > 0)) {
    return { chip: { text: 'Most days have no day-to-day spending — no typical day to compare with yet', tone: 'note' }, line: '', subject: '' };
  }
  const chip = s.highDay ? { text: 'One of your highest days in the last ' + s.historyDays + ' days', tone: 'warn' }
    : s.vsTypical < -0.1 ? { text: 'Below a typical day', tone: 'good' }
    : s.vsTypical > 0.1 ? { text: 'Above a typical day, not unusual', tone: 'info' }
    : { text: 'Close to a typical day', tone: 'info' };
  if (s.sporadic.length) {
    const first = s.sporadic[0];
    return { chip: chip,
      line: 'Includes ' + summarySporadicText(s.sporadic) + ' · ' + (s.restDayToDay > 0
        ? 'the rest of the day (' + summaryMoney(s.restDayToDay) + ') ' + summarySignedPct(s.vsTypicalRest) + ' vs. a typical day (' + typical + ')'
        : 'nothing else day-to-day (a typical day is ' + typical + ')'),
      subject: ' (incl. ' + first.merchant + ' ' + summaryMoney(first.amount) + (s.sporadic.length > 1 ? ' and ' + (s.sporadic.length - 1) + ' more' : '') + ')' };
  }
  return { chip: chip, line: summarySignedPct(s.vsTypical) + ' vs. a typical day (' + typical + ')',
    subject: ' (' + summarySignedPct(s.vsTypical) + ' vs. typical day)' };
}

/** Card tips for yesterday's purchases, only for banks with a cashback rate in the card table. */
function summaryCardTips(items, cards) {
  const rateByBank = {}, nameByBank = {};
  cards.forEach(c => { if (c.rate > 0) { rateByBank[c.bank] = c.rate; nameByBank[c.bank] = c.name || ''; } });
  const tips = [];
  const seen = {};
  items.forEach(it => {
    if (it.type !== 'Transaction' || !(it.amount > 0) || seen[it.merchant]) return;
    const text = normalizeKeyword(it.merchant);
    const rule = CARD_TIP_RULES.find(r => r.test.test(text));
    if (!rule || !rateByBank[rule.bank] || it.bank === rule.bank) return;
    seen[it.merchant] = true;
    const rate = rateByBank[rule.bank];
    tips.push({ text: it.merchant + ' (' + summaryMoney(it.amount) + ' on ' + it.bank + '): ' + rule.label.toLowerCase() +
      ' earn ' + summaryPct(rate) + ' with your ' + rule.bank + (nameByBank[rule.bank] ? ' ' + nameByBank[rule.bank] : '') + ' card — about ' + summaryMoney(it.amount * rate) + ' back.' });
  });
  return tips;
}

/**
 * EMAIL KIT — v1.1.26. Small building blocks shared by the daily and the
 * monthly summary. Email clients (Gmail above all) only reliably render
 * TABLE layout with INLINE styles: no flexbox/grid, no <style> classes, no
 * scripts, no external images. Every block returns an HTML string; text
 * passed in is escaped here. The final HTML goes through toAsciiHtml().
 */
const EK = {
  font: "-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif",
  navy: '#1F3864', accent: '#4472C4', text: '#1F2937', muted: '#6B7280', faint: '#9CA3AF',
  line: '#E5E7EB', soft: '#F5F7FB', page: '#EEF1F6', track: '#E8ECF3',
  tones: {
    good: { fg: '#2E7D32', bg: '#E8F5E9', icon: '🟢' },
    warn: { fg: '#B45309', bg: '#FFF4E5', icon: '🟠' },
    bad:  { fg: '#B91C1C', bg: '#FDECEC', icon: '🔴' },
    info: { fg: '#1D4ED8', bg: '#EAF1FE', icon: '🔵' },
    note: { fg: '#4B5563', bg: '#F3F4F6', icon: '📝' }
  },
  heat: ['#F3F4F6', '#DBE6FA', '#B4CCF3', '#7FA6E8', '#4472C4', '#1F3864']
};
const SUMMARY_DAYS_FULL = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const SUMMARY_MONTHS_FULL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
                             'September', 'October', 'November', 'December'];

function ekTone(name) { return EK.tones[name] || EK.tones.note; }

function ekChip(text, tone) {
  const t = ekTone(tone);
  return '<span style="display:inline-block;padding:3px 10px;border-radius:999px;font-size:12px;font-weight:600;' +
    'line-height:18px;background:' + t.bg + ';color:' + t.fg + '">' + summaryEscape(text) + '</span>';
}

/** A section row inside the card: small caps title + content. */
function ekSection(title, inner) {
  return '<tr><td style="padding:22px 28px 0">' +
    '<div style="font-size:11px;font-weight:700;letter-spacing:.8px;text-transform:uppercase;color:' + EK.muted +
    ';margin:0 0 10px">' + summaryEscape(title) + '</div>' + inner + '</td></tr>';
}

/** Tiles in one row: [{label, value, sub, tone}] */
function ekKpis(tiles) {
  const w = Math.floor(100 / tiles.length);
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;' +
    'border-spacing:0"><tr>' + tiles.map((t, i) =>
      '<td width="' + w + '%" valign="top" style="padding:0 ' + (i < tiles.length - 1 ? '8px' : '0') + ' 0 0">' +
      '<div style="background:' + EK.soft + ';border-radius:10px;padding:12px 12px 11px">' +
      '<div style="font-size:11px;color:' + EK.muted + ';text-transform:uppercase;letter-spacing:.5px">' + summaryEscape(t.label) + '</div>' +
      '<div style="font-size:19px;font-weight:700;color:' + (t.tone ? ekTone(t.tone).fg : EK.navy) + ';margin:4px 0 2px;' +
      'white-space:nowrap">' + summaryEscape(t.value) + '</div>' +
      '<div style="font-size:12px;color:' + EK.muted + '">' + summaryEscape(t.sub || '') + '</div></div></td>').join('') +
    '</tr></table>';
}

/** Labelled horizontal bar: fraction 0..1 (clamped), colour, right-hand text. */
function ekBar(label, fraction, color, right, sub) {
  const pct = Math.max(0, Math.min(1, Number(fraction) || 0));
  const fill = pct > 0 ? '<div style="width:' + Math.max(2, Math.round(pct * 100)) + '%;height:8px;background:' + color +
    ';border-radius:999px;font-size:0;line-height:0">&nbsp;</div>' : '';
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 12px">' +
    '<tr><td style="font-size:14px;color:' + EK.text + ';padding:0 0 5px">' + summaryEscape(label) + '</td>' +
    '<td align="right" style="font-size:14px;font-weight:600;color:' + EK.text + ';padding:0 0 5px;white-space:nowrap">' +
    summaryEscape(right) + '</td></tr>' +
    '<tr><td colspan="2" style="background:' + EK.track + ';border-radius:999px;height:8px;font-size:0;line-height:0">' +
    fill + '</td></tr>' + (sub ? '<tr><td colspan="2" style="font-size:12px;color:' + EK.muted + ';padding:4px 0 0">' + sub +
    '</td></tr>' : '') + '</table>';
}

/** Vertical mini columns: [{label, sub, value, highlight}] — max height 70 px. */
function ekColumns(cols) {
  const max = Math.max.apply(null, cols.map(c => c.value).concat([1]));
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' + cols.map(c => {
    const h = c.value > 0 ? Math.max(4, Math.round(c.value / max * 70)) : 2;
    const color = c.highlight ? EK.accent : (c.value > 0 ? '#B4CCF3' : EK.line);
    return '<td valign="bottom" align="center" style="padding:0 3px;height:92px">' +
      '<div style="font-size:10px;color:' + (c.highlight ? EK.navy : EK.faint) + ';font-weight:' + (c.highlight ? 700 : 400) +
      ';margin:0 0 3px;white-space:nowrap">' + summaryEscape(c.sub) + '</div>' +
      '<div style="height:' + h + 'px;background:' + color + ';border-radius:5px 5px 2px 2px;font-size:0;line-height:0">&nbsp;</div></td>';
  }).join('') + '</tr><tr>' + cols.map(c =>
    '<td align="center" style="font-size:11px;padding:6px 0 0;color:' + (c.highlight ? EK.navy : EK.muted) +
    ';font-weight:' + (c.highlight ? 700 : 400) + '">' + summaryEscape(c.label) + '</td>').join('') + '</tr></table>';
}

/** List rows: [{title, meta, right, rightTone}] */
function ekList(rows) {
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + rows.map((r, i) =>
    '<tr><td style="padding:9px 0;' + (i ? 'border-top:1px solid ' + EK.line + ';' : '') + '">' +
    '<div style="font-size:14px;font-weight:600;color:' + EK.text + '">' + summaryEscape(r.title) + '</div>' +
    (r.meta ? '<div style="font-size:12px;color:' + EK.muted + ';margin-top:2px">' + summaryEscape(r.meta) + '</div>' : '') +
    '</td><td align="right" valign="top" style="padding:9px 0 9px 12px;white-space:nowrap;font-size:14px;font-weight:700;color:' +
    (r.rightTone ? ekTone(r.rightTone).fg : EK.text) + ';' + (i ? 'border-top:1px solid ' + EK.line + ';' : '') + '">' +
    summaryEscape(r.right) + '</td></tr>').join('') + '</table>';
}

/** Coloured note cards: [{tone, text}] */
function ekNotes(notes) {
  return notes.map(n => {
    const t = ekTone(n.tone);
    return '<div style="border-left:4px solid ' + t.fg + ';background:' + t.bg + ';padding:10px 12px;border-radius:6px;' +
      'margin:0 0 8px;font-size:14px;line-height:1.45;color:' + EK.text + '">' + t.icon + '&nbsp; ' + summaryEscape(n.text) + '</div>';
  }).join('');
}

/** Buttons: [{text, url, primary}] (entries without a url are skipped). */
function ekButtons(buttons) {
  return buttons.filter(b => b.url).map(b =>
    '<a href="' + summaryEscape(b.url) + '" style="display:inline-block;margin:0 8px 8px 0;padding:11px 16px;border-radius:8px;' +
    'font-size:14px;font-weight:600;text-decoration:none;' + (b.primary
      ? 'background:' + EK.accent + ';color:#FFFFFF;border:1px solid ' + EK.accent
      : 'background:#FFFFFF;color:' + EK.accent + ';border:1px solid #C7D5EE') + '">' + summaryEscape(b.text) + '</a>').join('');
}

/** The whole email: hidden preheader, header band, content rows, footer. */
function ekShell(o) {
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<meta name="color-scheme" content="light only"><title>' + summaryEscape(o.title) + '</title></head>' +
    '<body style="margin:0;padding:0;background:' + EK.page + '">' +
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0;color:' + EK.page + '">' + summaryEscape(o.preheader || '') +
    '&#8199;&#65279;&#847;&#8199;&#65279;&#847;&#8199;&#65279;&#847;&#8199;&#65279;&#847;</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + EK.page + '"><tr>' +
    '<td align="center" style="padding:24px 10px">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;background:#FFFFFF;' +
    'border-radius:14px;overflow:hidden;font-family:' + EK.font + ';color:' + EK.text + '">' +
    '<tr><td style="background:' + EK.navy + ';padding:22px 28px 20px">' +
    '<div style="font-size:12px;color:#B4CCF3;letter-spacing:.6px;text-transform:uppercase">' + summaryEscape(o.eyebrow) + '</div>' +
    '<div style="font-size:22px;font-weight:700;color:#FFFFFF;margin-top:4px">' + summaryEscape(o.title) + '</div>' +
    (o.subtitle ? '<div style="font-size:13px;color:#C9D6EE;margin-top:3px">' + summaryEscape(o.subtitle) + '</div>' : '') +
    '</td></tr>' + o.body +
    '<tr><td style="padding:22px 28px 26px">' + (o.buttons || '') + '</td></tr></table>' +
    '<div style="max-width:600px;font-family:' + EK.font + ';font-size:11px;line-height:1.5;color:' + EK.faint +
    ';padding:14px 12px 0">' + (o.footer || '') + '</div></td></tr></table></body></html>';
}

function summaryLongDate(d) {
  return SUMMARY_DAYS_FULL[d.getDay()] + ' ' + d.getDate() + ' ' + SUMMARY_MONTHS_FULL[d.getMonth()] + ' ' + d.getFullYear();
}
function summaryShortDate(d) {
  return SUMMARY_DAYS[d.getDay()] + ' ' + d.getDate() + ' ' + SUMMARY_MONTHS[d.getMonth()];
}
function summaryTime(d) {
  const h = d.getHours(), m = d.getMinutes();
  return (h % 12 || 12) + ':' + (m < 10 ? '0' : '') + m + (h < 12 ? ' AM' : ' PM');
}
/** Compact money for small spaces: RD$1.2K, RD$45K, RD$1.3M. */
function summaryMoneyShort(n) {
  const a = Math.abs(Number(n) || 0);
  const sign = n < 0 ? '−' : '';
  if (a >= 1e6) return sign + 'RD$' + (a / 1e6).toFixed(1).replace(/\.0$/, '') + 'M';
  if (a >= 1e4) return sign + 'RD$' + Math.round(a / 1e3) + 'K';
  if (a >= 1e3) return sign + 'RD$' + (a / 1e3).toFixed(1).replace(/\.0$/, '') + 'K';
  return sign + 'RD$' + Math.round(a);
}
/** v1.1.32: US$ amounts for the investments section. */
function summaryUsd(n) {
  const v = Number(n) || 0;
  return (v < 0 ? '−' : '') + 'US$' + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function summarySignedUsd(n) { return (n > 0 ? '+' : '') + summaryUsd(n); }
/** Investment moves are mostly under 1% a day: two decimals (spending percentages stay whole numbers). */
function summarySignedPctFine(x) { const v = Number(x) || 0; return (v >= 0 ? '+' : '−') + Math.abs(v * 100).toFixed(2) + '%'; }

function summarySignedPct(x) {
  if (Math.round(Math.abs(x) * 100) === 0) return '±0%';   // v1.1.26: no "−0%"
  return (x >= 0 ? '+' : '−') + summaryPct(x);
}

/**
 * v1.1.26: one line about the morning update (written by runGmailMonitorCore()
 * into Script Properties) — so the email says when the data is stale or when
 * an email couldn't be read. Returns {tone, text}.
 */
function summaryDataHealth(lastRun, now) {
  if (!lastRun || !lastRun.at) {
    return { tone: 'note', text: 'Update status appears here after the next automatic morning update.' };
  }
  const at = new Date(lastRun.at);
  const hours = (now - at) / 3600000;
  const when = (hours < 24 && at.getDate() === now.getDate() ? 'today' : summaryShortDate(at)) + ' at ' + summaryTime(at);
  if (hours > 26) {
    return { tone: 'warn', text: 'No automatic update since ' + when + ' — numbers may be incomplete. Check Extensions › Apps Script › Executions.' };
  }
  let text = 'Data updated ' + when + ' · ' + (lastRun.saved || 0) + ' new transaction(s)';
  if (lastRun.unparsed) text += ' · ' + lastRun.unparsed + ' email(s) couldn\'t be read (left unread in Gmail)';
  if (lastRun.errors) text += ' · ' + lastRun.errors + ' error(s) in the log';
  if (lastRun.unrecognized) text += ' · ' + lastRun.unrecognized + ' email(s) in Unrecognized';   // v1.1.35
  return { tone: lastRun.unparsed || lastRun.errors || lastRun.unrecognized ? 'warn' : 'good', text: text };
}

/** Subject, HTML and plain-text bodies of the daily email. Pure — see tests/. */
function buildDailySummaryEmail(s, opts) {
  opts = opts || {};
  const sec = Object.assign({}, SUMMARY_SECTIONS_DEFAULT, opts.sections || {});
  const links = opts.links || { dashboard: opts.sheetUrl };
  const d = s.day;
  const view = summaryTypicalView(s);   // v1.1.64
  const subject = summaryShortDate(d) + ' ' + d.getFullYear() + ': ' + summaryMoney(s.spent) + ' spent' + view.subject;
  const rows = [], T = [];

  // ---- hero
  if (sec.totals) {
    const chip = ekChip(view.chip.text, view.chip.tone) + (view.line ? '<div style="font-size:13px;color:' + EK.muted +
      ';margin-top:8px">' + summaryEscape(view.line) + '</div>' : '');
    rows.push('<tr><td style="padding:24px 28px 0">' +
      '<div style="font-size:13px;color:' + EK.muted + '">Spent yesterday · ' + s.purchases + ' purchase(s)</div>' +
      '<div style="font-size:40px;font-weight:800;color:' + EK.navy + ';line-height:1.15;margin:4px 0 10px">' +
      summaryMoney(s.spent) + '</div>' + chip + '</td></tr>');
    T.push('Yesterday (' + summaryLongDate(d) + '): ' + summaryMoney(s.spent) + ' · ' + s.purchases + ' purchase(s)' +
      (s.typicalDay !== null ? ' · typical day ' + summaryMoney(s.typicalDay) : '') + ' · ' + view.chip.text);
    if (view.line) T.push(view.line);
  }

  // ---- month KPIs + pace
  if (sec.vsAverage) {
    const tiles = [{ label: 'This month', value: summaryMoneyShort(s.mtd),
      sub: s.mtdPct !== null ? summaryPct(s.mtdPct) + ' of net income' : 'day ' + s.dayOfMonth + ' of ' + s.daysInMonth }];
    tiles.push({ label: 'Month-end pace', value: s.dayOfMonth >= 5 ? summaryMoneyShort(s.projected) : '—',
      sub: s.dayOfMonth < 5 ? 'from day 5' : (s.projectedPct !== null ? summaryPct(s.projectedPct) + ' of net income' : 'projected'),
      tone: s.projectedPct !== null && s.dayOfMonth >= 5 ? (s.projectedPct > 1 ? 'bad' : s.projectedPct > 0.85 ? 'warn' : null) : null });
    tiles.push({ label: 'Typical day', value: s.typicalDay !== null ? summaryMoneyShort(s.typicalDay) : '—',
      sub: s.typicalDay !== null ? 'median of ' + s.historyDays + ' days' : 'not enough history yet' });
    let inner = ekKpis(tiles);
    if (s.mtdPct !== null) {
      const used = s.mtdPct, elapsed = s.elapsedPct;
      const color = used > 1 ? ekTone('bad').fg : used > elapsed + 0.05 ? ekTone('warn').fg : ekTone('good').fg;
      inner += '<div style="height:14px;font-size:0">&nbsp;</div>' +
        ekBar('Net income used', used, color, summaryPct(used)) +
        ekBar('Month elapsed', elapsed, '#C7D5EE', summaryPct(elapsed));
    }
    rows.push(ekSection('This month', inner));
    T.push('', 'This month so far: ' + summaryMoney(s.mtd) + (s.mtdPct !== null ? ' (' + summaryPct(s.mtdPct) + ' of net income)' : ''));
    if (s.dayOfMonth >= 5) T.push('At this pace: ' + summaryMoney(s.projected) + (s.projectedPct !== null ? ' (' + summaryPct(s.projectedPct) + ' of net income)' : '') + ' by month-end');
    if (s.typicalDay === null) T.push('Typical day: not enough history yet (needs ' + SUMMARY_MIN_HISTORY_DAYS + ' days).');
  }

  // ---- last 7 days + purchases
  if (sec.totals) {
    rows.push(ekSection('Last 7 days', ekColumns(s.last7.map((x, i) => ({
      label: SUMMARY_DAYS[x.date.getDay()], sub: x.amount ? summaryMoneyShort(x.amount) : '·',
      value: Math.max(0, x.amount), highlight: i === s.last7.length - 1 })))));
    if (s.items.length) {
      const shown = s.items.slice(0, 6);
      let inner = ekList(shown.map(it => ({ title: it.merchant, meta: it.bank + ' · ' + it.category,
        right: summaryMoney(it.amount), rightTone: it.amount < 0 ? 'good' : null })));
      if (s.items.length > shown.length) {
        inner += '<div style="font-size:12px;color:' + EK.muted + ';padding-top:6px">+ ' + (s.items.length - shown.length) + ' more in Transactions</div>';
      }
      rows.push(ekSection('Yesterday\'s purchases', inner));
      shown.forEach(it => T.push('  • ' + it.merchant + ' — ' + it.bank + ' — ' + summaryMoney(it.amount)));
    }
  }

  // ---- categories this month vs usual
  if (sec.vsAverage && s.mtdTop.length) {
    const max = Math.max.apply(null, s.mtdTop.map(c => Math.max(c.mtd, c.usual || 0)).concat([1]));
    rows.push(ekSection('Where this month is going', s.mtdTop.map(c => {
      const over = c.usual > 0 && c.mtd > c.usual * SUMMARY_ABOVE_USUAL;
      const sub = c.usual > 0 ? (over ? '<span style="color:' + ekTone('warn').fg + '">above usual ' +
        summaryEscape(summaryMoneyShort(c.usual)) + '</span>' : 'usual ' + summaryEscape(summaryMoneyShort(c.usual))) : '';
      return ekBar(c.cat, c.mtd / max, over ? ekTone('warn').fg : EK.accent, summaryMoney(c.mtd), sub);
    }).join('') + (s.mtdBills > 0 ? '<div style="font-size:12px;color:' + EK.muted + '">Plus bills and fixed costs: ' +
      summaryEscape(summaryMoney(s.mtdBills)) + ' (rent, gym, phone, subscriptions, electricity).</div>' : '')));
  }

  // ---- transfers to review
  if (sec.transfers) {
    let inner;
    if (s.openTransfers.length) {
      inner = ekList(s.openTransfers.map(t => ({ title: t.merchant, meta: t.bank + ' · no category', right: summaryMoney(t.amount) })));
    } else {
      inner = '<div style="font-size:14px;color:' + EK.muted + '">None yesterday.' +
        (s.mtdOpenCount ? ' ' + s.mtdOpenCount + ' still open this month (' + summaryMoney(s.mtdOpenTotal) + ').' : '') + '</div>';
    }
    rows.push(ekSection('Transfers to review', inner));
    T.push('', 'Transfers to review: ' + (s.openTransfers.length
      ? s.openTransfers.map(t => t.merchant + ' ' + summaryMoney(t.amount)).join(', ') : 'none yesterday'));
  }

  // ---- recommendations + tips
  if (sec.recommendations) {
    const tone = { '🟢': 'good', '🟠': 'warn', '🔴': 'bad', '📝': 'note' };
    const recs = s.recommendations.length ? s.recommendations : [{ icon: '🟢', text: 'Nothing to flag.' }];
    rows.push(ekSection('Recommendations', ekNotes(recs.map(r => ({ tone: tone[r.icon] || 'info', text: r.text })))));
    T.push('', 'Recommendations:', ...recs.map(r => '  • ' + r.text));
  }
  if (sec.cashback) {
    rows.push(ekSection('Card & cashback tips', s.tips.length
      ? ekNotes(s.tips.map(t => ({ tone: 'info', text: t.text })))
      : '<div style="font-size:14px;color:' + EK.muted + '">No card tips for yesterday.</div>'));
    T.push('', 'Card tips:', ...(s.tips.length ? s.tips.map(t => '  • ' + t.text) : ['  • none']));
  }

  // ---- investments (v1.1.32)
  if (sec.investments && s.investments) {
    const iv = s.investments;
    const since = iv.prevDay ? summaryShortDate(keyToDate(iv.prevDay)) : null;
    let inner = ekKpis([
      { label: 'Portfolio', value: summaryUsd(iv.total), sub: iv.totalDop ? summaryMoney(iv.totalDop) : '' },
      { label: since ? 'Since ' + since : 'Change', value: iv.change === null ? '—' : summarySignedUsd(iv.change),
        sub: iv.change === null ? 'first day recorded' : summarySignedPctFine(iv.changePct || 0) + (iv.deposits ? ' · deposits left out' : '') +
          (iv.added && iv.added.length ? ' · new account(s) not counted' : ''),
        tone: iv.change > 0 ? 'good' : iv.change < 0 ? 'bad' : null },
      { label: 'Since tracking began', value: iv.returns && iv.returns.periodReturn !== null ? summarySignedPctFine(iv.returns.periodReturn) : '—',
        sub: iv.returns ? summarySignedUsd(iv.returns.gain) : '' }
    ]);
    if (iv.movers.length) {
      inner += '<div style="height:10px;font-size:0">&nbsp;</div>' + ekList(iv.movers.map(m => ({
        title: m.ticker, meta: m.account + (m.session ? ' · last session' : ' · since ' + (since || 'the last update')),   // v1.1.56
        right: summarySignedPctFine(m.change), rightTone: m.change >= 0 ? 'good' : 'bad' })));
    }
    if (iv.stale) {   // v1.1.56
      inner += '<div style="font-size:12px;color:' + EK.muted + ';margin-top:8px">' + iv.stale +
        ' position(s) had no live price at the update — valued at their last known price.</div>';
    }
    rows.push(ekSection('Investments', inner));
    T.push('', 'Investments: ' + summaryUsd(iv.total) + (iv.change !== null ? ' (' + summarySignedUsd(iv.change) + ' since ' + since + ')' : ''));
  }

  // ---- data health (always)
  const health = summaryDataHealth(s.lastRun, s.now || new Date());
  rows.push('<tr><td style="padding:20px 28px 0"><div style="font-size:12px;color:' + ekTone(health.tone).fg + ';background:' +
    ekTone(health.tone).bg + ';border-radius:8px;padding:9px 12px">' + summaryEscape(health.text) + '</div></td></tr>');
  T.push('', health.text);

  const name = opts.sheetName || 'the spreadsheet';
  const buttons = ekButtons([
    { text: 'Open ' + name + ' →', url: links.dashboard, primary: true },
    { text: 'Transactions', url: links.transactions },
    { text: 'Custom Rules', url: sec.transfers && s.mtdOpenCount ? links.rules : null }
  ]);
  if (links.dashboard) T.push('', 'Open ' + name + ': ' + links.dashboard);

  const preheader = summaryMoney(s.spent) + ' yesterday' + (s.mtdPct !== null
    ? ' · ' + summaryPct(s.mtdPct) + ' of net income used with ' + summaryPct(s.elapsedPct) + ' of the month gone' : '');
  const html = ekShell({
    eyebrow: 'Daily summary', title: summaryLongDate(d), subtitle: 'Amounts in DOP-equivalent',
    preheader: preheader, body: rows.join(''), buttons: buttons,
    footer: 'Financial Tracker · amounts converted at the Dashboard\'s exchange rates. ' +
      'Change what this email includes, or turn it off, in 📊 Tracker › Setup Wizard.'
  });
  return { subject: stripAstral(subject), html: toAsciiHtml(html), text: stripAstral(T.join('\n')), preheader: preheader };
}

/** Exchange rates as set on the Dashboard (defaults before the Dashboard exists). */
function readDashboardRates(ss) {
  const get = (name, key) => {
    const r = ss.getRangeByName(name);
    const v = r ? Number(r.getValue()) : 0;
    return v > 0 ? v : DASH_DEFAULT_RATES[key];
  };
  return { USD: get('RATE_USD', 'usd'), EUR: get('RATE_EUR', 'eur'), COP: get('RATE_COP', 'cop') };
}

/**
 * The user's cards: [{bank, rate, name}]. v1.1.27: from the Setup Wizard
 * (Configuration "cards" + the card catalogue); before the wizard has cards, from
 * an older Dashboard's card table.
 */
function readDashboardCards(ss) {
  const config = getConfig();
  if (config && Array.isArray(config.cards)) {
    return resolveCards(config.cards).map(c => ({ bank: c.bank, rate: c.cashback, name: c.name }));
  }
  const r = ss.getRangeByName('DASH_CARDS');
  if (!r) return [];
  return r.getValues()
    .map(row => ({ bank: String(row[0] || '').trim().toUpperCase(), rate: Number(row[1]) || 0 }))
    .filter(c => c.bank);
}

/** v1.1.26: deep links to the sheets the emails point at (missing sheets are left out). */
function summaryLinks(ss) {
  const url = ss.getUrl();
  const gid = name => { const sh = ss.getSheetByName(name); return sh ? url + '#gid=' + sh.getSheetId() : null; };
  return { dashboard: gid('Dashboard') || url, transactions: gid(TRANSACTIONS_SHEET), transfers: gid('Bank Transfers'),
           rules: gid(CUSTOM_RULES_SHEET) };
}

/** v1.1.26: outcome of the last monitor run, saved by runGmailMonitorCore(). */
const LAST_RUN_PROPERTY = 'FT_LAST_RUN';
function recordLastRun(info) {
  try {
    PropertiesService.getScriptProperties().setProperty(LAST_RUN_PROPERTY, JSON.stringify(info));
  } catch (error) {
    Logger.log('Could not record the run status: ' + error);
  }
}
function readLastRun() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(LAST_RUN_PROPERTY);
    return raw ? JSON.parse(raw) : null;
  } catch (error) {
    return null;
  }
}

/** Builds and sends the email. Returns { to, subject }. */
function deliverDailySummary(now) {
  refreshAutoDeductions();
  const config = getConfig();
  if (!config) throw new Error('Setup not completed.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  const values = sheet ? sheet.getDataRange().getValues() : [[]];
  const rates = readDashboardRates(ss);
  let investments = null;   // v1.1.32
  try { investments = investmentsReportData('daily', { rates: rates }); } catch (error) { Logger.log('Investments brief: ' + error); }
  const summary = computeDailySummary(values, {
    today: now || new Date(), rates: rates, netIncomeDop: computeNetIncomeDop(config, rates), cards: readDashboardCards(ss),
    lastRun: readLastRun(), now: now || new Date(), investments: investments
  });
  const mail = buildDailySummaryEmail(summary, {
    sections: config.notifySections, sheetName: ss.getName(), links: summaryLinks(ss)
  });
  const to = config.notifyEmail || config.email;
  GmailApp.sendEmail(to, mail.subject, mail.text, { htmlBody: mail.html, name: 'Financial Tracker' });
  Logger.log('📬 Daily summary sent to ' + to + ' — ' + mail.subject);
  return { to: to, subject: mail.subject };
}

/** Trigger handler (daily, at the hour chosen in the Setup Wizard). */
function sendDailySummary() {
  try {
    const config = getConfig();
    if (!config || !config.notifyEnabled) {
      Logger.log('Daily summary is turned off — nothing sent.');
      return;
    }
    deliverDailySummary();
  } catch (error) {
    Logger.log('❌ Daily summary failed: ' + error);
  }
}

/** Menu "📬 Send Daily Summary Now" — works even when the daily email is turned off. */
function sendDailySummaryNow() {
  const config = requireConfig();
  if (!config) return;
  try {
    const sent = deliverDailySummary();
    safeAlert('📬 Sent to ' + sent.to + '\n' + sent.subject);
  } catch (error) {
    safeAlert('❌ Could not send the summary: ' + error);
  }
}

// ====================================================================================================
// 06_monthlySummary.gs
// ====================================================================================================

/**
 * MONTHLY SUMMARY EMAIL — v1.1.26
 *
 * Sent on the 1st of each month about the PREVIOUS calendar month — the same
 * period as the Dashboard's month columns and as a monthly salary. By the
 * time it runs (the hour chosen in the Setup Wizard), the 6 AM update of the
 * 1st has already swept the whole previous month (see computeDailyWindow()),
 * so the month's last day is included.
 *
 * Same rules as the Dashboard and the daily email: spending = Transaction and
 * Transfer rows with a category other than "Exclude", in DOP-equivalent at
 * the Dashboard's rates. Uses the email kit in 05_dailySummary.gs.
 *
 * Menu "🗓️ Send Monthly Summary Now" sends the last complete month on demand.
 */
const SUMMARY_SUBSCRIPTIONS_CATEGORY = 'Streaming & Subscriptions';

/** Everything the monthly email needs. Pure — see tests/. opts as computeDailySummary() plus optional `month` (a Date in it). */
function computeMonthlySummary(values, opts) {
  const rates = opts.rates;
  const ref = opts.today;
  const month = opts.month
    ? new Date(opts.month.getFullYear(), opts.month.getMonth(), 1, 12)
    : new Date(ref.getFullYear(), ref.getMonth() - 1, 1, 12);
  const monthKeyOf = d => normalizeDateForCompare(d).slice(0, 7);
  const monthKey = monthKeyOf(month);
  const prevMonth = new Date(month.getFullYear(), month.getMonth() - 1, 1, 12);
  const prevKey = monthKeyOf(prevMonth);
  const usualKeys = [1, 2, 3].map(k => monthKeyOf(new Date(month.getFullYear(), month.getMonth() - k, 1, 12)));
  const daysInMonth = new Date(month.getFullYear(), month.getMonth() + 1, 0).getDate();
  const toDop = (amount, cur) => (Number(amount) || 0) * ({ USD: rates.USD, EUR: rates.EUR, COP: rates.COP }[cur] || 1);
  const cardRate = {}, cardName = {};
  (opts.cards || []).forEach(c => { if (c.rate > 0) { cardRate[c.bank] = c.rate; cardName[c.bank] = c.name || ''; } });

  const m = {
    month: month, monthKey: monthKey, prevMonth: prevMonth, daysInMonth: daysInMonth,
    netIncomeDop: Number(opts.netIncomeDop) || 0,
    spent: 0, byCategory: {}, byDay: new Array(daysInMonth + 1).fill(0), byDayToDay: new Array(daysInMonth + 1).fill(0),
    byBank: {}, merchants: {}, subs: {},
    prevSubs: {}, prevByCategory: {}, usualCatSums: {}, monthTotals: {},
    fixed: 0, variable: 0, transfersOut: 0, transfersCount: 0, openCount: 0, openTotal: 0,
    placeholders: 0, unmatchedReversals: 0, purchases: 0, biggest: null,
    missed: { total: 0, byRule: {} }
  };

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[TX_COL.DATE]) continue;
    const key = normalizeDateForCompare(r[TX_COL.DATE]);
    const mk = key.slice(0, 7);
    const type = r[TX_COL.TYPE] || 'Transaction';
    const cat = String(r[TX_COL.CATEGORY] || '');
    const amt = toDop(r[TX_COL.AMOUNT], r[TX_COL.CURRENCY]);
    const merchant = String(r[TX_COL.MERCHANT] || '');
    const bank = String(r[TX_COL.BANK] || '');
    const spend = (type === 'Transaction' || type === 'Transfer' || type === 'Incoming') && cat !== '' && cat !== EXCLUDE_CATEGORY;   // v1.1.51
    if (spend) m.monthTotals[mk] = (m.monthTotals[mk] || 0) + amt;

    if (mk === monthKey) {
      if (spend) {
        m.spent += amt;
        m.byCategory[cat] = (m.byCategory[cat] || 0) + amt;
        m.byDay[Number(key.slice(8, 10))] += amt;
        if (!isBillCategory(cat)) m.byDayToDay[Number(key.slice(8, 10))] += amt;   // calendar colours: no rent/bill spikes
        m.byBank[bank] = (m.byBank[bank] || 0) + amt;
        const g = m.merchants[merchant] || (m.merchants[merchant] = { merchant: merchant, total: 0, count: 0 });
        g.total += amt; g.count++;
        if (FIXED_CATEGORY_NAMES.indexOf(cat) !== -1) m.fixed += amt; else m.variable += amt;
        if (amt > 0 && type === 'Transaction') {
          m.purchases++;
          if (!m.biggest || amt > m.biggest.amount) m.biggest = { merchant: merchant, amount: amt, bank: bank, date: key };
        }
        if (cat === SUMMARY_SUBSCRIPTIONS_CATEGORY) {
          const sg = m.subs[merchant] || (m.subs[merchant] = { merchant: merchant, total: 0, count: 0 });
          sg.total += amt; sg.count++;
        }
        // cashback left on the table: a purchase whose best card (per the terms) is one you have, paid with another
        if (type === 'Transaction' && amt > 0) {
          const rule = CARD_TIP_RULES.find(x => x.test.test(normalizeKeyword(merchant)));
          if (rule && cardRate[rule.bank] && bank !== rule.bank) {
            const k = rule.label + '|' + bank + '|' + rule.bank;
            const e = m.missed.byRule[k] || (m.missed.byRule[k] = { label: rule.label, paidWith: bank,
              best: rule.bank + (cardName[rule.bank] ? ' ' + cardName[rule.bank] : ''), spent: 0, cashback: 0 });
            e.spent += amt; e.cashback += amt * cardRate[rule.bank];
            m.missed.total += amt * cardRate[rule.bank];
          }
        }
      }
      if (type === 'Transfer' && cat !== EXCLUDE_CATEGORY) { m.transfersOut += amt; m.transfersCount++; }
      if (type === 'Transfer' && cat === '') { m.openCount++; m.openTotal += amt; }
      if (merchant === GARBLED_PLACEHOLDER) m.placeholders++;
      if (merchant === REVERSAL_UNMATCHED) m.unmatchedReversals++;
    }
    if (spend && mk === prevKey) {
      m.prevByCategory[cat] = (m.prevByCategory[cat] || 0) + amt;
      if (cat === SUMMARY_SUBSCRIPTIONS_CATEGORY) m.prevSubs[merchant] = true;
    }
    if (spend && usualKeys.indexOf(mk) !== -1) m.usualCatSums[cat] = (m.usualCatSums[cat] || 0) + amt;
  }

  // comparisons
  m.prevTotal = m.monthTotals[prevKey] || 0;
  const usualWithData = usualKeys.filter(k => (m.monthTotals[k] || 0) > 0);
  m.usualTotal = usualWithData.length ? usualWithData.reduce((a, k) => a + m.monthTotals[k], 0) / usualWithData.length : 0;
  m.usualMonths = usualWithData.length;
  m.vsPrev = m.prevTotal > 0 ? m.spent / m.prevTotal - 1 : null;
  m.vsUsual = m.usualTotal > 0 ? m.spent / m.usualTotal - 1 : null;
  m.leftOver = m.netIncomeDop > 0 ? m.netIncomeDop - m.spent : null;
  m.savingsRate = m.netIncomeDop > 0 ? m.leftOver / m.netIncomeDop : null;
  m.incomeUsed = m.netIncomeDop > 0 ? m.spent / m.netIncomeDop : null;

  m.categories = Object.keys(m.byCategory).map(c => ({
    cat: c, amount: m.byCategory[c], share: m.spent > 0 ? m.byCategory[c] / m.spent : 0,
    prev: m.prevByCategory[c] || 0, usual: m.usualMonths ? (m.usualCatSums[c] || 0) / m.usualMonths : 0,
    fixed: FIXED_CATEGORY_NAMES.indexOf(c) !== -1
  })).sort((a, b) => b.amount - a.amount);
  m.topMerchants = Object.keys(m.merchants).map(k => m.merchants[k]).sort((a, b) => b.total - a.total).slice(0, 8);
  m.subscriptions = Object.keys(m.subs).map(k => Object.assign({ recurring: !!m.prevSubs[k] }, m.subs[k]))
    .sort((a, b) => b.total - a.total);
  m.subscriptionsTotal = m.subscriptions.reduce((a, x) => a + x.total, 0);
  m.banks = Object.keys(m.byBank).map(b => ({ bank: b, amount: m.byBank[b] })).sort((a, b) => b.amount - a.amount);
  m.missedList = Object.keys(m.missed.byRule).map(k => m.missed.byRule[k]).sort((a, b) => b.cashback - a.cashback);

  // year to date (same year, up to this month)
  const year = String(month.getFullYear());
  const ytdKeys = Object.keys(m.monthTotals).filter(k => k.slice(0, 4) === year && k <= monthKey && m.monthTotals[k] > 0);
  m.ytdTotal = ytdKeys.reduce((a, k) => a + m.monthTotals[k], 0);
  m.ytdMonths = ytdKeys.length;
  m.ytdAverage = m.ytdMonths ? m.ytdTotal / m.ytdMonths : 0;
  m.ytdSavingsRate = m.netIncomeDop > 0 && m.ytdMonths ? 1 - m.ytdTotal / (m.netIncomeDop * m.ytdMonths) : null;

  m.recommendations = monthlyRecommendations(m);
  m.investments = opts.investments || null;   // v1.1.32
  m.lastRun = opts.lastRun || null;
  m.now = opts.now || ref;
  return m;
}

/** Fixed rules → [{tone, text}]. */
function monthlyRecommendations(m) {
  const out = [];
  const name = SUMMARY_MONTHS_FULL[m.month.getMonth()];
  const prevName = SUMMARY_MONTHS_FULL[m.prevMonth.getMonth()];
  if (m.savingsRate !== null) {
    if (m.savingsRate < 0) out.push({ tone: 'bad', text: 'You spent ' + summaryMoney(-m.leftOver) + ' more than your net income in ' + name + '.' });
    else if (m.savingsRate < 0.1) out.push({ tone: 'warn', text: 'Only ' + summaryPct(m.savingsRate) + ' of your net income was left over (' + summaryMoney(m.leftOver) + ').' });
    else out.push({ tone: 'good', text: 'You kept ' + summaryPct(m.savingsRate) + ' of your net income — ' + summaryMoney(m.leftOver) + '.' });
  }
  if (m.vsPrev !== null && m.vsPrev > 0.15) {
    const driver = m.categories.map(c => ({ cat: c.cat, diff: c.amount - c.prev })).sort((a, b) => b.diff - a.diff)[0];
    out.push({ tone: 'warn', text: 'Spending was up ' + summaryPct(m.vsPrev) + ' vs. ' + prevName +
      (driver && driver.diff > 0 ? ', mostly ' + driver.cat + ' (+' + summaryMoney(driver.diff) + ')' : '') + '.' });
  } else if (m.vsPrev !== null && m.vsPrev < -0.1) {
    out.push({ tone: 'good', text: 'Spending was down ' + summaryPct(m.vsPrev) + ' vs. ' + prevName + '.' });
  }
  m.categories.filter(c => !c.fixed && c.usual > 0 && c.amount > c.usual * 1.15 && c.amount - c.usual > 1000)
    .sort((a, b) => (b.amount - b.usual) - (a.amount - a.usual)).slice(0, 2)
    .forEach(c => out.push({ tone: 'warn', text: c.cat + ': ' + summaryMoney(c.amount) + ', ' +
      summaryPct(c.amount / c.usual - 1) + ' above your usual month (' + summaryMoney(c.usual) + ').' }));
  if (m.subscriptions.length) {
    out.push({ tone: 'info', text: m.subscriptions.length + ' subscription(s) cost ' + summaryMoney(m.subscriptionsTotal) +
      ' this month — about ' + summaryMoney(m.subscriptionsTotal * 12) + ' a year. Worth a quick check that you still use them all.' });
  }
  if (m.missed.total >= 50) {
    const top = m.missedList[0];
    out.push({ tone: 'info', text: 'About ' + summaryMoney(m.missed.total) + ' in cashback was left on the table — e.g. ' +
      top.label.toLowerCase() + ' paid with ' + top.paidWith + ' (' + summaryMoney(top.spent) + ') would earn with your ' + top.best + ' card.' });
  }
  if (m.openCount) {
    out.push({ tone: 'note', text: m.openCount + ' transfer(s) have no category (' + summaryMoney(m.openTotal) +
      ') and aren\'t counted as spending — add a Custom Rule, or "Exclude" for your own accounts.' });
  }
  if (m.placeholders || m.unmatchedReversals) {
    out.push({ tone: 'note', text: 'Data to check: ' + [m.placeholders ? m.placeholders + ' row(s) with an unreadable merchant' : '',
      m.unmatchedReversals ? m.unmatchedReversals + ' reversal(s) without their original purchase' : ''].filter(Boolean).join(' and ') + '.' });
  }
  return out;
}

/** Month calendar (Mon–Sun rows) coloured by the day's spending. */
function ekCalendar(month, byDay) {
  const days = byDay.length - 1;
  const max = Math.max.apply(null, byDay.concat([1]));
  const level = v => v <= 0 ? 0 : Math.min(5, 1 + Math.floor(v / max * 4.999));
  const first = (new Date(month.getFullYear(), month.getMonth(), 1, 12).getDay() + 6) % 7;   // Monday = 0
  let html = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:4px">' +
    '<tr>' + ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].map(d =>
      '<td width="14%" align="center" style="font-size:10px;color:' + EK.faint + '">' + d + '</td>').join('') + '</tr><tr>';
  let col = 0;
  for (let i = 0; i < first; i++, col++) html += '<td></td>';
  for (let d = 1; d <= days; d++, col++) {
    if (col && col % 7 === 0) html += '</tr><tr>';
    const lv = level(byDay[d]);
    html += '<td width="14%" align="center" style="height:34px;border-radius:6px;background:' + EK.heat[lv] + ';font-size:11px;color:' +
      (lv >= 4 ? '#FFFFFF' : EK.muted) + '">' + d + '</td>';
  }
  while (col % 7) { html += '<td></td>'; col++; }
  html += '</tr></table><div style="font-size:11px;color:' + EK.faint + ';padding:4px 4px 0">Less ' +
    EK.heat.map(c => '<span style="display:inline-block;width:12px;height:10px;border-radius:3px;background:' + c +
      ';vertical-align:middle"></span>').join('&nbsp;') + ' More</div>';
  return html;
}

/** Subject, HTML and plain-text bodies of the monthly email. Pure — see tests/. */
function buildMonthlySummaryEmail(m, opts) {
  opts = opts || {};
  const links = opts.links || { dashboard: opts.sheetUrl };
  const name = SUMMARY_MONTHS_FULL[m.month.getMonth()] + ' ' + m.month.getFullYear();
  const prevName = SUMMARY_MONTHS_FULL[m.prevMonth.getMonth()];
  const subject = name + ': ' + summaryMoney(m.spent) + ' spent' +
    (m.savingsRate !== null ? ' · ' + summaryPct(m.savingsRate) + (m.savingsRate < 0 ? ' over income' : ' left over')
      : (m.vsPrev !== null ? ' (' + summarySignedPct(m.vsPrev) + ' vs. ' + prevName + ')' : ''));
  const rows = [], T = [];

  // hero
  const chips = [];
  if (m.vsPrev !== null) chips.push(ekChip(summarySignedPct(m.vsPrev) + ' vs. ' + prevName, m.vsPrev > 0.15 ? 'warn' : m.vsPrev < -0.05 ? 'good' : 'info'));
  if (m.incomeUsed !== null) chips.push(ekChip(summaryPct(m.incomeUsed) + ' of net income', m.incomeUsed > 1 ? 'bad' : m.incomeUsed > 0.9 ? 'warn' : 'good'));
  rows.push('<tr><td style="padding:24px 28px 0"><div style="font-size:13px;color:' + EK.muted + '">Spent in ' +
    SUMMARY_MONTHS_FULL[m.month.getMonth()] + ' · ' + m.purchases + ' purchase(s)</div>' +
    '<div style="font-size:40px;font-weight:800;color:' + EK.navy + ';line-height:1.15;margin:4px 0 10px">' + summaryMoney(m.spent) +
    '</div>' + chips.join('&nbsp; ') + '</td></tr>');
  T.push(name + ': ' + summaryMoney(m.spent) + ' spent · ' + m.purchases + ' purchase(s)');

  // KPIs + income used
  const tiles = [
    { label: 'Left over', value: m.leftOver !== null ? summaryMoneyShort(m.leftOver) : '—',
      sub: m.savingsRate !== null ? summaryPct(m.savingsRate) + ' of net income' : 'set income in the wizard',
      tone: m.savingsRate === null ? null : m.savingsRate < 0 ? 'bad' : m.savingsRate < 0.1 ? 'warn' : 'good' },
    { label: 'vs. ' + prevName, value: m.vsPrev !== null ? summarySignedPct(m.vsPrev) : '—',
      sub: m.prevTotal ? summaryMoneyShort(m.prevTotal) + ' then' : 'no data', tone: m.vsPrev !== null && m.vsPrev > 0.15 ? 'warn' : null },
    { label: 'Usual month', value: m.usualTotal ? summaryMoneyShort(m.usualTotal) : '—',
      sub: m.usualMonths ? 'avg. of ' + m.usualMonths + ' month(s)' : 'no history yet' }
  ];
  let inner = ekKpis(tiles);
  if (m.incomeUsed !== null) {
    inner += '<div style="height:14px;font-size:0">&nbsp;</div>' + ekBar('Net income used', m.incomeUsed,
      m.incomeUsed > 1 ? ekTone('bad').fg : m.incomeUsed > 0.9 ? ekTone('warn').fg : ekTone('good').fg, summaryPct(m.incomeUsed),
      summaryEscape(summaryMoney(m.spent) + ' of ' + summaryMoney(m.netIncomeDop)));
  }
  const fixedShare = m.spent > 0 ? m.fixed / m.spent : 0;
  inner += '<div style="font-size:13px;color:' + EK.text + ';margin:4px 0 6px">Fixed vs. variable</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-radius:999px;overflow:hidden"><tr>' +
    (m.fixed > 0 ? '<td width="' + Math.max(1, Math.round(fixedShare * 100)) + '%" style="background:' + EK.navy + ';height:10px;font-size:0">&nbsp;</td>' : '') +
    (m.variable > 0 ? '<td style="background:#7FA6E8;height:10px;font-size:0">&nbsp;</td>' : '') + '</tr></table>' +
    '<div style="font-size:12px;color:' + EK.muted + ';padding-top:5px"><span style="color:' + EK.navy + '">■</span> Fixed ' +
    summaryEscape(summaryMoney(m.fixed)) + ' (' + summaryPct(fixedShare) + ') &nbsp; <span style="color:#7FA6E8">■</span> Variable ' +
    summaryEscape(summaryMoney(m.variable)) + '</div>';
  rows.push(ekSection('The month in numbers', inner));
  if (m.leftOver !== null) T.push('Left over: ' + summaryMoney(m.leftOver) + ' (' + summaryPct(m.savingsRate) + ' of net income)');
  if (m.vsPrev !== null) T.push('vs. ' + prevName + ': ' + summarySignedPct(m.vsPrev));

  rows.push(ekSection('Day by day · day-to-day spending', ekCalendar(m.month, m.byDayToDay)));

  // categories
  const maxCat = Math.max.apply(null, m.categories.map(c => c.amount).concat([1]));
  rows.push(ekSection('Categories', m.categories.map(c => {
    const parts = [summaryPct(c.share) + ' of the month'];
    if (c.prev > 0) {
      const d = c.amount / c.prev - 1;
      const col = d > 0.15 ? ekTone('warn').fg : d < -0.1 ? ekTone('good').fg : EK.muted;
      parts.push('<span style="color:' + col + '">' + summarySignedPct(d) + ' vs. ' + summaryEscape(prevName) + '</span>');
    } else if (!c.fixed) {
      parts.push('new vs. ' + summaryEscape(prevName));
    }
    return ekBar(c.cat, c.amount / maxCat, c.fixed ? EK.navy : EK.accent, summaryMoney(c.amount), parts.join(' · '));
  }).join('')));
  m.categories.forEach(c => T.push('  • ' + c.cat + ': ' + summaryMoney(c.amount)));

  if (m.topMerchants.length) {
    rows.push(ekSection('Top merchants', ekList(m.topMerchants.map(g => ({
      title: g.merchant, meta: g.count + (g.count === 1 ? ' payment' : ' payments'), right: summaryMoney(g.total) })))));
  }
  if (m.subscriptions.length) {
    rows.push(ekSection('Subscriptions', ekList(m.subscriptions.map(x => ({
      title: x.merchant, meta: x.recurring ? 'also last month' : 'new this month', right: summaryMoney(x.total) }))) +
      '<div style="font-size:12px;color:' + EK.muted + ';padding-top:6px">≈ ' + summaryEscape(summaryMoney(m.subscriptionsTotal * 12)) +
      ' a year at this rate.</div>'));
    T.push('', 'Subscriptions: ' + summaryMoney(m.subscriptionsTotal) + ' (' + m.subscriptions.map(x => x.merchant).join(', ') + ')');
  }

  // banks + cashback left on the table
  if (m.banks.length) {
    const maxBank = Math.max.apply(null, m.banks.map(b => b.amount).concat([1]));
    let bankHtml = m.banks.map(b => ekBar(b.bank, b.amount / maxBank, '#7FA6E8', summaryMoney(b.amount))).join('');
    if (m.missedList.length) {
      bankHtml += '<div style="height:6px;font-size:0">&nbsp;</div>' + ekNotes(m.missedList.slice(0, 3).map(e => ({ tone: 'info',
        text: e.label + ' paid with ' + e.paidWith + ' (' + summaryMoney(e.spent) + '): about ' + summaryMoney(e.cashback) +
          ' back with your ' + e.best + ' card.' })));
    }
    rows.push(ekSection('By bank', bankHtml));
  }

  // transfers + year to date
  rows.push(ekSection('Transfers', '<div style="font-size:14px;color:' + EK.text + '">' + m.transfersCount + ' transfer(s) out · ' +
    summaryEscape(summaryMoney(m.transfersOut)) + (m.openCount ? ' · <span style="color:' + ekTone('warn').fg + '">' + m.openCount +
      ' without a category (' + summaryEscape(summaryMoney(m.openTotal)) + ')</span>' : '') + '</div>'));
  rows.push(ekSection(m.month.getFullYear() + ' so far', ekKpis([
    { label: 'Spent', value: summaryMoneyShort(m.ytdTotal), sub: m.ytdMonths + ' month(s)' },
    { label: 'Per month', value: summaryMoneyShort(m.ytdAverage), sub: 'average' },
    { label: 'Kept', value: m.ytdSavingsRate !== null ? summaryPct(m.ytdSavingsRate) : '—', sub: 'of net income',
      tone: m.ytdSavingsRate === null ? null : m.ytdSavingsRate < 0 ? 'bad' : m.ytdSavingsRate < 0.1 ? 'warn' : 'good' }
  ])));
  T.push('', m.month.getFullYear() + ' so far: ' + summaryMoney(m.ytdTotal) + ' in ' + m.ytdMonths + ' month(s)');

  // investments (v1.1.32)
  if (m.investments && m.investments.none) {   // v1.1.49: the history starts after this month
    rows.push(ekSection('Investments', '<div style="font-size:13px;color:' + EK.muted + '">Your investment history starts on ' +
      summaryEscape(m.investments.firstDay) + ' — this section fills in from that month\'s summary, sent on ' +
      summaryEscape(m.investments.firstReport) + '.</div>'));
    T.push('', 'Investments: history starts on ' + m.investments.firstDay + ' — first monthly figures on ' + m.investments.firstReport);
  } else if (m.investments) {
    const iv = m.investments;
    let inner = ekKpis([
      { label: 'Value at month end', value: summaryUsd(iv.endValue), sub: iv.partial ? 'tracking began this month'
        : iv.added && iv.added.length ? 'added this month: ' + iv.added.join(', ') : 'from ' + summaryUsd(iv.startValue) },
      { label: 'Gain in ' + SUMMARY_MONTHS[m.month.getMonth()], value: summarySignedUsd(iv.gain),
        sub: (iv.gainPct !== null ? summarySignedPctFine(iv.gainPct) + ' · ' : '') + 'deposits left out', tone: iv.gain > 0 ? 'good' : iv.gain < 0 ? 'bad' : null },
      { label: 'Deposited', value: summaryUsd(iv.deposits), sub: 'dividends ' + summaryUsd(iv.dividends) + ' · fees ' + summaryUsd(iv.fees) }
    ]);
    if (iv.allocation.length) {
      inner += '<div style="height:12px;font-size:0">&nbsp;</div>' + iv.allocation.map(a =>
        ekBar(a.account, a.share, '#0F766E', summaryUsd(a.value), summaryEscape(summaryPct(a.share) + ' of the portfolio'))).join('');
    }
    if (iv.returns && iv.returns.periodReturn !== null) {
      inner += '<div style="font-size:12px;color:' + EK.muted + '">Since tracking began (' + summaryEscape(iv.returns.start) + '): ' +
        summaryEscape(summarySignedPctFine(iv.returns.periodReturn) + ' · ' + summarySignedUsd(iv.returns.gain)) +
        (iv.returns.annualized !== null ? ' · ' + summaryEscape(summarySignedPctFine(iv.returns.annualized)) + ' a year' : '') + '</div>';
    }
    rows.push(ekSection('Investments', inner));
    T.push('', 'Investments: ' + summaryUsd(iv.endValue) + ' at month end · gain ' + summarySignedUsd(iv.gain));
  }

  const recs = m.recommendations.length ? m.recommendations : [{ tone: 'good', text: 'Nothing to flag.' }];
  rows.push(ekSection('Recommendations', ekNotes(recs)));
  T.push('', 'Recommendations:', ...recs.map(r => '  • ' + r.text));

  const health = summaryDataHealth(m.lastRun, m.now || new Date());
  rows.push('<tr><td style="padding:20px 28px 0"><div style="font-size:12px;color:' + ekTone(health.tone).fg + ';background:' +
    ekTone(health.tone).bg + ';border-radius:8px;padding:9px 12px">' + summaryEscape(health.text) + '</div></td></tr>');

  const sheetName = opts.sheetName || 'the spreadsheet';
  const buttons = ekButtons([
    { text: 'Open ' + sheetName + ' →', url: links.dashboard, primary: true },
    { text: 'Transactions', url: links.transactions },
    { text: 'Custom Rules', url: m.openCount ? links.rules : null }
  ]);
  if (links.dashboard) T.push('', 'Open ' + sheetName + ': ' + links.dashboard);
  const preheader = summaryMoney(m.spent) + ' spent' + (m.leftOver !== null ? ' · ' + summaryMoney(m.leftOver) + ' left over' : '') +
    (m.vsPrev !== null ? ' · ' + summarySignedPct(m.vsPrev) + ' vs. ' + prevName : '');
  const html = ekShell({
    eyebrow: 'Monthly summary', title: name,
    subtitle: '1–' + m.daysInMonth + ' ' + SUMMARY_MONTHS[m.month.getMonth()] + ' · amounts in DOP-equivalent',
    preheader: preheader, body: rows.join(''), buttons: buttons,
    footer: 'Financial Tracker · amounts converted at the Dashboard\'s exchange rates. ' +
      'Turn this email off in 📊 Tracker › Setup Wizard.'
  });
  return { subject: stripAstral(subject), html: toAsciiHtml(html), text: stripAstral(T.join('\n')), preheader: preheader };
}

/** Builds and sends the monthly email for the month before `now`. Returns { to, subject }. */
function deliverMonthlySummary(now) {
  refreshAutoDeductions();
  const config = getConfig();
  if (!config) throw new Error('Setup not completed.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  const values = sheet ? sheet.getDataRange().getValues() : [[]];
  const rates = readDashboardRates(ss);
  const ref = now || new Date();
  let investments = null;   // v1.1.32
  try {
    investments = investmentsReportData('monthly', { rates: rates, month: new Date(ref.getFullYear(), ref.getMonth() - 1, 1, 12) });
  } catch (error) { Logger.log('Investments brief: ' + error); }
  const summary = computeMonthlySummary(values, {
    today: ref, rates: rates, netIncomeDop: computeNetIncomeDop(config, rates),
    cards: readDashboardCards(ss), lastRun: readLastRun(), now: ref, investments: investments
  });
  const mail = buildMonthlySummaryEmail(summary, { sheetName: ss.getName(), links: summaryLinks(ss) });
  const to = config.notifyEmail || config.email;
  GmailApp.sendEmail(to, mail.subject, mail.text, { htmlBody: mail.html, name: 'Financial Tracker' });
  Logger.log('🗓️ Monthly summary sent to ' + to + ' — ' + mail.subject);
  return { to: to, subject: mail.subject };
}

/** Trigger handler (1st of each month, at the hour chosen in the Setup Wizard). */
function sendMonthlySummary() {
  try {
    const config = getConfig();
    if (!config || !config.notifyMonthly) {
      Logger.log('Monthly summary is turned off — nothing sent.');
      return;
    }
    deliverMonthlySummary();
  } catch (error) {
    Logger.log('❌ Monthly summary failed: ' + error);
  }
}

/** Menu "🗓️ Send Monthly Summary Now" — last complete month, even when the monthly email is off. */
function sendMonthlySummaryNow() {
  const config = requireConfig();
  if (!config) return;
  try {
    const sent = deliverMonthlySummary();
    safeAlert('🗓️ Sent to ' + sent.to + '\n' + sent.subject);
  } catch (error) {
    safeAlert('❌ Could not send the summary: ' + error);
  }
}

// ====================================================================================================
// 07_investments.gs
// ====================================================================================================

/**
 * INVESTMENTS — v1.1.29
 *
 * The same idea as the bank side: every movement becomes a row of a ledger,
 * whatever its source, and values are computed from it.
 *
 *   Investment Ledger  every movement: Buy / Sell / Dividend / Deposit /
 *                      Withdrawal / Fee, plus Snapshot (positions as shown by
 *                      the broker on a date) and Valuation (the value of an
 *                      account tracked by balance, e.g. a fund or pension)
 *   Holdings           positions computed from the ledger and valued with
 *                      GOOGLEFINANCE, plus the other accounts and a total
 *   Investment Accounts  accounts, and the keyword that identifies the bank
 *                      transfers that fund each one
 *
 * Sources:
 *   · HAPI emails (no-reply@hapi.trade): executed orders (ticker, quantity,
 *     average price, cost and the per-order fee) and dividends (net amount).
 *     "Deposit Completed" has no amount — deposits come from the bank side.
 *   · Bank transfers already in Transactions whose beneficiary contains an
 *     account's deposit keyword → Deposit rows (the transfer itself stays in
 *     whatever category you give it, e.g. "Exclude", so it isn't spending).
 *   · Rows you type: Snapshot, Valuation, anything a broker doesn't email.
 *
 * Positions start from each account's most recent Snapshot and apply the
 * movements dated AFTER that day — so re-reading old emails never counts a
 * trade twice. Earlier movements stay in the ledger as history (dividends,
 * fees and contributions of the year still count them).
 */
const INVESTMENT_LEDGER_SHEET = 'Investment Ledger';
const HOLDINGS_SHEET = 'Holdings';
const INVESTMENT_ACCOUNTS_SHEET = 'Investment Accounts';
const LEDGER_HEADERS = ['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id'];
const LG = { DATE: 0, ACCOUNT: 1, TYPE: 2, TICKER: 3, QTY: 4, PRICE: 5, AMOUNT: 6, FEE: 7, CURRENCY: 8, SOURCE: 9, NOTES: 10, ID: 11 };
// v1.1.42: Notice — a broker said a deposit happened but not how much (HAPI's "Deposit Completed"); it moves nothing
const LEDGER_TYPES = ['Buy', 'Sell', 'Dividend', 'Deposit', 'Withdrawal', 'Fee', 'Snapshot', 'Valuation', 'Notice'];
const ACCOUNTS_HEADERS = ['Account', 'Kind', 'Deposit keyword', 'Notes'];
const DEFAULT_INVESTMENT_ACCOUNTS = [
  ['HAPI', 'Broker', 'OUROSR', "Orders and dividends come from HAPI's emails; deposits from bank transfers whose " +
    "beneficiary contains the keyword (HAPI's collection account in the DR)."]
];
const BROKER_PATTERNS = {
  HAPI: { searchQuery: 'from:no-reply@hapi.trade', sender: 'hapi.trade', parse: parseHapiMessage }
};
// GOOGLEFINANCE symbols for tickers it doesn't know as they are. v1.1.30: any crypto pair — 6+ letters ending in USD
// (ETHUSD, SHIBUSD…) — is CURRENCY:<pair>; US stock tickers have at most 5 letters, so none is caught by mistake.
const PRICE_SYMBOLS = {};
function priceSymbol(ticker) {
  if (PRICE_SYMBOLS[ticker]) return PRICE_SYMBOLS[ticker];
  return /^[A-Z]{3,}USD$/.test(ticker) && ticker.length >= 6 ? 'CURRENCY:' + ticker : ticker;
}
/**
 * v1.1.31: GOOGLEFINANCE doesn't price crypto pairs — ETHUSD and SHIBUSD returned errors in a live sheet (a known
 * limitation). Crypto comes from Coinbase's public spot price (no key), fetched by the script on every refresh;
 * stocks and ETFs stay on GOOGLEFINANCE. Crypto prices therefore update when the tracker runs, not continuously.
 */
const CRYPTO_PRICE_URL = pair => 'https://api.coinbase.com/v2/prices/' + pair + '/spot';
function isCryptoPair(ticker) { return /^[A-Z]{3,}USD$/.test(ticker) && ticker.length >= 6; }
// v1.1.31: the sanity check only trusts a RECENT last known price — after months without trades a 50%+ move is real
// (the first version would have frozen such a position at its old price, labelled "check symbol").
const SANITY_MAX_AGE_DAYS = 90;
function daysBetween(fromKey, toKey) {
  const d = k => { const p = String(k).split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]); };
  return Math.round((d(toKey) - d(fromKey)) / 86400000);
}

/** Coinbase spot prices for the crypto pairs among `tickers` → { TICKER: price }; missing on any failure. */
function fetchCryptoPrices(tickers) {
  const pairs = tickers.filter(isCryptoPair).filter((t, i, a) => a.indexOf(t) === i);
  const out = {};
  if (!pairs.length) return out;
  let responses;
  try {
    responses = UrlFetchApp.fetchAll(pairs.map(t => ({ url: CRYPTO_PRICE_URL(t.slice(0, -3) + '-USD'), muteHttpExceptions: true })));
  } catch (error) {
    Logger.log('Crypto prices unavailable: ' + error);
    return out;
  }
  responses.forEach((response, i) => {
    try {
      if (response.getResponseCode() !== 200) return;
      const amount = Number(JSON.parse(response.getContentText()).data.amount);
      if (amount > 0) out[pairs[i]] = amount;
    } catch (error) {
      Logger.log('Crypto price for ' + pairs[i] + ' unreadable: ' + error);
    }
  });
  return out;
}

// prices under a cent (e.g. SHIB) keep 8 decimals instead of showing US$0.00
const USD_PRICE_FORMAT = '[<0.01]"US$"0.00000000;"US$"#,##0.00';
// A live price this far from the last known one is treated as a wrong symbol (e.g. a ticker Google maps to another security)
const PRICE_SANITY = 0.5;

const invNumber = value => {
  if (typeof value === 'number') return value;
  const n = Number(String(value || '').replace(/[,$\s]/g, ''));
  return isNaN(n) ? 0 : n;
};

/** Visible text of an email, whitespace collapsed (plain part, or the HTML stripped). */
function investmentMessageText(message) {
  let text = '';
  try { text = message.getPlainBody() || ''; } catch (error) { text = ''; }
  if (!/[A-Za-z]{3}/.test(text) && typeof message.getBody === 'function') {
    text = String(message.getBody() || '').replace(/<(style|script|head)[\s\S]*?<\/\1>/gi, ' ')
      .replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&');
  }
  return text.replace(/[\u200b\u200c\u034f\u00a0]/g, ' ').replace(/\s+/g, ' ').trim();
}

/**
 * One HAPI email → { kind: 'event', event } | { kind: 'skipped', reason } |
 * { kind: 'unparsed', reason }. An order is only accepted when its own numbers
 * agree (quantity × average price = cost), so a changed template can't write
 * a wrong position silently. Pure — see tests/.
 */
function parseHapiMessage(message) {
  const subject = String(message.getSubject() || '');
  const text = investmentMessageText(message);
  const base = { account: 'HAPI', currency: 'USD', source: 'email', id: 'gmail:' + message.getId(), fee: 0, notes: '' };

  if (/Order Executed/i.test(subject) || /Your order has been executed/i.test(text)) {
    const side = (text.match(/Buy\/Sell:\s*(Buy|Sell)\b/i) || [])[1];
    const ticker = (text.match(/Ticker:\s*([A-Z0-9.\-]{1,12})\b/) || [])[1];
    // the quantity is followed by the order fee, in a row whose label HAPI's template leaves empty
    const qty = text.match(/Quantity(?: Shares)?:\s*([\d,]*\.?\d+)(?:\s*US\$\s*([\d,]*\.?\d+))?/i);
    const price = (text.match(/Average price:\s*US\$\s*([\d,]*\.?\d+)/i) || [])[1];
    const cost = (text.match(/Cost:\s*US\$\s*([\d,]*\.?\d+)/i) || [])[1];
    const status = ((text.match(/Status:\s*(.*?)(?:\s+(?:Your funds|Any questions)\b|$)/i) || [])[1] || '').trim();
    if (!side || !ticker || !qty || !price || !cost) {
      return { kind: 'unparsed', reason: 'order email without side, ticker, quantity, price or cost' };
    }
    if (!/completed/i.test(status)) return { kind: 'unparsed', reason: 'order status "' + status + '"' };
    const q = invNumber(qty[1]), p = invNumber(price), cost$ = invNumber(cost), fee = qty[2] ? invNumber(qty[2]) : 0;
    const tolerance = Math.max(0.05, cost$ * 0.005), gross = q * p;
    // v1.1.38: a market order's Cost is quantity × price (the fee apart); a LIMIT order's Cost includes the fee
    // (e.g. 2 × 150.25 + 2.99 = 303.49 — the rule seen in a live email). Either way the cost basis is quantity × price, as HAPI shows it.
    const feeApart = Math.abs(gross - cost$) <= tolerance;
    const feeIncluded = fee > 0 && (Math.abs(gross + fee - cost$) <= tolerance || Math.abs(gross - fee - cost$) <= tolerance);
    if (!(q > 0) || !(p > 0) || !(feeApart || feeIncluded)) {
      return { kind: 'unparsed', reason: 'quantity × average price does not match the cost' };
    }
    const orderType = ((text.match(/Order type:\s*(.+?)\s+Buy\/Sell:/i) || [])[1] || '').trim();
    return { kind: 'event', event: Object.assign(base, {
      date: message.getDate(), type: side.charAt(0).toUpperCase() + side.slice(1).toLowerCase(), ticker: ticker,
      qty: q, price: p, amount: feeApart ? cost$ : +gross.toFixed(2), fee: fee,
      notes: /limit/i.test(orderType) ? 'Limit order' : ''
    }) };
  }
  if (/Dividend/i.test(subject)) {
    const ticker = (subject.match(/Dividend from ([A-Z0-9.\-]{1,12})\b/i) ||
      text.match(/dividends? from (?:your shares in )?([A-Z0-9.\-]{1,12})\b/) || [])[1];
    const amount = (text.match(/Net amount received:\s*(?:US)?\$\s*([\d,]*\.?\d+)/i) || [])[1];
    if (!ticker || !amount) return { kind: 'unparsed', reason: 'dividend email without ticker or amount' };
    const paid = text.match(/Payment date:\s*(\d{4})-(\d{2})-(\d{2})/);
    const date = paid ? new Date(Number(paid[1]), Number(paid[2]) - 1, Number(paid[3]), 12) : message.getDate();
    return { kind: 'event', event: Object.assign(base, {
      date: date, type: 'Dividend', ticker: ticker.toUpperCase(), qty: '', price: '', amount: invNumber(amount),
      notes: 'net amount received'
    }) };
  }
  if (/Deposit Completed/i.test(subject)) {
    // v1.1.42: no amount, but the date says a deposit happened — kept as a Notice, checked against recorded deposits
    return { kind: 'event', event: Object.assign(base, { date: message.getDate(), type: 'Notice', ticker: '', qty: '', price: '',
      amount: '', notes: 'Deposit completed — HAPI sends no amount' }) };
  }
  return { kind: 'skipped', reason: 'not an order or a dividend' };
}

/** Broker emails in the run's range → { events, parsed, skipped, unparsed, threads, failedThreadIds }. */
function captureBrokerEmails(range, skipIds) {
  skipIds = skipIds || new Set();   // v1.1.36: emails already in the ledger aren't read again
  const out = { events: [], parsed: 0, skipped: 0, unparsed: 0, threads: [], failedThreadIds: new Set(), unrecognized: [], readIds: [] };
  Object.keys(BROKER_PATTERNS).forEach(broker => {
    const p = BROKER_PATTERNS[broker];
    const query = p.searchQuery + ' after:' + toEpochSeconds(range.start) + ' before:' + toEpochSeconds(range.endExclusive);
    gmailSearchAll(query, MAX_THREADS_PER_RUN).threads.forEach(thread => {
      thread.getMessages().forEach(message => {
        if (String(message.getFrom()).toLowerCase().indexOf(p.sender) === -1) return;
        const when = message.getDate();
        if (when < range.start || when >= range.endExclusive) return;
        if (skipIds.has('gmail:' + message.getId())) return;
        let result;
        try { result = p.parse(message); } catch (error) { result = { kind: 'unparsed', reason: String(error) }; }
        if (result.kind === 'event') {
          out.events.push(result.event); out.parsed++;
          if (result.event.type !== 'Notice') out.readIds.push(message.getId());   // a Notice's Unrecognized row is managed by checkDepositNotices
        }
        else if (result.kind === 'skipped') { out.skipped++; out.readIds.push(message.getId()); }
        else {
          out.unparsed++;
          out.failedThreadIds.add(thread.getId());
          out.unrecognized.push({ id: message.getId(), date: when, bank: broker, subject: message.getSubject() || '',   // v1.1.35
            reason: 'Broker email not read: ' + result.reason, snippet: investmentMessageText(message).substring(0, 700) });
          Logger.log('⚠️ ' + broker + ' email not read (' + result.reason + '): "' + message.getSubject() + '" ' + when);
        }
      });
      out.threads.push(thread);
    });
  });
  return out;
}

/** Investment Accounts sheet (created with the defaults) → [{account, kind, keyword}]. */
function readInvestmentAccounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INVESTMENT_ACCOUNTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(INVESTMENT_ACCOUNTS_SHEET);
    sheet.getRange(1, 1, 1, ACCOUNTS_HEADERS.length).setValues([ACCOUNTS_HEADERS]);
    sheet.getRange(2, 1, DEFAULT_INVESTMENT_ACCOUNTS.length, ACCOUNTS_HEADERS.length).setValues(DEFAULT_INVESTMENT_ACCOUNTS);
  }
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 3).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ account: String(r[0]).trim(), kind: String(r[1]).trim(), keyword: String(r[2]).trim() }));
}

/** Bank transfers (Transactions) whose beneficiary contains an account's deposit keyword → Deposit events. Pure. */
function brokerDepositsFromTransactions(values, accounts) {
  const withKeyword = accounts.filter(a => a.keyword);
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if ((r[TX_COL.TYPE] || '') !== 'Transfer' || !r[TX_COL.DATE]) continue;
    const merchant = String(r[TX_COL.MERCHANT] || '').toUpperCase();
    const account = withKeyword.find(a => merchant.indexOf(a.keyword.toUpperCase()) !== -1);
    if (!account) continue;
    const amount = Math.abs(invNumber(r[TX_COL.AMOUNT]));
    if (!amount) continue;
    const ref = r[TX_COL.MESSAGE_ID] || r[TX_COL.TX_REF] || (normalizeDateForCompare(r[TX_COL.DATE]) + ':' + amount);
    out.push({ date: r[TX_COL.DATE], account: account.account, type: 'Deposit', ticker: '', qty: '', price: '',
      amount: amount, fee: '', currency: r[TX_COL.CURRENCY] || 'DOP', source: 'bank',
      notes: (r[TX_COL.BANK] || '') + ' → ' + r[TX_COL.MERCHANT], id: 'bank:' + ref });
  }
  return out;
}

function ledgerRow(e) {
  return [e.date, e.account, e.type, e.ticker || '', e.qty === '' || e.qty === undefined ? '' : e.qty,
    e.price === '' || e.price === undefined ? '' : e.price, e.amount, e.fee || '', e.currency || 'USD',
    e.source || 'manual', e.notes || '', e.id || ''];
}

function getOrCreateLedgerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INVESTMENT_LEDGER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(INVESTMENT_LEDGER_SHEET);
    sheet.getRange(1, 1, 1, LEDGER_HEADERS.length).setValues([LEDGER_HEADERS]);
    sheet.hideColumns(LG.ID + 1);          // internal: email / bank reference, used to skip duplicates
    ensureAutoFilter(sheet, LEDGER_HEADERS.length, 1000);
  }
  return sheet;
}

/** Appends the events not saved yet (by Id), newest first. Rows you typed (no Id) are never touched. */
function saveInvestmentEvents(events) {
  const sheet = getOrCreateLedgerSheet();
  const last = sheet.getLastRow();
  const seen = new Set(last > 1 ? sheet.getRange(2, LG.ID + 1, last - 1, 1).getValues().map(r => String(r[0])).filter(Boolean) : []);
  const fresh = [];
  let duplicates = 0;
  events.forEach(e => {
    if (e.id && seen.has(e.id)) { duplicates++; return; }
    if (e.id) seen.add(e.id);
    fresh.push(ledgerRow(e));
  });
  if (fresh.length) {
    ensureRowCapacity(sheet, last + fresh.length);
    sheet.getRange(last + 1, 1, fresh.length, LEDGER_HEADERS.length).setValues(fresh);
    sortSheetByDateDesc(sheet, LG.DATE + 1);
  }
  return { saved: fresh.length, duplicates: duplicates };
}

/**
 * Ledger → positions, cash, other accounts and the year's figures. Pure — see tests/.
 * Average-cost method: a sale removes cost at the position's average cost and the
 * difference is realized P/L. DOP deposits after a snapshot are converted at opts.usdRate
 * to estimate the broker's cash (the broker converts at its own rate).
 */
function computeHoldings(values, opts) {
  opts = opts || {};
  const year = String(opts.year || new Date().getFullYear());
  const usdRate = Number(opts.usdRate) || 0;
  // v1.1.43: a row without a date can't be placed in time, so it isn't counted — and that is now said, not silent
  const undated = values.slice(1).filter(r => !r[LG.DATE] && String(r[LG.ACCOUNT]).trim() && r[LG.TYPE]).length;
  const rows = values.slice(1).filter(r => r[LG.DATE] && String(r[LG.ACCOUNT]).trim() && LEDGER_TYPES.indexOf(r[LG.TYPE]) !== -1)
    .map(r => ({ r: r, key: normalizeDateForCompare(r[LG.DATE]) }))
    .sort((a, b) => a.key < b.key ? -1 : a.key > b.key ? 1 : 0);
  const snapshotDay = {};
  rows.filter(x => x.r[LG.TYPE] === 'Snapshot').forEach(x => {
    const a = String(x.r[LG.ACCOUNT]).trim();
    if (!snapshotDay[a] || x.key > snapshotDay[a]) snapshotDay[a] = x.key;
  });
  const positions = {}, cash = {}, cashEstimated = {}, valuations = {};
  // v1.1.39: a Valuation of an account that has positions (a broker's value on a past date) only marks where its return
  // starts — it isn't an "other account", and its value isn't added twice
  const positionAccounts = new Set(rows.filter(x => ['Snapshot', 'Buy', 'Sell'].indexOf(x.r[LG.TYPE]) !== -1)
    .map(x => String(x.r[LG.ACCOUNT]).trim()));
  const totals = { dividendsYtd: 0, feesYtd: 0, realizedYtd: 0, contributionsYtd: { USD: 0, DOP: 0 } };
  const warnings = [];
  if (undated) warnings.push(undated + ' Investment Ledger row(s) have no date and are not counted — add their date');
  const pos = (a, t) => positions[a + '|' + t] ||
    (positions[a + '|' + t] = { account: a, ticker: t, qty: 0, cost: 0, realized: 0, dividends: 0, lastPrice: 0, lastPriceDay: '' });

  rows.forEach(({ r, key }) => {
    const a = String(r[LG.ACCOUNT]).trim(), type = r[LG.TYPE], t = String(r[LG.TICKER] || '').toUpperCase().trim();
    const q = invNumber(r[LG.QTY]), p = invNumber(r[LG.PRICE]), amount = invNumber(r[LG.AMOUNT]), fee = invNumber(r[LG.FEE]);
    const cur = String(r[LG.CURRENCY] || 'USD').toUpperCase();
    const inYear = key.slice(0, 4) === year;
    if (type === 'Valuation') {
      if (positionAccounts.has(a)) return;
      const value = amount || q * p;
      if (!valuations[a] || key >= valuations[a].day) valuations[a] = { account: a, value: value, currency: cur, day: key, units: q, unitPrice: p };
      return;
    }
    // the year's figures count every movement, before or after a snapshot
    if (inYear) {
      if (type === 'Dividend') totals.dividendsYtd += amount;
      if (type === 'Fee') totals.feesYtd += amount;
      if (fee && (type === 'Buy' || type === 'Sell')) totals.feesYtd += fee;
      if (type === 'Deposit') totals.contributionsYtd[cur === 'DOP' ? 'DOP' : 'USD'] += amount;
    }
    if (t && t !== 'CASH' && p > 0 && (type === 'Buy' || type === 'Sell' || type === 'Snapshot')) {
      const x = pos(a, t);
      if (key >= x.lastPriceDay) { x.lastPrice = p; x.lastPriceDay = key; }
    }
    if (type === 'Dividend' && t) pos(a, t).dividends += amount;

    const snap = snapshotDay[a];
    if (type === 'Snapshot') {
      if (key !== snap) return;                  // only the account's latest snapshot is the base
      if (t === 'CASH') cash[a] = (cash[a] || 0) + amount;
      else { const x = pos(a, t); x.qty += q; x.cost += amount; }
      return;
    }
    if (snap && key <= snap) return;             // already inside the snapshot
    const toUsd = value => {
      if (cur !== 'DOP') return value;
      cashEstimated[a] = true;
      return usdRate > 0 ? value / usdRate : 0;
    };
    if (type === 'Buy') {
      const x = pos(a, t); x.qty += q; x.cost += amount;
      cash[a] = (cash[a] || 0) - amount - fee;
    } else if (type === 'Sell') {
      const x = pos(a, t);
      if (x.qty <= 1e-9) {
        warnings.push(a + ' ' + t + ': sale on ' + key + ' with no position — add a Snapshot of this account');
        return;
      }
      const sold = Math.min(q, x.qty), average = x.cost / x.qty;
      if (q > x.qty + 1e-6) warnings.push(a + ' ' + t + ': sold ' + q + ' but held ' + +x.qty.toFixed(6) + ' on ' + key);
      x.qty -= sold; x.cost -= average * sold;
      const gain = amount * (sold / q) - average * sold;
      x.realized += gain;
      if (inYear) totals.realizedYtd += gain;
      cash[a] = (cash[a] || 0) + amount - fee;
    } else if (type === 'Dividend') {
      cash[a] = (cash[a] || 0) + amount;
    } else if (type === 'Deposit') {
      cash[a] = (cash[a] || 0) + toUsd(amount);
    } else if (type === 'Withdrawal') {
      cash[a] = (cash[a] || 0) - toUsd(amount);
    } else if (type === 'Fee') {
      cash[a] = (cash[a] || 0) - amount;
    }
  });

  const open = Object.keys(positions).map(k => positions[k]).filter(x => x.qty > 1e-9)
    .map(x => Object.assign(x, { qty: +x.qty.toFixed(8), cost: +x.cost.toFixed(2), average: x.cost / x.qty }))
    .sort((a, b) => b.cost - a.cost);
  Object.keys(cashEstimated).forEach(a => {
    if (usdRate <= 0) warnings.push(a + ': cash leaves out DOP deposits — set the USD rate on the Dashboard');
  });
  return {
    positions: open,
    closed: Object.keys(positions).map(k => positions[k]).filter(x => x.qty <= 1e-9 && (x.realized || x.dividends)),
    cash: Object.keys(cash).map(a => ({ account: a, amount: +cash[a].toFixed(2), estimated: !!cashEstimated[a] })),
    valuations: Object.keys(valuations).map(a => valuations[a]),
    totals: totals, warnings: warnings, snapshotDay: snapshotDay
  };
}

/**
 * A number as formula text: 10 significant digits, never scientific notation. v1.1.30: the last
 * known price used to be rounded to 6 decimals, which turned a SHIB-sized 0.0000061234 into 0.000006 (−2%).
 */
function formulaNumber(x) {
  const n = Number(Number(x).toPrecision(10));
  return /e/i.test(String(n)) ? n.toFixed(20).replace(/0+$/, '').replace(/\.$/, '') : String(n);
}

/**
 * Price and Price source cells of a position. Stocks/ETFs: GOOGLEFINANCE formulas (live in the sheet). Crypto: the
 * value fetched from Coinbase (opts.fetched). Both fall back to the last known price, and both are compared with it
 * when that price is recent (see PRICE_SANITY, SANITY_MAX_AGE_DAYS).
 */
function holdingPriceFormulas(ticker, lastPrice, opts) {
  opts = opts || {};
  const recent = !(opts.lastPriceDay && opts.today) || daysBetween(opts.lastPriceDay, opts.today) <= SANITY_MAX_AGE_DAYS;
  if (isCryptoPair(ticker)) {
    const fetched = Number(opts.fetched) || 0;
    if (fetched > 0 && !(lastPrice > 0 && recent && Math.abs(fetched / lastPrice - 1) > PRICE_SANITY)) {
      return { price: fetched, source: 'Coinbase' + (opts.fetchedAt ? ' · ' + opts.fetchedAt : '') };
    }
    if (fetched > 0) return { price: lastPrice, source: 'last known — check Coinbase price' };
    return { price: lastPrice > 0 ? lastPrice : 0, source: lastPrice > 0 ? 'last known (Coinbase unavailable)' : 'no price' };
  }
  const symbol = priceSymbol(ticker);
  const live = 'GOOGLEFINANCE("' + symbol + '","price")';
  if (!(lastPrice > 0)) {
    return { price: '=IFERROR(' + live + ',0)', source: '=IF(ISERROR(' + live + '),"no price","live")' };
  }
  const last = formulaNumber(lastPrice);
  if (!recent) {
    return { price: '=IFERROR(' + live + ',' + last + ')', source: '=IF(ISERROR(' + live + '),"last known","live")' };
  }
  const off = 'ABS(' + live + '/' + last + '-1)>' + PRICE_SANITY;
  return {
    price: '=IFERROR(IF(' + off + ',' + last + ',' + live + '),' + last + ')',
    source: '=IFERROR(IF(' + off + ',"last known — check symbol","live"),"last known")'
  };
}

/**
 * Writes the Holdings sheet from computeHoldings() — v1.1.34: the Dashboard's look (title band, KPI cards,
 * sections, striped tables, total rows). Prices stay live: stocks/ETFs are GOOGLEFINANCE formulas.
 * Columns B–M of the positions table are fixed (the history reads Market value from column H).
 */
function buildHoldingsSheet(h, prices) {
  prices = prices || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HOLDINGS_SHEET);
  if (sheet) {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    sheet.clear();
    sheet.setConditionalFormatRules([]);
    sheet.getCharts().forEach(c => sheet.removeChart(c));   // sheet.clear() keeps embedded charts
  } else sheet = ss.insertSheet(HOLDINGS_SHEET);
  const T = DASH_THEME, border = SpreadsheetApp.BorderStyle, usd = '"US$"#,##0.00', year = new Date().getFullYear();
  const nPos = Math.max(h.positions.length, 1), nCash = h.cash.length, nVal = Math.max(h.valuations.length, 1);
  const R = { kpiLabel: 4, kpiValue: 5, kpiNote: 6, posHead: 8, posCols: 9, posFirst: 10 };
  R.posLast = R.posFirst + nPos - 1; R.cashFirst = R.posLast + 1; R.posTotal = R.cashFirst + nCash;
  R.valHead = R.posTotal + 2; R.valCols = R.valHead + 1; R.valFirst = R.valCols + 1; R.valLast = R.valFirst + nVal - 1;
  R.grand = R.valLast + 2; R.notes = R.grand + 2;
  ensureRowCapacity(sheet, R.notes + h.warnings.length + 40);

  const section = (row, c1, c2, text) => {
    sheet.getRange(row, c1, 1, c2 - c1 + 1).merge().setValue(text).setFontWeight('bold').setFontSize(11).setFontColor(T.navy)
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

  // ---- 1-2 title band
  sheet.setRowHeight(1, 46);
  sheet.setRowHeight(2, 24);
  sheet.getRange(1, 1, 2, 13).setBackground(T.navy);
  sheet.getRange('B1:M1').merge().setValue('📈  Investments').setFontSize(20).setFontWeight('bold').setFontColor('#FFFFFF');
  sheet.getRange('B2:M2').merge().setValue('Holdings  ·  stocks and ETFs from GOOGLEFINANCE, crypto from Coinbase  ·  built ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + '  ·  amounts in US$').setFontColor(T.line);
  sheet.setRowHeight(3, 10);

  // ---- 4-6 KPI cards (the return is filled in once the Performance block exists)
  const kpis = [
    { c: 2, label: 'PORTFOLIO VALUE', value: '=C' + R.grand, fmt: usd, note: '=IFERROR("RD$"&TEXT(C' + R.grand + '*RATE_USD,"#,##0")&" · DOP-equivalent","")' },
    { c: 4, label: 'INVESTED (COST BASIS)', value: '=F' + R.posTotal, fmt: usd, note: 'Positions + cash, from the ledger' },
    { c: 6, label: 'UNREALIZED P/L', value: '=I' + R.posTotal, fmt: '"+US$"#,##0.00;"−US$"#,##0.00',
      note: '=IFERROR(TEXT(J' + R.posTotal + ',"+0.00%;-0.00%")&" on cost","")' },
    { c: 8, label: 'RETURN SINCE START', value: '—', fmt: '+0.00%;-0.00%;0.00%', note: 'Deposits left out · Modified Dietz' },
    { c: 10, label: 'DIVIDENDS ' + year, value: h.totals.dividendsYtd, fmt: usd, note: 'Fees ' + year + ': US$' + h.totals.feesYtd.toFixed(2) },
    { c: 12, label: 'DEPOSITED ' + year, value: h.totals.contributionsYtd.DOP, fmt: '"RD$"#,##0',
      note: h.totals.contributionsYtd.USD ? '+ US$' + h.totals.contributionsYtd.USD.toFixed(2) : 'From bank transfers' }
  ];
  sheet.setRowHeight(R.kpiLabel, 22); sheet.setRowHeight(R.kpiValue, 36); sheet.setRowHeight(R.kpiNote, 22);
  kpis.forEach(k => {
    sheet.getRange(R.kpiLabel, k.c, 3, 2).mergeAcross().setBackground(T.soft).setHorizontalAlignment('left')
      .setBorder(null, true, null, true, null, null, '#FFFFFF', border.SOLID_THICK);
    sheet.getRange(R.kpiLabel, k.c, 1, 2).setBorder(true, null, null, null, null, null, T.accent, border.SOLID_THICK);
    sheet.getRange(R.kpiLabel, k.c).setValue(k.label).setFontSize(8).setFontWeight('bold').setFontColor(T.muted);
    sheet.getRange(R.kpiValue, k.c).setValue(k.value).setFontSize(18).setFontWeight('bold').setFontColor(T.navy).setNumberFormat(k.fmt);
    sheet.getRange(R.kpiNote, k.c).setValue(k.note).setFontSize(8).setFontColor(T.muted);
  });
  R.kpiReturnCol = 8;
  sheet.setRowHeight(7, 12);

  // ---- positions
  section(R.posHead, 2, 13, '📊  Positions');
  const heads = ['Account', 'Ticker', 'Shares', 'Avg cost', 'Cost basis', 'Price', 'Market value', 'Unrealized P/L', 'P/L %', 'Weight', 'Dividends', 'Price source'];
  tableHead(R.posCols, 2, heads);
  const posRows = h.positions.length ? h.positions.map((p, i) => {
    const r = R.posFirst + i, f = holdingPriceFormulas(p.ticker, p.lastPrice, { lastPriceDay: p.lastPriceDay,
      today: prices.today, fetched: (prices.fetched || {})[p.ticker], fetchedAt: prices.fetchedAt });
    return [p.account, p.ticker, p.qty, '=F' + r + '/D' + r, p.cost, f.price, '=D' + r + '*G' + r, '=H' + r + '-F' + r,
      '=IF(F' + r + '>0,I' + r + '/F' + r + ',"")', '=IF($H$' + R.posTotal + '>0,H' + r + '/$H$' + R.posTotal + ',"")', p.dividends, f.source];
  }) : [['No positions yet — add a Snapshot to the Investment Ledger', '', '', '', '', '', '', '', '', '', '', '']];
  sheet.getRange(R.posFirst, 2, posRows.length, heads.length).setValues(posRows);
  // v1.1.56: each stock's change in the last session, in a hidden helper column — the "biggest moves" of the daily report
  sheet.getRange(1, DAY_CHANGE_COL, sheet.getMaxRows(), 1).clearContent();
  if (h.positions.length) {
    h.positions.forEach((p, i) => {   // crypto has no session: left empty (its move comes from the history's prices)
      if (!isCryptoPair(p.ticker)) sheet.getRange(R.posFirst + i, DAY_CHANGE_COL).setFormula('=IFERROR(GOOGLEFINANCE("' + priceSymbol(p.ticker) + '","changepct")/100,"")');
    });
    sheet.hideColumns(DAY_CHANGE_COL);
  }
  if (nCash) {
    sheet.getRange(R.cashFirst, 2, nCash, heads.length).setValues(h.cash.map(c =>
      [c.account, 'Cash' + (c.estimated ? ' (≈)' : ''), '', '', c.amount, '', c.amount, '', '', '=IF($H$' + R.posTotal + '>0,H' +
        (R.cashFirst + h.cash.indexOf(c)) + '/$H$' + R.posTotal + ',"")', '', c.estimated ? 'DOP deposits at the Dashboard rate' : 'ledger']));
  }
  stripe(R.posFirst, 2, R.posTotal - R.posFirst, heads.length);
  sheet.getRange(R.posTotal, 2, 1, heads.length).setValues([['TOTAL', '', '', '',
    '=SUM(F' + R.posFirst + ':F' + (R.posTotal - 1) + ')', '', '=SUM(H' + R.posFirst + ':H' + (R.posTotal - 1) + ')',
    '=SUM(I' + R.posFirst + ':I' + (R.posTotal - 1) + ')', '=IF(F' + R.posTotal + '>0,I' + R.posTotal + '/F' + R.posTotal + ',"")', '',
    '=SUM(L' + R.posFirst + ':L' + (R.posTotal - 1) + ')', '']]);
  totalRow(R.posTotal, 2, heads.length);
  const nRows = R.posTotal - R.posFirst + 1;
  sheet.getRange(R.posFirst, 2, nRows - 1, 1).setFontColor(T.muted);
  sheet.getRange(R.posFirst, 3, nRows - 1, 1).setFontWeight('bold').setFontColor(T.navy);
  sheet.getRange(R.posFirst, 4, nRows, 1).setNumberFormat('#,##0.00000');
  sheet.getRange(R.posFirst, 5, nRows, 5).setNumberFormat(usd);
  sheet.getRange(R.posFirst, 5, nRows, 1).setNumberFormat(USD_PRICE_FORMAT);   // Avg cost
  sheet.getRange(R.posFirst, 7, nRows, 1).setNumberFormat(USD_PRICE_FORMAT);   // Price
  sheet.getRange(R.posFirst, 9, nRows, 1).setNumberFormat('"+US$"#,##0.00;"−US$"#,##0.00;"US$"0.00');
  sheet.getRange(R.posFirst, 10, nRows, 2).setNumberFormat('+0.00%;-0.00%;0.00%');
  sheet.getRange(R.posFirst, 11, nRows, 1).setNumberFormat('0.0%');
  sheet.getRange(R.posFirst, 12, nRows, 1).setNumberFormat(usd);
  sheet.getRange(R.posFirst, 13, nRows, 1).setFontSize(8).setHorizontalAlignment('center');

  // ---- other accounts (tracked by balance)
  section(R.valHead, 2, 7, '🏦  Other accounts');
  tableHead(R.valCols, 2, ['Account', 'Value', 'Currency', 'As of', 'Units × unit price', 'Value (US$)']);
  const dayDate = k => { const p = String(k).split('-').map(Number); return new Date(p[0], p[1] - 1, p[2], 12); };
  const valRows = h.valuations.length ? h.valuations.map((v, i) => {
    const r = R.valFirst + i;
    return [v.account, v.value, v.currency, dayDate(v.day),   // v1.1.34: a real date (was text, shown two ways)
      v.units && v.unitPrice ? (+Number(v.units).toFixed(4)).toLocaleString('en-US') + ' × ' +
        Number(v.unitPrice).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '',
      '=IF(D' + r + '="DOP",IFERROR(C' + r + '/RATE_USD,""),C' + r + ')'];
  }) : [['Add a balance: 📊 Tracker › ➕ Add Balance or Deposit', '', '', '', '', '']];
  sheet.getRange(R.valFirst, 2, valRows.length, 6).setValues(valRows);
  stripe(R.valFirst, 2, nVal, 6);
  sheet.getRange(R.valFirst, 3, nVal, 1).setNumberFormat('#,##0.00');
  sheet.getRange(R.valFirst, 4, nVal, 1).setHorizontalAlignment('center');
  sheet.getRange(R.valFirst, 5, nVal, 1).setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');
  sheet.getRange(R.valFirst, 6, nVal, 1).setFontSize(9).setFontColor(T.muted);
  sheet.getRange(R.valFirst, 7, nVal, 1).setNumberFormat(usd);

  // ---- grand total
  sheet.getRange(R.grand, 2, 1, 12).setBackground(T.navy).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12);
  sheet.setRowHeight(R.grand, 32);
  sheet.getRange(R.grand, 2, 1, 4).setValues([['TOTAL INVESTED',
    '=H' + R.posTotal + '+SUM(G' + R.valFirst + ':G' + R.valLast + ')', 'DOP-equivalent', '=IFERROR(C' + R.grand + '*RATE_USD,"")']]);
  sheet.getRange(R.grand, 3).setNumberFormat(usd);
  sheet.getRange(R.grand, 4).setHorizontalAlignment('right').setFontWeight('normal').setFontSize(9);
  sheet.getRange(R.grand, 5).setNumberFormat('"RD$"#,##0.00');

  // ---- notes and warnings
  const notes = ['Positions start from each account\'s latest Snapshot and add the movements after that day. ' +
    'Funds and pensions: 📊 Tracker › ➕ Add Balance or Deposit. This sheet is rebuilt on every refresh — edit the ledger, not this.']
    .concat(h.warnings.map(w => '⚠️ ' + w));
  sheet.getRange(R.notes, 2, notes.length, 1).setValues(notes.map(n => [n])).setFontSize(8).setFontColor(T.muted);
  R.notesEnd = R.notes + notes.length - 1;

  // ---- conditional looks: P/L colours, price-source chips, weight scale
  const rows = n => sheet.getRange(R.posFirst, n, nRows, 1);
  const rule = () => SpreadsheetApp.newConditionalFormatRule();
  sheet.setConditionalFormatRules([
    rule().whenNumberGreaterThan(0).setFontColor(T.good).setRanges([sheet.getRange(R.posFirst, 9, nRows, 2)]).build(),
    rule().whenNumberLessThan(0).setFontColor(T.bad).setRanges([sheet.getRange(R.posFirst, 9, nRows, 2)]).build(),
    rule().whenTextEqualTo('live').setBackground('#E8F5E9').setFontColor('#2E7D32').setRanges([rows(13)]).build(),
    rule().whenTextStartsWith('Coinbase').setBackground('#E0F2F1').setFontColor('#00695C').setRanges([rows(13)]).build(),
    rule().whenTextStartsWith('last known').setBackground('#FFF4D6').setFontColor('#92400E').setRanges([rows(13)]).build(),
    rule().setGradientMinpoint('#FFFFFF').setGradientMaxpoint('#C9D7EF').setRanges([sheet.getRange(R.posFirst, 11, nRows - 1, 1)]).build()
  ]);

  // ---- columns
  sheet.setColumnWidth(1, 16);
  sheet.setColumnWidth(2, 170);
  for (let c = 3; c <= 12; c++) sheet.setColumnWidth(c, 108);
  sheet.setColumnWidth(13, 200);
  sheet.setFrozenRows(2);
  sheet.setHiddenGridlines(true);
  return { sheet: sheet, rows: R };
}

function styleLedgerSheet(sheet) {
  styleHeader(sheet, LEDGER_HEADERS.length);
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, LG.DATE + 1, rows, 1).setNumberFormat('yyyy-MM-dd').setHorizontalAlignment('center');
  sheet.getRange(2, LG.ACCOUNT + 1, rows, 1).setFontWeight('bold').setFontColor(DASH_THEME.navy);
  sheet.getRange(2, LG.TICKER + 1, rows, 1).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, LG.QTY + 1, rows, 1).setNumberFormat('#,##0.00000###');
  sheet.getRange(2, LG.PRICE + 1, rows, 1).setNumberFormat('#,##0.00######');   // up to 8 decimals (SHIB-sized prices)
  sheet.getRange(2, LG.AMOUNT + 1, rows, 2).setNumberFormat('#,##0.00');
  sheet.getRange(2, LG.CURRENCY + 1, rows, 2).setHorizontalAlignment('center');
  sheet.getRange(2, LG.NOTES + 1, rows, 1).setFontColor('#6B7280');
  sheet.getRange(2, LG.TYPE + 1, rows, 1).setHorizontalAlignment('center').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(LEDGER_TYPES, true).setAllowInvalid(false).build());
  const colours = { Buy: ['#E8F5E9', '#2E7D32'], Sell: ['#FDECEC', '#B91C1C'], Dividend: ['#F3E8FF', '#7E22CE'],
    Deposit: ['#EAF1FE', '#1D4ED8'], Withdrawal: ['#FFF4E5', '#B45309'], Fee: ['#F3F4F6', '#4B5563'],
    Snapshot: ['#E0F2F1', '#00695C'], Valuation: ['#FEF9C3', '#854D0E'], Notice: ['#EEF2F7', '#475569'] };
  const sources = { email: ['#EAF1FE', '#1D4ED8'], bank: ['#E0F2F1', '#00695C'], manual: ['#F3F4F6', '#4B5563'] };
  const typeCol = sheet.getRange(2, LG.TYPE + 1, rows, 1), sourceCol = sheet.getRange(2, LG.SOURCE + 1, rows, 1);
  const rules = Object.keys(colours).map(t => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t)
    .setBackground(colours[t][0]).setFontColor(colours[t][1]).setRanges([typeCol]).build())
    .concat(Object.keys(sources).map(s => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s)
      .setBackground(sources[s][0]).setFontColor(sources[s][1]).setRanges([sourceCol]).build()));
  rules.unshift(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2="",$B2<>"")')   // v1.1.43: no date
    .setBackground(SHEET_THEME.problem).setRanges([sheet.getRange(2, 1, rows, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",ISEVEN(ROW()))')
    .setBackground(SHEET_THEME.stripe).setRanges([sheet.getRange(2, 1, rows, LEDGER_HEADERS.length)]).build());
  sheet.setConditionalFormatRules(rules);
  [95, 150, 95, 80, 115, 115, 110, 70, 75, 75, 280].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

/** v1.1.34: Portfolio History — day bands, TOTAL rows stand out, number formats. */
function styleHistorySheet(sheet) {
  styleHeader(sheet, HISTORY_HEADERS.length);
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 1, rows, 1).setNumberFormat('yyyy-MM-dd').setHorizontalAlignment('center');
  sheet.getRange(2, 3, rows, 1).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 4, rows, 1).setNumberFormat('#,##0.00000###');
  sheet.getRange(2, 5, rows, 1).setNumberFormat(USD_PRICE_FORMAT);
  sheet.getRange(2, 6, rows, 1).setNumberFormat('"US$"#,##0.00');
  const all = sheet.getRange(2, 1, rows, HISTORY_HEADERS.length);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="' + HISTORY_TOTAL + '"')
      .setBackground(DASH_THEME.total).setFontColor(DASH_THEME.navy).setBold(true).setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",ISEVEN(INT($A2)))')   // alternate by day
      .setBackground(HISTORY_DAY_BAND).setRanges([all]).build()]);   // v1.1.39: the row stripe was too faint to see days
  [95, 150, 90, 120, 120, 130].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

/** v1.1.34: Investment Accounts — kind chips, readable notes. */
function styleAccountsSheet(sheet) {
  styleHeader(sheet, ACCOUNTS_HEADERS.length);
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 1, rows, 1).setFontWeight('bold').setFontColor(DASH_THEME.navy);
  sheet.getRange(2, 2, rows, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 4, rows, 1).setWrap(true).setFontSize(9).setFontColor('#6B7280');
  const kinds = { Broker: ['#EAF1FE', '#1D4ED8'], Fund: ['#FEF9C3', '#854D0E'], Pension: ['#F3E8FF', '#7E22CE'], Other: ['#F3F4F6', '#4B5563'] };
  sheet.setConditionalFormatRules(Object.keys(kinds).map(k => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(k)
    .setBackground(kinds[k][0]).setFontColor(kinds[k][1]).setRanges([sheet.getRange(2, 2, rows, 1)]).build()));
  [170, 90, 150, 460].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

/** The investments part of a run: broker emails → ledger, bank deposits → ledger, Holdings rebuilt. */
function runInvestmentsStep(range) {
  const ledgerSheet = getOrCreateLedgerSheet();
  const ledgerIds = new Set(ledgerSheet.getLastRow() > 1
    ? ledgerSheet.getRange(2, LG.ID + 1, ledgerSheet.getLastRow() - 1, 1).getValues().map(r => String(r[0])).filter(Boolean) : []);
  const capture = range ? captureBrokerEmails(range, ledgerIds)
    : { events: [], parsed: 0, skipped: 0, unparsed: 0, threads: [], failedThreadIds: new Set(), unrecognized: [], readIds: [] };
  const accounts = readInvestmentAccounts();
  const tx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSACTIONS_SHEET);
  const deposits = tx ? brokerDepositsFromTransactions(tx.getDataRange().getValues(), accounts) : [];
  const saved = saveInvestmentEvents(capture.events.concat(deposits));
  if (capture.threads.length) markEmailsAsProcessed(capture.threads, capture.failedThreadIds);
  const holdings = refreshHoldings();
  return { parsed: capture.parsed, skipped: capture.skipped, unparsed: capture.unparsed, deposits: deposits.length,
    saved: saved.saved, duplicates: saved.duplicates, positions: holdings.positions.length, warnings: holdings.warnings,
    unrecognized: capture.unrecognized, readIds: capture.readIds };
}

/** Rebuilds Holdings from the ledger. */
function refreshHoldings() {
  const ledger = getOrCreateLedgerSheet();
  let usdRate = 0;
  try { usdRate = readDashboardRates(SpreadsheetApp.getActiveSpreadsheet()).USD; } catch (error) { usdRate = 0; }
  const h = computeHoldings(ledger.getDataRange().getValues(), { usdRate: usdRate, year: new Date().getFullYear() });
  const now = new Date(), todayKey = normalizeDateForCompare(now);
  const fetched = fetchCryptoPrices(h.positions.map(p => p.ticker));
  const built = buildHoldingsSheet(h, { today: todayKey, fetched: fetched,
    fetchedAt: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') });
  // v1.1.32: read what GOOGLEFINANCE computed, record today's values, then the Performance block
  let marketValues = null, live = null, dayChange = null;
  if (h.positions.length) {
    SpreadsheetApp.flush();
    // v1.1.56: GOOGLEFINANCE loads after the formulas are written; read too soon, every stock fell back to its last known
    // (ledger) price — so the history kept the same prices day after day: no "biggest moves", stale daily changes
    live = waitForLivePrices(built.sheet, built.rows.posFirst, h.positions.length);
    marketValues = built.sheet.getRange(built.rows.posFirst, 8, h.positions.length, 1).getValues().map(r => r[0]);
    dayChange = built.sheet.getRange(built.rows.posFirst, DAY_CHANGE_COL, h.positions.length, 1).getValues().map(r => r[0]);
  }
  const valueRows = holdingsValueRows(h, marketValues, { fetched: fetched, usdRate: usdRate, live: live, dayChange: dayChange,
    prevPrices: latestHistoryPrices(readHistory(), todayKey) });
  recordPortfolioHistory(valueRows, todayKey);
  const current = {};
  valueRows.forEach(r => { current[r.account] = (current[r.account] || 0) + (Number(r.value) || 0); });
  h.returns = computeReturns(ledger.getDataRange().getValues(), current, { usdRate: usdRate, today: todayKey });
  const byAccount = {};
  valueRows.forEach(r => { byAccount[r.account] = (byAccount[r.account] || 0) + (Number(r.value) || 0); });
  writePerformanceBlock(built.sheet, built.rows.notesEnd + 3, h.returns, readHistory(),
    Object.keys(byAccount).map(a => ({ account: a, value: byAccount[a] })), { row: built.rows.kpiValue, col: built.rows.kpiReturnCol });
  // v1.1.42: every deposit notice without a recorded deposit goes to Unrecognized; recorded ones drop off it
  const notices = checkDepositNotices(ledger.getDataRange().getValues());
  if (notices.missing.length || notices.matchedIds.length) {
    recordUnrecognized(notices.missing, notices.matchedIds, new Date());
    h.unmatchedDeposits = notices.missing.length;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNRECOGNIZED_SHEET);
    if (sheet) styleUnrecognizedSheet(sheet);
  }
  return h;
}

const NOTICE_WINDOW_BEFORE_DAYS = 7;   // a bank transfer usually leaves a few days before the broker completes the deposit
const NOTICE_WINDOW_AFTER_DAYS = 2;

/**
 * v1.1.42: pairs each deposit Notice with a recorded Deposit of the same account dated from 7 days before to 2 days after
 * it (each deposit used once, the closest first). Notices on or before the account's tracking start (its first Snapshot
 * or balance) are inside the start value and aren't checked. Returns the Unrecognized entries for the unpaired ones and
 * the Gmail ids of the paired ones. Pure.
 */
function checkDepositNotices(ledgerValues) {
  const rows = ledgerValues.slice(1).filter(r => r[LG.DATE] && String(r[LG.ACCOUNT]).trim())
    .map(r => ({ r: r, key: normalizeDateForCompare(r[LG.DATE]), account: String(r[LG.ACCOUNT]).trim() }));
  const start = {};
  rows.filter(x => x.r[LG.TYPE] === 'Snapshot' || x.r[LG.TYPE] === 'Valuation')
    .forEach(x => { if (!start[x.account] || x.key < start[x.account]) start[x.account] = x.key; });
  const deposits = rows.filter(x => x.r[LG.TYPE] === 'Deposit').map(x => ({ account: x.account, key: x.key, used: false }));
  const missing = [], matchedIds = [];
  // v1.1.45: a Notice you turned into the Deposit itself (type changed, amount typed — same row, same Id) is resolved;
  // its Unrecognized row used to stay forever, since no Notice was left to pair
  rows.filter(x => x.r[LG.TYPE] === 'Deposit' && /^gmail:/.test(String(x.r[LG.ID] || '')))
    .forEach(x => matchedIds.push(String(x.r[LG.ID]).replace(/^gmail:/, '')));
  rows.filter(x => x.r[LG.TYPE] === 'Notice').sort((a, b) => a.key < b.key ? -1 : 1).forEach(n => {
    const id = String(n.r[LG.ID] || '').replace(/^gmail:/, '');
    if (start[n.account] && n.key <= start[n.account]) { if (id) matchedIds.push(id); return; }
    const candidates = deposits.filter(d => !d.used && d.account === n.account &&
      daysBetween(d.key, n.key) <= NOTICE_WINDOW_BEFORE_DAYS && daysBetween(n.key, d.key) <= NOTICE_WINDOW_AFTER_DAYS)
      .sort((a, b) => Math.abs(daysBetween(a.key, n.key)) - Math.abs(daysBetween(b.key, n.key)));
    if (candidates.length) {
      candidates[0].used = true;
      if (id) matchedIds.push(id);
    } else {
      missing.push({ id: id || ('notice:' + n.account + ':' + n.key), date: n.r[LG.DATE], bank: n.account, subject: 'Deposit Completed',
        reason: 'Deposit without amount — add it: 📊 Tracker › ➕ Add Balance or Deposit › A deposit (' + n.key + ')',
        snippet: n.account + ' confirmed a deposit on ' + n.key + ' but its email has no amount, and no deposit is recorded near that date.' });
    }
  });
  return { missing: missing, matchedIds: matchedIds };
}

/** Menu "📈 Refresh Investments": deposits from bank transfers + Holdings (emails are read by the regular runs). */
function refreshInvestmentsNow() {
  const config = requireConfig();
  if (!config) return;
  withRunLock(() => {
    try {
      safeToast('Refreshing investments...', '📈 Investments', -1);
      const r = runInvestmentsStep(null);
      formatDataSheets();
      ensureSheetOrder();
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOLDINGS_SHEET);
      if (sheet) SpreadsheetApp.setActiveSheet(sheet);
      safeToast('Done.', '📈 Investments', 3);   // v1.1.33: the progress toast used to stay open
      safeAlert('📈 Investments refreshed\n\n' + r.positions + ' position(s) · ' + r.saved + ' new ledger row(s) from bank deposits' +
        (r.warnings.length ? '\n\n⚠️ ' + r.warnings.join('\n⚠️ ') : ''));
    } catch (error) {
      safeAlert('❌ Could not refresh investments: ' + error);
    }
  });
}

/* ======================================================================
 * PORTFOLIO HISTORY, RETURNS AND REPORT DATA — v1.1.32
 * ====================================================================== */
const HISTORY_SHEET = 'Portfolio History';
// v1.1.56: Day change — the market's own change of the last session (GOOGLEFINANCE "changepct"), for stocks and ETFs
const HISTORY_HEADERS = ['Date', 'Account', 'Ticker', 'Quantity', 'Price (US$)', 'Value (US$)', 'Day change'];
const DAY_CHANGE_COL = 20;          // Holdings: a hidden helper column (T) with each position's changepct
const LIVE_PRICE_WAIT_MS = 15000;   // how long a refresh waits for GOOGLEFINANCE to load
const LIVE_PRICE_CALM_POLLS = 4;    // …and stops earlier after this many checks without progress (a symbol it doesn't know)
const HISTORY_TOTAL = 'TOTAL';
const HISTORY_DAY_BAND = '#E8EEF8';   // v1.1.39: every other day, so one day's rows read as a block
const ANNUALIZE_MIN_DAYS = 180;   // annualizing a few weeks of returns gives absurd numbers

/**
 * Every position, cash balance and balance-tracked account in US$ — the Market value
 * GOOGLEFINANCE already computed in Holdings (marketValues, one per position row), or
 * quantity × (Coinbase price or last known price) when that cell isn't a number yet. Pure.
 */
function holdingsValueRows(h, marketValues, opts) {
  opts = opts || {};
  const fetched = opts.fetched || {}, usdRate = Number(opts.usdRate) || 0;
  const rows = h.positions.map((p, i) => {
    const mv = marketValues ? Number(marketValues[i]) : NaN;
    const price = isFinite(mv) && mv > 0 ? mv / p.qty : (fetched[p.ticker] > 0 ? fetched[p.ticker] : p.lastPrice);
    const dc = opts.dayChange ? opts.dayChange[i] : '';
    const dayChange = typeof dc === 'number' && isFinite(dc) ? dc : '';
    if (opts.live && !opts.live[i]) {
      // v1.1.56: no live price this time — valued at the latest price the history has (not the snapshot's), and no
      // price recorded for today, so it can't make a move that didn't happen
      const prev = (opts.prevPrices || {})[p.account + '|' + p.ticker];
      const at = prev > 0 ? prev : price;
      return { account: p.account, ticker: p.ticker, qty: p.qty, price: '', value: p.qty * at, dayChange: dayChange, stale: true };
    }
    return { account: p.account, ticker: p.ticker, qty: p.qty, price: price, value: isFinite(mv) && mv > 0 ? mv : p.qty * price, dayChange: dayChange };
  });
  h.cash.forEach(c => rows.push({ account: c.account, ticker: 'CASH', qty: '', price: '', value: c.amount }));
  h.valuations.forEach(v => rows.push({ account: v.account, ticker: '', qty: v.units || '', price: '',
    value: v.currency === 'DOP' ? (usdRate > 0 ? v.value / usdRate : 0) : v.value }));
  return rows;
}

/** Writes today's rows (replacing any earlier run of the same day) and a TOTAL row. */
function recordPortfolioHistory(valueRows, todayKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET);
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
  }
  const last = sheet.getLastRow();
  sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);   // v1.1.56: a history from before has 6 columns
  const kept = last > 1 ? sheet.getRange(2, 1, last - 1, HISTORY_HEADERS.length).getValues()
    .filter(r => r[0] && normalizeDateForCompare(r[0]) !== todayKey) : [];
  const parts = todayKey.split('-').map(Number);
  const day = new Date(parts[0], parts[1] - 1, parts[2], 12);
  const total = valueRows.reduce((s, r) => s + (Number(r.value) || 0), 0);
  const today = valueRows.map(r => [day, r.account, r.ticker, r.qty, r.price === '' ? '' : +Number(r.price).toPrecision(10), +Number(r.value).toFixed(2),
    r.dayChange === '' || r.dayChange === undefined ? '' : +Number(r.dayChange).toPrecision(6)])
    .concat([[day, HISTORY_TOTAL, '', '', '', +total.toFixed(2), '']]);
  const all = today.concat(kept);
  if (last > 1) sheet.getRange(2, 1, last - 1, HISTORY_HEADERS.length).clearContent();
  ensureRowCapacity(sheet, all.length + 1);
  sheet.getRange(2, 1, all.length, HISTORY_HEADERS.length).setValues(all);
  sortSheetByDateDesc(sheet, 1);
  return { rows: today.length, total: total };
}

function readHistory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HISTORY_SHEET);
  return sheet ? sheet.getDataRange().getValues() : [HISTORY_HEADERS];
}

/**
 * Newton's method on Σ cf_i / (1+r)^(t_i/365) = 0. Flows: [{ day: 'yyyy-MM-dd', amount }]
 * (money in = negative). Returns the annual rate, or null when it doesn't converge.
 */
function xirr(flows) {
  if (flows.length < 2 || !flows.some(f => f.amount < 0) || !flows.some(f => f.amount > 0)) return null;
  const t0 = flows[0].day;
  const years = flows.map(f => daysBetween(t0, f.day) / 365);
  let r = 0.1;
  for (let k = 0; k < 100; k++) {
    let f = 0, df = 0;
    flows.forEach((c, i) => { const d = Math.pow(1 + r, years[i]); f += c.amount / d; df -= years[i] * c.amount / (d * (1 + r)); });
    if (Math.abs(df) < 1e-12) return null;
    const next = r - f / df;
    if (!isFinite(next) || next <= -0.9999) return null;
    if (Math.abs(next - r) < 1e-9) return next;
    r = next;
  }
  return null;
}

/**
 * Return since each account started being tracked — its latest Snapshot, first Valuation,
 * or first movement — and for all accounts together. Period return: Modified Dietz
 * (deposits weighted by how long they were invested). Annualized (XIRR) only after
 * ANNUALIZE_MIN_DAYS. DOP flows and values are converted at opts.usdRate. Pure.
 */
function computeReturns(ledgerValues, currentValues, opts) {
  opts = opts || {};
  const usdRate = Number(opts.usdRate) || 0, today = opts.today;
  const toUsd = (amount, cur) => String(cur || 'USD').toUpperCase() === 'DOP' ? (usdRate > 0 ? amount / usdRate : 0) : amount;
  const rows = ledgerValues.slice(1).filter(r => r[LG.DATE] && String(r[LG.ACCOUNT]).trim())
    .map(r => ({ r: r, key: normalizeDateForCompare(r[LG.DATE]), account: String(r[LG.ACCOUNT]).trim() }))
    .sort((a, b) => a.key < b.key ? -1 : a.key > b.key ? 1 : 0);
  const accounts = {};
  Object.keys(currentValues).forEach(a => { accounts[a] = { account: a, start: null, startValue: 0, flows: [] }; });
  const snapDay = {};
  rows.filter(x => x.r[LG.TYPE] === 'Snapshot').forEach(x => { if (!snapDay[x.account] || x.key > snapDay[x.account]) snapDay[x.account] = x.key; });
  Object.keys(accounts).forEach(a => {
    const acc = accounts[a];
    const mine = rows.filter(x => x.account === a);
    if (snapDay[a]) {
      // v1.1.39: from the EARLIEST of its first Snapshot and its first balance (e.g. the broker's value on January 1)
      // — it used to start at the latest snapshot, so a portfolio snapshotted in September was "tracked since" September
      const firstSnap = mine.find(x => x.r[LG.TYPE] === 'Snapshot');
      const firstVal = mine.find(x => x.r[LG.TYPE] === 'Valuation');
      if (firstVal && firstVal.key < firstSnap.key) {
        acc.start = firstVal.key;
        acc.startValue = toUsd(invNumber(firstVal.r[LG.AMOUNT]) || invNumber(firstVal.r[LG.QTY]) * invNumber(firstVal.r[LG.PRICE]),
          firstVal.r[LG.CURRENCY]);
      } else {
        acc.start = firstSnap.key;
        mine.filter(x => x.r[LG.TYPE] === 'Snapshot' && x.key === firstSnap.key).forEach(x => {
          const t = String(x.r[LG.TICKER]).toUpperCase().trim();
          acc.startValue += t === 'CASH' ? invNumber(x.r[LG.AMOUNT]) : invNumber(x.r[LG.QTY]) * invNumber(x.r[LG.PRICE]);
        });
      }
    } else {
      const firstVal = mine.find(x => x.r[LG.TYPE] === 'Valuation');
      if (firstVal && !mine.some(x => x.r[LG.TYPE] === 'Deposit' && x.key < firstVal.key)) {
        acc.start = firstVal.key;
        const v = invNumber(firstVal.r[LG.AMOUNT]) || invNumber(firstVal.r[LG.QTY]) * invNumber(firstVal.r[LG.PRICE]);
        acc.startValue = toUsd(v, firstVal.r[LG.CURRENCY]);
      } else if (mine.length) {
        acc.start = mine[0].key;                         // starts from zero; its first deposit is a flow
        acc.startValue = 0;
      }
    }
    if (!acc.start) return;
    mine.forEach(x => {
      const type = x.r[LG.TYPE];
      if (type !== 'Deposit' && type !== 'Withdrawal') return;
      if (acc.startValue > 0 ? x.key <= acc.start : x.key < acc.start) return;
      const amount = toUsd(invNumber(x.r[LG.AMOUNT]), x.r[LG.CURRENCY]);
      acc.flows.push({ day: x.key, amount: type === 'Deposit' ? amount : -amount });
    });
  });
  const measure = (start, startValue, flows, value) => {
    if (!start || !today) return null;
    const T = Math.max(daysBetween(start, today), 0);
    const net = flows.reduce((s, f) => s + f.amount, 0);
    const weighted = flows.reduce((s, f) => s + f.amount * (T > 0 ? (T - daysBetween(start, f.day)) / T : 1), 0);
    const gain = value - startValue - net;
    const base = startValue + weighted;
    const annual = T >= ANNUALIZE_MIN_DAYS
      ? xirr([{ day: start, amount: -startValue }].concat(flows.map(f => ({ day: f.day, amount: -f.amount })))
        .concat([{ day: today, amount: value }]).filter(f => f.amount !== 0)) : null;
    return { start: start, days: T, startValue: startValue, netDeposits: net, value: value, gain: gain,
      periodReturn: base > 0 ? gain / base : null, annualized: annual };
  };
  const list = Object.keys(accounts).map(a => {
    const acc = accounts[a];
    const m = Object.assign({ account: a }, measure(acc.start, acc.startValue, acc.flows, currentValues[a]) || {});
    if (m.start) {
      // v1.1.40: money spent on purchases since the start (net of sales and dividends) had to come from somewhere —
      // deposits, or cash already in the account. Far more bought than deposited usually means deposits are missing,
      // and every missing deposit shows up as gain (reported: HAPI's 2026 gain included purchases' funding).
      let bought = 0;
      rows.filter(x => x.account === a && x.key > m.start).forEach(x => {
        const type = x.r[LG.TYPE], amt = toUsd(invNumber(x.r[LG.AMOUNT]), x.r[LG.CURRENCY]), fee = toUsd(invNumber(x.r[LG.FEE]), x.r[LG.CURRENCY]);
        if (type === 'Buy') bought += amt + fee;
        else if (type === 'Sell') bought -= amt - fee;
        else if (type === 'Dividend') bought -= amt;
      });
      m.netBought = bought;
      // v1.1.44: a sale made on the start day is already inside the start value as cash (e.g. a sale on Dec 31 that
      // settles in January) — its proceeds fund purchases without any deposit
      m.startCash = rows.filter(x => x.account === a && x.key === m.start && x.r[LG.TYPE] === 'Sell')
        .reduce((t, x) => t + toUsd(invNumber(x.r[LG.AMOUNT]) - invNumber(x.r[LG.FEE]), x.r[LG.CURRENCY]), 0);
      const gap = bought - m.netDeposits - m.startCash;
      m.unfunded = gap > Math.max(50, bought * 0.05) ? gap : 0;
    }
    return m;
  }).filter(x => x.start);
  // all accounts together: accounts that start later enter as a flow on their start day
  const started = list.map(x => x.start).sort();
  const start = started[0] || null;
  const flows = [];
  let startValue = 0, value = 0;
  list.forEach(x => {
    value += x.value;
    if (x.start === start) startValue += x.startValue;
    else if (x.startValue) flows.push({ day: x.start, amount: x.startValue });
    accounts[x.account].flows.forEach(f => flows.push(f));
  });
  flows.sort((a, b) => a.day < b.day ? -1 : 1);
  const total = start ? measure(start, startValue, flows, value) : null;
  if (total) {
    // v1.1.34: the return counts later-starting accounts as flows, but the row must read right — "Start value" is every
    // account's start value and "Net deposits" only real deposits (it showed accounts' start values as deposits)
    total.startValue = list.reduce((s, x) => s + x.startValue, 0);
    total.netDeposits = list.reduce((s, x) => s + x.netDeposits, 0);
  }
  return { accounts: list, total: total };
}

/**
 * What the daily email shows: today's total vs the previous recorded day, the gain in
 * between without deposits, and the positions whose price moved most. Pure.
 */
function investmentsDailyBrief(historyValues, ledgerValues, opts) {
  opts = opts || {};
  const usdRate = Number(opts.usdRate) || 0;
  const rows = historyValues.slice(1).filter(r => r[0]).map(r => ({ day: normalizeDateForCompare(r[0]), account: r[1],
    ticker: String(r[2] || ''), price: invNumber(r[4]), value: invNumber(r[5]),
    dayChange: r[6] === '' || r[6] === null || r[6] === undefined ? null : Number(r[6]) }));
  const days = rows.filter(r => r.account === HISTORY_TOTAL).map(r => r.day).sort();
  if (!days.length) return null;
  const today = days[days.length - 1], prev = days.length > 1 ? days[days.length - 2] : null;
  const totalOn = d => rows.filter(r => r.day === d && r.account === HISTORY_TOTAL).reduce((s, r) => s + r.value, 0);
  const brief = { day: today, total: totalOn(today), totalDop: usdRate > 0 ? totalOn(today) * usdRate : null, prevDay: prev,
    change: null, changePct: null, deposits: 0, movers: [], added: [], returns: opts.returns || null,
    // v1.1.56: positions that had no live price at the update (recorded without a price)
    stale: rows.filter(r => r.day === today && r.account !== HISTORY_TOTAL && r.ticker && r.ticker !== 'CASH' && !(r.price > 0) && r.value > 0).length };
  // v1.1.56: the market's own change of the last session — so a Monday shows Friday's moves instead of nothing
  const session = rows.filter(r => r.day === today && r.dayChange !== null && isFinite(r.dayChange))
    .map(r => ({ ticker: r.ticker, account: r.account, change: r.dayChange, session: true }));
  if (prev) {
    // v1.1.37: compared account by account. An account added to the tracker since the previous day (a fund's first
    // balance, say) isn't gain — comparing totals counted it as one (+US$17,104 on the day two accounts were added).
    const byAccount = d => {
      const m = {};
      rows.filter(r => r.day === d && r.account !== HISTORY_TOTAL).forEach(r => { m[r.account] = (m[r.account] || 0) + r.value; });
      return m;
    };
    const now = byAccount(today), then = byAccount(prev);
    const common = Object.keys(now).filter(a => a in then);
    brief.added = Object.keys(now).filter(a => !(a in then));
    ledgerValues.slice(1).forEach(r => {
      if (!r[LG.DATE] || (r[LG.TYPE] !== 'Deposit' && r[LG.TYPE] !== 'Withdrawal')) return;
      if (common.indexOf(String(r[LG.ACCOUNT]).trim()) === -1) return;
      const k = normalizeDateForCompare(r[LG.DATE]);
      if (k <= prev || k > today) return;
      const cur = String(r[LG.CURRENCY] || 'USD').toUpperCase();
      const amount = cur === 'DOP' ? (usdRate > 0 ? invNumber(r[LG.AMOUNT]) / usdRate : 0) : invNumber(r[LG.AMOUNT]);
      brief.deposits += r[LG.TYPE] === 'Deposit' ? amount : -amount;
    });
    const before = common.reduce((t, a) => t + then[a], 0), after = common.reduce((t, a) => t + now[a], 0);
    brief.change = after - before - brief.deposits;
    brief.changePct = before > 0 ? brief.change / before : null;
    const prevPrice = {};
    rows.filter(r => r.day === prev && r.price > 0).forEach(r => { prevPrice[r.account + '|' + r.ticker] = r.price; });
    const bySession = {};
    session.forEach(m => { bySession[m.account + '|' + m.ticker] = true; });
    const byPrice = rows.filter(r => r.day === today && r.price > 0 && prevPrice[r.account + '|' + r.ticker] > 0 && !bySession[r.account + '|' + r.ticker])
      .map(r => ({ ticker: r.ticker, account: r.account, change: r.price / prevPrice[r.account + '|' + r.ticker] - 1 }));
    brief.movers = session.concat(byPrice).filter(m => Math.abs(m.change) >= 0.0005)
      .sort((a, b) => Math.abs(b.change) - Math.abs(a.change)).slice(0, 3);
  } else {
    brief.movers = session.filter(m => Math.abs(m.change) >= 0.0005).sort((a, b) => Math.abs(b.change) - Math.abs(a.change)).slice(0, 3);
  }
  return brief;
}

/** What the monthly email shows for the month of opts.month: value change, deposits, gain, dividends, fees, allocation. Pure. */
function investmentsMonthlyBrief(historyValues, ledgerValues, opts) {
  opts = opts || {};
  const usdRate = Number(opts.usdRate) || 0;
  const monthKey = normalizeDateForCompare(opts.month).slice(0, 7);
  const rows = historyValues.slice(1).filter(r => r[0]).map(r => ({ day: normalizeDateForCompare(r[0]), account: r[1], value: invNumber(r[5]) }));
  const days = rows.filter(r => r.account === HISTORY_TOTAL).map(r => r.day).sort();
  const inMonth = days.filter(d => d.slice(0, 7) === monthKey);
  if (!inMonth.length) return null;
  const endDay = inMonth[inMonth.length - 1];
  const before = days.filter(d => d.slice(0, 7) < monthKey);
  const startDay = before.length ? before[before.length - 1] : inMonth[0];
  const valueOn = d => {
    const m = {};
    rows.filter(r => r.day === d && r.account !== HISTORY_TOTAL).forEach(r => { m[r.account] = (m[r.account] || 0) + r.value; });
    return m;
  };
  // v1.1.37: account by account — an account's start is its value when the month began, or, if it was added during
  // the month, its first recorded value (so adding an account is not counted as gain)
  const end = valueOn(endDay), start = valueOn(startDay), from = {}, added = [];
  const accounts = Object.keys(end).concat(Object.keys(start).filter(a => !(a in end)));
  accounts.forEach(a => {
    if (a in start) { from[a] = { day: startDay, value: start[a] }; return; }
    const first = inMonth.find(d => rows.some(r => r.day === d && r.account === a));
    from[a] = { day: first, value: valueOn(first)[a] || 0 };
    added.push(a);
  });
  const brief = { endDay: endDay, startDay: startDay, startValue: accounts.reduce((t, a) => t + from[a].value, 0),
    endValue: accounts.reduce((t, a) => t + (end[a] || 0), 0), deposits: 0, dividends: 0, fees: 0, allocation: [],
    added: added, returns: opts.returns || null, partial: !before.length };
  ledgerValues.slice(1).forEach(r => {
    if (!r[LG.DATE]) return;
    const k = normalizeDateForCompare(r[LG.DATE]);
    const cur = String(r[LG.CURRENCY] || 'USD').toUpperCase();
    const usd = v => cur === 'DOP' ? (usdRate > 0 ? v / usdRate : 0) : v;
    if (k.slice(0, 7) === monthKey) {
      if (r[LG.TYPE] === 'Dividend') brief.dividends += usd(invNumber(r[LG.AMOUNT]));
      if (r[LG.TYPE] === 'Fee') brief.fees += usd(invNumber(r[LG.AMOUNT]));
      if (r[LG.TYPE] === 'Buy' || r[LG.TYPE] === 'Sell') brief.fees += usd(invNumber(r[LG.FEE]));
    }
    const a = String(r[LG.ACCOUNT]).trim();
    if ((r[LG.TYPE] === 'Deposit' || r[LG.TYPE] === 'Withdrawal') && from[a] && k > from[a].day && k <= endDay) {
      brief.deposits += (r[LG.TYPE] === 'Deposit' ? 1 : -1) * usd(invNumber(r[LG.AMOUNT]));
    }
  });
  brief.gain = brief.endValue - brief.startValue - brief.deposits;
  brief.gainPct = brief.startValue > 0 ? brief.gain / brief.startValue : null;
  brief.allocation = Object.keys(end).map(a => ({ account: a, value: end[a], share: brief.endValue > 0 ? end[a] / brief.endValue : 0 }))
    .sort((x, y) => y.value - x.value);
  return brief;
}

/** Holdings: the Performance block (Dashboard look), the KPI return, an allocation pie and the history line chart. */
function writePerformanceBlock(sheet, startRow, returns, historyValues, allocation, kpi) {
  const T = DASH_THEME, border = SpreadsheetApp.BorderStyle, usd = '"US$"#,##0.00';
  sheet.getRange(startRow, 2, 1, 8).merge().setValue('📐  Performance').setFontWeight('bold').setFontSize(11).setFontColor(T.navy)
    .setBorder(null, null, true, null, null, null, T.accent, border.SOLID_MEDIUM);
  sheet.setRowHeight(startRow, 28);
  const heads = ['Account', 'Tracked since', 'Start value', 'Net deposits', 'Value now', 'Gain', 'Return', 'Annualized'];
  sheet.getRange(startRow + 1, 2, 1, heads.length).setValues([heads]).setFontWeight('bold').setFontSize(8).setFontColor(T.muted)
    .setBackground(T.soft).setBorder(null, null, true, null, null, null, T.line, border.SOLID);
  const dayDate = k => { const p = String(k).split('-').map(Number); return new Date(p[0], p[1] - 1, p[2], 12); };
  const line = x => [x.account, dayDate(x.start), x.startValue, x.netDeposits, x.value, x.gain,   // v1.1.34: real dates
    x.periodReturn === null ? '—' : x.periodReturn, x.annualized === null ? 'after ' + ANNUALIZE_MIN_DAYS + ' days' : x.annualized];
  const body = returns.accounts.map(line);
  if (returns.total) body.push(line(Object.assign({ account: 'All accounts' }, returns.total)));
  if (!body.length) body.push(['Performance appears after the first refresh', '', '', '', '', '', '', '']);
  const first = startRow + 2;
  sheet.getRange(first, 2, body.length, heads.length).setValues(body)
    .setBackgrounds(body.map((_, i) => new Array(heads.length).fill(i % 2 ? T.stripe : '#FFFFFF')));
  sheet.getRange(first, 3, body.length, 1).setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');
  sheet.getRange(first, 4, body.length, 4).setNumberFormat(usd);
  sheet.getRange(first, 7, body.length, 1).setNumberFormat('"+US$"#,##0.00;"−US$"#,##0.00;"US$"0.00');
  sheet.getRange(first, 8, body.length, 2).setNumberFormat('+0.00%;-0.00%;0.00%').setHorizontalAlignment('right');
  sheet.getRange(first, 9, body.length, 1).setFontColor(T.muted);
  const totalRow = first + body.length - 1;
  if (returns.total) {
    sheet.getRange(totalRow, 2, 1, heads.length).setFontWeight('bold').setBackground(T.total)
      .setBorder(true, null, null, null, null, null, T.accent, border.SOLID);
    if (kpi) {   // the KPI card shows the all-accounts return
      sheet.getRange(kpi.row, kpi.col).setFormula('=H' + totalRow);
      sheet.getRange(kpi.row + 1, kpi.col).setValue('Since ' + returns.total.start + ' · deposits left out');
    }
  }
  // v1.1.40: accounts whose purchases since the start exceed their recorded deposits
  const gaps = returns.accounts.filter(x => x.unfunded > 0).map(x => '⚠️ ' + x.account + ': purchases since ' + x.start + ' (net US$' +
    x.netBought.toFixed(2) + ') exceed the deposits recorded (US$' + x.netDeposits.toFixed(2) + ')' +
    (x.startCash ? ' and the start-day sale proceeds (US$' + x.startCash.toFixed(2) + ')' : '') + ' by US$' + x.unfunded.toFixed(2) +
    ' — unless that came from cash already in the account, deposits are missing and show up as gain. Add them: 📊 Tracker › ➕ Add Balance or Deposit.');
  if (gaps.length) {
    sheet.getRange(totalRow + 1, 2, gaps.length, 1).setValues(gaps.map(g => [g])).setFontSize(9).setFontColor('#B45309').setFontWeight('bold');
  }
  const signed = sheet.getRange(first, 7, body.length, 2);
  sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(T.good).setRanges([signed]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(T.bad).setRanges([signed]).build()]));

  // charts: value over time (line) and where it is (pie), side by side
  const chartRow = totalRow + 3 + gaps.length;
  const col = 16;   // P:S, hidden helper tables
  const totals = historyValues.slice(1).filter(r => r[0] && r[1] === HISTORY_TOTAL)
    .map(r => [r[0], invNumber(r[5])]).sort((a, b) => normalizeDateForCompare(a[0]) < normalizeDateForCompare(b[0]) ? -1 : 1);
  const alloc = (allocation || []).filter(a => a.value > 0).sort((a, b) => b.value - a.value);
  sheet.getRange(chartRow - 1, 2, 1, 12).merge().setValue('📈  Value over time and allocation').setFontWeight('bold').setFontSize(11)
    .setFontColor(T.navy).setBorder(null, null, true, null, null, null, T.accent, border.SOLID_MEDIUM);
  if (totals.length >= 2) {
    sheet.getRange(1, col, 1, 2).setValues([['Date', 'Total (US$)']]);
    ensureRowCapacity(sheet, totals.length + 1);
    sheet.getRange(2, col, totals.length, 2).setValues(totals);
    sheet.getRange(2, col, totals.length, 1).setNumberFormat('yyyy-mm-dd');
    sheet.insertChart(sheet.newChart().setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(1, col, totals.length + 1, 2)).setNumHeaders(1)
      .setPosition(chartRow + 1, 2, 0, 0).setOption('title', 'Portfolio value (US$)').setOption('legend', { position: 'none' })
      .setOption('colors', [T.accent]).setOption('width', 620).setOption('height', 280).build());
  } else {
    sheet.getRange(chartRow + 1, 2, 1, 5).merge().setValue('The value chart appears once there are two days of history.')
      .setFontSize(9).setFontColor(T.muted);
  }
  if (alloc.length) {
    sheet.getRange(1, col + 2, 1, 2).setValues([['Account', 'Value (US$)']]);
    ensureRowCapacity(sheet, alloc.length + 1);
    sheet.getRange(2, col + 2, alloc.length, 2).setValues(alloc.map(a => [a.account, +a.value.toFixed(2)]));
    sheet.insertChart(sheet.newChart().setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange(1, col + 2, alloc.length + 1, 2)).setNumHeaders(1)
      .setPosition(chartRow + 1, 8, 20, 0).setOption('title', 'Allocation by account').setOption('pieHole', 0.45)
      .setOption('colors', [T.navy, T.accent, '#0F766E', '#7FA6E8', '#C9D7EF', '#A6A6A6'])
      .setOption('width', 460).setOption('height', 280).build());
  }
  sheet.hideColumns(col, 4);
  return chartRow;
}

/** For the summary emails: the investments brief from Portfolio History and the ledger (null when there's none). */
function investmentsReportData(kind, opts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName(INVESTMENT_LEDGER_SHEET);
  const history = readHistory();
  if (!ledgerSheet || history.length < 2) return null;
  const ledger = ledgerSheet.getDataRange().getValues();
  const usdRate = opts.rates && opts.rates.USD || 0;
  const rows = history.slice(1).filter(r => r[0]);
  const latest = rows.map(r => normalizeDateForCompare(r[0])).sort().pop();
  const current = {};
  rows.filter(r => normalizeDateForCompare(r[0]) === latest && r[1] !== HISTORY_TOTAL)
    .forEach(r => { current[r[1]] = (current[r[1]] || 0) + invNumber(r[5]); });
  const returns = computeReturns(ledger, current, { usdRate: usdRate, today: latest }).total;
  if (kind === 'daily') return investmentsDailyBrief(history, ledger, { usdRate: usdRate, returns: returns });
  const monthly = investmentsMonthlyBrief(history, ledger, { usdRate: usdRate, returns: returns, month: opts.month });
  if (monthly) return monthly;
  // v1.1.49: a month before the history began used to leave the section out without a word (reported: "nothing about
  // investments in the monthly report") — say when it starts instead
  const first = rows.map(r => normalizeDateForCompare(r[0])).sort()[0];
  const p = first.split('-').map(Number);
  return { none: true, firstDay: first, firstReport: Utilities.formatDate(new Date(p[0], p[1], 1, 12), Session.getScriptTimeZone(), 'yyyy-MM-dd') };
}

/* ======================================================================
 * FUND / PENSION BALANCES — v1.1.33
 * A dialog that adds a Valuation row, so balances aren't typed by hand into the ledger.
 * ====================================================================== */
const VALUATION_KINDS = ['Fund', 'Pension', 'Broker', 'Other'];

/** null when the entry can be saved, otherwise what's wrong. Pure. */
function validateValuationEntry(e, todayKey) {
  if (!e || !String(e.account || '').trim()) return 'Give the account a name.';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(e.date || ''))) return 'Pick the statement date.';
  if (todayKey && e.date > todayKey) return 'The statement date is in the future.';
  if (['DOP', 'USD'].indexOf(e.currency) === -1) return 'Currency must be DOP or USD.';
  if (e.mode === 'units') {
    if (!(Number(e.units) > 0) || !(Number(e.unitPrice) > 0)) return 'Units and unit price must both be greater than 0.';
  } else if (!(Number(e.amount) > 0)) {
    return e.mode === 'deposit' ? 'The deposit must be greater than 0.' : 'The balance must be greater than 0.';
  }
  return null;
}

/** Ledger row of a validated entry. Pure. */
function valuationRow(e) {
  const p = e.date.split('-').map(Number);
  const units = e.mode === 'units';
  if (e.mode === 'deposit') {   // v1.1.40: money put into the account
    return [new Date(p[0], p[1] - 1, p[2], 12), String(e.account).trim(), 'Deposit', '', '', '', Number(e.amount), '', e.currency,
      'manual', String(e.notes || '').trim(), ''];
  }
  return [new Date(p[0], p[1] - 1, p[2], 12), String(e.account).trim(), 'Valuation', '', units ? Number(e.units) : '',
    units ? Number(e.unitPrice) : '', units ? '' : Number(e.amount), '', e.currency, 'manual', String(e.notes || '').trim(), ''];
}

function openValuationDialog() {
  const config = getConfig();
  if (!config) {
    SpreadsheetApp.getUi().alert("❌ Setup not completed. Please run Setup Wizard first.");
    return;
  }
  const ledger = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVESTMENT_LEDGER_SHEET);
  const known = [];
  if (ledger && ledger.getLastRow() > 1) {
    ledger.getRange(2, 1, ledger.getLastRow() - 1, LEDGER_HEADERS.length).getValues().forEach(r => {
      const a = String(r[LG.ACCOUNT] || '').trim();
      if (r[LG.TYPE] === 'Valuation' && a && known.indexOf(a) === -1) known.push(a);
    });
  }
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_top"><style>
  body { font-family: -apple-system, 'Segoe UI', Roboto, Arial, sans-serif; margin: 0; padding: 18px 20px; color: #1F2937; font-size: 14px; }
  label { display: block; font-weight: 600; font-size: 13px; margin: 12px 0 5px; }
  input, select { width: 100%; box-sizing: border-box; padding: 9px; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 14px; }
  .row { display: flex; gap: 10px; } .row > div { flex: 1; }
  .modes { display: flex; gap: 16px; margin-top: 4px; } .modes label { font-weight: 400; margin: 0; display: flex; gap: 6px; align-items: center; }
  .modes input { width: auto; }
  .hint { font-size: 12px; color: #6B7280; margin-top: 5px; line-height: 1.4; }
  button { margin-top: 18px; width: 100%; padding: 11px; border: 0; border-radius: 7px; background: #0F766E; color: #fff; font-size: 15px; font-weight: 600; cursor: pointer; }
  button:disabled { opacity: .6; }
  #status { display: none; margin-top: 12px; padding: 10px; border-radius: 6px; font-size: 13px; }
  #status.error { display: block; background: #FDECEC; color: #B91C1C; } #status.info { display: block; background: #EAF1FE; color: #1D4ED8; }
</style></head><body>
  <label for="account">Account</label>
  <input id="account" list="known" placeholder="e.g. Liquidity fund, Pension fund" value="">
  <datalist id="known">${known.map(a => `<option value="${esc(a)}">`).join('')}</datalist>
  <div class="hint">Use the same name every time — Holdings keeps the latest balance of each account.
    A broker with positions (e.g. HAPI): its total value on a past date sets where its return starts.</div>
  <div class="row">
    <div><label for="kind">Kind</label><select id="kind">${VALUATION_KINDS.map(k => `<option>${k}</option>`).join('')}</select></div>
    <div><label for="date" id="dateLabel">Statement date</label><input id="date" type="date" value="${today}" max="${today}"></div>
  </div>
  <label>What to record</label>
  <div class="modes">
    <label><input type="radio" name="mode" value="units" checked onchange="mode()"> Units × unit price</label>
    <label><input type="radio" name="mode" value="amount" onchange="mode()"> A balance</label>
    <label><input type="radio" name="mode" value="deposit" onchange="mode()"> A deposit</label>
  </div>
  <div id="unitsBox" class="row">
    <div><label for="units">Units</label><input id="units" type="number" step="any" min="0"></div>
    <div><label for="unitPrice">Unit price</label><input id="unitPrice" type="number" step="any" min="0"></div>
  </div>
  <div class="hint" id="unitsHint">A fund: units (cuotas) are in your statement; the unit price (valor cuota) in the fund's fact sheet.</div>
  <div id="amountBox" style="display:none"><label for="amount" id="amountLabel">Balance</label><input id="amount" type="number" step="any" min="0">
    <div class="hint" id="amountHint">A pension: the balance of your latest statement.</div></div>
  <div class="row">
    <div><label for="currency">Currency</label><select id="currency"><option>DOP</option><option>USD</option></select></div>
    <div><label for="notes">Notes</label><input id="notes" placeholder="optional"></div>
  </div>
  <button id="save" onclick="save()">Save balance</button>
  <div id="status"></div>
<script>
  function el(id) { return document.getElementById(id); }
  function pick() { return document.querySelector('input[name="mode"]:checked').value; }
  function mode() {
    const u = pick() === 'units';
    el('unitsBox').style.display = u ? 'flex' : 'none'; el('unitsHint').style.display = u ? 'block' : 'none';
    el('amountBox').style.display = u ? 'none' : 'block';
    const d = pick() === 'deposit';
    el('amountLabel').textContent = d ? 'Deposit' : 'Balance';
    el('amountHint').textContent = d ? 'Money you put into the account on that date — only deposits NOT already in the Investment Ledger ' +
      '(bank transfers that match the keyword in Investment Accounts are added by themselves).' : 'A pension: the balance of your latest statement.';
    el('save').textContent = d ? 'Save deposit' : 'Save balance';
    el('dateLabel').textContent = d ? 'Deposit date' : 'Statement date';
  }
  function fail(m) { el('status').className = 'error'; el('status').textContent = m; }
  function save() {
    const e = { account: el('account').value.trim(), kind: el('kind').value, date: el('date').value, mode: pick(),
      units: el('units').value, unitPrice: el('unitPrice').value, amount: el('amount').value,
      currency: el('currency').value, notes: el('notes').value };
    if (!e.account) return fail('Give the account a name.');
    if (e.mode === 'units' && !(Number(e.units) > 0 && Number(e.unitPrice) > 0)) return fail('Units and unit price must both be greater than 0.');
    if (e.mode !== 'units' && !(Number(e.amount) > 0)) return fail(e.mode === 'deposit' ? 'The deposit must be greater than 0.' : 'The balance must be greater than 0.');
    el('save').disabled = true;
    el('status').className = 'info';
    el('status').textContent = '⏳ Saving and updating Holdings — a summary pops up in the sheet. This window will close.';
    google.script.run.withFailureHandler(function(err) { fail('Error: ' + err); el('save').disabled = false; }).addValuationEntry(e);
    setTimeout(function() { google.script.host.close(); }, 1500);
  }
</script></body></html>`).setWidth(460).setHeight(610);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ Balance or deposit');
}

/** Saves one balance as a Valuation row (a new row every statement — history is kept), then rebuilds Holdings. */
function addValuationEntry(entry) {
  const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const problem = validateValuationEntry(entry, todayKey);
  if (problem) {
    safeAlert('❌ Balance not saved: ' + problem);
    return false;
  }
  return withRunLock(() => {
    try {
      safeToast('Saving the balance...', '📈 Investments', -1);
      const sheet = getOrCreateLedgerSheet();
      const row = valuationRow(entry);
      // v1.1.35: the same account and statement date again UPDATES that balance instead of adding a duplicate
      const last = sheet.getLastRow();
      const existing = last > 1 ? sheet.getRange(2, 1, last - 1, LEDGER_HEADERS.length).getValues() : [];
      const same = existing.findIndex(r => r[LG.TYPE] === 'Valuation' && r[LG.DATE] &&
        String(r[LG.ACCOUNT]).trim().toLowerCase() === row[LG.ACCOUNT].toLowerCase() &&
        normalizeDateForCompare(r[LG.DATE]) === entry.date);
      if (entry.mode === 'deposit') {   // v1.1.40
        const twice = existing.some(r => r[LG.TYPE] === 'Deposit' && r[LG.DATE] && normalizeDateForCompare(r[LG.DATE]) === entry.date &&
          String(r[LG.ACCOUNT]).trim().toLowerCase() === row[LG.ACCOUNT].toLowerCase() && Math.abs(invNumber(r[LG.AMOUNT]) - row[LG.AMOUNT]) < 0.005);
        if (twice) {
          safeAlert('ℹ️ Not saved: ' + row[LG.ACCOUNT] + ' already has a deposit of ' + entry.currency + ' ' + row[LG.AMOUNT] + ' on ' + entry.date + '.');
          return false;
        }
      }
      const updated = entry.mode !== 'deposit' && same !== -1;
      if (updated) {
        sheet.getRange(same + 2, 1, 1, LEDGER_HEADERS.length).setValues([row]);
      } else {
        ensureRowCapacity(sheet, last + 1);
        sheet.getRange(last + 1, 1, 1, LEDGER_HEADERS.length).setValues([row]);
      }
      sortSheetByDateDesc(sheet, LG.DATE + 1);
      // the account is listed in Investment Accounts (no deposit keyword needed for a balance)
      readInvestmentAccounts();
      const accounts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVESTMENT_ACCOUNTS_SHEET);
      const names = accounts.getLastRow() > 1 ? accounts.getRange(2, 1, accounts.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase()) : [];
      if (names.indexOf(row[LG.ACCOUNT].toLowerCase()) === -1) {
        accounts.getRange(accounts.getLastRow() + 1, 1, 1, ACCOUNTS_HEADERS.length)
          .setValues([[row[LG.ACCOUNT], VALUATION_KINDS.indexOf(entry.kind) !== -1 ? entry.kind : 'Other', '', 'Balance from statements (Valuation rows)']]);
      }
      safeToast('Updating Holdings...', '📈 Investments', -1);
      refreshHoldings();
      formatDataSheets();
      ensureSheetOrder();
      safeToast('Done.', '📈 Investments', 3);
      const value = entry.mode === 'units' ? Number(entry.units) * Number(entry.unitPrice) : Number(entry.amount);
      if (entry.mode === 'deposit') {
        safeAlert('✅ Deposit saved\n\n' + row[LG.ACCOUNT] + ' — ' + entry.currency + ' ' + Number(entry.amount).toLocaleString('en-US',
          { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' on ' + entry.date + '\n\nIt counts as money put in, not as gain (Holdings › Performance).');
        return true;
      }
      safeAlert((updated ? '✅ Balance updated (same account and date)\n\n' : '✅ Balance saved\n\n') + row[LG.ACCOUNT] + ' — ' + entry.currency + ' ' +
        value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' as of ' + entry.date +
        (entry.mode === 'units' ? ' (' + entry.units + ' × ' + entry.unitPrice + ')' : '') +
        (isPositionAccount(row[LG.ACCOUNT]) ? '\n\nThis account has positions: the balance sets where its return starts (Holdings › Performance).'
          : '\n\nIt now shows in Holdings › Other accounts. Add a new balance with every statement.'));
      return true;
    } catch (error) {
      safeAlert('❌ Could not save the balance: ' + error);
      return false;
    }
  });
}

/** v1.1.39: whether the ledger holds positions (Snapshot, Buy or Sell rows) for this account. */
function isPositionAccount(account) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVESTMENT_LEDGER_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return false;
  const name = String(account).trim().toLowerCase();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, LEDGER_HEADERS.length).getValues()
    .some(r => String(r[LG.ACCOUNT]).trim().toLowerCase() === name && ['Snapshot', 'Buy', 'Sell'].indexOf(r[LG.TYPE]) !== -1);
}

/** v1.1.43: the investment tabs only (after the investments step; the full styling runs with recategorize). */
function styleInvestmentSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [[INVESTMENT_LEDGER_SHEET, styleLedgerSheet], [HISTORY_SHEET, styleHistorySheet], [INVESTMENT_ACCOUNTS_SHEET, styleAccountsSheet]]
    .forEach(([name, style]) => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) return;
      try { style(sheet); } catch (error) { Logger.log('Could not style ' + name + ': ' + error); }
    });
}

/* ======================================================================
 * PASTE A BROKER'S PORTFOLIO — v1.1.48
 * The positions screen (HAPI's "My Assets") copied from the browser and pasted in a dialog becomes a Snapshot:
 * per ticker, quantity; value; gain — cost = value − gain, price = value ÷ quantity. "Total assets" in the paste is
 * checked against the positions' sum (a position left out of the copy shows up), "Total money" becomes the CASH row.
 * ====================================================================== */
const PASTE_TOTAL_TOLERANCE = 0.1;   // US$: values on the screen are rounded to cents, one rounding per position

/** Text as copied from the broker's screen → { positions, cash, totalAssets, warnings }. Pure. */
function parsePortfolioPaste(text) {
  const lines = String(text || '').split(/\r?\n/)
    .map(l => l.replace(/\[([^\]]*)\]\([^)]*\)/g, '$1').replace(/[\u00a0\u200b]/g, ' ').trim())   // [text](link) → text
    .filter(Boolean);
  const money = s => { const m = String(s).match(/^([+\-−]?)\s*(?:US)?\$\s*([\d,]*\.?\d+)/); return m ? (m[1] && m[1] !== '+' ? -1 : 1) * Number(m[2].replace(/,/g, '')) : null; };
  const number = s => /^[\d,]*\.?\d+$/.test(s) ? Number(s.replace(/,/g, '')) : null;
  const isTicker = s => /^[A-Z][A-Z0-9.\-]{0,11}$/.test(s);
  const out = { positions: [], cash: null, totalAssets: null, warnings: [] };
  const labelled = (label, i) => {   // "Total money $7.77" or "Total money" / "$7.77" on the next line
    const same = lines[i].slice(label.length).trim();
    if (money(same) !== null) return money(same);
    return i + 1 < lines.length && money(lines[i + 1]) !== null ? money(lines[i + 1]) : null;
  };
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (/^total money/i.test(line)) { out.cash = labelled('Total money', i); continue; }
    if (/^total assets/i.test(line)) { out.totalAssets = labelled('Total assets', i); continue; }
    if (!isTicker(line)) continue;
    const qty = number(lines[i + 1] || ''), value = money(lines[i + 2] || ''), gain = money(lines[i + 3] || '');
    if (qty === null || value === null || gain === null) continue;   // a word in capitals, not a position
    if (!(qty > 0)) { out.warnings.push(line + ': quantity is 0 — skipped'); continue; }
    if (out.positions.some(p => p.ticker === line)) { out.warnings.push(line + ' appears twice — the first one is kept'); i += 3; continue; }
    const cost = +(value - gain).toFixed(2);
    out.positions.push({ ticker: line, qty: qty, value: value, gain: gain, cost: cost, price: Number((value / qty).toPrecision(10)) });
    i += 3;
  }
  if (!out.positions.length) out.warnings.push('No positions found. Copy the list of assets from the broker\'s portfolio screen — ticker, quantity, value and gain for each.');
  const sum = out.positions.reduce((t, p) => t + p.value, 0);
  out.positionsValue = +sum.toFixed(2);
  if (out.totalAssets !== null && Math.abs(out.totalAssets - sum) > PASTE_TOTAL_TOLERANCE) {
    out.warnings.push('The positions add up to US$' + sum.toFixed(2) + ' but the screen says US$' + out.totalAssets.toFixed(2) +
      ' — US$' + Math.abs(out.totalAssets - sum).toFixed(2) + (out.totalAssets > sum ? ' is missing: a position was probably left out of the copy (scroll down the list and copy again).'
        : ' too much: a position may be pasted twice.'));
  }
  return out;
}

/** Rows of the Snapshot to save. Pure. */
function snapshotRows(parsed, account, dateKey) {
  const p = dateKey.split('-').map(Number), when = new Date(p[0], p[1] - 1, p[2], 12);
  const rows = parsed.positions.map(x => [when, account, 'Snapshot', x.ticker, x.qty, x.price, x.cost, '', 'USD', 'manual', 'pasted from the broker\'s screen', '']);
  if (parsed.cash !== null) rows.push([when, account, 'Snapshot', 'CASH', '', '', parsed.cash, '', 'USD', 'manual', 'pasted from the broker\'s screen', '']);
  return rows;
}

/** From the dialog: what the paste would save (nothing is written). */
function previewPortfolioPaste(text) {
  return parsePortfolioPaste(text);
}

/** From the dialog: saves the Snapshot — replacing that account's snapshot of the same day — and rebuilds Holdings. */
function savePortfolioPaste(entry) {
  const account = String(entry && entry.account || '').trim();
  const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (!account) { safeAlert('❌ Not saved: give the account a name.'); return false; }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(entry.date || '')) || entry.date > todayKey) { safeAlert('❌ Not saved: pick the date the screen shows (today or earlier).'); return false; }
  const parsed = parsePortfolioPaste(entry.text);
  if (!parsed.positions.length) { safeAlert('❌ Not saved: ' + parsed.warnings.join(' ')); return false; }
  return withRunLock(() => {
    try {
      safeToast('Saving the snapshot...', '📈 Investments', -1);
      const sheet = getOrCreateLedgerSheet();
      const last = sheet.getLastRow();
      const width = LEDGER_HEADERS.length;
      const existing = last > 1 ? sheet.getRange(2, 1, last - 1, width).getValues() : [];
      const same = r => r[LG.TYPE] === 'Snapshot' && r[LG.DATE] && normalizeDateForCompare(r[LG.DATE]) === entry.date &&
        String(r[LG.ACCOUNT]).trim().toLowerCase() === account.toLowerCase();
      const kept = existing.filter(r => !same(r));
      const replaced = existing.length - kept.length;
      const all = kept.concat(snapshotRows(parsed, account, entry.date));
      if (last > 1) sheet.getRange(2, 1, last - 1, width).clearContent();
      ensureRowCapacity(sheet, all.length + 1);
      sheet.getRange(2, 1, all.length, width).setValues(all);
      sortSheetByDateDesc(sheet, LG.DATE + 1);
      readInvestmentAccounts();
      refreshHoldings();
      formatDataSheets();
      ensureSheetOrder();
      safeToast('Done.', '📈 Investments', 3);
      safeAlert('✅ Snapshot saved — ' + account + ', ' + entry.date + '\n\n' + parsed.positions.length + ' position(s), US$' +
        parsed.positionsValue.toFixed(2) + (parsed.cash !== null ? ' + cash US$' + parsed.cash.toFixed(2) : '') +
        (replaced ? '\nIt replaced the ' + replaced + ' row(s) of that day\'s earlier snapshot.' : '') +
        (parsed.warnings.length ? '\n\n⚠️ ' + parsed.warnings.join('\n⚠️ ') : '') +
        '\n\nPositions now start from this snapshot; movements after that day are added from the broker\'s emails.');
      return true;
    } catch (error) {
      safeAlert('❌ Could not save the snapshot: ' + error);
      return false;
    }
  });
}

function openPortfolioPasteDialog() {
  const ledger = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVESTMENT_LEDGER_SHEET);
  const known = ['HAPI'];
  if (ledger && ledger.getLastRow() > 1) {
    ledger.getRange(2, 1, ledger.getLastRow() - 1, LEDGER_HEADERS.length).getValues().forEach(r => {
      const a = String(r[LG.ACCOUNT] || '').trim();
      if (a && ['Snapshot', 'Buy', 'Sell'].indexOf(r[LG.TYPE]) !== -1 && known.indexOf(a) === -1) known.push(a);
    });
  }
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_top"><style>
  body { font-family: -apple-system, 'Segoe UI', Roboto, Arial, sans-serif; margin: 0; padding: 16px 18px; color: #1F2937; font-size: 13px; }
  label { display: block; font-weight: 600; margin: 10px 0 4px; }
  input, textarea { width: 100%; box-sizing: border-box; padding: 8px; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 13px; }
  textarea { height: 150px; font-family: ui-monospace, Menlo, Consolas, monospace; font-size: 12px; }
  .row { display: flex; gap: 10px; } .row > div { flex: 1; }
  .hint { font-size: 12px; color: #6B7280; margin-top: 4px; line-height: 1.4; }
  table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 12px; }
  th, td { padding: 4px 6px; border-bottom: 1px solid #EEF1F5; text-align: right; } th:first-child, td:first-child { text-align: left; }
  th { color: #6B7280; font-weight: 600; font-size: 11px; }
  .warn { background: #FFF4D6; color: #92400E; padding: 8px; border-radius: 6px; margin-top: 8px; font-size: 12px; }
  .ok { background: #E8F5E9; color: #2E7D32; padding: 8px; border-radius: 6px; margin-top: 8px; font-size: 12px; }
  .buttons { display: flex; gap: 8px; margin-top: 12px; }
  button { flex: 1; padding: 10px; border-radius: 7px; font-size: 14px; font-weight: 600; cursor: pointer; border: 1px solid #0F766E; }
  #preview { background: #fff; color: #0F766E; } #save { background: #0F766E; color: #fff; } button:disabled { opacity: .5; cursor: default; }
  #result { max-height: 200px; overflow-y: auto; }
</style></head><body>
  <div class="row">
    <div><label for="account">Account</label><input id="account" list="known" value="${esc(known[0])}">
      <datalist id="known">${known.map(a => '<option value="' + esc(a) + '">').join('')}</datalist></div>
    <div><label for="date">Date on the screen</label><input id="date" type="date" value="${today}" max="${today}"></div>
  </div>
  <label for="text">The broker's portfolio screen, pasted</label>
  <textarea id="text" placeholder="Total assets&#10;$0.00&#10;Total money&#10;$0.00&#10;GOOGL&#10;1.5&#10;$450.00&#10;+$50.00 (+12.50%)&#10;..."></textarea>
  <div class="hint">In the broker's app on the web, open the portfolio, select from <b>Total balance</b> down to the last asset
    (scroll to the end of the list), copy and paste here. Each asset needs its ticker, quantity, value and gain.
    Including <b>Total assets</b> lets the tracker check nothing was left out; <b>Total money</b> becomes the cash.</div>
  <div class="buttons"><button id="preview" onclick="preview()">Preview</button><button id="save" onclick="save()" disabled>Save snapshot</button></div>
  <div id="result"></div>
<script>
  function el(id) { return document.getElementById(id); }
  function money(n) { return 'US$' + Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;'); }
  function show(r) {
    var h = '';
    if (r.positions.length) {
      h += '<table><tr><th>Ticker</th><th>Quantity</th><th>Value</th><th>Cost</th></tr>';
      r.positions.forEach(function (p) { h += '<tr><td>' + esc(p.ticker) + '</td><td>' + p.qty + '</td><td>' + money(p.value) + '</td><td>' + money(p.cost) + '</td></tr>'; });
      h += '</table>';
      h += '<div class="' + (r.warnings.length ? 'warn' : 'ok') + '">' + r.positions.length + ' position(s), ' + money(r.positionsValue) +
        (r.cash !== null ? ' + cash ' + money(r.cash) : ' · no Total money in the paste: no cash row') +
        (r.totalAssets !== null && !r.warnings.length ? ' · matches Total assets ✓' : '') + '</div>';
    }
    r.warnings.forEach(function (w) { h += '<div class="warn">⚠️ ' + esc(w) + '</div>'; });
    el('result').innerHTML = h;
    el('save').disabled = !r.positions.length;
    el('save').textContent = r.warnings.length ? 'Save anyway' : 'Save snapshot';   // a known gap is saved on purpose
  }
  function preview() {
    el('save').disabled = true;
    el('result').innerHTML = '<div class="hint">Reading…</div>';
    google.script.run.withSuccessHandler(show).withFailureHandler(function (e) { el('result').innerHTML = '<div class="warn">' + esc(e) + '</div>'; })
      .previewPortfolioPaste(el('text').value);
  }
  function save() {
    el('save').disabled = true; el('preview').disabled = true;
    el('result').innerHTML = '<div class="ok">⏳ Saving and updating Holdings — a summary pops up in the sheet. This window will close.</div>';
    google.script.run.withFailureHandler(function (e) { el('result').innerHTML = '<div class="warn">' + esc(e) + '</div>'; el('preview').disabled = false; })
      .savePortfolioPaste({ account: el('account').value.trim(), date: el('date').value, text: el('text').value });
    setTimeout(function () { google.script.host.close(); }, 1500);
  }
  el('text').addEventListener('input', function () { el('save').disabled = true; });
</script></body></html>`).setWidth(520).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Paste broker positions');
}

/**
 * v1.1.56: waits for GOOGLEFINANCE after Holdings is written — checks each position's Price source until every one is
 * live (stocks "live", crypto "Coinbase · …"), stopping early after LIVE_PRICE_CALM_POLLS checks without progress (a
 * symbol Google Finance doesn't know never loads) and at LIVE_PRICE_WAIT_MS. Returns which positions are live.
 */
function waitForLivePrices(sheet, firstRow, count, opts) {
  opts = opts || {};
  const clock = opts.clock || (() => Date.now());
  const sleep = opts.sleep || (ms => Utilities.sleep(ms));
  const maxMs = opts.maxMs || LIVE_PRICE_WAIT_MS, started = clock();
  const isLive = s => s === 'live' || /^Coinbase/.test(s);
  const read = () => sheet.getRange(firstRow, 13, count, 1).getValues().map(r => String(r[0]));
  let sources = read(), pending = sources.filter(s => !isLive(s)).length, calm = 0;
  while (pending && calm < LIVE_PRICE_CALM_POLLS && clock() - started < maxMs) {
    sleep(1500);
    SpreadsheetApp.flush();
    sources = read();
    const now = sources.filter(s => !isLive(s)).length;
    calm = now < pending ? 0 : calm + 1;
    pending = now;
  }
  if (pending) Logger.log('Holdings: ' + pending + ' position(s) without a live price after ' + Math.round((clock() - started) / 1000) + ' s');
  return sources.map(isLive);
}

/** The latest price the history has for each account|ticker, from days before `beforeKey`. */
function latestHistoryPrices(historyValues, beforeKey) {
  const best = {};
  historyValues.slice(1).forEach(r => {
    if (!r[0] || !(invNumber(r[4]) > 0)) return;
    const day = normalizeDateForCompare(r[0]), key = r[1] + '|' + r[2];
    if (day >= beforeKey) return;
    if (!best[key] || best[key].day < day) best[key] = { day: day, price: invNumber(r[4]) };
  });
  const out = {};
  Object.keys(best).forEach(k => { out[k] = best[k].price; });
  return out;
}

// ====================================================================================================
// 08_startHere.gs
// ====================================================================================================

/**
 * START HERE — v1.1.47
 *
 * A sidebar that walks through setting the tracker up and keeping it healthy. Every step is CHECKED against the sheet
 * and the project, never ticked by hand, so the same list is the health check: a step that breaks later (the daily
 * update unscheduled, emails waiting in Unrecognized, a ledger row without a date) shows up again with what to do.
 *
 * startHereStatus() is what the sidebar renders; it returns plain data (no Date objects), so the sidebar can refresh it.
 */
const START_HERE_STALE_BALANCE_DAYS = 45;
const START_HERE_STALE_RUN_HOURS = 36;
// the only functions the sidebar's buttons may call, and the only tabs it may open
const START_HERE_ACTIONS = ['openSetupWizard', 'openDateRangeDialog', 'openValuationDialog', 'refreshInvestmentsNow', 'openSheetByName',
  'openPortfolioPasteDialog'];
const START_HERE_SHEETS = ['Unrecognized', 'Bank Transfers', 'Incoming Transfers', 'Investment Ledger', 'Holdings', 'Custom Rules', 'Dashboard'];

function openStartHere() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(startHereHtml()).setTitle('📘 Start here'));
}

function openSheetByName(name) {
  if (START_HERE_SHEETS.indexOf(name) === -1) return false;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) return false;
  if (sheet.isSheetHidden()) sheet.showSheet();
  SpreadsheetApp.setActiveSheet(sheet);
  return true;
}

/** Every step, checked. Each: { group, id, title, status: done|todo|warn|error|optional, detail, action: {label, fn, arg} }. */
function startHereStatus(now) {
  now = now || new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const steps = [];
  const add = (group, id, title, status, detail, action) => steps.push({ group: group, id: id, title: title, status: status,
    detail: detail, action: action || null });
  const day = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const open = (label, sheet) => ({ label: label, fn: 'openSheetByName', arg: sheet });

  // ---- Setup
  let config = null;
  try { config = getConfig(); } catch (error) { config = null; }
  const banks = config && config.banksToTrack ? Object.keys(config.banksToTrack).filter(b => config.banksToTrack[b]) : [];
  const configured = !!(config && config.email && banks.length);
  add('Setup', 'code', 'Tracker v' + SCRIPT_VERSION + ' is installed', 'done',
    'One file, FinancialTracker.gs. To update: select everything in it, paste the new version, save.');
  add('Setup', 'settings', 'Your settings', configured ? 'done' : 'todo',
    configured ? 'Banks: ' + banks.join(', ') + '. Change income, banks, cards or emails in the Setup Wizard.'
      : 'Income, banks, cards and summary emails — the Setup Wizard asks for all of it.',
    { label: configured ? 'Change settings' : 'Open the Setup Wizard', fn: 'openSetupWizard' });
  let handlers = [];
  try { handlers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction()); } catch (error) { handlers = []; }
  const scheduled = handlers.indexOf('runGmailMonitor') !== -1;
  add('Setup', 'schedule', 'Daily update scheduled', !configured ? 'todo' : scheduled ? 'done' : 'error',
    scheduled ? 'New bank emails are read every morning at 6 AM.'
      : configured ? 'Nothing is scheduled, so nothing updates by itself. Saving the Setup Wizard schedules it.'
        : 'Scheduled when you save the Setup Wizard.',
    scheduled ? null : { label: 'Open the Setup Wizard', fn: 'openSetupWizard' });
  if (configured && config.notifyEnabled) {
    const ok = handlers.indexOf('sendDailySummary') !== -1;
    add('Setup', 'dailyEmail', 'Daily summary email', ok ? 'done' : 'error',
      ok ? 'Sent every morning around ' + config.notifyHour + ':00.' : 'Turned on but not scheduled — save the Setup Wizard again.',
      ok ? null : { label: 'Open the Setup Wizard', fn: 'openSetupWizard' });
  }
  if (configured && config.notifyMonthly) {
    const ok = handlers.indexOf('sendMonthlySummary') !== -1;
    add('Setup', 'monthlyEmail', 'Monthly summary email', ok ? 'done' : 'error',
      ok ? 'Sent on the 1st of every month.' : 'Turned on but not scheduled — save the Setup Wizard again.',
      ok ? null : { label: 'Open the Setup Wizard', fn: 'openSetupWizard' });
  }

  // ---- Your data
  let last = null;
  try { last = readLastRun(); } catch (error) { last = null; }
  if (!last || !last.at) {
    add('Your data', 'firstRun', "Read this year's bank emails", configured ? 'todo' : 'optional',
      'Run Monitor by Date Range from January 1 to today. Repeat it until its summary has no ⏸ or ⏭ line — each run gets further.',
      configured ? { label: 'Monitor by Date Range', fn: 'openDateRangeDialog' } : null);
  } else {
    const at = new Date(last.at), hours = (now - at) / 3600000;
    const bits = ['Last update: ' + Utilities.formatDate(at, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + ' · ' + (last.saved || 0) + ' saved.'];
    if (last.partial) bits.push('It stopped early to stay within the time limit — run it again to finish.');
    if (last.errors) bits.push(last.errors + ' error(s): see Extensions › Apps Script › Executions.');
    if (hours > START_HERE_STALE_RUN_HOURS) bits.push('That is more than a day ago — is the daily update running?');
    add('Your data', 'lastRun', 'Bank emails up to date', last.partial || last.errors || hours > START_HERE_STALE_RUN_HOURS ? 'warn' : 'done',
      bits.join(' '), last.partial ? { label: 'Monitor by Date Range', fn: 'openDateRangeDialog' } : null);
  }
  const unrec = ss.getSheetByName(UNRECOGNIZED_SHEET);
  const waiting = unrec && unrec.getLastRow() > 1
    ? unrec.getRange(2, 7, unrec.getLastRow() - 1, 1).getValues().filter(r => r[0] === 'New').length : 0;
  add('Your data', 'unrecognized', waiting ? waiting + ' email(s) could not be read' : 'Every email was read', waiting ? 'warn' : 'done',
    waiting ? 'Each one says why and links to the email. Send it to get its format supported, or set its Status to Ignore.'
      : 'Emails the tracker cannot read are listed in Unrecognized.', waiting ? open('Open Unrecognized', 'Unrecognized') : null);
  const tx = ss.getSheetByName(TRANSACTIONS_SHEET);
  if (tx && tx.getLastRow() > 1) {
    const month = day(now).slice(0, 7);
    const open$ = tx.getRange(2, 1, tx.getLastRow() - 1, TX_NUM_COLS).getValues().filter(r => r[TX_COL.DATE] &&
      r[TX_COL.TYPE] === 'Transfer' && !String(r[TX_COL.CATEGORY] || '').trim() && normalizeDateForCompare(r[TX_COL.DATE]).slice(0, 7) === month).length;
    add('Your data', 'transfers', open$ ? open$ + ' transfer(s) this month without a category' : 'Transfers categorized',
      open$ ? 'warn' : 'done', open$ ? 'Type a category in Bank Transfers (Exclude for moves between your own accounts), or add a Custom Rule for ones that repeat.'
        : 'Every transfer this month has a category.', open$ ? open('Open Bank Transfers', 'Bank Transfers') : null);
  }

  // v1.1.51: money received — only counted once it has a category
  if (tx && tx.getLastRow() > 1) {
    const received = tx.getRange(2, 1, tx.getLastRow() - 1, TX_NUM_COLS).getValues()
      .filter(r => r[TX_COL.DATE] && r[TX_COL.TYPE] === 'Incoming' && !String(r[TX_COL.CATEGORY] || '').trim()).length;
    if (received) {
      add('Your data', 'incoming', received + ' incoming transfer(s) without a category', 'warn',
        'Give each one the category it pays back (a roommate\'s share of the rent → Rent) and it is taken off what you spent there; ' +
        'Exclude for money that pays nothing back. A Custom Rule on the sender\'s name does it by itself.',
        open('Open Incoming Transfers', 'Incoming Transfers'));
    }
  }

  // ---- Investments (optional)
  const ledger = ss.getSheetByName(INVESTMENT_LEDGER_SHEET);
  const values = ledger ? ledger.getDataRange().getValues() : [];
  if (!ledger || values.length < 2) {
    add('Investments (optional)', 'investments', 'Track your investments', 'optional',
      'Broker orders and dividends are read from their emails (HAPI today). Start by pasting your positions from the broker\'s screen.',
      { label: 'Paste broker positions', fn: 'openPortfolioPasteDialog' });
  } else {
    const rows = values.slice(1).filter(r => String(r[LG.ACCOUNT] || '').trim());
    const accounts = [...new Set(rows.map(r => String(r[LG.ACCOUNT]).trim()))];
    const positionAccounts = accounts.filter(a => rows.some(r => String(r[LG.ACCOUNT]).trim() === a && ['Snapshot', 'Buy', 'Sell'].indexOf(r[LG.TYPE]) !== -1));
    const hasSnapshot = a => rows.some(r => String(r[LG.ACCOUNT]).trim() === a && r[LG.TYPE] === 'Snapshot');
    if (!positionAccounts.length) {
      add('Investments (optional)', 'snapshot', "Your broker's positions", 'todo',
        "Paste the broker's portfolio screen — the tracker turns it into a Snapshot (one row per ticker and the cash).",
        { label: 'Paste broker positions', fn: 'openPortfolioPasteDialog' });
    }
    positionAccounts.forEach(a => {
      if (!hasSnapshot(a)) {
        add('Investments (optional)', 'snapshot-' + a, a + ': positions', 'todo',
          "Orders are recorded, but positions need a Snapshot to start from — paste the broker's portfolio screen.",
          { label: 'Paste broker positions', fn: 'openPortfolioPasteDialog' });
        return;
      }
      const firstSnap = rows.filter(r => String(r[LG.ACCOUNT]).trim() === a && r[LG.TYPE] === 'Snapshot' && r[LG.DATE])
        .map(r => normalizeDateForCompare(r[LG.DATE])).sort()[0];
      const start = rows.filter(r => String(r[LG.ACCOUNT]).trim() === a && r[LG.TYPE] === 'Valuation' && r[LG.DATE])
        .map(r => normalizeDateForCompare(r[LG.DATE])).sort()[0];
      const fromStart = start && firstSnap && start < firstSnap;
      add('Investments (optional)', 'start-' + a, a + ': return measured from ' + (fromStart ? start : firstSnap || '—'), fromStart ? 'done' : 'todo',
        fromStart ? "From its value on that date, with every deposit since counted as money put in."
          : "It starts at the snapshot. To measure the year, add the account's value on its opening date (from the broker's statement) as a balance.",
        fromStart ? null : { label: 'Add a balance', fn: 'openValuationDialog' });
    });
    const undated = rows.filter(r => !r[LG.DATE] && r[LG.TYPE]).length;
    if (undated) {
      add('Investments (optional)', 'undated', undated + ' ledger row(s) without a date', 'error',
        "They aren't counted anywhere until they have one (highlighted in red in the Investment Ledger).",
        open('Open the Investment Ledger', 'Investment Ledger'));
    }
    let noAmount = 0;
    try { noAmount = checkDepositNotices(values).missing.length; } catch (error) { noAmount = 0; }
    if (noAmount) {
      add('Investments (optional)', 'deposits', noAmount + ' deposit(s) without an amount', 'warn',
        'The broker confirmed them but its email has no amount — add each one as a deposit, or it counts as gain.',
        open('Open Unrecognized', 'Unrecognized'));
    }
    accounts.filter(a => positionAccounts.indexOf(a) === -1).forEach(a => {
      const latest = rows.filter(r => String(r[LG.ACCOUNT]).trim() === a && r[LG.TYPE] === 'Valuation' && r[LG.DATE])
        .map(r => normalizeDateForCompare(r[LG.DATE])).sort().pop();
      if (!latest) return;
      const age = daysBetween(latest, day(now));
      add('Investments (optional)', 'balance-' + a, a + ': balance of ' + latest, age > START_HERE_STALE_BALANCE_DAYS ? 'warn' : 'done',
        age > START_HERE_STALE_BALANCE_DAYS ? age + ' days ago — add the latest statement.' : 'Recent enough. Add each new statement when it arrives.',
        age > START_HERE_STALE_BALANCE_DAYS ? { label: 'Add a balance', fn: 'openValuationDialog' } : null);
    });
  }

  const counted = steps.filter(s => s.status !== 'optional');
  return { version: SCRIPT_VERSION, checkedAt: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    steps: steps, done: counted.filter(s => s.status === 'done').length, total: counted.length,
    problems: steps.filter(s => s.status === 'error').length };
}

function startHereHtml() {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_top"><style>
  body { font-family: -apple-system, 'Segoe UI', Roboto, Arial, sans-serif; margin: 0; padding: 14px 14px 24px; color: #1F2937; font-size: 13px; }
  h1 { font-size: 16px; margin: 0 0 4px; color: #1F3864; }
  .sub { color: #6B7280; font-size: 12px; margin-bottom: 10px; }
  .bar { height: 6px; background: #E5E7EB; border-radius: 3px; overflow: hidden; margin: 6px 0 4px; }
  .bar div { height: 100%; background: #0F766E; }
  h2 { font-size: 11px; letter-spacing: .06em; text-transform: uppercase; color: #6B7280; margin: 16px 0 6px; }
  .step { display: flex; gap: 9px; padding: 9px 0; border-bottom: 1px solid #F1F3F6; }
  .icon { width: 20px; flex: none; font-size: 15px; line-height: 18px; text-align: center; }
  .title { font-weight: 600; color: #111827; }
  .detail { color: #4B5563; margin-top: 2px; line-height: 1.4; }
  .error .title { color: #B91C1C; } .warn .title { color: #92400E; }
  button { margin-top: 6px; padding: 5px 10px; border: 1px solid #0F766E; background: #fff; color: #0F766E; border-radius: 5px;
    font-size: 12px; cursor: pointer; } button:hover { background: #E0F2F1; }
  .top { display: flex; justify-content: space-between; align-items: center; }
  .refresh { margin: 0; border-color: #D1D5DB; color: #374151; }
  .foot { margin-top: 18px; color: #6B7280; font-size: 12px; line-height: 1.5; }
  a { color: #1D4ED8; }
</style></head><body>
  <div class="top"><h1>📘 Start here</h1><button class="refresh" onclick="load()">↻ Check again</button></div>
  <div id="summary" class="sub">Checking your tracker…</div>
  <div class="bar"><div id="fill" style="width:0"></div></div>
  <div id="steps"></div>
  <div class="foot">Every step is checked against your sheet — nothing to tick by hand. Something breaks later? It shows up
    here again. The full guide: <a href="https://github.com/joaquinganan/agentic-fin-tracker/blob/main/docs/USER_GUIDE.md" target="_blank">USER_GUIDE</a>
    · <a href="https://github.com/joaquinganan/agentic-fin-tracker/blob/main/docs/GUIA.md" target="_blank">Guía en español</a>.</div>
<script>
  var ICONS = { done: '✅', todo: '⬜', warn: '⚠️', error: '❌', optional: '○' };
  var current = [];
  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function render(r) {
    current = r.steps;
    document.getElementById('summary').textContent = r.done + ' of ' + r.total + ' done' +
      (r.problems ? ' · ' + r.problems + ' problem(s) to fix' : '') + ' · v' + r.version + ' · checked ' + r.checkedAt;
    document.getElementById('fill').style.width = (r.total ? Math.round(100 * r.done / r.total) : 0) + '%';
    var html = '', group = '';
    r.steps.forEach(function (s, i) {
      if (s.group !== group) { group = s.group; html += '<h2>' + esc(group) + '</h2>'; }
      html += '<div class="step ' + s.status + '"><div class="icon">' + ICONS[s.status] + '</div><div><div class="title">' + esc(s.title) +
        '</div><div class="detail">' + esc(s.detail) + '</div>' +
        (s.action ? '<button onclick="act(' + i + ')">' + esc(s.action.label) + '</button>' : '') + '</div></div>';
    });
    document.getElementById('steps').innerHTML = html;
  }
  function fail(e) { document.getElementById('summary').textContent = 'Could not check: ' + e; }
  function load() {
    document.getElementById('summary').textContent = 'Checking your tracker…';
    google.script.run.withSuccessHandler(render).withFailureHandler(fail).startHereStatus();
  }
  function act(i) {
    var a = current[i] && current[i].action;
    if (!a) return;
    var run = google.script.run.withSuccessHandler(function () { setTimeout(load, 800); }).withFailureHandler(fail);
    if (a.fn === 'openSetupWizard') run.openSetupWizard();
    else if (a.fn === 'openDateRangeDialog') run.openDateRangeDialog();
    else if (a.fn === 'openValuationDialog') run.openValuationDialog();
    else if (a.fn === 'refreshInvestmentsNow') run.refreshInvestmentsNow();
    else if (a.fn === 'openSheetByName') run.openSheetByName(a.arg);
    else if (a.fn === 'openPortfolioPasteDialog') run.openPortfolioPasteDialog();
  }
  load();
</script></body></html>`;
}

// ====================================================================================================
// 09_pdfText.gs
// ====================================================================================================

/**
 * PDF TEXT — v1.1.52
 *
 * Reads the text of a bank statement PDF inside Apps Script, with no Drive API: v1.1.51 converted the PDF to a
 * Google Doc through Drive, and on real statements that conversion returned text with no totals and no rows (every
 * statement ended in Unrecognized). This reads the PDF itself: finds each page's content, inflates it (FlateDecode),
 * follows the text operators with their positions and rebuilds the lines — the way `pdftotext` does.
 *
 * Scope: what bank statements use — PDF 1.4-style files (objects not packed in object streams), FlateDecode content,
 * simple fonts (WinAnsi, /Differences with glyph names) and Type0 fonts with a ToUnicode map. Anything else throws, and
 * the statement goes to Unrecognized with the reason.
 */

/* ---------- inflate (RFC 1951), for FlateDecode streams ---------- */
const PDF_LENGTH_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const PDF_LENGTH_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const PDF_DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const PDF_DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const PDF_CLEN_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

/** zlib data (2-byte header) or raw deflate → bytes. Pure. */
function pdfInflate(src) {
  let pos = 0, bit = 0, bitCount = 0;
  if (src.length > 2 && (src[0] & 0x0f) === 8 && ((src[0] << 8) | src[1]) % 31 === 0) pos = 2;   // zlib header
  let out = new Uint8Array(Math.max(1024, src.length * 4)), outLen = 0;
  const put = b => {
    if (outLen === out.length) { const bigger = new Uint8Array(out.length * 2); bigger.set(out); out = bigger; }
    out[outLen++] = b;
  };
  const need = n => { while (bitCount < n) { if (pos >= src.length) throw new Error('PDF: compressed data ends early'); bit |= src[pos++] << bitCount; bitCount += 8; } };
  const bits = n => { if (!n) return 0; need(n); const v = bit & ((1 << n) - 1); bit >>>= n; bitCount -= n; return v; };
  // a Huffman table: counts per length and symbols in canonical order
  const build = lengths => {
    const counts = new Uint16Array(16), offs = new Uint16Array(16), symbols = new Uint16Array(lengths.length);
    lengths.forEach(l => counts[l]++);
    counts[0] = 0;
    for (let i = 1; i < 16; i++) offs[i] = offs[i - 1] + counts[i - 1];
    lengths.forEach((l, s) => { if (l) symbols[offs[l]++] = s; });
    return { counts: counts, symbols: symbols };
  };
  const decode = t => {
    let code = 0, first = 0, index = 0;
    for (let len = 1; len < 16; len++) {
      code |= bits(1);
      const count = t.counts[len];
      if (code - count < first) return t.symbols[index + (code - first)];
      index += count; first += count; first <<= 1; code <<= 1;
    }
    throw new Error('PDF: bad compressed data');
  };
  let fixedLit = null, fixedDist = null;
  let final = 0;
  while (!final) {
    final = bits(1);
    const type = bits(2);
    if (type === 0) {                                   // stored
      bit = 0; bitCount = 0;
      const len = src[pos] | (src[pos + 1] << 8); pos += 4;
      for (let i = 0; i < len; i++) put(src[pos++]);
      continue;
    }
    let lit, dist;
    if (type === 1) {                                   // fixed Huffman
      if (!fixedLit) {
        const l = []; for (let i = 0; i < 288; i++) l.push(i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8);
        fixedLit = build(l); fixedDist = build(new Array(30).fill(5));
      }
      lit = fixedLit; dist = fixedDist;
    } else if (type === 2) {                            // dynamic Huffman
      const hlit = bits(5) + 257, hdist = bits(5) + 1, hclen = bits(4) + 4;
      const clen = new Array(19).fill(0);
      for (let i = 0; i < hclen; i++) clen[PDF_CLEN_ORDER[i]] = bits(3);
      const ct = build(clen), lengths = [];
      while (lengths.length < hlit + hdist) {
        const sym = decode(ct);
        if (sym < 16) lengths.push(sym);
        else if (sym === 16) { const prev = lengths[lengths.length - 1], n = bits(2) + 3; for (let i = 0; i < n; i++) lengths.push(prev); }
        else if (sym === 17) { const n = bits(3) + 3; for (let i = 0; i < n; i++) lengths.push(0); }
        else { const n = bits(7) + 11; for (let i = 0; i < n; i++) lengths.push(0); }
      }
      lit = build(lengths.slice(0, hlit)); dist = build(lengths.slice(hlit));
    } else {
      throw new Error('PDF: bad compressed block');
    }
    for (;;) {
      const sym = decode(lit);
      if (sym === 256) break;
      if (sym < 256) { put(sym); continue; }
      const li = sym - 257, len = PDF_LENGTH_BASE[li] + bits(PDF_LENGTH_EXTRA[li]);
      const di = decode(dist), d = PDF_DIST_BASE[di] + bits(PDF_DIST_EXTRA[di]);
      for (let i = 0; i < len; i++) put(out[outLen - d]);
    }
  }
  return out.slice(0, outLen);
}

/* ---------- PDF objects ---------- */
function pdfLatin1(bytes, from, to) {
  let s = '';
  for (let i = from; i < to; i += 8192) s += String.fromCharCode.apply(null, Array.prototype.slice.call(bytes, i, Math.min(to, i + 8192)));
  return s;
}

/** One PDF value from text at pos → { v, pos }. Names are {n}, refs {ref}, strings {s} (bytes as a latin1 string). */
function pdfParseValue(t, pos) {
  const ws = () => { for (;;) { while (pos < t.length && /[\s\0]/.test(t[pos])) pos++; if (t[pos] === '%') { while (pos < t.length && t[pos] !== '\n' && t[pos] !== '\r') pos++; } else break; } };
  ws();
  const c = t[pos];
  if (c === '<' && t[pos + 1] === '<') {
    pos += 2; const d = {};
    for (;;) {
      ws();
      if (t[pos] === '>' && t[pos + 1] === '>') { pos += 2; return { v: d, pos: pos }; }
      const k = pdfParseValue(t, pos); pos = k.pos;
      const v = pdfParseValue(t, pos); pos = v.pos;
      d[k.v.n] = v.v;
    }
  }
  if (c === '[') {
    pos++; const a = [];
    for (;;) { ws(); if (t[pos] === ']') return { v: a, pos: pos + 1 }; const x = pdfParseValue(t, pos); a.push(x.v); pos = x.pos; }
  }
  if (c === '/') {
    let e = pos + 1; while (e < t.length && !/[\s\/\[\]<>()\{\}%]/.test(t[e])) e++;
    return { v: { n: t.slice(pos + 1, e).replace(/#([0-9A-Fa-f]{2})/g, (m, h) => String.fromCharCode(parseInt(h, 16))) }, pos: e };
  }
  if (c === '(') {
    let depth = 1, s = '', i = pos + 1;
    while (i < t.length && depth) {
      const ch = t[i];
      if (ch === '\\') {
        const n = t[i + 1];
        const map = { n: '\n', r: '\r', t: '\t', b: '\b', f: '\f', '(': '(', ')': ')', '\\': '\\' };
        if (map[n] !== undefined) { s += map[n]; i += 2; }
        else if (/[0-7]/.test(n)) { let o = ''; i++; while (o.length < 3 && /[0-7]/.test(t[i])) o += t[i++]; s += String.fromCharCode(parseInt(o, 8) & 255); }
        else if (n === '\r' || n === '\n') { i += 2; if (n === '\r' && t[i] === '\n') i++; }
        else { s += n; i += 2; }
        continue;
      }
      if (ch === '(') depth++;
      if (ch === ')') { depth--; if (!depth) { i++; break; } }
      s += ch; i++;
    }
    return { v: { s: s }, pos: i };
  }
  if (c === '<') {
    const e = t.indexOf('>', pos);
    let h = t.slice(pos + 1, e).replace(/\s+/g, ''); if (h.length % 2) h += '0';
    let s = ''; for (let i = 0; i < h.length; i += 2) s += String.fromCharCode(parseInt(h.substr(i, 2), 16));
    return { v: { s: s, hex: true }, pos: e + 1 };
  }
  const m = /^([+-]?(?:\d+\.?\d*|\.\d+))(?:\s+(\d+)\s+R\b)?/.exec(t.slice(pos, pos + 40));
  if (m) {
    if (m[2] !== undefined && /^\d+$/.test(m[1])) return { v: { ref: Number(m[1]) }, pos: pos + m[0].length };
    return { v: Number(m[1]), pos: pos + m[1].length };
  }
  const w = /^[A-Za-z*'"]+/.exec(t.slice(pos, pos + 20));
  if (w) return { v: w[0] === 'true' ? true : w[0] === 'false' ? false : w[0] === 'null' ? null : { op: w[0] }, pos: pos + w[0].length };
  throw new Error('PDF: cannot read a value at ' + pos);
}

/** Every "n g obj … endobj" of the file → { n: { dict, stream: bytes|null } }. Streams are inflated when FlateDecode. */
function pdfObjects(bytes) {
  const text = pdfLatin1(bytes, 0, bytes.length);
  // v1.1.59: a password-protected statement says so, instead of failing on unreadable content
  if (/\/Encrypt\b/.test(text)) throw new Error('the PDF is password-protected — not supported yet (send one to get it added)');
  if (/\/ObjStm\b/.test(text)) throw new Error('PDF: objects packed in object streams are not supported');
  const objs = {};
  const re = /(\d+)\s+\d+\s+obj\b/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    let parsed;
    try { parsed = pdfParseValue(text, m.index + m[0].length); } catch (error) { continue; }
    const o = { dict: parsed.v, stream: null };
    const after = text.slice(parsed.pos, parsed.pos + 20);
    const sm = /^\s*stream\r?\n/.exec(after);
    if (sm && parsed.v && typeof parsed.v === 'object') {
      const start = parsed.pos + sm[0].length;
      let len = parsed.v.Length;
      if (len && len.ref !== undefined) len = null;   // indirect: found after all objects are read
      o.streamStart = start; o.streamLen = typeof len === 'number' ? len : null;
      re.lastIndex = typeof len === 'number' ? start + len : text.indexOf('endstream', start);
    }
    objs[m[1]] = o;
  }
  Object.keys(objs).forEach(k => {
    const o = objs[k];
    if (o.streamStart === undefined) return;
    let len = o.streamLen;
    if (len === null) {
      const l = o.dict.Length && o.dict.Length.ref !== undefined ? objs[o.dict.Length.ref] : null;
      len = l && typeof l.dict === 'number' ? l.dict : text.indexOf('endstream', o.streamStart) - o.streamStart;
    }
    let data = bytes.slice(o.streamStart, o.streamStart + len);
    const filters = [].concat(o.dict.Filter || []).map(f => f.n);
    if (filters.length && filters.some(f => f !== 'FlateDecode')) { o.unsupported = filters.join(','); return; }
    if (filters.length) data = pdfInflate(data);
    o.stream = data;
  });
  return objs;
}

/* ---------- text: fonts, positions, lines ---------- */
const PDF_WINANSI_HIGH = { 128: 8364, 130: 8218, 131: 402, 132: 8222, 133: 8230, 134: 8224, 135: 8225, 136: 710, 137: 8240, 138: 352, 139: 8249,
  140: 338, 142: 381, 145: 8216, 146: 8217, 147: 8220, 148: 8221, 149: 8226, 150: 8211, 151: 8212, 152: 732, 153: 8482, 154: 353, 155: 8250,
  156: 339, 158: 382, 159: 376 };
const PDF_GLYPHS = { space: ' ', nbspace: ' ', exclam: '!', quotedbl: '"', numbersign: '#', dollar: '$', percent: '%', ampersand: '&',
  quotesingle: "'", quoteright: '\u2019', parenleft: '(', parenright: ')', asterisk: '*', plus: '+', comma: ',', hyphen: '-', minus: '-', period: '.',
  slash: '/', zero: '0', one: '1', two: '2', three: '3', four: '4', five: '5', six: '6', seven: '7', eight: '8', nine: '9', colon: ':',
  semicolon: ';', less: '<', equal: '=', greater: '>', question: '?', at: '@', bracketleft: '[', backslash: '\\', bracketright: ']',
  underscore: '_', grave: '`', braceleft: '{', bar: '|', braceright: '}', degree: '°', ordfeminine: 'ª', ordmasculine: 'º',
  endash: '–', emdash: '—', bullet: '•', quotedblleft: '“', quotedblright: '”', quoteleft: '‘', exclamdown: '¡', questiondown: '¿',
  aacute: 'á', eacute: 'é', iacute: 'í', oacute: 'ó', uacute: 'ú', ntilde: 'ñ', udieresis: 'ü', Aacute: 'Á', Eacute: 'É', Iacute: 'Í',
  Oacute: 'Ó', Uacute: 'Ú', Ntilde: 'Ñ', Udieresis: 'Ü', agrave: 'à', egrave: 'è', ccedilla: 'ç', Ccedilla: 'Ç' };

function pdfGlyphChar(name) {
  if (name.length === 1) return name;
  if (PDF_GLYPHS[name] !== undefined) return PDF_GLYPHS[name];
  const u = /^uni([0-9A-Fa-f]{4})$/.exec(name);
  if (u) { const c = String.fromCharCode(parseInt(u[1], 16)); return c === '\u00a0' ? ' ' : c; }
  return '';
}

/** A font → { decode(bytes) → text, codes(bytes) → [code…], width(code) → glyph width in 1/1000 em }. */
function pdfFont(font, get) {
  const val = x => (x && x.ref !== undefined ? get(x).dict : x);
  const type0 = font.Subtype && font.Subtype.n === 'Type0';
  const widths = {};
  let defaultWidth = 500;
  if (type0) {
    const desc = val((val(font.DescendantFonts) || [])[0]) || {};
    defaultWidth = typeof desc.DW === 'number' ? desc.DW : 1000;
    const w = val(desc.W) || [];
    for (let i = 0; i < w.length;) {
      const first = w[i], next = val(w[i + 1]);
      if (Array.isArray(next)) { next.forEach((x, k) => { widths[first + k] = x; }); i += 2; }
      else { for (let c = first; c <= next; c++) widths[c] = w[i + 2]; i += 3; }
    }
  } else {
    const list = val(font.Widths), first = typeof font.FirstChar === 'number' ? font.FirstChar : 0;
    if (Array.isArray(list)) list.forEach((x, k) => { widths[first + k] = val(x); });
    const d = val(font.FontDescriptor);
    if (d && typeof d.MissingWidth === 'number') defaultWidth = d.MissingWidth;
  }
  const decode = pdfFontDecoder(font, get);
  const codes = s => { const out = []; if (type0) { for (let i = 0; i + 1 < s.length; i += 2) out.push((s.charCodeAt(i) << 8) | s.charCodeAt(i + 1)); }
    else { for (let i = 0; i < s.length; i++) out.push(s.charCodeAt(i)); } return out; };
  return { decode: decode, codes: codes, width: c => (widths[c] !== undefined ? widths[c] : (c === 32 ? 278 : defaultWidth)), type0: type0 };
}

/** A font → decoder of its string bytes. */
function pdfFontDecoder(font, get) {
  const toUni = font.ToUnicode ? get(font.ToUnicode) : null;
  if (toUni && toUni.stream) {
    const cmap = pdfLatin1(toUni.stream, 0, toUni.stream.length), map = {};
    let width = 1;
    const hex = h => parseInt(h, 16);
    const uni = h => { let s = ''; for (let i = 0; i < h.length; i += 4) s += String.fromCharCode(hex(h.substr(i, 4))); return s; };
    (cmap.match(/beginbfchar([\s\S]*?)endbfchar/g) || []).forEach(block => {
      (block.match(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g) || []).forEach(pair => {
        const p = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/.exec(pair); width = p[1].length / 2; map[hex(p[1])] = uni(p[2]);
      });
    });
    (cmap.match(/beginbfrange([\s\S]*?)endbfrange/g) || []).forEach(block => {
      (block.match(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g) || []).forEach(tr => {
        const p = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/.exec(tr); width = p[1].length / 2;
        for (let c = hex(p[1]), k = 0; c <= hex(p[2]); c++, k++) map[c] = String.fromCharCode(hex(p[3]) + k);
      });
    });
    const twoBytes = (font.Subtype && font.Subtype.n === 'Type0') || width === 2;
    return s => {
      let out = '';
      if (twoBytes) { for (let i = 0; i + 1 < s.length; i += 2) out += map[(s.charCodeAt(i) << 8) | s.charCodeAt(i + 1)] || ''; }
      else { for (let i = 0; i < s.length; i++) out += map[s.charCodeAt(i)] !== undefined ? map[s.charCodeAt(i)] : ''; }
      return out;
    };
  }
  let enc = font.Encoding ? (font.Encoding.ref !== undefined ? get(font.Encoding).dict : font.Encoding) : null;
  const diff = {};
  if (enc && enc.Differences) {
    let code = 0;
    enc.Differences.forEach(x => { if (typeof x === 'number') code = x; else if (x && x.n !== undefined) diff[code++] = pdfGlyphChar(x.n); });
  }
  return s => {
    let out = '';
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      if (diff[c] !== undefined) out += diff[c];
      else if (c >= 32 && c < 127) out += String.fromCharCode(c);
      else if (c >= 160) out += c === 160 ? ' ' : String.fromCharCode(c);
      else if (PDF_WINANSI_HIGH[c]) out += String.fromCharCode(PDF_WINANSI_HIGH[c]);
    }
    return out;
  };
}

/** The text pieces of one page's content, with their positions. */
function pdfPagePieces(content, fonts) {
  const t = pdfLatin1(content, 0, content.length);
  const mul = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
    a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
  let ctm = [1, 0, 0, 1, 0, 0], tm = [1, 0, 0, 1, 0, 0], lm = [1, 0, 0, 1, 0, 0], leading = 0, font = null, size = 0;
  let charSpace = 0, wordSpace = 0, hScale = 1;
  const stack = [], pieces = [], ops = [];
  let pos = 0;
  // a string at the current text position: its text, where it starts and how wide it is (device space)
  const show = str => {
    if (!font) return;
    let advance = 0;
    font.codes(str).forEach(c => { advance += (font.width(c) / 1000 * size + charSpace + (c === 32 && !font.type0 ? wordSpace : 0)) * hScale; });
    const m = mul(tm, ctm), devScale = Math.sqrt(m[0] * m[0] + m[1] * m[1]);
    const text = font.decode(str);
    if (text) pieces.push({ x: m[4], y: m[5], text: text, size: Math.abs(size * devScale), width: advance * devScale, order: pieces.length });
    tm = mul([1, 0, 0, 1, advance, 0], tm);
  };
  while (pos < t.length) {
    while (pos < t.length && /[\s\0]/.test(t[pos])) pos++;
    if (pos >= t.length) break;
    if (t[pos] === '%') { while (pos < t.length && t[pos] !== '\n') pos++; continue; }
    if (t.startsWith('BI', pos) && /\s/.test(t[pos + 2] || ' ')) { const e = t.indexOf('EI', pos); pos = e < 0 ? t.length : e + 2; continue; }   // inline image
    let v;
    try { v = pdfParseValue(t, pos); } catch (error) { pos++; continue; }
    pos = v.pos;
    if (!v.v || v.v.op === undefined) { ops.push(v.v); continue; }
    const op = v.v.op, a = ops.splice(0, ops.length);
    if (op === 'q') stack.push(ctm);
    else if (op === 'Q') ctm = stack.pop() || [1, 0, 0, 1, 0, 0];
    else if (op === 'cm' && a.length === 6) ctm = mul(a, ctm);
    else if (op === 'BT') { tm = [1, 0, 0, 1, 0, 0]; lm = tm; }
    else if (op === 'Tf') { font = fonts[a[0] && a[0].n] || null; size = a[1] || size; }
    else if (op === 'Tc') charSpace = a[0] || 0;
    else if (op === 'Tw') wordSpace = a[0] || 0;
    else if (op === 'Tz') hScale = (typeof a[0] === 'number' ? a[0] : 100) / 100;
    else if (op === 'Tm' && a.length === 6) { tm = a.slice(); lm = tm; }
    else if (op === 'Td' || op === 'TD') { if (op === 'TD') leading = -a[1]; lm = mul([1, 0, 0, 1, a[0], a[1]], lm); tm = lm; }
    else if (op === 'TL') leading = a[0];
    else if (op === 'T*') { lm = mul([1, 0, 0, 1, 0, -leading], lm); tm = lm; }
    else if (op === 'Tj') { if (a[0] && a[0].s !== undefined) show(a[0].s); }
    else if (op === "'" || op === '"') {
      if (op === '"') { wordSpace = a[0] || 0; charSpace = a[1] || 0; }
      lm = mul([1, 0, 0, 1, 0, -leading], lm); tm = lm; const s = a[a.length - 1]; if (s && s.s !== undefined) show(s.s);
    }
    else if (op === 'TJ' && Array.isArray(a[0])) {
      a[0].forEach(x => {
        if (x && x.s !== undefined) show(x.s);
        else if (typeof x === 'number') tm = mul([1, 0, 0, 1, -x / 1000 * size * hScale, 0], tm);   // kerning / spacing
      });
    }
  }
  return pieces;
}

/** The PDF's text, page by page, rebuilt into lines (top to bottom, left to right). */
function pdfToText(bytes) {
  const objs = pdfObjects(bytes);
  const get = r => (r && r.ref !== undefined ? objs[r.ref] : null) || { dict: r, stream: null };
  const catalog = Object.keys(objs).map(k => objs[k]).find(o => o.dict && o.dict.Type && o.dict.Type.n === 'Catalog');
  if (!catalog) throw new Error('PDF: no document catalog');
  const pages = [];
  const walk = (node, inherited) => {
    const d = get(node).dict || {};
    const res = d.Resources || inherited;
    if (d.Type && d.Type.n === 'Pages') (d.Kids || []).forEach(k => walk(k, res));
    else pages.push({ dict: d, resources: res });
  };
  walk(catalog.dict.Pages, null);
  const pageTexts = pages.map(p => {
    const res = p.resources && p.resources.ref !== undefined ? get(p.resources).dict : (p.resources || {});
    const fontDict = res.Font && res.Font.ref !== undefined ? get(res.Font).dict : (res.Font || {});
    const fonts = {};
    Object.keys(fontDict || {}).forEach(name => { const f = get(fontDict[name]).dict; if (f) fonts[name] = pdfFont(f, get); });
    const parts = [].concat(p.dict.Contents || []).map(c => get(c));
    parts.forEach(c => { if (c.unsupported) throw new Error('PDF: content compressed with ' + c.unsupported + ' is not supported'); });
    const total = parts.reduce((n, c) => n + (c.stream ? c.stream.length + 1 : 0), 0), content = new Uint8Array(total);
    let at = 0;
    parts.forEach(c => { if (c.stream) { content.set(c.stream, at); at += c.stream.length; content[at++] = 10; } });
    const pieces = pdfPagePieces(content, fonts).sort((a, b) => (b.y - a.y) || (a.x - b.x) || (a.order - b.order));
    const lines = [];
    pieces.forEach(pc => {
      const line = lines.length && Math.abs(lines[lines.length - 1].y - pc.y) <= Math.max(2, pc.size * 0.3) ? lines[lines.length - 1] : null;
      if (line) line.items.push(pc); else lines.push({ y: pc.y, items: [pc] });
    });
    return lines.map(l => l.items.sort((a, b) => (a.x - b.x) || (a.order - b.order)).reduce((s, pc, i, all) => {
      if (!i) return pc.text;
      const prev = all[i - 1], gap = pc.x - (prev.x + prev.width);   // a word space is ~0.25 em; kerning is far less
      return s + (gap > Math.min(prev.size, pc.size) * 0.12 ? ' ' : '') + pc.text;
    }, '').replace(/\s+/g, ' ').trim()).filter(Boolean).join('\n');
  });
  return pageTexts.join('\n\f\n');
}
