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
const SCRIPT_VERSION = "1.1.56"; // bump on every release (v1.1.19 fixed it being stuck at 1.1.12)
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
    otherIncome: existing.otherIncome, otherIncomeCurrency: existing.otherIncomeCurrency, cards: existing.cards
  } : { email: defaultEmail };
  // v1.1.27: card catalogue for the "Credit Cards" section (public product facts only)
  const catalogJson = JSON.stringify(Object.keys(CARD_PRODUCTS).map(id =>
    ({ id: id, bank: CARD_PRODUCTS[id].bank, name: CARD_PRODUCTS[id].name }))).replace(/</g, '\\u003c');
  const prefillJson = JSON.stringify(prefill).replace(/</g, '\\u003c');
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
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

          <div class="row">
            <div class="form-group">
              <label>➕ Other Monthly Income</label>
              <input type="number" id="otherIncome" placeholder="0.00" step="0.01" min="0">
            </div>
            <div class="form-group">
              <label>Other income currency</label>
              <select id="otherIncomeCurrency">
                <option value="DOP">DOP</option>
                <option value="USD">USD</option>
              </select>
            </div>
          </div>
          <div class="hint" style="margin:-6px 0 4px">Freelance work, rent you receive, etc. Added in full to your net income — no deductions.</div>

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

          <div class="section-title">🏦 Banks to Track</div>
          <div class="checkbox-group">
            <div class="checkbox-item"><input type="checkbox" id="bank_lafise" checked onchange="renderCards()"><label for="bank_lafise">LAFISE</label></div>
            <div class="checkbox-item"><input type="checkbox" id="bank_banesco" checked onchange="renderCards()"><label for="bank_banesco">BANESCO</label></div>
            <div class="checkbox-item"><input type="checkbox" id="bank_bhd" checked onchange="renderCards()"><label for="bank_bhd">BHD</label></div>
            <div class="checkbox-item"><input type="checkbox" id="bank_popular" checked onchange="renderCards()"><label for="bank_popular">POPULAR</label></div>
            <div class="checkbox-item"><input type="checkbox" id="bank_bdi" checked onchange="renderCards()"><label for="bank_bdi">BDI</label></div>
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
              <div class="checkbox-item"><input type="checkbox" id="sec_vsAverage" checked><label for="sec_vsAverage">Comparison with your averages</label></div>
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
        const BANKS = ['LAFISE', 'BANESCO', 'BHD', 'POPULAR', 'BDI'];
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

        const EXISTING = ${prefillJson};
        (function prefill() {
          const setVal = function(id, v) { if (v !== undefined && v !== null && v !== '') el(id).value = v; };
          setVal('email', EXISTING.email);
          setVal('incomeCurrency', EXISTING.incomeCurrency);
          setVal('monthlyIncome', EXISTING.monthlyIncome);
          setVal('otherIncome', EXISTING.otherIncome);
          setVal('otherIncomeCurrency', EXISTING.otherIncomeCurrency);
          if (EXISTING.deductionMode === 'manual') el('dedManual').checked = true;
          if (EXISTING.deductionMode !== 'auto' || EXISTING.incomeCurrency !== 'DOP') {
            setVal('arsAmount', EXISTING.ARS);
            setVal('afpAmount', EXISTING.AFP);
            setVal('taxRate', EXISTING.taxRate);
            setVal('isrAmount', EXISTING.ISR);
          }
          if (EXISTING.banksToTrack) {
            ['LAFISE', 'BANESCO', 'BHD', 'POPULAR', 'BDI'].forEach(function(b) {
              if (b in EXISTING.banksToTrack) el('bank_' + b.toLowerCase()).checked = !!EXISTING.banksToTrack[b];
            });
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
            otherIncome: num('otherIncome'),
            otherIncomeCurrency: el('otherIncomeCurrency').value,
            cards: cards,
            deductionMode: mode(),
            deductions: {
              ARS: num('arsAmount'), AFP: num('afpAmount'),
              taxRate: cur === 'DOP' ? 0 : num('taxRate'),
              ISR: cur === 'DOP' ? num('isrAmount') : 0
            },
            banksToTrack: {
              LAFISE: el('bank_lafise').checked, BANESCO: el('bank_banesco').checked, BHD: el('bank_bhd').checked,
              POPULAR: el('bank_popular').checked, BDI: el('bank_bdi').checked
            },
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
  if (Number(config.otherIncome || 0) < 0) return "Other income can't be negative.";
  if (config.otherIncomeCurrency && ['USD', 'DOP'].indexOf(config.otherIncomeCurrency) === -1) return "Other income currency must be USD or DOP.";
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
        ["otherIncome", Number(config.otherIncome) || 0],
        ["otherIncomeCurrency", config.otherIncomeCurrency || 'DOP'],
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
      ];
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
      (config.otherIncome > 0 ? ", including other income" : ""),
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
                              investments: investments, unrecognizedOpen: unrecognizedOpen, stopped: stopped, deferred: deferred }));
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
  if (s.reversals) lines.push("↩️ Reversals: " + s.reversals + " — saved as negative rows that cancel the original purchase");
  if (r.results.reversalsUnmatched) {
    lines.push("⚠️ Reversal without its original purchase: " + r.results.reversalsUnmatched +
      " — look for \"" + REVERSAL_UNMATCHED + "\" in Transactions");
  }
  if (r.search && r.search.capped) {
    lines.push("⚠️ Hit the " + MAX_THREADS_PER_RUN + "-thread cap — split this date range into smaller pieces.");
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

➕ Other income: ${config.otherIncome} ${config.otherIncomeCurrency} (no deductions)
💳 Cards: ${resolveCards(config.cards).map(c => c.bank + (c.name ? ' ' + c.name : '') + (c.closeDay ? ' · closes day ' + c.closeDay : '') + (c.dueDay ? ' · due day ' + c.dueDay : '')).join('; ') || 'none set'}

📬 Daily summary: ${config.notifyEnabled ? 'on — around ' + config.notifyHour + ':00 to ' + (config.notifyEmail || config.email) : 'off'}
🗓️ Monthly summary: ${config.notifyMonthly ? 'on — the 1st, around ' + config.notifyHour + ':00' : 'off'}

🏦 Banks Tracked:
  LAFISE: ${config.banksToTrack.LAFISE ? '✅' : '❌'}
  BANESCO: ${config.banksToTrack.BANESCO ? '✅' : '❌'}
  BHD: ${config.banksToTrack.BHD ? '✅' : '❌'}
  POPULAR: ${config.banksToTrack.POPULAR ? '✅' : '❌'}
  BDI: ${config.banksToTrack.BDI ? '✅' : '❌'}

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
