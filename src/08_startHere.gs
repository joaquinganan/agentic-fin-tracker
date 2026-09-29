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
