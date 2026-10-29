'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

// v1.1.77 — reported: a manual Recategorize finished in about a minute ("✅ Recategorized — …" in the log), then
// Executions showed it as "Timed Out" at 6 minutes. Its "✅ Done" window was closed minutes later, and an alert waits
// until it is closed — inside withRunLock(), so the run lock was held the whole time and a 6 AM run in that window
// would have been skipped ("Another run is in progress").

function configured(options) {
  const mock = makeServices(Object.assign({ ui: true }, options || {}));
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ BHD: true })], ['setupDate', '2026-09-01']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions');
  h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  return { h, mock };
}

test('Recategorize: the "Done" window is shown after the run lock is released (reported)', () => {
  const { h, mock } = configured();
  h.ctx.recategorizeAllTransactionsPrompt();
  assert.equal(mock.ui.alerts.length, 2, 'the confirmation and the result');
  assert.match(mock.ui.alerts[0], /Continue\?/);
  assert.match(mock.ui.alerts[1], /^✅ Done — \d+ cell\(s\) updated\.$/);
  assert.deepEqual(mock.ui.alertsLocked, [false, false], 'no window waits with the lock held');
  assert.equal(mock.lock.held, false);
  assert.ok(h.logs.some(l => /waiting for its window to be closed/.test(l)), 'the log explains a later "Timed Out"');
});

test('every action under the run lock shows its result after releasing it — summary and errors alike', () => {
  let { h, mock } = configured();
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.ok(mock.ui.alerts.length >= 1);
  assert.ok(mock.ui.alertsLocked.every(l => l === false), 'run summary: ' + mock.ui.alertsLocked);

  ({ h, mock } = configured());
  h.ctx.refreshInvestmentsNow();
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /Investments refreshed/);
  assert.ok(mock.ui.alertsLocked.every(l => l === false));

  ({ h, mock } = configured());   // an error half-way: its window too
  mock.services.GmailApp.search = () => { throw new Error('Gmail is unavailable'); };
  h.ctx.runGmailMonitor();
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /Gmail is unavailable/);
  assert.ok(mock.ui.alertsLocked.every(l => l === false));
});

test('withRunLock: alerts keep their order, the result is returned, and an exception still shows them', () => {
  const { h, mock } = configured();
  const out = h.ctx.withRunLock(() => { h.ctx.safeAlert('first'); h.ctx.safeAlert('second'); return 42; });
  assert.equal(out, 42);
  assert.deepEqual(mock.ui.alerts, ['first', 'second']);
  assert.deepEqual(mock.ui.alertsLocked, [false, false]);
  assert.throws(() => h.ctx.withRunLock(() => { h.ctx.safeAlert('before the error'); throw new Error('boom'); }), /boom/);
  assert.equal(mock.ui.alerts[2], 'before the error');
  assert.equal(mock.lock.held, false);
  // outside a run, an alert is shown at once, as before
  h.ctx.safeAlert('outside');
  assert.equal(mock.ui.alerts[3], 'outside');
});

test('a run already in progress: the "⏳" window is shown at once, and the other run\'s lock is untouched', () => {
  const { h, mock } = configured({ lockBusy: true });
  h.ctx.recategorizeAllTransactionsPrompt();
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /Another run is already in progress/);
});

test('the 6 AM trigger (no UI): the summary is logged, nothing throws, no "waiting" line', () => {
  const { h, mock } = configured({ ui: false });
  h.ctx.runGmailMonitor();
  assert.equal(mock.lock.held, false);
  assert.ok(h.logs.some(l => /^\[No UI available — trigger-based run\]/.test(l)));
  assert.ok(!h.logs.some(l => /waiting for its window/.test(l)));
});
