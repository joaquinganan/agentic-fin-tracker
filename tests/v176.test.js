'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

// v1.1.76 — reported: the 6 AM run hit Apps Script's 6-minute limit inside recategorize. The investments step came
// after it, so the day's Portfolio History snapshot was never taken (the next email: "—, no update on Mon 26 Oct"),
// and that morning's email showed the previous day's update in green ("Data updated Sun 25 Oct at 6:39 AM").

test('the summary sent after a morning without an update says so — it was green at 25 hours (reported)', () => {
  const h = load();
  const last = { at: h.date(2026, 10, 25, 6, 39).toISOString(), saved: 1 };
  const d = h.plain(h.ctx.summaryDataHealth(last, h.date(2026, 10, 26, 7, 47)));
  assert.equal(d.tone, 'warn');
  assert.match(d.text, /^No automatic update this morning — the last one was Sun 25 Oct at 6:39 AM/);
  // a summary sent before the morning update (around 6:00) still shows yesterday's update as it is
  const early = h.plain(h.ctx.summaryDataHealth(last, h.date(2026, 10, 26, 6, 10)));
  assert.equal(early.tone, 'good');
  assert.match(early.text, /^Data updated Sun 25 Oct at 6:39 AM/);
  // more than 26 hours: as before
  assert.match(h.plain(h.ctx.summaryDataHealth(last, h.date(2026, 10, 26, 9))).text, /^No automatic update since Sun 25 Oct/);
  // today's update: green
  assert.equal(h.plain(h.ctx.summaryDataHealth({ at: h.date(2026, 10, 26, 6, 40).toISOString(), saved: 2 }, h.date(2026, 10, 26, 7, 47))).tone, 'good');
});

test('a run cut off by the time limit is reported as such, naming the step', () => {
  const h = load();
  const last = { at: h.date(2026, 10, 26, 6, 40).toISOString(), saved: 3, unfinished: true, step: 'recategorize' };
  const d = h.plain(h.ctx.summaryDataHealth(last, h.date(2026, 10, 26, 7, 47)));
  assert.equal(d.tone, 'warn');
  assert.match(d.text, /^The update today at 6:40 AM did not finish: Apps Script stopped it at its 6-minute limit during "recategorize" · 3 new transaction\(s\) saved/);
  // started less than 6 minutes ago: still running, not cut off
  const running = h.plain(h.ctx.summaryDataHealth(last, h.date(2026, 10, 26, 6, 43)));
  assert.equal(running.tone, 'note');
  assert.match(running.text, /still running/);
});

test('steps left for the next run show in the summary line', () => {
  const h = load();
  const d = h.plain(h.ctx.summaryDataHealth({ at: h.date(2026, 10, 26, 6, 40).toISOString(), saved: 2,
    deferred: ['recategorize', 'sheet order'] }, h.date(2026, 10, 26, 7, 47)));
  assert.equal(d.tone, 'warn');
  assert.match(d.text, /· left for the next run \(time limit\): recategorize, sheet order$/);
});

function book() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ BHD: true })], ['setupDate', '2026-09-01']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  mock.gmail.threads.push(fakeThread('t0', [h.fakeMessage({ subject: 'BHD Notificación de Transacciones', from: 'Alertas@bhd.com.do',
    body: fixture('bhd_consumo_cacharepa'), date: h.date(2026, 9, 3, 10), id: 'm0' })]));
  return { h, mock };
}

test('the day\'s snapshot is taken before recategorize, and a run killed there is recorded as unfinished', () => {
  const { h } = book();
  // what a kill inside recategorize would leave behind: what was recorded and written up to that moment
  h.get('var SEEN = null');
  h.get('recategorizeAllTransactions = function () { SEEN = { last: readLastRun(), ' +
    'history: !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HISTORY_SHEET) }; return 0; }');
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const seen = h.plain(h.get('SEEN'));
  assert.equal(seen.history, true, 'Portfolio History already written when recategorize starts');
  assert.deepEqual([seen.last.unfinished, seen.last.step, seen.last.saved], [true, 'recategorize', 1]);
  // a run that finishes replaces it with its outcome
  const done = h.plain(h.ctx.readLastRun());
  assert.equal(done.unfinished, undefined);
  assert.deepEqual([done.saved, done.deferred], [1, []]);
});

test('📘 Start here shows a cut-off run as a problem', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  h.ctx.recordLastRun({ at: h.date(2026, 10, 26, 6, 40).toISOString(), saved: 3, unfinished: true, step: 'recategorize' });
  const s = h.plain(h.ctx.startHereStatus(h.date(2026, 10, 26, 9))).steps.find(x => x.id === 'lastRun');
  assert.equal(s.status, 'warn');
  assert.match(s.detail, /6-minute limit during "recategorize"/);
});
