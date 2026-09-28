'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

function counted(h, name, id, day) {
  const m = h.fakeMessage({ subject: 'BHD Notificación de Transacciones', from: 'Alertas@bhd.com.do', body: fixture(name), date: h.date(2026, 9, day, 10), id });
  const read = m.getPlainBody;
  m.reads = 0;
  m.getPlainBody = () => { m.reads++; return read.call(m); };
  return m;
}

test('emails already saved are not read again (reading is the slow part)', () => {
  const h = load();
  const saved = counted(h, 'bhd_consumo_cacharepa', 'm1', 3), fresh = counted(h, 'bhd_consumo_medicar', 'm2', 4);
  const r = h.ctx.extractTransactionsFromThreads([fakeThread('t1', [saved]), fakeThread('t2', [fresh])], {}, null, { skipIds: new Set(['m1']) });
  assert.equal(saved.reads, 0, 'never opened');
  assert.ok(fresh.reads > 0);
  assert.equal(r.stats.alreadySaved, 1);
  assert.equal(r.transactions.length, 1);
  assert.equal(r.processedThreads.length, 2, 'still counts as processed (it can be marked read)');
});

test('reading stops at the deadline; the rest is left for the next run', () => {
  const h = load();
  let t = 0;
  const threads = ['bhd_consumo_cacharepa', 'bhd_consumo_medicar', 'bhd_consumo_pedidosya'].map((n, i) => fakeThread('t' + i, [counted(h, n, 'm' + i, 3 + i)]));
  const r = h.ctx.extractTransactionsFromThreads(threads, {}, null, { deadline: 250, clock: () => (t += 100) });   // 100, 200, 300…
  assert.equal(r.processedThreads.length, 2);
  assert.deepEqual(h.plain(r.stopped), { remaining: 1, processed: 2 });
  assert.equal(threads[2].getMessages()[0].reads, 0);
});

function configured() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ BHD: true })], ['setupDate', '2026-09-01']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions');
  h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  return { h, mock };
}

test('a run that would hit Google\'s limit stops in time, finishes its steps, and the next run completes it', () => {
  const { h, mock } = configured();
  const names = ['bhd_consumo_cacharepa', 'bhd_consumo_medicar', 'bhd_consumo_pedidosya'];
  mock.gmail.threads.push(...names.map((n, i) => fakeThread('t' + i, [counted(h, n, 'm' + i, 3 + i)])));
  // a slow clock: every check is 2 minutes later — reading stops after the first thread (budget: 3.5 minutes)
  // (v1.1.44: time stands still after reading — marking has its own time check now)
  h.get('runClock = (() => { let n = 0, t = 0; return () => (n++ < 3 ? (t += 120000) : t); })()');
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const first = mock.ui.alerts[mock.ui.alerts.length - 1];
  assert.match(first, /⏸ Stopped reading early to stay within Google's 6-minute limit — 2 of 3 thread\(s\) left/);
  assert.equal(mock.ss.getSheetByName('Transactions')._rows(14).length, 1, 'what was read is saved');
  assert.deepEqual(mock.gmail.markedRead.slice().sort(), ['t0'], 'only the threads read are marked processed');
  assert.equal(JSON.parse(mock.props.FT_LAST_RUN).partial, true);
  assert.notEqual(mock.ss.toastLog[mock.ss.toastLog.length - 1].timeout, -1, 'no toast left open');
  assert.ok(first.includes('Recategorized'), 'the cleanup steps ran');
  // the next run, with time to spare, picks up the rest and doesn't read the first email again
  h.get('runClock = () => Date.now()');
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.equal(mock.ss.getSheetByName('Transactions')._rows(14).length, 3);
  assert.equal(mock.gmail.threads[0].getMessages()[0].reads, 1, 'read once in total');
  const second = mock.ui.alerts[mock.ui.alerts.length - 1];
  assert.doesNotMatch(second, /Stopped reading early/);
  assert.match(second, /Read before, skipped: 1 email/);
  assert.equal(JSON.parse(mock.props.FT_LAST_RUN).partial, false);
});

test('broker emails already in the ledger are not read again either', () => {
  const { h, mock } = configured();
  const m = h.fakeMessage({ subject: '✅ Order Executed ', from: 'Hapi App <no-reply@hapi.trade>', body: fixture('investments/hapi_order_buy'),
    date: h.date(2026, 9, 8, 9), id: 'm-hapi' });
  let reads = 0; const read = m.getPlainBody; m.getPlainBody = () => { reads++; return read.call(m); };
  mock.gmail.threads.push(fakeThread('t-hapi', [m]));
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.equal(reads, 1);
  assert.equal(mock.ss.getSheetByName('Investment Ledger')._rows(12).length, 1);
});
