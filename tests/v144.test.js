'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

function book() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ BHD: true })], ['setupDate', '2026-09-01']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  return { h, mock };
}
const bhd = (h, name, id, day) => fakeThread('t-' + id, [h.fakeMessage({ subject: 'BHD Notificación de Transacciones', from: 'Alertas@bhd.com.do',
  body: fixture(name), date: h.date(2026, 9, day, 10), id })]);

test('threads marked by an earlier run are not marked again (reported: a year-long range spent minutes re-marking)', () => {
  const { h, mock } = book();
  mock.gmail.threads.push(bhd(h, 'bhd_consumo_cacharepa', 'a', 3), bhd(h, 'bhd_consumo_medicar', 'b', 4));
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.deepEqual(mock.gmail.markedRead.slice().sort(), ['t-a', 't-b']);
  assert.ok(mock.gmail.queries.some(q => /\{is:unread -label:Procesado\}$/.test(q)), 'asked Gmail which threads are still pending');
  // same range again, plus one new thread: only the new one is marked
  mock.gmail.threads.push(bhd(h, 'bhd_consumo_pedidosya', 'c', 5));
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.deepEqual(mock.gmail.markedRead.slice().sort(), ['t-a', 't-b', 't-c'], 't-a and t-b not marked twice');
  assert.ok(h.logs.some(l => /Marked 1 thread\(s\) as processed \(2 already were\)/.test(l)));
});

test('marking stops at the time budget; the rest is left for the next run', () => {
  const h = load();
  const threads = [fakeThread('x1', []), fakeThread('x2', [])];
  let t = 0;
  const r = h.plain(h.ctx.markEmailsAsProcessed(threads, new Set(), { deadline: 50, clock: () => (t += 100) }));
  assert.equal(r.marked, 0);
  assert.ok(h.logs.some(l => /Stopped marking at the time budget: 2 thread\(s\) left/.test(l)));
});

test('every phase logs how long it took, so a slow one shows in View › Executions', () => {
  const { h, mock } = book();
  mock.gmail.threads.push(bhd(h, 'bhd_consumo_cacharepa', 'a', 3));
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  for (const phase of ['save', 'mark', 'read log', 'sort', 'recategorize', 'investments', 'unrecognized']) {
    assert.ok(h.logs.some(l => l.startsWith('⏱ ' + phase + ': ')), phase);
  }
});

test('a sale on the start day is cash at the start: it funds purchases without a deposit (a Dec 31 sale settling in January)', () => {
  const h = load();
  const L = rows => [['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id']]
    .concat(rows.map(r => { const x = r.slice(); x[0] = h.date(...x[0]); while (x.length < 12) x.push(''); return x; }));
  // fictitious: US$3,000 on Dec 31 including a US$300 sale that day; US$1,000 deposited; US$1,300 bought since
  const rows = [[[2025, 12, 31], 'HAPI', 'Valuation', '', '', '', 3000, '', 'USD'],
    [[2025, 12, 31], 'HAPI', 'Sell', 'SCHD', 10, 30.01, 300.1, 0.1, 'USD'],
    [[2026, 1, 5], 'HAPI', 'Deposit', '', '', '', 1000, '', 'USD'],
    [[2026, 1, 6], 'HAPI', 'Buy', 'META', 1, 1299.85, 1299.85, 0.15, 'USD']];
  const withSale = h.plain(h.ctx.computeReturns(L(rows), { HAPI: 4400 }, { usdRate: 60, today: '2026-09-28' })).accounts[0];
  assert.equal(withSale.startCash, 300);
  assert.equal(withSale.unfunded, 0, 'covered by the deposit and the start-day sale');
  const without = h.plain(h.ctx.computeReturns(L(rows.filter(r => r[2] !== 'Sell')), { HAPI: 4400 }, { usdRate: 60, today: '2026-09-28' })).accounts[0];
  assert.equal(+without.unfunded.toFixed(2), 300);
});
