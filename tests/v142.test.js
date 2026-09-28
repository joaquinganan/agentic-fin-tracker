'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

const L = (h, rows) => [['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id']]
  .concat(rows.map(r => { const x = r.slice(); x[0] = h.date(...x[0]); while (x.length < 12) x.push(''); return x; }));
const notice = (d, id) => [d, 'HAPI', 'Notice', '', '', '', '', '', 'USD', 'email', 'Deposit completed', 'gmail:' + id];
// fictitious dates and amounts
const ledgerRows = () => [
  [[2025, 12, 31], 'HAPI', 'Valuation', '', '', '', 4000, '', 'USD', 'manual'],   // tracking starts here
  notice([2025, 12, 31], 'n0'),                                                  // inside the start value: not checked
  notice([2026, 2, 10], 'n1'), notice([2026, 5, 20], 'n2'), notice([2026, 7, 15], 'n3'), notice([2026, 8, 20], 'n4'),
  [[2026, 8, 18], 'HAPI', 'Deposit', '', '', '', 60000, '', 'DOP', 'bank', 'BANESCO → OUROSR SRL', 'bank:b1']];   // 2 days before n4

test('each deposit notice is paired with a recorded deposit near its date; the rest are missing', () => {
  const h = load();
  const r = h.plain(h.ctx.checkDepositNotices(L(h, ledgerRows())));
  assert.deepEqual(r.missing.map(m => m.id), ['n1', 'n2', 'n3']);
  assert.deepEqual(r.matchedIds.sort(), ['n0', 'n4']);
  assert.match(r.missing[0].reason, /^Deposit without amount — add it: .*A deposit \(2026-02-10\)/);
  assert.equal(r.missing[0].bank, 'HAPI');
});

test('the window: up to 7 days before and 2 after; one deposit covers one notice', () => {
  const h = load();
  const run = rows => h.plain(h.ctx.checkDepositNotices(L(h, rows))).missing.map(m => m.id);
  const dep = d => [d, 'HAPI', 'Deposit', '', '', '', 100, '', 'USD', 'manual'];
  assert.deepEqual(run([notice([2026, 3, 10], 'a'), dep([2026, 3, 3])]), [], '7 days before: paired');
  assert.deepEqual(run([notice([2026, 3, 10], 'a'), dep([2026, 3, 2])]), ['a'], '8 days before: not');
  assert.deepEqual(run([notice([2026, 3, 10], 'a'), dep([2026, 3, 12])]), [], '2 days after: paired');
  assert.deepEqual(run([notice([2026, 3, 10], 'a'), dep([2026, 3, 13])]), ['a']);
  assert.deepEqual(run([notice([2026, 3, 10], 'a'), notice([2026, 3, 11], 'b'), dep([2026, 3, 10])]), ['b'], 'one deposit, two notices');
  assert.deepEqual(run([notice([2026, 3, 10], 'a'), [[2026, 3, 10], 'IBKR', 'Deposit', '', '', '', 100, '', 'USD']]), ['a'], 'same account only');
});

test('notices move no money: positions, cash and returns are the same with or without them', () => {
  const h = load();
  const without = L(h, ledgerRows().filter(r => r[2] !== 'Notice'));
  const withN = L(h, ledgerRows());
  assert.deepEqual(h.plain(h.ctx.computeHoldings(withN, { usdRate: 60, year: 2026 })), h.plain(h.ctx.computeHoldings(without, { usdRate: 60, year: 2026 })));
  assert.deepEqual(h.plain(h.ctx.computeReturns(withN, { HAPI: 6000 }, { usdRate: 60, today: '2026-09-28' })),
    h.plain(h.ctx.computeReturns(without, { HAPI: 6000 }, { usdRate: 60, today: '2026-09-28' })));
});

function book() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"BHD":true}'], ['monthlyIncome', 1000], ['incomeCurrency', 'USD']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  return { h, mock };
}

test('missing deposits appear in Unrecognized with a Gmail link; adding one takes its row off', () => {
  const { h, mock } = book();
  const rows = L(h, ledgerRows()).slice(1);
  h.ctx.getOrCreateLedgerSheet().getRange(2, 1, rows.length, 12).setValues(rows);
  h.ctx.refreshHoldings();
  const un = () => h.plain(mock.ss.getSheetByName('Unrecognized')._rows(10));
  assert.deepEqual(un().map(r => r[9]).sort(), ['n1', 'n2', 'n3']);
  assert.equal(un().find(r => r[9] === 'n2')[5], '=HYPERLINK("https://mail.google.com/mail/u/0/#all/n2","Open")');
  assert.ok(mock.ss.getSheetByName('Unrecognized').cfRules.some(r => r.whenTextStartsWith === 'Deposit without amount'));
  // you add the May deposit (HAPI's date): it pairs with its notice and the row goes
  h.ctx.addValuationEntry({ account: 'HAPI', kind: 'Broker', date: '2026-05-20', mode: 'deposit', amount: '59.99', currency: 'USD' });
  assert.deepEqual(un().map(r => r[9]).sort(), ['n1', 'n3']);
});

test('a run reads the HAPI deposit email into a Notice and flags it when no deposit is recorded', () => {
  const { h, mock } = book();
  mock.gmail.threads.push(fakeThread('t-dep', [h.fakeMessage({ subject: 'Deposit Completed 🟢', from: 'Hapi App <no-reply@hapi.trade>',
    body: fixture('investments/hapi_deposit'), date: h.date(2026, 9, 18, 12), id: 'm-dep' })]));
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const ledger = mock.ss.getSheetByName('Investment Ledger')._rows(12);
  assert.deepEqual(h.plain(ledger.map(r => [r[2], r[11]])), [['Notice', 'gmail:m-dep']]);
  const row = mock.ss.getSheetByName('Unrecognized')._rows(10).find(r => r[9] === 'm-dep');
  assert.ok(row, 'flagged');
  assert.match(row[3], /Deposit without amount/);
  // a second run doesn't read the email again, and the flag stays until the deposit is recorded
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.equal(mock.ss.getSheetByName('Investment Ledger')._rows(12).length, 1);
  assert.ok(mock.ss.getSheetByName('Unrecognized')._rows(10).find(r => r[9] === 'm-dep'));
});

test('a Notice turned into the Deposit itself (type changed, amount typed) clears its Unrecognized row (v1.1.45, reported)', () => {
  const { h, mock } = book();
  const led = h.ctx.getOrCreateLedgerSheet();
  led.getRange(2, 1, 2, 12).setValues([
    [h.date(2025, 12, 31), 'HAPI', 'Valuation', '', '', '', 3000, '', 'USD', 'manual', '', ''],
    [h.date(2026, 1, 2), 'HAPI', 'Notice', '', '', '', '', '', 'USD', 'email', 'Deposit completed', 'gmail:n-jan']]);
  h.ctx.refreshHoldings();
  assert.deepEqual(h.plain(mock.ss.getSheetByName('Unrecognized')._rows(10).map(r => r[9])), ['n-jan']);
  const at = led._rows(12).findIndex(r => r[11] === 'gmail:n-jan') + 2;
  led.getRange(at, 3).setValue('Deposit');           // you edit the notice row itself
  led.getRange(at, 7).setValue(250);
  h.ctx.refreshHoldings();
  assert.equal(mock.ss.getSheetByName('Unrecognized')._rows(10).length, 0);
});
