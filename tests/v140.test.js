'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

const L = (h, rows) => [['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id']]
  .concat(rows.map(r => { const x = r.slice(); x[0] = h.date(...x[0]); while (x.length < 12) x.push(''); return x; }));
// fictitious: US$5,000 on Dec 31; US$1,000 deposited; US$3,000 bought and US$500 sold since
const hapi = h => [
  [[2025, 12, 31], 'HAPI', 'Valuation', '', '', '', 5000, '', 'USD'],
  [[2026, 2, 1], 'HAPI', 'Deposit', '', '', '', 1000, '', 'USD', 'bank'],
  [[2026, 3, 1], 'HAPI', 'Buy', 'GOOGL', 5, 400, 2000, 0.15, 'USD', 'email'],
  [[2026, 5, 1], 'HAPI', 'Buy', 'NVDA', 5, 200, 1000, 0.15, 'USD', 'email'],
  [[2026, 6, 1], 'HAPI', 'Sell', 'NVDA', 2, 250, 500, 0.15, 'USD', 'email'],
  [[2026, 9, 26], 'HAPI', 'Snapshot', 'CASH', '', '', 10, '', 'USD']];

test('missing deposits are flagged: far more bought since the start than deposited (reported: gain included purchases)', () => {
  const h = load();
  const r = h.plain(h.ctx.computeReturns(L(h, hapi(h)), { HAPI: 8400 }, { usdRate: 60, today: '2026-09-28' })).accounts[0];
  assert.equal(+r.netBought.toFixed(2), 2500.45, '3,000 bought + fees − 500 sold');
  assert.equal(+r.unfunded.toFixed(2), 1500.45);
  // once the missing deposits are recorded, no warning — and the gain drops accordingly
  const covered = hapi(h).concat([[[2026, 4, 1], 'HAPI', 'Deposit', '', '', '', 1500, '', 'USD', 'manual']]);
  const r2 = h.plain(h.ctx.computeReturns(L(h, covered), { HAPI: 8400 }, { usdRate: 60, today: '2026-09-28' })).accounts[0];
  assert.equal(r2.unfunded, 0);
  assert.equal(r2.gain, r.gain - 1500);
});

test('small differences are not flagged (cash already in the account, rounding)', () => {
  const h = load();
  const rows = hapi(h).map(x => (x[2] === 'Deposit' ? [x[0], x[1], x[2], '', '', '', 2480, '', 'USD'] : x));
  assert.equal(h.plain(h.ctx.computeReturns(L(h, rows), { HAPI: 8400 }, { usdRate: 60, today: '2026-09-28' })).accounts[0].unfunded, 0);
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

test('Performance shows the warning with the figures', () => {
  const { h, mock } = book();
  const rows = L(h, hapi(h)).slice(1);
  h.ctx.getOrCreateLedgerSheet().getRange(2, 1, rows.length, 12).setValues(rows);
  h.ctx.refreshHoldings();
  const note = [...mock.ss.getSheetByName('Holdings').cells.values()].map(String).find(v => v.startsWith('⚠️ HAPI: purchases since 2025-12-31'));
  assert.ok(note, 'warning written under Performance');
  assert.match(note, /exceed the deposits recorded \(US\$1000\.00\) by US\$1500\.45/);
  assert.match(note, /Add Balance or Deposit/);
});

test('a deposit from the dialog is money put in, not gain; the same one twice is refused', () => {
  const { h, mock } = book();
  const dep = { account: 'HAPI', kind: 'Broker', date: '2026-04-01', mode: 'deposit', amount: '1500', currency: 'USD', notes: 'from Airtm' };
  assert.equal(h.ctx.validateValuationEntry(dep, '2026-09-28'), null);
  assert.equal(h.ctx.addValuationEntry(dep), true);
  assert.match(mock.ui.alerts.pop(), /Deposit saved[\s\S]*HAPI — USD 1,500.00 on 2026-04-01[\s\S]*money put in, not as gain/);
  const row = h.plain(mock.ss.getSheetByName('Investment Ledger')._rows(12)[0]);
  assert.deepEqual(row.slice(1, 11), ['HAPI', 'Deposit', '', '', '', 1500, '', 'USD', 'manual', 'from Airtm']);
  assert.equal(h.ctx.addValuationEntry(dep), false);
  assert.match(mock.ui.alerts.pop(), /already has a deposit of USD 1500 on 2026-04-01/);
  assert.equal(mock.ss.getSheetByName('Investment Ledger')._rows(12).length, 1);
  // a second deposit the same day of a different amount is fine
  assert.equal(h.ctx.addValuationEntry(Object.assign({}, dep, { amount: '200' })), true);
  assert.match(h.ctx.validateValuationEntry(Object.assign({}, dep, { amount: '0' }), '2026-09-28'), /deposit must be greater than 0/);
  h.ctx.openValuationDialog();
  assert.ok(mock.ui.dialogs.pop().html.includes('value="deposit"'));
});
