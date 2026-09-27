'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

const HAPI = 'Hapi App <no-reply@hapi.trade>';
const order = (h, body, id) => h.fakeMessage({ subject: '✅ Order Executed ', from: HAPI, body, date: h.date(2026, 7, 6, 17, 55), id: id || 'lim' });

test('HAPI limit orders: Cost includes the fee — cost basis is quantity × price (from a live email)', () => {
  const h = load();
  const r = h.plain(h.ctx.parseHapiMessage(order(h, fixture('investments/hapi_order_limit_buy'))));
  assert.equal(r.kind, 'event', r.reason);
  assert.deepEqual([r.event.type, r.event.ticker, r.event.qty, r.event.price, r.event.amount, r.event.fee, r.event.notes],
    ['Buy', 'MSFT', 2, 150.25, 300.5, 2.99, 'Limit order']);
  // a market order keeps its own Cost (fee apart), and a cost matching neither rule is still refused
  const market = h.plain(h.ctx.parseHapiMessage(order(h, fixture('investments/hapi_order_buy'))));
  assert.deepEqual([market.event.amount, market.event.fee, market.event.notes], [400, 0.15, '']);
  const wrong = h.plain(h.ctx.parseHapiMessage(order(h, fixture('investments/hapi_order_limit_buy').replace('US$ 303.49', 'US$ 310.00'))));
  assert.match(wrong.reason, /does not match the cost/);
});

test('the limit-order email that was in Unrecognized drops off once read (history before the snapshot)', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"BHD":true}'], ['monthlyIncome', 1000], ['incomeCurrency', 'USD']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  h.ctx.recordUnrecognized([{ id: 'lim', date: h.date(2026, 7, 6), bank: 'HAPI', subject: 'Order Executed', reason: 'Broker email not read', snippet: '' }], [], h.date(2026, 9, 27));
  mock.gmail.threads.push(fakeThread('t-lim', [order(h, fixture('investments/hapi_order_limit_buy'))]));
  h.ctx.runGmailMonitorForDateRange('2026-07-06', '2026-07-06');
  assert.equal(mock.ss.getSheetByName('Unrecognized')._rows(10).length, 0);
  const buy = mock.ss.getSheetByName('Investment Ledger')._rows(12).find(r => r[2] === 'Buy');
  assert.deepEqual([buy[3], buy[4], buy[6], buy[7], buy[10]], ['MSFT', 2, 300.5, 2.99, 'Limit order']);
});

function tabs(names) {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  names.forEach(n => mock.ss.insertSheet(n));
  return { h, mock, order: () => mock.ss.getSheets().map(s => s.getName()),
    hidden: () => mock.ss.getSheets().filter(s => s.isSheetHidden()).map(s => s.getName()).sort() };
}

test('reordering keeps hidden tabs hidden, and moves nothing when the order is already right', () => {
  const t = tabs(['Configuration', 'Holdings', 'Dashboard', 'Categories', 'Transactions', 'Investment Accounts']);
  ['Configuration', 'Categories', 'Investment Accounts'].forEach(n => t.mock.ss.getSheetByName(n).hideSheet());
  t.h.ctx.ensureSheetOrder();
  assert.deepEqual(t.order(), ['Dashboard', 'Transactions', 'Holdings', 'Investment Accounts', 'Configuration', 'Categories']);
  assert.deepEqual(t.hidden(), ['Categories', 'Configuration', 'Investment Accounts'], 'hidden again after moving');
  const before = t.mock.ss.activations;
  t.h.ctx.ensureSheetOrder();
  assert.ok(t.mock.ss.activations - before <= 1, 'already in order: nothing moved');
  assert.deepEqual(t.hidden(), ['Categories', 'Configuration', 'Investment Accounts']);
});

test('the Recategorize menu item reorders the tabs (reported: tabs unchanged after recategorizing)', () => {
  const t = tabs(['Configuration', 'Custom Rules', 'Transactions', 'Holdings', 'Dashboard']);
  [['Key', 'Value'], ['email', 'user@example.com']].forEach(r => t.mock.ss.getSheetByName('Configuration').appendRow(r));
  t.mock.ss.getSheetByName('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  t.h.ctx.initializeTransactionsSheet();
  t.h.ctx.recategorizeAllTransactionsPrompt();
  assert.deepEqual(t.order(), ['Dashboard', 'Transactions', 'Bank Transfers', 'Holdings', 'Custom Rules', 'Configuration'],
    'recategorize also builds Bank Transfers, which goes after Transactions');
});
