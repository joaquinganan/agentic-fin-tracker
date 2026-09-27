'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

const HAPI = 'Hapi App <no-reply@hapi.trade>';
const hapi = (h, name, subject, date, id) =>
  h.fakeMessage({ subject, from: HAPI, body: fixture('investments/' + name), date, id: id || name });

test('HAPI emails: orders (with the per-order fee), dividends; deposit confirmations skipped', () => {
  const h = load();
  const buy = h.plain(h.ctx.parseHapiMessage(hapi(h, 'hapi_order_buy', '✅ Order Executed ', h.date(2026, 9, 8, 9, 43))));
  assert.equal(buy.kind, 'event');
  assert.deepEqual([buy.event.type, buy.event.ticker, buy.event.qty, buy.event.price, buy.event.amount, buy.event.fee, buy.event.currency],
    ['Buy', 'GOOGL', 1.25, 320, 400, 0.15, 'USD']);
  assert.equal(buy.event.id, 'gmail:hapi_order_buy');
  const sell = h.plain(h.ctx.parseHapiMessage(hapi(h, 'hapi_order_sell', '✅ Order Executed ', h.date(2026, 6, 18))));
  assert.deepEqual([sell.event.type, sell.event.ticker, sell.event.qty, sell.event.amount, sell.event.fee], ['Sell', 'UNH', 0.4, 164.2, 0.15]);
  const div = h.plain(h.ctx.parseHapiMessage(hapi(h, 'hapi_dividend', '💸 Dividend from UNH received!', h.date(2026, 8, 18, 8))));
  assert.deepEqual([div.event.type, div.event.ticker, div.event.amount], ['Dividend', 'UNH', 1.95]);
  assert.equal(new Date(div.event.date).getDate(), 18, 'the payment date in the email');
  const dep = h.plain(h.ctx.parseHapiMessage(hapi(h, 'hapi_deposit', 'Deposit Completed 🟢', h.date(2026, 9, 8))));
  assert.equal(dep.kind, 'skipped');
  assert.match(dep.reason, /no amount/);
});

test('HAPI emails: HTML-only bodies are read; inconsistent or unfinished orders are never saved', () => {
  const h = load();
  const text = fixture('investments/hapi_order_buy');
  const htmlOnly = { getSubject: () => '✅ Order Executed ', getFrom: () => HAPI, getPlainBody: () => '', getDate: () => h.date(2026, 9, 8),
    getId: () => 'h1', getBody: () => '<table>' + text.split('\n').map(l => '<tr><td><p>' + l + '</p></td></tr>').join('') + '</table>' };
  assert.equal(h.plain(h.ctx.parseHapiMessage(htmlOnly)).event.qty, 1.25);
  const tampered = hapi(h, 'hapi_order_buy', '✅ Order Executed ', h.date(2026, 9, 8));
  const body = fixture('investments/hapi_order_buy').replace('US$ 400', 'US$ 450');
  tampered.getPlainBody = () => body;
  const bad = h.plain(h.ctx.parseHapiMessage(tampered));
  assert.equal(bad.kind, 'unparsed');
  assert.match(bad.reason, /does not match the cost/);
  const pending = hapi(h, 'hapi_order_buy', '✅ Order Executed ', h.date(2026, 9, 8));
  const pendingBody = fixture('investments/hapi_order_buy').replace('Order completed', 'Order pending');
  pending.getPlainBody = () => pendingBody;
  assert.match(h.plain(h.ctx.parseHapiMessage(pending)).reason, /status "Order pending"/);
});

// ---- positions
function ledger(h, rows) {
  return [['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id']]
    .concat(rows.map(r => { const x = r.slice(); x[0] = h.date(...x[0]); while (x.length < 12) x.push(''); return x; }));
}

test('positions: latest snapshot is the base, later movements apply, average cost and realized P/L', () => {
  const h = load();
  const rows = ledger(h, [
    [[2026, 5, 1], 'HAPI', 'Snapshot', 'GOOGL', 9, 100, 900],               // an older snapshot: ignored
    [[2026, 6, 10], 'HAPI', 'Buy', 'GOOGL', 5, 100, 500, 0.15, 'USD'],       // before the latest snapshot: history only
    [[2026, 8, 31], 'HAPI', 'Snapshot', 'GOOGL', 2, 300, 500],
    [[2026, 8, 31], 'HAPI', 'Snapshot', 'UNH', 1, 400, 380],
    [[2026, 8, 31], 'HAPI', 'Snapshot', 'CASH', '', '', 10],
    [[2026, 8, 31], 'HAPI', 'Buy', 'GOOGL', 1, 300, 300, 0.15, 'USD'],       // same day as the snapshot: already in it
    [[2026, 8, 18], 'HAPI', 'Dividend', 'UNH', '', '', 1.95, '', 'USD'],
    [[2026, 9, 5], 'HAPI', 'Deposit', '', '', '', 5950, '', 'DOP'],
    [[2026, 9, 8], 'HAPI', 'Buy', 'GOOGL', 1.25, 320, 400, 0.15, 'USD'],
    [[2026, 9, 10], 'HAPI', 'Sell', 'UNH', 0.4, 410.5, 164.2, 0.15, 'USD'],
    [[2026, 9, 20], 'IBKR', 'Buy', 'VOO', 1, 500, 500, 0, 'USD'],            // an account with no snapshot starts from zero
    [[2026, 7, 31], 'BHD Fondo', 'Valuation', '', 40, 2262.1, '', '', 'DOP'],
    [[2026, 8, 31], 'BHD Fondo', 'Valuation', '', 40, 2277.6, '', '', 'DOP'],
  ]);
  const r = h.plain(h.ctx.computeHoldings(rows, { usdRate: 59.5, year: 2026 }));
  const p = Object.fromEntries(r.positions.map(x => [x.account + ' ' + x.ticker, x]));
  assert.deepEqual([p['HAPI GOOGL'].qty, p['HAPI GOOGL'].cost], [3.25, 900]);
  assert.deepEqual([p['HAPI UNH'].qty, p['HAPI UNH'].cost], [0.6, 228]);
  assert.equal(+p['HAPI UNH'].realized.toFixed(2), 12.2, 'sold 0.4 at 410.50 with an average cost of 380');
  assert.equal(p['HAPI UNH'].dividends, 1.95);
  assert.equal(p['HAPI GOOGL'].lastPrice, 320, 'last known price: the latest trade');
  assert.deepEqual([p['IBKR VOO'].qty, p['IBKR VOO'].cost], [1, 500]);
  const cash = Object.fromEntries(r.cash.map(c => [c.account, c]));
  assert.equal(cash.HAPI.amount, +(10 + 5950 / 59.5 - 400.15 + 164.05).toFixed(2));
  assert.equal(cash.HAPI.estimated, true, 'a DOP deposit makes the cash an estimate');
  assert.deepEqual(r.valuations.map(v => [v.account, +v.value.toFixed(2), v.day]), [['BHD Fondo', 91104, '2026-08-31']]);
  assert.equal(+r.totals.dividendsYtd.toFixed(2), 1.95);
  assert.equal(+r.totals.feesYtd.toFixed(2), 0.6, 'every order fee of the year, before or after the snapshot');
  assert.equal(r.totals.contributionsYtd.DOP, 5950);
  assert.deepEqual(r.warnings, []);
});

test('positions: a sale with no position is flagged, never turned into negative shares', () => {
  const h = load();
  const r = h.plain(h.ctx.computeHoldings(ledger(h, [[[2026, 9, 1], 'HAPI', 'Sell', 'TSLA', 1, 400, 400, 0.15, 'USD']]), { usdRate: 60, year: 2026 }));
  assert.equal(r.positions.length, 0);
  assert.match(r.warnings[0], /HAPI TSLA: sale on 2026-09-01 with no position — add a Snapshot/);
});

test('prices: GOOGLEFINANCE with a sanity check against the last known price; crypto symbols mapped', () => {
  const h = load();
  const f = h.plain(h.ctx.holdingPriceFormulas('ETHUSD', 2700));
  assert.equal(f.price, '=IFERROR(IF(ABS(GOOGLEFINANCE("CURRENCY:ETHUSD","price")/2700-1)>0.5,2700,GOOGLEFINANCE("CURRENCY:ETHUSD","price")),2700)');
  assert.match(f.source, /check symbol/);
  assert.equal(h.plain(h.ctx.holdingPriceFormulas('XYZ', 0)).price, '=IFERROR(GOOGLEFINANCE("XYZ","price"),0)');
});

// ---- the whole run
function setup() {
  const mock = makeServices();
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ LAFISE: true, BANESCO: true, BHD: true, POPULAR: true })], ['setupDate', '2026-09-01']]
    .forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions');
  h.ctx.initializeTransactionsSheet();
  const cr = mock.ss.insertSheet('Custom Rules');
  [['UserEmail', 'Category', 'Keyword', 'Timestamp'], ['user@example.com', 'Exclude', 'OUROSR SRL', '']].forEach(r => cr.appendRow(r));
  // a transfer to HAPI's collection account, already in Transactions
  const tx = new Array(14).fill('');
  tx[0] = h.date(2026, 9, 5); tx[1] = 'BANESCO'; tx[2] = 'OUROSR SRL'; tx[3] = 5950; tx[4] = 'DOP'; tx[5] = 'Exclude'; tx[11] = 'Transfer'; tx[12] = 'bank-msg-1';
  mock.ss.getSheetByName('Transactions').getRange(2, 1, 1, 14).setValues([tx]);
  // the positions as HAPI showed them on Aug 31
  const ledgerSheet = h.ctx.getOrCreateLedgerSheet();
  ledgerSheet.getRange(2, 1, 3, 12).setValues([
    [h.date(2026, 8, 31), 'HAPI', 'Snapshot', 'GOOGL', 2, 300, 500, '', 'USD', 'manual', 'portfolio screen', ''],
    [h.date(2026, 8, 31), 'HAPI', 'Snapshot', 'UNH', 1, 400, 380, '', 'USD', 'manual', 'portfolio screen', ''],
    [h.date(2026, 8, 31), 'HAPI', 'Snapshot', 'CASH', '', '', 10, '', 'USD', 'manual', 'portfolio screen', '']]);
  mock.gmail.threads.push(
    fakeThread('t-hapi-buy', [hapi(h, 'hapi_order_buy', '✅ Order Executed ', h.date(2026, 9, 8, 9, 43), 'm-buy')]),
    fakeThread('t-hapi-sell', [hapi(h, 'hapi_order_sell', '✅ Order Executed ', h.date(2026, 9, 10, 9, 41), 'm-sell')]),
    fakeThread('t-hapi-div', [hapi(h, 'hapi_dividend', '💸 Dividend from UNH received!', h.date(2026, 8, 18, 8), 'm-div')]),
    fakeThread('t-hapi-dep', [hapi(h, 'hapi_deposit', 'Deposit Completed 🟢', h.date(2026, 9, 6), 'm-dep')]));
  return { h, mock };
}

test('a run: HAPI emails and the bank deposit reach the ledger, Holdings is rebuilt, nothing is counted twice', () => {
  const { h, mock } = setup();
  h.ctx.buildOrRefreshDashboard();
  h.ctx.runGmailMonitorForDateRange('2026-01-01', '2026-12-31');
  assert.deepEqual(h.logs.filter(l => /Error during investments|Could not style/.test(l)), []);
  const rows = mock.ss.getSheetByName('Investment Ledger')._rows(12);
  const kinds = rows.map(r => r[2] + ' ' + (r[3] || r[1]) + ' ' + r[9]).sort();
  assert.deepEqual(kinds, ['Buy GOOGL email', 'Deposit HAPI bank', 'Dividend UNH email', 'Sell UNH email',
    'Snapshot CASH manual', 'Snapshot GOOGL manual', 'Snapshot UNH manual']);
  assert.ok(mock.gmail.queries.some(q => q.startsWith('from:no-reply@hapi.trade after:')), 'searched by sender and range');
  assert.ok(mock.gmail.markedRead.includes('t-hapi-buy') && mock.gmail.markedRead.includes('t-hapi-dep'), 'read → processed');
  assert.equal(mock.ss.getSheetByName('Transactions')._rows(14).length, 1, 'HAPI emails never become spending rows');
  // Holdings: formulas with live prices, and the position numbers from the ledger
  const hold = mock.ss.getSheetByName('Holdings');
  const cells = [...hold.cells.values()].map(String);
  assert.ok(cells.some(c => c.includes('GOOGLEFINANCE("GOOGL","price")/320-1)>0.5')), 'GOOGL priced live, last known 320');
  const values = hold._rows(13);
  const googl = values.find(r => r[2] === 'GOOGL');
  assert.deepEqual([googl[1], googl[3], googl[5]], ['HAPI', 3.25, 900]);
  assert.ok(cells.some(c => /RATE_USD/.test(c)), 'DOP-equivalent through the Dashboard rate');
  assert.ok(mock.ss.getRangeByName('RATE_USD'), 'that named range exists');
  // a second run adds nothing
  const before = rows.length;
  h.ctx.runGmailMonitorForDateRange('2026-01-01', '2026-12-31');
  assert.equal(mock.ss.getSheetByName('Investment Ledger')._rows(12).length, before, 'emails and deposits are not saved twice');
});

test('run summary reports the investments step', () => {
  const h = load();
  const text = h.ctx.buildRunSummary({ search: { capped: false }, threads: [], transactions: [], recatChanged: 0, errors: [],
    results: { success: 0, failed: 0, duplicates: 0 }, marked: null,
    stats: { messagesSeen: 0, promotional: 0, nonTransactional: 0, declined: 0, notOwnBank: 0, amountNotFound: 0, parseErrors: 0, placeholders: 0, reversals: 0 },
    investments: { parsed: 3, deposits: 1, saved: 4, duplicates: 0, unparsed: 1, warnings: ['HAPI TSLA: sale with no position'] } });
  assert.match(text, /📈 Investments: 4 new ledger row\(s\) \(3 from broker emails, 1 deposit\(s\) from bank transfers\)/);
  assert.match(text, /Broker emails not read: 1 — left UNREAD/);
  assert.match(text, /Holdings: HAPI TSLA: sale with no position/);
});
