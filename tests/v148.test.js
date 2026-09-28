'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

// fictitious portfolio, laid out as the broker's web screen copies (and as it pastes with links)
const PLAIN = ['Total balance', '$2,010.30', 'Total assets', '$2,000.30', 'Total money', '$10.00', 'My Assets',
  'GOOGL', '2.5', '$800.00', '+$150.00 (+23.08%)', 'SHIBUSD', '50,000.5', '$0.30', '-$0.10 (-25.00%)',
  'AAPL', '5', '$1,200.00', '+$200.00 (+20.00%)'].join('\n');
const LINKED = '[NVDA](https://app.example/asset/NVDA)\n[3.1](https://app.example/asset/NVDA)\n[$600.00](https://app.example/asset/NVDA)\n' +
  '[−$20.00 (−3.23%)](https://app.example/asset/NVDA)';

test('a pasted portfolio screen becomes positions: cost = value − gain, price = value ÷ quantity, cash from Total money', () => {
  const h = load();
  const r = h.plain(h.ctx.parsePortfolioPaste(PLAIN));
  assert.deepEqual(r.positions.map(p => [p.ticker, p.qty, p.value, p.cost]), [['GOOGL', 2.5, 800, 650], ['SHIBUSD', 50000.5, 0.3, 0.4], ['AAPL', 5, 1200, 1000]]);
  assert.equal(r.positions[0].price, 320);
  assert.equal(r.positions[1].price, 0.000005999940001, 'tiny prices keep their precision');
  assert.deepEqual([r.cash, r.totalAssets, r.positionsValue, r.warnings], [10, 2000.3, 2000.3, []], 'matches Total assets');
  const linked = h.plain(h.ctx.parsePortfolioPaste(LINKED));
  assert.deepEqual(linked.positions.map(p => [p.ticker, p.qty, p.value, p.cost]), [['NVDA', 3.1, 600, 620]], 'links stripped; a loss raises the cost');
  assert.equal(h.plain(h.ctx.parsePortfolioPaste('Total money $7.25\n' + LINKED)).cash, 7.25, 'label and amount on one line');
});

test('a position left out of the copy is caught against Total assets (as SHIB was, by hand, before this existed)', () => {
  const h = load();
  const missing = PLAIN.replace('AAPL\n5\n$1,200.00\n+$200.00 (+20.00%)', '');
  const r = h.plain(h.ctx.parsePortfolioPaste(missing));
  assert.equal(r.positions.length, 2);
  assert.match(r.warnings[0], /add up to US\$800\.30 but the screen says US\$2000\.30 — US\$1200\.00 is missing/);
  // cents of rounding across positions are not a problem
  assert.deepEqual(h.plain(h.ctx.parsePortfolioPaste(PLAIN.replace('$2,000.30', '$2,000.37'))).warnings, []);
});

test('what is not a position is not taken as one; duplicates, zero quantities and empty pastes are reported', () => {
  const h = load();
  assert.equal(h.plain(h.ctx.parsePortfolioPaste('USD\nPortfolio\nHISTORY\nReports')).positions.length, 0);
  assert.match(h.plain(h.ctx.parsePortfolioPaste('nothing useful')).warnings[0], /No positions found/);
  const twice = h.plain(h.ctx.parsePortfolioPaste(LINKED + '\n' + LINKED));
  assert.equal(twice.positions.length, 1);
  assert.match(twice.warnings[0], /NVDA appears twice/);
  assert.match(h.plain(h.ctx.parsePortfolioPaste('QQQ\n0\n$0.00\n+$0.00 (0%)')).warnings[0], /QQQ: quantity is 0/);
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

test('saving: a Snapshot row per position and the cash; the same account and day again replaces it; Holdings built', () => {
  const { h, mock } = book();
  const led = h.ctx.getOrCreateLedgerSheet();
  led.getRange(2, 1, 1, 12).setValues([[h.date(2026, 9, 1), 'HAPI', 'Buy', 'GOOGL', 1, 300, 300, 0.15, 'USD', 'email', '', 'gmail:x']]);
  assert.equal(h.ctx.savePortfolioPaste({ account: 'HAPI', date: '2026-09-20', text: PLAIN }), true);
  const rows = () => h.plain(mock.ss.getSheetByName('Investment Ledger')._rows(12));
  const snap = rows().filter(r => r[2] === 'Snapshot');
  assert.deepEqual(snap.map(r => [r[3], r[4], r[6]]).sort(), [['AAPL', 5, 1000], ['CASH', '', 10], ['GOOGL', 2.5, 650], ['SHIBUSD', 50000.5, 0.4]]);
  assert.match(mock.ui.alerts.pop(), /Snapshot saved — HAPI, 2026-09-20[\s\S]*3 position\(s\), US\$2000\.30 \+ cash US\$10\.00/);
  // pasted again (corrected), same day: replaced, not added
  assert.equal(h.ctx.savePortfolioPaste({ account: 'HAPI', date: '2026-09-20', text: LINKED }), true);
  assert.deepEqual(rows().filter(r => r[2] === 'Snapshot').map(r => r[3]), ['NVDA']);
  assert.match(mock.ui.alerts.pop(), /replaced the 4 row\(s\) of that day's earlier snapshot/);
  assert.equal(rows().filter(r => r[2] === 'Buy').length, 1, 'other rows untouched');
  const holdings = h.plain(mock.ss.getSheetByName('Holdings')._rows(13));
  assert.ok(holdings.some(r => r[2] === 'NVDA' && r[3] === 3.1), 'Holdings rebuilt from it');
});

test('saving refuses an empty account, a future date and a paste with no positions', () => {
  const { h, mock } = book();
  assert.equal(h.ctx.savePortfolioPaste({ account: ' ', date: '2026-09-20', text: PLAIN }), false);
  assert.match(mock.ui.alerts.pop(), /give the account a name/);
  assert.equal(h.ctx.savePortfolioPaste({ account: 'HAPI', date: '2999-01-01', text: PLAIN }), false);
  assert.equal(h.ctx.savePortfolioPaste({ account: 'HAPI', date: '2026-09-20', text: 'hello' }), false);
  assert.match(mock.ui.alerts.pop(), /No positions found/);
  assert.equal(mock.ss.getSheetByName('Investment Ledger'), null, 'nothing written');
});

test('Start here and the menu lead to it', () => {
  const { h, mock } = book();
  const s = h.plain(h.ctx.startHereStatus(h.date(2026, 9, 28, 9))).steps.find(x => x.id === 'investments');
  assert.equal(s.action.fn, 'openPortfolioPasteDialog');
  assert.ok(h.plain(h.get('START_HERE_ACTIONS')).includes('openPortfolioPasteDialog'));
  h.ctx.onOpen();
  assert.ok(mock.ui.menus[0].items.some(i => i.fn === 'openPortfolioPasteDialog'));
});
