'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

// Reported: the daily report's investments looked wrong and its "biggest moves" were missing. The history was being
// recorded before GOOGLEFINANCE loaded, so stocks got their last known (ledger) price day after day.

let sheets = 0;
function sourcesSheet(h, mock, values) {
  const s = mock.ss.insertSheet('S' + (++sheets));
  s.getRange(5, 13, values.length, 1).setValues(values.map(v => [v]));
  return s;
}

test('waiting for live prices: stops as soon as they load; gives up early on a symbol that never loads', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const s = sourcesSheet(h, mock, ['last known', 'live', 'Coinbase · 2026-09-29 06:04']);
  let sleeps = 0;
  const live = h.plain(h.ctx.waitForLivePrices(s, 5, 3, { sleep: () => { sleeps++; s.getRange(5, 13).setValue('live'); } }));
  assert.deepEqual(live, [true, true, true]);
  assert.equal(sleeps, 1, 'one check after it loaded');
  const stuck = sourcesSheet(h, mock, ['last known — check symbol', 'last known']);
  let n = 0;
  const r = h.plain(h.ctx.waitForLivePrices(stuck, 5, 2, { sleep: () => { n++; } }));
  assert.deepEqual(r, [false, false]);
  assert.equal(n, h.get('LIVE_PRICE_CALM_POLLS'), 'gave up after the checks without progress, not the whole wait');
  let t = 0, m = 0;
  h.plain(h.ctx.waitForLivePrices(stuck, 5, 2, { sleep: () => { m++; t += 20000; }, clock: () => t, maxMs: 15000 }));
  assert.equal(m, 1, 'never past the time limit');
});

test('a stock without a live price is valued at the history\'s latest price and recorded without a price', () => {
  const h = load();
  const holdings = { positions: [{ account: 'HAPI', ticker: 'GOOGL', qty: 2, lastPrice: 150 }, { account: 'HAPI', ticker: 'NVDA', qty: 1, lastPrice: 90 }],
    cash: [], valuations: [] };
  const rows = h.plain(h.ctx.holdingsValueRows(holdings, [300, 90], { live: [true, false], dayChange: [0.021, ''], prevPrices: { 'HAPI|NVDA': 180 } }));
  assert.deepEqual(rows.map(r => [r.ticker, r.price, r.value, r.dayChange, !!r.stale]),
    [['GOOGL', 150, 300, 0.021, false], ['NVDA', '', 180, '', true]], 'NVDA at yesterday\'s 180, not the snapshot\'s 90');
  const hist = [['Date', 'Account', 'Ticker', 'Quantity', 'Price (US$)', 'Value (US$)'],
    [h.date(2026, 9, 26), 'HAPI', 'NVDA', 1, 170, 170], [h.date(2026, 9, 28), 'HAPI', 'NVDA', 1, 180, 180], [h.date(2026, 9, 29), 'HAPI', 'NVDA', 1, 185, 185]];
  assert.deepEqual(h.plain(h.ctx.latestHistoryPrices(hist, '2026-09-29')), { 'HAPI|NVDA': 180 }, 'the latest day before today');
});

test('the history gains a Day change column — an older 6-column history too', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const old = mock.ss.insertSheet('Portfolio History');
  old.getRange(1, 1, 2, 6).setValues([['Date', 'Account', 'Ticker', 'Quantity', 'Price (US$)', 'Value (US$)'], [h.date(2026, 9, 28), 'HAPI', 'GOOGL', 2, 148, 296]]);
  h.ctx.recordPortfolioHistory([{ account: 'HAPI', ticker: 'GOOGL', qty: 2, price: 150, value: 300, dayChange: 0.0135 }], '2026-09-29');
  const v = h.plain(old.getDataRange().getValues());
  assert.equal(v[0][6], 'Day change');
  assert.deepEqual(v.slice(1).map(r => [h.ctx.normalizeDateForCompare(r[0]), r[2], r[6]]),
    [['2026-09-29', 'GOOGL', 0.0135], ['2026-09-29', '', ''], ['2026-09-28', 'GOOGL', '']]);
});

const HIST = h => [['Date', 'Account', 'Ticker', 'Quantity', 'Price (US$)', 'Value (US$)', 'Day change'],
  // Monday 6 AM vs Sunday 6 AM: both at Friday's close — the prices didn't move, the last session did
  [h.date(2026, 9, 27), 'HAPI', 'GOOGL', 2, 150, 300, 0.012], [h.date(2026, 9, 27), 'HAPI', 'NVDA', 1, 180, 180, -0.031],
  [h.date(2026, 9, 27), 'HAPI', 'ETHUSD', 0.5, 4000, 2000, ''], [h.date(2026, 9, 27), 'TOTAL', '', '', '', 2480, ''],
  [h.date(2026, 9, 28), 'HAPI', 'GOOGL', 2, 150, 300, 0.012], [h.date(2026, 9, 28), 'HAPI', 'NVDA', 1, 180, 180, -0.031],
  [h.date(2026, 9, 28), 'HAPI', 'ETHUSD', 0.5, 4100, 2050, ''], [h.date(2026, 9, 28), 'HAPI', 'AMZN', 1, '', 220, ''],
  [h.date(2026, 9, 28), 'TOTAL', '', '', '', 2750, '']];

test('biggest moves come from the last session (a Monday shows Friday\'s), crypto from its price; unpriced positions are counted', () => {
  const h = load();
  const L = [['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id']];
  const b = h.plain(h.ctx.investmentsDailyBrief(HIST(h), L, { usdRate: 60 }));
  assert.deepEqual(b.movers.map(m => [m.ticker, +m.change.toFixed(4), !!m.session]), [['NVDA', -0.031, true], ['ETHUSD', 0.025, false], ['GOOGL', 0.012, true]]);
  assert.equal(b.stale, 1, 'AMZN had no live price');
});

test('the daily email says "last session" and when prices were not live', () => {
  const h = load();
  const L = [['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id']];
  const iv = h.ctx.investmentsDailyBrief(HIST(h), L, { usdRate: 60 });
  const d = h.ctx.computeDailySummary([new Array(15).fill('h')], { today: h.date(2026, 9, 29, 8), rates: { USD: 60, EUR: 64, COP: 0.015 },
    netIncomeDop: 80000, cards: [], investments: iv });
  const decode = x => x.replace(/&#(\d+);/g, (a, n) => String.fromCodePoint(Number(n)));
  const html = decode(h.ctx.buildDailySummaryEmail(d, { sheetName: 'T' }).html);
  assert.ok(html.includes('HAPI · last session'));
  assert.ok(html.includes('1 position(s) had no live price at the update'));
});

test('a refresh records the price GOOGLEFINANCE loads, not the ledger\'s it shows while loading (the reported bug)', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"BHD":true}'], ['monthlyIncome', 1000], ['incomeCurrency', 'USD']].forEach(r => cfg.appendRow(r));
  h.ctx.getOrCreateLedgerSheet().getRange(2, 1, 1, 12).setValues([[h.date(2026, 9, 1), 'HAPI', 'Snapshot', 'GOOGL', 2, 150, 300, '', 'USD', 'manual', '', '']]);
  // while it waits, GOOGLEFINANCE finishes loading: 170 instead of the ledger's 150
  h.get('Utilities').sleep = () => {
    const s = mock.ss.getSheetByName('Holdings');
    const row = s.getDataRange().getValues().findIndex(r => r[2] === 'GOOGL') + 1;   // row 1 is index 0
    s.getRange(row, 8).setValue(340); s.getRange(row, 13).setValue('live');
  };
  h.ctx.refreshHoldings();
  const today = h.ctx.normalizeDateForCompare(new Date());
  const googl = h.plain(mock.ss.getSheetByName('Portfolio History').getDataRange().getValues())
    .find(r => r[2] === 'GOOGL' && h.ctx.normalizeDateForCompare(r[0]) === today);
  assert.deepEqual([googl[4], googl[5]], [170, 340]);
});
