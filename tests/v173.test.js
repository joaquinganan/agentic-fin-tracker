'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

// v1.1.73: the daily summary reports yesterday up to 12:00 AM. Asked: "the daily report should be based only on what
// happened the previous day up to 12 AM" — spending already was (by the transaction's date); investments were not for
// crypto, which trades around the clock: the 6 AM snapshot had six hours of today in its price. Prices invented.

const candleUrl = (pair, iso, isoEnd) => 'https://api.exchange.coinbase.com/products/' + pair +
  '/candles?granularity=3600&start=' + iso + '&end=' + isoEnd;
// Santo Domingo is UTC−4 all year: 12:00 AM on the 21st is 04:00 UTC
const MID_21 = Date.UTC(2026, 9, 21, 4) / 1000, MID_22 = Date.UTC(2026, 9, 22, 4) / 1000;
const candles = (t, open) => JSON.stringify([[t + 3600, 1, 1, 999999, 1, 1], [t, open - 50, open + 50, open, open + 10, 5]]);

function setup() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  h.ctx.getOrCreateLedgerSheet().getRange(2, 1, 1, 12).setValues([[h.date(2026, 10, 1), 'HAPI', 'Buy', 'GOOGL', 2, 150, 300, '', 'USD', 'manual', '', '']]);
  // snapshots taken at 6 AM: GOOGL at its last close, BTC at the spot price of the moment
  h.ctx.recordPortfolioHistory([{ account: 'HAPI', ticker: 'GOOGL', qty: 2, price: 150, value: 300 },
    { account: 'Wallet', ticker: 'BTCUSD', qty: 0.01, price: 80000, value: 800 }], '2026-10-21');
  h.ctx.recordPortfolioHistory([{ account: 'HAPI', ticker: 'GOOGL', qty: 2, price: 160, value: 320 },
    { account: 'Wallet', ticker: 'BTCUSD', qty: 0.01, price: 90000, value: 900 }], '2026-10-22');
  return { mock, h };
}

test('the daily brief values crypto at 12:00 AM of each snapshot; stocks keep their last close', () => {
  const { mock, h } = setup();
  mock.http.responses[candleUrl('BTC-USD', '2026-10-21T04:00:00.000Z', '2026-10-21T05:00:00.000Z')] = { code: 200, body: candles(MID_21, 81000) };
  mock.http.responses[candleUrl('BTC-USD', '2026-10-22T04:00:00.000Z', '2026-10-22T05:00:00.000Z')] = { code: 200, body: candles(MID_22, 84000) };
  const iv = h.plain(h.ctx.investmentsReportData('daily', { rates: { USD: 60 } }));
  assert.equal(iv.total, 320 + 840, 'GOOGL at its close + 0.01 BTC at 84,000, not the 6 AM 90,000');
  assert.equal(iv.change, (320 + 840) - (300 + 810), 'midnight to midnight, not 6 AM to 6 AM');
  const btc = iv.movers.find(m => m.ticker === 'BTCUSD');
  assert.equal(+btc.change.toFixed(4), +(84000 / 81000 - 1).toFixed(4));
  assert.deepEqual(iv.cryptoAt, { repriced: 1, missing: 0 });
  assert.equal(mock.http.requests.filter(u => u.includes('candles')).length, 2, 'one candle per crypto per day compared');
  // Portfolio History itself is not rewritten
  const kept = h.ctx.readHistory().filter(r => r[2] === 'BTCUSD').map(r => r[5]).sort();
  assert.deepEqual(kept, [800, 900]);

  const d = h.ctx.computeDailySummary([new Array(15).fill('h')], { today: h.date(2026, 10, 22, 8), rates: { USD: 60, EUR: 64, COP: 0.015 },
    netIncomeDop: 80000, cards: [], investments: iv });
  const html = h.ctx.buildDailySummaryEmail(d, { sheetName: 'T' }).html.replace(/&#(\d+);/g, (a, n) => String.fromCodePoint(Number(n)));
  assert.ok(html.includes('Crypto at its 12:00 AM price (end of Wed 21 Oct), stocks at their last close.'), 'says what the values mean');
});

test('without a 12:00 AM price, crypto keeps the update\'s price and the email says so', () => {
  const { mock, h } = setup();   // no candle responses: Coinbase down or the coin is not on the exchange
  const iv = h.plain(h.ctx.investmentsReportData('daily', { rates: { USD: 60 } }));
  assert.equal(iv.total, 320 + 900);
  assert.deepEqual(iv.cryptoAt, { repriced: 0, missing: 1 });
  const d = h.ctx.computeDailySummary([new Array(15).fill('h')], { today: h.date(2026, 10, 22, 8), rates: { USD: 60, EUR: 64, COP: 0.015 },
    netIncomeDop: 80000, cards: [], investments: iv });
  const html = h.ctx.buildDailySummaryEmail(d, { sheetName: 'T' }).html.replace(/&#(\d+);/g, (a, n) => String.fromCodePoint(Number(n)));
  assert.ok(html.includes('1 crypto position(s) without a 12:00 AM price'));
  assert.ok(!html.includes('Crypto at its 12:00 AM price'));
});

test('a candle that does not start at 12:00 AM is not taken as the midnight price', () => {
  const { mock, h } = setup();
  const wrong = JSON.stringify([[MID_22 + 3600, 1, 1, 84000, 1, 1]]);   // only the 1:00 AM candle came back
  mock.http.responses[candleUrl('BTC-USD', '2026-10-22T04:00:00.000Z', '2026-10-22T05:00:00.000Z')] = { code: 200, body: wrong };
  const prices = h.plain(h.ctx.fetchCryptoMidnightPrices(['BTCUSD', 'GOOGL'], ['2026-10-22']));
  assert.deepEqual(prices, { '2026-10-22': {} });
});

test('no crypto: no calls to Coinbase and no note', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  h.ctx.getOrCreateLedgerSheet().getRange(2, 1, 1, 12).setValues([[h.date(2026, 10, 1), 'HAPI', 'Buy', 'GOOGL', 2, 150, 300, '', 'USD', 'manual', '', '']]);
  h.ctx.recordPortfolioHistory([{ account: 'HAPI', ticker: 'GOOGL', qty: 2, price: 150, value: 300 }], '2026-10-21');
  h.ctx.recordPortfolioHistory([{ account: 'HAPI', ticker: 'GOOGL', qty: 2, price: 160, value: 320 }], '2026-10-22');
  const iv = h.plain(h.ctx.investmentsReportData('daily', { rates: { USD: 60 } }));
  assert.equal(iv.cryptoAt, null);
  assert.equal(mock.http.requests.length, 0);
});

test('spending in the daily summary counts only yesterday, up to 12:00 AM (the transaction\'s date, not when it was read)', () => {
  const h = load();
  const rows = [new Array(14).fill('h')].concat([[h.date(2026, 10, 21, 23, 50), 100], [h.date(2026, 10, 22, 0, 30), 5000],
    ['2026-10-22', 7000]].map(([date, amount]) => {
    const r = new Array(14).fill('');
    r[0] = date; r[1] = 'LAFISE'; r[2] = 'X'; r[3] = amount; r[4] = 'DOP'; r[5] = 'Restaurants & cafés'; r[11] = 'Transaction';
    return r;
  }));
  const s = h.plain(h.ctx.computeDailySummary(rows, { today: h.date(2026, 10, 22, 8), rates: { USD: 60, EUR: 64, COP: 0.015 }, netIncomeDop: 0, cards: [] }));
  assert.equal(s.spent, 100);
  assert.equal(s.mtd, 100);
});
