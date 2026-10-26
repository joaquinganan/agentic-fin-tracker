'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');

// v1.1.75: reported — the email of Sunday the 25th ("23 → 24 Oct", a Saturday) listed three stocks with Friday's
// session moves while the day's change counted only crypto. On a day without a market session only crypto moves.
// Numbers invented, same shape.

const L = [['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id']];
const hist = (h, rows) => [['Date', 'Account', 'Ticker', 'Quantity', 'Price (US$)', 'Value (US$)', 'Day change']]
  .concat(rows.map(r => { const x = r.slice(); x[0] = h.date(...x[0]); while (x.length < 7) x.push(''); return x; }));
// the 6 AM snapshot of day d: GOOGL and UNH at their close, with that session's day change; BTC by its price
const day = (d, googl, gdc, unh, udc, btc) => [[[2026, 10, d], 'HAPI', 'GOOGL', 2, googl, 2 * googl, gdc],
  [[2026, 10, d], 'HAPI', 'UNH', 1, unh, unh, udc], [[2026, 10, d], 'Wallet', 'BTCUSD', 0.01, btc, btc / 100],
  [[2026, 10, d], 'TOTAL', '', '', '', 2 * googl + unh + btc / 100]];
const emailOf = (h, iv) => {
  const s = h.ctx.computeDailySummary([new Array(15).fill('h')], { today: h.date(2026, 10, 25, 8), rates: { USD: 60, EUR: 64, COP: 0.015 },
    netIncomeDop: 80000, cards: [], investments: iv });
  const mail = h.ctx.buildDailySummaryEmail(s, { sheetName: 'T' });
  return mail.html.replace(/&#(\d+);/g, (a, n) => String.fromCodePoint(Number(n)));
};

// Saturday's and Sunday's snapshots both hold Friday's close and Friday's day change; only BTC moved (−1%)
const WEEKEND = h => hist(h, day(24, 160, -0.0124, 300, -0.0162, 100000).concat(day(25, 160, -0.0124, 300, -0.0162, 99000)));

test('on a weekend the biggest moves are crypto only, as the change (reported)', () => {
  const h = load();
  const b = h.plain(h.ctx.investmentsDailyBrief(WEEKEND(h), L, {}));
  assert.equal(+b.change.toFixed(2), -10, 'only BTC: 0.01 × −1,000');
  assert.deepEqual(b.movers.map(m => m.ticker), ['BTCUSD'], 'not Friday\'s GOOGL and UNH');
  assert.equal(b.stocksClosed, true);
  const html = emailOf(h, b);
  assert.ok(!html.includes('last session'), 'no stock listed with its last session');
  assert.ok(html.includes('Stocks unchanged (no market session).'));
});

test('the crypto note says the stocks did not move on a weekend, and their last close on a weekday', () => {
  const h = load();
  const weekend = h.ctx.investmentsDailyBrief(WEEKEND(h), L, {});
  weekend.cryptoAt = { repriced: 1, missing: 0 };
  assert.ok(emailOf(h, weekend).includes('Crypto at its 12:00 AM price (end of Sat 24 Oct), stocks unchanged (no market session).'));
  const weekday = h.ctx.investmentsDailyBrief(hist(h, day(23, 158, 0.004, 305, 0.01, 100000)
    .concat(day(24, 160, 0.0127, 300, -0.0164, 99000))), L, {});
  weekday.cryptoAt = { repriced: 1, missing: 0 };
  const html = emailOf(h, weekday);
  assert.ok(html.includes('stocks at their last close.'));
  assert.ok(!html.includes('no market session'));
});

test('on a trading day the stocks\' sessions are listed, as before', () => {
  const h = load();
  const b = h.plain(h.ctx.investmentsDailyBrief(hist(h, day(23, 158, 0.004, 305, 0.01, 100000)
    .concat(day(24, 160, 0.0127, 300, -0.0164, 99000))), L, {}));
  assert.deepEqual(b.movers.map(m => [m.ticker, !!m.session]), [['UNH', true], ['GOOGL', true], ['BTCUSD', false]]);
  assert.equal(b.stocksClosed, false);
});

test('a holiday for one stock\'s market but not another\'s: only the one whose price moved is listed', () => {
  const h = load();
  const b = h.plain(h.ctx.investmentsDailyBrief(hist(h, day(23, 158, 0.004, 300, -0.0164, 100000)
    .concat(day(24, 160, 0.0127, 300, -0.0164, 100000))), L, {}));
  assert.deepEqual(b.movers.map(m => m.ticker), ['GOOGL']);
  assert.equal(b.stocksClosed, false);
});
