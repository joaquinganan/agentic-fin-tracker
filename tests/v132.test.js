'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

const L = (h, rows) => [['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id']]
  .concat(rows.map(r => { const x = r.slice(); x[0] = h.date(...x[0]); while (x.length < 12) x.push(''); return x; }));
const near = (a, b, eps = 1e-6) => assert.ok(Math.abs(a - b) < eps, `${a} ≈ ${b}`);

test('XIRR: 1,000 → 1,100 in a year is 10%; a mid-year deposit is weighted by time', () => {
  const h = load();
  near(h.ctx.xirr([{ day: '2025-01-01', amount: -1000 }, { day: '2026-01-01', amount: 1100 }]), 0.1, 1e-6);
  const r = h.ctx.xirr([{ day: '2025-01-01', amount: -1000 }, { day: '2025-07-02', amount: -1000 }, { day: '2026-01-01', amount: 2150 }]);
  assert.ok(r > 0.09 && r < 0.11, 'about 10% a year: +150 on 1,000 for a year and 1,000 for half a year');
  assert.equal(h.ctx.xirr([{ day: '2025-01-01', amount: 1000 }]), null, 'needs money in and out');
});

test('returns since tracking began: snapshot value, deposits (DOP converted), Modified Dietz; annualized only after 180 days', () => {
  const h = load();
  const ledger = L(h, [
    [[2026, 9, 1], 'HAPI', 'Snapshot', 'GOOGL', 2, 300, 500, '', 'USD'],     // start value 2 × 300 + 100 cash = 700
    [[2026, 9, 1], 'HAPI', 'Snapshot', 'CASH', '', '', 100, '', 'USD'],
    [[2026, 9, 11], 'HAPI', 'Deposit', '', '', '', 6000, '', 'DOP'],         // 100 US$ at 60, 20 of 30 days invested
    [[2026, 8, 31], 'Fondo', 'Valuation', '', 40, 2250, '', '', 'DOP'],       // 90,000 DOP = 1,500 US$
    [[2026, 9, 30], 'Fondo', 'Valuation', '', 40, 2265, '', '', 'DOP'],
    [[2026, 9, 5], 'IBKR', 'Deposit', '', '', '', 1000, '', 'USD'],          // no snapshot: starts from zero
  ]);
  const r = h.plain(h.ctx.computeReturns(ledger, { HAPI: 850, Fondo: 1510, IBKR: 1010 }, { usdRate: 60, today: '2026-10-01' }));
  const by = Object.fromEntries(r.accounts.map(a => [a.account, a]));
  assert.deepEqual([by.HAPI.start, by.HAPI.startValue, by.HAPI.netDeposits], ['2026-09-01', 700, 100]);
  near(by.HAPI.gain, 50);
  near(by.HAPI.periodReturn, 50 / (700 + 100 * 20 / 30));
  assert.equal(by.HAPI.annualized, null, '30 days: not annualized');
  assert.deepEqual([by.Fondo.start, by.Fondo.startValue], ['2026-08-31', 1500]);
  near(by.Fondo.periodReturn, 10 / 1500);
  assert.deepEqual([by.IBKR.start, by.IBKR.startValue, by.IBKR.netDeposits], ['2026-09-05', 0, 1000]);
  near(by.IBKR.gain, 10);
  // all together: Fondo starts first; HAPI and IBKR enter as flows on their start days
  assert.equal(r.total.start, '2026-08-31');
  near(r.total.value, 3370);
  near(r.total.gain, 3370 - 1500 - (700 + 100 + 1000));
  near(r.total.startValue, 1500 + 700 + 0, 1e-9);                 // every account's start value (v1.1.34)
  near(r.total.netDeposits, 100 + 1000, 1e-9);                    // real deposits only — not the start values
  near(r.total.value - r.total.startValue - r.total.netDeposits, r.total.gain, 1e-9);
});

test('returns: annualized after 180 days', () => {
  const h = load();
  const ledger = L(h, [[[2025, 9, 1], 'HAPI', 'Snapshot', 'CASH', '', '', 1000, '', 'USD']]);
  const r = h.plain(h.ctx.computeReturns(ledger, { HAPI: 1100 }, { usdRate: 60, today: '2026-09-01' }));
  near(r.accounts[0].annualized, 0.1, 1e-4);
  near(r.total.periodReturn, 0.1);
});

test('values for the history: GOOGLEFINANCE results when they are numbers, quantity × price otherwise', () => {
  const h = load();
  const holdings = { positions: [{ account: 'HAPI', ticker: 'GOOGL', qty: 2, lastPrice: 300 }, { account: 'HAPI', ticker: 'ETHUSD', qty: 0.5, lastPrice: 2600 },
    { account: 'HAPI', ticker: 'NVDA', qty: 1, lastPrice: 180 }],
    cash: [{ account: 'HAPI', amount: 7.5 }], valuations: [{ account: 'Fondo', value: 90000, currency: 'DOP', units: 40 }] };
  const rows = h.plain(h.ctx.holdingsValueRows(holdings, [640, '=D11*G11', '#N/A'], { fetched: { ETHUSD: 2800 }, usdRate: 60 }));
  assert.deepEqual(rows.map(r => [r.ticker, r.price, r.value]),
    [['GOOGL', 320, 640], ['ETHUSD', 2800, 1400], ['NVDA', 180, 180], ['CASH', '', 7.5], ['', '', 1500]]);
});

function workbook() {
  const mock = makeServices();
  const h = load({ services: mock.services });
  return { h, mock };
}
const hist = (h, rows) => [['Date', 'Account', 'Ticker', 'Quantity', 'Price (US$)', 'Value (US$)']]
  .concat(rows.map(r => { const x = r.slice(); x[0] = h.date(...x[0]); return x; }));

test('history: one set of rows per day — a second run the same day replaces it; TOTAL row', () => {
  const { h, mock } = workbook();
  const v = [{ account: 'HAPI', ticker: 'GOOGL', qty: 2, price: 300, value: 600 }, { account: 'HAPI', ticker: 'CASH', qty: '', price: '', value: 10 }];
  h.ctx.recordPortfolioHistory(v, '2026-09-25');
  h.ctx.recordPortfolioHistory(v, '2026-09-26');
  h.ctx.recordPortfolioHistory([Object.assign({}, v[0], { price: 310, value: 620 }), v[1]], '2026-09-26');   // same day again
  const rows = mock.ss.getSheetByName('Portfolio History')._rows(6);
  assert.equal(rows.length, 6, '3 rows per day, 2 days');
  const totals = rows.filter(r => r[1] === 'TOTAL').map(r => [h.ctx.normalizeDateForCompare(r[0]), r[5]]);
  assert.deepEqual(h.plain(totals), [['2026-09-26', 630], ['2026-09-25', 610]], 'newest first, latest run of the day kept');
});

test('daily brief: change since the previous day without deposits, biggest movers', () => {
  const h = load();
  const history = hist(h, [
    [[2026, 9, 24], 'HAPI', 'GOOGL', 2, 300, 600], [[2026, 9, 24], 'HAPI', 'NVDA', 1, 200, 200], [[2026, 9, 24], 'TOTAL', '', '', '', 800],
    [[2026, 9, 25], 'HAPI', 'GOOGL', 2, 306, 612], [[2026, 9, 25], 'HAPI', 'NVDA', 1, 190, 190], [[2026, 9, 25], 'HAPI', 'CASH', '', '', 100],
    [[2026, 9, 25], 'TOTAL', '', '', '', 902]]);
  const ledger = L(h, [[[2026, 9, 25], 'HAPI', 'Deposit', '', '', '', 6000, '', 'DOP']]);
  const b = h.plain(h.ctx.investmentsDailyBrief(history, ledger, { usdRate: 60 }));
  assert.deepEqual([b.day, b.prevDay, b.total], ['2026-09-25', '2026-09-24', 902]);
  near(b.deposits, 100);
  near(b.change, 2, 1e-9);                     // 902 − 800 − 100 deposited
  assert.deepEqual(b.movers.map(m => m.ticker), ['NVDA', 'GOOGL'], 'largest move first');
  near(b.movers[0].change, -0.05);
  assert.equal(h.plain(h.ctx.investmentsDailyBrief(hist(h, [[[2026, 9, 25], 'TOTAL', '', '', '', 900]]), ledger, {})).change, null, 'first day');
});

test('monthly brief: from the last value of the previous month; deposits, dividends, fees, allocation', () => {
  const h = load();
  const history = hist(h, [
    [[2026, 8, 31], 'HAPI', 'GOOGL', 2, 300, 600], [[2026, 8, 31], 'TOTAL', '', '', '', 600],
    [[2026, 9, 30], 'HAPI', 'GOOGL', 3, 310, 930], [[2026, 9, 30], 'Fondo', '', 40, '', 1500], [[2026, 9, 30], 'TOTAL', '', '', '', 2430]]);
  const ledger = L(h, [[[2026, 9, 3], 'HAPI', 'Deposit', '', '', '', 300, '', 'USD'], [[2026, 9, 12], 'HAPI', 'Dividend', 'GOOGL', '', '', 1.5, '', 'USD'],
    [[2026, 9, 10], 'HAPI', 'Buy', 'GOOGL', 1, 300, 300, 0.15, 'USD'], [[2026, 9, 20], 'Fondo', 'Deposit', '', '', '', 90000, '', 'DOP']]);
  const b = h.plain(h.ctx.investmentsMonthlyBrief(history, ledger, { usdRate: 60, month: h.date(2026, 9, 1) }));
  // v1.1.37: "Fondo" first appears on Sep 30 — added during the month, so its start is that first value and its
  // earlier deposit isn't this month's contribution; the gain is the same (+30 on HAPI), the row reads right
  assert.deepEqual([b.startDay, b.endDay, b.startValue, b.endValue], ['2026-08-31', '2026-09-30', 2100, 2430]);
  assert.deepEqual(b.added, ['Fondo']);
  near(b.deposits, 300);
  near(b.gain, 30);
  near(b.dividends, 1.5); near(b.fees, 0.15);
  assert.deepEqual(b.allocation.map(a => a.account), ['Fondo', 'HAPI']);
  assert.equal(h.plain(h.ctx.investmentsMonthlyBrief(history, ledger, { month: h.date(2026, 7, 1) })), null, 'no data that month');
});

test('emails: an Investments section in the daily and monthly summaries (and it can be turned off)', () => {
  const h = load();
  const brief = { day: '2026-09-25', total: 902, totalDop: 54120, prevDay: '2026-09-24', change: 2, changePct: 0.0025, deposits: 100,
    movers: [{ ticker: 'NVDA', account: 'HAPI', change: -0.05 }], returns: { periodReturn: 0.071, gain: 50, start: '2026-09-01' } };
  const s = h.ctx.computeDailySummary([new Array(14).fill('h')], { today: h.date(2026, 9, 26), rates: { USD: 60, EUR: 64, COP: 0.015 },
    netIncomeDop: 80000, cards: [], investments: brief });
  const decode = x => x.replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(Number(n)));
  const on = decode(h.ctx.buildDailySummaryEmail(s, { sheetName: 'T' }).html);
  for (const piece of ['Investments', 'US$902.00', '+US$2.00', 'deposits left out', 'NVDA', '−5.00%', '+7.10%', '+0.25%']) assert.ok(on.includes(piece), piece);
  const off = decode(h.ctx.buildDailySummaryEmail(s, { sheetName: 'T', sections: { investments: false } }).html);
  assert.ok(!off.includes('US$902.00'));
  const m = h.ctx.computeMonthlySummary([new Array(14).fill('h')], { today: h.date(2026, 10, 1), rates: { USD: 60, EUR: 64, COP: 0.015 },
    netIncomeDop: 80000, cards: [], investments: { startDay: '2026-08-31', endDay: '2026-09-30', startValue: 600, endValue: 2430, deposits: 1800,
      gain: 30, gainPct: 0.05, dividends: 1.5, fees: 0.15, partial: false, allocation: [{ account: 'Fondo', value: 1500, share: 0.617 }],
      returns: { periodReturn: 0.02, gain: 45, start: '2026-08-31', annualized: null } } });
  const mm = decode(h.ctx.buildMonthlySummaryEmail(m, { sheetName: 'T' }).html);
  for (const piece of ['Investments', 'US$2,430.00', 'Gain in Sep', '+US$30.00', 'dividends US$1.50', 'Fondo', 'Since tracking began (2026-08-31)']) {
    assert.ok(mm.includes(piece), piece);
  }
});

test('Holdings: a Performance block, and the chart is replaced — not stacked — on every rebuild', () => {
  const { h, mock } = workbook();
  h.ctx.getOrCreateLedgerSheet().getRange(2, 1, 2, 12).setValues([
    [h.date(2026, 9, 1), 'HAPI', 'Snapshot', 'GOOGL', 2, 300, 500, '', 'USD', 'manual', '', ''],
    [h.date(2026, 9, 1), 'HAPI', 'Snapshot', 'CASH', '', '', 100, '', 'USD', 'manual', '', '']]);
  const history = mock.ss.insertSheet('Portfolio History');   // a previous day already recorded
  history.getRange(1, 1, 3, 6).setValues([['Date', 'Account', 'Ticker', 'Quantity', 'Price (US$)', 'Value (US$)'],
    [h.date(2026, 9, 1), 'HAPI', 'GOOGL', 2, 300, 600], [h.date(2026, 9, 1), 'TOTAL', '', '', '', 700]]);
  h.ctx.refreshHoldings();
  h.ctx.refreshHoldings();
  const sheet = mock.ss.getSheetByName('Holdings');
  assert.deepEqual(sheet.charts.map(c => c.cfg.type).sort(), ['LINE', 'PIE'], 'one of each after two rebuilds — not stacked (v1.1.34: + allocation pie)');
  const values = [...sheet.cells.values()].map(String);
  assert.ok(values.some(v => v.endsWith('Performance')) && values.includes('All accounts'), 'Performance section and its total row');
  assert.equal(mock.ss.getSheetByName('Portfolio History')._rows(6).filter(r => r[1] === 'TOTAL').length, 2, 'today recorded once');
});
