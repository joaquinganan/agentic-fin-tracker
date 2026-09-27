'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

const hist = (h, rows) => [['Date', 'Account', 'Ticker', 'Quantity', 'Price (US$)', 'Value (US$)']]
  .concat(rows.map(r => { const x = r.slice(); x[0] = h.date(...x[0]); return x; }));
const L = (h, rows) => [['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id']]
  .concat(rows.map(r => { const x = r.slice(); x[0] = h.date(...x[0]); while (x.length < 12) x.push(''); return x; }));

test('daily change: accounts added since the previous day are not gain (reported: +US$17,104 the day two were added)', () => {
  const h = load();
  const history = hist(h, [
    [[2026, 9, 26], 'HAPI', 'GOOGL', 2, 300, 8000], [[2026, 9, 26], 'HAPI', 'CASH', '', '', 411], [[2026, 9, 26], 'TOTAL', '', '', '', 8411],
    [[2026, 9, 28], 'HAPI', 'GOOGL', 2, 302.5, 8100], [[2026, 9, 28], 'HAPI', 'CASH', '', '', 411],
    [[2026, 9, 28], 'Pension fund', '', '', '', 15008], [[2026, 9, 28], 'Liquidity fund', '', 54.6, '', 2091], [[2026, 9, 28], 'TOTAL', '', '', '', 25610]]);
  const b = h.plain(h.ctx.investmentsDailyBrief(history, L(h, []), { usdRate: 60 }));
  assert.equal(b.total, 25610, 'the portfolio is still everything');
  assert.equal(b.change, 100, 'only HAPI moved: 8,511 − 8,411');
  assert.ok(Math.abs(b.changePct - 100 / 8411) < 1e-12);
  assert.deepEqual(b.added.sort(), ['Liquidity fund', 'Pension fund']);
  const email = h.ctx.buildDailySummaryEmail(h.ctx.computeDailySummary([new Array(14).fill('h')], { today: h.date(2026, 9, 29),
    rates: { USD: 60, EUR: 64, COP: 0.015 }, netIncomeDop: 80000, cards: [], investments: b }), { sheetName: 'T' }).html;
  assert.ok(email.includes('+US$100.00') && email.includes('new account(s) not counted'));
});

test('status colours are far from every category chip (ΔE ≥ 25) — reported: "uncategorized" looked like Electricity', () => {
  const h = load();
  const lab = hex => {
    const [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255).map(c => c > 0.04045 ? Math.pow((c + 0.055) / 1.055, 2.4) : c / 12.92);
    const f = v => v > 0.008856 ? Math.cbrt(v) : 7.787 * v + 16 / 116;
    const X = (r * 0.4124 + g * 0.3576 + b * 0.1805) / 0.95047, Y = r * 0.2126 + g * 0.7152 + b * 0.0722, Z = (r * 0.0193 + g * 0.1192 + b * 0.9505) / 1.08883;
    return [116 * f(Y) - 16, 500 * (f(X) - f(Y)), 200 * (f(Y) - f(Z))];
  };
  const dE = (a, b) => Math.hypot(...lab(a).map((v, i) => v - lab(b)[i]));
  const chips = h.plain(h.ctx.categoryPalette(['A', 'B', 'C', 'D', 'E'])).map(p => [p.name, p.bg])
    .concat(Object.entries(h.plain(h.get('TYPE_COLORS'))).map(([t, c]) => [t, c[0]]));
  const theme = h.plain(h.get('SHEET_THEME'));
  for (const [state, colour] of [['attention', theme.attention], ['problem', theme.problem]]) {
    for (const [name, bg] of chips) assert.ok(dE(colour, bg) >= 25,
      `${state} ${colour} vs ${name} ${bg}: ΔE ${dE(colour, bg).toFixed(1)}`);
  }
});

test('tabs: spending, then investments together, then settings; settings tabs can be hidden and shown again', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  for (const n of ['Categories', 'Portfolio History', 'Custom Rules', 'Transactions', 'Configuration', 'Investment Accounts', 'Holdings',
                   'Dashboard', 'Investment Ledger', 'Unrecognized']) mock.ss.insertSheet(n);
  h.ctx.ensureSheetOrder();
  assert.deepEqual(mock.ss.getSheets().map(s => s.getName()), ['Dashboard', 'Transactions', 'Unrecognized', 'Holdings', 'Investment Ledger',
    'Portfolio History', 'Custom Rules', 'Investment Accounts', 'Configuration', 'Categories']);
  h.ctx.toggleSettingsTabs();
  const hidden = () => mock.ss.getSheets().filter(s => s.isSheetHidden()).map(s => s.getName()).sort();
  assert.deepEqual(hidden(), ['Categories', 'Configuration', 'Investment Accounts'], 'Custom Rules stays visible — you edit it, and emails link to it');
  assert.match(mock.ui.alerts.pop(), /hidden/i);
  h.ctx.ensureSheetOrder();
  assert.deepEqual(hidden(), ['Categories', 'Configuration', 'Investment Accounts'], 'reordering never unhides them');
  h.ctx.toggleSettingsTabs();
  assert.deepEqual(hidden(), []);
});

test('Unrecognized: link noise (tracking URLs) is removed from what the email says', () => {
  const mock = makeServices();
  const h = load({ services: mock.services });
  h.ctx.recordUnrecognized([{ id: 'x', date: h.date(2026, 7, 6, 14, 55), bank: 'HAPI', subject: 'Order Executed', reason: 'r',
    snippet: '<https://cwk5.r.us-east-1.awstrack.me/L0/https:%2F%2Fapp.hapi.trade/1/0100abc> *Hello Alex, * Your order has been executed. See https://app.hapi.trade/x?y=1 now' }], [], h.date(2026, 9, 27));
  assert.equal(mock.ss.getSheetByName('Unrecognized')._rows(10)[0][4], '*Hello Alex, * Your order has been executed. See now');
});
