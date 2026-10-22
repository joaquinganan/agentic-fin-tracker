'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');

// v1.1.74: the daily summary shows only the change of the day before. Asked: "Since Wed 21 Oct" — the change from the
// 20th to the 21st, or from the 21st to today? It was the 21st's (the snapshot of a day holds the end of the day
// before), now labelled with both days: "20 → 21 Oct". After a day without an update the change spanned every day since: now it is "—".

const L = [['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id']];
const hist = (h, rows) => [['Date', 'Account', 'Ticker', 'Quantity', 'Price (US$)', 'Value (US$)', 'Day change']]
  .concat(rows.map(r => { const x = r.slice(); x[0] = h.date(...x[0]); while (x.length < 7) x.push(''); return x; }));
const day = (d, googl, btc) => [[[2026, 10, d], 'HAPI', 'GOOGL', 2, googl / 2, googl], [[2026, 10, d], 'Wallet', 'BTCUSD', 0.01, btc * 100, btc],
  [[2026, 10, d], 'TOTAL', '', '', '', googl + btc]];
const emailOf = (h, iv) => {
  const s = h.ctx.computeDailySummary([new Array(15).fill('h')], { today: h.date(2026, 10, 22, 8), rates: { USD: 60, EUR: 64, COP: 0.015 },
    netIncomeDop: 80000, cards: [], investments: iv });
  const mail = h.ctx.buildDailySummaryEmail(s, { sheetName: 'T' });
  return { html: mail.html.replace(/&#(\d+);/g, (a, n) => String.fromCodePoint(Number(n))), text: mail.text };
};

test('the change is the previous day\'s, labelled with both days, "20 → 21 Oct" (not "Since …")', () => {
  const h = load();
  const iv = h.ctx.investmentsDailyBrief(hist(h, day(20, 300, 800).concat(day(21, 310, 810), day(22, 320, 840))), L, { usdRate: 60 });
  assert.equal(h.plain(iv).change, 40, 'end of the 21st (snapshot of the 22nd) − end of the 20th (snapshot of the 21st)');
  const { html, text } = emailOf(h, iv);
  assert.ok(html.includes('>20 → 21 Oct<'), 'the label names both ends: the 20th against the 21st');
  assert.ok(!html.includes('Since Wed 21 Oct'));
  assert.ok(html.includes('Wallet · 20 → 21 Oct'), 'a crypto move says its days too');
  assert.ok(text.includes('(+US$40.00, 20 → 21 Oct)'));
});

test('without the snapshot of the day before, no change of several days: "—" and which day had no update', () => {
  const h = load();
  const b = h.plain(h.ctx.investmentsDailyBrief(hist(h, day(19, 300, 800).concat(day(22, 320, 840))), L, { usdRate: 60 }));
  assert.equal(b.change, null);
  assert.equal(b.prevDay, null);
  assert.equal(b.missingDay, '2026-10-21');
  assert.equal(b.movers.filter(m => !m.session).length, 0, 'no price moves over three days');
  const { html } = emailOf(h, h.ctx.investmentsDailyBrief(hist(h, day(19, 300, 800).concat(day(22, 320, 840))), L, { usdRate: 60 }));
  assert.ok(html.includes('no update on Wed 21 Oct'));
  const first = h.plain(h.ctx.investmentsDailyBrief(hist(h, day(22, 320, 840)), L, { usdRate: 60 }));
  assert.equal(first.missingDay, null, 'the first day recorded is not a missing update');
  assert.ok(emailOf(h, h.ctx.investmentsDailyBrief(hist(h, day(22, 320, 840)), L, { usdRate: 60 })).html.includes('first day recorded'));
});

test('the day before, and the label, across a month end', () => {
  const h = load();
  const rows = [[[2026, 10, 31], 'TOTAL', '', '', '', 100], [[2026, 11, 1], 'TOTAL', '', '', '', 110]];
  assert.equal(h.plain(h.ctx.investmentsDailyBrief(hist(h, rows), L, {})).prevDay, '2026-10-31');
  assert.equal(h.ctx.summaryDayChangeLabel('2026-11-01'), '31 Oct → 1 Nov');
});

// Reported: a run during market hours on the 21st replaced that day's snapshot (the 20th's close) with midday prices;
// the email on the 22nd showed a loss for a day the stocks went up. Numbers invented, same shape.
const stock = (d, price, dc, qty) => [[2026, 10, d], 'HAPI', 'TSLA', qty || 2, price, +((qty || 2) * price).toFixed(2), dc];
const totalRow = (d, v) => [[2026, 10, d], 'TOTAL', '', '', '', v];

test('a stock the day before is at its close before the last session, not the price of a midday snapshot (reported)', () => {
  const h = load();
  // close of the 20th: 100. Midday on the 21st: 104 (+4%). Close of the 21st: 103 (+3%).
  const b = h.plain(h.ctx.investmentsDailyBrief(hist(h, [stock(21, 104, 0.04), totalRow(21, 208), stock(22, 103, 0.03), totalRow(22, 206)]), L, {}));
  assert.equal(+b.change.toFixed(2), +(206 - 2 * 103 / 1.03).toFixed(2), 'the 20th\'s close (100) to the 21st\'s (103): +6, not −2');
  assert.ok(Math.abs(b.changePct - 0.03) < 1e-9);
});

test('with no session in between (a weekend, a holiday) the stocks do not move, whatever their last day change', () => {
  const h = load();
  // Saturday's and Sunday's 6 AM snapshots both have Friday's close and Friday's day change
  const b = h.plain(h.ctx.investmentsDailyBrief(hist(h, [stock(24, 103, 0.03), totalRow(24, 206), stock(25, 103, 0.03), totalRow(25, 206)]), L, {}));
  assert.equal(b.change, 0);
});

test('the quantity the day before stays its own (shares bought that day are not counted as gain)', () => {
  const h = load();
  // 1 share at the close of the 20th (100); 1 more bought on the 21st with a deposit of 103
  const ledger = L.concat([[h.date(2026, 10, 21), 'HAPI', 'Deposit', 'CASH', '', '', 103, '', 'USD', 'email', '', '']]);
  const b = h.plain(h.ctx.investmentsDailyBrief(hist(h, [stock(21, 104, 0.04, 1), totalRow(21, 104), stock(22, 103, 0.03, 2), totalRow(22, 206)]), ledger, {}));
  assert.equal(+b.change.toFixed(2), +(206 - 103 - 103 / 1.03).toFixed(2), 'the share held gained 3; the new one is the deposit');
});

test('a deposit and a buy taken in by a snapshot of the day (a mid-morning run, seen in a real sheet) are not gain', () => {
  const h = load();
  // the 16th's snapshot, taken mid-morning, already has the deposit of 1,000 and the shares it bought
  const ledger = L.concat([[h.date(2026, 10, 16), 'HAPI', 'Deposit', 'CASH', '', '', 1000, '', 'USD', 'email', '', '']]);
  const rows = [[[2026, 10, 15], 'HAPI', 'CASH', '', '', 2], totalRow(15, 2),
    stock(16, 330, 0.01, 3), [[2026, 10, 16], 'HAPI', 'CASH', '', '', 12], totalRow(16, 1002),
    stock(17, 333.3, 0.01, 3), [[2026, 10, 17], 'HAPI', 'CASH', '', '', 12], totalRow(17, 1011.9)];
  const first = h.plain(h.ctx.investmentsDailyBrief(hist(h, rows.slice(0, 5)), ledger, {}));
  assert.equal(first.change, 0, 'the 15th held only cash: depositing and buying is not gain');
  const next = h.plain(h.ctx.investmentsDailyBrief(hist(h, rows.slice(2)), ledger, {}));
  assert.equal(+next.change.toFixed(2), +(3 * (333.3 - 333.3 / 1.01)).toFixed(2), 'the next day: only the market move');
});

test('a fund or pension (a balance, no ticker) changes by its balance, less its deposits', () => {
  const h = load();
  const ledger = L.concat([[h.date(2026, 10, 22), 'Fund', 'Deposit', '', '', '', 100, '', 'USD', 'manual', '', '']]);
  const rows = [[[2026, 10, 21], 'Fund', '', 54.6, '', 2000], totalRow(21, 2000), [[2026, 10, 22], 'Fund', '', 56, '', 2110], totalRow(22, 2110)];
  const b = h.plain(h.ctx.investmentsDailyBrief(hist(h, rows), ledger, {}));
  assert.equal(+b.change.toFixed(2), 10);
  assert.equal(b.deposits, 100);
});
