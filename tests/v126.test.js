'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

const D = 'Dining/Delivery + Entertainment + Other';
function rowsFor(h, list) {
  return [new Array(14).fill('h')].concat(list.map(o => {
    const r = new Array(14).fill('');
    r[0] = o.date; r[1] = o.bank || 'LAFISE'; r[2] = o.merchant || 'X'; r[3] = o.amount; r[4] = o.currency || 'DOP';
    r[5] = o.category === undefined ? D : o.category; r[11] = o.type || 'Transaction';
    return r;
  }));
}
const RATES = { USD: 60, EUR: 64, COP: 0.015 };

// ---------- daily: bills don't make a day look like a spike
test('daily: rent/electricity are left out of the day-to-day comparison, but still in the total', () => {
  const h = load();
  const list = [];
  for (let d = 1; d <= 23; d++) list.push({ date: h.date(2026, 9, d), amount: 1000 });
  list.push({ date: h.date(2026, 9, 24), merchant: 'EDESUR APP-WB', amount: 3600, category: 'Electricity' });
  list.push({ date: h.date(2026, 9, 24), merchant: 'UBER*EATS', amount: 900 });
  const s = h.plain(h.ctx.computeDailySummary(rowsFor(h, list), { today: h.date(2026, 9, 25), rates: RATES, netIncomeDop: 60000, cards: [] }));
  assert.equal(s.spent, 4500, 'hero total includes the bill');
  assert.equal(s.spentDayToDay, 900);
  assert.equal(Math.round(s.vsTypical * 100), -10, 'compared on day-to-day spending only');   // v1.1.64: typical day, was the average
  assert.equal(s.highDay, false);
  assert.ok(!s.recommendations.some(r => /highest days/.test(r.text)), 'no false alarm from the bill');
  assert.equal(s.last7.length, 7);
  assert.deepEqual(s.last7.map(x => x.amount), [1000, 1000, 1000, 1000, 1000, 1000, 4500]);
  assert.ok(s.mtdTop.every(c => c.cat !== 'Electricity'), 'bars show variable categories only');
  assert.equal(s.mtdBills, 3600);
});

test('daily email: redesigned blocks, preheader, links, data-health line', () => {
  const h = load();
  const list = [{ date: h.date(2026, 9, 24), merchant: 'SM POLA', amount: 1200, category: 'Groceries + Barbershop', bank: 'BANESCO' }];
  const s = h.ctx.computeDailySummary(rowsFor(h, list), { today: h.date(2026, 9, 25, 8), rates: RATES, netIncomeDop: 60000, cards: [],
    lastRun: { at: h.date(2026, 9, 25, 6, 4).toISOString(), saved: 3, unparsed: 1, errors: 0 }, now: h.date(2026, 9, 25, 8, 10) });
  const mail = h.plain(h.ctx.buildDailySummaryEmail(s, { sheetName: 'Tracker',
    links: { dashboard: 'https://d', transactions: 'https://t', rules: 'https://r' } }));
  for (const piece of ['Last 7 days', 'Yesterday&#39;s purchases', 'This month', 'Month elapsed', 'Net income used', 'https://d', 'https://t']) {
    assert.ok(mail.html.includes(piece) || mail.html.includes(piece.replace('&#39;', "'")), piece);
  }
  assert.match(mail.preheader, /^RD\$1,200 yesterday · \d+% of net income used with \d+% of the month gone$/);
  assert.ok(mail.html.includes('display:none'), 'hidden preheader for the inbox preview');
  const decoded = mail.html.replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(Number(n)));   // the HTML is ASCII-only
  assert.ok(decoded.includes('Data updated today at 6:04 AM · 3 new transaction(s) · 1 email(s) couldn'));
  assert.ok(/^[\x00-\x7F]*$/.test(mail.html));
});

test('data health: stale update and unknown status', () => {
  const h = load();
  const stale = h.plain(h.ctx.summaryDataHealth({ at: h.date(2026, 9, 22, 6, 4).toISOString(), saved: 1 }, h.date(2026, 9, 25, 8)));
  assert.equal(stale.tone, 'warn');
  assert.match(stale.text, /No automatic update since Tue 22 Sep at 6:04 AM/);
  assert.equal(h.plain(h.ctx.summaryDataHealth(null, h.date(2026, 9, 25))).tone, 'note');
});

// ---------- monthly
function monthData(h) {
  const list = [];
  for (const [m, last, uber] of [[6, 30, 300], [7, 31, 300], [8, 31, 450]]) {
    for (let d = 1; d <= last; d++) list.push({ date: h.date(2026, m, d), merchant: 'UBER*RIDES', amount: uber, category: 'Transportation' });
    list.push({ date: h.date(2026, m, 2), merchant: 'LANDLORD', amount: 30000, category: 'Rent', type: 'Transfer' });
    list.push({ date: h.date(2026, m, 5), merchant: 'Spotify', amount: 12, currency: 'USD', category: 'Streaming & Subscriptions' });
  }
  list.push({ date: h.date(2026, 8, 6), merchant: 'Netflix', amount: 10, currency: 'USD', category: 'Streaming & Subscriptions' });
  list.push({ date: h.date(2026, 8, 10), merchant: 'FARMA VALUE FD05', amount: 2000, bank: 'POPULAR', category: 'Health + Vet + Pharmacy' });
  list.push({ date: h.date(2026, 8, 12), merchant: 'MARIA PRUEBA', amount: 5000, category: '', type: 'Transfer' });
  list.push({ date: h.date(2026, 8, 15), merchant: 'PAGO TARJETA', amount: 20000, type: 'Card Payment', category: 'Exclude' });
  list.push({ date: h.date(2026, 9, 1), merchant: 'UBER*RIDES', amount: 999, category: 'Transportation' });   // next month: ignored
  return rowsFor(h, list);
}

test('monthly: totals, comparisons, subscriptions, missed cashback, year to date (pure)', () => {
  const h = load();
  const m = h.plain(h.ctx.computeMonthlySummary(monthData(h), { today: h.date(2026, 9, 1, 9), rates: RATES, netIncomeDop: 80000,
    cards: [{ bank: 'BHD', rate: 0.05 }] }));
  assert.equal(m.monthKey, '2026-08');
  const aug = 31 * 450 + 30000 + 720 + 600 + 2000;
  assert.equal(m.spent, aug, 'card payment and open transfer are not spending; September is ignored');
  assert.equal(m.prevTotal, 31 * 300 + 30000 + 720);
  assert.equal(Math.round(m.vsPrev * 1000), Math.round((aug / (31 * 300 + 30720) - 1) * 1000));
  assert.equal(m.fixed, 30000 + 1320, 'rent and subscriptions are fixed categories');
  assert.equal(m.leftOver, 80000 - aug);
  assert.deepEqual(m.subscriptions.map(x => [x.merchant, x.recurring]), [['Spotify', true], ['Netflix', false]]);
  assert.equal(m.missed.total, 100, 'pharmacy on POPULAR, 5% with the BHD card you have');
  assert.equal(m.openCount, 1);
  assert.equal(m.ytdMonths, 3);
  assert.equal(m.byDayToDay[2], 450, 'calendar colours leave the rent out');
  assert.equal(m.byDay[2], 30450);
  const texts = m.recommendations.map(r => r.text).join(' | ');
  assert.match(texts, /Spending was up \d+% vs\. July, mostly Transportation/);
  assert.match(texts, /2 subscription\(s\) cost RD\$1,320 this month — about RD\$15,840 a year/);
  assert.match(texts, /RD\$100 in cashback was left on the table — e\.g\. pharmacies paid with POPULAR/);
  assert.match(texts, /1 transfer\(s\) have no category/);
});

test('monthly email: subject, preheader, calendar, sections, ASCII-safe', () => {
  const h = load();
  const m = h.ctx.computeMonthlySummary(monthData(h), { today: h.date(2026, 9, 1, 9), rates: RATES, netIncomeDop: 80000, cards: [{ bank: 'BHD', rate: 0.05 }] });
  const mail = h.plain(h.ctx.buildMonthlySummaryEmail(m, { sheetName: 'Tracker', links: { dashboard: 'https://d', rules: 'https://r' } }));
  assert.match(mail.subject, /^August 2026: RD\$[\d,]+ spent · \d+% left over$/);
  assert.match(mail.preheader, /left over · \+\d+% vs\. July$/);
  for (const piece of ['Day by day', 'Categories', 'Top merchants', 'Subscriptions', 'By bank', 'Transfers', '2026 so far', 'Recommendations', 'https://r']) {
    assert.ok(mail.html.includes(piece), piece);
  }
  assert.equal((mail.html.match(/height:34px/g) || []).length, 31, 'one calendar cell per day of August');
  assert.ok(/^[\x00-\x7F]*$/.test(mail.html) && !/[\u{10000}-\u{10FFFF}]/u.test(mail.subject + mail.text));
  assert.equal(h.ctx.summarySignedPct(-0.001), '±0%');
});

test('monthly delivery: 1st-of-month trigger, on/off switch, menu always sends; run status recorded', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  mock.ss.insertSheet('Transactions'); mock.ss.insertSheet('Custom Rules');
  const input = monthly => ({ email: 'owner@example.com', incomeCurrency: 'USD', monthlyIncome: 3000, deductionMode: 'manual',
    deductions: { ARS: 0, AFP: 0, ISR: 0, taxRate: 10 }, banksToTrack: { LAFISE: true },
    notify: { enabled: false, monthly: monthly, email: '', hour: 9, sections: { totals: true } }, timestamp: 'x' });
  h.ctx.saveSetupConfig(input(true));
  assert.deepEqual(mock.triggers.map(t => [t.handler, t.hour, t.monthDay]).sort(),
    [['runGmailMonitor', 6, undefined], ['sendMonthlySummary', 9, 1]]);
  h.ctx.sendMonthlySummary();
  assert.equal(mock.gmail.sent.length, 1);
  assert.match(mock.gmail.sent[0].subject, /^[A-Z][a-z]+ \d{4}: RD\$/);

  h.ctx.saveSetupConfig(input(false));
  assert.deepEqual(mock.triggers.map(t => t.handler), ['runGmailMonitor']);
  h.ctx.sendMonthlySummary();
  assert.equal(mock.gmail.sent.length, 1, 'turned off → the trigger sends nothing');
  h.ctx.sendMonthlySummaryNow();
  assert.equal(mock.gmail.sent.length, 2, 'menu item always sends');

  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-01');
  const run = JSON.parse(mock.props.FT_LAST_RUN);
  assert.ok(run.at && run.saved === 0 && run.errors === 0, 'the monitor records its outcome for the emails');
});

test('setup validation: monthly alone needs no daily sections', () => {
  const h = load();
  const cfg = { email: 'a@b.co', incomeCurrency: 'USD', monthlyIncome: 3000, deductionMode: 'manual',
    deductions: { ARS: 0, AFP: 0, ISR: 0, taxRate: 0 }, banksToTrack: { BHD: true },
    notify: { enabled: false, monthly: true, email: '', hour: 9, sections: { totals: false } } };
  assert.equal(h.ctx.validateSetupInput(cfg), null);
  cfg.notify.email = 'nope';
  assert.match(h.ctx.validateSetupInput(cfg), /valid email for the summaries/);
});
