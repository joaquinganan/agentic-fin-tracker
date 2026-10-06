'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');

// v1.1.65: the month-end pace counts bills and fixed costs once (plus the usual ones not paid yet) and extrapolates
// only day-to-day spending. Before, everything spent so far was multiplied by days-in-month / day-of-month.
const D = 'Dining/Delivery + Entertainment + Other';
function rowsFor(list) {
  return [new Array(14).fill('h')].concat(list.map(o => {
    const r = new Array(14).fill('');
    r[0] = o.date; r[1] = o.bank || 'LAFISE'; r[2] = o.merchant || 'X'; r[3] = o.amount; r[4] = o.currency || 'DOP';
    r[5] = o.category === undefined ? D : o.category; r[11] = o.type || 'Transaction';
    return r;
  }));
}
const RATES = { USD: 60, EUR: 64, COP: 0.015 };
const NET = 200000;
// the email is sent on Tue 6 Oct: yesterday is day 5 of a 31-day month
const opts = h => ({ today: h.date(2026, 10, 6, 8), rates: RATES, netIncomeDop: NET, cards: [] });
const summarize = (h, list) => h.plain(h.ctx.computeDailySummary(rowsFor(list), opts(h)));
const email = (h, list) => {
  const s = h.ctx.computeDailySummary(rowsFor(list), opts(h));
  const mail = h.plain(h.ctx.buildDailySummaryEmail(s, { sheetUrl: 'https://x', sheetName: 'Tracker' }));
  mail.decoded = mail.html.replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(Number(n)));
  return mail;
};

// Reported: rent paid on day 1, then a normal week; on day 5 the pace said 166% of net income (amounts invented).
function reportedCase(h) {
  const d = n => h.date(2026, 10, n);
  return [
    { date: d(1), merchant: 'LANDLORD EJEMPLO', amount: 45000, category: 'Rent', type: 'Transfer', bank: 'BANESCO' },
    { date: d(1), merchant: 'OWN ACCOUNT', amount: 60000, category: 'Exclude', type: 'Transfer' },
    { date: d(1), merchant: 'OWN ACCOUNT', amount: -60000, category: 'Exclude', type: 'Incoming', bank: 'BANESCO' },
    { date: d(1), merchant: 'UBER*EATS', amount: 400 }, { date: d(1), merchant: 'UBER*EATS', amount: 500 },
    { date: d(1), merchant: 'UBER*RIDES', amount: 120, category: 'Transportation' },
    { date: d(1), merchant: 'UBER*RIDES', amount: 160, category: 'Transportation' },
    { date: d(2), merchant: 'UBER*EATS', amount: 360 },
    { date: d(3), merchant: 'SHELL EJEMPLO', amount: 2800, category: 'Vehicle Gas' },
    { date: d(3), merchant: 'UBER*EATS', amount: 400 }, { date: d(3), merchant: 'UBER*EATS', amount: 600 },
    { date: d(3), merchant: 'UBER*EATS', amount: 460 },
    { date: d(4), merchant: 'SPOTIFY', amount: 3.5, currency: 'USD', category: 'Streaming & Subscriptions' },
    { date: d(4), merchant: 'UBER*EATS', amount: 400 }, { date: d(4), merchant: 'UBER*EATS', amount: 500 },
    { date: d(5), merchant: 'TRAINER EJEMPLO', amount: 2500, category: 'Gym + Calisthenics', type: 'Transfer', bank: 'BANESCO' },
    { date: d(5), merchant: 'UBER*EATS', amount: 400 }, { date: d(5), merchant: 'UBER*EATS', amount: 380 }
  ];
}

test('reported case: the old pace multiplied the rent by 31/5 and said over 100% of net income', () => {
  const h = load();
  const s = summarize(h, reportedCase(h));
  assert.equal(s.dayOfMonth, 5);
  assert.equal(s.mtd, 55190, 'own-account transfers stay out; the subscription is converted');
  const old = s.mtd / 5 * 31;
  assert.ok(old / NET > 1.66, 'the fixture reproduces the reported alarm (old formula: ' + Math.round(old) + ')');

  assert.equal(s.mtdFixed, 47710, 'rent + gym + subscription');
  assert.equal(s.mtdDayToDay, 7480);
  assert.equal(s.pendingFixed, 0, 'no previous months: nothing pending');
  assert.equal(Math.round(s.projected), 47710 + 7480 / 5 * 31, 'bills once, day-to-day at its pace');
  assert.equal(Math.round(s.projectedPct * 100), 47);
});

test('reported case in the email: on track, no red warning, the tile and the text line agree', () => {
  const h = load();
  const mail = email(h, reportedCase(h));
  assert.ok(!mail.decoded.includes('🔴'), 'no red warning');
  assert.ok(!/17[01]%|166%/.test(mail.decoded + mail.text));
  assert.ok(mail.decoded.includes('On track: at this pace the month ends at about 47% of your net income.'));
  assert.ok(mail.text.includes('At this pace: RD$94,086 (47% of net income) by month-end'));
});

test('bills of a usual month not paid yet are added once; bills already paid are not added again', () => {
  const h = load();
  const list = [];
  // July, August, September: rent RD$45,000, electricity RD$3,000, phone RD$1,500, and day-to-day spending
  [7, 8, 9].forEach(m => {
    list.push({ date: h.date(2026, m, 1), amount: 45000, category: 'Rent', type: 'Transfer' });
    list.push({ date: h.date(2026, m, 12), amount: 3000, category: 'Electricity' });
    list.push({ date: h.date(2026, m, 15), amount: 1500, category: 'Telecommunications' });
    list.push({ date: h.date(2026, m, 20), amount: 9000 });
  });
  // October so far: rent paid, phone paid above its usual, electricity not yet; RD$1,000 a day of day-to-day
  list.push({ date: h.date(2026, 10, 1), amount: 45000, category: 'Rent', type: 'Transfer' });
  list.push({ date: h.date(2026, 10, 2), amount: 2000, category: 'Telecommunications' });
  for (let n = 1; n <= 5; n++) list.push({ date: h.date(2026, 10, n), amount: 1000 });
  const s = summarize(h, list);
  assert.equal(s.mtdFixed, 47000);
  assert.equal(s.pendingFixed, 3000, 'electricity pending; rent paid; phone paid over its usual counts nothing negative');
  assert.equal(s.projected, 47000 + 3000 + 31000);
});

test('without bills, the pace is the same extrapolation as before', () => {
  const h = load();
  const list = [1, 2, 3, 4, 5].map(n => ({ date: h.date(2026, 10, n), amount: 1200 }));
  const s = summarize(h, list);
  assert.equal(s.mtdFixed, 0);
  assert.equal(s.projected, 6000 / 5 * 31);
});
