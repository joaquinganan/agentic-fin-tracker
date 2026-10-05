'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');

// v1.1.64: the daily summary compares yesterday with a "typical day" (the median of the last 60 days of day-to-day
// spending), warns only when the day is among the highest of those 60 days, and names big one-off purchases apart.
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
const YESTERDAY = [2026, 10, 4];   // a Sunday; the email is sent on Mon 5 Oct
// `n` days before yesterday (1 = the day before yesterday)
const before = (h, n) => h.date(YESTERDAY[0], YESTERDAY[1], YESTERDAY[2] - n);
const yesterday = h => h.date(YESTERDAY[0], YESTERDAY[1], YESTERDAY[2]);
// one entry per day: amounts[0] is the day before yesterday, amounts[1] the day before that, and so on
const history = (h, amounts, extra) => amounts.map((a, i) => Object.assign({ date: before(h, i + 1), amount: a }, extra || {}))
  .filter(o => o.amount !== 0);
const summarize = (h, list) => h.plain(h.ctx.computeDailySummary(rowsFor(list),
  { today: h.date(2026, 10, 5, 8), rates: RATES, netIncomeDop: 60000, cards: [] }));
const email = (h, list) => {
  const s = h.ctx.computeDailySummary(rowsFor(list), { today: h.date(2026, 10, 5, 8), rates: RATES, netIncomeDop: 60000, cards: [] });
  const mail = h.plain(h.ctx.buildDailySummaryEmail(s, { sheetUrl: 'https://x', sheetName: 'Tracker' }));
  mail.decoded = mail.html.replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(Number(n)));
  return mail;
};
const fill = (n, v) => new Array(n).fill(v);

// ---------- the reported case
// Reported: a day with a tank of fuel and three Uber Eats came out "+730%" over the daily average (amounts invented).
function reportedCase(h) {
  // 60 days of history: mostly RD$425, two supermarket days of RD$845 in the last 30
  const amounts = [845].concat(fill(13, 425), [845], fill(15, 425), fill(30, 425));
  const list = history(h, amounts);
  list.push({ date: yesterday(h), merchant: 'SHELL EJEMPLO', amount: 2500 });
  for (let i = 0; i < 3; i++) list.push({ date: yesterday(h), merchant: 'UBER*EATS', amount: 420 });
  return { list, amounts };
}

test('reported case: the old 30-day average gave +730% for a day with a tank of fuel and three Uber Eats', () => {
  const h = load();
  const { list, amounts } = reportedCase(h);
  const avg30 = amounts.slice(0, 30).reduce((a, b) => a + b, 0) / 30;
  assert.equal(Math.round((3760 / avg30 - 1) * 100), 730, 'the fixture reproduces the reported number');

  const s = summarize(h, list);
  assert.equal(s.typicalDay, 425, 'typical day = median of 60 days, not pulled up by the supermarket days');
  assert.equal(s.historyDays, 60);
  assert.deepEqual(s.sporadic.map(x => [x.merchant, x.amount]), [['SHELL EJEMPLO', 2500]], 'the fuel is named apart');
  assert.equal(s.restDayToDay, 1260, 'the three Uber Eats are the rest of the day');
  assert.equal(Math.round(s.vsTypicalRest * 100), 196);

  const mail = email(h, list);
  assert.equal(mail.subject, 'Sun 4 Oct 2026: RD$3,760 spent (incl. SHELL EJEMPLO RD$2,500)');
  assert.ok(!/730%|785%/.test(mail.decoded + mail.text), 'no percentage that counts the fuel as a habit');
  assert.ok(mail.decoded.includes('Includes SHELL EJEMPLO, RD$2,500 · the rest of the day (RD$1,260) +196% vs. a typical day (RD$425)'));
  assert.ok(mail.text.includes('typical day RD$425'));
});

// ---------- rule 1: the typical day is the median of the day-to-day totals of the last 60 days
test('typical day: the median, not the mean — a few big days do not move it', () => {
  const h = load();
  const list = history(h, fill(54, 400).concat(fill(6, 6000)));     // mean RD$960, median RD$400
  list.push({ date: yesterday(h), amount: 1000 });
  const s = summarize(h, list);
  assert.equal(s.typicalDay, 400);
  assert.equal(Math.round(s.vsTypical * 100), 150);
});

test('typical day: over 60 days (not 30), counting days with nothing spent as RD$0, and without bills', () => {
  const h = load();
  // the last 30 days at RD$400, the 30 before at RD$1,000: the median of all 60 is RD$700 (an even count: the middle two)
  const list = history(h, fill(30, 400).concat(fill(30, 1000)));
  for (let i = 1; i <= 60; i += 5) list.push({ date: before(h, i), merchant: 'EDESUR', amount: 9000, category: 'Electricity' });
  list.push({ date: before(h, 90), amount: 50000 });                 // older than 60 days: left out
  list.push({ date: yesterday(h), amount: 700 });
  const s = summarize(h, list);
  assert.equal(s.typicalDay, 700, 'median of 60 days, bills excluded');
  assert.equal(s.historyDays, 60);

  // spending on only 20 of the last 60 days: most days are RD$0, so the typical day is RD$0
  const sparse = history(h, fill(60, 0).map((v, i) => (i % 3 === 2 ? 1000 : 0)));
  sparse.push({ date: yesterday(h), amount: 800 });
  const z = summarize(h, sparse);
  assert.equal(z.historyDays, 60);
  assert.equal(z.typicalDay, 0, 'days with nothing spent count');
});

// ---------- rule 2: the warning comes from the day's rank among the 60 days, not from a percentage
test('high day: only above the 90th percentile of the last 60 days', () => {
  const h = load();
  const steps = Array.from({ length: 60 }, (x, i) => (i + 1) * 100);  // RD$100 … RD$6,000; 90th percentile RD$5,400
  const day = amount => summarize(h, history(h, steps).concat([{ date: yesterday(h), amount }]));
  const high = day(5500), notHigh = day(5300);
  assert.equal(high.highDayThreshold, 5400);
  assert.equal(high.highDay, true);
  assert.ok(high.recommendations.some(r => r.icon === '🟠' && /one of your highest days in the last 60 days/.test(r.text)));
  assert.equal(notHigh.highDay, false, 'well above the typical day (+74%) but not among the top 10%');
  assert.ok(!notHigh.recommendations.some(r => /highest days|above your daily average/.test(r.text)), 'no warning');
});

test('a big fuel day is not a warning when such days are common in your history; it is still named', () => {
  const h = load();
  const amounts = [];
  for (let i = 0; i < 60; i++) amounts.push(i % 6 === 0 ? 4500 : 450);   // 10 big days in 60
  const list = history(h, amounts);
  list.push({ date: yesterday(h), merchant: 'SHELL EJEMPLO', amount: 2500 }, { date: yesterday(h), merchant: 'UBER*EATS', amount: 420 });
  const s = summarize(h, list);
  assert.equal(s.highDay, false);
  assert.ok(!s.recommendations.some(r => r.icon === '🟠' && /highest/.test(r.text)));
  const mail = email(h, list);
  assert.ok(mail.decoded.includes('Includes SHELL EJEMPLO, RD$2,500'));
  assert.ok(!mail.decoded.includes('highest days'), 'the chip is not a warning either');
});

// ---------- rule 3: big one-off purchases are named apart
test('one-off purchases: three times a typical day or more, named apart; smaller ones are not', () => {
  const h = load();
  const list = history(h, fill(60, 500));
  list.push({ date: yesterday(h), merchant: 'FERRETERIA OCHOA', amount: 1500 },  // exactly 3× → named
    { date: yesterday(h), merchant: 'SUPERMERCADO NACIONAL', amount: 1400 },        // 2.8× → part of the rest
    { date: yesterday(h), merchant: 'ALQUILER', amount: 25000, category: 'Rent' }); // bills are never one-off purchases
  const s = summarize(h, list);
  assert.deepEqual(s.sporadic.map(x => x.merchant), ['FERRETERIA OCHOA']);
  assert.equal(s.restDayToDay, 1400);
  assert.equal(s.highDay, true, 'the whole day (RD$2,900) still decides the warning');
  assert.ok(s.recommendations.some(r => /highest days/.test(r.text) && /It includes FERRETERIA OCHOA, RD\$1,500/.test(r.text)));
});

test('a day that is only one big purchase says so instead of comparing an empty rest', () => {
  const h = load();
  const list = history(h, fill(60, 500));
  list.push({ date: yesterday(h), merchant: 'SHELL EJEMPLO', amount: 2500 });
  const mail = email(h, list);
  assert.ok(mail.decoded.includes('Includes SHELL EJEMPLO, RD$2,500 · nothing else day-to-day'));
});

// ---------- rule 4: little history, or a typical day of RD$0, is said in words — no percentage
test('little history: no typical day and no percentage, the email says why', () => {
  const h = load();
  const list = history(h, fill(5, 400));
  list.push({ date: yesterday(h), amount: 4000 });
  const s = summarize(h, list);
  assert.equal(s.historyDays, 5);
  assert.equal(s.typicalDay, null);
  assert.equal(s.vsTypical, null);
  assert.equal(s.highDay, false);
  const mail = email(h, list);
  assert.equal(mail.subject, 'Sun 4 Oct 2026: RD$4,000 spent');
  assert.ok(mail.decoded.includes('Typical day available after 7 days of history'));
  assert.ok(!/% vs\./.test(mail.decoded + mail.text), 'no percentage comparison anywhere');
});

test('a typical day of RD$0: said in words; no percentage, no one-off names, no warning', () => {
  const h = load();
  const list = history(h, fill(60, 0).map((v, i) => (i % 3 === 2 ? 1000 : 0)));
  list.push({ date: yesterday(h), merchant: 'SHELL EJEMPLO', amount: 2500 });
  const s = summarize(h, list);
  assert.equal(s.typicalDay, 0);
  assert.equal(s.vsTypical, null);
  assert.deepEqual(s.sporadic, []);
  assert.equal(s.highDay, false, 'with most days at RD$0, any purchase would be "high"');
  const mail = email(h, list);
  assert.equal(mail.subject, 'Sun 4 Oct 2026: RD$2,500 spent');
  assert.ok(mail.decoded.includes('Most days have no day-to-day spending'));
  assert.ok(!/% vs\./.test(mail.decoded + mail.text), 'no percentage comparison anywhere');
});

test('below a typical day: the good news keeps its percentage (it cannot run away: at most 100%)', () => {
  const h = load();
  const list = history(h, fill(60, 1000));
  list.push({ date: yesterday(h), amount: 400 });
  const s = summarize(h, list);
  assert.equal(Math.round(s.vsTypical * 100), -60);
  assert.ok(s.recommendations.some(r => /^Nice — day-to-day spending was 60% below a typical day \(RD\$1,000\)/.test(r.text)));
  assert.equal(email(h, list).subject, 'Sun 4 Oct 2026: RD$400 spent (−60% vs. typical day)');
});
