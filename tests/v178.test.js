'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

// v1.1.78 — the LAFISE Visa Gold in the card catalogue (its 10/10 terms: other categories, minimums and caps than the
// Clásica's, RD$10,000/month in total), several cards per bank in the Setup Wizard, and "Which card for what"
// following the LAFISE card you set up. Statement and payment days below are fictitious.

const input = cards => ({
  email: 'owner@example.com', incomeCurrency: 'DOP', monthlyIncome: 95000, deductionMode: 'manual',
  deductions: { ARS: 0, AFP: 0, ISR: 0, taxRate: 0 }, banksToTrack: { LAFISE: true, BANESCO: true, BHD: true }, cards: cards,
  notify: { enabled: false, monthly: false, email: '', hour: 8, sections: { totals: true } }, timestamp: 'x'
});
const GOLD = { bank: 'LAFISE', product: 'LAFISE_VISA_GOLD', closeDay: 3, dueDay: 25 };
const CLASICA = { bank: 'LAFISE', product: 'LAFISE_CLASICA', closeDay: 14, dueDay: 7 };

function configured() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  mock.ss.insertSheet('Transactions'); mock.ss.insertSheet('Custom Rules');
  return { h, mock };
}

test('the Visa Gold is in the catalogue with its public terms', () => {
  const h = load();
  const [card] = h.plain(h.ctx.resolveCards([GOLD]));
  assert.deepEqual([card.bank, card.name, card.cashback, card.closeDay, card.dueDay], ['LAFISE', 'Visa Gold', 0.10, 3, 25]);
  assert.match(card.note, /RD\$10,000\/month in total/);
  assert.match(card.note, /supermarkets, gas, Amazon/);
});

test('several cards from one bank: validated, saved, and each one on the Dashboard', () => {
  const { h, mock } = configured();
  const v = cards => h.ctx.validateSetupInput(input(cards));
  assert.equal(v([CLASICA, GOLD]), null);
  assert.match(v([GOLD, Object.assign({}, GOLD, { closeDay: 9 })]), /LAFISE: the same card is listed twice/);
  const other = name => ({ bank: 'LAFISE', product: 'other', name: name, cashback: 1 });
  assert.equal(v([GOLD, other('Infinite')]), null);
  assert.match(v([other('Infinite'), other(' infinite ')]), /the same card is listed twice/);
  assert.match(v([GOLD, CLASICA, other('A'), other('B'), other('C')]), /at most 4 cards per bank/);
  assert.equal(v([GOLD, CLASICA, other('A'), other('B'), { bank: 'LAFISE', product: 'none' }]), null, '"No credit card" rows don\'t count');

  assert.equal(h.ctx.saveSetupConfig(input([CLASICA, GOLD, { bank: 'BHD', product: 'BHD_MIPAIS' }])), true);
  assert.deepEqual(h.plain(h.ctx.getConfig().cards).map(c => c.product), ['LAFISE_CLASICA', 'LAFISE_VISA_GOLD', 'BHD_MIPAIS']);
  assert.deepEqual(h.plain(mock.ss.getRangeByName('DASH_CARDS').getValues()).map(r => r.slice(0, 4)), [
    ['LAFISE · Clásica Mastercard', 0.1, 'Day 14', 'Day 7'], ['LAFISE · Visa Gold', 0.1, 'Day 3', 'Day 25'], ['BHD · Mi País', 0.05, '—', '—']]);
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /Credit cards: LAFISE Clásica Mastercard, LAFISE Visa Gold, BHD Mi País/);
});

test('cashback tips with two cards from one bank use the better rate (not the last one listed)', () => {
  const h = load();
  const cards = [{ bank: 'BHD', rate: 0.08, name: 'Better' }, { bank: 'BHD', rate: 0.05, name: 'Mi País' }];
  const tips = h.plain(h.ctx.summaryCardTips([{ merchant: 'FARMA VALUE FD05', bank: 'POPULAR', amount: 2000, type: 'Transaction' }], cards));
  assert.match(tips[0].text, /earn 8% with your BHD Better card — about RD\$160 back/);
});

test('monthly "cashback left on the table" with two cards from one bank: the better rate, and its name', () => {
  const h = load();
  const row = new Array(14).fill('');
  row[0] = h.date(2026, 8, 10); row[1] = 'POPULAR'; row[2] = 'FARMA VALUE FD05'; row[3] = 2000; row[4] = 'DOP';
  row[5] = 'Health & pharmacy'; row[11] = 'Transaction';
  const m = h.plain(h.ctx.computeMonthlySummary([new Array(14).fill('h'), row], { today: h.date(2026, 9, 1, 9),
    rates: { USD: 60, EUR: 64, COP: 0.015 }, netIncomeDop: 80000,
    cards: [{ bank: 'BHD', rate: 0.08, name: 'Better' }, { bank: 'BHD', rate: 0.05, name: 'Mi País' }] }));
  assert.equal(m.missed.total, 160);
  assert.equal(Object.values(m.missed.byRule)[0].best, 'BHD Better');
});

test('"Which card for what" follows your LAFISE card', () => {
  const h = load();
  const mx = cards => h.plain(h.ctx.cashbackMatrixFor(cards === undefined ? null : { cards: cards }));
  const row = (m, name) => m.rows.find(r => r[0] === name);
  const categories = h.plain(h.ctx.getCategories());

  // no configuration, no LAFISE card, the Clásica: the table as before
  for (const m of [mx(), mx([]), mx([CLASICA])]) {
    assert.equal(m.header, 'LAFISE 10/10');
    assert.deepEqual(m.rows, h.plain(h.get('CASHBACK_MATRIX')));
    assert.deepEqual(m.notes, h.plain(h.get('CASHBACK_NOTES')));
  }

  const gold = mx([GOLD]);
  assert.equal(gold.header, 'LAFISE Visa Gold 10/10');
  assert.equal(row(gold, 'Supermarkets')[2], '10% if ≥ RD$10,000 · cap RD$4,000');
  assert.match(row(gold, 'Supermarkets')[5], /^LAFISE once the cycle reaches RD\$10,000; otherwise BANESCO/);
  assert.match(row(gold, 'Restaurants')[2], /≥ RD\$3,000 \(with fast food\)/);
  for (const name of ['Amazon', 'Insurance', 'Laundries']) assert.ok(row(gold, name), name);
  assert.ok(!row(gold, 'Taxi apps (Uber, DiDi)') && !row(gold, 'Gyms'), 'not in the Gold\'s program');
  assert.match(gold.notes[0], /^LAFISE Visa Gold 10\/10 .*Total cap RD\$10,000\/month/);
  assert.equal(gold.notes.length, h.get('CASHBACK_NOTES').length);

  const both = mx([CLASICA, GOLD]);
  assert.equal(both.header, 'LAFISE 10/10 (Clásica · Gold)');
  assert.equal(row(both, 'Restaurants')[2], 'Clásica: 10% if ≥ RD$5,000 · cap RD$1,000\nGold: 10% if ≥ RD$3,000 (with fast food) · cap RD$1,000 (shared)');
  assert.equal(row(both, 'Supermarkets')[2], 'Gold: 10% if ≥ RD$10,000 · cap RD$4,000', 'only the Gold earns there');
  assert.equal(row(both, 'Gas stations')[2], '10% if ≥ RD$5,000 · cap RD$1,000', 'same terms: once');
  assert.match(row(both, 'Taxi apps (Uber, DiDi)')[2], /^Clásica: 10% if ≥ RD\$3,000/);
  assert.match(row(both, 'Taxi apps (Uber, DiDi)')[5], /^LAFISE/);
  assert.equal(both.rows.filter(r => /^Everything else/.test(r[0])).length, 1);
  assert.equal(both.notes.length, h.get('CASHBACK_NOTES').length + 1, 'both programs\' notes');

  for (const m of [gold, both]) {
    assert.ok(/^Everything else/.test(m.rows[m.rows.length - 1][0]), 'the catch-all row stays last');
    for (const r of m.rows.slice(0, -1)) assert.ok(categories.includes(r[6]), r[0] + ' → ' + r[6]);
    assert.equal(new Set(m.rows.map(r => r[0])).size, m.rows.length, 'each kind of spending once');
  }
});

test('the Dashboard shows the table of your LAFISE card, and gives room to cells naming both', () => {
  let { h, mock } = configured();
  h.ctx.saveSetupConfig(input([GOLD]));
  let dash = mock.ss.getSheetByName('Dashboard');
  let all = [...dash.cells.values()];
  assert.ok(all.includes('LAFISE Visa Gold 10/10'));
  assert.ok(all.includes('Laundries\n→ Shopping & home') && all.includes('10% if ≥ RD$10,000 · cap RD$4,000'));
  assert.ok(!all.includes('Taxi apps (Uber, DiDi)\n→ Ride apps & taxis'));

  ({ h, mock } = configured());
  h.ctx.saveSetupConfig(input([CLASICA, GOLD]));
  dash = mock.ss.getSheetByName('Dashboard');
  all = [...dash.cells.values()];
  assert.ok(all.includes('LAFISE 10/10 (Clásica · Gold)'));
  const tall = Object.keys(dash.rowHeights).filter(r => dash.rowHeights[r] === 56).length;
  const combined = h.plain(h.ctx.cashbackMatrixFor({ cards: [CLASICA, GOLD] })).rows.filter(r => /\n/.test(r[2])).length;
  assert.ok(combined > 0 && tall === combined, tall + ' tall rows for ' + combined + ' cells naming both cards');
});

test('Setup Wizard: a row per card, "+ Another card" per bank, and the same card twice refused before saving', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  h.ctx.openSetupWizard();
  const html = mock.ui.dialogs[0].html;
  for (const piece of ['"name":"Visa Gold"', 'const MAX_CARDS_PER_BANK = 4;', 'function addCardRow(bank)', 'class="c-add"',
                       '+ Another \' + b + \' card', 'the same card is listed twice']) {
    assert.ok(html.includes(piece), piece);
  }
});
