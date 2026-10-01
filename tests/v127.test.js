'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

// Fictitious days and amounts — nobody's real statement dates.
const CARDS = [
  { bank: 'LAFISE', product: 'LAFISE_CLASICA', closeDay: 10, dueDay: 5 },
  { bank: 'BANESCO', product: 'BANESCO_SUPERCASHBACK', closeDay: 20, dueDay: 12 },
  { bank: 'BHD', product: 'BHD_MIPAIS', closeDay: 1, dueDay: 22 },
  { bank: 'POPULAR', product: 'other', name: 'Visa Oro', cashback: 2, closeDay: 28, dueDay: 18 },
  { bank: 'BDI', product: 'none' }
];

function configured() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  mock.ss.insertSheet('Transactions'); mock.ss.insertSheet('Custom Rules');
  return { h, mock };
}
const input = over => Object.assign({
  email: 'owner@example.com', incomeCurrency: 'USD', monthlyIncome: 3000, otherIncome: 25000, otherIncomeCurrency: 'DOP',
  deductionMode: 'manual', deductions: { ARS: 0, AFP: 0, ISR: 0, taxRate: 10 },
  banksToTrack: { LAFISE: true, BANESCO: true, BHD: true, POPULAR: true, BDI: true }, cards: CARDS,
  notify: { enabled: false, monthly: false, email: '', hour: 8, sections: { totals: true } }, timestamp: 'x'
}, over || {});

test('cards: catalogue products get their name, rate and note; "other" keeps what was typed; "none" is skipped', () => {
  const h = load();
  const cards = h.plain(h.ctx.resolveCards(CARDS));
  assert.deepEqual(cards.map(c => [c.bank, c.name, c.cashback, c.closeDay, c.dueDay]), [
    ['LAFISE', 'Clásica Mastercard', 0.10, 10, 5], ['BANESCO', 'Super Cashback', 0.07, 20, 12],
    ['BHD', 'Mi País', 0.05, 1, 22], ['POPULAR', 'Visa Oro', 0.02, 28, 18]]);
  assert.match(cards[0].note, /^10\/10/);
  assert.equal(cards[3].note, '');
  const rows = h.plain(h.ctx.dashboardCardRows({ cards: CARDS }, null));
  assert.deepEqual(rows[0].slice(0, 4), ['LAFISE · Clásica Mastercard', 0.1, 'Day 10', 'Day 5']);
  assert.deepEqual(rows[3].slice(0, 4), ['POPULAR · Visa Oro', 0.02, 'Day 28', 'Day 18']);
  // before the wizard has cards, an older Dashboard's typed rows are kept
  const legacy = [['BHD', 0.05, 'Day 1', 'Day 22', 'mine'], ['', '', '', '', '']];
  assert.deepEqual(h.plain(h.ctx.dashboardCardRows({}, legacy)), [legacy[0]]);
});

test('setup: other income and cards are validated', () => {
  const h = load();
  assert.equal(h.ctx.validateSetupInput(input()), null);
  assert.match(h.ctx.validateSetupInput(input({ otherIncome: -1 })), /Other income can't be negative/);
  assert.match(h.ctx.validateSetupInput(input({ cards: [{ bank: 'BHD', product: 'BHD_MIPAIS', closeDay: 32 }] })), /between 1 and 31/);
  assert.match(h.ctx.validateSetupInput(input({ cards: [{ bank: 'BHD', product: 'NOPE' }] })), /unknown card/);
  assert.match(h.ctx.validateSetupInput(input({ cards: [{ bank: 'BDI', product: 'other', name: ' ' }] })), /give the other card a name/);
});

test('saving: stored, Dashboard rebuilt with the cards and other income, progress + summary shown', () => {
  const { h, mock } = configured();
  assert.equal(h.ctx.saveSetupConfig(input()), true);
  const c = h.plain(h.ctx.getConfig());
  assert.deepEqual(h.plain(c.otherIncomes), [{ label: '', amount: 25000, currency: 'DOP' }], 'v1.1.59: the single field became a line');
  assert.equal(c.cards.length, 4, '"none" is not stored');
  assert.deepEqual(h.plain(mock.ss.getRangeByName('DASH_CARDS').getValues()).map(r => r.slice(0, 4)), [
    ['LAFISE · Clásica Mastercard', 0.1, 'Day 10', 'Day 5'], ['BANESCO · Super Cashback', 0.07, 'Day 20', 'Day 12'],
    ['BHD · Mi País', 0.05, 'Day 1', 'Day 22'], ['POPULAR · Visa Oro', 0.02, 'Day 28', 'Day 18']]);
  const dash = mock.ss.getSheetByName('Dashboard');
  const formulas = [...dash.cells.values()].filter(v => typeof v === 'string');
  assert.ok(formulas.some(f => f.includes('CFG_OTHER_INCOME_DOP+CFG_OTHER_INCOME_USD*')), 'other income, every currency at its rate (v1.1.59)');
  assert.ok(formulas.some(f => /\+N\d+-N\d+$/.test(f)), 'net income adds other income and takes other deductions off');
  assert.equal(mock.ss.getRangeByName('CFG_OTHER_INCOME_DOP').getValue(), 25000);
  assert.deepEqual(mock.ss.toasts.slice(0, 3), ['Saving your configuration...', 'Preparing sheets and daily triggers...', 'Rebuilding the Dashboard...']);
  const alert = mock.ui.alerts[mock.ui.alerts.length - 1];
  assert.match(alert, /Configuration saved and Dashboard updated/);
  assert.match(alert, /Net income: RD\$185,650 \/ month \(DOP-equivalent\), with 1 other income\(s\)/);   // 3000×0.9×59.5 = 160,650 + 25,000
  assert.match(alert, /Credit cards: LAFISE Clásica Mastercard, BANESCO Super Cashback, BHD Mi País, POPULAR Visa Oro/);
  assert.equal(mock.lock.held, false);
});

test('saving: invalid input is reported even though the window already closed; nothing is written', () => {
  const { h, mock } = configured();
  assert.equal(h.ctx.saveSetupConfig(input({ monthlyIncome: 0 })), false);
  assert.match(mock.ui.alerts[0], /Setup not saved: Monthly income must be/);
  assert.equal(mock.ss.getSheetByName('Configuration'), null);
});

test('net income for the emails includes other income in its own currency', () => {
  const h = load();
  const rates = { USD: 60, EUR: 64, COP: 0.015 };
  const base = { incomeCurrency: 'DOP', monthlyIncome: 100000, deductions: { ARS: 3040, AFP: 2870, ISR: 9000, taxRate: 0 } };
  assert.equal(h.ctx.computeNetIncomeDop(Object.assign({}, base, { otherIncome: 500, otherIncomeCurrency: 'USD' }), rates), 100000 - 14910 + 30000);
  assert.equal(h.ctx.computeNetIncomeDop(Object.assign({}, base, { otherIncome: 20000, otherIncomeCurrency: 'DOP' }), rates), 100000 - 14910 + 20000);
  assert.equal(h.ctx.computeNetIncomeDop(base, rates), 100000 - 14910, 'older configurations: no other income');
});

test('card tips name your card; they come from the wizard\'s cards', () => {
  const { h, mock } = configured();
  h.ctx.saveSetupConfig(input());
  const cards = h.plain(h.ctx.readDashboardCards(mock.ss));
  assert.deepEqual(cards.find(c => c.bank === 'BHD'), { bank: 'BHD', rate: 0.05, name: 'Mi País' });
  const tips = h.plain(h.ctx.summaryCardTips([{ merchant: 'FARMA VALUE FD05', bank: 'POPULAR', amount: 2000, type: 'Transaction' }], cards));
  assert.match(tips[0].text, /pharmacies earn 5% with your BHD Mi País card — about RD\$100 back/);
});

test('Setup Wizard: card section from the catalogue, other income field, closes while saving', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  h.ctx.openSetupWizard();
  const html = mock.ui.dialogs[0].html;
  for (const piece of ['id="incomeLines"', 'id="deductionLines"', 'id="cardsBox"', 'id="saveBtn"', '"name":"Mi País"',
                       '"name":"Clásica Mastercard"', '"name":"Super Cashback"', 'This window will close', 'google.script.host.close()']) {
    assert.ok(html.includes(piece), piece);
  }
  assert.ok(!html.includes('withSuccessHandler(function() {\n              statusEl.className = \'status success\''), 'no longer waits in the window');
});
