'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

const RATES = { USD: 60, EUR: 65, COP: 0.015 };
const base = { incomeCurrency: 'DOP', monthlyIncome: 100000, deductionMode: 'auto', deductions: {} };

test('several other incomes, each in its currency, and other deductions — net income converts every line (requested)', () => {
  const h = load();
  const net = h.ctx.computeNetIncomeDop(Object.assign({}, base, {
    otherIncomes: [{ label: 'Freelance', amount: 500, currency: 'USD' }, { label: 'Rent', amount: 15000, currency: 'DOP' }, { amount: 100, currency: 'EUR' }],
    otherDeductions: [{ label: 'Loan', amount: 8000, currency: 'DOP' }, { label: 'Insurance', amount: 20, currency: 'USD' }] }), RATES);
  const salary = h.ctx.computeNetIncomeDop(base, RATES);
  assert.equal(net, salary + 500 * 60 + 15000 + 100 * 65 - 8000 - 20 * 60);
  assert.deepEqual(h.plain(h.ctx.moneyLineTotals([{ amount: 1, currency: 'USD' }, { amount: 2.5, currency: 'USD' }, { amount: 0, currency: 'DOP' }])),
    { DOP: 0, USD: 3.5, EUR: 0 }, 'empty lines are ignored');
});

test('the wizard checks the lines: amounts of 0 or more, DOP/USD/EUR, at most 10', () => {
  const h = load();
  const input = extra => Object.assign({ email: 'u@example.com', incomeCurrency: 'DOP', monthlyIncome: 50000, deductionMode: 'manual',
    deductions: { ARS: 0, AFP: 0, ISR: 0 }, banksToTrack: { BHD: true }, cards: [] }, extra);
  assert.equal(h.ctx.validateSetupInput(input({ otherIncomes: [{ amount: 10, currency: 'EUR' }], otherDeductions: [{ amount: 5, currency: 'USD' }] })), null);
  assert.match(h.ctx.validateSetupInput(input({ otherIncomes: [{ amount: -1, currency: 'DOP' }] })), /Other income amounts must be numbers of 0 or more/);
  assert.match(h.ctx.validateSetupInput(input({ otherDeductions: [{ amount: 5, currency: 'COP' }] })), /Other deduction currency must be DOP, USD, EUR/);
  assert.match(h.ctx.validateSetupInput(input({ otherIncomes: new Array(11).fill({ amount: 1, currency: 'DOP' }) })), /Up to 10 other incomes/);
});

function book() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  return { h, mock };
}

test('saving: the lines and their totals per currency are stored; the Dashboard adds incomes, takes deductions off', () => {
  const { h, mock } = book();
  h.ctx.saveSetupConfig({ email: 'u@example.com', incomeCurrency: 'DOP', monthlyIncome: 100000, deductionMode: 'auto', deductions: {},
    banksToTrack: { BHD: true }, cards: [], notify: {},
    otherIncomes: [{ label: 'Freelance', amount: 500, currency: 'USD' }, { label: 'Rent', amount: 15000, currency: 'DOP' }],
    otherDeductions: [{ label: 'Loan', amount: 8000, currency: 'DOP' }] });
  const c = h.plain(h.ctx.getConfig());
  assert.deepEqual(c.otherIncomes.map(x => [x.label, x.amount, x.currency]), [['Freelance', 500, 'USD'], ['Rent', 15000, 'DOP']]);
  assert.deepEqual(c.otherDeductions.map(x => [x.label, x.amount]), [['Loan', 8000]]);
  const name = n => mock.ss.getRangeByName(n).getValue();
  assert.deepEqual([name('CFG_OTHER_INCOME_DOP'), name('CFG_OTHER_INCOME_USD'), name('CFG_OTHER_INCOME_EUR'), name('CFG_OTHER_DED_DOP')], [15000, 500, 0, 8000]);
  const dash = mock.ss.getSheetByName('Dashboard');
  assert.equal(dash.getRange(18, 10).getValue(), 'Other deductions (DOP-equivalent)');
  assert.match(String(dash.getRange(19, 14).getFormula ? dash.getRange(19, 14).getFormula() : dash.getRange(19, 14).getValue()), /\+N17-N18$/);
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /with 2 other income\(s\) and 1 other deduction\(s\)/);
});

test('a sheet set up before keeps its single other income — on the Dashboard too, without opening the wizard again', () => {
  const { h, mock } = book();
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'u@example.com'], ['incomeCurrency', 'DOP'], ['monthlyIncome', 100000], ['otherIncome', 25000],
   ['otherIncomeCurrency', 'USD'], ['banksToTrack', '{"BHD":true}']].forEach(r => cfg.appendRow(r));
  assert.deepEqual(h.plain(h.ctx.getConfig().otherIncomes), [{ label: '', amount: 25000, currency: 'USD' }]);
  h.ctx.ensureConfigNamedRanges();
  assert.equal(mock.ss.getRangeByName('CFG_OTHER_INCOME_USD').getValue(), 25000, 'not 0: taken from the older setup');
  assert.equal(mock.ss.getRangeByName('CFG_OTHER_DED_DOP').getValue(), 0);
});

test('a password-protected statement PDF is named as such (reason shown in Unrecognized), not a confusing parse error', () => {
  const h = load();
  const pdf = Buffer.from('%PDF-1.4\n1 0 obj << /Type /Catalog >> endobj\ntrailer << /Root 1 0 R /Encrypt 2 0 R >>\n%%EOF\n');
  assert.throws(() => h.ctx.pdfToText(new Uint8Array(pdf)), /password-protected/);
});
