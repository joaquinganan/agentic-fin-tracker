'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

function book(rules) {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"POPULAR":true}'], ['monthlyIncome', 80000], ['incomeCurrency', 'DOP']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  const cr = mock.ss.insertSheet('Custom Rules'); cr.appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  (rules || []).forEach(r => cr.appendRow(['user@example.com', r[0], r[1], '']));
  const row = h.ctx.transactionToRow({ date: h.date(2026, 8, 26), bank: 'BANESCO', merchant: 'PAGO RENTA', amount: 30000, currency: 'DOP',
    category: 'Rent', autoCategory: h.get('AUTO_NONE'), description: 'PAGO RENTA', subject: 's', timestamp: 't', isCredit: false,
    isCashback: false, type: 'Transfer', messageId: 'rent-1', txRef: '' });
  mock.ss.getSheetByName('Transactions').getRange(2, 1, 1, 15).setValues([row]);
  h.ctx.recategorizeAllTransactions('user@example.com');   // builds Bank Transfers
  return { h, mock, tx: () => h.plain(mock.ss.getSheetByName('Transactions')._rows(15)), bt: () => mock.ss.getSheetByName('Bank Transfers') };
}
const type = (sheet, cells) => sheet.getRange(sheet.getLastRow() + 1, 1, 1, 6).setValues([cells]);

test('a transfer SENT typed by hand in Bank Transfers is saved (positive), keeps your category, only once (requested)', () => {
  const { h, tx, bt } = book();
  type(bt(), ['2026-01-10', 'popular', 'MB a 0837271691 LA CASITA DE B', 'Groceries', '8745', '']);
  h.ctx.recategorizeAllTransactions('user@example.com');
  const saved = tx().filter(r => r[2] === 'MB a 0837271691 LA CASITA DE B');
  assert.equal(saved.length, 1);
  assert.deepEqual([saved[0][1], saved[0][3], saved[0][4], saved[0][5], saved[0][11], saved[0][9]], ['POPULAR', 8745, 'DOP', 'Groceries', 'Transfer', 'NO']);
  assert.match(saved[0][12], /^manual:\d+:out:/);
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(tx().filter(r => r[2] === 'MB a 0837271691 LA CASITA DE B').length, 1, 'not twice');
  assert.equal(tx().find(r => r[2] === 'MB a 0837271691 LA CASITA DE B')[5], 'Groceries', 'kept by recategorizing');
});

test('typed without a category, a Custom Rule categorizes it like any transfer', () => {
  const { h, tx, bt } = book([['Pharmacy', 'FARMA VALUE']]);
  type(bt(), ['2026-01-03', 'POPULAR', 'MB a 0796973725 FARMA VALUE RD', '', '4,569.17', 'DOP']);
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(tx().find(r => /FARMA VALUE/.test(r[2]))[5], 'Pharmacy');
});

test('rows of a Bank Transfers sheet from before v1.1.39 (no Ids) are not taken for typed ones — nothing is duplicated', () => {
  const { h, tx, bt } = book();
  const sheet = bt();
  sheet.getRange(2, 8).setValue('');                     // an old sheet: the saved rent transfer's row has no Id
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(tx().filter(r => r[2] === 'PAGO RENTA').length, 1);
});

test('a typed row without a date or amount goes to Unrecognized, naming Bank Transfers', () => {
  const { h, mock, tx, bt } = book();
  type(bt(), ['', 'POPULAR', 'MB a someone', '', '500', '']);
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(tx().filter(r => r[2] === 'MB a someone').length, 0);
  const row = h.plain(mock.ss.getSheetByName('Unrecognized')._rows(10)).find(r => r[1] === 'Bank Transfers');
  assert.match(row[3], /type it again in Bank Transfers/);
});
