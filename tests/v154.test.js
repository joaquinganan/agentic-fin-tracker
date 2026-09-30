'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

function book() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"BANESCO":true}'], ['monthlyIncome', 80000], ['incomeCurrency', 'DOP']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  // categorized by hand (as a rent transfer is: no rule guesses it), so recategorizing keeps it
  const row = (d, type, merchant, amount, category, id) => h.ctx.transactionToRow({ date: h.date(2026, 8, d), bank: 'BANESCO', merchant, amount,
    currency: 'DOP', category, autoCategory: h.get('AUTO_NONE'), description: merchant, subject: 's', timestamp: 't', isCredit: amount < 0,
    isCashback: false, type, messageId: id, txRef: '' });
  const rows = [row(26, 'Transfer', 'PAGO RENTA', 30000, 'Rent', 'rent-1'), row(17, 'Incoming', 'ACH transfer (sender not in the statement)', -6500, '', 'in-1')];
  mock.ss.getSheetByName('Transactions').getRange(2, 1, rows.length, 15).setValues(rows);
  h.ctx.recategorizeAllTransactions('user@example.com');   // builds Bank Transfers and Incoming Transfers
  const tx = () => h.plain(mock.ss.getSheetByName('Transactions')._rows(15));
  return { h, mock, tx, sheet: n => mock.ss.getSheetByName(n) };
}
const rowOf = (sheet, id) => sheet._rows(8).findIndex(r => r[7] === id) + 2;

test('a category typed in Incoming Transfers or Bank Transfers reaches Transactions at once (reported: the Dashboard didn\'t change)', () => {
  const { h, tx, sheet } = book();
  const inc = sheet('Incoming Transfers'), at = rowOf(inc, 'in-1');
  inc.getRange(at, 4).setValue('Rent');
  h.ctx.onEdit({ range: inc.getRange(at, 4) });
  assert.equal(tx().find(r => r[12] === 'in-1')[5], 'Rent', 'no update needed — the Dashboard reads Transactions');
  inc.getRange(at, 4).setValue('');
  h.ctx.onEdit({ range: inc.getRange(at, 4) });
  assert.equal(tx().find(r => r[12] === 'in-1')[5], '', 'clearing it too');
  const bt = sheet('Bank Transfers'), bat = rowOf(bt, 'rent-1');
  bt.getRange(bat, 4).setValue('Housing');
  h.ctx.onEdit({ range: bt.getRange(bat, 4) });
  assert.equal(tx().find(r => r[12] === 'rent-1')[5], 'Housing', 'Bank Transfers as well');
  inc.getRange(at, 3).setValue('someone');
  h.ctx.onEdit({ range: inc.getRange(at, 3) });
  assert.equal(tx().find(r => r[12] === 'in-1')[2], 'ACH transfer (sender not in the statement)', 'other columns are not copied');
});

test('a row typed by hand in Incoming Transfers is saved as money received, keeps your category, and only once', () => {
  const { h, mock, tx, sheet } = book();
  const inc = sheet('Incoming Transfers'), next = inc.getLastRow() + 1;
  inc.getRange(next, 1, 1, 6).setValues([['2026-08-28', '', 'Roommate', 'Rent', '15,000.00', '']]);
  h.ctx.recategorizeAllTransactions('user@example.com');
  const saved = tx().filter(r => r[11] === 'Incoming' && r[2] === 'Roommate');
  assert.equal(saved.length, 1);
  assert.deepEqual([saved[0][1], saved[0][3], saved[0][4], saved[0][5], saved[0][14]], ['MANUAL', -15000, 'DOP', 'Rent', h.get('AUTO_NONE')], 'categorized by you, not by the tracker');
  assert.match(saved[0][12], /^manual:/);
  assert.equal(h.ctx.normalizeDateForCompare(saved[0][0]), '2026-08-28');
  assert.ok(inc._rows(8).some(r => r[2] === 'Roommate' && r[7] === saved[0][12]), 'listed again, now with its Id');
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(tx().filter(r => r[2] === 'Roommate').length, 1, 'not twice');
  assert.equal(tx().find(r => r[2] === 'Roommate')[5], 'Rent', 'your category kept by recategorizing');
  // and it is taken off Rent
  const m = h.plain(h.ctx.computeMonthlySummary(mock.ss.getSheetByName('Transactions').getDataRange().getValues(),
    { today: h.date(2026, 9, 1, 9), rates: { USD: 60, EUR: 64, COP: 0.015 }, netIncomeDop: 80000, cards: [] }));
  assert.equal(m.categories.find(c => c.cat === 'Rent').amount, 15000, '30,000 paid − 15,000 received');
});

test('a typed row is imported once even if the sheet is not rebuilt right after (an update interrupted in between)', () => {
  const { h, tx, sheet } = book();
  const inc = sheet('Incoming Transfers'), next = inc.getLastRow() + 1;
  inc.getRange(next, 1, 1, 6).setValues([['2026-08-29', 'LAFISE', 'Roommate', 'Electricity', '2500', 'DOP']]);
  assert.deepEqual(h.plain(h.ctx.importTypedIncomingRows()), { added: 1, rejected: 0, repaired: 0 });
  assert.deepEqual(h.plain(h.ctx.importTypedIncomingRows()), { added: 0, rejected: 0, repaired: 0 }, 'the row now carries its Id');
  assert.equal(tx().filter(r => r[2] === 'Roommate').length, 1);
});

test('a typed row without a date or an amount is not lost: Unrecognized keeps what was typed (no Gmail link)', () => {
  const { h, mock, tx, sheet } = book();
  const inc = sheet('Incoming Transfers'), next = inc.getLastRow() + 1;
  inc.getRange(next, 1, 1, 6).setValues([['', '', 'Roommate', 'Electricity', '2500', '']]);
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(tx().filter(r => r[2] === 'Roommate').length, 0);
  const row = h.plain(mock.ss.getSheetByName('Unrecognized')._rows(10)).find(r => r[1] === 'Incoming Transfers');
  assert.ok(row, 'listed');
  assert.match(row[3], /needs a date \(yyyy-mm-dd\) and an amount/);
  assert.match(row[4], /Roommate · Electricity · 2500/);
  assert.equal(row[5], '', 'no Gmail link — it is not an email');
});
