'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

// Reported: transfers sent, pasted from Excel into Bank Transfers, were saved WITHOUT a date; typed again after deleting
// them, they were saved a second time.
function book() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"POPULAR":true}'], ['monthlyIncome', 80000], ['incomeCurrency', 'DOP']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  h.ctx.getOrCreateTransfersSheet('Bank Transfers');
  return { h, mock, tx: () => h.plain(mock.ss.getSheetByName('Transactions')._rows(15)), bt: () => mock.ss.getSheetByName('Bank Transfers') };
}
const day = (h, v) => h.ctx.normalizeDateForCompare(v);

test('a yyyy-mm-dd typed as text is that day — not the day before (UTC midnight in Santo Domingo)', () => {
  const h = load();
  assert.equal(h.ctx.normalizeDateForCompare('2026-08-28'), '2026-08-28');
  assert.equal(h.ctx.normalizeDateForCompare(' 2026-01-01 '), '2026-01-01');
});

test('dates pasted as real dates (from Excel) and as text are both saved, on their day', () => {
  const { h, tx, bt } = book();
  bt().getRange(2, 1, 2, 6).setValues([[h.date(2026, 1, 10), 'POPULAR', 'MB a LA CASITA DE B', '', 8745, 'DOP'], ['2026-01-11', 'POPULAR', 'MB a PATRICIA G VEL', '', '650', 'DOP']]);
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.deepEqual(tx().map(r => [day(h, r[0]), r[2]]).sort(), [['2026-01-10', 'MB a LA CASITA DE B'], ['2026-01-11', 'MB a PATRICIA G VEL']]);
});

test('typed again, a row saved without its date gets the date back — not a second copy (the reported duplicates)', () => {
  const { h, mock, tx, bt } = book();
  // what v1.1.57 left: a typed row in Transactions with no date
  const lost = h.ctx.transactionToRow({ date: '', bank: 'POPULAR', merchant: 'MB a LA CASITA DE B', amount: 8745, currency: 'DOP', category: '',
    autoCategory: h.get('AUTO_NONE'), description: 'MB a LA CASITA DE B', subject: 'Added by hand in Bank Transfers', timestamp: 't',
    isCredit: false, isCashback: false, type: 'Transfer', messageId: 'manual:1:out:5', txRef: '' });
  mock.ss.getSheetByName('Transactions').getRange(2, 1, 1, 15).setValues([lost]);
  bt().getRange(3, 1, 1, 6).setValues([[h.date(2026, 1, 10), 'POPULAR', 'MB a LA CASITA DE B', '', 8745, 'DOP']]);
  const r = h.plain(h.ctx.importTypedTransferRows('Bank Transfers', 'Transfer'));
  assert.deepEqual(r, { added: 0, rejected: 0, repaired: 1 });
  assert.deepEqual(tx().map(x => [day(h, x[0]), x[2], x[12]]), [['2026-01-10', 'MB a LA CASITA DE B', 'manual:1:out:5']], 'dated, same row, same Id');
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(tx().length, 1, 'still one');
});

test('whatever loses a date on the way, the check after saving puts it back and logs it', () => {
  const { h, tx, bt } = book();
  h.get('var __realToRow = transactionToRow; transactionToRow = t => { const r = __realToRow(t); r[0] = ""; return r; }');   // simulate the loss
  bt().getRange(2, 1, 1, 6).setValues([[h.date(2026, 2, 3), 'POPULAR', 'MB a 7Bars Mac', '', 1700, 'DOP']]);
  h.ctx.importTypedTransferRows('Bank Transfers', 'Transfer');
  assert.equal(day(h, tx()[0][0]), '2026-02-03');
  assert.ok(h.logs.some(l => /saved without its date and was dated again .*2026-02-03/.test(l)));
});
