'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

function book(customRules) {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"BANESCO":true}'], ['monthlyIncome', 1000], ['incomeCurrency', 'USD']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  const cr = mock.ss.insertSheet('Custom Rules'); cr.appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  (customRules || []).forEach(r => cr.appendRow(['user@example.com', r[0], r[1], '']));
  return { h, mock, tx: () => mock.ss.getSheetByName('Transactions')._rows(15) };
}
const row = (h, merchant, category, type, auto) => { const x = new Array(15).fill(''); x[0] = h.date(2026, 9, 10); x[1] = 'BANESCO'; x[2] = merchant;
  x[3] = 2500; x[4] = 'DOP'; x[5] = category; x[6] = merchant; x[7] = type === 'Transfer' ? 'Transferencia realizada' : 'Consumo'; x[9] = 'NO';
  x[11] = type; x[12] = 'm-' + merchant.replace(/\s/g, ''); x[14] = auto === undefined ? '' : auto; return x; };

test('recategorize keeps the categories you set by hand (reported: it wiped them)', () => {
  const { h, mock, tx } = book();
  mock.ss.getSheetByName('Transactions').getRange(2, 1, 4, 15).setValues([
    row(h, 'JUAN PEREZ', 'Rent', 'Transfer'),                                              // before v1.1.39: a transfer you categorised
    row(h, 'SUPERMERCADO BRAVO', 'Health + Vet + Pharmacy', 'Transaction', 'Groceries + Barbershop'),   // you changed the tracker's choice
    row(h, 'FARMACIA CAROL', 'Groceries + Barbershop', 'Transaction'),                    // before v1.1.39, a purchase: re-evaluated
    row(h, 'SUPERMERCADO NACIONAL', '', 'Transaction', 'Health + Vet + Pharmacy')]);     // you cleared it: the automatic one comes back
  h.ctx.recategorizeAllTransactions('user@example.com');
  const by = Object.fromEntries(tx().map(r => [r[2], [r[5], r[14]]]));
  assert.deepEqual(by['JUAN PEREZ'], ['Rent', '(none)'], 'kept; the tracker\'s own choice (none) recorded');
  assert.deepEqual(by['SUPERMERCADO BRAVO'], ['Health + Vet + Pharmacy', 'Groceries + Barbershop'], 'kept');
  assert.deepEqual(by['FARMACIA CAROL'], ['Health + Vet + Pharmacy', 'Health + Vet + Pharmacy']);
  assert.deepEqual(by['SUPERMERCADO NACIONAL'], ['Groceries + Barbershop', 'Groceries + Barbershop']);
  // and again: nothing you set moves
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(Object.fromEntries(tx().map(r => [r[2], r[5]]))['JUAN PEREZ'], 'Rent');
  assert.equal(Object.fromEntries(tx().map(r => [r[2], r[5]]))['SUPERMERCADO BRAVO'], 'Health + Vet + Pharmacy');
});

test('a Custom Rule still applies to rows whose category the tracker set; a hand-set one wins over it', () => {
  const { h, mock, tx } = book([['Rent', 'MARIA GOMEZ'], ['Exclude', 'JUAN PEREZ']]);
  mock.ss.getSheetByName('Transactions').getRange(2, 1, 2, 15).setValues([
    row(h, 'MARIA GOMEZ', '', 'Transfer', '(none)'),
    row(h, 'JUAN PEREZ', 'Gym + Calisthenics', 'Transfer', '(none)')]);
  h.ctx.recategorizeAllTransactions('user@example.com');
  const by = Object.fromEntries(tx().map(r => [r[2], r[5]]));
  assert.equal(by['MARIA GOMEZ'], 'Rent', 'the rule fills an empty one');
  assert.equal(by['JUAN PEREZ'], 'Gym + Calisthenics', 'your choice beats a rule added later — clear the cell to let the rule apply');
});

test('a category typed in Bank Transfers reaches Transactions and survives the rebuild and later runs', () => {
  const { h, mock, tx } = book();
  mock.ss.getSheetByName('Transactions').getRange(2, 1, 2, 15).setValues([
    row(h, 'JUAN PEREZ', '', 'Transfer', '(none)'), row(h, 'LUISA DIAZ', '', 'Transfer', '(none)')]);
  h.ctx.recategorizeAllTransactions('user@example.com');           // builds Bank Transfers (with the hidden Id)
  const bt = mock.ss.getSheetByName('Bank Transfers');
  assert.equal(bt.getRange(1, 8).getValue(), 'Id');
  assert.ok(bt.isColumnHiddenByUser(8));
  const at = bt._rows(8).findIndex(r => r[2] === 'JUAN PEREZ') + 2;
  bt.getRange(at, 4).setValue('Rent');                            // you categorise it where you review transfers
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(Object.fromEntries(tx().map(r => [r[2], r[5]]))['JUAN PEREZ'], 'Rent');
  assert.equal(bt._rows(8).find(r => r[2] === 'JUAN PEREZ')[3], 'Rent', 'still there after the rebuild');
  assert.equal(bt._rows(8).find(r => r[2] === 'LUISA DIAZ')[3], '');
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(Object.fromEntries(tx().map(r => [r[2], r[5]]))['JUAN PEREZ'], 'Rent', 'and on every later run');
});

test('an old Bank Transfers sheet (no Id yet) still hands back its edits, matched by date, beneficiary and amount', () => {
  const { h, mock, tx } = book();
  mock.ss.getSheetByName('Transactions').getRange(2, 1, 1, 15).setValues([row(h, 'JUAN PEREZ', '', 'Transfer')]);
  const bt = mock.ss.insertSheet('Bank Transfers');
  bt.getRange(1, 1, 2, 7).setValues([['Date', 'Bank', 'Beneficiary / Description', 'Category', 'Amount', 'Currency', 'Email Subject'],
    [h.date(2026, 9, 10), 'BANESCO', 'JUAN PEREZ', 'Rent', 2500, 'DOP', 'Transferencia realizada']]);
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(tx()[0][5], 'Rent');
});

test('HAPI\'s return can start at its value on an earlier date — not only at the snapshot (reported)', () => {
  const h = load();
  const L = rows => [['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id']]
    .concat(rows.map(r => { const x = r.slice(); x[0] = h.date(...x[0]); while (x.length < 12) x.push(''); return x; }));
  const ledger = L([
    [[2026, 1, 1], 'HAPI', 'Valuation', '', '', '', 5000, '', 'USD'],               // the broker's value on Jan 1 (fictitious)
    [[2026, 3, 1], 'HAPI', 'Deposit', '', '', '', 59500, '', 'DOP'],                // 1,000 US$ at 59.5
    [[2026, 6, 1], 'HAPI', 'Deposit', '', '', '', 500, '', 'USD'],
    [[2026, 9, 26], 'HAPI', 'Snapshot', 'GOOGL', 20, 300, 4000, '', 'USD'],
    [[2026, 9, 26], 'HAPI', 'Snapshot', 'CASH', '', '', 1000, '', 'USD']]);
  const r = h.plain(h.ctx.computeReturns(ledger, { HAPI: 8000 }, { usdRate: 59.5, today: '2026-09-28' }));
  const hapi = r.accounts[0];
  assert.deepEqual([hapi.start, hapi.startValue, hapi.netDeposits], ['2026-01-01', 5000, 1500]);
  assert.equal(hapi.gain, 1500);
  assert.ok(hapi.annualized !== null, 'more than 180 days');
  // the balance is only a starting point: Holdings doesn't list HAPI under Other accounts
  const holdings = h.plain(h.ctx.computeHoldings(ledger, { usdRate: 59.5, year: 2026 }));
  assert.deepEqual(holdings.valuations, []);
  assert.equal(holdings.positions[0].qty, 20);
  // without it, the first snapshot is the start
  const snapOnly = h.plain(h.ctx.computeReturns(L([[[2026, 9, 26], 'HAPI', 'Snapshot', 'CASH', '', '', 1000, '', 'USD']]),
    { HAPI: 1000 }, { usdRate: 59.5, today: '2026-09-28' })).accounts[0];
  assert.equal(snapOnly.start, '2026-09-26');
});

test('Portfolio History: every other day is banded in a colour you can see', () => {
  const mock = makeServices();
  const h = load({ services: mock.services });
  h.ctx.recordPortfolioHistory([{ account: 'HAPI', ticker: 'CASH', qty: '', price: '', value: 10 }], '2026-09-27');
  h.ctx.styleHistorySheet(mock.ss.getSheetByName('Portfolio History'));
  const band = mock.ss.getSheetByName('Portfolio History').cfRules.find(r => /ISEVEN\(INT/.test(r.whenFormulaSatisfied || ''));
  assert.equal(band.setBackground, '#E8EEF8');
});
