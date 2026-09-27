'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');

const DAY = 24 * 60 * 60 * 1000;

test('date range includes the end date (C1)', () => {
  const h = load();
  const r = h.ctx.buildRangeBounds('2026-08-01', '2026/08/31');
  assert.equal(r.start.getDate(), 1);
  assert.equal(r.endExclusive.getMonth(), 8, 'end bound is Sept 1 00:00 → all of Aug 31 is included');
  assert.equal(r.endExclusive.getDate(), 1);
  assert.equal(r.endExclusive - r.start, 31 * DAY);
  assert.throws(() => h.ctx.buildRangeBounds('2026-09-10', '2026-09-01'), /after end date/);
});

test('daily window never leaves a gap at month boundaries (C1)', () => {
  const h = load();
  const w = d => h.plain(h.ctx.computeDailyWindow(d));
  assert.deepEqual(w(h.date(2026, 9, 30, 6)), { startStr: '2026-09-01', endStr: '2026-09-30' }, 'today is included');
  assert.deepEqual(w(h.date(2026, 10, 1, 6)), { startStr: '2026-09-01', endStr: '2026-10-01' },
    'on the 1st, the whole previous month is still covered');
  assert.deepEqual(w(h.date(2026, 10, 2, 6)), { startStr: '2026-10-01', endStr: '2026-10-02' });
  assert.deepEqual(w(h.date(2027, 1, 1, 6)), { startStr: '2026-12-01', endStr: '2027-01-01' }, 'year boundary');
});

test('dedup: messageId, legacy rows and same-run repeats (E1)', () => {
  const h = load();
  const COLS = 14;
  const row = (date, bank, amount, id, ref) => {
    const r = new Array(COLS).fill('');
    r[0] = date; r[1] = bank; r[3] = amount; r[12] = id; r[13] = ref || '';
    return r;
  };
  const existing = [
    new Array(COLS).fill('header'),
    row(h.date(2026, 9, 2), 'LAFISE', 297.25, 'm1_0'),
    row(h.date(2026, 3, 1), 'BHD', 500, ''),         // legacy row, no MessageId
    row(h.date(2026, 9, 5), 'BANESCO', 1000, 'g1_0', 'BANESCO:E000000.A0000001')
  ];
  const index = h.ctx.buildExistingIndex(existing);
  const tx = (id, date, bank, amount, txRef) => ({ messageId: id, date, bank, amount, txRef });
  const incoming = [
    tx('m1_0', h.date(2026, 9, 2), 'LAFISE', 297.25),   // same id → dup
    tx('m9_0', h.date(2026, 3, 1), 'BHD', 500),         // matches legacy row → dup
    tx('m2_0', h.date(2026, 9, 5), 'LAFISE', 189.19),   // new
    tx('m3_0', h.date(2026, 9, 5), 'LAFISE', 189.19),   // same day + amount, DIFFERENT email → new (v1.1.13)
    tx('m2_0', h.date(2026, 9, 5), 'LAFISE', 189.19),   // repeated within this run → dup
    tx('g2_0', h.date(2026, 9, 5), 'BANESCO', 1000, 'BANESCO:E000000.A0000001'), // same bank id, 2nd email → dup (v1.1.23)
    tx('g3_0', h.date(2026, 9, 6), 'BHD', 250, 'BHD:M1'),
    tx('g4_0', h.date(2026, 9, 6), 'BHD', 250, 'BHD:M1')  // same bank id twice in one run → dup
  ];
  const sel = h.ctx.selectNewTransactions(incoming, index);
  assert.equal(sel.duplicates, 5);
  assert.deepEqual(h.plain(sel.fresh).map(t => t.messageId), ['m2_0', 'm3_0', 'g3_0']);
  assert.equal(h.ctx.transactionToRow(sel.fresh[0]).length, h.get('TX_NUM_COLS'), 'a saved row fills every Transactions column (v1.1.39: 15, with Auto Category)');
});

test('recategorize repairs the CACHAREPA row and never touches the merchant (C3, C5)', () => {
  const h = load();
  const COL = h.get('TX_COL');
  const mkRow = fields => {
    const r = new Array(13).fill('');
    Object.keys(fields).forEach(k => { r[COL[k]] = fields[k]; });
    return r;
  };

  const cacharepa = mkRow({ SUBJECT: 'BHD Notificación de Transacciones', MERCHANT: 'CACHAREPA CHURCHILL',
    DESCRIPTION: 'CACHAREPA CHURCHILL', TYPE: 'Transfer', CATEGORY: '', CURRENCY: 'DOP', IS_CREDIT: 'NO' });
  const r1 = h.plain(h.ctx.computeRecategorization(cacharepa, {}));
  assert.equal(r1.type, 'Transaction');
  assert.equal(r1.category, 'Dining/Delivery + Entertainment + Other');
  assert.equal('merchant' in r1, false, 'recategorize has no way to rewrite the merchant any more');

  const bdi = mkRow({ SUBJECT: 'Aviso BDI', MERCHANT: 'X', TYPE: 'Transfer', CURRENCY: 'DOP' });
  assert.equal(h.ctx.computeRecategorization(bdi, {}).type, 'Transfer', 'unknown subject keeps the saved Type');

  const refund = mkRow({ SUBJECT: 'Notificación de Consumo', MERCHANT: 'X', TYPE: 'Cashback', CURRENCY: 'DOP' });
  const r3 = h.plain(h.ctx.computeRecategorization(refund, {}));
  assert.equal(r3.type, 'Cashback');
  assert.equal(r3.isCashback, 'YES');

  const payment = mkRow({ SUBJECT: 'Joaquín, ¡Realizaste un pago a tu tarjeta LAFISE!', MERCHANT: 'LAFISE',
    TYPE: 'Transaction', CATEGORY: 'Dining/Delivery + Entertainment + Other', CURRENCY: 'FKR', IS_CREDIT: 'NO' });
  const r4 = h.plain(h.ctx.computeRecategorization(payment, {}));
  assert.equal(r4.type, 'Card Payment');
  assert.equal(r4.category, 'Exclude', 'card payments are explicitly excluded (v1.1.23)');
  assert.equal(r4.isCredit, 'YES');
  assert.equal(r4.currency, 'DOP', 'invalid currency code is repaired');

  const selfTransfer = mkRow({ SUBJECT: 'Transacciones entre mis productos', MERCHANT: 'BHD Transfer (to account ...1111)',
    TYPE: 'Transfer', CURRENCY: 'DOP' });
  assert.equal(h.ctx.computeRecategorization(selfTransfer, { Exclude: ['1111'] }).category, 'Exclude');
});

test('run summary reports counters and warnings (E7)', () => {
  const h = load();
  const stats = Object.assign(h.plain(h.ctx.newParseStats()),
    { messagesSeen: 12, promotional: 2, declined: 1, amountNotFound: 1, placeholders: 1 });
  const text = h.ctx.buildRunSummary({
    search: { capped: true }, threads: new Array(10).fill({}), transactions: new Array(8).fill({}), stats,
    results: { success: 5, duplicates: 3, failed: 0 }, marked: null, recatChanged: 4, errors: ['recategorize — boom']
  });
  assert.match(text, /Email threads found: 10 \(12 message\(s\) in range\)/);
  assert.match(text, /Saved: 5 \| Duplicates: 3 \| Failed: 0/);
  assert.match(text, /promotional 2 · declined 1/);
  assert.match(text, /Could not parse: 1 email\(s\) — left UNREAD/);
  assert.match(text, /placeholder merchant: 1/);
  assert.match(text, /2000-thread cap/);
  assert.match(text, /recategorize — boom/);
});

test('reversals pair with their original and net to zero (v1.1.23)', () => {
  const h = load();
  const COL = h.get('TX_COL');
  const tx = (o) => Object.assign({ bank: 'BHD', currency: 'DOP', type: 'Transaction', date: h.date(2026, 5, 6) }, o);
  const batch = [
    tx({ merchant: 'PedidosYa*Som Cafe', category: 'Dining/Delivery + Entertainment + Other', amount: 488, timeKey: '06/05/2026 11:47 am' }),
    tx({ merchant: h.get('REVERSAL_UNMATCHED'), category: '', amount: -488, reversal: true, timeKey: '06/05/2026 11:47 am' }),
    tx({ merchant: h.get('REVERSAL_UNMATCHED'), category: '', amount: -488, reversal: true, timeKey: '06/05/2026 11:47 am' }) // a 2nd reversal can't reuse it
  ];
  const unmatched = h.ctx.resolveReversals(batch, [[]]);
  assert.equal(unmatched, 1);
  assert.equal(batch[1].merchant, 'PedidosYa*Som Cafe');
  assert.equal(batch[1].description, 'PedidosYa*Som Cafe (reversal)');
  assert.equal(batch[1].category, 'Dining/Delivery + Entertainment + Other');
  assert.equal(batch[0].amount + batch[1].amount, 0);
  assert.equal(batch[2].merchant, h.get('REVERSAL_UNMATCHED'));

  // original already saved on an earlier run: matched by bank + amount + day
  const saved = new Array(14).fill('');
  saved[COL.DATE] = h.date(2026, 5, 6); saved[COL.BANK] = 'BHD'; saved[COL.AMOUNT] = 488;
  saved[COL.MERCHANT] = 'PedidosYa*Som Cafe'; saved[COL.CATEGORY] = 'Dining/Delivery + Entertainment + Other';
  const later = [tx({ merchant: h.get('REVERSAL_UNMATCHED'), category: '', amount: -488, reversal: true, timeKey: '06/05/2026 11:47 am' })];
  assert.equal(h.ctx.resolveReversals(later, [new Array(14).fill('h'), saved]), 0);
  assert.equal(later[0].category, 'Dining/Delivery + Entertainment + Other');

  // recategorize keeps a paired reversal in its original's category, and leaves an unmatched one blank
  const row = f => { const r = new Array(14).fill(''); Object.keys(f).forEach(k => { r[COL[k]] = f[k]; }); return r; };
  const subject = 'BHD Notificación de Transacciones';
  assert.equal(h.ctx.computeRecategorization(row({ SUBJECT: subject, MERCHANT: 'PedidosYa*Som Cafe',
    DESCRIPTION: 'PedidosYa*Som Cafe (reversal)', AMOUNT: -488, TYPE: 'Transaction', CURRENCY: 'DOP' }), {}).category,
    'Dining/Delivery + Entertainment + Other');
  assert.equal(h.ctx.computeRecategorization(row({ SUBJECT: subject, MERCHANT: h.get('REVERSAL_UNMATCHED'),
    DESCRIPTION: h.get('REVERSAL_UNMATCHED'), AMOUNT: -488, TYPE: 'Transaction', CURRENCY: 'DOP' }), {}).category, '');
});
