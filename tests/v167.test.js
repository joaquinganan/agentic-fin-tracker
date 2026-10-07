'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

// v1.1.67: Uber authorizes an estimate when a ride is requested; when the fare changes it authorizes the final amount
// (or charges less without a new alert). The bank emails every authorization as a purchase. Uber's trip receipt says
// what was really charged, to which card and when the ride was requested: each ride alert is checked against it.
// Shapes from real LAFISE alerts and Uber receipts (Sept–Oct 2026); every amount, time and name here is invented.

const RECEIPT_FROM = 'Uber Receipts <noreply@uber.com>';
const LAFISE_CARD = 'LAFISE CF <notificaciones@bancolafise.com>';
const M = 60000;

// ---------- the receipt ----------
test('Uber trip receipt: when it was requested, and what each card was charged (the card and its bank)', () => {
  const h = load();
  const r = h.plain(h.ctx.parseUberTripReceipt(h.ctx.htmlToPlainText(fixture('uber_trip_receipt_html'))));
  assert.equal(h.ctx.normalizeDateForCompare(new Date(r.requested)), '2026-10-08');
  assert.equal(new Date(r.requested).getHours() + ':' + new Date(r.requested).getMinutes(), '18:12');
  assert.deepEqual(r.payments, [{ last4: '1234', bank: 'LAFISE', currency: 'DOP', amount: 245.6 }]);
});

test('the receipt as Gmail may flatten it (cells on one line), with Uber Cash and a card, paid in cash, or not a receipt', () => {
  const h = load();
  const p = text => h.plain(h.ctx.parseUberTripReceipt(text));
  const flat = 'Oct 8, 2026 6:12 PM Thanks for riding Total DOP 245.60 Payments Mastercard ••••1234 (LAFISE) DOP 245.60 10/8/26 6:41 PM';
  assert.deepEqual(p(flat).payments, [{ last4: '1234', bank: 'LAFISE', currency: 'DOP', amount: 245.6 }]);
  const split = p('Sep 21, 2026\n4:15 PM\nTotal\nRD$ 1,318.40\nPayments\nUber Cash\nRD$ 200.00\nVisa ••••9876 (BHD León)\nRD$ 1,118.40\n');
  assert.deepEqual(split.payments, [{ last4: '9876', bank: 'BHD', currency: 'DOP', amount: 1118.4 }], 'only the card; Uber Cash is not a bank');
  assert.equal(new Date(split.requested).getHours(), 16);
  assert.deepEqual(p('Oct 1, 2026 12:05 AM Total DOP 99.00 Payments Cash DOP 99.00').payments, []);
  assert.equal(new Date(p('Oct 1, 2026 12:05 AM Total DOP 99.00').requested).getHours(), 0, '12 AM is midnight');
  assert.deepEqual(p('Mastercard ••••1234 (OTHER BANK) DOP 10.00 Oct 1, 2026 1:00 PM').payments, [], 'a bank the tracker does not know');
  assert.equal(p('Your weekly summary — no trips'), null);
});

// ---------- the matching (pure) ----------
const at = (hh, mm) => new Date(2026, 9, 8, hh, mm).getTime();
const alert = (id, hh, mm, amount, extra) => Object.assign({ id, bank: 'LAFISE', currency: 'DOP', amount, at: at(hh, mm), state: '' }, extra || {});
const receipt = (h1, m1, h2, m2, amount, extra) => Object.assign({ bank: 'LAFISE', currency: 'DOP', amount, requested: at(h1, m1), sent: at(h2, m2) }, extra || {});
const decide = (h, alerts, receipts) => h.plain(h.ctx.matchRideReceipts(alerts, receipts));

test('reported: the estimate and the final fare were both saved — the estimate is a hold, the final fare is the charge', () => {
  const h = load();
  const r = decide(h, [alert('est', 18, 12, 230.10), alert('fin', 18, 41, 245.60)], [receipt(18, 12, 18, 41, 245.60)]);
  assert.deepEqual(r.decisions, { est: { kind: 'hold' }, fin: { kind: 'charge' } });
  assert.deepEqual(r.unmatched, []);
});

test('one alert of the final fare: the charge, nothing excluded; a lower final fare with no new alert: the amount is corrected', () => {
  const h = load();
  assert.deepEqual(decide(h, [alert('a', 15, 10, 150)], [receipt(15, 10, 15, 48, 150)]).decisions, { a: { kind: 'charge' } });
  assert.deepEqual(decide(h, [alert('b', 18, 30, 260)], [receipt(18, 30, 18, 57, 248.30)]).decisions,
    { b: { kind: 'adjust', amount: 248.3, was: 260 } });
});

test('requests that never became a trip, just before one that did, are holds; an alert far from any receipt is left alone', () => {
  const h = load();
  const r = decide(h, [alert('try1', 18, 12, 120.40), alert('try2', 18, 24, 120.30), alert('ride', 18, 30, 260), alert('lone', 13, 0, 120.30),
    alert('later', 21, 0, 120.30)], [receipt(18, 30, 18, 57, 248.30)]);
  assert.deepEqual(r.decisions, { try1: { kind: 'hold' }, try2: { kind: 'hold' }, ride: { kind: 'adjust', amount: 248.3, was: 260 } });
  assert.deepEqual(r.unmatched, ['lone', 'later'], 'no receipt near them (before or after the ride): counted as charged, and said in the log');
});

test('the estimate equal to the fare (an earlier "UBR* PENDING" alert, then the charge): the later one is the charge', () => {
  const h = load();
  const r = decide(h, [alert('pending', 18, 12, 245.60, { merchant: 'UBR* PENDING.UBER.COM' }), alert('fin', 18, 41, 245.60)],
    [receipt(18, 12, 18, 41, 245.60)]);
  assert.deepEqual(r.decisions, { pending: { kind: 'hold' }, fin: { kind: 'charge' } });
  assert.ok(h.get('RIDE_MERCHANT_RE').test('UBR* PENDING.UBER.COM') && h.get('RIDE_MERCHANT_RE').test('UBER RIDES-*UBER RIDES SANTO DOMINGO DOM'));
  assert.ok(!h.get('RIDE_MERCHANT_RE').test('UBER*EATS SANTO DOMINGO DOM') && !h.get('RIDE_MERCHANT_RE').test('UBER EATS-W*UBER EATS-'));
});

test('back-to-back rides: the second ride\'s only alert belongs to the second receipt, not to the first ride as a hold', () => {
  const h = load();
  const r = decide(h, [alert('h1', 18, 0, 180), alert('f1', 18, 25, 190.50), alert('r2', 18, 20, 99.99)],
    [receipt(18, 0, 18, 25, 190.50), receipt(18, 20, 18, 40, 99.99)]);
  assert.deepEqual(r.decisions, { h1: { kind: 'hold' }, f1: { kind: 'charge' }, r2: { kind: 'charge' } });
});

test('another bank, another currency, an alert already settled, a duplicated receipt: nothing guessed', () => {
  const h = load();
  // the receipt was paid with a BHD card: LAFISE alerts are not its charges
  assert.deepEqual(decide(h, [alert('x', 18, 12, 245.60)], [receipt(18, 12, 18, 41, 245.60, { bank: 'BHD' })]).decisions, {});
  assert.deepEqual(decide(h, [alert('x', 18, 12, 245.60, { currency: 'USD' })], [receipt(18, 12, 18, 41, 245.60)]).decisions, {});
  // settled before (a later run): no decision again
  assert.deepEqual(decide(h, [alert('est', 18, 12, 230.10, { state: 'hold' }), alert('fin', 18, 41, 245.60, { state: 'charged' })],
    [receipt(18, 12, 18, 41, 245.60)]).decisions, {});
  // a hold settled before is never turned into the corrected charge of a receipt
  assert.deepEqual(decide(h, [alert('est', 18, 12, 230.10, { state: 'hold' })], [receipt(18, 12, 18, 41, 245.60)]).decisions, {});
  // Uber sends some receipts twice: the copy must not turn the estimate into a second, corrected charge
  assert.deepEqual(decide(h, [alert('est', 18, 12, 230.10), alert('fin', 18, 41, 245.60)],
    [receipt(18, 12, 18, 41, 245.60), receipt(18, 12, 18, 42, 245.60)]).decisions, { est: { kind: 'hold' }, fin: { kind: 'charge' } });
});

// ---------- on the sheet, with Gmail ----------
const lafiseBody = (merchant, amount) => 'Servicio de Alerta - Nuevo Consumo\nEstimado cliente:\nComercio/Ciudad/País:\n' + merchant +
  '\nFecha:\n06/10/2026 20:05\nTarjeta:\nXXXX-XXXX-XXXX-1234\nAutorización:\n000000\nMonto:\nDOP ' + amount.toFixed(2) + '\n';
function gmailMessage(h, id, from, subject, date, body, html) {
  return { getId: () => id, getFrom: () => from, getSubject: () => subject, getDate: () => date, getPlainBody: () => body,
    getBody: () => html || '', getThread: () => null };
}

function book(alerts, receiptMessages) {
  const messages = [];
  const holder = { threads: [] };
  const mock = makeServices({ ui: true, threads: holder.threads });
  const h = load({ services: mock.services });
  mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"LAFISE":true}']].forEach(r => mock.ss.getSheetByName('Configuration').appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  const items = [];
  alerts.forEach(a => {
    const msg = gmailMessage(h, a.id, LAFISE_CARD, 'Servicio de Alerta - Nuevo Consumo', a.date, lafiseBody(a.merchant || 'UBER*RIDES SANTO DOMINGO DOM', a.amount));
    messages.push(msg);
    items.push(...h.ctx.parseEmailMessage(msg, {}, h.ctx.newParseStats()).items);
  });
  h.ctx.saveTransactions(items);
  holder.threads.push(fakeThread('t-bank', messages));
  (receiptMessages || []).forEach((m, i) => holder.threads.push(fakeThread('t-uber-' + i, [m])));
  const tx = () => h.plain(mock.ss.getSheetByName('Transactions').getDataRange().getValues().slice(1));
  return { h, mock, tx };
}
const receiptMsg = (h, id, sent, html) => gmailMessage(h, id, RECEIPT_FROM, 'Your Tuesday evening trip with Uber', sent, '', html);

test('reported case on the sheet: estimate excluded, final fare kept, rides add up to what Uber charged; a second run changes nothing', () => {
  const html = fixture('uber_trip_receipt_html');
  const pre = load();
  const d = (hh, mm) => pre.date(2026, 10, 8, hh, mm);
  const { h, mock, tx } = book([{ id: 'm-est', date: d(18, 12), amount: 230.10 }, { id: 'm-fin', date: d(18, 41), amount: 245.60 },
    { id: 'm-eats', date: d(18, 13), amount: 230.10, merchant: 'UBER*EATS SANTO DOMINGO DOM' }], []);
  mock.gmail.threads.push(fakeThread('t-uber', [receiptMsg(h, 'r1', h.date(2026, 10, 8, 18, 41), html)]));
  const r = h.plain(h.ctx.reconcileRideReceipts(h.date(2026, 10, 9, 6, 0)));
  assert.deepEqual([r.receipts, r.charged, r.holds, r.adjusted, r.unmatched], [1, 1, 1, 0, 0]);
  assert.equal(r.holdsTotal, 230.1);
  const rows = tx();
  const est = rows.find(x => x[3] === 230.1 && /RIDES/.test(x[2])), fin = rows.find(x => x[3] === 245.6);
  assert.equal(est[5], 'Exclude');
  assert.match(est[6], /\(Uber hold, not charged\)$/);
  assert.match(fin[6], /\(Uber receipt\)$/);
  assert.notEqual(fin[5], 'Exclude');
  const eats = rows.find(x => /EATS/.test(x[2]));
  assert.doesNotMatch(eats[6], /Uber (hold|receipt)/, 'Uber Eats is not a ride');
  const counted = rows.filter(x => /RIDES/.test(x[2]) && x[5] !== 'Exclude').reduce((s, x) => s + x[3], 0);
  assert.equal(Math.round(counted * 100) / 100, 245.6, 'what Uber charged');
  // the end-of-run Recategorize keeps the hold out; and a second run decides nothing new
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(tx().find(x => x[3] === 230.1 && /RIDES/.test(x[2]))[5], 'Exclude');
  const settled = tx();
  const again = h.plain(h.ctx.reconcileRideReceipts(h.date(2026, 10, 9, 6, 0)));
  assert.deepEqual([again.charged, again.holds, again.adjusted], [0, 0, 0]);
  assert.deepEqual(tx(), settled, 'nothing moved');
});

test('a lower final fare without a new alert: the saved amount becomes what Uber charged, and says what the alert said', () => {
  const pre = load();
  const { h, mock, tx } = book([{ id: 'm-b', date: pre.date(2026, 10, 8, 18, 12), amount: 260 }], []);
  const html = fixture('uber_trip_receipt_html');   // requested 6:12 PM, charged 245.60
  mock.gmail.threads.push(fakeThread('t-uber', [receiptMsg(h, 'r1', h.date(2026, 10, 8, 18, 41), html)]));
  const r = h.plain(h.ctx.reconcileRideReceipts(h.date(2026, 10, 9, 6, 0)));
  assert.equal(r.adjusted, 1);
  const row = tx()[0];
  assert.equal(row[3], 245.6);
  assert.match(row[6], /\(Uber receipt: DOP 245\.60, the alert said 260\.00\)$/);
  assert.notEqual(row[5], 'Exclude');
});

test('a hold put back by hand stays yours; no receipts in Gmail: nothing changes; a row typed by hand is never looked up', () => {
  const pre = load();
  const html = fixture('uber_trip_receipt_html');
  const { h, mock, tx } = book([{ id: 'm-est', date: pre.date(2026, 10, 8, 18, 12), amount: 230.10 }, { id: 'm-fin', date: pre.date(2026, 10, 8, 18, 41), amount: 245.60 }], []);
  const before = tx();
  assert.deepEqual(h.plain(h.ctx.reconcileRideReceipts(h.date(2026, 10, 9, 6, 0))).receipts, 0);
  assert.deepEqual(tx(), before, 'without receipts nothing is guessed');
  mock.gmail.threads.push(fakeThread('t-uber', [receiptMsg(h, 'r1', h.date(2026, 10, 8, 18, 41), html)]));
  h.ctx.reconcileRideReceipts(h.date(2026, 10, 9, 6, 0));
  const sheet = mock.ss.getSheetByName('Transactions');
  const i = tx().findIndex(x => x[3] === 230.1);
  sheet.getRange(i + 2, 6).setValue('Transportation');   // you say it was charged
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(tx()[i][5], 'Transportation');
  // a typed row has no Gmail message: it is skipped, not an error
  const typed = h.ctx.transactionToRow({ date: h.date(2026, 10, 8), bank: 'LAFISE', merchant: 'UBER*RIDES', amount: 50, currency: 'DOP',
    category: 'Transportation', description: 'UBER*RIDES', subject: 'Added by hand', timestamp: 't', type: 'Transaction', messageId: 'manual:1:out:9' });
  sheet.appendRow(typed);
  assert.doesNotThrow(() => h.ctx.reconcileRideReceipts(h.date(2026, 10, 9, 6, 0)));
});

test('the run summary says what the receipts changed', () => {
  const h = load();
  const base = { search: {}, threads: [], transactions: [], stats: h.ctx.newParseStats(), results: { success: 0, duplicates: 0, failed: 0 },
    recatChanged: 0, errors: [] };
  const text = h.ctx.buildRunSummary(Object.assign({}, base, { rides: { receipts: 3, charged: 2, holds: 2, adjusted: 1, unmatched: 1, holdsTotal: 350.5, currency: 'DOP' } }));
  assert.match(text, /🚕 Uber: 3 trip receipt\(s\) checked · 2 authorization\(s\) never charged left out \(DOP 350\.50\) · 1 amount\(s\) corrected · 1 ride alert\(s\) without a receipt kept/);
  assert.doesNotMatch(h.ctx.buildRunSummary(base), /Uber/);
});

test('a whole run: ride alerts read, then checked against the receipt before Recategorize; the summary says it', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 80000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', '{"LAFISE":true}']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  const alertMsg = (id, hh, mm, amount) => gmailMessage(h, id, LAFISE_CARD, 'Servicio de Alerta - Nuevo Consumo', h.date(2026, 10, 8, hh, mm),
    lafiseBody('UBER*RIDES SANTO DOMINGO DOM', amount));
  mock.gmail.threads.push(fakeThread('t-bank', [alertMsg('m-est', 18, 12, 230.10), alertMsg('m-fin', 18, 41, 245.60)]));
  mock.gmail.threads.push(fakeThread('t-uber', [receiptMsg(h, 'r1', h.date(2026, 10, 8, 18, 41), fixture('uber_trip_receipt_html'))]));
  h.ctx.runGmailMonitorForDateRange('2026-10-01', '2026-10-09');
  const rows = h.plain(mock.ss.getSheetByName('Transactions').getDataRange().getValues().slice(1));
  assert.deepEqual(rows.map(r => [r[3], r[5] === 'Exclude']).sort((a, b) => a[0] - b[0]), [[230.1, true], [245.6, false]]);
  assert.match(mock.ui.alerts.pop(), /🚕 Uber: 1 trip receipt\(s\) checked · 1 authorization\(s\) never charged left out \(DOP 230\.10\)/);
});
