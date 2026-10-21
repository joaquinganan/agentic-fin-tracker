'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

// v1.1.72: the QIK credit card statement (PDF) — every purchase, the card payment, the cashback credited and any refund,
// all-or-nothing against the statement's own totals; the purchases already saved from their alerts are not saved again.
// The fixture keeps the real statement's structure (as read by pdfToText) with invented data.

const STMT = () => fixture('statements/qik_card');
const QIK_FROM = 'Qik Banco Digital <no-reply-qik@qik.com.do>';
const STMT_SUBJECT = 'Estado de cuenta de tu Tarjeta de Crédito QIK';

test('the statement is read: cut date, every row, a description on two lines, the totals add up', () => {
  const h = load();
  const r = h.plain(h.ctx.parseQIKCardStatement(STMT()));
  assert.deepEqual(r.problems, []);
  assert.equal(r.cut, '2026-09-05');
  assert.deepEqual(r.entries.map(e => [e.day, e.type, e.amount, e.merchant]), [
    ['2026-08-07', 'Transaction', 20, 'Opret Metro Sd Distrito Naci'],
    ['2026-08-07', 'Cashback', 12.5, 'QIK cashback (statement)'],
    ['2026-08-10', 'Transaction', 1250.75, 'Supermercado Ejemplo Santo Domingo'],
    ['2026-08-10', 'Transaction', 20, 'Opret Metro Sd Distrito Naci'],
    ['2026-08-10', 'Transaction', 20, 'Opret Metro Sd Distrito Naci'],
    ['2026-08-15', 'Transaction', 630.4, 'Cafe Ejemplo Bar Panama City'],
    ['2026-08-22', 'Cashback', 7.5, 'QIK cashback (statement)'],
    ['2026-08-28', 'Card Payment', 5230, 'Pago A Tarjeta - Cuenta De Ahorro Qik'],
    ['2026-09-02', 'Transaction', 845, 'Farmacia Ejemplo Distrito Naci'],
    ['2026-09-04', 'Transaction', 1213.85, 'Tienda Ejemplo Santo Domingo']]);
  assert.match(r.rows[1].description, /Recompensas Qik Rebate Modom$/, 'the line after a row goes on its description');
  assert.equal(r.rows[2].description, 'Supermercado Ejemplo Santo Domingododom', 'a row is never joined to the next row');
  const refs = r.entries.map(e => e.ref);
  assert.equal(new Set(refs).size, refs.length, 'two equal metro rides on one day are two rows');
  assert.ok(refs.includes('20260905:2026-08-10:P:2000:1') && refs.includes('20260905:2026-08-10:P:2000:2'));
  assert.ok(refs.includes('20260905:2026-08-28:Y:523000:1'), 'the payment and a purchase never share a reference');
});

test('all or nothing: a row misread, a total that differs, a row in dollars or no totals → nothing taken, it says why', () => {
  const h = load();
  const bad = [
    [STMT().replace('Supermercado Ejemplo Santo Domingododom RD$ 1,250.75', 'Supermercado Ejemplo Santo Domingododom RD$ 1,205.75'), /rows go from|purchases add up/],
    [STMT().replace('RD$ 4,000.00 RD$ 5,250.00 RD$ 22.40', 'RD$ 4,100.00 RD$ 5,150.00 RD$ 22.40'), /purchases add up to 4000\.00/],
    [STMT().replace('Farmacia Ejemplo Distrito Nacisddom RD$ 845.00', 'Farmacia Ejemplo Distrito Nacisddom US$ 845.00'), /another currency/],
    [STMT().replace('Pago A Tarjeta - Cuenta De Ahorro Qik - RD$', 'Pago A Tarjeta - Cuenta De Ahorro Qik RD$'), /rows go from/],   // a credit read as a purchase
    [STMT().replace(/Fecha de corte: [^\n]+/, ''), /cut date or totals/]];
  bad.forEach(([text, why], i) => {
    const r = h.ctx.parseQIKCardStatement(text);
    assert.match(r.problems.join('; '), why, 'case ' + i);
    assert.equal(r.entries.length, 0, 'case ' + i);
  });
});

test('a refund (a credit that is not the payment or the cashback) is a negative purchase in its merchant\'s category', () => {
  // no real refund seen yet: the shape of a credit row ("- RD$") with a merchant's descriptor
  const h = load();
  const refund = '02/09/2026 03/09/2026 Farmacia Ejemplo Distrito Nacisddom - RD$ 845.00';
  // either way the bank adds it up: in the payments, or taken off the purchases
  const inPayments = STMT().replace('Tienda Ejemplo Santo Domingododom RD$ 1,213.85', 'Tienda Ejemplo Santo Domingododom RD$ 1,213.85\n' + refund)
    .replace('RD$ 4,000.00 RD$ 111.11 5,250.00', 'RD$ 3,155.00 RD$ 111.11 5,250.00').replace('RD$ 4,000.00 RD$ 5,250.00', 'RD$ 4,000.00 RD$ 6,095.00');
  const netted = inPayments.replace('RD$ 4,000.00 RD$ 6,095.00', 'RD$ 3,155.00 RD$ 5,250.00');
  [inPayments, netted].forEach(text => {
    const r = h.plain(h.ctx.parseQIKCardStatement(text));
    assert.deepEqual(r.problems, []);
    const last = r.entries[r.entries.length - 1];
    assert.deepEqual([last.type, last.amount, last.merchant, last.refund], ['Transaction', -845, 'Farmacia Ejemplo Distrito Naci', true]);
  });
});

test('the merchant without the region and country the card network glues to the city', () => {
  const h = load();
  [['Libreria Ejemplo Santo Domingododom', 'Libreria Ejemplo Santo Domingo'],
   ['Opret Metro Sd Distrito Nacisddom', 'Opret Metro Sd Distrito Naci'],
   ['Kiosko Ejemplo 12 Colon Papan', 'Kiosko Ejemplo 12 Colon'],          // a short city: region + country on their own
   ['Netflix.com Los Gatoscausa', 'Netflix.com Los Gatos'],
   ['Pago A Tarjeta - Cuenta De Ahorro Qik -', 'Pago A Tarjeta - Cuenta De Ahorro Qik'],
   ['Random', 'Random'],                                   // one word: nothing to take off
   ['Tienda Ejemplo Bogota', 'Tienda Ejemplo Bogota']]     // no country at the end: as it is
    .forEach(([raw, clean]) => assert.equal(h.ctx.cleanCardStatementMerchant(raw), clean, raw));
});

// ---- the statement email
const statementMessage = (h, id) => ({ getId: () => id || 'm-stmt', getFrom: () => QIK_FROM, getSubject: () => STMT_SUBJECT,
  getDate: () => h.date(2026, 9, 7, 14), getPlainBody: () => 'Hola, estás listo para vivir la experiencia.', getBody: () => '',
  getAttachments: () => [{ getContentType: () => 'application/pdf', getName: () => 'Estado de Cuenta - TC 0000 - 05 sep 2026.pdf', copyBlob: () => ({}) }],
  getThread: () => null });

test('the statement email: each row as the item it is — the payment Exclude, the cashback a credit, purchases categorized', () => {
  const h = load();
  h.ctx.statementPdfText = () => STMT();   // the PDF reader is tested in v1.1.51/52; the real PDF was read by it as this text
  const r = h.plain(h.ctx.parseEmailMessage(statementMessage(h), {}, h.ctx.newParseStats()));
  assert.equal(r.status, 'ok', 'a statement, not filtered by "ESTADO DE CUENTA" or its advertising');
  const by = t => r.items.filter(i => i.type === t);
  assert.equal(by('Transaction').length, 7);
  assert.deepEqual(by('Card Payment').map(i => [i.amount, i.category, i.isCredit]), [[5230, 'Exclude', true]]);
  assert.deepEqual(by('Cashback').map(i => [i.amount, i.isCashback, i.isCredit]), [[12.5, true, true], [7.5, true, true]]);
  assert.equal(r.items.find(i => /Opret/.test(i.merchant)).category, 'Public transport & parking');
  assert.ok(r.items.every(i => /^QIK:STMT:20260905:/.test(i.txRef) && /^m-stmt_20260905-/.test(i.messageId)));
  // a Custom Rule still wins
  const ruled = h.plain(h.ctx.parseEmailMessage(statementMessage(h), { 'Gifts & donations': ['TIENDA EJEMPLO'] }, h.ctx.newParseStats()));
  assert.equal(ruled.items.find(i => /Tienda/.test(i.merchant)).category, 'Gifts & donations');
  // Recategorize keeps every type (the subject says nothing) and the payment's Exclude
  r.items.forEach(i => {
    const row = h.ctx.transactionToRow(Object.assign({ timestamp: 't' }, i));
    assert.equal(h.ctx.computeRecategorization(row, {}, []).type, i.type, i.merchant);
  });
});

// ---- already saved from the alerts
const card = (h, day, amount, extra) => Object.assign({ date: h.date(2026, 8, day), bank: 'QIK', merchant: 'CAFETERIA EJEMPLO', amount: amount,
  currency: 'DOP', type: 'Transaction', category: '', messageId: 'alert-' + day + '-' + amount + '_0', txRef: '' }, extra || {});
const savedValues = (h, items) => [new Array(15).fill('h')].concat(items.map(t => h.ctx.transactionToRow(Object.assign({ description: t.merchant, subject: 's', timestamp: 't' }, t))));
const stmtItems = h => {
  h.ctx.statementPdfText = () => STMT();
  return h.ctx.parseEmailMessage(statementMessage(h), {}, h.ctx.newParseStats()).items;
};

test('reported: purchases without an alert are saved from the statement; those with one are not saved twice', () => {
  const h = load();
  const alerts = [card(h, 7, 20), card(h, 10, 1250.75), card(h, 10, 20),
    // abroad: the alert in dollars, the statement in pesos — the same merchant's first word
    card(h, 15, 10.5, { currency: 'USD', merchant: 'CAFE EJEMPLO BAR PANAMA' })];
  const sel = h.ctx.selectNewTransactions(stmtItems(h), h.ctx.buildExistingIndex(savedValues(h, alerts)));
  assert.deepEqual(h.plain(sel.fresh.map(t => [t.type, t.amount])).sort(), [['Card Payment', 5230], ['Cashback', 12.5], ['Cashback', 7.5],
    ['Transaction', 1213.85], ['Transaction', 20], ['Transaction', 845]].sort(), 'one of the two metro rides on the 10th had no alert');
  assert.equal(sel.duplicates, 4);
  assert.ok(h.logs.some(l => /Only in the statement \(no alert saved\): QIK Transaction DOP 845\.00 2026-09-02/.test(l)));
});

test('the statement read again (every run) saves nothing more, and the alerts it paired with stay paired', () => {
  const h = load();
  const alerts = [card(h, 7, 20), card(h, 10, 20)];
  const first = h.ctx.selectNewTransactions(stmtItems(h), h.ctx.buildExistingIndex(savedValues(h, alerts)));
  const sheet = savedValues(h, alerts.concat(first.fresh));
  const again = h.ctx.selectNewTransactions(stmtItems(h), h.ctx.buildExistingIndex(sheet));
  assert.equal(again.fresh.length, 0);
});

test('statement first, alert later (a run by dates); or both in one run, in either order', () => {
  const h = load();
  const index = h.ctx.buildExistingIndex(savedValues(h, stmtItems(h)));
  assert.equal(h.ctx.selectNewTransactions([Object.assign(card(h, 0, 845), { date: h.date(2026, 9, 2) })], index).fresh.length, 0,
    'the alert of a purchase the statement already saved');
  const empty = () => h.ctx.buildExistingIndex([new Array(15).fill('h')]);
  const alert = Object.assign(card(h, 0, 845), { date: h.date(2026, 9, 3) });
  assert.equal(h.ctx.selectNewTransactions([alert].concat(stmtItems(h)), empty()).fresh.length, 10);
  assert.equal(h.ctx.selectNewTransactions(stmtItems(h).concat([alert]), empty()).fresh.length, 10);
});

test('not the same purchase: another bank, amount, type, sign, more than 3 days apart, after the cut, a hold never charged', () => {
  const h = load();
  const notMatching = [
    Object.assign(card(h, 0, 845), { date: h.date(2026, 9, 2), bank: 'BHD' }),
    Object.assign(card(h, 0, 845.01), { date: h.date(2026, 9, 2) }),
    Object.assign(card(h, 0, 845), { date: h.date(2026, 9, 2), type: 'Transfer' }),
    Object.assign(card(h, 0, -845), { date: h.date(2026, 9, 2) }),
    Object.assign(card(h, 0, 845), { date: h.date(2026, 8, 29) }),
    // within 3 days of the 4th, but after the cut (Sept 5) plus a day: a purchase of the next statement
    Object.assign(card(h, 0, 1213.85), { date: h.date(2026, 9, 7) }),
    // in dollars, another merchant
    card(h, 15, 10.5, { currency: 'USD', merchant: 'TIENDA DUTY FREE' })];
  const index = h.ctx.buildExistingIndex(savedValues(h, notMatching));
  assert.equal(h.ctx.selectNewTransactions(stmtItems(h), index).fresh.length, 10);
  // an Uber hold (v1.1.67) is not what was charged: the statement row is saved, not taken as its copy
  const hold = Object.assign(card(h, 0, 845), { date: h.date(2026, 9, 2), merchant: 'UBER *TRIP', description: 'UBER *TRIP (Uber hold, not charged)' });
  const values = [new Array(15).fill('h'), h.ctx.transactionToRow(Object.assign({ subject: 's', timestamp: 't' }, hold))];
  assert.equal(h.ctx.selectNewTransactions(stmtItems(h), h.ctx.buildExistingIndex(values)).fresh.length, 10);
});

test('two equal alerts are two purchases; equal statement rows pair one alert each', () => {
  const h = load();
  const empty = () => h.ctx.buildExistingIndex([new Array(15).fill('h')]);
  assert.equal(h.ctx.selectNewTransactions([card(h, 10, 20), Object.assign(card(h, 10, 20), { messageId: 'other_0' })], empty()).fresh.length, 2);
  const both = h.ctx.buildExistingIndex(savedValues(h, [card(h, 10, 20), Object.assign(card(h, 10, 20), { messageId: 'other_0' })]));
  const sel = h.ctx.selectNewTransactions(stmtItems(h), both);
  assert.equal(sel.fresh.filter(t => t.amount === 20).length, 1, 'the 7th\'s ride had no alert; the 10th\'s two did');
});

test('an Uber ride from the statement is what was charged: never checked against the receipts (no Gmail search for it)', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  const ride = Object.assign(card(h, 0, 245.6, { merchant: 'Uber Trip Help.uber.com', messageId: 'm-stmt_20260905-2026-09-02-P-24560-1',
    txRef: 'QIK:STMT:20260905:2026-09-02:P:24560:1' }), { date: h.date(2026, 9, 2) });
  mock.ss.getSheetByName('Transactions').appendRow(h.ctx.transactionToRow(Object.assign({ description: ride.merchant, subject: STMT_SUBJECT, timestamp: 't' }, ride)));
  h.ctx.reconcileRideReceipts(h.date(2026, 9, 7));
  assert.deepEqual(mock.gmail.queries, []);
});

test('a whole run: QIK alerts and the statement in the same search; a second run changes nothing', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 80000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', '{"QIK":true}']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  h.ctx.statementPdfText = () => STMT();
  const alert = (id, day, amount, merchant) => ({ getId: () => id, getFrom: () => 'Qik <notificaciones@qik.do>', getSubject: () => 'Usaste tu tarjeta de crédito Qik',
    getDate: () => h.date(2026, 8, day, 11), getPlainBody: () => fixture('qik_consumo').replace(/CAFETERIA EJEMPLO/g, merchant)
      .replace(/RD\$ 640\.00/g, 'RD$ ' + amount.toFixed(2)), getBody: () => '', getThread: () => null });
  mock.gmail.threads.push(fakeThread('t-alerts', [alert('a1', 10, 1250.75, 'SUPERMERCADO EJEMPLO'), alert('a2', 10, 20, 'OPRET METRO SD')]));
  mock.gmail.threads.push(fakeThread('t-stmt', [statementMessage(h, 'm-stmt')]));
  const rows = () => h.plain(mock.ss.getSheetByName('Transactions').getDataRange().getValues().slice(1));
  h.ctx.runGmailMonitorForDateRange('2026-08-01', '2026-09-08');
  assert.equal(rows().length, 2 + 10 - 2, 'two purchases from their alerts, the other eight rows from the statement');
  assert.deepEqual(rows().filter(r => r[11] === 'Card Payment').map(r => [r[3], r[5]]), [[5230, 'Exclude']]);
  h.ctx.runGmailMonitorForDateRange('2026-08-01', '2026-09-08');
  assert.equal(rows().length, 10);
});
