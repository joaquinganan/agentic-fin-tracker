'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

// Real samples shared by another user (LAFISE, QIK, BDI, Scotiabank, Banreservas); the fixtures keep their structure with
// invented data.
const EMAILS = {
  lafisePayment: { subject: '    Notificación de pago de tarjeta de crédito', from: '"Banco LAFISE" <bancanet@notificaciones.lafise.com>', body: 'lafise_pago_tarjeta' },
  lafiseDetail: { subject: 'Detalle de Transaccion Tarjeta de Crédito ', from: 'Alerta de Consumo Banco LAFISE <notificacioneslafisedo@lafise.com.do>', body: 'lafise_consumo_detalle_tc' },
  qikCash: { subject: 'Retiro con Código CASH exitoso', from: 'Qik Banco Digital <no-reply-qik@qik.com.do>', body: 'qik_codigo_cash' },
  bdiIn: { subject: 'BDI Digital: Comprobante transacción Interbancaria Beneficiario', from: '"bdinforma@bdi.com.do" <bdinforma@bdi.com.do>', body: 'bdi_transferencia_recibida' },
  scotiaIn: { subject: 'Pago al Instante recibido', from: 'Alertas Scotiabank <alertas@scotiabank.com>', body: 'scotiabank_pago_instante' },
  brTransfer: { subject: 'Recibo de la transacción', from: '<NotificacionesTuBancoApp@banreservas.com>', body: 'banreservas_transferencia_tercero' },
  brCash: { subject: 'Recibo de la transacción', from: '<NotificacionesTuBancoApp@banreservas.com>', body: 'banreservas_retiro_tuefectivo' },
  brIn: { subject: 'Notificaciones Banreservas', from: 'notificaciones@banreservas.com', body: 'banreservas_transferencia_recibida' }
};
const parseBody = (h, e, body, id) => h.plain(h.ctx.parseEmailMessage(h.fakeMessage({ subject: e.subject, from: e.from, body: body,
  date: h.date(2026, 9, 27, 10), id: id || 'm' }), {}, h.ctx.newParseStats()));
const parse = (h, e, id) => parseBody(h, e, fixture(e.body), id);
const row = t => [t.bank, t.type, t.currency, t.amount, t.merchant, t.category];

test('LAFISE card payment (a sender that was never searched): a Card Payment, Exclude, the card in the name; only Exitoso counts', () => {
  const h = load();
  const r = parse(h, EMAILS.lafisePayment);
  assert.equal(r.status, 'ok');
  assert.deepEqual(r.items.map(row), [['LAFISE', 'Card Payment', 'DOP', 2345.67, 'Pago de Tarjeta de Crédito XXXX-1111', 'Exclude']],
    'it used to be read by the generic guesser, with "Concepto: … Monto: DOP 3" as the name');
  const failed = parseBody(h, EMAILS.lafisePayment, fixture('lafise_pago_tarjeta').replace('Exitoso', 'Rechazado'));
  assert.equal(failed.status, 'filtered', 'a payment that did not go through is not saved');
  const q = h.get('BANK_PATTERNS').LAFISE.searchQuery;
  assert.match(q, /from:bancanet@notificaciones\.lafise\.com/);
  assert.match(q, /from:notificacioneslafisedo@lafise\.com\.do/);
});

test('LAFISE second card template: the purchase in DOP; a currency word it does not know is not guessed', () => {
  const h = load();
  assert.deepEqual(parse(h, EMAILS.lafiseDetail).items.map(r => row(r).slice(0, 5)), [['LAFISE', 'Transaction', 'DOP', 1234.56, 'SUPERMERCADO EJEMPLO SANTO DOMINGODO']]);
  assert.deepEqual(h.plain(h.ctx.extractLAFISECardDetailTransactions(fixture('lafise_consumo_detalle_tc').replace('PESOS DOMI', 'DOLARES EST'))).map(x => x.currency), ['USD']);
  assert.deepEqual(h.plain(h.ctx.extractLAFISECardDetailTransactions(fixture('lafise_consumo_detalle_tc').replace('PESOS DOMI', 'XYZ'))), []);
  assert.equal(h.ctx.detectTypeFromSubject(EMAILS.lafiseDetail.subject), 'Transaction', 'Recategorize keeps it a purchase');
});

test('QIK Código CASH: a cash withdrawal, kept despite the security footer; one not Exitoso is not saved', () => {
  const h = load();
  const r = parse(h, EMAILS.qikCash);
  assert.equal(r.status, 'ok', 'reported: filtered as a security-code email');
  assert.deepEqual(r.items.map(x => row(x).slice(0, 5)), [['QIK', 'Transaction', 'DOP', 1500, 'Código Cash (cash withdrawal)']]);
  assert.equal(parseBody(h, EMAILS.qikCash, fixture('qik_codigo_cash').replace('Exitoso', 'Fallido')).status, 'filtered');
  // the purchase template still reads as before
  assert.deepEqual(h.plain(h.ctx.extractQIKConsumoTransactions(fixture('qik_consumo'))).map(x => [x.amount, x.merchant]), [[640, 'CAFETERIA EJEMPLO']]);
});

test('BDI transfer RECEIVED: money in (negative, Incoming), not a transfer sent; yours is Exclude; Recategorize keeps it', () => {
  const h = load();
  const r = parse(h, EMAILS.bdiIn);
  assert.equal(r.status, 'ok');
  assert.deepEqual(r.items.map(row), [['BDI', 'Incoming', 'DOP', -3500, 'MARIA ELENA GOMEZ', 'Exclude']],
    'reported: saved as a Transfer (spending) with a garbled name');
  assert.equal(r.items[0].description, 'MARIA ELENA GOMEZ (own account)');
  const third = parseBody(h, EMAILS.bdiIn, fixture('bdi_transferencia_recibida').replace(/Pagado Por\n\*?MARIA ELENA GOMEZ\*?/, 'Pagado Por\n*JUAN PEREZ*'));
  assert.deepEqual(third.items.map(row), [['BDI', 'Incoming', 'DOP', -3500, 'JUAN PEREZ', '']]);
  // its subject is the same as a sent one's ("… Interbancaria" → Transfer): the saved type must survive Recategorize
  const TX = h.get('TX_COL');
  const saved = []; saved[TX.SUBJECT] = EMAILS.bdiIn.subject; saved[TX.MERCHANT] = 'X'; saved[TX.DESCRIPTION] = 'X (own account)';
  saved[TX.TYPE] = 'Incoming'; saved[TX.CURRENCY] = 'DOP';
  assert.equal(h.plain(h.ctx.computeRecategorization(saved, [])).type, 'Incoming');
});

test('BDI is strict: an email of its own that no extractor knows goes to Unrecognized, never saved as a purchase', () => {
  const h = load();
  const other = { subject: 'BDI Digital: Información de tu cuenta', from: 'bdinforma@bdi.com.do' };
  const r = parseBody(h, other, 'Tu cuenta de ahorro está activa. Límite diario de transferencias RD$50,000.00. Depósito mínimo RD$1,000.00.');
  assert.equal(r.status, 'failed');
  assert.deepEqual(r.items, [], 'reported: the amounts of non-purchase BDI emails were saved as purchases');
  assert.match(r.reason, /nothing guessed/);
  const welcome = parseBody(h, other, '¡Bienvenida a BDI! Tu cuenta puede recibir desde RD$500.00 y tu límite diario es RD$25,000.00.');
  assert.equal(welcome.status, 'filtered', 'a welcome email is not a transaction');
  const opening = parseBody(h, other, 'Tu primer depósito puede ser desde RD$500.00.');
  assert.equal(opening.status, 'filtered', 'nor is one about the first deposit');
  // a real BDI purchase that says "bienvenido" is still read
  const buy = parseBody(h, { subject: 'Notificacion de Consumos', from: 'BDIinforma@bdi.com.do' }, 'Bienvenido. ' + fixture('bdi_consumo'));
  assert.deepEqual(buy.items.map(x => x.amount), [1850]);
  for (const b of ['BDI', 'SCOTIABANK', 'QIK', 'BANRESERVAS']) assert.equal(h.get('BANK_PATTERNS')[b].strict, true, b);
});

test('Scotiabank Pago al Instante received: Incoming, named after the sending bank', () => {
  const h = load();
  const r = parse(h, EMAILS.scotiaIn);
  assert.equal(r.status, 'ok', 'it was "Amount not found"');
  assert.deepEqual(r.items.map(row), [['SCOTIABANK', 'Incoming', 'DOP', -1250, 'Pago al Instante desde Banco BHD', '']]);
  assert.equal(h.ctx.detectTypeFromSubject('Pago al Instante recibido'), 'Incoming', 'by its subject (Recategorize)');
  const other = parseBody(h, { subject: 'Crédito a su cuenta', from: EMAILS.scotiaIn.from }, fixture('scotiabank_pago_instante'));
  assert.deepEqual(other.items.map(x => x.type), ['Incoming'], 'and by its body, whatever the subject');
});

test('Banreservas: a transfer sent (and its tax as its own row), a TuEfectivo withdrawal, a transfer received', () => {
  const h = load();
  const t = parse(h, EMAILS.brTransfer);
  assert.equal(t.status, 'ok');
  assert.deepEqual(t.items.map(row), [['BANRESERVAS', 'Transfer', 'DOP', 750, 'JUAN PEREZ', ''],
    ['BANRESERVAS', 'Transfer', 'DOP', 1.13, 'Banreservas: impuesto y comisión de transferencia', '']]);
  assert.equal(t.items[0].txRef, 'BANRESERVAS:100000000001');
  const c = parse(h, EMAILS.brCash);
  assert.deepEqual(c.items.map(x => row(x).slice(0, 5)), [['BANRESERVAS', 'Transaction', 'DOP', 1000, 'Retiro TuEfectivo (cash withdrawal)']],
    'empty Comisión/Impuestos make no fee row');
  assert.ok(!JSON.stringify(c.items).includes('8090000000'), 'the phone number is not kept');
  assert.equal(h.ctx.extractBANRESERVASWithdrawalTransactions(fixture('banreservas_retiro_tuefectivo')).length, 1, 'no zero fee row from the extractor');
  const withFee = h.plain(h.ctx.extractBANRESERVASWithdrawalTransactions(fixture('banreservas_retiro_tuefectivo').replace('Comisión:', 'Comisión:\nDOP 25.00')));
  assert.deepEqual(withFee.map(x => x.amount), [1000, 25]);
  const i = parse(h, EMAILS.brIn);
  assert.deepEqual(i.items.map(row), [['BANRESERVAS', 'Incoming', 'DOP', -2500, 'JUAN PEREZ', '']], 'not a Transfer, though its body says "Transferencia Recibida"');
  // an app receipt of another kind is not guessed
  const other = parseBody(h, EMAILS.brCash, fixture('banreservas_retiro_tuefectivo').replace('Retiro TuEfectivo', 'Pago de Servicio'));
  assert.equal(other.status, 'failed');
  assert.deepEqual(other.items, []);
  assert.equal(h.ctx.detectTypeFromSubject('Recibo de la transacción'), null, 'Recategorize keeps the type read from the body');
});

function book(banks) {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 80000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify(banks)]].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  return { h, mock };
}
const push = (h, mock, list) => list.forEach((e, i) => mock.gmail.threads.push(fakeThread('t' + i,
  [h.fakeMessage({ subject: e.subject, from: e.from, body: fixture(e.body), date: h.date(2026, 9, 10 + i, 10), id: 'm' + i })])));

test('a run: every new format saved; the summary counts by bank', () => {
  const { h, mock } = book({ LAFISE: true, BDI: true, SCOTIABANK: true, QIK: true, BANRESERVAS: true });
  push(h, mock, Object.values(EMAILS));
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const tx = h.plain(mock.ss.getSheetByName('Transactions')._rows(15));
  assert.deepEqual(tx.map(r => r[1] + ' ' + r[3]).sort(), ['BANRESERVAS -2500', 'BANRESERVAS 1.13', 'BANRESERVAS 1000', 'BANRESERVAS 750',
    'BDI -3500', 'LAFISE 1234.56', 'LAFISE 2345.67', 'QIK 1500', 'SCOTIABANK -1250']);
  assert.ok(mock.ss.getSheetByName('Raw_BANRESERVAS'));
  const summary = mock.ui.alerts.pop();
  assert.match(summary, /🏦 Parsed by bank: LAFISE 2 · BDI 1 · SCOTIABANK 1 · QIK 1 · BANRESERVAS 4/);
  assert.doesNotMatch(summary, /NOT ticked/);
});

test('a bank with emails but not ticked in the wizard is named in the summary (reported: Scotiabank purchases never showed)', () => {
  const { h, mock } = book({ LAFISE: true });
  push(h, mock, [EMAILS.lafiseDetail, { subject: 'Autorización fuera del país', from: 'Alertas Scotiabank <alertas@scotiabank.com>', body: 'scotiabank_consumo' }]);
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const tx = h.plain(mock.ss.getSheetByName('Transactions')._rows(15));
  assert.deepEqual(tx.map(r => r[1]), ['LAFISE'], 'an unticked bank is still not read');
  const summary = mock.ui.alerts.pop();
  assert.match(summary, /Emails found from banks NOT ticked in the Setup Wizard: SCOTIABANK\./);
  assert.doesNotMatch(summary, /NOT ticked[^\n]*(BDI|QIK|BANRESERVAS)/, 'only banks that do have emails');
});
