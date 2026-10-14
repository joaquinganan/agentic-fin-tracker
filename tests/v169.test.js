'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');

// v1.1.69: another user's Unrecognized sheet (19 rows). Real samples of each format; the fixtures keep their structure
// with invented data.
const SCOTIA = '"Alertas Scotiabank" <alertas@scotiabank.com>';
const EMAILS = {
  scotiaOwn: { subject: 'Transferencias a cuentas propias', from: SCOTIA, body: 'scotiabank_transferencia_propia' },
  scotiaThird: { subject: 'Transferencias a terceros', from: SCOTIA, body: 'scotiabank_transferencia_terceros' },
  bdiOnline: { subject: 'BDI Online: Comprobante de Transacción', from: '"bdinforma@bdi.com.do" <bdinforma@bdi.com.do>', body: 'bdi_online_pago_tarjeta' },
  brLoan: { subject: 'Recibo de la transacción', from: '<NotificacionesTuBancoApp@banreservas.com>', body: 'banreservas_pago_prestamo' }
};
const parseBody = (h, e, body) => h.plain(h.ctx.parseEmailMessage(h.fakeMessage({ subject: e.subject, from: e.from, body: body,
  date: h.date(2026, 9, 14, 12), id: 'm' }), {}, h.ctx.newParseStats()));
const parse = (h, e) => parseBody(h, e, fixture(e.body));
const row = t => [t.bank, t.type, t.currency, t.amount, t.merchant, t.category];

test('Scotiabank transfer to your own credit card: a Card Payment, Exclude, the card in the name; Recategorize keeps it', () => {
  const h = load();
  const r = parse(h, EMAILS.scotiaOwn);
  assert.equal(r.status, 'ok', 'reported: Unrecognized, "Format not known yet"');
  assert.deepEqual(r.items.map(row), [['SCOTIABANK', 'Card Payment', 'DOP', 1250.4, 'Pago de Tarjeta de Crédito ***2222', 'Exclude']]);
  // no subject rule says Transfer, so the saved type survives Recategorize
  const TX = h.get('TX_COL');
  const saved = []; saved[TX.SUBJECT] = EMAILS.scotiaOwn.subject; saved[TX.MERCHANT] = r.items[0].merchant;
  saved[TX.DESCRIPTION] = r.items[0].description; saved[TX.TYPE] = 'Card Payment'; saved[TX.CURRENCY] = 'DOP';
  const re = h.plain(h.ctx.computeRecategorization(saved, []));
  assert.deepEqual([re.type, re.category], ['Card Payment', 'Exclude']);
});

test('Scotiabank transfer to your own loan: money out, no category; to another account of yours: Exclude', () => {
  const h = load();
  const loan = parseBody(h, EMAILS.scotiaOwn, fixture(EMAILS.scotiaOwn.body).replace('Tarjeta de Crédito ***2222', 'Préstamo ***3333'));
  assert.deepEqual(loan.items.map(row), [['SCOTIABANK', 'Transfer', 'DOP', 1250.4, 'Pago de préstamo ***3333', '']]);
  const own = parseBody(h, EMAILS.scotiaOwn, fixture(EMAILS.scotiaOwn.body).replace('Tarjeta de Crédito ***2222', 'Cuenta de Ahorro ***4444'));
  assert.deepEqual(own.items.map(row), [['SCOTIABANK', 'Transfer', 'DOP', 1250.4, 'Transferencia a Cuenta de Ahorro ***4444', 'Exclude']]);
  assert.match(own.items[0].description, /\(own account\)$/);
  const usd = parseBody(h, EMAILS.scotiaOwn, fixture(EMAILS.scotiaOwn.body).replace('$1,250.40 DOP', '$80.00 USD'));
  assert.deepEqual(usd.items.map(x => [x.currency, x.amount]), [['USD', 80]]);
});

test('Scotiabank transfer to a third party: a Transfer, DOP when no currency is written; who it was is not in the email', () => {
  const h = load();
  const r = parse(h, EMAILS.scotiaThird);
  assert.equal(r.status, 'ok');
  assert.deepEqual(r.items.map(row), [['SCOTIABANK', 'Transfer', 'DOP', 4500, 'Transferencia a terceros (Scotiabank)', '']]);
  // still strict: another Scotiabank email nobody knows is not guessed
  const other = parseBody(h, EMAILS.scotiaThird, 'Hola ANA M, su pago de $1,000.00 fue programado.');
  assert.equal(other.status, 'failed');
});

test('BDI Online card payment (receipt in columns): the amount paid is Exclude, the tax a Transfer of its own', () => {
  const h = load();
  const r = parse(h, EMAILS.bdiOnline);
  assert.equal(r.status, 'ok', 'reported: Unrecognized, "Format not known yet"');
  assert.deepEqual(r.items.map(row), [['BDI', 'Card Payment', 'DOP', 2500, 'Pago Tarjetas de Crédito (BDI)', 'Exclude'],
    ['BDI', 'Transfer', 'DOP', 5, 'BDI: impuesto y comisión de pago de tarjeta', '']]);
  assert.equal(r.items[0].txRef, 'BDI:PAY:4000001');
  assert.ok(!r.items.some(x => /PEREZ|1234/.test(x.merchant)), 'neither the holder nor the card number is kept');
  // a commission is added to the tax
  const fee = parseBody(h, EMAILS.bdiOnline, fixture(EMAILS.bdiOnline.body).replace('1.00\nRD$0.00', '1.00\nRD$25.00').replace('RD$2,505.00', 'RD$2,530.00'));
  assert.deepEqual(fee.items.map(x => x.amount), [2500, 30]);
  // without the second table: the total less the tax
  const short = parseBody(h, EMAILS.bdiOnline, fixture(EMAILS.bdiOnline.body).replace(/No\. Referencia[\s\S]*?RD\$0\.00\n/, ''));
  assert.deepEqual(short.items.map(x => x.amount), [2500, 5]);
  // the first layout still reads as before
  assert.deepEqual(h.plain(h.ctx.extractBDICardPaymentTransactions(fixture('bdi_pago_tarjeta'))).map(x => x.amount), [4000, 8]);
});

test('Banreservas loan payment: money out named after the loan, no category; its tax a row of its own', () => {
  const h = load();
  const r = parse(h, EMAILS.brLoan);
  assert.equal(r.status, 'ok', 'reported: Unrecognized, "Format not known yet"');
  assert.deepEqual(r.items.map(row), [['BANRESERVAS', 'Transfer', 'DOP', 3200.5, 'Pago de préstamo *3333', ''],
    ['BANRESERVAS', 'Transfer', 'DOP', 4.8, 'Banreservas: impuesto y comisión de pago de préstamo', '']]);
  assert.equal(r.items[0].txRef, 'BANRESERVAS:200000000001');
  // a transfer to a third party still reads as before
  assert.deepEqual(parse(h, { subject: 'Recibo de la transacción', from: EMAILS.brLoan.from, body: 'banreservas_transferencia_tercero' }).items.map(x => x.type),
    ['Transfer', 'Transfer']);
});

test('emails with no movement of money are filtered, not sent to Unrecognized', () => {
  const h = load();
  const cases = [
    [{ subject: 'Cancelacion tarjeta de débito', from: 'notificaciones@popularenlinea.com' },
      '*Estimado ANA PEREZ* Le informamos que la cancelación de su tarjeta ha sido completado exitosamente. *Producto*: Tarjeta Débito Popular *Terminal*: 1111 *Fecha*: 8/10/2026'],
    [{ subject: '¡Hemos recibido tu mensaje! Estamos aquí para ayudarte', from: 'bdinforma@bdi.com.do' },
      'Solicitud en Línea Solicitud de Contáctanos Ana, Gracias por ponerte en contacto con nosotros. Hemos recibido tu solicitud. Monto disponible RD$1,000.00'],
    [{ subject: 'Solicitud Recibida', from: SCOTIA },
      'Hola ANA, Gracias por realizar su solicitud. Le confirmamos que ha sido recibida por el departamento correspondiente bajo el número INC-00000001.'],
    [{ subject: 'Solicitud completada', from: SCOTIA },
      'Hola ANA, Nos complace en informarle que su número de solicitud INC-00000001 ha sido completado.']];
  for (const [e, body] of cases) assert.equal(parseBody(h, e, body).status, 'filtered', e.subject);
});

test('BHD "Denegada- Confirme Banco": a declined attempt, filtered — not "Amount not found", never saved', () => {
  const h = load();
  const e = { subject: 'BHD Notificación de Transacciones', from: 'BHD <Alertas@bhd.com.do>' };
  const r = parseBody(h, e, fixture('bhd_consumo_denegada'));
  assert.equal(r.status, 'filtered', 'reported: three of these in Unrecognized as "Amount not found"');
  assert.deepEqual(r.items, []);
  const rows = h.plain(h.ctx.extractBHDConsumoTransactions(fixture('bhd_consumo_denegada')));
  assert.deepEqual(rows.map(x => [x.amount, x.currency, x.merchant, x.declined]), [[7.5, 'USD', 'TIENDA EJEMPLO ONLINE', true]],
    'the status (and the bank\'s note after it) is not taken for the merchant');
  // the same row approved is still a purchase
  const ok = parseBody(h, e, fixture('bhd_consumo_denegada').replace('Denegada- Confirme Banco', 'Aprobada'));
  assert.deepEqual(ok.items.map(x => [x.amount, x.currency, x.merchant]), [[7.5, 'USD', 'TIENDA EJEMPLO ONLINE']]);
});

test('Banreservas cash deposit at a branch ("Dep de ahorros sin libreta"): money in; yours when you deposited it', () => {
  const h = load();
  const e = { subject: 'Notificación - Dep de ahorros sin libreta', from: '"notificaciones@banreservas.com" <notificaciones@banreservas.com>' };
  const r = parseBody(h, e, fixture('banreservas_deposito_sin_libreta'));
  assert.equal(r.status, 'ok', 'reported: Unrecognized, "Format not known yet"');
  assert.deepEqual(r.items.map(row), [['BANRESERVAS', 'Incoming', 'USD', -500, 'Depósito en oficina: ANA PEREZ', 'Exclude']]);
  assert.match(r.items[0].description, /\(own account\)$/);
  assert.equal(r.items[0].txRef, 'BANRESERVAS:IN:260914000000000001');
  // someone else deposited it: money received, the category is yours to set
  const other = parseBody(h, e, fixture('banreservas_deposito_sin_libreta').replace('ANA PEREZ - ID', 'LUIS GOMEZ - ID'));
  assert.deepEqual(other.items.map(row), [['BANRESERVAS', 'Incoming', 'USD', -500, 'Depósito en oficina: LUIS GOMEZ', '']]);
  // a transfer received still reads as before
  assert.deepEqual(parse(h, { subject: 'Notificaciones Banreservas', from: e.from, body: 'banreservas_transferencia_recibida' }).items.map(x => [x.type, x.amount]),
    [['Incoming', -2500]]);
});
