'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');

// v1.1.68: Banesco's September 2026 statement layout writes every amount with its currency ("RD$100.00 RD$253,675.36"),
// and the holder's name sits after the account number. The real statement went to Unrecognized as "no transactions
// found; credits add up to 0.00". The fixture keeps the real structure (as read by pdfToText) with invented data.
const NEW = () => fixture('statements/banesco_savings_rd_amounts');

test('Banesco statement with RD$ before every amount: every row read, the totals add up, every credit taken', () => {
  const h = load();
  const r = h.plain(h.ctx.parseBanescoSavingsStatement(NEW()));
  assert.deepEqual(r.problems, [], 'reported: no transactions found; credits add up to 0.00');
  assert.equal(r.rows.length, 9);
  assert.equal(r.currency, 'DOP');
  assert.equal(r.holder, 'ANA MARIA PEÑA SOTO', 'the holder after "Detalle de tus transacciones", before "Moneda:"');
  assert.deepEqual(r.incoming.map(x => [x.day, x.amount, x.merchant, x.own]), [
    ['2026-09-05', 5000, 'Ana Maria Pe A Soto', true],                       // the holder decides own money
    ['2026-09-15', 2500, 'ACH transfer (sender not in the statement)', false],
    ['2026-09-30', 120.5, 'Intereses', false]]);
  // still all-or-nothing: one amount off and nothing is taken
  const bad = h.plain(h.ctx.parseBanescoSavingsStatement(NEW().replace('RD$2,500.00', 'RD$2,600.00')));
  assert.ok(bad.problems.length > 0);
  assert.deepEqual(bad.incoming, []);
});

test('the holder is also found after the account number when the transactions header does not name it', () => {
  const h = load();
  const text = NEW().replace(/Detalle de tus transacciones\nANA MARIA PEÑA SOTO\n/, 'Detalle de tus transacciones\n');
  const r = h.plain(h.ctx.parseBanescoSavingsStatement(text));
  assert.deepEqual(r.problems, []);
  assert.equal(r.holder, 'ANA MARIA PEÑA SOTO');
  assert.equal(r.incoming[0].own, true);
  // and the other way round: only the transactions header names it
  const onlyDetail = h.plain(h.ctx.parseBanescoSavingsStatement(NEW().replace(/N° de cuenta: 00000000000\nANA MARIA PEÑA SOTO\n/, 'N° de cuenta: 00000000000\n')));
  assert.equal(onlyDetail.holder, 'ANA MARIA PEÑA SOTO');
});

test('the currency comes from "Moneda:", not from the RD$ of the rate table', () => {
  const h = load();
  const usd = NEW().replace('Moneda: DOP', 'Moneda: USD').replace(/RD\$(?=[\d,]+\.\d{2} RD\$|[\d,]+\.\d{2}\n)/g, 'US$');
  const r = h.plain(h.ctx.parseBanescoSavingsStatement(usd));
  assert.deepEqual(r.problems, []);
  assert.equal(r.currency, 'USD', 'a USD account whose statement still lists the RD$ rate ranges');
  assert.ok(r.incoming.every(x => x.currency === 'USD'));
  // the old layout, without "Moneda:", still decides by the signs
  assert.equal(h.ctx.parseBanescoSavingsStatement(fixture('statements/banesco_savings_flat').replace('Moneda: DOP', '')).currency, 'DOP');
});

test('the statement email with the new layout is saved, not sent to Unrecognized', () => {
  const h = load();
  h.ctx.statementPdfText = () => NEW();                                    // the PDF reader is tested in v1.1.51/52
  const message = h.fakeMessage({ subject: 'Estado de Cuenta de Ahorros Banesco', from: 'Banesco <estadodecuenta@banesco.com.do>',
    body: 'Adjunto su estado de cuenta.', date: h.date(2026, 10, 9, 8), id: 'stmt-sep' });
  message.getAttachments = () => [{ getContentType: () => 'application/pdf', getName: () => 'Septiembre_2026.pdf', copyBlob: () => ({}) }];
  const r = h.plain(h.ctx.parseStatementEmail(message, 'BANESCO', h.get('BANK_PATTERNS').BANESCO.statement, {}, h.ctx.newParseStats()));
  assert.equal(r.status, 'ok', r.reason);
  assert.deepEqual(r.items.map(x => [x.amount, x.category, x.type]), [[-5000, 'Exclude', 'Incoming'], [-2500, '', 'Incoming'], [-120.5, '', 'Incoming']]);
});
