'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture, manifest } = require('./harness');

test('type detection: confirmed subjects (C3)', () => {
  const h = load();
  const t = (subject, body) => h.ctx.detectTransactionType(subject, body || '');
  assert.equal(t('BHD Notificación de Transacciones', fixture('bhd_consumo_cacharepa')), 'Transaction',
    "the real CACHAREPA email used to be typed Transfer because of 'ACH'");
  assert.equal(t('Transacciones entre mis productos'), 'Transfer');
  assert.equal(t('Juan, ¡Transferencia exitosa!'), 'Transfer');
  assert.equal(t('Notificación de Transferencia Realizada'), 'Transfer');
  assert.equal(t('Joaquín, ¡Realizaste un pago a tu tarjeta LAFISE!'), 'Card Payment');
  assert.equal(t('Servicio de Alerta - Nuevo Consumo'), 'Transaction');
  assert.equal(t('Alerta de Consumo Banesco RD'), 'Transaction');
  assert.equal(t('Notificación de Consumo'), 'Transaction');
  // Unknown subject (e.g. BDI): body keywords still apply, whole-word
  assert.equal(t('Aviso BDI', 'Se realizó una TRANSFERENCIA ACH por RD$ 500'), 'Transfer');
  assert.equal(t('Aviso BDI', 'Consumo en NACHOS BAR por RD$ 500'), 'Transaction');
});

test('subject-only detection returns null for unknown subjects (used by recategorize)', () => {
  const h = load();
  assert.equal(h.ctx.detectTypeFromSubject('Aviso BDI'), null);
  assert.equal(h.ctx.detectTypeFromSubject('BHD Notificación de Transacciones'), 'Transaction');
});

for (const fx of manifest()) {
  test(`end-to-end parse: ${fx.name}${fx.synthetic ? ' (synthetic)' : ' (real, anonymized)'}`, () => {
    const h = load();
    const [y, mo, d, hh, mi] = fx.date;
    const msg = h.fakeMessage({ subject: fx.subject, from: fx.from, body: fixture(fx.name), date: h.date(y, mo, d, hh, mi), id: 'id-' + fx.name });
    const stats = h.ctx.newParseStats();
    const res = h.plain(h.ctx.parseEmailMessage(msg, {}, stats));
    assert.equal(res.items.length, fx.expected.length, 'item count — log:\n' + h.logs.join('\n'));
    fx.expected.forEach((exp, i) => {
      const got = res.items[i];
      assert.equal(got.merchant, exp.merchant);
      assert.equal(got.amount, exp.amount);
      assert.equal(got.currency, exp.currency);
      assert.equal(got.type, exp.type);
      if (exp.category !== undefined) assert.equal(got.category, exp.category);
      assert.equal(got.messageId, 'id-' + fx.name + '_' + i);
      assert.equal(got.description, got.merchant);
    });
  });
}

test('declined BANESCO email: skipped, and NO fallback rows (M1)', () => {
  const h = load();
  const msg = h.fakeMessage({ subject: 'Alerta de Consumo Banesco RD', from: 'notificaciones@banesco.com.do',
    body: fixture('banesco_consumo_declined'), date: h.date(2026, 2, 16) });
  const stats = h.ctx.newParseStats();
  const res = h.plain(h.ctx.parseEmailMessage(msg, {}, stats));
  assert.equal(res.items.length, 0, 'used to save the declined charge AND the "Dispones de" balance');
  assert.equal(res.status, 'filtered');
  assert.equal(stats.declined, 1);
});

test('multi-row BHD email with one declined row keeps the approved row (M1)', () => {
  const h = load();
  assert.equal(h.ctx.isDeclinedTransactionEmail('BHD Notificación de Transacciones', fixture('bhd_consumo_multirow_mixed')), true);
  assert.equal(h.ctx.hasApprovedRow(fixture('bhd_consumo_multirow_mixed')), true);
  assert.equal(h.ctx.hasApprovedRow('La transacción no fue aprobada'), false);
});

test('balance amounts are never transactions', () => {
  const h = load();
  const body = fixture('banesco_consumo_hola_plaza');
  const idx = body.indexOf('RD$ 25,410.08');
  assert.ok(idx > 0);
  assert.equal(h.ctx.isBalanceAmount(body, idx), true);
  const amounts = h.plain(h.ctx.extractAllAmounts(body, h.get('BANK_PATTERNS').BANESCO)).map(i => i.amount);
  assert.deepEqual(amounts, [387.45]);
});

test('IsCredit means money in, scoped to the item (M3)', () => {
  const h = load();
  assert.equal(h.ctx.computeIsCredit('Card Payment', ''), true);
  assert.equal(h.ctx.computeIsCredit('Cashback', ''), true);
  assert.equal(h.ctx.computeIsCredit('Transaction', 'consumo con tu TARJETA DE CRÉDITO VISA'), false);
  assert.equal(h.ctx.computeIsCredit('Transaction', 'CRÉDITO a tu cuenta por RD$ 500'), true);
});

test('looksGarbled: real merchants pass, broken captures are caught', () => {
  const h = load();
  const g = t => h.ctx.looksGarbled(t);
  assert.equal(g('HOLA PLAZA LAS AMERICAS'), false, 'v1.1.18 false positive');
  assert.equal(g('PedidosYa*Som Cafe'), false);
  assert.equal(g('cia a banco local] Hola Juan, Acabas de realiza'), true);
  assert.equal(g('Producto destino: XXXX1111'), true);
});

test('unknown sender, promotional and non-transactional emails are counted', () => {
  const h = load();
  const stats = h.ctx.newParseStats();
  const run = (subject, from, body) =>
    h.ctx.parseEmailMessage(h.fakeMessage({ subject, from, body, date: h.date(2026, 9, 1) }), {}, stats);
  run('Recibo de tu viaje', 'Uber <noreply@uber.com>', 'Pagado con LAFISE RD$ 200');
  run('¡Participa y gana!', 'notificaciones@banesco.com.do', 'PROMOCIÓN especial RD$ 100');
  run('Tu código OTP', 'Alertas@bhd.com.do', 'Tu OTP es 123456');
  assert.equal(stats.notOwnBank, 1);
  assert.equal(stats.promotional, 1);
  assert.equal(stats.nonTransactional, 1);
});
