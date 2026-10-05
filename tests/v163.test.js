'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');

const parse = (h, subject, from, body, stats) => h.plain(h.ctx.parseEmailMessage(h.fakeMessage({ subject, from, body,
  date: h.date(2026, 10, 2, 10), id: 'm' + Math.random() }), {}, stats || h.ctx.newParseStats()));
const LAFISE_CARD = 'LAFISE CF <notificaciones@bancolafise.com>';
const LAFISE_BANK = 'bancanet@notificaciones.lafise.com';

test('an authorization code starting with "US" doesn\'t make a DOP purchase dollars (reported: Uber Eats DOP 390 saved as USD)', () => {
  const h = load();
  const r = parse(h, 'Servicio de Alerta - Nuevo Consumo', LAFISE_CARD, fixture('lafise_consumo_ubereats').replace('Autorización:\n000000', 'Autorización:\nUS34K7'));
  assert.deepEqual(r.items.map(t => [t.currency, t.amount]), [['DOP', 297.25]]);
  const cur = s => h.ctx.detectCurrencyFromMatch(s);
  assert.deepEqual([cur('Autorización: US34K7 Monto: 390'), cur('US$ 25.00'), cur('USD 25.00'), cur('RD$ 1,000.00')], ['DOP', 'USD', 'USD', 'DOP']);
});

test('LAFISE "Aviso de transferencia": a failed one ("Estado: Error") is not a transfer — it is filtered as failed', () => {
  const h = load();
  const st = h.ctx.newParseStats();
  const r = parse(h, 'Aviso de transferencia en banco local', LAFISE_BANK, fixture('lafise_aviso_transferencia_local'), st);
  assert.equal(r.status, 'filtered', 'reported: two failed rent payments came as Unrecognized and would have counted');
  assert.equal(st.declined, 1);
  assert.equal(st.unrecognized.length, 0);
});

test('LAFISE "Aviso de transferencia" that went through: a transfer to the beneficiary; to your own name, Exclude', () => {
  const h = load();
  // the success wording isn't known from a sample yet: any status that isn't a failure is read
  const ok = fixture('lafise_aviso_transferencia_local').replace('Estado: Error', 'Estado: Exitosa');
  const r = parse(h, 'Aviso de transferencia en banco local', LAFISE_BANK, ok);
  assert.deepEqual(r.items.map(t => [t.bank, t.type, t.currency, t.amount, t.merchant, t.category || '', t.txRef]),
    [['LAFISE', 'Transfer', 'DOP', 12500, 'PEDRO ANTONIO RAMIREZ', '', 'LAFISE:OUT:100000001']]);
  const own = parse(h, 'Aviso de transferencia en banco local', LAFISE_BANK, ok.replace('Titular: PEDRO ANTONIO RAMIREZ', 'Titular: MARIA ELENA GOMEZ'));
  assert.deepEqual(own.items.map(t => [t.category, t.description]), [['Exclude', 'MARIA ELENA GOMEZ (own account)']], 'money moved to your own account');
});

test('Unrecognized noise: replies and forwards, LAFISE account notices and BDI registration emails are skipped on purpose', () => {
  const h = load();
  const cases = [
    // a reply with nothing else that would filter it — only the subject says it's a conversation
    ['RE: Consulta sobre mi cuenta', 'Oficial <oficial@bdi.com.do>', 'Buenos días, quedamos atentos a su respuesta. Saludos.'],
    ['RE: Firma de Contrato - Apertura Cuenta Digital', 'Oficial <oficial@bdi.com.do>', 'Buenos días, le confirmamos su firma ha sido recibida.'],
    ['RV: Solicitud Código de Activación', 'Oficial <oficial@bdi.com.do>', 'Su código es 12345678.'],
    ['Aquí tiene su contraseña temporal!', LAFISE_BANK, 'Esta es su contraseña temporal de Bancanet abc123'],
    ['Advertencia de sesión duplicada', LAFISE_BANK, 'Advertencia de sesión duplicada Información sobre la sesión activa'],
    ['Tu usuario de Bancanet está por bloquearse', LAFISE_BANK, 'Ingresa a Bancanet antes de que tu usuario se bloquee'],
    ['El Alias del usuario en Bancanet se ha actualizado', LAFISE_BANK, 'Este es su nuevo alias/usuario para Bancanet'],
    ['Desbloqueo de usuario', LAFISE_BANK, 'Desbloqueo de usuario. Hemos reactivado nuevamente tu cuenta'],
    ['BDI Digital: Solicitud de Registro BDI en Línea', 'BDI Digital <BDIDigital@bdi.com.do>', 'Solicitud de Registro en BDI Digital. Activar usuario.']
  ];
  for (const [subject, from, body] of cases) {
    const st = h.ctx.newParseStats();
    assert.equal(parse(h, subject, from, body, st).status, 'filtered', subject);
    assert.equal(st.unrecognized.length, 0, subject);
  }
});

test('a real purchase whose footer mentions Bancanet is still read', () => {
  const h = load();
  const r = parse(h, 'Servicio de Alerta - Nuevo Consumo', LAFISE_CARD,
    fixture('lafise_consumo_ubereats') + '\nRecuerda tu usuario de Bancanet. Si olvidaste tu contraseña temporal, recupérala en Bancanet.\n');
  assert.equal(r.status, 'ok');
  assert.deepEqual(r.items.map(t => [t.currency, t.amount]), [['DOP', 297.25]]);
});
