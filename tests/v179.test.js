'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

// v1.1.79: the LAFISE credit card statement (PDF): two currency sections, each checked against its own total line and
// the summary at the top; purchases already saved from their alerts are not saved again (the QIK pairing, v1.1.72).
// The fixture keeps the real statement's layout (as read by pdfToText) with invented data.

const STMT = () => fixture('statements/lafise_card');
const FROM = 'info@lafise.com.do';
const SUBJECT = 'Estado de Cuenta de Tarjeta BANCO LAFISE - 0000 - Fecha de corte 2026-10-20';

test('the statement is read: cut date, both currencies, a page that goes on with the pesos section, each kind of credit', () => {
  const h = load();
  const r = h.plain(h.ctx.parseLAFISECardStatement(STMT()));
  assert.deepEqual(r.problems, []);
  assert.equal(r.cut, '2026-10-20');
  assert.deepEqual(r.entries.map(e => [e.day, e.type, e.currency, e.amount, e.merchant]), [
    ['2026-09-15', 'Card Payment', 'USD', 20, 'ABONO A SU CUENTA'],
    ['2026-09-21', 'Transaction', 'USD', 40, 'Amazon.com*1A2B3C4D5 SEATTLE'],
    ['2026-10-04', 'Transaction', 'USD', 4.99, 'SPOTIFY*P0000000AA SPOTIFY.COM'],
    ['2026-10-11', 'Transaction', 'USD', -1.5, 'Purchase Returns - AMAZON PRIME PMTS SEATTLE'],
    ['2026-10-12', 'Transaction', 'USD', 7.99, 'NETFLIX.COM LOS GATOS'],
    ['2026-10-14', 'Cashback', 'USD', 1, 'LAFISE cashback (statement)'],
    ['2026-09-12', 'Transaction', 'DOP', 500, 'RESTAURANTE EJEMPLO SANTO DOMINGO'],
    ['2026-09-12', 'Transaction', 'DOP', 650, 'UBER*EATS SANTO DOMINGO'],
    ['2026-09-15', 'Transaction', 'DOP', 200, 'UBER RIDES-*UBER RIDES SANTO DOMINGO'],
    // the bank's own word for a cashback granted by hand (its "Crédito" comes garbled in the PDF itself)
    ['2026-09-15', 'Cashback', 'DOP', 750, 'LAFISE cashback (statement)'],
    ['2026-09-17', 'Transaction', 'DOP', 2400, 'CLARO PAGO 8090000000 SANTO DOMINGO'],
    ['2026-09-20', 'Transaction', 'DOP', 150, 'UBER*RIDES SANTO DOMINGO'],
    ['2026-09-20', 'Transaction', 'DOP', 150, 'UBER*RIDES SANTO DOMINGO'],
    ['2026-09-30', 'Transaction', 'DOP', 600, 'GIMNASIO EJEMPLO DISTRITO NACI'],
    ['2026-10-03', 'Transaction', 'DOP', 2000, 'ESTACION EJEMPLO SANTO DOMINGO'],
    ['2026-10-14', 'Cashback', 'DOP', 120, 'LAFISE cashback (statement)']]);
  const refs = r.entries.map(e => e.ref);
  assert.equal(new Set(refs).size, refs.length, 'two equal rides on one day are two rows');
  assert.ok(refs.includes('20261020:2026-09-20:P:15000:1') && refs.includes('20261020:2026-09-20:P:15000:2'));
  assert.ok(refs.includes('20261020:2026-09-21:PUSD:4000:1'), 'a dollar row says so in its reference');
});

test('all or nothing: a row misread, a total or summary that differs, a row that is both or neither → nothing taken, it says why', () => {
  const h = load();
  const bad = [
    [STMT().replace('RESTAURANTE EJEMPLO SANTO DOMINGODODOM 500.00', 'RESTAURANTE EJEMPLO SANTO DOMINGODODOM 50.00'), /DOP rows add up to 6200\.00/],
    [STMT().replace('**Total transacciones en pesos RD$** 6,650.00', '**Total transacciones en pesos RD$** 6,560.00'), /DOP rows .* the section total says 6560\.00/],
    [STMT().replace(/^US\$ 200\.00 20\.00 22\.50/m, 'US$ 200.00 20.00 21.50'), /USD rows .* the summary says 52\.98 and 21\.50/],
    [STMT().replace('NETFLIX.COM LOS GATOS CAUSA 7.99 0.00', 'NETFLIX.COM LOS GATOS CAUSA 7.99 7.99'), /not one debit or one credit/],
    [STMT().replace(/\*\*Total transacciones en Dólares[^\n]+\n/, ''), /USD: the section total was not found/],
    [STMT().replace('Transacciones de Dólares US$\n', ''), /before any currency section/],
    [STMT().replace(/Fecha de corte: [^\n]+/g, ''), /cut date/]];
  bad.forEach(([text, why], i) => {
    const r = h.ctx.parseLAFISECardStatement(text);
    assert.match(r.problems.join('; '), why, 'case ' + i);
    assert.equal(r.entries.length, 0, 'case ' + i);
  });
});

test('the balance chain is not required: the real one did not close in pesos while every row matched both totals', () => {
  const h = load();
  const r = h.ctx.parseLAFISECardStatement(STMT().replace(/^RD\$ 30,000\.00 0\.00 870\.00 6,650\.00 5,780\.00$/gm, 'RD$ 30,000.00 0.00 870.00 6,650.00 1,000.00'));
  assert.deepEqual(h.plain(r.problems), []);
  assert.equal(r.entries.length, 16);
});

test('a card without dollar movements: no dollar section, nothing missing', () => {
  const h = load();
  const pesosOnly = STMT().replace(/Transacciones de Dólares US\$\n[\s\S]*?\*\*Total transacciones en Dólares[^\n]+\n/, '')
    .replace(/^US\$ 200\.00 20\.00 22\.50 52\.98 50\.48$/gm, 'US$ 200.00 0.00 0.00 0.00 0.00');
  const r = h.plain(h.ctx.parseLAFISECardStatement(pesosOnly));
  assert.deepEqual(r.problems, []);
  assert.ok(r.entries.length === 10 && r.entries.every(e => e.currency === 'DOP'));
});

test('the merchant without the region and country, in capitals too (LAFISE) and for Sweden', () => {
  const h = load();
  [['UBER*EATS SANTO DOMINGODODOM', 'UBER*EATS SANTO DOMINGO'], ['GIMNASIO EJEMPLO DISTRITO NACISDDOM', 'GIMNASIO EJEMPLO DISTRITO NACI'],
   ['Amazon.com*1A2B SEATTLE WAUSA', 'Amazon.com*1A2B SEATTLE'], ['SPOTIFY*P0 SPOTIFY.COM ABSWE', 'SPOTIFY*P0 SPOTIFY.COM'],
   ['Tienda Ejemplo Santo Domingododom', 'Tienda Ejemplo Santo Domingo']]   // QIK, as before
    .forEach(([raw, clean]) => assert.equal(h.ctx.cleanCardStatementMerchant(raw), clean, raw));
});

// ---- the statement email
const statementMessage = (h, id) => ({ getId: () => id || 'm-stmt', getFrom: () => FROM, getSubject: () => SUBJECT,
  getDate: () => h.date(2026, 10, 16, 11), getPlainBody: () => '', getBody: () => '<p>Estado de cuenta</p>',
  getAttachments: () => [{ getContentType: () => 'application/pdf', getName: () => '0000_2026-10-20.pdf', copyBlob: () => ({}) }],
  getThread: () => null });

test('the statement email: LAFISE by its sender; payment Exclude, cashback a credit, purchases categorized; searched alone', () => {
  const h = load();
  h.ctx.statementPdfText = () => STMT();   // the real PDF was read by pdfToText as this layout
  const r = h.plain(h.ctx.parseEmailMessage(statementMessage(h), {}, h.ctx.newParseStats()));
  assert.equal(r.status, 'ok');
  assert.ok(r.items.every(i => i.bank === 'LAFISE' && /^LAFISE:STMT:20261020:/.test(i.txRef)));
  assert.deepEqual(r.items.filter(i => i.type === 'Card Payment').map(i => [i.currency, i.amount, i.category]), [['USD', 20, 'Exclude']]);
  assert.deepEqual(r.items.filter(i => i.type === 'Cashback').map(i => [i.currency, i.amount, i.isCredit]),
    [['USD', 1, true], ['DOP', 750, true], ['DOP', 120, true]]);
  assert.equal(r.items.find(i => /UBER\*EATS/.test(i.merchant)).category, 'Delivery');
  assert.equal(r.items.find(i => /NETFLIX/.test(i.merchant)).category, 'Subscriptions');
  assert.ok(h.ctx.isStatementEmail(statementMessage(h)), 'read again on every run, like the other statements');
  // the search asks info@ only for its statements: it also sends advertising, and LAFISE is not a strict bank
  const q = h.get('BANK_PATTERNS').LAFISE.searchQuery;
  assert.match(q, / OR \(from:info@lafise\.com\.do subject:"Estado de Cuenta de Tarjeta"\)$/);
  assert.ok(!/from:info@lafise\.com\.do(?! subject)/.test(q));
});

// ---- already saved from the alerts
const alert = (h, m, d, amount, extra) => Object.assign({ date: h.date(2026, m, d), bank: 'LAFISE', merchant: 'UBER EATS-W*UBER EATS- SANTO DOMINGO DOM',
  amount: amount, currency: 'DOP', type: 'Transaction', category: '', messageId: 'alert-' + m + '-' + d + '-' + amount + '_0', txRef: '' }, extra || {});
const savedValues = (h, items) => [new Array(15).fill('h')].concat(items.map(t => h.ctx.transactionToRow(Object.assign({ description: t.merchant, subject: 's', timestamp: 't' }, t))));
const stmtItems = h => {
  h.ctx.statementPdfText = () => STMT();
  return h.ctx.parseEmailMessage(statementMessage(h), {}, h.ctx.newParseStats()).items;
};

test('purchases with an alert (pesos and dollars) and the payment notified are not saved twice; the rest is', () => {
  const h = load();
  const notified = [alert(h, 9, 12, 650), alert(h, 9, 20, 150, { merchant: 'UBER *TRIP' }),
    alert(h, 9, 22, 40, { currency: 'USD', merchant: 'AMAZON.COM' }),                       // two days after the statement's date
    alert(h, 9, 15, 20, { currency: 'USD', type: 'Card Payment', category: 'Exclude', merchant: 'Pago a tarjeta 0000' })];
  const sel = h.ctx.selectNewTransactions(stmtItems(h), h.ctx.buildExistingIndex(savedValues(h, notified)));
  assert.equal(sel.duplicates, 4);
  assert.equal(sel.fresh.length, 12);
  assert.equal(sel.fresh.filter(t => t.amount === 150).length, 1, 'one of the two rides of the 20th had an alert');
  assert.ok(h.logs.some(l => /Only in the statement \(no alert saved\): LAFISE Transaction DOP 2400\.00 2026-09-17/.test(l)));
  const again = h.ctx.selectNewTransactions(stmtItems(h), h.ctx.buildExistingIndex(savedValues(h, notified.concat(sel.fresh))));
  assert.equal(again.fresh.length, 0, 'read again on the next run: nothing more');
});

test('a whole run: a LAFISE alert and the statement in the same search; a second run changes nothing', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 80000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', '{"LAFISE":true}']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  h.ctx.statementPdfText = () => STMT();
  const consumo = { getId: () => 'a1', getFrom: () => 'LAFISE CF <notificaciones@bancolafise.com>', getSubject: () => 'Servicio de Alerta - Nuevo Consumo',
    getDate: () => h.date(2026, 9, 12, 20), getPlainBody: () => fixture('lafise_consumo_ubereats').replace('02/09/2026', '12/09/2026').replace('DOP 297.25', 'DOP 650.00'),
    getBody: () => '', getThread: () => null };
  mock.gmail.threads.push(fakeThread('t-alert', [consumo]));
  mock.gmail.threads.push(fakeThread('t-stmt', [statementMessage(h, 'm-stmt')]));
  const rows = () => h.plain(mock.ss.getSheetByName('Transactions').getDataRange().getValues().slice(1));
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-10-20');
  assert.equal(rows().length, 16, 'the alert, and fifteen rows from the statement');
  assert.equal(rows().filter(r => r[3] === 650).length, 1);
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-10-20');
  assert.equal(rows().length, 16);
});
