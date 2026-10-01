'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

// Real samples shared by another user (BDI, Scotiabank, QIK); the fixtures keep their structure with invented data.
const EMAILS = {
  bdiTransfer: { subject: 'BDI Digital: Comprobante transacción Interbancaria', from: '"bdinforma@bdi.com.do" <bdinforma@bdi.com.do>', body: 'bdi_transferencia_interbancaria' },
  bdiConsumo: { subject: 'Notificacion de Consumos', from: 'BDI Informa <BDIinforma@bdi.com.do>', body: 'bdi_consumo' },
  scotia: { subject: 'Autorización fuera del país', from: 'Alertas Scotiabank <alertas@scotiabank.com>', body: 'scotiabank_consumo' },
  qik: { subject: 'Usaste tu tarjeta de crédito Qik', from: 'notificaciones@qik.do', body: 'qik_consumo' }
};
const parse = (h, e, id) => h.plain(h.ctx.parseEmailMessage(h.fakeMessage({ subject: e.subject, from: e.from, body: fixture(e.body),
  date: h.date(2026, 9, 27, 10), id: id || 'm' }), {}, h.ctx.newParseStats()));
const row = t => [t.bank, t.type, t.currency, t.amount, t.merchant];

test('BDI interbank transfer sent: the amount to the beneficiary, and the tax and commission as their own row', () => {
  const h = load();
  const r = parse(h, EMAILS.bdiTransfer);
  assert.equal(r.status, 'ok');
  assert.deepEqual(r.items.map(row), [['BDI', 'Transfer', 'DOP', 2500, 'Luis Ramirez'], ['BDI', 'Transfer', 'DOP', 5, 'BDI — impuesto y comisión de transferencia']]);
  // one that isn't marked as sent is not guessed
  assert.deepEqual(h.plain(h.ctx.extractBDITransferTransactions(fixture('bdi_transferencia_interbancaria').replace('[Salida]', '[Entrada]'))), []);
});

test('BDI card purchases: approved rows only, from the table', () => {
  const h = load();
  const r = parse(h, EMAILS.bdiConsumo);
  assert.equal(r.status, 'ok');
  assert.deepEqual(r.items.map(row), [['BDI', 'Transaction', 'DOP', 1850, 'SUPERMERCADO EJEMPLO']], 'the RECHAZADA row is left out');
});

test('Scotiabank: a new bank, recognized by its sender; the amount in its own currency', () => {
  const h = load();
  assert.deepEqual(parse(h, EMAILS.scotia).items.map(row), [['SCOTIABANK', 'Transaction', 'USD', 25.5, 'TIENDA ONLINE EJEMPLO']]);
});

test('QIK: the purchase, never the available balance; its security footer does not get it filtered — a real code email still is', () => {
  const h = load();
  const r = parse(h, EMAILS.qik);
  assert.equal(r.status, 'ok', 'reported: "código de seguridad" in the footer filtered the purchase');
  assert.deepEqual(r.items.map(row), [['QIK', 'Transaction', 'DOP', 640, 'CAFETERIA EJEMPLO']], 'not RD$9,999.99');
  const code = h.plain(h.ctx.parseEmailMessage(h.fakeMessage({ subject: 'Tu código de seguridad', from: 'notificaciones@qik.do',
    body: 'Tu código de seguridad es 482913. No lo compartas con nadie.', date: h.date(2026, 9, 27, 10), id: 'c' }), {}, h.ctx.newParseStats()));
  assert.equal(code.status, 'filtered');
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

test('a run with the new banks: every email saved, each bank its Raw_ sheet, in the tab order', () => {
  const { h, mock } = book({ BDI: true, SCOTIABANK: true, QIK: true });
  Object.keys(EMAILS).forEach((k, i) => { const e = EMAILS[k];
    mock.gmail.threads.push(fakeThread('t' + i, [h.fakeMessage({ subject: e.subject, from: e.from, body: fixture(e.body), date: h.date(2026, 9, 20 + i, 10), id: 'm' + i })])); });
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const tx = h.plain(mock.ss.getSheetByName('Transactions')._rows(15));
  assert.deepEqual(tx.map(r => r[1] + ' ' + r[3]).sort(), ['BDI 1850', 'BDI 2500', 'BDI 5', 'QIK 640', 'SCOTIABANK 25.5']);
  for (const s of ['Raw_BDI', 'Raw_SCOTIABANK', 'Raw_QIK']) assert.ok(mock.ss.getSheetByName(s), s);
  const order = h.plain(h.get('CANONICAL_SHEET_ORDER'));
  // v1.1.61: Raw_BANRESERVAS added at the end (BANK_ORDER)
  assert.deepEqual(order.filter(n => n.indexOf('Raw_') === 0), ['Raw_LAFISE', 'Raw_BANESCO', 'Raw_BHD', 'Raw_BDI', 'Raw_POPULAR', 'Raw_SCOTIABANK', 'Raw_QIK', 'Raw_BANRESERVAS']);
});

test('one list of banks: the wizard, View Config and the tab order all come from it; a new bank starts unticked in a saved setup', () => {
  const { h, mock } = book({ LAFISE: true, BHD: true });
  const banks = h.plain(h.get('BANK_ORDER'));
  assert.deepEqual(Object.keys(h.plain(h.get('Object.fromEntries(Object.keys(BANK_PATTERNS).map(k => [k, 1]))'))).sort(), banks.slice().sort(), 'every bank read is listed');
  h.ctx.openSetupWizard();
  const html = mock.ui.dialogs.pop().html;
  for (const b of banks) assert.ok(html.includes('id="bank_' + b.toLowerCase() + '"'), b + ' checkbox');
  assert.match(html, /BANKS\.forEach\(function\(b\) \{ el\('bank_' \+ b\.toLowerCase\(\)\)\.checked = !!EXISTING\.banksToTrack\[b\]; \}\)/);
  h.ctx.viewConfig();
  const shown = mock.ui.alerts.pop();
  assert.match(shown, /SCOTIABANK: ❌/);
  assert.match(shown, /QIK: ❌/);
  assert.match(shown, /LAFISE: ✅/);
});

test('the catalogue: QIK pays 1% on everything; BDI, AMEX Gold and LifeMiles have no base cashback (their note says why)', () => {
  const h = load();
  const cards = h.plain(h.ctx.resolveCards([{ bank: 'QIK', product: 'QIK_MASTERCARD' }, { bank: 'BDI', product: 'BDI_VISA_CLASICA' },
    { bank: 'SCOTIABANK', product: 'SCOTIABANK_AMEX_GOLD' }, { bank: 'BHD', product: 'BHD_LIFEMILES' }]));
  assert.deepEqual(cards.map(c => [c.bank, c.name, c.cashback]), [['QIK', 'Mastercard', 0.01], ['BDI', 'Visa Clásica', 0],
    ['SCOTIABANK', 'American Express Gold', 0], ['BHD', 'Visa LifeMiles', 0]]);
  assert.ok(cards.every(c => c.note.length > 30));
  assert.ok(cards.every(c => !/\d{2,3},?\d{3}/.test(c.note)), 'no credit limits or amounts of a particular holder');
});
