'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

const LAFISE_IN = 'PagosAlInstanteMT103@lafise.com';
const SUBJECT_IN = 'TRANSFERENCIA ENTRANTE APLICADA EXITOSAMENTE.';

test('a LAFISE incoming transfer: Incoming, negative, no category by itself — your own account is Exclude', () => {
  const h = load();
  assert.equal(h.ctx.detectTransactionType(SUBJECT_IN, fixture('lafise_incoming_third_party')), 'Incoming',
    'not a Transfer, although its body says "TRANSFERENCIA ... RECIBIDA" and "PAGOS AL INSTANTE"');
  const msg = name => h.fakeMessage({ subject: SUBJECT_IN, from: LAFISE_IN, body: fixture(name), date: h.date(2026, 9, 12, 9), id: name });
  const other = h.plain(h.ctx.parseEmailMessage(msg('lafise_incoming_third_party'), {}, h.ctx.newParseStats())).items[0];
  assert.deepEqual([other.type, other.amount, other.currency, other.merchant, other.category, other.isCredit],
    ['Incoming', -4500, 'DOP', 'LUIS ALBERTO GOMEZ DIAZ', '', true]);
  const own = h.plain(h.ctx.parseEmailMessage(msg('lafise_incoming_own'), {}, h.ctx.newParseStats())).items[0];
  assert.deepEqual([own.amount, own.category, own.description], [-25000, 'Exclude', 'ANA MARIA PEREZ SOTO (own account)']);
  // a Custom Rule on the sender's name is the user's way of categorizing it
  const rules = { 'user@example.com': { Rent: ['LUIS ALBERTO'] } };
  assert.ok(h.ctx.sameHolder('MARIA JOSE NUÑEZ DE LEON', 'Maria Jose Nu Ez De'), 'Ñ written as a space, name cut');
  assert.ok(!h.ctx.sameHolder('ANA MARIA PEREZ', 'ANA LUCIA PEREZ'));
});

test('a Banesco savings statement is read the same whatever the text layout; its own totals must add up', () => {
  const h = load();
  const results = ['layout', 'flat', 'cells'].map(v => h.plain(h.ctx.parseBanescoSavingsStatement(fixture('statements/banesco_savings_' + v))));
  for (const r of results) {
    assert.deepEqual(r.problems, []);
    assert.equal(r.rows.length, 10);
    assert.equal(r.holder, 'ANA MARIA PEREZ SOTO');
    assert.deepEqual(r.incoming.map(x => [x.day, x.amount, x.merchant, x.own]), [
      ['2026-08-17', 6500, 'ACH transfer (sender not in the statement)', false], ['2026-08-24', 2300, 'ACH transfer (sender not in the statement)', false],
      ['2026-08-28', 3000, 'Ana Maria Perez Sot', true], ['2026-08-31', 1200, 'Luis Gomez Diaz', false]], 'interest and debits left out');
  }
  // one amount misread → the totals don't add up → nothing is taken
  const bad = h.plain(h.ctx.parseBanescoSavingsStatement(fixture('statements/banesco_savings_layout').replace('6,500.00', '6,600.00')));
  assert.ok(bad.problems.length > 0);
  assert.deepEqual(bad.incoming, []);
  // every row follows the balance, but the statement's credit total says otherwise: still nothing taken
  const total = h.plain(h.ctx.parseBanescoSavingsStatement(fixture('statements/banesco_savings_layout').replace('13,150.25', '13,250.25')));
  assert.deepEqual(total.problems, ['credits add up to 13150.25, the statement says 13250.25']);
  assert.deepEqual(total.incoming, []);
});

function book(opts) {
  opts = opts || {};
  const mock = makeServices({ ui: true });
  const trashed = [];
  if (opts.drive !== false) {
    mock.services.Drive = { Files: { create: (meta, blob) => { assert.equal(meta.mimeType, 'application/vnd.google-apps.document'); return { id: 'doc-1' }; } } };
    mock.services.DocumentApp = { openById: id => ({ getBody: () => ({ getText: () => fixture('statements/banesco_savings_flat') }) }) };
    mock.services.DriveApp = { getFileById: id => ({ setTrashed: v => trashed.push(id) }) };
  }
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ BANESCO: true, LAFISE: true })]].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  const cr = mock.ss.insertSheet('Custom Rules'); cr.appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  (opts.rules || []).forEach(r => cr.appendRow(['user@example.com', r[0], r[1], '']));
  const statement = h.fakeMessage({ subject: 'Estado de Cuenta de Ahorros Banesco', from: 'Banesco <estadodecuenta@banesco.com.do>',
    body: 'Adjunto su estado de cuenta.', date: h.date(2026, 9, 2, 10), id: 'stmt-aug' });
  statement.getAttachments = () => [{ getContentType: () => 'application/pdf', getName: () => 'Agosto_2026.pdf', copyBlob: () => ({ pdf: true }) }];
  mock.gmail.threads.push(fakeThread('t-stmt', [statement]),
    fakeThread('t-in', [h.fakeMessage({ subject: SUBJECT_IN, from: LAFISE_IN, body: fixture('lafise_incoming_third_party'), date: h.date(2026, 9, 12, 9), id: 'in-1' })]));
  return { h, mock, trashed, tx: () => h.plain(mock.ss.getSheetByName('Transactions')._rows(15)) };
}

test('a run reads the statement PDF (via Drive) and the LAFISE email; Incoming Transfers lists them; nothing twice', () => {
  const { h, mock, trashed, tx } = book({ rules: [['Rent', 'LUIS ALBERTO']] });
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const incoming = tx().filter(r => r[11] === 'Incoming');
  assert.deepEqual(incoming.map(r => [r[1], r[2], r[3], r[5]]).sort(), [
    ['BANESCO', 'ACH transfer (sender not in the statement)', -2300, ''], ['BANESCO', 'ACH transfer (sender not in the statement)', -6500, ''],
    ['BANESCO', 'Ana Maria Perez Sot', -3000, 'Exclude'], ['BANESCO', 'Luis Gomez Diaz', -1200, ''],
    ['LAFISE', 'LUIS ALBERTO GOMEZ DIAZ', -4500, 'Rent']].sort(), 'the Custom Rule categorized the LAFISE one');
  assert.equal(h.ctx.normalizeDateForCompare(incoming.find(r => r[3] === -6500)[0]), '2026-08-17', 'the statement row\'s own date');
  assert.deepEqual(trashed, ['doc-1'], 'the converted copy is trashed');
  const sheet = mock.ss.getSheetByName('Incoming Transfers');
  assert.ok(sheet, 'Incoming Transfers created');
  assert.deepEqual(h.plain(sheet.getRange(1, 1, 1, 8).getValues()[0]).slice(0, 4), ['Date', 'Bank', 'From', 'Category']);
  assert.ok(sheet._rows(8).every(r => r[4] > 0), 'shown as received (positive)');
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.equal(tx().filter(r => r[11] === 'Incoming').length, 5, 'not saved twice');
});

test('money received with a category is taken off that category; without one it counts nowhere', () => {
  const h = load();
  const row = (d, type, merchant, amount, category) => { const x = new Array(15).fill(''); x[0] = h.date(2026, 8, d); x[1] = 'BANESCO';
    x[2] = merchant; x[3] = amount; x[4] = 'DOP'; x[5] = category; x[6] = merchant; x[11] = type; return x; };
  const values = [new Array(15).fill('h'), row(26, 'Transfer', 'PAGO RENTA', 30000, 'Rent'), row(28, 'Incoming', 'ROOMMATE', -15000, 'Rent'),
    row(29, 'Incoming', 'ACH', -9000, ''), row(30, 'Incoming', 'ME', -5000, 'Exclude')];
  const m = h.plain(h.ctx.computeMonthlySummary(values, { today: h.date(2026, 9, 1, 9), rates: { USD: 60, EUR: 64, COP: 0.015 }, netIncomeDop: 80000, cards: [] }));
  const rent = m.categories.find(c => c.cat === 'Rent');
  assert.equal(rent.amount, 15000, '30,000 paid − 15,000 paid back');
  assert.equal(m.spent, 15000, 'the uncategorized and the excluded ones count nowhere');
});

test('recategorizing keeps a row Incoming, your own money Exclude, and a category typed in Incoming Transfers', () => {
  const { h, mock, tx } = book();
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const sheet = mock.ss.getSheetByName('Incoming Transfers');
  const at = sheet._rows(8).findIndex(r => r[4] === 6500) + 2;
  sheet.getRange(at, 4).setValue('Electricity');                      // typed where you review them
  h.ctx.recategorizeAllTransactions('user@example.com');
  h.ctx.recategorizeAllTransactions('user@example.com');
  const incoming = tx().filter(r => r[3] < 0);
  assert.ok(incoming.every(r => r[11] === 'Incoming'), 'still Incoming');
  assert.equal(incoming.find(r => r[3] === -6500)[5], 'Electricity', 'your category kept');
  assert.equal(incoming.find(r => r[3] === -3000)[5], 'Exclude', 'own money stays Exclude');
});

test('without the Drive API the statement goes to Unrecognized with how to turn it on; Start here says so too', () => {
  const { h, mock, tx } = book({ drive: false });
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.equal(tx().filter(r => r[1] === 'BANESCO').length, 0);
  const row = mock.ss.getSheetByName('Unrecognized')._rows(10).find(r => r[9] === 'stmt-aug');
  assert.match(row[3], /Statement not read: .*Turn on the Drive API service.*Services › \+ › Drive API/);
  const steps = Object.fromEntries(h.plain(h.ctx.startHereStatus(h.date(2026, 9, 28, 9))).steps.map(s => [s.id, s]));
  assert.equal(steps.statements.status, 'optional');
  assert.match(steps.statements.detail, /Drive API/);
  assert.equal(steps.incoming.status, 'warn', 'the LAFISE one still needs a category');
});
