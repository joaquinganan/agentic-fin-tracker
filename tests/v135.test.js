'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

function configured() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ LAFISE: true, BHD: true, POPULAR: true })], ['setupDate', '2026-09-01']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions');
  h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  return { h, mock };
}
const unreadable = h => h.fakeMessage({ subject: 'Alerta de consumo LAFISE', from: 'notificaciones@bancolafise.com',
  body: 'Hola Alex, tu operación fue registrada. Detalle disponible en la app.', date: h.date(2026, 9, 20, 10), id: 'm-broken' });
const readable = h => h.fakeMessage({ subject: 'BHD Notificación de Transacciones', from: 'Alertas@bhd.com.do',
  body: fixture('bhd_consumo_cacharepa'), date: h.date(2026, 9, 21, 10), id: 'm-ok' });

test('a bank email that can\'t be read is listed in Unrecognized — once, with reason, what it says and a Gmail link', () => {
  const { h, mock } = configured();
  mock.gmail.threads.push(fakeThread('t1', [unreadable(h)]), fakeThread('t2', [readable(h)]));
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const sheet = mock.ss.getSheetByName('Unrecognized');
  assert.ok(sheet, 'created when there is something to list');
  const rows = sheet._rows(10);
  assert.equal(rows.length, 1, 'only the email that failed');
  const [date, bank, subject, reason, says, link, status, , , id] = rows[0];
  assert.deepEqual([bank, subject, reason, status, id], ['LAFISE', 'Alerta de consumo LAFISE', 'Amount not found', 'New', 'm-broken']);
  assert.match(says, /tu operación fue registrada/);
  assert.equal(link, '=HYPERLINK("https://mail.google.com/mail/u/0/#all/m-broken","Open")');
  assert.equal(typeof date.getFullYear, 'function');
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /🔎 Unrecognized: 1 email\(s\) need a look/);
  assert.equal(JSON.parse(mock.props.FT_LAST_RUN).unrecognized, 1);
  // run again: updated, not duplicated; "Ignore" is kept and leaves the count
  sheet.getRange(2, 7).setValue('Ignore');
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const again = sheet._rows(10);
  assert.equal(again.length, 1);
  assert.equal(again[0][6], 'Ignore', 'your Status survives the next run');
  assert.equal(JSON.parse(mock.props.FT_LAST_RUN).unrecognized, 0, 'ignored ones are not counted');
});

test('an email read cleanly later drops off the list', () => {
  const { h, mock } = configured();
  assert.equal(h.ctx.recordUnrecognized([{ id: 'a', date: h.date(2026, 9, 1), bank: 'BHD', subject: 's', reason: 'Amount not found', snippet: 'x' },
    { id: 'b', date: h.date(2026, 9, 2), bank: 'BHD', subject: 's', reason: 'Amount not found', snippet: 'y' }], [], h.date(2026, 9, 3)), 2);
  assert.equal(h.ctx.recordUnrecognized([], ['a'], h.date(2026, 9, 4)), 1, 'a newer version read "a"');
  assert.deepEqual(h.plain(mock.ss.getSheetByName('Unrecognized')._rows(10).map(r => r[9])), ['b']);
  assert.equal(h.ctx.recordUnrecognized([], [], h.date(2026, 9, 4)), 1);
});

test('rows saved with an unreadable merchant are listed too; broker emails that can\'t be read are listed', () => {
  const { h, mock } = configured();
  const odd = h.fakeMessage({ subject: 'Alerta de consumo', from: 'notificaciones@popularenlinea.com',
    body: 'Notificación de Consumo\nMonto\tMoneda\tFecha\tEstatus\nRD$250.00\tRD\t06/09/2026\tAprobada\nGracias\n', date: h.date(2026, 9, 6, 9), id: 'm-odd' });
  const hapi = h.fakeMessage({ subject: '✅ Order Executed ', from: 'Hapi App <no-reply@hapi.trade>',
    body: fixture('investments/hapi_order_buy').replace('Order completed', 'Order pending'), date: h.date(2026, 9, 7, 9), id: 'm-hapi' });
  mock.gmail.threads.push(fakeThread('t1', [odd]), fakeThread('t2', [hapi]));
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const rows = h.plain(mock.ss.getSheetByName('Unrecognized')._rows(10));
  const by = Object.fromEntries(rows.map(r => [r[9], r]));
  assert.match(by['m-odd'][3], /^Saved — merchant unreadable/);
  assert.deepEqual([by['m-hapi'][1], by['m-hapi'][3]], ['HAPI', 'Broker email not read: order status "Order pending"']);
  assert.ok(mock.ss.getSheets().map(s => s.getName()).indexOf('Unrecognized') > mock.ss.getSheets().map(s => s.getName()).indexOf('Transactions'));
  const sheet = mock.ss.getSheetByName('Unrecognized');
  assert.ok(sheet.cfRules.some(r => r.whenTextEqualTo === 'New') && sheet.cfRules.some(r => r.whenTextStartsWith === 'Broker email'), 'status and reason chips');
  assert.ok(sheet.cfRules.some(r => r.whenFormulaSatisfied === '=$G2="Ignore"'), 'ignored rows greyed out');
});

test('the daily email\'s data-health line mentions what is waiting in Unrecognized', () => {
  const h = load();
  const d = h.plain(h.ctx.summaryDataHealth({ at: h.date(2026, 9, 25, 6).toISOString(), saved: 3, unrecognized: 2 }, h.date(2026, 9, 25, 8)));
  assert.equal(d.tone, 'warn');
  assert.match(d.text, /2 email\(s\) in Unrecognized/);
});

test('balances: the same account and date again updates it; another date adds a row (v1.1.35)', () => {
  const { h, mock } = configured();
  const add = (date, amount) => h.ctx.addValuationEntry({ account: 'Pension fund', kind: 'Pension', date, mode: 'amount', amount, currency: 'DOP' });
  add('2026-09-08', '890000');
  add('2026-09-08', '892972.13');                       // same statement, corrected
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /Balance updated \(same account and date\)/);
  add('2026-01-01', '807733.26');                       // an earlier statement: the start
  const rows = h.plain(mock.ss.getSheetByName('Investment Ledger')._rows(12).filter(r => r[2] === 'Valuation'));
  assert.deepEqual(rows.map(r => [h.ctx.normalizeDateForCompare(r[0]), r[6]]).sort(), [['2026-01-01', 807733.26], ['2026-09-08', 892972.13]]);
  const perf = mock.ss.getSheetByName('Holdings')._rows(8).find(r => r[1] === 'Pension fund' && typeof r[2] !== 'number');
  assert.equal(h.ctx.normalizeDateForCompare(perf[2]), '2026-01-01', 'tracked since the earliest balance');
});
