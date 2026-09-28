'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

function counted(h, subject, body, id, day) {
  const m = h.fakeMessage({ subject, from: 'notificaciones@banesco.com.do', body, date: h.date(2026, 9, day || 5, 10), id });
  const read = m.getPlainBody; m.reads = 0; m.getPlainBody = () => { m.reads++; return read.call(m); };
  return m;
}
function book() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ BANESCO: true, BHD: true })], ['setupDate', '2026-09-01']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  return { h, mock };
}

test('emails read without saving anything (statements, promotions) are not read again — reported: a year-long range never finished', () => {
  const { h, mock } = book();
  const statement = counted(h, 'Tu estado de cuenta está disponible', 'Estimado cliente, tu estado de cuenta ya está disponible.', 'st1');
  const promo = counted(h, 'Promoción exclusiva', 'Aprovecha 20% de descuento con tu tarjeta', 'pr1');
  const broken = counted(h, 'Alerta de consumo', 'Hola, tu operación fue registrada.', 'bk1');
  mock.gmail.threads.push(fakeThread('t1', [statement]), fakeThread('t2', [promo]), fakeThread('t3', [broken]));
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.deepEqual([statement.reads, promo.reads], [1, 1], 'read once');
  assert.equal(broken.reads, 2, 'an email that failed is retried');
  const log = mock.ss.getSheetByName('Read Emails');
  assert.ok(log.isSheetHidden());
  assert.deepEqual(h.plain(log._rows(2).map(r => r[0])).sort(), ['pr1', 'st1']);
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /Read before, skipped: 2 email/);
  // after an update, they're read once more (the new version may read them differently)
  log.getRange(2, 2, 2, 1).setValues([['1.1.0'], ['1.1.0']]);
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.deepEqual([statement.reads, promo.reads], [2, 2]);
});

test('a stopped run leaves the heavy steps to the next one; the next run does them', () => {
  const { h, mock } = book();
  ['bhd_consumo_cacharepa', 'bhd_consumo_medicar', 'bhd_consumo_pedidosya'].forEach((n, i) => mock.gmail.threads.push(fakeThread('t' + i,
    [h.fakeMessage({ subject: 'BHD Notificación de Transacciones', from: 'Alertas@bhd.com.do', body: fixture(n), date: h.date(2026, 9, 3 + i, 10), id: 'm' + i })])));
  // 2 minutes per check while reading (it stops after the first thread), then time stands still: only the
  // stopped-run rule — not the 5-minute brake — can hold the heavy steps back
  h.get('runClock = (() => { let n = 0, t = 0; return () => (n++ < 3 ? (t += 120000) : t); })()');
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const first = mock.ui.alerts[mock.ui.alerts.length - 1];
  assert.match(first, /⏸ Stopped reading early/);
  assert.match(first, /⏭ Left for the next run: sort, recategorize, investments/);
  assert.ok(!h.logs.some(l => /Auto-recategorize/.test(l)), 'recategorize didn\'t run');
  assert.equal(mock.ss.getSheetByName('Transactions')._rows(15).length, 1, 'what was read is saved');
  h.get('runClock = () => Date.now()');
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const second = mock.ui.alerts[mock.ui.alerts.length - 1];
  assert.doesNotMatch(second, /Left for the next run/);
  assert.ok(h.logs.some(l => /Auto-recategorize/.test(l)));
  assert.equal(mock.ss.getSheetByName('Transactions')._rows(15).length, 3);
});

test('no step starts after 5 minutes, even when reading finished in time', () => {
  const { h, mock } = book();
  mock.gmail.threads.push(fakeThread('t0', [h.fakeMessage({ subject: 'BHD Notificación de Transacciones', from: 'Alertas@bhd.com.do',
    body: fixture('bhd_consumo_cacharepa'), date: h.date(2026, 9, 3, 10), id: 'm0' })]));
  // time jumps 100 s at every check: reading ends in time, but the clock passes 5 minutes during the steps
  h.get('runClock = (() => { let t = 0; return () => (t += 100000); })()');
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const summary = mock.ui.alerts[mock.ui.alerts.length - 1];
  assert.doesNotMatch(summary, /Stopped reading early/);
  assert.match(summary, /⏭ Left for the next run: .*(investments|unrecognized|sheet order)/);
  assert.notEqual(mock.ss.toastLog[mock.ss.toastLog.length - 1].timeout, -1, 'no toast left open');
});

test('ledger rows without a date are flagged — they are not counted (reported: manual deposits without a date)', () => {
  const { h, mock } = book();
  const led = h.ctx.getOrCreateLedgerSheet();
  led.getRange(2, 1, 2, 12).setValues([
    [h.date(2026, 9, 26), 'HAPI', 'Snapshot', 'CASH', '', '', 100, '', 'USD', 'manual', '', ''],
    ['', 'HAPI', 'Deposit', '', '', '', 200, '', 'USD', 'manual', '', '']]);
  const res = h.plain(h.ctx.computeHoldings(led.getDataRange().getValues(), { usdRate: 60, year: 2026 }));
  assert.match(res.warnings[0], /^1 Investment Ledger row\(s\) have no date and are not counted/);
  h.ctx.styleLedgerSheet(led);
  assert.equal(led.cfRules[0].whenFormulaSatisfied, '=AND($A2="",$B2<>"")', 'highlighted, first rule');
  assert.equal(led.cfRules[0].setBackground, '#FCA5A5');
});

test('a deposit from the dialog keeps its date (the check v1.1.40 missed)', () => {
  const { h, mock } = book();
  h.ctx.addValuationEntry({ account: 'HAPI', kind: 'Broker', date: '2026-06-12', mode: 'deposit', amount: '59.99', currency: 'USD' });
  const row = mock.ss.getSheetByName('Investment Ledger')._rows(12)[0];
  assert.equal(typeof row[0].getFullYear, 'function', 'a real date');
  assert.equal(h.ctx.normalizeDateForCompare(row[0]), '2026-06-12');
});
