'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

function book() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ BHD: true, BANESCO: true })]].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  mock.gmail.threads.push(
    fakeThread('t1', [h.fakeMessage({ subject: 'BHD Notificación de Transacciones', from: 'Alertas@bhd.com.do', body: fixture('bhd_consumo_cacharepa'), date: h.date(2026, 9, 3, 10), id: 'm1' })]),
    fakeThread('t2', [h.fakeMessage({ subject: 'Tu estado de cuenta está disponible', from: 'notificaciones@banesco.com.do', body: 'Tu estado de cuenta ya está disponible.', date: h.date(2026, 9, 4, 10), id: 'st1' })]));
  return { h, mock };
}

test('a transaction row you delete is read again on the next run — only emails with nothing to save go to the read log', () => {
  const { h, mock } = book();
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const tx = mock.ss.getSheetByName('Transactions');
  assert.equal(tx._rows(15).length, 1);
  assert.deepEqual(h.plain(mock.ss.getSheetByName('Read Emails')._rows(2).map(r => r[0])), ['st1'], 'the statement only');
  tx.getRange(2, 1, 1, 15).clearContent();
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.equal(tx._rows(15).filter(r => r[0]).length, 1, 'back from its email');
});

test('Reset System starts over for real: a new setup and a date range bring every transaction back; investments are kept', () => {
  const { h, mock } = book();
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  h.ctx.getOrCreateLedgerSheet().getRange(2, 1, 1, 12).setValues([[h.date(2026, 9, 1), 'HAPI', 'Snapshot', 'CASH', '', '', 10, '', 'USD', 'manual', '', '']]);
  h.ctx.recordUnrecognized([{ id: 'x', date: h.date(2026, 9, 2), bank: 'BHD', subject: 's', reason: 'Amount not found', snippet: '' }], [], h.date(2026, 9, 5));
  h.ctx.resetSystem();
  const confirm = mock.ui.alerts[mock.ui.alerts.length - 2];
  assert.match(confirm, /cannot be undone/);
  assert.match(confirm, /Kept: the investment tabs/);
  const names = mock.ss.getSheets().map(s => s.getName());
  for (const gone of ['Configuration', 'Transactions', 'Custom Rules', 'Unrecognized', 'Read Emails']) assert.ok(names.indexOf(gone) === -1, gone + ' deleted');
  assert.ok(names.indexOf('Investment Ledger') !== -1, 'investment ledger kept');
  assert.equal(mock.props.FT_LAST_RUN, undefined, 'no "last update" from before the reset');
  // set up again and re-read: the transaction comes back
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ BHD: true, BANESCO: true })]].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.equal(mock.ss.getSheetByName('Transactions')._rows(15).length, 1);
});

test('the Spanish guide names every menu item as it appears, and its links point to files that exist', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  h.ctx.onOpen();
  const guide = fs.readFileSync(path.join(__dirname, '..', 'docs', 'GUIA.md'), 'utf8');
  const missing = mock.ui.menus[0].items.filter(i => i.caption).map(i => i.caption).filter(c => !guide.includes(c));
  assert.deepEqual(missing, [], 'menu items the guide does not mention');
  const links = [...guide.matchAll(/\]\((\.\.\/[^)#]+)\)/g)].map(m => m[1]);
  assert.ok(links.length > 0);
  assert.deepEqual(links.filter(l => !fs.existsSync(path.join(__dirname, '..', 'docs', l))), []);
});

test('the monthly summary of a month before the investment history says when it starts, instead of nothing (reported)', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  h.ctx.getOrCreateLedgerSheet().getRange(2, 1, 1, 12).setValues([[h.date(2026, 9, 26), 'HAPI', 'Snapshot', 'CASH', '', '', 10, '', 'USD', 'manual', '', '']]);
  h.ctx.recordPortfolioHistory([{ account: 'HAPI', ticker: 'CASH', qty: '', price: '', value: 10 }], '2026-09-26');
  const august = h.plain(h.ctx.investmentsReportData('monthly', { rates: { USD: 60 }, month: h.date(2026, 8, 1) }));
  assert.deepEqual(august, { none: true, firstDay: '2026-09-26', firstReport: '2026-10-01' });
  const m = h.ctx.computeMonthlySummary([new Array(15).fill('h')], { today: h.date(2026, 9, 28), rates: { USD: 60, EUR: 64, COP: 0.015 },
    netIncomeDop: 80000, cards: [], investments: august });
  const decode = x => x.replace(/&#(\d+);/g, (a, n) => String.fromCodePoint(Number(n)));
  const html = decode(h.ctx.buildMonthlySummaryEmail(m, { sheetName: 'T' }).html);
  assert.ok(html.includes('Investments') && html.includes('starts on 2026-09-26') && html.includes('sent on 2026-10-01'));
  const september = h.plain(h.ctx.investmentsReportData('monthly', { rates: { USD: 60 }, month: h.date(2026, 9, 1) }));
  assert.equal(september.none, undefined, 'a month with history gets its figures');
  assert.equal(september.endValue, 10);
});
