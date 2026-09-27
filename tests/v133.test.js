'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

function configured() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ BHD: true })], ['setupDate', '2026-09-01']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions');
  h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  return { h, mock };
}
const lastToast = mock => mock.ss.toastLog[mock.ss.toastLog.length - 1];

test('no action leaves a progress toast open — when it finishes, and when it fails (v1.1.33)', () => {
  // the reported case: "Refreshing investments..." stayed on screen after its summary
  let { h, mock } = configured();
  h.ctx.refreshInvestmentsNow();
  assert.ok(mock.ss.toastLog.some(t => t.msg === 'Refreshing investments...' && t.timeout === -1));
  assert.notEqual(lastToast(mock).timeout, -1, 'closed after refreshing');
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /Investments refreshed/);
  // an error half-way through a run (Gmail unavailable): the progress toast is closed too
  ({ h, mock } = configured());
  mock.services.GmailApp.search = () => { throw new Error('Gmail is unavailable'); };
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.ok(mock.ss.toastLog.some(t => t.timeout === -1), 'a progress toast was shown');
  assert.notEqual(lastToast(mock).timeout, -1, 'and closed after the error');
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /Gmail is unavailable/);
  // a normal run ends with "Done." (not an extra "Finished.")
  ({ h, mock } = configured());
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  assert.equal(lastToast(mock).msg, 'Done.');
});

test('balance entries: validated, then saved as a Valuation row', () => {
  const h = load();
  const ok = { account: 'Liquidity fund', kind: 'Fund', date: '2026-08-31', mode: 'units', units: '50.5', unitPrice: '2277.6', currency: 'DOP' };
  assert.equal(h.ctx.validateValuationEntry(ok, '2026-09-26'), null);
  assert.match(h.ctx.validateValuationEntry(Object.assign({}, ok, { account: ' ' }), '2026-09-26'), /name/);
  assert.match(h.ctx.validateValuationEntry(Object.assign({}, ok, { date: '2026-10-02' }), '2026-09-26'), /future/);
  assert.match(h.ctx.validateValuationEntry(Object.assign({}, ok, { unitPrice: '0' }), '2026-09-26'), /greater than 0/);
  assert.match(h.ctx.validateValuationEntry({ account: 'Pension', date: '2026-08-31', mode: 'amount', amount: '', currency: 'DOP' }, '2026-09-26'), /balance must be/);
  const row = h.plain(h.ctx.valuationRow(ok));
  assert.deepEqual(row.slice(1), ['Liquidity fund', 'Valuation', '', 50.5, 2277.6, '', '', 'DOP', 'manual', '', '']);
  assert.equal(h.ctx.normalizeDateForCompare(h.ctx.valuationRow(ok)[0]), '2026-08-31');
  const pension = h.plain(h.ctx.valuationRow({ account: 'Pension', date: '2026-08-31', mode: 'amount', amount: '800000', currency: 'DOP' }));
  assert.deepEqual([pension[4], pension[5], pension[6]], ['', '', 800000]);
});

test('Add Fund / Pension Balance: a new row every statement, the account is listed, Holdings shows the latest', () => {
  const { h, mock } = configured();
  h.ctx.buildOrRefreshDashboard();
  const entry = (date, units, price) => ({ account: 'Liquidity fund', kind: 'Fund', date, mode: 'units', units, unitPrice: price, currency: 'DOP', notes: '' });
  assert.equal(h.ctx.addValuationEntry(entry('2026-07-31', '50', '2262.1')), true);
  assert.equal(h.ctx.addValuationEntry(entry('2026-08-31', '50', '2277.6')), true);
  assert.equal(h.ctx.addValuationEntry({ account: 'Pension fund', kind: 'Pension', date: '2026-08-31', mode: 'amount', amount: '800000', currency: 'DOP' }), true);
  const ledger = mock.ss.getSheetByName('Investment Ledger')._rows(12);
  assert.equal(ledger.filter(r => r[2] === 'Valuation').length, 3, 'rows are added, never overwritten');
  const accounts = mock.ss.getSheetByName('Investment Accounts')._rows(4).map(r => [r[0], r[1]]);
  assert.deepEqual(h.plain(accounts.filter(a => a[0] !== 'HAPI')), [['Liquidity fund', 'Fund'], ['Pension fund', 'Pension']], 'listed once each');
  const holdings = mock.ss.getSheetByName('Holdings')._rows(8);
  const fund = holdings.find(r => r[1] === 'Liquidity fund');
  assert.equal(fund[2], 50 * 2277.6, 'the latest statement');
  assert.ok(holdings.find(r => r[1] === 'Pension fund'));
  assert.notEqual(lastToast(mock).timeout, -1);
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /Balance saved[\s\S]*Pension fund — DOP 800,000.00 as of 2026-08-31/);
  // invalid: nothing saved
  assert.equal(h.ctx.addValuationEntry({ account: '', date: '2026-08-31', mode: 'amount', amount: '5', currency: 'DOP' }), false);
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 1], /Balance not saved: Give the account a name/);
  assert.equal(mock.ss.getSheetByName('Investment Ledger')._rows(12).length, ledger.length);
});

test('the balance dialog: fields, the accounts already used, and it closes while saving', () => {
  const { h, mock } = configured();
  h.ctx.addValuationEntry({ account: 'Liquidity fund', kind: 'Fund', date: '2026-08-31', mode: 'amount', amount: '1000', currency: 'DOP' });
  mock.ui.dialogs.length = 0;
  h.ctx.openValuationDialog();
  const d = mock.ui.dialogs[0];
  for (const piece of ['id="account"', '<option value="Liquidity fund">', 'id="units"', 'id="unitPrice"', 'id="amount"', 'type="date"',
                       'addValuationEntry(e)', 'google.script.host.close()', 'valor cuota']) {
    assert.ok(d.html.includes(piece), piece);
  }
  assert.match(d.title, /Fund or pension balance/);
});

test('onOpen builds the 📊 Tracker menu, and every item calls a function that exists', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  h.ctx.onOpen();
  assert.equal(mock.ui.menus.length, 1);
  const menu = mock.ui.menus[0];
  assert.equal(menu.caption, '📊 Tracker');
  const items = menu.items.filter(i => i.fn);
  assert.ok(items.length >= 11);
  const missing = items.filter(i => h.get('typeof ' + i.fn) !== 'function').map(i => i.fn);
  assert.deepEqual(missing, [], 'a menu item pointing to a missing function fails only when clicked');
  for (const fn of ['refreshInvestmentsNow', 'openValuationDialog', 'openSetupWizard']) assert.ok(items.some(i => i.fn === fn), fn);
});
