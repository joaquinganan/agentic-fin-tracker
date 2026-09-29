'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

// Incoming Transfers is built like Bank Transfers and must be treated like it everywhere a list of sheets decides —
// v1.1.51 left it out of the column fitting, the sort and the reset (reported: "not formatted with the recategorizer").
function book() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"BANESCO":true}']].forEach(r => cfg.appendRow(r));
  for (const name of ['Bank Transfers', 'Incoming Transfers']) {
    const s = h.ctx.getOrCreateTransfersSheet(name);
    s.getRange(2, 1, 3, 8).setValues([   // out of date order, a long name
      [h.date(2026, 8, 3), 'BANESCO', 'SOMEONE WITH A VERY LONG NAME INDEED', '', 100, 'DOP', 'subject', 'a'],
      [h.date(2026, 8, 24), 'BANESCO', 'B', '', 200, 'DOP', 'subject', 'b'],
      [h.date(2026, 8, 17), 'BANESCO', 'C', '', 300, 'DOP', 'subject', 'c']]);
  }
  return { h, mock };
}

test('Incoming Transfers gets what Bank Transfers gets: styling, fitted columns, date order', () => {
  const { h, mock } = book();
  h.ctx.formatDataSheets();
  h.ctx.sortAllDataSheets();
  const bt = mock.ss.getSheetByName('Bank Transfers'), inc = mock.ss.getSheetByName('Incoming Transfers');
  assert.equal(inc.cfRules.length, bt.cfRules.length, 'the same conditional formats');
  assert.ok(Object.keys(inc.widths).length > 0, 'columns fitted');
  assert.deepEqual(inc.widths, bt.widths, 'the same widths for the same content');
  assert.deepEqual(h.plain(inc._rows(8).map(r => r[7])), ['b', 'c', 'a'], 'newest first');
  assert.deepEqual(h.plain(inc._rows(8).map(r => r[7])), h.plain(bt._rows(8).map(r => r[7])));
});

test('Reset System deletes Incoming Transfers with the rest of the derived sheets', () => {
  const { h, mock } = book();
  h.ctx.getOrCreateLedgerSheet();   // a real book has tabs the reset keeps (Sheets can't delete the last one)
  h.ctx.resetSystem();
  const names = mock.ss.getSheets().map(s => s.getName());
  assert.ok(names.indexOf('Bank Transfers') === -1 && names.indexOf('Incoming Transfers') === -1);
  assert.ok(names.indexOf('Investment Ledger') !== -1, 'investments kept');
  assert.match(mock.ui.alerts[mock.ui.alerts.length - 2], /Bank Transfers, Incoming Transfers/);
});
