'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('vm');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

const NOW = [2026, 9, 28, 9];
const byId = r => Object.fromEntries(r.steps.map(s => [s.id, s]));

test('a new sheet: settings to do, nothing scheduled yet, investments optional', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const r = h.plain(h.ctx.startHereStatus(h.date(...NOW)));
  const s = byId(r);
  assert.equal(s.code.status, 'done');
  assert.match(s.code.title, /^Tracker v\d+\.\d+\.\d+ is installed$/);
  assert.deepEqual([s.settings.status, s.settings.action.fn], ['todo', 'openSetupWizard']);
  assert.equal(s.schedule.status, 'todo');
  assert.equal(s.firstRun.status, 'optional');
  assert.equal(s.investments.status, 'optional');
  assert.deepEqual([r.done, r.problems], [2, 0]);   // code + "every email was read"
  assert.ok(r.steps.every(x => typeof x.title === 'string' && typeof x.detail === 'string'), 'plain data for the sidebar');
});

function configured(opts) {
  opts = opts || {};
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 95000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ BHD: true, BANESCO: true })], ['notifyEnabled', true], ['notifyHour', 8]].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  if (!opts.noTriggers) h.ctx.createTrigger();
  return { h, mock };
}

test('a configured sheet with real problems: each one shows, with what to do', () => {
  const { h, mock } = configured();
  h.ctx.recordLastRun({ at: h.date(2026, 9, 28, 6).toISOString(), saved: 3, duplicates: 0, failed: 0, unparsed: 0, errors: 0, partial: true });
  h.ctx.recordUnrecognized([{ id: 'a', date: h.date(2026, 9, 2), bank: 'BHD', subject: 's', reason: 'Amount not found', snippet: '' },
    { id: 'b', date: h.date(2026, 9, 3), bank: 'BHD', subject: 's', reason: 'Amount not found', snippet: '' },
    { id: 'c', date: h.date(2026, 9, 4), bank: 'BHD', subject: 's', reason: 'Amount not found', snippet: '' }], [], h.date(2026, 9, 5));
  mock.ss.getSheetByName('Unrecognized').getRange(4, 7).setValue('Ignore');
  const tx = new Array(15).fill(''); tx[0] = h.date(2026, 9, 10); tx[1] = 'BANESCO'; tx[2] = 'JUAN PEREZ'; tx[3] = 500; tx[4] = 'DOP'; tx[11] = 'Transfer';
  mock.ss.getSheetByName('Transactions').getRange(2, 1, 1, 15).setValues([tx]);
  const led = h.ctx.getOrCreateLedgerSheet();
  led.getRange(2, 1, 4, 12).setValues([
    [h.date(2026, 9, 26), 'HAPI', 'Snapshot', 'CASH', '', '', 10, '', 'USD', 'manual', '', ''],
    [h.date(2026, 7, 15), 'Pension fund', 'Valuation', '', '', '', 700000, '', 'DOP', 'manual', '', ''],   // 75 days old
    ['', 'HAPI', 'Deposit', '', '', '', 50, '', 'USD', 'manual', '', ''],                                   // no date
    [h.date(2026, 9, 27), 'HAPI', 'Notice', '', '', '', '', '', 'USD', 'email', '', 'gmail:n1']]);          // after the snapshot, no deposit near it
  const s = byId(h.plain(h.ctx.startHereStatus(h.date(...NOW))));
  assert.deepEqual([s.settings.status, s.schedule.status, s.dailyEmail.status], ['done', 'done', 'done']);
  assert.equal(s.lastRun.status, 'warn');
  assert.match(s.lastRun.detail, /stopped early/);
  assert.deepEqual(s.lastRun.action, { label: 'Monitor by Date Range', fn: 'openDateRangeDialog' });
  assert.deepEqual([s.unrecognized.status, s.unrecognized.title], ['warn', '2 email(s) could not be read'], 'Ignore not counted');
  assert.deepEqual(s.unrecognized.action, { label: 'Open Unrecognized', fn: 'openSheetByName', arg: 'Unrecognized' });
  assert.deepEqual([s.transfers.status, s.transfers.title], ['warn', '1 transfer(s) this month without a category']);
  assert.equal(s['start-HAPI'].status, 'todo');
  assert.equal(s['start-HAPI'].action.fn, 'openValuationDialog');
  assert.equal(s.undated.status, 'error');
  assert.equal(s.deposits.status, 'warn');
  assert.deepEqual([s['balance-Pension fund'].status, s['balance-Pension fund'].detail.slice(0, 11)], ['warn', '75 days ago']);
  // HAPI's value on Jan 1 recorded → its start is done
  led.getRange(6, 1, 1, 12).setValues([[h.date(2025, 12, 31), 'HAPI', 'Valuation', '', '', '', 3000, '', 'USD', 'manual', '', '']]);
  assert.equal(byId(h.plain(h.ctx.startHereStatus(h.date(...NOW))))['start-HAPI'].status, 'done');
});

test('a sheet whose daily update or summary email is not scheduled shows it as a problem', () => {
  const { h } = configured({ noTriggers: true });
  const r = h.plain(h.ctx.startHereStatus(h.date(...NOW)));
  const s = byId(r);
  assert.equal(s.schedule.status, 'error');
  assert.equal(s.dailyEmail.status, 'error');
  assert.equal(r.problems, 2);
  assert.equal(s.firstRun.status, 'todo', 'never ran: read this year\'s emails');
});

test('the sidebar: its script compiles, and every button calls an allowed server function that exists', () => {
  const { h, mock } = configured();
  h.ctx.openStartHere();
  const d = mock.ui.dialogs.pop();
  assert.equal(d.sidebar, true);
  assert.equal(d.title, '📘 Start here');
  const script = d.html.match(/<script>([\s\S]*?)<\/script>/)[1];
  assert.doesNotThrow(() => new vm.Script(script));
  const handlers = [...d.html.matchAll(/\son[a-z]+="\s*([A-Za-z_$][\w$]*)\s*\(/g)].map(m => m[1]);
  assert.deepEqual([...new Set(handlers)].filter(fn => !new RegExp('function\\s+' + fn + '\\s*\\(').test(script)), []);
  const calls = [...new Set([...script.matchAll(/run\.([A-Za-z_$][\w$]*)\(/g)].map(m => m[1]))].filter(fn => !/^with/.test(fn));   // withSuccessHandler etc. are the client API
  const allowed = h.plain(h.get('START_HERE_ACTIONS')).concat(['startHereStatus']);
  assert.deepEqual(calls.filter(fn => allowed.indexOf(fn) === -1), [], 'only allowed functions');
  assert.deepEqual(calls.filter(fn => h.get('typeof ' + fn) !== 'function'), [], 'all exist');
});

test('the sidebar only opens the tabs it lists, and shows a hidden one', () => {
  const { h, mock } = configured();
  mock.ss.insertSheet('Bank Transfers').hideSheet();
  assert.equal(h.ctx.openSheetByName('Bank Transfers'), true);
  assert.equal(mock.ss.getSheetByName('Bank Transfers').isSheetHidden(), false);
  assert.equal(h.ctx.openSheetByName('Configuration'), false, 'not on the list');
});

test('the menu starts with 📘 Start here, and the Setup Wizard points to it when done', () => {
  const { h, mock } = configured();
  h.ctx.onOpen();
  assert.deepEqual(mock.ui.menus[0].items[0], { caption: '📘 Start here', fn: 'openStartHere' });
  assert.match(h.ctx.setupSavedMessage(h.ctx.getConfig(), mock.ss), /Next: 📊 Tracker › 📘 Start here/);
});
