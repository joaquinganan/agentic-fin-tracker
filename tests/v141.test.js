'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('vm');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

// Every dialog's page script as the browser gets it: it must compile, and every function its buttons and options call
// must exist. (v1.1.40 shipped a balance dialog whose script didn't load — one apostrophe — so no option switched and
// nothing could be saved; the tests only looked for text in the HTML.)
function checkDialog(html, name) {
  const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
  assert.ok(scripts.length, name + ': has a script');
  scripts.forEach((code, i) => {
    assert.doesNotThrow(() => new vm.Script(code, { filename: name + ' script ' + (i + 1) }), name + ': script ' + (i + 1) + ' compiles');
  });
  const all = scripts.join('\n');
  const handlers = [...html.matchAll(/\son[a-z]+="\s*([A-Za-z_$][\w$]*)\s*\(/g)].map(m => m[1]);
  const missing = [...new Set(handlers)].filter(fn => !new RegExp('function\\s+' + fn + '\\s*\\(').test(all));
  assert.deepEqual(missing, [], name + ': handlers call functions that exist');
  return handlers.length;
}

function book() {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"BHD":true}'], ['monthlyIncome', 1000], ['incomeCurrency', 'USD']].forEach(r => cfg.appendRow(r));
  return { h, mock };
}

test('every dialog\'s page script compiles and its buttons call functions that exist', () => {
  const { h, mock } = book();
  h.ctx.openSetupWizard();
  h.ctx.openDateRangeDialog();
  h.ctx.openValuationDialog();
  assert.equal(mock.ui.dialogs.length, 3);
  for (const d of mock.ui.dialogs) assert.ok(checkDialog(d.html, d.title) > 0, d.title + ': has handlers');
});

test('the check catches what broke in v1.1.40: an apostrophe that ends a string early', () => {
  assert.throws(() => checkDialog("<p onclick=\"go()\"></p><script>function go() { const t = 'the account's keyword'; }</script>", 'broken'));
  assert.throws(() => checkDialog('<p onclick="missing()"></p><script>function go() {}</script>', 'unknown handler'));
});
