'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const vm = require('vm');
const { build, FILES, OUT } = require('../tools/bundle');

test('dist/FinancialTracker.gs is exactly what src/ builds (run `npm run bundle` after changing src/)', () => {
  assert.ok(fs.existsSync(OUT), 'dist/FinancialTracker.gs exists');
  assert.equal(fs.readFileSync(OUT, 'utf8'), build());
});

test('the single file: every source file, in load order, under a header with the version', () => {
  const text = build();
  const version = fs.readFileSync(require('path').join(__dirname, '..', 'src', '01_main.gs'), 'utf8').match(/const SCRIPT_VERSION = "([^"]+)"/)[1];
  assert.match(text.split('\n')[1], new RegExp('Financial Tracker v' + version.replace(/\./g, '\\.')));
  const at = FILES.map(f => text.indexOf('\n// ' + f + '\n'));
  assert.ok(at.every(i => i > 0), 'every file is in it');
  assert.deepEqual(at.slice().sort((a, b) => a - b), at, 'in load order');
  assert.doesNotThrow(() => new vm.Script(text, { filename: 'FinancialTracker.gs' }), 'compiles as one script');
  assert.equal((text.match(/const SCRIPT_VERSION = /g) || []).length, 1, 'declared once');
});
