'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');
const { FILES } = require('../tools/bundle');

// "📊 Tracker › X" in the code sends people to a menu item: X must exist. (Holdings still said "Add Fund / Pension
// Balance" ten versions after that item was renamed — found reviewing the demo video.)
test('every "📊 Tracker › …" the tracker shows names a menu item that exists', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  h.ctx.onOpen();
  const bare = s => s.replace(/^[^A-Za-z]+/, '').trim();
  const items = mock.ui.menus[0].items.filter(i => i.caption).map(i => bare(i.caption));
  const refs = [];
  FILES.forEach(f => {
    const code = fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8');
    for (const m of code.matchAll(/📊 Tracker › ([^'"`<\\]+)/g)) refs.push({ file: f, text: bare(m[1]) });
  });
  assert.ok(refs.length >= 8, 'references found');
  const broken = refs.filter(r => !items.some(i => r.text.startsWith(i))).map(r => r.file + ': ' + r.text.slice(0, 50));
  assert.deepEqual(broken, []);
});
