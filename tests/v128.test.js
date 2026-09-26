'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture, manifest } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

const EMAIL = 'user@example.com';

/** Same configured system as integration.test.js: every fixture email in Gmail, one custom category ("Pets"). */
function setup() {
  const mock = makeServices();
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', EMAIL], ['monthlyIncome', 3400], ['ARS', 150], ['AFP', 200], ['taxRate', 18],
   ['banksToTrack', JSON.stringify({ LAFISE: true, BANESCO: true, BHD: true, POPULAR: true, BDI: false })],
   ['setupDate', '2026-09-01'], ['incomeCurrency', 'USD']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions');
  h.ctx.initializeTransactionsSheet();
  const rules = mock.ss.insertSheet('Custom Rules');
  [['UserEmail', 'Category', 'Keyword', 'Timestamp'], ['(example — delete this row)', 'Exclude', 'YOUR NAME', ''],
   [EMAIL, 'Pets', 'PETSMART', '']].forEach(r => rules.appendRow(r));
  const msg = fx => h.fakeMessage({ subject: fx.subject, from: fx.from, body: fixture(fx.name), date: h.date(...fx.date), id: 'id-' + fx.name });
  const byName = Object.fromEntries(manifest().map(fx => [fx.name, fx]));
  mock.gmail.threads.push(
    fakeThread('t-bhd', ['bhd_consumo_cacharepa', 'bhd_consumo_medicar'].map(n => msg(byName[n]))),
    fakeThread('t-lafise', [msg(byName.lafise_consumo_ubereats)]),
    fakeThread('t-transfer', [msg(byName.bhd_transfer_self)]));
  return { h, mock };
}

// WCAG 2.x relative luminance / contrast ratio
const lum = hex => {
  const [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255)
    .map(v => v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4));
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
};
const contrast = (a, b) => { const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p); return (x + 0.05) / (y + 0.05); };

test('category colours: one per category, distinct, readable (WCAG AA 4.5:1), custom ones rotate', () => {
  const h = load();
  const palette = h.plain(h.ctx.categoryPalette(['Pets', 'Kids', 'Exclude', 'Rent']));
  const names = palette.map(p => p.name);
  const T = h.plain(h.get('SHEET_THEME'));
  assert.deepEqual(names.slice(0, 10), h.plain(h.ctx.getCategories()), 'every default category, in their order');
  assert.deepEqual(names.slice(10), ['Exclude', 'Pets', 'Kids'], 'Exclude, then custom ones — no duplicates of defaults/Exclude');
  assert.equal(new Set(palette.map(p => p.bg)).size, palette.length, 'no two categories share a background');
  const checks = palette.map(p => [p.name, p.fg, p.bg])
    .concat(Object.entries(h.plain(h.get('TYPE_COLORS'))).map(([t, c]) => [t, c[1], c[0]]))
    .concat([['header', T.headerFg, T.headerBg], ['refund on stripe', T.refund, T.stripe], ['config keys', T.keyFg, T.keyBg]]);
  for (const [what, fg, bg] of checks) assert.ok(contrast(fg, bg) >= 4.5, `${what}: ${fg} on ${bg} = ${contrast(fg, bg).toFixed(2)}:1`);
  // the four chips from the demo video are kept
  const byName = Object.fromEntries(palette.map(p => [p.name, [p.bg, p.fg]]));
  assert.deepEqual(byName['Groceries + Barbershop'], ['#FFF4E5', '#B45309']);
  assert.deepEqual(byName['Transportation'], ['#EAF1FE', '#1D4ED8']);
});

test('data sheets get header, chips and attention rules after a run — and styling raised no error', () => {
  const { h, mock } = setup();
  h.ctx.runGmailMonitorForDateRange('2026-01-01', '2026-12-31');
  assert.deepEqual(h.logs.filter(l => /Could not style/.test(l)), [], 'every sheet styled');
  const tx = mock.ss.getSheetByName('Transactions');
  assert.equal(tx.frozenRows, 1);
  assert.ok(tx.styles.some(s => s.m === 'setBackground' && s.rect.r1 === 1 && s.rect.r2 === 1 && s.v === '#1F3864'), 'navy header');
  const rules = tx.cfRules;
  const colOf = r => r.ranges[0].getColumn();
  assert.match(rules[0].whenFormulaSatisfied, /^=OR\(\$C2="\(unparsed/, 'problem rows first (highest priority)');
  assert.equal(rules[1].whenFormulaSatisfied, '=AND($A2<>"",$L2="Transfer",$F2="")', 'uncategorized transfers next');
  const chips = rules.filter(r => r.whenTextEqualTo && colOf(r) === 6).map(r => [r.whenTextEqualTo, r.setBackground]);
  assert.equal(chips.length, 12, '10 categories + Exclude + Pets on column F');
  assert.deepEqual(chips.find(c => c[0] === 'Pets'), ['Pets', '#CFFAFE']);
  assert.ok(rules.some(r => r.whenTextEqualTo === 'Transfer' && colOf(r) === 12), 'Type chip');
  assert.ok(rules.some(r => r.whenNumberLessThan === 0 && colOf(r) === 4 && r.setFontColor === '#2E7D32'), 'refunds green');
  assert.equal(rules[rules.length - 1].whenFormulaSatisfied, '=AND($A2<>"",ISEVEN(ROW()))', 'stripe last (lowest priority)');
  // every rule covers the whole sheet below the header, not just current rows
  assert.ok(rules.every(r => r.ranges[0].getRow() === 2 && r.ranges[0].getNumRows() === tx.getMaxRows() - 1));
  const raw = mock.ss.getSheetByName('Raw_BHD');
  assert.ok(raw.cfRules.some(r => r.whenTextEqualTo === 'Health + Vet + Pharmacy' && colOf(r) === 4), 'Raw_ sheets: chips on column D');
});

test('Bank Transfers shows each transfer\'s Category; an old sheet is migrated and its filter recreated', () => {
  const { h, mock } = setup();
  // an existing pre-1.1.28 Bank Transfers sheet: 6 columns and a filter criterion on "Amount" (column D)
  const old = mock.ss.insertSheet('Bank Transfers');
  old.appendRow(['Date', 'Bank', 'Beneficiary / Description', 'Amount', 'Currency', 'Email Subject']);
  old.appendRow([h.date(2026, 9, 1), 'BHD', 'OLD ROW', 500, 'DOP', 's']);
  old.getRange(1, 1, 1000, 6).createFilter();
  old.getFilter().criteria[4] = 'amount > 100';
  h.ctx.runGmailMonitorForDateRange('2026-01-01', '2026-12-31');
  const bt = mock.ss.getSheetByName('Bank Transfers');
  assert.deepEqual(h.plain(bt.getRange(1, 1, 1, 7).getValues()[0]),
    ['Date', 'Bank', 'Beneficiary / Description', 'Category', 'Amount', 'Currency', 'Email Subject']);
  const rows = bt._rows(7);                                   // data rows (from row 2)
  assert.ok(rows.length >= 1 && rows.every(r => typeof r[4] === 'number'), 'Amount now in column E');
  assert.ok(!rows.some(r => r[2] === 'OLD ROW'), 'rewritten from Transactions');
  assert.ok(bt.getFilter() && bt.getFilter().getRange().getNumColumns() >= 7, 'filter covers the new column');
  assert.equal(bt.getFilter().criteria[4], undefined, 'the old "Amount" criterion did not land on Category');
  assert.equal(bt.cfRules[1].whenFormulaSatisfied, '=AND($A2<>"",$D2="")', 'transfers without a category highlighted');
});

test('Custom Rules: a dropdown of every category (new names allowed), chips, example row muted', () => {
  const { h, mock } = setup();
  h.ctx.runGmailMonitorForDateRange('2026-01-01', '2026-12-31');
  const cr = mock.ss.getSheetByName('Custom Rules');
  const dv = h.plain(cr.styles.find(s => s.m === 'setDataValidation').v);
  assert.deepEqual(dv.list.slice(-2), ['Exclude', 'Pets']);
  assert.equal(dv.list.length, 12);
  assert.equal(dv.allowInvalid, true, 'typing a new category still works');
  assert.match(dv.helpText, /type a new name/);
  assert.ok(cr.cfRules.some(r => r.whenFormulaSatisfied === '=LEFT($A2,8)="(example"' && r.setItalic === true));
  assert.ok(cr.cfRules.some(r => r.whenTextEqualTo === 'Pets' && r.ranges[0].getColumn() === 2));
  const cfg = mock.ss.getSheetByName('Configuration');
  assert.ok(cfg.styles.some(s => s.m === 'setNote' && /Setup Wizard/.test(s.v)));
});

test('Categories sheet shows each category in its chip colours', () => {
  const { h, mock } = setup();
  h.ctx.buildOrRefreshCategoriesSheet(EMAIL);
  const sheet = mock.ss.getSheetByName('Categories');
  const bg = h.plain(sheet.styles.find(s => s.m === 'setBackgrounds' && s.rect.c1 === 2).v.map(r => r[0]));
  const all = sheet._rows(4);
  const names = all.slice(all.findIndex(r => r[1] === 'Category') + 1).map(r => r[1]).filter(Boolean);
  const palette = Object.fromEntries(h.plain(h.ctx.categoryPalette(['Pets'])).map(p => [p.name, p.bg]));
  assert.deepEqual(bg, names.map(n => palette[n]));
  assert.ok(names.includes('Pets') && names.includes('Exclude'));
});
