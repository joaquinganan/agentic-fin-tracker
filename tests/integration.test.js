'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture, manifest } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

const EMAIL = 'user@example.com';
const DINING = 'Dining/Delivery + Entertainment + Other';

/** A configured system with every fixture email sitting in Gmail. */
function setup(options) {
  const mock = makeServices(options);
  const h = load({ services: mock.services });
  const bare = options && options.bare;   // no standard fixture threads
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', EMAIL], ['monthlyIncome', 3400], ['ARS', 150], ['AFP', 200], ['taxRate', 18],
   ['banksToTrack', JSON.stringify({ LAFISE: true, BANESCO: true, BHD: true, POPULAR: true, BDI: false })],
   ['setupDate', '2026-09-01'], ['incomeCurrency', 'USD']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions');
  h.ctx.initializeTransactionsSheet();
  const rules = mock.ss.insertSheet('Custom Rules');
  [['UserEmail', 'Category', 'Keyword', 'Timestamp'],
   [EMAIL, 'Exclude', 1111, ''],            // numeric keyword (C4)
   [EMAIL, 'Pets', 'PETSMART', '']].forEach(r => rules.appendRow(r));

  const msg = fx => h.fakeMessage({ subject: fx.subject, from: fx.from, body: fixture(fx.name),
    date: h.date(...fx.date), id: 'id-' + fx.name });
  const byName = Object.fromEntries(manifest().map(fx => [fx.name, fx]));
  // Gmail groups same-subject BHD alerts into ONE thread spanning months (M6)
  if (!bare) mock.gmail.threads.push(
    fakeThread('t-bhd', ['bhd_consumo_cacharepa', 'bhd_consumo_medicar', 'bhd_consumo_pedidosya'].map(n => msg(byName[n]))),
    fakeThread('t-banesco', ['banesco_consumo_hola_plaza', 'banesco_consumo_declined'].map(n => msg(byName[n]))),
    fakeThread('t-lafise', [msg(byName.lafise_consumo_ubereats)]),
    fakeThread('t-popular', [msg(byName.popular_consumo_casita)]),
    fakeThread('t-bhd-transfer', [msg(byName.bhd_transfer_self)]),
    fakeThread('t-broken', [h.fakeMessage({ subject: 'Servicio de Alerta - Nuevo Consumo',
      from: 'notificaciones@bancolafise.com', body: 'Contenido sin montos', date: h.date(2026, 9, 3), id: 'id-broken' })])
  );
  return { h, mock, msg, byName, tx: () => mock.ss.getSheetByName('Transactions')._rows(14) };
}

test('date-range run: exact epoch bounds, only in-range messages, raw sheets rebuilt (C1, M6)', () => {
  const { h, mock, tx } = setup();
  h.ctx.runGmailMonitorForDateRange('2026-03-01', '2026-03-31');

  const start = Math.floor(new Date(2026, 2, 1).getTime() / 1000);
  const endExclusive = Math.floor(new Date(2026, 3, 1).getTime() / 1000);
  assert.ok(mock.gmail.queries[0].includes(`after:${start} before:${endExclusive}`), mock.gmail.queries[0]);

  const rows = tx();
  assert.deepEqual(rows.map(r => r[2]).sort(), ['CACHAREPA CHURCHILL', 'FCIA MEDICAR GBC 30 DE MA'],
    'May message in the same BHD thread is out of range and skipped');
  assert.ok(rows.every(r => r[11] === 'Transaction'));
  assert.equal(mock.ss.getSheetByName('Raw_BHD')._rows(8).length, 2);
  assert.ok(h.logs.some(l => l.includes('Email threads found')), 'summary reached safeAlert → Logger (no UI in tests)');
  assert.equal(mock.lock.held, false, 'lock released');
});

test('full-year run: every fixture lands in the right place; re-run is fully deduplicated', () => {
  const { h, mock, tx } = setup();
  h.ctx.runGmailMonitorForDateRange('2026-01-01', '2026-12-31');
  const rows = tx();
  const byMerchant = Object.fromEntries(rows.map(r => [r[2], r]));

  assert.equal(rows.length, 7, 'declined BANESCO and the broken email add nothing: ' + rows.map(r => r[2]).join(' | '));
  assert.equal(byMerchant['HOLA PLAZA LAS AMERICAS'][5], 'Groceries + Barbershop');
  assert.equal(byMerchant['CACHAREPA CHURCHILL'][11], 'Transaction');
  assert.equal(byMerchant['LA CASITA DE'][5], 'Health + Vet + Pharmacy');
  assert.equal(byMerchant['UBER EATS-W*UBER EATS- SANTO DOMINGO DOM'][5], DINING);
  assert.equal(byMerchant['JUAN PEREZ'][11], 'Transfer');
  assert.ok(!rows.some(r => r[3] === 39766.33), 'balance never saved');

  assert.equal(mock.ss.getSheetByName('Bank Transfers')._rows(6).length, 1);
  assert.equal(mock.ss.getSheetByName('Raw_BHD')._rows(8).length, 3);
  assert.ok(!mock.gmail.markedRead.includes('t-broken'), 'unparseable thread stays unread (M4)');
  assert.ok(mock.gmail.markedRead.includes('t-bhd'));
  assert.ok(mock.gmail.labels['Procesado'], 'label created (M4)');
  const summary = h.logs.find(l => l.includes('Email threads found'));
  assert.match(summary, /Could not parse: 1 email/);
  assert.match(summary, /declined 1/);

  h.ctx.runGmailMonitorForDateRange('2026-01-01', '2026-12-31');
  assert.equal(tx().length, 7, 'second run saves nothing new');
  // v1.1.36: already-saved emails are skipped before being read (they used to be read and dropped as duplicates)
  const second = h.logs.filter(l => l.includes('Email threads found')).pop();
  assert.match(second, /Saved: 0 \| Duplicates: 0/);
  assert.match(second, /Read before, skipped: 8 email\(s\)/);   // v1.1.43: the declined one too (read log)
});

test('recategorize repairs a mis-typed row and keeps Raw_ notes (C3, C5, M7)', () => {
  const { h, mock, tx } = setup();
  h.ctx.runGmailMonitorForDateRange('2026-03-01', '2026-03-31');
  const txSheet = mock.ss.getSheetByName('Transactions');
  const rowIdx = tx().findIndex(r => r[2] === 'CACHAREPA CHURCHILL') + 2;
  txSheet.getRange(rowIdx, 12).setValue('Transfer');   // what v1.1.18 produced
  txSheet.getRange(rowIdx, 6).setValue('');

  const raw = mock.ss.getSheetByName('Raw_BHD');
  const medicarRow = raw._rows(8).findIndex(r => r[2] === 'FCIA MEDICAR GBC 30 DE MA') + 2;
  raw.getRange(medicarRow, 7).setValue('receta mensual');   // user note

  h.ctx.recategorizeAllTransactions(EMAIL);
  const fixed = tx().find(r => r[2] === 'CACHAREPA CHURCHILL');
  assert.equal(fixed[11], 'Transaction');
  assert.equal(fixed[5], DINING);
  assert.equal(fixed[2], 'CACHAREPA CHURCHILL', 'merchant untouched');
  const rawRows = mock.ss.getSheetByName('Raw_BHD')._rows(8);
  assert.equal(rawRows.length, 2, 'CACHAREPA is back in Raw_BHD');
  assert.equal(rawRows.find(r => r[2] === 'FCIA MEDICAR GBC 30 DE MA')[6], 'receta mensual', 'note survived the rebuild');

  assert.equal(h.ctx.recategorizeAllTransactions(EMAIL), 0, 'a second pass changes nothing (stable)');
});

// ---------- Dashboard (v1.1.21 layout)
const SHEETS_FUNCTIONS = new Set(['IF', 'IFERROR', 'INDEX', 'MATCH', 'MONTH', 'TODAY', 'DATE', 'EDATE', 'SUMIFS',
  'SUM', 'SUMPRODUCT', 'AVERAGEIF', 'TEXT', 'UPPER', 'SPARKLINE', 'MAX', 'MIN', 'GOOGLEFINANCE', 'COLUMN', 'ROW', 'LEFT']);

/** Static checks a formula must pass before it ever reaches Sheets. Returns a list of problems. */
function lintFormula(f, namedRanges) {
  const problems = [];
  let paren = 0, brace = 0, inStr = false;
  for (const ch of f) {
    if (ch === '"') { inStr = !inStr; continue; }
    if (inStr) continue;
    if (ch === '(') paren++;
    if (ch === ')') paren--;
    if (ch === '{') brace++;
    if (ch === '}') brace--;
    if (paren < 0 || brace < 0) problems.push('closes before it opens');
  }
  if (inStr) problems.push('unclosed string');
  if (paren || brace) problems.push('unbalanced () or {}');
  const code = f.replace(/"[^"]*"/g, '""');
  for (const m of code.matchAll(/([A-Z][A-Z0-9_.]*)\(/g)) {
    if (!SHEETS_FUNCTIONS.has(m[1])) problems.push('unknown function ' + m[1]);
  }
  for (const m of code.matchAll(/(?<![A-Za-z0-9_$!:])([A-Z][A-Z0-9_]*)(?![A-Za-z0-9_(!$:])/g)) {
    const id = m[1];
    if (/^[A-Z]{1,3}\d+$/.test(id)) continue;             // a cell reference like C13
    if (!namedRanges[id]) problems.push('unknown name ' + id);
  }
  return problems;
}

function dashboardFormulas(sheet) {
  const out = [];
  for (const [key, v] of sheet.cells) if (typeof v === 'string' && v.startsWith('=')) out.push([key, v]);
  return out;
}

test('Dashboard: every formula passes the lint; KPIs and tables point at the right rows (v1.1.21)', () => {
  const { h, mock } = setup();
  h.ctx.runGmailMonitorForDateRange('2026-01-01', '2026-12-31');
  h.ctx.buildOrRefreshDashboard();
  const dash = mock.ss.getSheetByName('Dashboard');

  const formulas = dashboardFormulas(dash);
  assert.ok(formulas.length > 250, 'formulas written: ' + formulas.length);
  const bad = formulas.map(([k, f]) => [k, f, lintFormula(f, mock.ss.namedRanges)]).filter(x => x[2].length);
  assert.deepEqual(bad, [], 'formula problems');
  for (const rule of dash.cfRules) if (rule.whenFormulaSatisfied) {
    assert.deepEqual(lintFormula(rule.whenFormulaSatisfied, mock.ss.namedRanges), []);
  }

  const cell = a1 => dash.getRange(a1).getValue();
  // KPI "spent" points at the left table's TOTAL row, which sums every category + transfers row
  const totalRow = Number(cell('C8').match(/^=\$C\$(\d+)$/)[1]);
  assert.equal(dash.getRange(totalRow, 2).getValue(), 'TOTAL');
  // each left-table row reads the grid row of the SAME category (icon stripped)
  const labels = h.plain(h.ctx.getCategories()).concat(['Pets']);
  for (let i = 0; i < labels.length; i++) {
    const r = 13 + i;
    const g = Number(dash.getRange(r, 3).getValue().match(/^=INDEX\(\$C\$(\d+):/)[1]);
    assert.equal(dash.getRange(g, 2).getValue(), labels[i], 'row ' + r);
    assert.ok(String(dash.getRange(r, 2).getValue()).endsWith(labels[i]));
  }
  // grid headers are midnight DATE() formulas following the Year cell (M9)
  const gridHeaderRow = [...dash.cells.entries()].find(([k, v]) => v === '=DATE($F$4,1,1)')[0].split(',')[0];
  assert.equal(dash.getRange(Number(gridHeaderRow), 2).getValue(), 'Category');

  // named ranges for every input, the helpers and Configuration
  assert.deepEqual(Object.keys(mock.ss.namedRanges).sort(), ['CFG_AFP', 'CFG_ARS', 'CFG_DEDUCTION_MODE', 'CFG_INCOME_CURRENCY',
    'CFG_ISR', 'CFG_MONTHLY_INCOME', 'CFG_OTHER_CURRENCY', 'CFG_OTHER_INCOME', 'CFG_TAX_RATE', 'DASH_CARDS', 'DASH_MONTH', 'DASH_MONTH_NUM', 'DASH_PERIOD_START',
    'DASH_YEAR', 'RATE_COP', 'RATE_EUR', 'RATE_UPDATED', 'RATE_USD']);
  assert.equal(cell('C4'), 'Current month');
  assert.equal(cell('F4'), new Date().getFullYear());
  assert.ok(dash.hiddenRows.has(3), 'helper row hidden');
  assert.equal(dash.frozenRows, 5);

  // one chart, the conditional-format rules, the LAFISE reference
  assert.equal(dash.charts.length, 1);
  assert.equal(dash.charts[0].cfg.type, 'COLUMN');
  assert.equal(dash.cfRules.length, 11);
  const all = [...dash.cells.values()];
  // "Which card for what": the three programs side by side (v1.1.22)
  const row = name => h.plain(h.get('CASHBACK_MATRIX')).find(r => r[0] === name);
  assert.ok(all.includes('Supermarkets') && all.includes('Taxi apps (Uber, DiDi)') && all.includes('4121'));
  assert.equal(row('Supermarkets')[5], 'BANESCO');
  assert.equal(row('Pharmacies')[5], 'BHD');
  assert.match(row('Restaurants')[2], /10% if ≥ RD\$5,000/);
  assert.ok(all.includes('Pets') && !all.includes('Exclude'), 'custom-only category row, never Exclude');
});

test('Dashboard: first build migrates the v1.1.20 inputs; later builds keep every edit (C7, v1.1.21)', () => {
  const { h, mock } = setup();
  const legacy = mock.ss.insertSheet('Dashboard');
  legacy.getRange('D4').setValue(9);                       // the old hard-coded month
  legacy.getRange('G4').setValue(2025);
  legacy.getRange('G5').setValue(60.1);
  legacy.getRange('I5').setValue(65);
  legacy.getRange('K5').setValue(0.016);
  legacy.getRange('M5').setValue(h.date(2026, 9, 20));
  legacy.getRange('I8:L10').setValues([
    ['LAFISE', 0.1, 'Restaurants, Gas, Groceries', 'Day 9'],   // untouched default → corrected
    ['BANESCO', 0.07, 'Luz y servicios', 'Day 26'],            // user text → kept
    ['BHD', 0.05, 'Salud', 'Day 25']
  ]);

  h.ctx.buildOrRefreshDashboard();
  const named = name => mock.ss.getRangeByName(name);
  assert.equal(named('DASH_MONTH').getValue(), 'Current month', 'old month number is not carried over');
  assert.equal(named('DASH_YEAR').getValue(), 2025);
  assert.equal(named('RATE_USD').getValue(), 60.1);
  assert.equal(named('RATE_EUR').getValue(), 65);
  assert.equal(named('RATE_COP').getValue(), 0.016);
  assert.equal(Object.prototype.toString.call(named('RATE_UPDATED').getValue()), '[object Date]');
  const cards = named('DASH_CARDS').getValues();
  assert.equal(cards[0][4], h.get('CARD_NOTES').LAFISE, 'untouched old default upgraded');
  assert.equal(cards[0][2], '', 'no built-in statement dates (v1.1.24)');
  assert.deepEqual(h.plain(cards[1]), ['BANESCO', 0.07, '', 'Day 26', 'Luz y servicios'], 'note kept, no built-in dates');

  // user edits through the new cells, then rebuilds twice
  named('DASH_MONTH').setValue('March');
  named('RATE_EUR').setValue(66);
  named('RATE_UPDATED').setValue('—');
  const cardsRange = named('DASH_CARDS');
  mock.ss.getSheetByName('Dashboard').getRange(cardsRange.getRow() + 2, 6).setValue('Farmacias 5%');
  mock.ss.getSheetByName('Dashboard').getRange(cardsRange.getRow(), 6).setValue('Restaurants, Gas, Groceries'); // an old default again
  h.ctx.onEdit({ range: named('DASH_YEAR') });
  assert.equal(named('RATE_UPDATED').getValue(), '—', 'editing the year does not stamp the rate date');
  h.ctx.onEdit({ range: named('RATE_EUR') });
  assert.equal(Object.prototype.toString.call(named('RATE_UPDATED').getValue()), '[object Date]', 'rate edit stamps it');

  h.ctx.buildOrRefreshDashboard();
  h.ctx.buildOrRefreshDashboard();
  const dash = mock.ss.getSheetByName('Dashboard');
  assert.equal(named('DASH_MONTH').getValue(), 'March');
  assert.equal(named('RATE_EUR').getValue(), 66);
  assert.equal(named('DASH_CARDS').getValues()[2][4], 'Farmacias 5%', 'your own note is kept');
  assert.equal(named('DASH_CARDS').getValues()[0][4], h.get('CARD_NOTES').LAFISE, 'old default upgraded on the named path too');
  assert.equal(Object.prototype.toString.call(named('RATE_UPDATED').getValue()), '[object Date]');
  assert.equal(dash.charts.length, 1, 'rebuilds never stack charts');
  assert.equal(dash.cfRules.length, 11, 'rebuilds never stack rules');
});

test('run lock: a second run while one is in progress does nothing (C2)', () => {
  const { h, tx } = setup({ lockBusy: true });
  h.ctx.runGmailMonitorForDateRange('2026-01-01', '2026-12-31');
  assert.equal(tx().length, 0);
  assert.ok(h.logs.some(l => l.includes('Another run is in progress')));
});

test('bulk save past the default 1,000 rows and filter growth past 2,000 keep working (E1)', () => {
  const { h, mock, tx } = setup();
  const many = Array.from({ length: 2300 }, (_, i) => ({
    date: h.date(2026, 1, 1 + (i % 28)), bank: 'LAFISE', merchant: 'M' + i, amount: i + 1, currency: 'DOP',
    category: DINING, type: 'Transaction', description: 'M' + i, subject: 's', timestamp: '', messageId: 'bulk_' + i
  }));
  const sheet = mock.ss.getSheetByName('Transactions');
  const first = h.plain(h.ctx.saveTransactions(many.slice(0, 1500)));
  assert.equal(first.success, 1500);
  const criteria = { mock: 'criteria', copy() { return { build: () => ({ mock: 'criteria' }) }; } };
  sheet.getFilter().setColumnFilterCriteria(2, criteria);
  const second = h.plain(h.ctx.saveTransactions(many));
  assert.deepEqual(second, { success: 800, failed: 0, duplicates: 1500, reversalsUnmatched: 0 });
  assert.equal(tx().length, 2300);
  assert.ok(sheet.getFilter().getRange().getLastRow() >= 2301, 'filter grew with the data');
  assert.deepEqual(sheet.getFilter().getColumnFilterCriteria(2), { mock: 'criteria' }, 'user filter kept');
});

test('real thread: purchase + "Reversada" 15 s later → +488 and −488, same merchant and category (v1.1.23)', () => {
  const { h, mock, msg, byName, tx } = setup({ bare: true });
  mock.gmail.threads.push(fakeThread('t-rev', [msg(byName.bhd_consumo_reversal_original), msg(byName.bhd_consumo_reversal)]));
  h.ctx.runGmailMonitorForDateRange('2026-05-06', '2026-05-06');
  const rows = tx();
  assert.equal(rows.length, 2);
  assert.deepEqual(rows.map(r => r[3]).sort((a, b) => a - b), [-488, 488]);
  assert.ok(rows.every(r => r[2] === 'PedidosYa*Som Cafe' && r[5] === DINING), JSON.stringify(rows.map(r => r.slice(2, 7))));
  assert.equal(rows.find(r => r[3] < 0)[6], 'PedidosYa*Som Cafe (reversal)');
  assert.equal(rows.reduce((sum, r) => sum + r[3], 0), 0, 'nets to zero');
  assert.match(h.logs.find(l => l.includes('Email threads found')), /Reversals: 1/);
  assert.ok(!h.logs.some(l => l.includes('Unknown Merchant')), 'no fallback row any more');

  h.ctx.runGmailMonitorForDateRange('2026-05-06', '2026-05-06');
  assert.equal(tx().length, 2, 're-run adds nothing');
});

test('same BANESCO transfer notice twice in one thread → one row, keyed by the bank reference (v1.1.23)', () => {
  const { h, mock, msg, byName, tx } = setup({ bare: true });
  const fx = byName.banesco_transfer_ref;
  const copy = Object.assign({}, fx, { name: fx.name });
  const second = h.fakeMessage({ subject: fx.subject, from: fx.from, body: fixture(fx.name), date: h.date(2026, 9, 5, 12, 22), id: 'id-dup-copy' });
  mock.gmail.threads.push(fakeThread('t-dup', [msg(copy), second]));
  h.ctx.runGmailMonitorForDateRange('2026-09-05', '2026-09-05');
  const rows = tx();
  assert.equal(rows.length, 1);
  assert.equal(rows[0][13], 'BANESCO:E000000.A0000001');
  assert.equal(rows[0][11], 'Transfer');
  assert.match(h.logs.find(l => l.includes('Email threads found')), /Saved: 1 \| Duplicates: 1/);
  assert.equal(mock.ss.getSheetByName('Transactions').getRange(1, 14).getValue(), 'TxRef');
});

test('payroll deposit notice is skipped on purpose, not reported as "could not parse" (v1.1.23)', () => {
  const { h, mock, msg, byName, tx } = setup({ bare: true });
  mock.gmail.threads.push(fakeThread('t-pay', [msg(byName.popular_payroll_deposit)]));
  h.ctx.runGmailMonitorForDateRange('2026-05-07', '2026-05-07');
  const summary = h.logs.find(l => l.includes('Email threads found'));
  assert.equal(tx().length, 0);
  assert.match(summary, /non-transactional 1/);
  assert.doesNotMatch(summary, /Could not parse/);
  assert.ok(mock.gmail.markedRead.includes('t-pay'), 'filtered on purpose → marked processed');
});
