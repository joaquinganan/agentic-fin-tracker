'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices } = require('./sheets_mock');

// A POPULAR "Código Cash" withdrawal: the consumo table without a merchant column (synthetic; format from a live
// Raw_POPULAR row whose merchant was saved as "Aprobada").
const CODIGO_CASH = 'Notificación de retiro Código Cash\nMonto\tMoneda\tFecha\tEstatus\nRD$400.00\tRD\t06/05/2026\tAprobada\nGracias por preferirnos\n';

test('Código Cash: the status word is never saved as the merchant; category Dining/Other', () => {
  const h = load();
  const msg = h.fakeMessage({ subject: 'Notificación de retiro Código Cash', from: 'notificaciones@popularenlinea.com',
    body: CODIGO_CASH, date: h.date(2026, 5, 6, 10), id: 'cc1' });
  const r = h.plain(h.ctx.parseEmailMessage(msg, {}, h.ctx.newParseStats()));
  assert.equal(r.items.length, 1);
  assert.deepEqual([r.items[0].merchant, r.items[0].amount, r.items[0].category],
    ['Código Cash (cash withdrawal)', 400, 'Dining/Delivery + Entertainment + Other']);
  assert.equal(h.ctx.fixStatusAsMerchant('Aprobada', 'Alerta de Consumo', ''), h.get('GARBLED_PLACEHOLDER'), 'any other status-only row is flagged, not invented');
  assert.equal(h.ctx.fixStatusAsMerchant('APROBADO FOODS SRL', 'Notificación de retiro Código Cash', ''), 'APROBADO FOODS SRL', 'only an exact status word');
});

test('Código Cash: rows already saved as "Aprobada" are repaired by recategorize; real merchants untouched', () => {
  const mock = makeServices();
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"POPULAR":true}']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions');
  h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  const row = (merchant, subject) => { const x = new Array(14).fill(''); x[0] = h.date(2026, 5, 6); x[1] = 'POPULAR'; x[2] = merchant;
    x[3] = 400; x[4] = 'DOP'; x[5] = 'Groceries + Barbershop'; x[6] = merchant; x[7] = subject; x[9] = 'NO'; x[11] = 'Transaction'; x[12] = 'm-' + merchant; return x; };
  mock.ss.getSheetByName('Transactions').getRange(2, 1, 2, 14).setValues([row('Aprobada', 'Notificación de retiro Código Cash'),
    row('SUPERMERCADO NACIONAL', 'Notificación de Consumo')]);
  h.ctx.recategorizeAllTransactions();
  const rows = mock.ss.getSheetByName('Transactions')._rows(14);
  const cc = rows.find(r => r[7] === 'Notificación de retiro Código Cash');
  assert.deepEqual([cc[2], cc[6], cc[5]], ['Código Cash (cash withdrawal)', 'Código Cash (cash withdrawal)', 'Dining/Delivery + Entertainment + Other']);
  assert.equal(rows.find(r => r[7] === 'Notificación de Consumo')[2], 'SUPERMERCADO NACIONAL');
});

function holdingsBook() {
  const mock = makeServices();
  const h = load({ services: mock.services });
  const led = h.ctx.getOrCreateLedgerSheet();
  led.getRange(2, 1, 5, 12).setValues([
    [h.date(2026, 9, 1), 'HAPI', 'Snapshot', 'GOOGL', 2, 300, 500, '', 'USD', 'manual', '', ''],
    [h.date(2026, 9, 1), 'HAPI', 'Snapshot', 'ETHUSD', 0.5, 2600, 1100, '', 'USD', 'manual', '', ''],
    [h.date(2026, 9, 1), 'HAPI', 'Snapshot', 'CASH', '', '', 10, '', 'USD', 'manual', '', ''],
    [h.date(2026, 9, 27), 'Liquidity fund', 'Valuation', '', 50, 2277.6, '', '', 'DOP', 'manual', '', ''],   // units × price
    [h.date(2026, 9, 8), 'Pension fund', 'Valuation', '', '', '', 800000, '', 'DOP', 'manual', '', '']]);     // a balance
  return { h, mock };
}

test('Holdings: dates are real dates with one format — units and balance entries alike (reported)', () => {
  const { h, mock } = holdingsBook();
  h.ctx.refreshHoldings();
  const sheet = mock.ss.getSheetByName('Holdings');
  const rows = sheet._rows(8);
  const fund = rows.find(r => r[1] === 'Liquidity fund'), pension = rows.find(r => r[1] === 'Pension fund');
  for (const r of [fund, pension]) assert.equal(typeof r[4].getFullYear, 'function', r[1] + ': As of is a Date, not text');
  const fundRow = rows.indexOf(fund) + 2;
  assert.ok(sheet.styles.some(s => s.m === 'setNumberFormat' && s.v === 'yyyy-mm-dd' && s.rect.c1 === 5 && s.rect.r1 <= fundRow && s.rect.r2 >= fundRow));
  const perf = rows.find(r => r[1] === 'All accounts');
  assert.equal(typeof perf[2].getFullYear, 'function', 'Tracked since is a Date too');
});

test('Holdings: the Dashboard look — title band, KPI cards, sections, chips, two charts; the return KPI points at Performance', () => {
  const { h, mock } = holdingsBook();
  h.ctx.refreshHoldings();
  const sheet = mock.ss.getSheetByName('Holdings');
  const values = [...sheet.cells.values()].map(String);
  for (const v of ['📈  Investments', 'PORTFOLIO VALUE', 'UNREALIZED P/L', 'RETURN SINCE START', '📊  Positions', '🏦  Other accounts',
                   'TOTAL INVESTED', '📐  Performance', 'All accounts']) assert.ok(values.includes(v), v);
  assert.ok(sheet.styles.some(s => s.m === 'setBackground' && s.v === '#1F3864' && s.rect.r1 === 1 && s.rect.r2 === 2), 'navy title band');
  assert.ok(sheet.merges.length >= 6, 'title, subtitle and KPI cards are merged');
  const kpiReturn = sheet.getRange(5, 8).getValue();
  const allRow = sheet._rows(8).findIndex(r => r[1] === 'All accounts') + 2;
  assert.equal(kpiReturn, '=H' + allRow, 'the RETURN SINCE START card shows the all-accounts return');
  const rules = sheet.cfRules;
  assert.ok(rules.some(r => r.whenTextEqualTo === 'live') && rules.some(r => r.whenTextStartsWith === 'Coinbase') &&
    rules.some(r => r.whenTextStartsWith === 'last known'), 'price-source chips');
  assert.ok(rules.some(r => r.setGradientMaxpoint), 'weight colour scale');
  assert.deepEqual(sheet.charts.map(c => c.cfg.type).sort(), ['PIE'], 'one day of history: allocation pie only');
});

test('the other investment tabs are styled: ledger chips, history day bands and TOTAL rows, account kinds', () => {
  const { h, mock } = holdingsBook();
  h.ctx.refreshHoldings();
  h.ctx.styleTrackerSheets();
  const ledger = mock.ss.getSheetByName('Investment Ledger');
  assert.ok(ledger.cfRules.some(r => r.whenTextEqualTo === 'Snapshot') && ledger.cfRules.some(r => r.whenTextEqualTo === 'manual'), 'type and source chips');
  const history = mock.ss.getSheetByName('Portfolio History');
  assert.ok(history.cfRules.some(r => r.whenFormulaSatisfied === '=$B2="TOTAL"' && r.setBold === true));
  assert.ok(history.cfRules.some(r => r.whenFormulaSatisfied === '=AND($A2<>"",ISEVEN(INT($A2)))'), 'bands alternate by day');
  mock.ss.insertSheet('Investment Accounts').getRange(1, 1, 2, 4).setValues([['Account', 'Kind', 'Deposit keyword', 'Notes'], ['HAPI', 'Broker', 'OUROSR', '']]);
  h.ctx.styleTrackerSheets();
  assert.ok(mock.ss.getSheetByName('Investment Accounts').cfRules.some(r => r.whenTextEqualTo === 'Pension'), 'kind chips');
  assert.deepEqual(h.logs.filter(l => /Could not style/.test(l)), []);
});
