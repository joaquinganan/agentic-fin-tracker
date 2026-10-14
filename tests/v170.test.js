'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices } = require('./sheets_mock');

// v1.1.70: the categories were rebuilt — 25 in 8 groups, Unclassified as the fallback, fixed/variable from one place,
// transfers the tracker can name by itself (fees, loans, broker deposits), and old names moved by Recategorize.

const EMAIL = 'user@example.com';
function book(customRules, opts) {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', EMAIL], ['banksToTrack', '{"BANESCO":true}'], ['monthlyIncome', 1000], ['incomeCurrency', 'USD']]
    .forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  const cr = mock.ss.insertSheet('Custom Rules'); cr.appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  (customRules || []).forEach(r => cr.appendRow([EMAIL, r[0], r[1], '']));
  if (opts && opts.accounts) {
    const acc = mock.ss.insertSheet('Investment Accounts');
    acc.appendRow(['Account', 'Kind', 'Deposit keyword', 'Notes']);
    opts.accounts.forEach(a => acc.appendRow(a));
  }
  return { h, mock, tx: () => mock.ss.getSheetByName('Transactions')._rows(15) };
}
// [date, bank, merchant, amount, currency, category, description, subject, …, type (11), id (12), …, auto (14)]
const row = (h, merchant, category, type, auto) => { const x = new Array(15).fill(''); x[0] = h.date(2026, 10, 10); x[1] = 'BANESCO';
  x[2] = merchant; x[3] = 2500; x[4] = 'DOP'; x[5] = category; x[6] = merchant; x[7] = type === 'Transfer' ? 'Transferencia realizada' : 'Consumo';
  x[9] = 'NO'; x[11] = type; x[12] = 'm-' + merchant.replace(/\W/g, ''); x[14] = auto === undefined ? '' : auto; return x; };
const put = (mock, rows) => mock.ss.getSheetByName('Transactions').getRange(2, 1, rows.length, 15).setValues(rows);

test('the table: 25 categories in 8 groups, each keyword list tried once, in an order of its own', () => {
  const h = load();
  const cats = h.plain(h.get('DEFAULT_CATEGORIES'));
  const groups = h.plain(h.get('CATEGORY_GROUPS'));
  assert.deepEqual(groups, ['Housing', 'Food', 'Mobility', 'Health & care', 'Lifestyle', 'Travel', 'Bank & debt', 'Other']);
  assert.equal(Object.keys(cats).length, 25);
  // display order follows the groups
  const order = Object.keys(cats).map(c => groups.indexOf(cats[c].group));
  assert.ok(order.every((g, i) => g >= 0 && (i === 0 || g >= order[i - 1])), 'grouped, in group order');
  // every category with keywords is tried exactly once; the ones without keywords are never tried
  const match = h.plain(h.get('CATEGORY_MATCH_ORDER'));
  const withKeywords = Object.keys(cats).filter(c => cats[c].keywords.length);
  assert.deepEqual([...match].sort(), [...withKeywords].sort());
  assert.equal(new Set(match).size, match.length);
  assert.deepEqual(Object.keys(cats).filter(c => !cats[c].keywords.length),
    ['Loan payments', 'Payments to people', 'Unclassified'], 'chosen by hand, a Custom Rule, or by the tracker itself (Travel has keywords since v1.1.71)');
  assert.equal(h.get('FALLBACK_CATEGORY'), 'Unclassified');
  assert.deepEqual(h.plain(h.ctx.fixedCategoryNames()),
    ['Rent & housing', 'Utilities', 'Phone & internet', 'Gym', 'Subscriptions', 'Loan payments', 'Insurance']);
});

test('order matters: each pair that used to sit in one category still resolves the same way', () => {
  const h = load();
  const c = t => h.ctx.categorizeTransaction(t);
  const cases = [
    ['BONJOUR TOTAL ARENOSO', 'Restaurants & cafés'],      // not Fuel's 'TOTAL'
    ['TOTAL ARENOSO', 'Fuel & vehicle'],
    ['ALISS METRO PZA', 'Shopping & home'],                // a mall, not the metro
    ['METRO SANTO DOMINGO', 'Public transport & parking'],
    ['UBER EATS-W*UBER EATS- SANTO DOMINGO DOM', 'Delivery'],
    ['UBER*TRIP SANTO DOMINGO', 'Ride apps & taxis'],
    ['DIDI FOOD BOGOTA', 'Delivery'],
    ['DiDi CO Ride', 'Ride apps & taxis'],
    ['AMAZON PRIME*2K3', 'Subscriptions'],
    ['AMAZON MKTPL*AB12', 'Shopping & home'],
    ['IKEA REST SANTO DOMINGO', 'Restaurants & cafés'],
    ['IKEA SANTO DOMINGO', 'Shopping & home'],
    ['CLINICA VETERINARIA EJEMPLO', 'Pets'],               // a vet, not a clinic
    ['CLINICA EJEMPLO', 'Health & pharmacy'],
    ['PUMA STORE BLUE MALL', 'Shopping & home'],           // not the PUMA gas station
    ['PUMA LAS AMERICAS', 'Fuel & vehicle'],
    ['BARBERIA EJEMPLO', 'Personal care'],
    ['SUPERMERCADO BRAVO', 'Groceries']
  ];
  for (const [text, expected] of cases) assert.equal(c(text), expected, text);
});

test('agreed decisions: PayPal by its merchant or Unclassified; Cash its own; repairs, laundry and gear not fixed', () => {
  const h = load();
  const c = t => h.ctx.categorizeTransaction(t);
  assert.equal(c('PAYPAL *IKEA'), 'Shopping & home', 'PAYPAL *<merchant>: that merchant\'s category');
  assert.equal(c('PAYPAL *SPOTIFY'), 'Subscriptions');
  assert.equal(c('PAYPAL *UBERBV'), 'Ride apps & taxis');
  assert.equal(c('PAYPAL *XK29PL'), 'Unclassified', 'PayPal alone says nothing of what was bought');
  assert.equal(c('PAYPAL'), 'Unclassified');
  assert.equal(c('Código Cash (cash withdrawal)'), 'Cash');
  assert.equal(c('Retiro TuEfectivo (cash withdrawal)'), 'Cash');
  for (const m of ['FERRETERIA EJEMPLO', 'LAVANDERIA EJEMPLO', 'TIENDA DE DEPORTES EJEMPLO']) {
    assert.equal(c(m), 'Shopping & home', m);
    assert.equal(h.ctx.isFixedCategory(c(m)), false, m + ' is not a bill');
  }
  assert.equal(c('SEGUROS EJEMPLO'), 'Insurance');
  assert.equal(h.ctx.isFixedCategory('Insurance'), true);
  assert.equal(c('CAASD FACTURA'), 'Utilities', 'water is a utility');
  assert.equal(c('ALGO QUE NADIE CONOCE'), 'Unclassified');
  assert.equal(c(''), 'Unclassified');
});

test('transfers the tracker names by itself: tax rows, loan installments, deposits to a broker', () => {
  const h = load();
  const a = (t, k) => h.ctx.autoTransferCategory(t, k);
  assert.equal(a('BDI — impuesto y comisión de transferencia'), 'Bank fees & taxes');
  assert.equal(a('Banreservas: impuesto y comisión de pago de préstamo'), 'Bank fees & taxes');
  assert.equal(a('Pago de préstamo ***3333'), 'Loan payments');
  assert.equal(a('Pago de préstamo *3333'), 'Loan payments');
  assert.equal(a('JUAN PEREZ'), '');
  assert.equal(a('OUROSR CAPITAL', ['OUROSR']), 'Investments & savings');
  assert.equal(a('OUROSR CAPITAL', []), '', 'only with an account\'s deposit keyword');
  // through Recategorize: a broker deposit wins over a Custom Rule (it is never spending); the others only fill a blank
  const TX = h.get('TX_COL');
  const mk = (merchant, type) => { const r = new Array(15).fill(''); r[TX.MERCHANT] = merchant; r[TX.DESCRIPTION] = merchant;
    r[TX.TYPE] = type; r[TX.CURRENCY] = 'DOP'; r[TX.SUBJECT] = 'Recibo de la transacción'; return r; };
  const rules = { Exclude: ['OUROSR'], 'Payments to people': ['PAGO DE PRÉSTAMO'] };
  assert.equal(h.ctx.computeRecategorization(mk('OUROSR CAPITAL', 'Transfer'), rules, ['OUROSR']).category, 'Investments & savings');
  assert.equal(h.ctx.computeRecategorization(mk('Pago de préstamo *3333', 'Transfer'), rules, []).category, 'Payments to people',
    'your Custom Rule beats the automatic one');
  assert.equal(h.ctx.computeRecategorization(mk('Pago de préstamo *3333', 'Transfer'), {}, []).category, 'Loan payments');
  assert.equal(h.ctx.computeRecategorization(mk('Pago de préstamo *3333', 'Transaction'), {}, []).category, 'Unclassified',
    'only transfers: a purchase is categorized by its keywords');
  // the deposit keywords come from Investment Accounts — before that sheet exists, the defaults (HAPI)
  assert.deepEqual(h.plain(book().h.ctx.investmentDepositKeywords()), ['OUROSR']);
  assert.deepEqual(h.plain(book([], { accounts: [['HAPI', 'Broker', 'OUROSR', ''], ['FUND', 'Fund', '', ''], ['IBKR', 'Broker', 'IBKR LLC', '']] })
    .h.ctx.investmentDepositKeywords()), ['OUROSR', 'IBKR LLC']);
});

test('Investments & savings is never spending: summaries, Dashboard rows and categories typed by hand', () => {
  const h = load();
  assert.equal(h.ctx.isSpendingCategory('Investments & savings'), false);
  assert.equal(h.ctx.isSpendingCategory(' exclude '), false);
  assert.equal(h.ctx.isSpendingCategory(''), false);
  assert.equal(h.ctx.isSpendingCategory('Unclassified'), true);
  const { h: hb, mock } = book();
  put(mock, [row(hb, 'OUROSR CAPITAL', 'Investments & savings', 'Transfer'), row(hb, 'MI CATEGORIA', 'Viajes', 'Transaction')]);
  assert.deepEqual(hb.plain(hb.ctx.categoriesOnlyInTransactions(hb.plain(hb.ctx.getCategories()))), ['Viajes']);
  // the daily summary leaves it out of what was spent
  const TX = h.get('TX_COL');
  const r = (cat, amount) => { const x = new Array(15).fill(''); x[TX.DATE] = h.date(2026, 10, 13); x[TX.AMOUNT] = amount; x[TX.CURRENCY] = 'DOP';
    x[TX.CATEGORY] = cat; x[TX.TYPE] = 'Transfer'; x[TX.MERCHANT] = 'M'; return x; };
  const s = h.plain(h.ctx.computeDailySummary([new Array(15).fill('h'), r('Investments & savings', 50000), r('Rent & housing', 30000)],
    { today: h.date(2026, 10, 14, 9), rates: { USD: 60, EUR: 65, COP: 0.015 }, netIncomeDop: 100000 }));
  assert.equal(s.spent, 30000);
});

test('fixed or variable: one list for the Dashboard, the daily and the monthly summary', () => {
  const h = load();
  for (const c of h.plain(h.ctx.getCategories())) {
    assert.equal(h.ctx.isBillCategory(c), h.ctx.isFixedCategory(c), c);
  }
  assert.equal(h.ctx.isBillCategory('Utilities'), true, 'electricity was a bill in the summaries but variable on the Dashboard');
  assert.equal(h.get('SUMMARY_SUBSCRIPTIONS_CATEGORY'), 'Subscriptions');
});

test('Recategorize moves old category names: one-to-one as they are, merged ones within their family', () => {
  const { h, mock, tx } = book([['Rent', 'LANDLORD'], ['Groceries + Barbershop', 'CORTE'], ['Electricity', 'PLANTA']],
    { accounts: [['HAPI', 'Broker', 'OUROSR', '']] });
  put(mock, [
    row(h, 'JUAN PEREZ', 'Rent', 'Transfer'),                                             // set by hand (before v1.1.39)
    row(h, 'BARBERIA EJEMPLO', 'Groceries + Barbershop', 'Transaction', 'Dining/Delivery + Entertainment + Other'),   // set by hand
    row(h, 'TIENDA NUEVA', 'Health + Vet + Pharmacy', 'Transaction', 'Dining/Delivery + Entertainment + Other'),     // set by hand
    row(h, 'ALGUIEN', 'Dining/Delivery + Entertainment + Other', 'Transfer', '(none)'),   // set by hand on a transfer
    row(h, 'UBER*RIDES', 'Transportation', 'Transaction', 'Transportation'),              // the tracker's own: recomputed
    row(h, 'CACHAREPA CHURCHILL', 'Dining/Delivery + Entertainment + Other', 'Transaction', 'Dining/Delivery + Entertainment + Other'),
    row(h, 'COMERCIO DESCONOCIDO', 'Dining/Delivery + Entertainment + Other', 'Transaction', 'Dining/Delivery + Entertainment + Other'),
    row(h, 'OUROSR CAPITAL', 'Exclude', 'Transfer', '(none)'),                             // your Exclude on a broker deposit
    row(h, 'MARIA GOMEZ', 'Exclude', 'Transfer', '(none)')]);                              // your Exclude elsewhere: kept
  h.ctx.recategorizeAllTransactions(EMAIL);
  const by = Object.fromEntries(tx().map(r => [r[2], r[5]]));
  assert.equal(by['JUAN PEREZ'], 'Rent & housing');
  assert.equal(by['BARBERIA EJEMPLO'], 'Personal care', 'the keywords found one of its own family');
  assert.equal(by['TIENDA NUEVA'], 'Health & pharmacy', 'no keyword in its family: the main one');
  assert.equal(by['ALGUIEN'], 'Unclassified');
  assert.equal(by['UBER*RIDES'], 'Ride apps & taxis');
  assert.equal(by['CACHAREPA CHURCHILL'], 'Restaurants & cafés', 'a keyword of its family (v1.1.71)');
  assert.equal(by['COMERCIO DESCONOCIDO'], 'Unclassified', 'no keyword knows it — visible now');
  assert.equal(by['OUROSR CAPITAL'], 'Investments & savings');
  assert.equal(by['MARIA GOMEZ'], 'Exclude');
  // Custom Rules: one-to-one names renamed in the sheet, a merged one left for you to choose
  const rules = mock.ss.getSheetByName('Custom Rules')._rows(4).map(r => r[1]);
  assert.deepEqual(rules, ['Rent & housing', 'Groceries + Barbershop', 'Utilities']);
  // a second pass changes nothing, and what you set by hand stays yours
  const before = JSON.stringify(tx());
  assert.equal(h.ctx.recategorizeAllTransactions(EMAIL), 0);
  assert.equal(JSON.stringify(tx()), before);
  // Start here asks about the merged one
  const steps = Object.fromEntries(h.plain(h.ctx.startHereStatus(h.date(2026, 10, 14, 9))).steps.map(s => [s.id, s]));
  assert.equal(steps.oldCategories.status, 'warn');
  assert.match(steps.oldCategories.detail, /"Groceries \+ Barbershop" → Groceries or Personal care/);
});

test('legacyCategoryFor: every old name has a new one that exists', () => {
  const h = load();
  const cats = h.plain(h.ctx.getCategories());
  const legacy = h.plain(h.get('LEGACY_CATEGORIES'));
  assert.equal(Object.keys(legacy).length, 11);
  for (const [old, m] of Object.entries(legacy)) {
    assert.ok(cats.includes(m.to), old);
    for (const w of m.within || []) assert.ok(cats.includes(w), old + ' → ' + w);
    assert.ok(!cats.includes(old), old + ' is not a name any more');
  }
  assert.equal(h.ctx.legacyCategoryFor('Transportation', 'Fuel & vehicle'), 'Fuel & vehicle');
  assert.equal(h.ctx.legacyCategoryFor('Transportation', 'Groceries'), 'Ride apps & taxis');
  assert.equal(h.ctx.legacyCategoryFor('Rent', 'Groceries'), 'Rent & housing', 'one-to-one: always its new name');
  assert.equal(h.ctx.legacyCategoryFor('Viajes', 'Groceries'), '', 'not an old name');
});

test('Dashboard: groups with their subtotal, categories under them, and what is not spending on its own', () => {
  const { h, mock } = book();
  put(mock, [row(h, 'SUPERMERCADO BRAVO', 'Groceries', 'Transaction')]);
  h.ctx.buildOrRefreshDashboard();
  const dash = mock.ss.getSheetByName('Dashboard');
  const value = (r, c) => dash.getRange(r, c).getValue();
  // the grid: FOOD adds the three food categories right below it
  let food = null;
  for (let r = 1; r <= 200; r++) if (value(r, 2) === 'FOOD') food = r;
  assert.ok(food);
  assert.deepEqual([value(food + 1, 2), value(food + 2, 2), value(food + 3, 2)], ['Groceries', 'Restaurants & cafés', 'Delivery']);
  assert.equal(value(food, 3), `=SUM(C${food + 1}:C${food + 3})`);
  // TOTAL adds the group rows and the transfers, so no category is counted twice
  let total = null;
  for (let r = 1; r <= 200; r++) if (value(r, 2) === 'TOTAL') total = r;
  const groups = []; for (let r = 1; r < total; r++) if (/^[A-Z &]+$/.test(String(value(r, 2))) && String(value(r, 3)).startsWith('=SUM(C')) groups.push(r);
  assert.equal(groups.length, 8);
  assert.equal(value(total, 3), '=SUM(' + groups.map(r => 'C' + r).concat(['C' + (total - 1)]).join(',') + ')');
  // the heat map colours categories only
  const heat = dash.cfRules.find(r => r.setGradientMaxpoint);
  assert.equal(heat.ranges.length, 8);
  assert.ok(heat.ranges.every(rg => !groups.includes(rg.getRow())));
  // "Not spending": card payments, your own accounts, investments, money received, purchases left out
  const labels = [...dash.cells.entries()].filter(([k, v]) => typeof v === 'string' && /^(💳|🔁|📈|📥|🚫) [A-Z][a-z]/.test(v));
  assert.deepEqual(labels.map(([, v]) => v), ['💳 Credit card payments', '🔁 Between your accounts', '📈 Investments & savings',
    '📥 Money received', '🚫 Purchases left out']);
  const formulaRight = ([k]) => dash.getRange(Number(k.split(',')[0]), 12).getValue();
  assert.match(formulaRight(labels[0]), /^=ABS\(SUMIFS\(Transactions!\$D:\$D,Transactions!\$L:\$L,"Card Payment",/);
  assert.match(formulaRight(labels[2]), /Transactions!\$F:\$F,"Investments & savings"/);
  // by bank leaves investments out too
  assert.ok([...dash.cells.values()].some(v => typeof v === 'string' && v.includes('Transactions!$F:$F,"<>Investments & savings"')));
});

test('Categories sheet: each category with its group and type; Exclude and Investments & savings listed', () => {
  const { h, mock } = book([['Kids', 'TOYSRUS']]);
  h.ctx.buildOrRefreshCategoriesSheet(EMAIL);
  const rows = mock.ss.getSheetByName('Categories')._rows(5).slice(3);
  const by = Object.fromEntries(rows.filter(r => r[1]).map(r => [r[1], r]));
  assert.deepEqual(by['Utilities'].slice(2, 4), ['Housing', 'Fixed']);
  assert.deepEqual(by['Delivery'].slice(2, 4), ['Food', 'Variable']);
  assert.match(by['Payments to people'][4], /chosen by hand or with a Custom Rule/);
  assert.match(by['Travel'][4], /ARAJET/, 'Travel has keywords since v1.1.71');
  assert.match(by['Unclassified'][4], /Anything no keyword/);
  assert.deepEqual(by['Investments & savings'].slice(2, 4), ['—', 'Not spending']);
  assert.deepEqual(by['Kids'].slice(2, 4), ['Other', 'Variable']);
});

// ---- BDI Digital: a transfer to another BDI account (a sample shared on Oct 14, 2026; the fixture keeps its structure)
const BDI = { subject: 'BDI Digital: Comprobante de Transacción', from: '"bdinforma@bdi.com.do" <bdinforma@bdi.com.do>' };
const parseBdi = (h, body, rules) => h.plain(h.ctx.parseEmailMessage(h.fakeMessage({ subject: BDI.subject, from: BDI.from, body: body,
  date: h.date(2026, 10, 5, 10), id: 'm' }), rules || {}, h.ctx.newParseStats()));
const rowOf = t => [t.bank, t.type, t.currency, t.amount, t.merchant, t.category];

test('BDI Digital transfer between BDI accounts: what left the account, a Transfer to review; your name makes it Exclude', () => {
  const h = load();
  const r = parseBdi(h, fixture('bdi_digital_transferencia'));
  assert.equal(r.status, 'ok', 'it was Unrecognized: "Format not known yet"');
  assert.deepEqual(r.items.map(rowOf), [['BDI', 'Transfer', 'DOP', 3050, 'RAMIREZ PEREZ, LUIS ALBERTO', '']]);
  assert.match(h.plain(h.ctx.extractBDIInternalTransferTransactions(fixture('bdi_digital_transferencia')))[0].context, /to BDI account \*\*\*0001 \(US\$50\.00 at 61\.00\)/);
  // a Custom Rule on your own name: a move between your accounts
  const own = parseBdi(h, fixture('bdi_digital_transferencia'), { Exclude: ['RAMIREZ PEREZ'] });
  assert.equal(own.items[0].category, 'Exclude');
  // tax and commission, each in its own currency, as Bank fees & taxes
  const fees = parseBdi(h, fixture('bdi_digital_transferencia').replace('Impuesto 0.20%\nRD$0.00', 'Impuesto 0.20%\nRD$6.10')
    .replace('*US$0.00*', '*US$1.00*'));
  assert.deepEqual(fees.items.map(rowOf), [['BDI', 'Transfer', 'DOP', 3050, 'RAMIREZ PEREZ, LUIS ALBERTO', ''],
    ['BDI', 'Transfer', 'DOP', 6.1, 'BDI — impuesto y comisión de transferencia', 'Bank fees & taxes'],
    ['BDI', 'Transfer', 'USD', 1, 'BDI — impuesto y comisión de transferencia', 'Bank fees & taxes']]);
  // the same currency on both sides: no exchange in the context
  assert.equal(h.plain(h.ctx.extractBDIInternalTransferTransactions(fixture('bdi_digital_transferencia').replace('*US$50.00*', '*RD$3,050.00*')))[0].context,
    'to BDI account ***0001');
  // the interbank transfer and the card payment read as before
  assert.deepEqual(h.plain(h.ctx.extractBDIAnyTransferTransactions(fixture('bdi_transferencia_interbancaria'))).map(x => [x.amount, x.merchant]),
    [[2500, 'Luis Ramirez'], [5, 'BDI — impuesto y comisión de transferencia']]);
  assert.deepEqual(h.plain(h.ctx.extractBDIInternalTransferTransactions(fixture('bdi_transferencia_interbancaria'))), []);
  assert.deepEqual(h.plain(h.ctx.extractBDIInternalTransferTransactions(fixture('bdi_online_pago_tarjeta'))), []);
  // still strict: a receipt of this kind without its amounts is not guessed
  assert.equal(parseBdi(h, 'Comprobante de Transacción\nTipo de Transacción\n*Transferencias*\nBeneficiario\nLUIS\n4100000001').status, 'failed');
});
