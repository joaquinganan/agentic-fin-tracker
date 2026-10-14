'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

// v1.1.71: POPULAR purchases with the merchant on the line after the date (below), and merchants that stayed in Unclassified after v1.1.70, reviewed by hand and given to every user as keywords
// (public businesses only). Descriptors keep the shape banks print (city, country, cut-short names); branch names and
// numbers are made up.

test('reviewed Unclassified merchants land in their category', () => {
  const h = load();
  const cases = [
    ['LUSERGAS LOS PRADOS SANTO DOMINGO DOM', 'Rent & housing'],
    ['EJEMPLO FITNES DISTRITO NACI DOM', 'Gym'],
    ['LINKEDINPREA*00000000 DUBLIN 2 IRL', 'Subscriptions'],
    ['PORKBUN.COM SHERWOOD USA', 'Subscriptions'],
    ['VOCAL STAR', 'Subscriptions'],
    ['SURTIDORA EJEMPLO DISTRITO NACI DOM', 'Groceries'],
    ['ALM UNIDOS', 'Groceries'],
    ['ALMACENES', 'Groceries'],
    ['EXITO JUNIN', 'Groceries'],
    ['MARKET PASTEUR EJEMPLO', 'Groceries'],
    ['COLAITO PIANTINI SANTO DOMINGO DOM', 'Restaurants & cafés'],
    ['GOYURT CENTRO SD SANTO DOMINGO DOM', 'Restaurants & cafés'],
    ['IN&OUT EJEMPLO SANTO DOMINGO DOM', 'Restaurants & cafés'],
    ["WENDY'S EJEMPLO SANTO DOMINGO DOM", 'Restaurants & cafés'],
    ['MC DONALDS EJEMPLO', 'Restaurants & cafés'],
    ['CASA DE EJEMPLO PIZZER SANTO DOMINGO DOM', 'Restaurants & cafés'],   // "PIZZERIA" cut short
    ['TST*EJEMPLO GRILL -', 'Restaurants & cafés'],                         // Toast, a restaurant-only terminal
    ['BOLD*EJEMPLO CARNES PARR', 'Restaurants & cafés'],
    ['HAAGEN DAZS EJEMPLO', 'Restaurants & cafés'],
    ['BASKIN ROBBINS EJEMPLO', 'Restaurants & cafés'],
    ['EJEMPLO DESSERTS', 'Restaurants & cafés'],
    ['EJEMPLO CAFFE', 'Restaurants & cafés'],
    ['EJEMPLO COFFEE ROASTER', 'Restaurants & cafés'],
    ['EJEMPLO AREPAS', 'Restaurants & cafés'],
    ['CACHAREPA CHURCHILL', 'Restaurants & cafés'],
    ['GRANIER EJEMPLO', 'Restaurants & cafés'],
    ['EJEMPLO PARKIN SANTO DOMINGO DOM', 'Public transport & parking'],   // "PARKING" cut short
    ['RDVIAL', 'Public transport & parking'],
    ['LASA EJEMPLO', 'Public transport & parking'],
    ['APTPRA EJEMPLO', 'Ride apps & taxis'],
    ['GRUPO COMETA', 'Fuel & vehicle'],
    ['DROGUERIA EJEMPLO 000', 'Health & pharmacy'],
    ['CLINICORP', 'Health & pharmacy'],
    ['BANCO POPULAR CAROL AV. I', 'Health & pharmacy'],
    ['ARENA DE GATOS', 'Pets'],
    ['MINISO EJEMPLO', 'Shopping & home'],
    ['DOLLARCITY', 'Shopping & home'],
    ['BERSHKA EJEMPLO', 'Shopping & home'],
    ['CECOMSA ECOMMPP SANTIAGO DOM', 'Shopping & home'],
    ['DOWNTOWN CENTER I SANTO DOMINGO DOM', 'Shopping & home'],
    ['PARTY FIESTA', 'Shopping & home'],
    ['SMART TICKET SANTO DOMINGO DOM', 'Leisure & going out'],
    ['DAVE BUSTERS', 'Leisure & going out'],
    ['PAYPAL *ENEBA', 'Leisure & going out'],                               // PayPal + a merchant: the merchant's
    ['Driffle VILNIAUS M LTU', 'Leisure & going out'],
    ['EldoradoGG', 'Leisure & going out'],
    ['PROCURADURIA - 000000 SANTO DOMINGO DOM', 'Government & procedures'],
    ['ARAJET, S.A.', 'Travel'],
    ['JETBLUE', 'Travel'],
    ['FORT LAUDERDALE AIRPOR', 'Travel'],
  ];
  for (const [text, expected] of cases) assert.equal(h.ctx.categorizeTransaction(text), expected, text);
});

test('Travel is tried after what a trip is made of: a meal or parking at an airport keeps its own category', () => {
  const h = load();
  const c = t => h.ctx.categorizeTransaction(t);
  assert.equal(c('AIRPORT PARKING'), 'Public transport & parking');
  assert.equal(c('STARBUCKS MIAMI AIRPORT'), 'Restaurants & cafés');
  assert.equal(c('MINISO AIRPORT'), 'Shopping & home');
  assert.equal(c('MIAMI INTL AIRPORT'), 'Travel');
  const order = h.plain(h.get('CATEGORY_MATCH_ORDER'));
  for (const before of ['Restaurants & cafés', 'Shopping & home', 'Public transport & parking']) {
    assert.ok(order.indexOf(before) < order.indexOf('Travel'), before + ' before Travel');
  }
});

test('the new keywords do not catch what they should not', () => {
  const h = load();
  const c = t => h.ctx.categorizeTransaction(t);
  assert.equal(c('TRANSFERENCIA EXITOSA'), 'Unclassified', 'no bare "EXITO": it is inside "EXITOSA"');
  assert.equal(c('CAROL MARTINEZ'), 'Unclassified', "'CAROL' alone is not a pharmacy");
  assert.equal(c('LASAGNA HOUSE'), 'Unclassified', "'LASA' is a whole word");
  assert.equal(c('TSTAR MOTORS'), 'Unclassified', "'TST' is a whole word");
  assert.equal(c('PAYPAL'), 'Unclassified', 'bare PayPal stays to review (v1.1.70)');
  // "UBR*" is how the bank prints an Uber ride's estimate (v1.1.67, real alert) — a ride, not a delivery
  assert.equal(c('UBR* PENDING.UBER.COM'), 'Ride apps & taxis');
});

test('a Custom Rule still wins over the new keywords', () => {
  const h = load();
  const COL = h.get('TX_COL');
  const r = new Array(15).fill('');
  r[COL.SUBJECT] = 'Notificación de Consumo'; r[COL.MERCHANT] = 'ARAJET, S.A.'; r[COL.DESCRIPTION] = 'ARAJET, S.A.';
  r[COL.TYPE] = 'Transaction'; r[COL.CURRENCY] = 'DOP'; r[COL.IS_CREDIT] = 'NO';
  assert.equal(h.plain(h.ctx.computeRecategorization(r, {})).category, 'Travel');
  assert.equal(h.plain(h.ctx.computeRecategorization(r, { 'Work trips': ['ARAJET'] })).category, 'Work trips');
});

// ---- POPULAR "Notificación de Consumo" with the merchant on the next line (real .eml structure, data made up)
const POP = { subject: 'Notificación de Consumo', from: '<notificaciones@popularenlinea.com>' };
const parsePopular = (h, body) => h.plain(h.ctx.parseEmailMessage(h.fakeMessage({ subject: POP.subject, from: POP.from, body: body,
  date: h.date(2026, 3, 15, 16), id: 'p1' }), {}, h.ctx.newParseStats()));
const items = r => (r.items || r).map(i => [i.merchant, i.amount, i.currency, i.type]);

test('POPULAR: the merchant on the line after the date is read (it was the text after "terminada en")', () => {
  const h = load();
  const text = fixture('popular_consumo_wrapped');
  assert.match(text, /15\/03\/2026\nPedidosYa\*Pizza Ejemplo Sa \tAprobada/, 'the fixture keeps the real line break');
  assert.deepEqual(items(parsePopular(h, text)), [['PedidosYa*Pizza Ejemplo Sa', 1250, 'DOP', 'Transaction']]);
  // the same email as its HTML reads (one cell per line, the status in bold) and flattened to one line
  const html = text.replace('Monto \tMoneda \tFecha \tComercio\tEstatus\t\nRD$1,250.00 \tPeso dominicano \t15/03/2026\nPedidosYa*Pizza Ejemplo Sa \tAprobada\t',
    'Monto\nMoneda\nFecha\nComercio\nEstatus\nRD$1,250.00\nPeso dominicano\n15/03/2026\nPedidosYa*Pizza Ejemplo Sa\n*Aprobada*');
  assert.notEqual(html, text);
  assert.deepEqual(items(parsePopular(h, html)), [['PedidosYa*Pizza Ejemplo Sa', 1250, 'DOP', 'Transaction']]);
  assert.deepEqual(items(parsePopular(h, text.replace(/\s+/g, ' '))), [['PedidosYa*Pizza Ejemplo Sa', 1250, 'DOP', 'Transaction']]);
});

test('POPULAR: a merchant cut over two lines is one name; declined rows flagged; the one-line form still read', () => {
  const h = load();
  const x = h.ctx.extractPOPULARTransactions;
  assert.deepEqual(h.plain(x('US$14.27 \tDólar estadounidense \t22/05/2026\nAMAZON MKTPLACE\nPMTS \tAprobada\t')).map(i => [i.merchant, i.currency, i.declined]),
    [['AMAZON MKTPLACE PMTS', 'USD', false]]);
  assert.deepEqual(h.plain(x('RD$500.00 \tPeso dominicano \t01/04/2026\nTIENDA EJEMPLO \tRechazada\t')).map(i => [i.merchant, i.declined]),
    [['TIENDA EJEMPLO', true]]);
  assert.deepEqual(h.plain(x(fixture('popular_consumo_casita'))).map(i => i.merchant), ['LA CASITA DE'], 'the earlier sample');
  // a status word outside the list: the one-line form, as before
  assert.deepEqual(h.plain(x('RD$75.00\tRD\t02/04/2026\tCOMERCIO UNO\tProcesada')).map(i => [i.merchant, i.amount]), [['COMERCIO UNO', 75]]);
});

test('POPULAR rows saved before with the text as merchant: Start here says so, and a deleted row comes back right', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 1000], ['incomeCurrency', 'USD'],
   ['banksToTrack', JSON.stringify({ POPULAR: true })]].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  const tx = mock.ss.getSheetByName('Transactions');
  // what v1.1.70 saved for this email
  const old = new Array(15).fill(''); old[0] = h.date(2026, 3, 15); old[1] = 'POPULAR'; old[3] = 1250; old[4] = 'DOP';
  old[2] = old[6] = '1234. A continuación le informamos el detalle de'; old[5] = 'Unclassified'; old[7] = POP.subject;
  old[9] = 'NO'; old[11] = 'Transaction'; old[12] = 'p1';
  tx.getRange(2, 1, 1, 15).setValues([old]);
  const steps = () => Object.fromEntries(h.plain(h.ctx.startHereStatus(h.date(2026, 10, 14, 9))).steps.map(s => [s.id, s]));
  const s = steps().popularMerchant;
  assert.equal(s.status, 'warn');
  assert.match(s.title, /^1 POPULAR purchase/);
  assert.match(s.detail, /2026-03-15 and 2026-03-15.*Monitor by Date Range/);
  assert.deepEqual(s.action, { label: 'Open Transactions', fn: 'openSheetByName', arg: 'Transactions' });
  assert.equal(h.ctx.openSheetByName('Transactions'), true);
  // delete the row, run the dates again: back with the merchant's name, and the warning is gone
  tx.getRange(2, 1, 1, 15).clearContent();
  mock.gmail.threads.push(fakeThread('tp', [h.fakeMessage({ subject: POP.subject, from: POP.from, body: fixture('popular_consumo_wrapped'),
    date: h.date(2026, 3, 15, 16), id: 'p1' })]));
  h.ctx.runGmailMonitorForDateRange('2026-03-01', '2026-03-31');
  const rows = tx._rows(15).filter(r => r[0]);
  assert.deepEqual(rows.map(r => [r[2], r[5]]), [['PedidosYa*Pizza Ejemplo Sa', 'Delivery']]);
  assert.equal(steps().popularMerchant, undefined);
});
