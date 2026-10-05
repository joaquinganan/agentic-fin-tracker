'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

// v1.1.62. The lesson of this version: Gmail's getPlainBody() writes bold as *text*, and the v1.1.60/61 fixtures had no
// markers, so they passed while the real emails failed. The new fixtures keep the markers where the real emails have
// them (seen in the Unrecognized sheet and in the original .eml files), with invented data.
const parseBody = (h, subject, from, body) => h.plain(h.ctx.parseEmailMessage(h.fakeMessage({ subject, from, body,
  date: h.date(2026, 9, 27, 10), id: 'm' }), [], h.ctx.newParseStats()));
const row = t => [t.bank, t.type, t.currency, t.amount, t.merchant, t.category];
const F = {
  qik: ['Usaste tu tarjeta de crédito Qik', 'notificaciones@qik.do'],
  qikCash: ['Retiro con Código CASH exitoso', 'Qik Banco Digital <no-reply-qik@qik.com.do>'],
  lafiseDetail: ['Detalle de Transaccion Tarjeta de Crédito ', 'Alerta de Consumo Banco LAFISE <notificacioneslafisedo@lafise.com.do>'],
  lafisePay: ['Notificación de pago de tarjeta de crédito', '"Banco LAFISE" <bancanet@notificaciones.lafise.com>'],
  bdi: ['BDI Digital: Comprobante transacción Interbancaria', '"bdinforma@bdi.com.do" <bdinforma@bdi.com.do>'],
  bdiConsumo: ['Notificacion de Consumos', 'BDI Informa <BDIinforma@bdi.com.do>'],
  bdiPay: ['BDI Digital: Comprobante de Transacción', '"bdinforma@bdi.com.do" <bdinforma@bdi.com.do>'],
  scotia: ['Autorización fuera del país', 'Alertas Scotiabank <alertas@scotiabank.com>'],
  scotiaIn: ['Pago al Instante recibido', 'Alertas Scotiabank <alertas@scotiabank.com>'],
  brApp: ['Recibo de la transacción', '<NotificacionesTuBancoApp@banreservas.com>'],
  brNotif: ['Notificaciones Banreservas', 'notificaciones@banreservas.com'],
  banescoIn: ['Notificación de Transferencia Recibida', 'Notificaciones Banesco <notificaciones@banesco.com.do>']
};

test('QIK purchase as Gmail gives it (amount and merchant in bold): read, not filtered by its footer (reported)', () => {
  const h = load();
  const r = parseBody(h, ...F.qik, fixture('qik_consumo_bold'));
  assert.equal(r.status, 'ok', 'it was filtered without a trace: the extractor missed "*RD$ 45.00*", so the security footer won');
  assert.deepEqual(r.items.map(x => row(x).slice(0, 5)), [['QIK', 'Transaction', 'DOP', 45, 'CAFETERIA EJEMPLO']]);
});

test('LAFISE second template as Gmail gives it ("*COMERCIO: *…"): read in DOP and in USD ("Dolares US"); the first template is untouched', () => {
  const h = load();
  const dop = parseBody(h, ...F.lafiseDetail, fixture('lafise_consumo_detalle_tc'));
  assert.equal(dop.status, 'ok', 'reported: "Amount not found" on every one of them');
  assert.deepEqual(dop.items.map(x => row(x).slice(0, 5)), [['LAFISE', 'Transaction', 'DOP', 1234.56, 'SUPERMERCADO EJEMPLO SANTO DOMINGODO']]);
  const usd = parseBody(h, ...F.lafiseDetail, fixture('lafise_consumo_detalle_tc_usd'));
  assert.deepEqual(usd.items.map(x => row(x).slice(0, 5)), [['LAFISE', 'Transaction', 'USD', 12.34, 'TIENDA*DIGITAL EJEMPLO Stockholm SE']], 'the * inside a name stays');
  const first = parseBody(h, 'Servicio de Alerta - Nuevo Consumo', 'LAFISE CF <notificaciones@bancolafise.com>', fixture('lafise_consumo_ubereats'));
  assert.deepEqual(first.items.map(x => row(x).slice(0, 5)), [['LAFISE', 'Transaction', 'DOP', 297.25, 'UBER EATS-W*UBER EATS- SANTO DOMINGO DOM']]);
});

test('LAFISE first template: the currency written next to the amount wins over anything else in the email', () => {
  const h = load();
  const withNoise = fixture('lafise_consumo_ubereats').replace('Autorización:\n000000', 'Autorización:\nUS 4000123');
  assert.deepEqual(h.plain(h.ctx.extractLAFISETransactions(withNoise)).map(x => [x.currency, x.amount]), [['DOP', 297.25]],
    'the context alone said USD: the first currency-like token anywhere between the merchant and the amount');
  assert.deepEqual(h.plain(h.ctx.extractLAFISETransactions(fixture('lafise_consumo_ubereats').replace('DOP 297.25', 'USD 6.50'))).map(x => x.currency), ['USD']);
  // v1.1.63: with no currency next to the amount, an authorization code like "US 4000123" is not dollars (reported: a
  // DOP purchase saved as USD) — a real "US$" in the context still is
  assert.deepEqual(h.plain(h.ctx.extractLAFISETransactions(withNoise.replace('DOP 297.25', '297.25'))).map(x => x.currency), ['DOP'],
    'no currency next to the amount: an authorization code is not a currency');
  assert.deepEqual(h.plain(h.ctx.extractLAFISETransactions(withNoise.replace('US 4000123', 'US$ 4000123').replace('DOP 297.25', '297.25'))).map(x => x.currency), ['USD']);
});

test('Gmail bold never changes what the newer extractors read: every fixture, each line in *…*, gives the same rows', () => {
  const h = load();
  const bold = text => text.split('\n').map(l => l.trim() ? '*' + l.trim() + '*' : l).join('\n');
  const cases = [['bdi_consumo', F.bdiConsumo], ['bdi_transferencia_interbancaria', F.bdi], ['bdi_transferencia_recibida', F.bdi],
    ['bdi_pago_tarjeta', F.bdiPay], ['scotiabank_consumo', F.scotia], ['scotiabank_pago_instante', F.scotiaIn], ['qik_consumo', F.qik],
    ['qik_codigo_cash', F.qikCash], ['lafise_pago_tarjeta', F.lafisePay], ['lafise_consumo_detalle_tc', F.lafiseDetail],
    ['banreservas_transferencia_tercero', F.brApp], ['banreservas_retiro_tuefectivo', F.brApp], ['banreservas_transferencia_recibida', F.brNotif],
    ['banesco_transferencia_recibida', F.banescoIn]];
  for (const [name, [subject, from]] of cases) {
    const plain = parseBody(h, subject, from, fixture(name).replace(/\*([^*\n]+)\*/g, '$1'));
    const bolded = parseBody(h, subject, from, bold(fixture(name)));
    assert.equal(plain.status, 'ok', name);
    assert.deepEqual(bolded.items.map(row), plain.items.map(row), name);
  }
  assert.equal(h.ctx.flatText('*COMERCIO: *UBER*EATS DO *MONTO: *10.00'), 'COMERCIO: UBER*EATS DO MONTO: 10.00');
});

test('BDI card payment from the account: a Card Payment (Exclude) and its tax as its own Transfer row; the "Completada" notice is not read twice', () => {
  const h = load();
  const r = parseBody(h, ...F.bdiPay, fixture('bdi_pago_tarjeta'));
  assert.equal(r.status, 'ok', 'it was Unrecognized');
  assert.deepEqual(r.items.map(row), [['BDI', 'Card Payment', 'DOP', 4000, 'Pago Tarjetas de Crédito (BDI)', 'Exclude'],
    ['BDI', 'Transfer', 'DOP', 8, 'BDI: impuesto y comisión de pago de tarjeta', '']]);
  const done = parseBody(h, 'BDI Digital: Aviso Notificación Transacción Interbancaria - Completada', F.bdi[1], fixture('bdi_interbancaria_completada'));
  assert.equal(done.status, 'filtered', 'it confirms a transfer already saved from its Comprobante (it even says [Salida])');
  assert.deepEqual(parseBody(h, ...F.bdi, fixture('bdi_transferencia_interbancaria')).items.map(x => x.amount), [2500, 5], 'the Comprobante is still read');
});

test('emails that are not a movement: QIK code created, Banreservas payroll and loan reminder, BHD purchase-validation code', () => {
  const h = load();
  assert.equal(parseBody(h, 'Código CASH para ti creado.', F.qikCash[1], fixture('qik_codigo_cash_creado')).status, 'filtered',
    'the money leaves with the withdrawal (its own email); counting the code too would count it twice');
  assert.deepEqual(parseBody(h, ...F.qikCash, fixture('qik_codigo_cash')).items.map(x => x.amount), [1500], 'the withdrawal is still read');
  assert.equal(parseBody(h, ...F.brNotif, fixture('banreservas_pago_nomina')).status, 'filtered', 'salary: set in the Setup Wizard');
  assert.equal(parseBody(h, 'Notificación de balances', F.brNotif[1], 'Notificaciones Banreservas Estimado Cliente: su préstamo No. 1234 cortará el 17/09/2026 y su cuota es de RD$ 5,000.00.').status, 'filtered');
  assert.equal(parseBody(h, 'Código de validación de compra', 'BHD <Alertas@bhd.com.do>', 'Para continuar la compra con tu Tarjeta por valor de 50 (USD), introduce el código de validación 000000.').status, 'filtered');
});

test('BANESCO transfer received: Incoming (negative), the sender from its field or from "LBTR <name>"; yours is Exclude', () => {
  const h = load();
  const r = parseBody(h, ...F.banescoIn, fixture('banesco_transferencia_recibida'));
  assert.deepEqual(r.items.map(row), [['BANESCO', 'Incoming', 'DOP', -15000, 'MARIA ELENA GOMEZ', 'Exclude']], 'reported: saved as a transfer sent');
  const other = parseBody(h, ...F.banescoIn, fixture('banesco_transferencia_recibida').replace('Concepto: LBTR MARIA ELENA GOMEZ', 'Concepto: alquiler'));
  assert.deepEqual(other.items.map(row), [['BANESCO', 'Incoming', 'DOP', -15000, 'Transferencia recibida desde BANCO EJEMPLO, S.A.', '']],
    'no sender named: the sending bank, never the beneficiary (that is you)');
  assert.equal(h.ctx.detectTypeFromSubject('Notificación de Transferencia Recibida'), 'Incoming');
  assert.equal(h.ctx.detectTypeFromSubject('Notificación de Transferencia Realizada'), 'Transfer', 'the one sent is unchanged');
});

test('Recategorize repairs a BANESCO transfer received saved before: Incoming, negative, a credit; a transfer sent is untouched', () => {
  const h = load();
  const TX = h.get('TX_COL');
  const mk = (subject, amount) => { const r = []; r[TX.SUBJECT] = subject; r[TX.MERCHANT] = 'X'; r[TX.DESCRIPTION] = 'X';
    r[TX.TYPE] = 'Transfer'; r[TX.CURRENCY] = 'DOP'; r[TX.AMOUNT] = amount; r[TX.IS_CREDIT] = 'NO'; return r; };
  const got = h.plain(h.ctx.computeRecategorization(mk('Notificación de Transferencia Recibida', 25000), []));
  assert.deepEqual([got.type, got.amount, got.isCredit], ['Incoming', -25000, 'YES']);
  const again = h.plain(h.ctx.computeRecategorization(mk('Notificación de Transferencia Recibida', -25000), []));
  assert.equal(again.amount, -25000, 'never flipped back');
  const sent = h.plain(h.ctx.computeRecategorization(mk('Notificación de Transferencia Realizada', 25000), []));
  assert.deepEqual([sent.type, sent.amount], ['Transfer', 25000]);
});

test('Recategorize on the sheet: the saved row becomes Incoming and negative, and lands in Incoming Transfers', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 80000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ BANESCO: true })]].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  const sheet = mock.ss.getSheetByName('Transactions');
  sheet.appendRow(h.ctx.transactionToRow({ date: h.date(2026, 10, 1), bank: 'BANESCO', merchant: 'MARIA ELENA GOMEZ', amount: 25000, currency: 'DOP',
    category: '', description: 'MARIA ELENA GOMEZ', subject: 'Notificación de Transferencia Recibida', timestamp: '', type: 'Transfer', messageId: 'old-1' }));
  h.ctx.recategorizeAllTransactions('user@example.com');
  const saved = h.plain(sheet._rows(15))[0];
  assert.deepEqual([saved[11], saved[3], saved[9]], ['Incoming', -25000, 'YES'], 'it counted as a transfer sent');
  const incoming = mock.ss.getSheetByName('Incoming Transfers');
  assert.ok(incoming && incoming._rows(6).some(r => r.includes('MARIA ELENA GOMEZ')), 'listed with money received');
  h.ctx.recategorizeAllTransactions('user@example.com');
  assert.equal(h.plain(sheet._rows(15))[0][3], -25000, 'stable on a second pass');
});

/* ---------- Dashboard: the formulas, computed ----------
 * A small evaluator for exactly what the Dashboard writes (SUMIFS by currency × rate, DATE, EDATE, SUM, INDEX on the
 * grid), run over Transactions. It checks the arithmetic, not only the formula text. */
function evaluator(ss) {
  const serial = d => Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
  const colN = L => L.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const tx = ss.getSheetByName('Transactions').getDataRange().getValues();
  const dash = ss.getSheetByName('Dashboard');
  const cache = {};
  const cellValue = (r, c) => {
    const key = r + ',' + c;
    if (key in cache) return cache[key];
    const v = dash.getRange(r, c).getValue();
    return (cache[key] = (typeof v === 'string' && v.charAt(0) === '=') ? evaluate(v.slice(1)) : v);
  };
  const match = (value, crit) => {
    const c = String(crit), m = c.match(/^(>=|<=|<>|>|<|=)?([\s\S]*)$/), op = m[1] || '=', rhs = m[2];
    const v = value && typeof value.getTime === 'function' ? serial(value) : value;   // dates from the script realm
    const num = Number(rhs);
    if (op === '=' ) return rhs === '' ? (v === '' || v === null || v === undefined) : String(v).toUpperCase() === rhs.toUpperCase();
    if (op === '<>') return String(v === null || v === undefined ? '' : v).toUpperCase() !== rhs.toUpperCase();
    if (typeof v !== 'number' || isNaN(num)) return false;
    return op === '>=' ? v >= num : op === '<=' ? v <= num : op === '>' ? v > num : v < num;
  };
  function evaluate(src) {
    let i = 0;
    const peek = () => src[i], ws = () => { while (src[i] === ' ') i++; };
    const expr = () => { let v = add(); ws(); while (peek() === '&') { i++; v = String(v) + String(add()); ws(); } return v; };
    const add = () => { let v = mul(); ws(); while (peek() === '+' || peek() === '-') { const o = src[i++]; const w = mul(); v = o === '+' ? v + w : v - w; ws(); } return v; };
    const mul = () => { let v = atom(); ws(); while (peek() === '*' || peek() === '/') { const o = src[i++]; const w = atom(); v = o === '*' ? v * w : v / w; ws(); } return v; };
    function atom() {
      ws();
      if (peek() === '"') { const j = src.indexOf('"', i + 1); const s = src.slice(i + 1, j); i = j + 1; return s; }
      if (peek() === '(') { i++; const v = expr(); ws(); i++; return v; }
      let m = src.slice(i).match(/^-?\d+(\.\d+)?/);
      if (m) { i += m[0].length; return Number(m[0]); }
      m = src.slice(i).match(/^Transactions!\$([A-Z]+):\$[A-Z]+/);
      if (m) { i += m[0].length; return { col: colN(m[1]) - 1 }; }
      m = src.slice(i).match(/^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/);
      if (m) { i += m[0].length; return { r1: +m[2], c1: colN(m[1]), r2: +m[4], c2: colN(m[3]) }; }
      m = src.slice(i).match(/^([A-Z]+)\(/);
      if (m) {
        i += m[0].length; const args = []; ws();
        while (peek() !== ')') { args.push(expr()); ws(); if (peek() === ',') i++; ws(); }
        i++;
        return call(m[1], args);
      }
      m = src.slice(i).match(/^\$?([A-Z]+)\$?(\d+)/);
      if (m) { i += m[0].length; return cellValue(+m[2], colN(m[1])); }
      throw new Error('evaluator: cannot read ' + src.slice(i, i + 30));
    }
    function call(fn, a) {
      if (fn === 'DATE') return serial(new Date(a[0], a[1] - 1, a[2]));
      if (fn === 'EDATE') { const d = new Date(Date.UTC(1899, 11, 30) + a[0] * 86400000); return serial(new Date(d.getUTCFullYear(), d.getUTCMonth() + a[1], d.getUTCDate())); }
      if (fn === 'SUMIFS') {
        let total = 0;
        for (let r = 1; r < tx.length; r++) {
          let ok = true;
          for (let k = 1; k < a.length; k += 2) if (!match(tx[r][a[k].col], a[k + 1])) { ok = false; break; }
          const v = tx[r][a[0].col];
          if (ok && typeof v === 'number') total += v;
        }
        return total;
      }
      if (fn === 'SUM') { let t = 0; const g = a[0]; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const v = cellValue(r, c); if (typeof v === 'number') t += v; } return t; }
      if (fn === 'IFERROR') return a[0];
      throw new Error('evaluator: no function ' + fn);
    }
    return expr();
  }
  return { cellValue };
}

test('Dashboard arithmetic, computed: every category, month and currency; Incoming lowers its category; payments, Exclude and uncategorized money in never count', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 80000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ LAFISE: true })]].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  const DIN = 'Dining/Delivery + Entertainment + Other';
  const sheet = mock.ss.getSheetByName('Transactions');
  const add = (m, d, amount, cur, cat, type) => sheet.appendRow(h.ctx.transactionToRow({ date: h.date(2026, m, d), bank: 'LAFISE', merchant: 'M',
    amount, currency: cur, category: cat, description: 'M', subject: 's', timestamp: '', type, messageId: 'id' + Math.random() }));
  // October: the reported case. A USD row is converted at the USD rate; a DOP row is itself.
  add(10, 1, 120, 'USD', DIN, 'Transaction'); add(10, 1, 432.10, 'DOP', DIN, 'Transaction');
  add(10, 2, 88.40, 'DOP', 'Transportation', 'Transaction'); add(10, 2, 141.75, 'DOP', 'Transportation', 'Transaction');
  add(10, 3, 10, 'EUR', 'Streaming & Subscriptions', 'Transaction');
  add(10, 1, 30000, 'DOP', 'Rent', 'Transfer');                 // a transfer with a category counts in it
  add(10, 4, 2000, 'DOP', '', 'Transfer');                      // uncategorized transfer: its own row
  add(10, 5, -300, 'DOP', 'Transportation', 'Incoming');        // money back for a shared ride
  add(10, 6, -25000, 'DOP', '', 'Incoming');                    // money received, no category: counts nowhere
  add(10, 6, 25000, 'DOP', 'Exclude', 'Transfer');              // between your own accounts
  add(10, 7, 5000, 'DOP', 'Exclude', 'Card Payment');           // paying the card
  add(10, 8, 25, 'DOP', DIN, 'Cashback');                       // cashback is not spending
  add(10, 9, 750, 'DOP', 'Viajes', 'Transaction');              // a category typed by hand, no Custom Rule
  add(9, 30, 255.50, 'DOP', DIN, 'Transaction');                // September stays in September
  add(11, 1, 100, 'DOP', DIN, 'Transaction');                   // and November in November
  h.ctx.buildOrRefreshDashboard();
  const ev = evaluator(mock.ss);
  const dash = mock.ss.getSheetByName('Dashboard');
  const usd = dash.getRange('K4').getValue(), eur = dash.getRange('M4').getValue();
  assert.ok(usd > 0 && eur > 0, 'rates are numbers');
  // the grid: find each category's row and the month columns (C = Jan … N = Dec)
  const findGrid = name => { let last = null; for (let r = 1; r <= 400; r++) if (dash.getRange(r, 2).getValue() === name) last = r; return last; };
  const at = (name, month) => Math.round(ev.cellValue(findGrid(name), 2 + month) * 100) / 100;
  const r2 = x => Math.round(x * 100) / 100;
  assert.equal(at(DIN, 10), r2(120 * usd + 432.10), 'the reported case: the USD row at the rate, plus the DOP row');
  assert.equal(at(DIN, 9), 255.50);
  assert.equal(at(DIN, 11), 100);
  assert.equal(at('Transportation', 10), r2(88.40 + 141.75 - 300), 'money back lowers its category');
  assert.equal(at('Streaming & Subscriptions', 10), r2(10 * eur));
  assert.equal(at('Rent', 10), 30000);
  assert.equal(at('Transfers (uncategorized)', 10), 2000, 'not the money received, not the Exclude transfer');
  assert.ok(findGrid('Viajes'), 'a category typed by hand gets its row (it was counted nowhere)');
  assert.equal(at('Viajes', 10), 750);
  assert.equal(findGrid('Exclude'), null, 'Exclude is never a row');
  for (const cat of h.plain(h.ctx.getCategories())) if (![DIN, 'Transportation', 'Streaming & Subscriptions', 'Rent'].includes(cat)) assert.equal(at(cat, 10), 0, cat);
  const total = r2(120 * usd + 432.10 + 88.40 + 141.75 - 300 + 10 * eur + 30000 + 2000 + 750);
  assert.equal(at('TOTAL', 10), total, 'TOTAL = every category + uncategorized transfers; no payment, cashback, Exclude or uncategorized money in');
  assert.equal(r2(ev.cellValue(findGrid(DIN), 15)), r2(120 * usd + 432.10 + 255.50 + 100), 'the year column adds the months');
});

test('a run with the new formats as Gmail gives them: each saved where it belongs', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['monthlyIncome', 80000], ['incomeCurrency', 'DOP'], ['deductionMode', 'manual'],
   ['banksToTrack', JSON.stringify({ LAFISE: true, BANESCO: true, BDI: true, QIK: true, BANRESERVAS: true })]].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  const list = [[F.qik, 'qik_consumo_bold'], [F.lafiseDetail, 'lafise_consumo_detalle_tc'], [F.lafiseDetail, 'lafise_consumo_detalle_tc_usd'],
    [F.bdiPay, 'bdi_pago_tarjeta'], [['BDI Digital: Aviso Notificación Transacción Interbancaria - Completada', F.bdi[1]], 'bdi_interbancaria_completada'],
    [['Código CASH para ti creado.', F.qikCash[1]], 'qik_codigo_cash_creado'], [F.brNotif, 'banreservas_pago_nomina'], [F.banescoIn, 'banesco_transferencia_recibida']];
  list.forEach(([[subject, from], body], i) => mock.gmail.threads.push(fakeThread('t' + i, [h.fakeMessage({ subject, from, body: fixture(body),
    date: h.date(2026, 9, 10 + i, 10), id: 'm' + i })])));
  h.ctx.runGmailMonitorForDateRange('2026-09-01', '2026-09-30');
  const tx = h.plain(mock.ss.getSheetByName('Transactions')._rows(15));
  assert.deepEqual(tx.map(r => [r[1], r[11], r[3], r[4]].join(' ')).sort(), ['BANESCO Incoming -15000 DOP', 'BDI Card Payment 4000 DOP',
    'BDI Transfer 8 DOP', 'LAFISE Transaction 12.34 USD', 'LAFISE Transaction 1234.56 DOP', 'QIK Transaction 45 DOP']);
  const unrec = mock.ss.getSheetByName('Unrecognized');
  assert.ok(!unrec || unrec._rows(5).length === 0, 'nothing left for Unrecognized');
});
