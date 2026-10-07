'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, fixture } = require('./harness');
const { makeServices } = require('./sheets_mock');

// v1.1.66: an email whose plain-text part is empty is read from its HTML (LAFISE's "¡Transferencia exitosa!"); one
// between your own accounts is Exclude; a statement credit already saved from its email is not saved again; an amount
// typed by hand in a format that can't be read safely ("1.234,56") is not saved as another number.

const LAFISE_APP = 'Banco LAFISE <digital@notificaciones.lafise.com>';
const SUBJECT = 'Ana, ¡Transferencia exitosa! 👌';
// as the real email: a text/plain part with nothing in it, everything in the HTML
const message = (h, plain, html) => ({ getSubject: () => SUBJECT, getFrom: () => LAFISE_APP, getPlainBody: () => plain,
  getBody: () => html, getDate: () => h.date(2026, 9, 18, 10), getId: () => 'm-lafise-app' });

test('reported: LAFISE "¡Transferencia exitosa!" has an empty plain-text part — it is read from its HTML, not sent to Unrecognized', () => {
  const h = load();
  const html = fixture('lafise_transferencia_exitosa_html');
  for (const empty of ['', '\r\n', ' \n \n']) {
    const r = h.plain(h.ctx.parseEmailMessage(message(h, empty, html), {}, h.ctx.newParseStats()));
    assert.equal(r.status, 'ok', 'plain part ' + JSON.stringify(empty) + ': ' + (r.reason || ''));
    assert.deepEqual(r.items.map(t => [t.bank, t.type, t.currency, t.amount, t.txRef]), [['LAFISE', 'Transfer', 'USD', 75.5, 'LAFISE:990011223344']]);
  }
});

test('the HTML as text: one line per cell or block, entities decoded, no styles, scripts or comments', () => {
  const h = load();
  const text = h.ctx.htmlToPlainText(fixture('lafise_transferencia_exitosa_html'));
  const lines = text.split('\n');
  ['Hola Ana,', 'Acabas de realizar una transferencia de', 'USD 75.50', 'entre tus cuentas.',
    'De paso te dejamos tu número de referencia: 990011223344', 'Banco LAFISE República Dominicana']
    .forEach(l => assert.ok(lines.indexOf(l) !== -1, 'line: ' + l));
  assert.doesNotMatch(text, /font-family|<|&[a-z]+;|BEGIN MODULE|\bmso\b/);
  assert.match(text, /donde lo necesitas 😃/, 'a numeric entity outside the BMP');
  assert.doesNotMatch(text, /\n\n\n/, 'no runs of empty lines');
  assert.equal(h.ctx.htmlToPlainText('<p>A&amp;B &lt;x&gt;</p><script>var a = 1;</script><div>C&nbsp;&nbsp;D</div>'), 'A&B <x>\nC D');
});

test('a plain-text part with text is used as before: the HTML is not even asked for', () => {
  const h = load();
  let asked = false;
  const m = { getSubject: () => 'x', getPlainBody: () => 'Monto: RD$ 100.00', getBody: () => { asked = true; return '<p>other</p>'; } };
  assert.equal(h.ctx.emailPlainText(m), 'Monto: RD$ 100.00');
  assert.equal(asked, false);
  const old = { getPlainBody: () => '' };   // a message without getBody (tests, older mocks): the empty text, no error
  assert.equal(h.ctx.emailPlainText(old), '');
});

test('LAFISE app transfer "entre tus cuentas" is yours: Exclude, "(own account)"; another wording is a normal transfer', () => {
  const h = load();
  const html = fixture('lafise_transferencia_exitosa_html');
  const own = h.plain(h.ctx.parseEmailMessage(message(h, '', html), {}, h.ctx.newParseStats())).items[0];
  assert.equal(own.category, 'Exclude');
  assert.match(own.description, /\(own account\)$/);
  const other = h.plain(h.ctx.parseEmailMessage(message(h, '', html.replace('entre tus cuentas.', 'a Pedro Ejemplo.')), {}, h.ctx.newParseStats())).items[0];
  assert.equal(other.category, '');
  assert.doesNotMatch(other.description, /own account/);
});

test('found in the regression: a transfer SENT to your own account (v1.1.63) lost its Exclude at the end of every run', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  const cfg = mock.ss.insertSheet('Configuration');
  [['Key', 'Value'], ['email', 'user@example.com'], ['banksToTrack', '{"LAFISE":true}']].forEach(r => cfg.appendRow(r));
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  mock.ss.insertSheet('Custom Rules').appendRow(['UserEmail', 'Category', 'Keyword', 'Timestamp']);
  const own = h.ctx.parseEmailMessage(message(h, '', fixture('lafise_transferencia_exitosa_html')), {}, h.ctx.newParseStats()).items;
  const notice = h.ctx.parseEmailMessage(h.fakeMessage({ subject: 'Aviso de transferencia en banco local', from: 'bancanet@notificaciones.lafise.com',
    body: fixture('lafise_aviso_transferencia_local').replace('Estado: Error', 'Estado: Exitosa'), date: h.date(2026, 10, 1), id: 'aviso' }), {}, h.ctx.newParseStats()).items;
  notice[0].category = 'Exclude'; notice[0].description = notice[0].merchant + ' (own account)';   // as one to your own name is saved
  h.ctx.saveTransactions(own.concat(notice));
  h.ctx.recategorizeAllTransactions('user@example.com');
  const rows = mock.ss.getSheetByName('Transactions').getDataRange().getValues().slice(1);
  assert.deepEqual(rows.map(r => [r[11], r[5]]), [['Transfer', 'Exclude'], ['Transfer', 'Exclude']]);
  const TX = h.get('TX_COL');
  const third = []; third[TX.SUBJECT] = 'x'; third[TX.MERCHANT] = 'PEDRO'; third[TX.DESCRIPTION] = 'PEDRO'; third[TX.TYPE] = 'Transfer';
  assert.equal(h.plain(h.ctx.computeRecategorization(third, {})).category, '', 'a transfer to someone else stays open');
});

// --- Banesco: a transfer received is notified by email (v1.1.62) AND listed in the monthly statement (v1.1.55) ---
const stmtRow = (h, day, amount, balance) => ({ date: h.date(2026, 7, day), bank: 'BANESCO', merchant: 'ACH transfer (sender not in the statement)',
  amount: -amount, currency: 'DOP', type: 'Incoming', category: '', messageId: 'stmt_202607' + day + '-' + amount * 100,
  txRef: 'BANESCO:STMT:2026-07-' + day + ':' + amount.toFixed(2) + ':' + balance.toFixed(2) });
const emailRow = (h, day, amount, ref) => ({ date: h.date(2026, 7, day), bank: 'BANESCO', merchant: 'Transferencia recibida desde BANCO EJEMPLO',
  amount: -amount, currency: 'DOP', type: 'Incoming', category: '', messageId: 'mail-' + ref + '_0', txRef: 'BANESCO:IN:' + ref });
const savedValues = (h, items) => [new Array(15).fill('h')].concat(items.map(t => h.ctx.transactionToRow(Object.assign({ description: t.merchant, subject: 's', timestamp: 't' }, t))));

test('reported: a Banesco transfer received was saved twice — from its email and again from the statement', () => {
  const h = load();
  // the email came first (on the 10th); the statement lists it 3 days later, when the bank posts it
  const index = h.ctx.buildExistingIndex(savedValues(h, [emailRow(h, 10, 3150, 'E1')]));
  const sel = h.ctx.selectNewTransactions([stmtRow(h, 13, 3150, 12500), stmtRow(h, 14, 800, 13300)], index);
  assert.deepEqual(h.plain(sel.fresh.map(t => -t.amount)), [800], 'the 3,150 is already saved from its email; the 800 was never notified');
  assert.equal(sel.duplicates, 1);
});

test('statement and email in the same run, in either order; the other bank, currency or amount is not a match', () => {
  const h = load();
  const empty = () => h.ctx.buildExistingIndex([new Array(15).fill('h')]);
  let sel = h.ctx.selectNewTransactions([emailRow(h, 10, 3150, 'E1'), stmtRow(h, 13, 3150, 12500)], empty());
  assert.deepEqual(h.plain(sel.fresh.map(t => t.txRef)), ['BANESCO:IN:E1'], 'the email is kept: it names the sending bank');
  // a statement saved first (a run by dates read the email later): the email is the copy
  sel = h.ctx.selectNewTransactions([emailRow(h, 10, 3150, 'E1')], h.ctx.buildExistingIndex(savedValues(h, [stmtRow(h, 13, 3150, 12500)])));
  assert.equal(sel.fresh.length, 0);
  const notMatching = [Object.assign(stmtRow(h, 13, 3150, 12500), { bank: 'BHD', txRef: 'BHD:STMT:x', messageId: 'bhd' }),
    Object.assign(stmtRow(h, 13, 3150, 12500), { currency: 'USD', txRef: 'BANESCO:STMT:usd', messageId: 'usd' }),
    stmtRow(h, 13, 3150.01, 12500), stmtRow(h, 23, 3150, 15000), stmtRow(h, 3, 3150, 9000)];
  sel = h.ctx.selectNewTransactions(notMatching, h.ctx.buildExistingIndex(savedValues(h, [emailRow(h, 10, 3150, 'E1')])));
  assert.equal(sel.fresh.length, 5, 'other bank, other currency, another amount, 13 and 7 days apart');
  // two emails of the same amount are two transfers (each its own notice); so are two statement credits
  assert.equal(h.ctx.selectNewTransactions([emailRow(h, 10, 3150, 'E1'), emailRow(h, 11, 3150, 'E2')], empty()).fresh.length, 2);
  assert.equal(h.ctx.selectNewTransactions([stmtRow(h, 10, 3150, 5000), stmtRow(h, 11, 3150, 7500)], empty()).fresh.length, 2);
});

test('one email pairs with one statement credit: two equal transfers, one notified, both in the statement → one more saved', () => {
  const h = load();
  const index = h.ctx.buildExistingIndex(savedValues(h, [emailRow(h, 10, 3150, 'E1')]));
  const sel = h.ctx.selectNewTransactions([stmtRow(h, 11, 3150, 10000), stmtRow(h, 13, 3150, 12500)], index);
  assert.equal(sel.fresh.length, 1);
  // pairs saved before this version (both rows already there) are paired again when the index is built: a later
  // email of the same amount is not taken as the copy of that statement credit
  const both = h.ctx.buildExistingIndex(savedValues(h, [emailRow(h, 10, 3150, 'E1'), stmtRow(h, 13, 3150, 12500)]));
  assert.equal(h.ctx.selectNewTransactions([emailRow(h, 12, 3150, 'E2')], both).fresh.length, 1);
  // a credit typed by hand in Incoming Transfers counts as notified, too
  const typed = Object.assign(emailRow(h, 18, 700, 'x'), { messageId: 'manual:1:in:5', txRef: '', bank: 'BANESCO' });
  assert.equal(h.ctx.selectNewTransactions([stmtRow(h, 19, 700, 13200)], h.ctx.buildExistingIndex(savedValues(h, [typed]))).fresh.length, 0);
  // the statement read again (every run): its rows are duplicates by their reference, as before
  const again = h.ctx.buildExistingIndex(savedValues(h, [emailRow(h, 10, 3150, 'E1'), stmtRow(h, 14, 800, 13300)]));
  const sel2 = h.ctx.selectNewTransactions([stmtRow(h, 13, 3150, 12500), stmtRow(h, 14, 800, 13300)], again);
  assert.equal(sel2.fresh.length, 0);
});

// --- amounts typed by hand ---
test('typed amounts: 1,234.56, 4800, RD$ 1,500, -300 and numbers are read; 1.234,56 or 1.500 are not guessed', () => {
  const h = load();
  const a = v => h.ctx.typedAmount(v);
  assert.deepEqual([a(1234.5), a('1,234.56'), a('4800'), a('RD$ 1,500'), a('-300'), a(' 25.5 '), a('US$12.00'), a('DOP 99.99')],
    [1234.5, 1234.56, 4800, 1500, 300, 25.5, 12, 99.99]);
  ['1.234,56', '1234,56', '1.500', '12,34', '1,2345', 'abc', '', '1-2', '0'].forEach(v => assert.ok(!(a(v) > 0), JSON.stringify(v) + ' → ' + a(v)));
});

test('reported shape: "1.234,56" typed in Incoming Transfers goes to Unrecognized instead of being saved as 1.23', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  mock.ss.insertSheet('Transactions'); h.ctx.initializeTransactionsSheet();
  const inc = mock.ss.insertSheet('Incoming Transfers');
  inc.getRange(1, 1, 1, 8).setValues([h.get('TRANSFERS_HEADERS')]);
  inc.getRange(2, 1, 2, 6).setValues([['2026-08-28', 'LAFISE', 'Roommate', 'Rent', '1.234,56', 'DOP'], ['2026-08-29', 'LAFISE', 'Roommate', 'Rent', '1,234.56', 'DOP']]);
  assert.deepEqual(h.plain(h.ctx.importTypedIncomingRows()), { added: 1, rejected: 1, repaired: 0 });
  const tx = mock.ss.getSheetByName('Transactions').getDataRange().getValues().slice(1);
  assert.deepEqual(tx.map(r => r[3]), [-1234.56]);
  const un = h.plain(mock.ss.getSheetByName('Unrecognized').getDataRange().getValues()).find(r => r[1] === 'Incoming Transfers');
  assert.match(un[3], /1,234\.56/, 'the reason shows the format to type');
  assert.match(un[4], /1\.234,56/, 'what was typed is kept');
});
