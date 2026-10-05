'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');
const { makeServices, fakeThread } = require('./sheets_mock');

// ---------- Dominican payroll (2026 parameters)
test('DR payroll 2026: matches the published RD$50,000 example (SFS 1,520 · AFP 1,435 · ISR 1,854 · net 45,191)', () => {
  const h = load();
  const r = h.plain(h.ctx.computeDrPayroll(50000, 2026));
  assert.deepEqual([r.sfs, r.afp, r.isr, r.net], [1520, 1435, 1854, 45191]);
  assert.equal(r.stale, false);
});

test('DR payroll: exempt salary, caps and the top bracket', () => {
  const h = load();
  assert.equal(h.ctx.computeDrPayroll(30000, 2026).isr, 0, 'below RD$34,685/month net of TSS → no ISR');
  const high = h.plain(h.ctx.computeDrPayroll(300000, 2026));
  assert.equal(high.sfs, 7059.79, 'SFS stops at the RD$232,230 cap (3.04%)');
  assert.equal(high.afp, 8610, 'AFP cap RD$464,460 not reached (2.87%)');
  const annual = (300000 - 7059.79 - 8610) * 12;
  assert.equal(high.isr, Math.round((79776 + (annual - 867123.01) * 0.25) / 12 * 100) / 100);
  const later = h.plain(h.ctx.computeDrPayroll(50000, 2027));
  assert.equal(later.stale, true, '2027 scale not loaded yet → flagged');
  assert.equal(later.year, 2026);
});

test('Setup validation: auto deductions only for DOP; daily summary options checked', () => {
  const h = load();
  const base = () => ({ email: 'a@b.co', incomeCurrency: 'DOP', monthlyIncome: 80000, deductionMode: 'auto',
    deductions: { ARS: 0, AFP: 0, ISR: 0, taxRate: 0 }, banksToTrack: { BHD: true }, notify: { enabled: false } });
  assert.equal(h.ctx.validateSetupInput(base()), null);
  assert.match(h.ctx.validateSetupInput(Object.assign(base(), { incomeCurrency: 'USD' })), /only available for a DOP salary/);
  assert.match(h.ctx.validateSetupInput(Object.assign(base(), { notify: { enabled: true, hour: 8, sections: { totals: false } } })), /at least one thing/);
  assert.match(h.ctx.validateSetupInput(Object.assign(base(), { notify: { enabled: true, email: 'nope', hour: 8, sections: { totals: true } } })), /valid email/);
  assert.equal(h.ctx.validateSetupInput(Object.assign(base(), { notify: { enabled: true, email: '', hour: 8, sections: { totals: true } } })), null);
});

// ---------- Setup Wizard: no personal data, neutral placeholders
test('Setup Wizard contains no personal values; a first run suggests the current account email', () => {
  const mock = makeServices({ ui: true });
  const h = load({ services: mock.services });
  h.ctx.openSetupWizard();
  const html = mock.ui.dialogs[0].html;
  // generic checks — they don't name anyone, so the test itself leaks nothing
  assert.ok(!/value="[^"]*@/.test(html), 'no email hard-coded as a field value');
  const placeholders = [...html.matchAll(/placeholder="([^"]*)"/g)].map(m => m[1]);
  assert.deepEqual([...new Set(placeholders)].sort(), ['% cashback', '0', '0.00', 'Card name', 'Closes (day)', 'Due (day)',
    'Same as the email above', 'you@example.com'], 'only neutral placeholders — no numbers anyone could mistake for real data');
  assert.ok(html.includes('placeholder="you@example.com"') && html.includes('placeholder="0.00"'));
  assert.ok(html.includes('"email":"new.user@example.com"'), 'prefill = the Google account opening the wizard');
  assert.ok(html.includes('id="dedAuto"') && html.includes('id="isrAmount"') && html.includes('id="notifyEnabled"'));
  const root = process.env.GS_ROOT || require('path').join(__dirname, '..', 'src');   // same copy the harness loads
  const code = require('fs').readdirSync(root).filter(f => f.endsWith('.gs'))
    .map(f => require('fs').readFileSync(require('path').join(root, f), 'utf8')).join('\n').toLowerCase();
  // bank and broker senders (v1.1.29: HAPI), and example.com — nobody's personal address
  const bankSenders = /@(bancolafise\.com|notificaciones\.lafise\.com|lafise\.com|lafise\.com\.do|banreservas\.com|scotiabank\.com|scotiabank\.com\.do|qik\.do|qik\.com\.do|bhd\.com\.do|banesco\.com\.do|popularenlinea\.com|bdi\.com\.do|hapi\.trade|example\.com)$/i;
  const emails = code.match(/[a-z0-9._+-]+@[a-z0-9-]+\.[a-z0-9.]+/g) || [];
  assert.deepEqual(emails.filter(e => !bankSenders.test(e)), [], 'only bank senders and example.com in the code');
  assert.ok(!/'day \d+'/.test(code), 'no statement or payment dates baked into the code');
});

// ---------- Save → Configuration, named ranges, triggers
function configured(extra) {
  const mock = makeServices(Object.assign({ ui: true }, extra || {}));
  const h = load({ services: mock.services });
  mock.ss.insertSheet('Transactions');
  mock.ss.insertSheet('Custom Rules');
  return { h, mock };
}
const wizardInput = over => Object.assign({
  email: 'owner@example.com', incomeCurrency: 'DOP', monthlyIncome: 50000, deductionMode: 'auto',
  deductions: { ARS: 0, AFP: 0, ISR: 0, taxRate: 0 },
  banksToTrack: { LAFISE: true, BANESCO: true, BHD: true, POPULAR: false, BDI: false },
  notify: { enabled: true, email: '', hour: 9, sections: { totals: true, vsAverage: true, transfers: true, recommendations: true, cashback: false } },
  timestamp: '2026-09-25T12:00:00Z'
}, over || {});

test('saving DOP + automatic: stores the calculated ARS/AFP/ISR and creates both daily triggers', () => {
  const { h, mock } = configured();
  h.ctx.saveSetupConfig(wizardInput());
  const c = h.plain(h.ctx.getConfig());
  assert.deepEqual([c.deductions.ARS, c.deductions.AFP, c.deductions.ISR, c.deductions.taxRate], [1520, 1435, 1854, 0]);
  assert.equal(c.deductionMode, 'auto');
  assert.equal(c.notifyEnabled, true);
  assert.equal(c.notifyHour, 9);
  assert.equal(c.notifySections.cashback, false);
  assert.deepEqual(mock.triggers.map(t => [t.handler, t.hour]).sort(), [['runGmailMonitor', 6], ['sendDailySummary', 9]]);

  h.ctx.saveSetupConfig(wizardInput({ notify: { enabled: false, hour: 8, sections: { totals: true } } }));
  assert.deepEqual(mock.triggers.map(t => t.handler), ['runGmailMonitor'], 'summary trigger removed when turned off');
});

test('Dashboard income box: ISR row for DOP, rate row for USD; card table points to the wizard until cards are set', () => {
  const { h, mock } = configured();
  h.ctx.saveSetupConfig(wizardInput());
  h.ctx.buildOrRefreshDashboard();
  const dash = mock.ss.getSheetByName('Dashboard');
  const all = [...dash.cells.values()].filter(v => typeof v === 'string');
  assert.ok(all.some(v => v.includes('CFG_ISR') && v.includes('CFG_TAX_RATE')), 'tax row switches on currency');
  assert.ok(all.some(v => v.includes('"ISR — Impuesto Sobre la Renta"')));
  assert.equal(mock.ss.getRangeByName('CFG_ISR').getValue(), 1854);
  assert.deepEqual(h.plain(mock.ss.getRangeByName('DASH_CARDS').getValues()),
    [['No credit cards set up yet', '', '', '', 'Add them in 📊 Tracker › Setup Wizard.']], 'v1.1.27: no empty bank rows');
  assert.ok(dash.getColumnWidth(2) >= 300, 'column B fits the longest category label');
});

// ---------- Auto-fit columns
test('columns widen to fit the data after a run; hidden columns and wider manual widths are kept', () => {
  const { h, mock } = configured();
  h.ctx.saveSetupConfig(wizardInput({ notify: { enabled: false } }));
  const tx = mock.ss.getSheetByName('Transactions');
  tx.setColumnWidth(8, 500);                                      // user widened Email Subject by hand
  const msg = h.fakeMessage({ subject: 'Alerta de Consumo Banesco RD', from: 'notificaciones@banesco.com.do',
    body: 'consumo de RD$ 946.08, en SM NACIONAL METRO PLAZA ARROYO HONDO SANTO DOMINGO y su estado es aprobada',
    date: h.date(2026, 9, 20), id: 'fit-1' });
  mock.gmail.threads.push(fakeThread('t-fit', [msg]));
  h.ctx.runGmailMonitorForDateRange('2026-09-20', '2026-09-20');
  assert.ok(tx.getColumnWidth(3) > 250, 'Merchant widened: ' + tx.getColumnWidth(3));
  assert.equal(tx.getColumnWidth(8), 500, 'never narrower than a width you set');
  assert.equal(tx.getColumnWidth(9), 100, 'hidden Timestamp untouched');
  assert.ok(tx.getColumnWidth(3) <= 420, 'capped');
  assert.ok(!h.logs.some(l => l.includes('Could not fit')));
});

// ---------- Education removed
test('Education is no longer a default category', () => {
  const h = load();
  assert.ok(!h.plain(h.ctx.getCategories()).includes('Education'));
  assert.equal(h.ctx.categorizeTransaction('ACADEMIA DE INGLES'), 'Dining/Delivery + Entertainment + Other');
});

// ---------- Daily summary
function rowsFor(h, list) {
  const header = new Array(14).fill('h');
  return [header].concat(list.map(o => {
    const r = new Array(14).fill('');
    r[0] = o.date; r[1] = o.bank || 'LAFISE'; r[2] = o.merchant || 'X'; r[3] = o.amount; r[4] = o.currency || 'DOP';
    r[5] = o.category === undefined ? 'Dining/Delivery + Entertainment + Other' : o.category;
    r[11] = o.type || 'Transaction';
    return r;
  }));
}

test('daily summary: totals, average, pace, open transfers, card tips (pure)', () => {
  const h = load();
  const list = [];
  for (let d = 1; d <= 23; d++) list.push({ date: h.date(2026, 9, d), amount: 1000 });          // steady RD$1,000/day
  list.push({ date: h.date(2026, 9, 24), merchant: 'FARMA VALUE FD05', bank: 'POPULAR', amount: 3000, category: 'Health + Vet + Pharmacy' });
  list.push({ date: h.date(2026, 9, 24), merchant: 'UBER*EATS', amount: 25, currency: 'USD' });
  list.push({ date: h.date(2026, 9, 24), merchant: 'PAGO TARJETA', amount: 20000, type: 'Card Payment', category: 'Exclude' });
  list.push({ date: h.date(2026, 9, 24), merchant: 'MARIA PRUEBA', amount: 1500, type: 'Transfer', category: '' });
  const s = h.plain(h.ctx.computeDailySummary(rowsFor(h, list), {
    today: h.date(2026, 9, 25, 8), rates: { USD: 60, EUR: 64, COP: 0.015 }, netIncomeDop: 60000,
    cards: [{ bank: 'BHD', rate: 0.05 }, { bank: 'BANESCO', rate: 0.07 }]
  }));
  assert.equal(s.dayKey, '2026-09-24');
  assert.equal(s.spent, 3000 + 25 * 60, 'card payment and open transfer are not spending');
  assert.equal(s.purchases, 2);
  // v1.1.64: the 30-day average became a typical day (median of up to 60 days); with a steady history both are RD$1,000
  assert.equal(s.historyDays, 23, 'typical day uses only the days that have history');
  assert.equal(s.typicalDay, 1000);
  assert.equal(Math.round(s.vsTypical * 100), 350);
  assert.equal(s.openTransfersTotal, 1500);
  assert.equal(s.mtd, 23000 + 4500);
  // v1.1.64: warned because the day is among the highest of its history, and the RD$3,000 pharmacy is named as a one-off
  assert.ok(s.recommendations.some(r => /one of your highest days/.test(r.text) && /It includes FARMA VALUE FD05, RD\$3,000/.test(r.text)));
  assert.ok(s.recommendations.some(r => /1 transfer\(s\) this month have no category/.test(r.text)));
  assert.ok(s.recommendations.some(r => /of your net income/.test(r.text)));
  assert.equal(s.tips.length, 1);
  assert.match(s.tips[0].text, /FARMA VALUE FD05 \(RD\$3,000 on POPULAR\): pharmacies earn 5% with your BHD card — about RD\$150 back/);

  const noBhd = h.plain(h.ctx.computeDailySummary(rowsFor(h, list), {
    today: h.date(2026, 9, 25, 8), rates: { USD: 60, EUR: 64, COP: 0.015 }, netIncomeDop: 60000, cards: [] }));
  assert.equal(noBhd.tips.length, 0, 'no tip for a card you don\'t have');
});

test('daily summary email: sections switch off, text is escaped, link to the Dashboard', () => {
  const h = load();
  const rows = rowsFor(h, [{ date: h.date(2026, 9, 24), merchant: '<b>Bad & Co</b>', amount: 500 }]);
  const s = h.ctx.computeDailySummary(rows, { today: h.date(2026, 9, 25), rates: { USD: 60, EUR: 64, COP: 0.015 }, netIncomeDop: 0, cards: [] });
  const mail = h.plain(h.ctx.buildDailySummaryEmail(s, { sections: { cashback: false, transfers: false },
    sheetUrl: 'https://docs.google.com/x#gid=5', sheetName: 'Financial_Tracker_2026' }));
  assert.equal(mail.subject, 'Thu 24 Sep 2026: RD$500 spent');
  assert.ok(mail.html.includes('&lt;b&gt;Bad &amp; Co&lt;/b&gt;') && !mail.html.includes('<b>Bad'));
  assert.ok(mail.html.includes('https://docs.google.com/x#gid=5') && mail.html.includes('Open Financial_Tracker_2026'));
  assert.ok(!mail.html.includes('Card &amp; cashback tips') && !mail.html.includes('Transfers to review'));
  assert.ok(mail.html.includes('Recommendations') && mail.text.includes('Recommendations:'));
  assert.ok(mail.html.includes('not enough history yet'));
});

test('daily summary delivery: trigger respects the on/off switch; the menu item always sends', () => {
  const { h, mock } = configured();
  h.ctx.saveSetupConfig(wizardInput({ notify: { enabled: false, hour: 8, sections: { totals: true } } }));
  h.ctx.sendDailySummary();
  assert.equal(mock.gmail.sent.length, 0, 'turned off → the trigger sends nothing');
  h.ctx.sendDailySummaryNow();
  assert.equal(mock.gmail.sent.length, 1);
  assert.equal(mock.gmail.sent[0].to, 'owner@example.com');
  assert.match(mock.gmail.sent[0].options.htmlBody, /TEST_ID\/edit/);

  const only = { totals: true, vsAverage: false, transfers: false, recommendations: false, cashback: false }; // the wizard always sends all five
  h.ctx.saveSetupConfig(wizardInput({ notify: { enabled: true, email: 'alerts@example.com', hour: 8, sections: only } }));
  h.ctx.sendDailySummary();
  assert.equal(mock.gmail.sent[1].to, 'alerts@example.com');
  assert.ok(!mock.gmail.sent[1].options.htmlBody.includes('Recommendations'), 'only the chosen sections');
});

test('summary email survives Gmail encoding: ASCII-only HTML, no emoji in subject or text (v1.1.25)', () => {
  const h = load();
  const list = [];
  for (let d = 1; d <= 23; d++) list.push({ date: h.date(2026, 9, d), amount: 1000 });
  list.push({ date: h.date(2026, 9, 24), merchant: 'Café Ñandú 🍕', amount: 200 });
  list.push({ date: h.date(2026, 9, 24), merchant: 'X', amount: 800, type: 'Transfer', category: '' });
  const s = h.ctx.computeDailySummary(rowsFor(h, list), { today: h.date(2026, 9, 25), rates: { USD: 60, EUR: 64, COP: 0.015 }, netIncomeDop: 60000, cards: [] });
  const mail = h.plain(h.ctx.buildDailySummaryEmail(s, { sheetUrl: 'https://x', sheetName: 'Tracker' }));
  assert.ok(/^[\x00-\x7F]*$/.test(mail.html), 'every non-ASCII character is an entity');
  const astral = /[\u{10000}-\u{10FFFF}]/u;
  assert.ok(!astral.test(mail.subject) && !astral.test(mail.text), 'no emoji where entities do not exist');
  assert.ok(mail.subject.includes('−'), 'BMP characters (which Gmail handles) are kept in the subject');
  assert.ok(mail.html.includes('&#128994;'), '🟢 sent as an entity');
  assert.ok(mail.html.includes('Caf&#233; &#209;and&#250; &#127829;'), 'accents and emoji in merchant names too');
  const decoded = mail.html.replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(Number(n)));
  assert.ok(decoded.includes('🟢&nbsp; Nice — day-to-day') && decoded.includes('Café Ñandú 🍕'), 'decodes back to the original text');
  assert.ok(mail.text.includes('Nice — day-to-day spending'), 'plain text keeps the words, drops only the icon');
});
