'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');

const DINING = 'Dining/Delivery + Entertainment + Other';

test('keyword engine: short keywords match whole words only (C3, M2)', () => {
  const h = load();
  const m = (text, kw) => h.ctx.keywordMatches(text, kw);
  assert.equal(m('CACHAREPA CHURCHILL', 'ACH'), false, "'ACH' must not match inside CACHAREPA");
  assert.equal(m('PAGO VIA ACH A TERCEROS', 'ACH'), true);
  assert.equal(m('UBER*RIDES SANTO DOMINGO DOM', 'UBER'), true, "'*' counts as a boundary");
  assert.equal(m('BUSINESS CENTER SDQ', 'BUS'), false);
  assert.equal(m('EUROPA TRAVEL', 'ROPA'), false);
  assert.equal(m('SUPERMERCADOS NACIONAL', 'SUPER'), true, 'longer keywords stay substring matches');
  assert.equal(m('CUALQUIER COSA', ''), false, 'an empty keyword never matches (C4)');
  assert.equal(m('BHD TRANSFER (TO ACCOUNT ...1111)', 1111), true, 'numeric keywords work (C4)');
});

test('categorization: verified false positives are gone (M2)', () => {
  const h = load();
  const c = t => h.ctx.categorizeTransaction(t);
  assert.notEqual(c('AUTO RENT A CAR SDQ'), 'Rent');
  assert.notEqual(c('LUZ DE LUNA CAFE'), 'Electricity');
  assert.notEqual(c('BUSINESS CENTER SDQ'), 'Transportation');
  assert.notEqual(c('EUROPA TRAVEL SHOP'), 'Groceries + Barbershop');
});

test('categorization: regressions from earlier releases still hold', () => {
  const h = load();
  const c = t => h.ctx.categorizeTransaction(t);
  const cases = [
    ['UBER EATS-W*UBER EATS- SANTO DOMINGO DOM', DINING],            // v1.1.17
    ['UBER*EATS SANTO DOMINGO DOM', DINING],
    ['UBER*RIDES SANTO DOMINGO DOM', 'Transportation'],
    ['UBER RIDES-*UBER RIDES SANTO DOMINGO DOM', 'Transportation'],
    ['LA CASITA DE', 'Health + Vet + Pharmacy'],                      // v1.1.18, POPULAR truncation
    ['LA CASITA DE BOB-EP SANTO DOMINGO DOM', 'Health + Vet + Pharmacy'],
    ['BONJOUR TOTAL NACO', DINING],                                   // v1.1.15 ordering
    ['TOTALENERGIES MIRAMAR SANTO DOMINGO DOM', 'Vehicle Gas'],
    ['HOLA PLAZA LAS AMERICAS', 'Groceries + Barbershop'],
    ['FCIA MEDICAR GBC 30 DE MA', 'Health + Vet + Pharmacy'],
    ['PedidosYa*Som Cafe', DINING],
    ['AMAZON PRIME*2K3', 'Streaming & Subscriptions'],
    ['AMAZON MKTPL*AB12', DINING],
    ['EDESUR DOMINICANA', 'Electricity'],
    ['SPORTS BAR PIANTINI', DINING],
    ['DiDi CO Ride', 'Transportation'],                               // v1.1.20, real POPULAR/COP rows
    ['DL*DIDI RIDES', 'Transportation'],
    ['DLO*Didi', 'Transportation'],
    ['DIDI FOOD BOGOTA', DINING],                                     // delivery, not a ride
    ['ALISS METRO PZA', DINING],                                      // v1.1.23 — a mall, not the metro
    ['SM NACIONAL METRO PLZA', 'Groceries + Barbershop'],
    ['UBER * EATS PENDING help.uber.com NLD', DINING],                // spaces around '*'
    ['UBER * PENDING Amsterdam NLD', 'Transportation'],               // ambiguous pending charge stays a ride
    ['SM POLA INDEPENDENCIA', 'Groceries + Barbershop'],
    ['FARMA VALUE FD05', 'Health + Vet + Pharmacy'],
    ['FARM CAROL ENRIQUILLO', 'Health + Vet + Pharmacy'],
    ['BANCO POPULAR CAROL AV. I', DINING],                            // 'CAROL' alone is not a pharmacy
    ['PAYPAL *UBERBV', 'Transportation'],
    ['PAYPAL *UBER BV', 'Transportation'],
    ['Google', 'Streaming & Subscriptions'],
    ['GOOGLE *Google', 'Streaming & Subscriptions'],
    ['ANALISA GAZCUE', 'Health + Vet + Pharmacy'],
    ['CTRO ESPEC MED', 'Health + Vet + Pharmacy'],
    ['ALGO QUE NADIE CONOCE', DINING]                                 // fallback
  ];
  for (const [text, expected] of cases) assert.equal(c(text), expected, text);
});

test('Custom Rules: blank and numeric keywords are safe (C4)', () => {
  const email = 'user@example.com';
  const h = load({ sheets: { 'Custom Rules': [
    ['UserEmail', 'Category', 'Keyword', 'Timestamp'],
    [email, 'Exclude', '', ''],              // blank keyword: used to capture EVERYTHING
    [email, 'Exclude', 1111, ''],            // numeric: used to throw on every parse
    [email, 'Rent', 'Landlord Name', ''],
    [email, '', 'ORPHAN KEYWORD', ''],       // no category
    ['someone-else@example.com', 'Exclude', 'UBER', '']
  ] } });
  const rules = h.plain(h.ctx.getUserCustomRules(email));
  assert.deepEqual(rules, { Exclude: ['1111'], Rent: ['LANDLORD NAME'] });

  const raw = h.ctx.getUserCustomRules(email);
  assert.equal(h.ctx.findCustomRuleOverride('UBER*RIDES SANTO DOMINGO', raw), '', 'blank keyword no longer matches everything');
  assert.equal(h.ctx.findCustomRuleOverride('BHD Transfer (to account ...1111)', raw), 'Exclude');
  assert.equal(h.ctx.findCustomRuleOverride('LANDLORD NAME', raw), 'Rent');
});

test('custom-only categories are listed for the Dashboard, Exclude is not (M8)', () => {
  const email = 'user@example.com';
  const h = load({ sheets: { 'Custom Rules': [
    ['UserEmail', 'Category', 'Keyword', 'Timestamp'],
    [email, 'Pets', 'PETSMART', ''],
    [email, 'Rent', 'LANDLORD', ''],
    [email, 'exclude', 'MY OWN NAME', '']
  ] } });
  assert.deepEqual(h.plain(h.ctx.getCustomCategoryNames(email)), ['Pets']);
});
