'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./harness');

// v1.1.70: the categories were rebuilt (02_categorizer.gs). Each case keeps the merchant of the release that fixed it;
// the expected category is the new one it belongs to — and the fallback is Unclassified, no longer Dining.

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
  assert.notEqual(c('AUTO RENT A CAR SDQ'), 'Rent & housing');
  assert.notEqual(c('LUZ DE LUNA CAFE'), 'Utilities');
  assert.notEqual(c('BUSINESS CENTER SDQ'), 'Public transport & parking');
  assert.notEqual(c('EUROPA TRAVEL SHOP'), 'Shopping & home');
});

test('categorization: regressions from earlier releases still hold', () => {
  const h = load();
  const c = t => h.ctx.categorizeTransaction(t);
  const cases = [
    ['UBER EATS-W*UBER EATS- SANTO DOMINGO DOM', 'Delivery'],        // v1.1.17
    ['UBER*EATS SANTO DOMINGO DOM', 'Delivery'],
    ['UBER*RIDES SANTO DOMINGO DOM', 'Ride apps & taxis'],
    ['UBER RIDES-*UBER RIDES SANTO DOMINGO DOM', 'Ride apps & taxis'],
    ['LA CASITA DE', 'Pets'],                                         // v1.1.18, POPULAR truncation (a vet clinic)
    ['LA CASITA DE BOB-EP SANTO DOMINGO DOM', 'Pets'],
    ['BONJOUR TOTAL NACO', 'Restaurants & cafés'],                    // v1.1.15 ordering: not Fuel's 'TOTAL'
    ['TOTALENERGIES MIRAMAR SANTO DOMINGO DOM', 'Fuel & vehicle'],
    ['HOLA PLAZA LAS AMERICAS', 'Groceries'],
    ['FCIA MEDICAR GBC 30 DE MA', 'Health & pharmacy'],
    ['PedidosYa*Som Cafe', 'Delivery'],                               // Delivery is tried before Restaurants' 'CAFE'
    ['AMAZON PRIME*2K3', 'Subscriptions'],
    ['AMAZON MKTPL*AB12', 'Shopping & home'],
    ['EDESUR DOMINICANA', 'Utilities'],
    ['SPORTS BAR PIANTINI', 'Leisure & going out'],
    ['DiDi CO Ride', 'Ride apps & taxis'],                            // v1.1.20, real POPULAR/COP rows
    ['DL*DIDI RIDES', 'Ride apps & taxis'],
    ['DLO*Didi', 'Ride apps & taxis'],
    ['DIDI FOOD BOGOTA', 'Delivery'],                                 // delivery, not a ride
    ['ALISS METRO PZA', 'Shopping & home'],                           // v1.1.23 — a mall, not the metro
    ['SM NACIONAL METRO PLZA', 'Groceries'],
    ['UBER * EATS PENDING help.uber.com NLD', 'Delivery'],            // spaces around '*'
    ['UBER * PENDING Amsterdam NLD', 'Ride apps & taxis'],            // ambiguous pending charge stays a ride
    ['SM POLA INDEPENDENCIA', 'Groceries'],
    ['FARMA VALUE FD05', 'Health & pharmacy'],
    ['FARM CAROL ENRIQUILLO', 'Health & pharmacy'],
    ['BANCO POPULAR CAROL AV. I', 'Unclassified'],                    // 'CAROL' alone is not a pharmacy
    ['PAYPAL *UBERBV', 'Ride apps & taxis'],
    ['PAYPAL *UBER BV', 'Ride apps & taxis'],
    ['Google', 'Subscriptions'],
    ['GOOGLE *Google', 'Subscriptions'],
    ['ANALISA GAZCUE', 'Health & pharmacy'],
    ['CTRO ESPEC MED', 'Health & pharmacy'],
    ['ALGO QUE NADIE CONOCE', 'Unclassified']                         // fallback (v1.1.70: no longer Dining)
  ];
  for (const [text, expected] of cases) assert.equal(c(text), expected, text);
});

test('Custom Rules: blank and numeric keywords are safe (C4)', () => {
  const email = 'user@example.com';
  const h = load({ sheets: { 'Custom Rules': [
    ['UserEmail', 'Category', 'Keyword', 'Timestamp'],
    [email, 'Exclude', '', ''],              // blank keyword: used to capture EVERYTHING
    [email, 'Exclude', 1111, ''],            // numeric: used to throw on every parse
    [email, 'Rent', 'Landlord Name', ''],    // v1.1.70: an old one-to-one name — read as its new one
    [email, '', 'ORPHAN KEYWORD', ''],       // no category
    ['someone-else@example.com', 'Exclude', 'UBER', '']
  ] } });
  const rules = h.plain(h.ctx.getUserCustomRules(email));
  assert.deepEqual(rules, { Exclude: ['1111'], 'Rent & housing': ['LANDLORD NAME'] });

  const raw = h.ctx.getUserCustomRules(email);
  assert.equal(h.ctx.findCustomRuleOverride('UBER*RIDES SANTO DOMINGO', raw), '', 'blank keyword no longer matches everything');
  assert.equal(h.ctx.findCustomRuleOverride('BHD Transfer (to account ...1111)', raw), 'Exclude');
  assert.equal(h.ctx.findCustomRuleOverride('LANDLORD NAME', raw), 'Rent & housing');
});

test('custom-only categories are listed for the Dashboard, Exclude is not (M8)', () => {
  const email = 'user@example.com';
  const h = load({ sheets: { 'Custom Rules': [
    ['UserEmail', 'Category', 'Keyword', 'Timestamp'],
    [email, 'Kids', 'TOYSRUS', ''],
    [email, 'Pets', 'PETSMART', ''],         // v1.1.70: a default category now
    [email, 'Rent', 'LANDLORD', ''],
    [email, 'exclude', 'MY OWN NAME', ''],
    [email, 'Investments & savings', 'MY BROKER', '']   // v1.1.70: reserved, like Exclude
  ] } });
  assert.deepEqual(h.plain(h.ctx.getCustomCategoryNames(email)), ['Kids']);
});
