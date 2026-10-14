/**
 * CATEGORIZER MODULE
 * Handles transaction categorization with extensible rules
 */

/**
 * v1.1.70: the categories were rebuilt — 25 categories in 8 groups (CATEGORY_GROUPS), agreed by both users after
 * looking at a year of real rows. Why:
 *  - 'Dining/Delivery + Entertainment + Other' was the fallback (v1.1.5), so every merchant the tracker didn't know
 *    landed there, next to restaurants: in one user's file it held half of all spending (shopping, a whole trip,
 *    delivery, bars…). The fallback is now 'Unclassified', so what isn't recognized stays visible.
 *  - the combined categories are split where they hid different decisions: food (groceries · restaurants · delivery),
 *    mobility (ride apps · public transport · fuel), care (health · personal care · pets · gym) and lifestyle.
 *  - each category says whether it is FIXED (a bill that comes every month) — the Dashboard, the daily and the monthly
 *    summary read it from here instead of two lists of their own that had drifted apart (Electricity was fixed in the
 *    summaries but not on the Dashboard).
 *  - 'Travel' and 'Payments to people' have no keywords: they are chosen by hand or with a Custom Rule.
 * The keywords of each old category moved to the new one they belong to; categories saved with an old name are moved
 * by Recategorize (LEGACY_CATEGORIES below). Categories from before (v1.1.4 → v1.1.69): Rent, Gym + Calisthenics,
 * Telecommunications, Streaming & Subscriptions, Electricity, Groceries + Barbershop,
 * Dining/Delivery + Entertainment + Other, Vehicle Gas, Health + Vet + Pharmacy, Transportation.
 *
 * Names are in English (v1.1.4); keywords stay in Spanish — they match Spanish bank-email text.
 * Each category: group, fixed, keywords, icon, color = chip [background, text] in the data sheets (v1.1.28).
 * DEFAULT_CATEGORIES is in DISPLAY order (by group); the order keywords are tried in is CATEGORY_MATCH_ORDER.
 */
const CATEGORY_GROUPS = ['Housing', 'Food', 'Mobility', 'Health & care', 'Lifestyle', 'Travel', 'Bank & debt', 'Other'];

const DEFAULT_CATEGORIES = {
  // ---- Housing
  'Rent & housing': {
    // v1.1.19: removed bare 'RENT' and 'RENTA' — substring matches sent
    // "AUTO RENT A CAR" (verified) and similar to Rent, a FIXED category.
    // A landlord is matched by a Custom Rule (e.g. the landlord's name),
    // which is how real rent payments have been categorized all along.
    // v1.1.70: repairs are NOT here (a one-off in a fixed category inflates the month-end pace) — Shopping & home.
    group: 'Housing', fixed: true,
    keywords: ['ALQUILER', 'ARRENDAMIENTO', 'CONDOMINIO'],
    icon: '🏠', color: ['#EDE9E4', '#57534E']
  },
  'Utilities': {
    // v1.1.19: removed bare 'LUZ' — "LUZ DE LUNA CAFE" landed here
    // (verified). The three distributors below cover every real bill.
    // v1.1.70: water too (CAASD, CORAASAN, INAPA — public utilities). Laundry is not a bill: Shopping & home.
    group: 'Housing', fixed: true,
    keywords: ['EDESUR', 'EDENORTE', 'EDEESTE', 'ELECTRICIDAD', 'CAASD', 'CORAASAN', 'INAPA'],
    icon: '💡', color: ['#FEF9C3', '#854D0E']
  },
  'Phone & internet': {
    // v1.1.5: merged Internet + Mobile Data. 'CLARO' bare (not just
    // 'CLARO MOVIL') so "CLARO P REC" (Claro Prepago Recarga, a real
    // merchant descriptor) matches too.
    group: 'Housing', fixed: true,
    keywords: ['INTERNET', 'WIFI', 'BANDA ANCHA', 'CLARO INTERNET',
               'DATOS MOVILES', 'RECARGA', 'PLAN MOVIL', 'CLARO MOVIL', 'CLARO', 'PREPAGO', 'ALTICE'],
    icon: '📶', color: ['#F3E8FF', '#7E22CE']
  },
  // ---- Food
  'Groceries': {
    // v1.1.23: 'POLA' — Supermercados Pola ("SM POLA INDEPENDENCIA"); whole word
    group: 'Food', fixed: false,
    keywords: ['SUPER', 'CARREFOUR', 'JUMBO', 'BRAVO', 'MERCADO', 'FRUTAS',
               'SUPERMERCADOS NACIONAL', 'SM NACIONAL', 'LA SIRENA', 'PLAZA LAMA', 'IBERIA',
               'APREZIO', 'LA FUENTE', 'PRICESMART', 'HIPER OLE', '365 EL CACIQUE', 'HOLA PLAZA', 'POLA',
               'COLMADO', 'PANADERIA', 'MINIMARKET', 'MINI MARKET'],
    icon: '🛒', color: ['#FFF4E5', '#B45309']
  },
  'Restaurants & cafés': {
    // v1.1.15: 'BONJOUR' is matched before Fuel & vehicle (see CATEGORY_MATCH_ORDER) so "BONJOUR TOTAL ARENOSO" (the
    // convenience store) never reaches Fuel's bare 'TOTAL'.
    // v1.1.5: Krispy Kreme and IKEA Restaurant from real data ('IKEA REST' before Shopping's 'IKEA').
    group: 'Food', fixed: false,
    keywords: ['RESTAURANTE', 'RESTAURANT', 'PIZZERIA', 'BURGER', 'COMIDA', 'FOOD', 'CAFE', 'SUSHI', 'POLLO',
               'HELADOS BON', 'BONJOUR', 'MCDONALD', 'BURGER KING', 'KFC', 'PIZZA HUT', 'SUBWAY', 'POLLOS VICTORINA',
               'COMEDOR', 'WENDYS', 'JADE TERIYAKI', 'SWEETFROG', 'SWEET FROG', 'KRISPY KREME', 'IKEA REST',
               'CAFETERIA', 'HELADERIA', 'STARBUCKS'],
    icon: '🍽️', color: ['#FFE4E6', '#BE123C']
  },
  'Delivery': {
    // v1.1.17: bare 'UBER EATS' (space, no asterisk) — "UBER EATS-W*UBER EATS- SANTO DOMINGO DOM" matched neither
    // 'UBER*EATS' nor 'UBEREATS' and fell through to the rides' bare 'UBER'. Delivery is matched before the rides.
    // v1.1.5: every PedidosYa* order is food. v1.1.70: 'DIDI FOOD' here (it went to Dining through 'FOOD' before).
    group: 'Food', fixed: false,
    keywords: ['UBER*EATS', 'UBEREATS', 'UBER EATS', 'PEDIDOSYA', 'DIDI FOOD', 'GLOVO', 'DELIVERY'],
    icon: '🛵', color: ['#FFEDD5', '#9A3412']
  },
  // ---- Mobility
  'Ride apps & taxis': {
    // v1.1.20: 'DIDI' — real POPULAR rows in COP ("DiDi CO Ride", "DL*DIDI RIDES", "DLO*Didi"). Whole word (4 letters).
    // v1.1.23: Uber rides billed through PayPal ("PAYPAL *UBERBV", "PAYPAL *UBER BV")
    group: 'Mobility', fixed: false,
    keywords: ['UBER', 'UBER*RIDES', 'UBER*TRIP', 'DIDI', 'TAXI', 'MOTOCONCHO', 'UBERBV', 'UBER BV', 'CABIFY', 'INDRIVE'],
    icon: '🚕', color: ['#EAF1FE', '#1D4ED8']
  },
  'Public transport & parking': {
    // v1.1.23: 'METRO' alone is transit; the Metro Plaza mall ("ALISS METRO PZA") is matched first by Shopping & home
    group: 'Mobility', fixed: false,
    keywords: ['BUS', 'METRO', 'PARKING', 'PARQUEO', 'GUAGUA', 'TELEFERICO', 'OMSA'],
    icon: '🚇', color: ['#E0F2FE', '#075985']
  },
  'Fuel & vehicle': {
    // v1.1.15: bare 'TOTAL' and 'NEXT LINCOLN' (a gas-station brand, "NEXT LINCOLN GASOPOLIS" from LAFISE, bare "NEXT
    // LINCOLN" from POPULAR). 'TOTAL' is safe because Restaurants & cafés ('BONJOUR') is matched first.
    group: 'Mobility', fixed: false,
    keywords: ['SHELL', 'HESS', 'PUMA', 'TEXACO', 'GASOLINA', 'COMBUSTIBLE', 'FUEL',
               'TOTALENERGIES', 'TOTAL', 'NEXT LINCOLN', 'GASOLINERA', 'PEAJE', 'PASO RAPIDO'],
    icon: '⛽', color: ['#E2E8F0', '#334155']
  },
  // ---- Health & care
  'Health & pharmacy': {
    // v1.1.23: 'FARMA' (FARMA VALUE…), 'FARM' whole word (FARM CAROL…), labs and clinics (ANALISA, CTRO ESPEC MED)
    group: 'Health & care', fixed: false,
    keywords: ['FARMACIA', 'FCIA', 'DOCTOR', 'MEDICO', 'HOSPITAL', 'CLINICA', 'SALUD',
               'MEDICINAS', 'MEDICAR', 'GBC', 'LOS HIDALGOS',
               'FARMAX', 'FARMAXTRA', 'CRUZ VERDE', 'FARMATODO', 'FARMAVALUE', 'FARMA', 'FARM', 'ANALISA', 'LABORATORIO',
               'ESPEC MED', 'ODONTO', 'DENTAL', 'OPTICA'],
    icon: '🏥', color: ['#E8F5E9', '#2E7D32']
  },
  'Personal care': {
    group: 'Health & care', fixed: false,
    keywords: ['BARBERIA', 'BARBER', 'SALON', 'PELUQUERIA', 'SPA'],
    icon: '💈', color: ['#FDF4FF', '#86198F']
  },
  'Pets': {
    // v1.1.18: 'LA CASITA DE' (a vet clinic) — POPULAR cuts the name short, so the shorter keyword covers every bank.
    // Matched before Health & pharmacy, so "CLINICA VETERINARIA" is a vet, not a clinic.
    group: 'Health & care', fixed: false,
    keywords: ['VETERINARIA', 'VET', 'LA CASITA DE', 'PET SHOP', 'PETSHOP', 'MASCOTA'],
    icon: '🐾', color: ['#ECFCCB', '#3F6212']
  },
  'Gym': {
    // v1.1.70: the membership and the trainer only — sports gear is Shopping & home (a one-off in a fixed category
    // inflates the month-end pace of the daily summary)
    group: 'Health & care', fixed: true,
    keywords: ['GYM', 'GIMNASIO', 'CALISTENIA', 'FITNESS', 'CROSSFIT', 'PILATES', 'YOGA', 'ENTRENADOR'],
    icon: '💪', color: ['#CCFBF1', '#0F766E']
  },
  // ---- Lifestyle
  'Subscriptions': {
    // v1.1.15: 'AMAZON PRIME' only — bare 'AMAZON' is a purchase (Shopping & home), matched after this.
    // v1.1.23: 'GOOGLE' — real POPULAR USD rows "Google" / "GOOGLE *Google" (Google One/Play)
    group: 'Lifestyle', fixed: true,
    keywords: ['NETFLIX', 'SPOTIFY', 'DISNEY', 'HBO', 'AMAZON PRIME', 'PRIME VIDEO', 'SUSCRIPCION',
               'YOUTUBE PREMIUM', 'OPENAI', 'ANTHROPIC', 'CLAUDE.AI', 'GOOGLE', 'ICLOUD', 'APPLE.COM/BILL', 'PARAMOUNT'],
    icon: '🎬', color: ['#E0E7FF', '#3730A3']
  },
  'Shopping & home': {
    // v1.1.5 (Clothing's keywords), v1.1.15 (bare 'AMAZON'), v1.1.23 (ALISS and the Metro Plaza mall — matched before
    // Public transport's 'METRO'). v1.1.70: home and hardware stores too; repairs and sports gear belong here.
    group: 'Lifestyle', fixed: false,
    keywords: ['AMAZON', 'ROPA', 'ZAPATOS', 'ADIDAS', 'NIKE', 'PUMA STORE', 'ZARA', 'FASHION', 'BOUTIQUE', 'TIENDA',
               'ALISS', 'METRO PZA', 'METRO PLZA', 'METRO PLAZA', 'IKEA', 'H&M', 'SHEIN', 'TEMU', 'ALIEXPRESS', 'EBAY',
               'FERRETERIA', 'CORRIPIO', 'LAVANDERIA', 'DEPORTES'],
    icon: '🛍️', color: ['#FDF2F8', '#9D174D']
  },
  'Leisure & going out': {
    group: 'Lifestyle', fixed: false,
    keywords: ['CINE', 'CINEMA', 'PELICULAS', 'CONCIERTO', 'ENTRETENIMIENTO', 'BAR', 'PUB', 'TEATRO', 'DISCOTECA',
               'CERVECERIA', 'BREWERY', 'TICKETMASTER'],
    icon: '🎉', color: ['#EDE9FE', '#5B21B6']
  },
  'Education & books': {
    // v1.1.24 removed 'Education' (unused then); v1.1.70 brings it back as a default. A Custom Rule named "Education"
    // is moved to this name by Recategorize.
    group: 'Lifestyle', fixed: false,
    keywords: ['COURSERA', 'UDEMY', 'EDX', 'DUOLINGO', 'LIBRERIA', 'BOOKSTORE', 'UNIVERSIDAD', 'COLEGIO', 'ACADEMIA'],
    icon: '📚', color: ['#F0FDFA', '#115E59']
  },
  // ---- Travel
  'Travel': {
    // v1.1.70: chosen by hand or with a Custom Rule (agreed): a trip's purchases keep the category of what they are.
    group: 'Travel', fixed: false,
    keywords: [],
    icon: '✈️', color: ['#ECFEFF', '#155E75']
  },
  // ---- Bank & debt
  'Bank fees & taxes': {
    // v1.1.70: a transfer's own tax-and-commission row ("BDI — impuesto y comisión de transferencia") lands here by
    // itself (autoTransferCategory); these keywords are for card charges
    group: 'Bank & debt', fixed: false,
    keywords: ['COMISION', 'IMPUESTO', 'CARGO POR', 'INTERESES', 'MEMBRESIA ANUAL', 'CARGO ANUAL'],
    icon: '🏦', color: ['#F1F5F9', '#475569']
  },
  'Loan payments': {
    // v1.1.70: "Pago de préstamo …" transfers (Scotiabank, Banreservas — v1.1.69) land here by themselves
    group: 'Bank & debt', fixed: true,
    keywords: [],
    icon: '📄', color: ['#FEE2E2', '#991B1B']
  },
  'Insurance': {
    group: 'Bank & debt', fixed: true,
    keywords: ['SEGURO', 'ASEGURADORA', 'POLIZA'],
    icon: '🛡️', color: ['#EEF2FF', '#4338CA']
  },
  // ---- Other
  'Cash': {
    // v1.1.70: its own category (agreed); a Custom Rule moves it for whoever prefers Leisure & going out
    group: 'Other', fixed: false,
    keywords: ['CASH WITHDRAWAL', 'CODIGO CASH', 'CÓDIGO CASH', 'CAJERO', 'ATM'],
    icon: '💵', color: ['#DCFCE7', '#166534']
  },
  'Government & procedures': {
    group: 'Other', fixed: false,
    keywords: ['DGII', 'JCE', 'PASAPORTE', 'INTRANT', 'MIGRACION', 'MINISTERIO', 'AYUNTAMIENTO', 'MARBETE', 'EMBAJADA'],
    icon: '🏛️', color: ['#F5F5F4', '#44403C']
  },
  'Payments to people': {
    // v1.1.70: a transfer to a person that fits no other category — chosen by hand or with a Custom Rule on the name.
    // Transfers are never put here by themselves: one without a category stays in "Transfers to review".
    group: 'Other', fixed: false,
    keywords: [],
    icon: '🤝', color: ['#F0FDF4', '#15803D']
  },
  'Unclassified': {
    // v1.1.70: the fallback — a merchant no keyword knows. Add a Custom Rule (or pick a category) to move it.
    group: 'Other', fixed: false,
    keywords: [],
    icon: '❓', color: ['#FEF2F2', '#B91C1C']
  }
};

/**
 * v1.1.70: the order keywords are tried in — first match wins. It is NOT the display order: several past fixes depend
 * on one category being tried before another, and those pairs now sit in different groups:
 *   Delivery ('UBER EATS', 'DIDI FOOD') before Ride apps ('UBER', 'DIDI');
 *   Restaurants ('BONJOUR' — a store inside a gas station) before Fuel ('TOTAL');
 *   Shopping ('METRO PLAZA' — a mall) before Public transport ('METRO');
 *   Subscriptions ('AMAZON PRIME') before Shopping ('AMAZON');  Restaurants ('IKEA REST') before Shopping ('IKEA');
 *   Pets ('VETERINARIA') before Health ('CLINICA');  Shopping ('PUMA STORE') before Fuel ('PUMA').
 * The rest keeps the relative order the old categories had. Every category with keywords is listed once (tested).
 */
const CATEGORY_MATCH_ORDER = ['Rent & housing', 'Gym', 'Phone & internet', 'Subscriptions', 'Utilities', 'Groceries',
  'Personal care', 'Delivery', 'Restaurants & cafés', 'Leisure & going out', 'Shopping & home', 'Fuel & vehicle', 'Pets',
  'Health & pharmacy', 'Ride apps & taxis', 'Public transport & parking', 'Education & books', 'Insurance',
  'Government & procedures', 'Cash', 'Bank fees & taxes'];

const FALLBACK_CATEGORY = 'Unclassified';
// v1.1.70: money that moves but isn't spent, apart from Exclude: deposits to a broker (Investment Accounts' deposit
// keyword). Like Exclude, it is never a spending row; the Dashboard shows it among the movements.
const INVESTMENT_CATEGORY = 'Investments & savings';

/**
 * v1.1.70: categories saved before this version → the new ones. A one-to-one name moves as it is. A name that merged
 * several kinds of spending moves to the new category the keywords find, when that is one of its own (`within`), or
 * else to its main one (`to`): a hand-set "Groceries + Barbershop" on a barbershop stays a barbershop.
 */
const LEGACY_CATEGORIES = {
  'Rent': { to: 'Rent & housing' },
  'Gym + Calisthenics': { to: 'Gym' },
  'Telecommunications': { to: 'Phone & internet' },
  'Streaming & Subscriptions': { to: 'Subscriptions' },
  'Electricity': { to: 'Utilities' },
  'Vehicle Gas': { to: 'Fuel & vehicle' },
  'Education': { to: 'Education & books' },
  'Groceries + Barbershop': { to: 'Groceries', within: ['Groceries', 'Personal care'] },
  'Health + Vet + Pharmacy': { to: 'Health & pharmacy', within: ['Health & pharmacy', 'Pets'] },
  'Transportation': { to: 'Ride apps & taxis', within: ['Ride apps & taxis', 'Public transport & parking', 'Fuel & vehicle'] },
  'Dining/Delivery + Entertainment + Other': { to: 'Unclassified',
    within: ['Restaurants & cafés', 'Delivery', 'Leisure & going out', 'Shopping & home'] }
};

/**
 * v1.1.19: SHARED KEYWORD ENGINE — used by categorization, Custom Rules,
 * type detection and the promotional/non-transactional/declined filters
 * (03_gmailMonitor.gs), so every keyword list follows one matching rule.
 *
 * Why: plain substring matching caused a whole class of bugs — 'ACH' matched
 * inside "CACHAREPA", 'RENT' inside "AUTO RENT A CAR", 'ROPA' inside
 * "EUROPA", 'BUS' inside "BUSINESS". Rule now:
 *   - keywords of 4 or fewer letters/digits (BAR, BUS, VET, ACH, OTP, UBER…)
 *     must match as a WHOLE WORD — the characters on each side must not be
 *     a letter or digit ('*', '-', space, start/end of text all count as
 *     boundaries, so "UBER*RIDES" still matches 'UBER');
 *   - longer keywords keep SUBSTRING matching, which several rely on on
 *     purpose ('SUPER' → SUPERMERCADO, 'FARMAX' → FARMAXTRA, 'TOTAL' →
 *     "TOTALENERGIES"-style descriptors).
 * Regexes are compiled once and cached per keyword.
 */
const KEYWORD_WORD_CHARS = 'A-Z0-9ÁÉÍÓÚÑÜ';
const KEYWORD_WHOLE_WORD_MAX_LEN = 4;
const _keywordRegexCache = {};

/**
 * Uppercase + v1.1.23 spacing rules, applied to BOTH the text and every
 * keyword: spaces around '*' are removed ("UBER * EATS" → "UBER*EATS" — real
 * LAFISE rows fell to Transportation because of the spaces) and runs of
 * whitespace collapse to one space.
 */
function normalizeKeyword(keyword) {
  return String(keyword === null || keyword === undefined ? '' : keyword)
    .toUpperCase().replace(/\s*\*\s*/g, '*').replace(/\s+/g, ' ').trim();
}

function keywordMatches(upperText, keyword) {
  const kw = normalizeKeyword(keyword);
  if (!kw || !upperText) return false;       // an empty keyword never matches (see C4 fix below)
  const coreLen = kw.replace(new RegExp('[^' + KEYWORD_WORD_CHARS + ']', 'g'), '').length;
  if (coreLen > KEYWORD_WHOLE_WORD_MAX_LEN) return upperText.includes(kw);
  let re = _keywordRegexCache[kw];
  if (!re) {
    const esc = kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    re = new RegExp('(^|[^' + KEYWORD_WORD_CHARS + '])' + esc + '(?=$|[^' + KEYWORD_WORD_CHARS + '])');
    _keywordRegexCache[kw] = re;
  }
  return re.test(upperText);
}

/** Returns the first keyword in `keywords` that matches, or null. */
function findMatchingKeyword(upperText, keywords) {
  for (const kw of keywords || []) {
    if (keywordMatches(upperText, kw)) return kw;
  }
  return null;
}

/**
 * Main categorization function — DEFAULT_CATEGORIES only, evaluated in
 * CATEGORY_MATCH_ORDER (v1.1.70; declaration order before — order matters, see the notes above it).
 * v1.1.5: fallback changed from 'Other' to FALLBACK_CATEGORY ('Unclassified' since v1.1.70).
 * v1.1.19: the old `customRules` (merged defaults + Custom Rules) parameter
 * is gone. Every caller already runs findCustomRuleOverride() FIRST, so the
 * merged rules could only ever be reached when no custom rule matched — at
 * which point they were identical to the defaults. Also uses the shared
 * keyword engine above instead of raw substring matching.
 */
function categorizeTransaction(text) {
  if (!text) return FALLBACK_CATEGORY;
  const upperText = normalizeKeyword(text);
  for (const category of CATEGORY_MATCH_ORDER) {
    if (findMatchingKeyword(upperText, DEFAULT_CATEGORIES[category].keywords)) {
      return category;
    }
  }
  return FALLBACK_CATEGORY;
}

/**
 * v1.1.70: a category for a Transfer the tracker can name by itself (transfers otherwise only get one from a Custom
 * Rule or by hand): a bank's own tax-and-commission row, a loan installment, or a deposit to a broker listed in
 * Investment Accounts (`investmentKeywords`: their deposit keywords). '' when it is none of them.
 */
const TRANSFER_FEE_RE = /\bimpuestos?\s+y\s+comisi[óo]n/i;
const LOAN_PAYMENT_RE = /^\s*pago de pr[ée]stamo\b/i;
function autoTransferCategory(text, investmentKeywords) {
  const t = String(text || '');
  const upper = t.toUpperCase();
  if ((investmentKeywords || []).some(k => k && upper.indexOf(String(k).toUpperCase()) !== -1)) return INVESTMENT_CATEGORY;
  if (TRANSFER_FEE_RE.test(t)) return 'Bank fees & taxes';
  if (LOAN_PAYMENT_RE.test(t)) return 'Loan payments';
  return '';
}

/**
 * v1.1.5: checks ONLY the user's Custom Rules (not the defaults) against
 * `text`, regardless of transaction Type. Used so a personal rule — e.g.
 * "<LANDLORD NAME>" → Rent for a landlord's transfer — can override the normal
 * "only Type=Transaction gets a category" rule (03_gmailMonitor.gs /
 * 04_sheetsWriter.gs both call this BEFORE deciding whether to categorize
 * at all). Returns the matched category, or '' if no custom rule matches.
 */
function findCustomRuleOverride(text, rawCustomRules) {
  // v1.1.19: String() + shared engine — a numeric keyword (e.g. 1111, an
  // account's last 4 digits, which Sheets stores as a number) used to throw
  // "toUpperCase is not a function" here.
  if (!text || !rawCustomRules) return '';
  const upperText = normalizeKeyword(text);
  for (let category in rawCustomRules) {
    if (findMatchingKeyword(upperText, rawCustomRules[category])) return category;
  }
  return '';
}

/**
 * Get all available categories
 */
function getCategories() {
  return Object.keys(DEFAULT_CATEGORIES);
}

/** v1.1.70: the group a category is shown under — a category of your own goes under 'Other'. */
function categoryGroup(category) {
  const c = DEFAULT_CATEGORIES[category];
  return c ? c.group : 'Other';
}

/** v1.1.70: fixed categories (a bill every month), read by the Dashboard and both summaries — one list, from here. */
function fixedCategoryNames() {
  return getCategories().filter(c => DEFAULT_CATEGORIES[c].fixed);
}
function isFixedCategory(category) {
  return !!(DEFAULT_CATEGORIES[category] && DEFAULT_CATEGORIES[category].fixed);
}

/** v1.1.70: categories that are never spending: Exclude and Investments & savings (case and spaces ignored). */
function isReservedCategory(category) {
  const c = String(category === null || category === undefined ? '' : category).trim().toUpperCase();
  return c === EXCLUDE_CATEGORY.toUpperCase() || c === INVESTMENT_CATEGORY.toUpperCase();
}
/** v1.1.70: a row with this category counts as spending (given a spending Type): set, and not a reserved one. */
function isSpendingCategory(category) {
  return String(category === null || category === undefined ? '' : category).trim() !== '' && !isReservedCategory(category);
}

/**
 * v1.1.70: the category a saved row gets for a name from before v1.1.70 (LEGACY_CATEGORIES). `computed` is what the
 * keywords give the row now. Returns '' for a name that isn't an old one.
 */
function legacyCategoryFor(name, computed) {
  const old = LEGACY_CATEGORIES[String(name === null || name === undefined ? '' : name).trim()];
  if (!old) return '';
  return old.within && old.within.indexOf(computed) !== -1 ? computed : old.to;
}
/** v1.1.70: old names that merged several kinds of spending — a Custom Rule with one needs a decision (Start here). */
function isMergedLegacyCategory(name) {
  const old = LEGACY_CATEGORIES[String(name === null || name === undefined ? '' : name).trim()];
  return !!(old && old.within);
}

/**
 * Get category with icon
 */
function getCategoryWithIcon(category) {
  if (DEFAULT_CATEGORIES[category]) {
    return DEFAULT_CATEGORIES[category].icon + ' ' + category;
  }
  return '📌 ' + category;
}

/**
 * Add custom rule to user config
 */
function addCustomCategoryRule(userEmail, category, keyword) {
  try {
    const sheet = getOrCreateSheet(CUSTOM_RULES_SHEET);
    sheet.appendRow([
      userEmail,
      category,
      normalizeKeyword(keyword),
      new Date().toISOString()
    ]);
    return true;
  } catch (error) {
    Logger.log("Error adding custom rule: " + error);
    return false;
  }
}

/**
 * Get custom rules for user
 * v1.1.6: email comparison now trims whitespace and ignores case — a
 * mismatched email (extra space, different case) between what's typed into
 * the Custom Rules sheet and what Configuration stores would otherwise
 * silently return zero rules, which is one of the ways the landlord/Rent
 * override could fail to apply even with the right keyword.
 */
function getUserCustomRules(userEmail) {
  try {
    const sheet = getOrCreateSheet(CUSTOM_RULES_SHEET);
    const data = sheet.getDataRange().getValues();
    const customRules = {};
    const targetEmail = String(userEmail).trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][0] || '').trim().toLowerCase();
      if (rowEmail === targetEmail) {
        // v1.1.19: BUG FIX (C4) — values used to be pushed raw. A blank
        // keyword cell became '' and `text.includes('')` is ALWAYS true, so
        // that one rule captured every transaction; a numeric keyword (1111)
        // threw inside every parse and every recategorize, silently dropping
        // every email of the run. Incomplete rows are now skipped.
        let category = String(data[i][1] === null || data[i][1] === undefined ? '' : data[i][1]).trim();
        const keyword = normalizeKeyword(data[i][2]);
        if (!category || !keyword) continue;
        // v1.1.70: a rule still on an old one-to-one name (Rent, Electricity…) works with the new name. Recategorize
        // also renames it in the sheet; an old name that merged several categories is kept until you choose.
        const renamed = LEGACY_CATEGORIES[category] && !LEGACY_CATEGORIES[category].within ? LEGACY_CATEGORIES[category].to : '';
        if (renamed) category = renamed;

        if (!customRules[category]) {
          customRules[category] = [];
        }
        customRules[category].push(keyword);
      }
    }
    
    return customRules;
  } catch (error) {
    Logger.log("Error reading custom rules: " + error);
    return {};
  }
}

/**
 * v1.1.70: Custom Rules whose category is an old one-to-one name (Rent → Rent & housing…) get the new name in the
 * sheet, so what you read there is what applies. An old name that merged several categories is left as it is: it
 * keeps working as a category of your own, and 📘 Start here asks you to choose its new one. Returns how many changed.
 */
function renameLegacyCustomRules() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CUSTOM_RULES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const range = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1);
  let changed = 0;
  const values = range.getValues().map(r => {
    const old = LEGACY_CATEGORIES[String(r[0] === null || r[0] === undefined ? '' : r[0]).trim()];
    if (!old || old.within) return [r[0]];
    changed++;
    return [old.to];
  });
  if (changed) {
    range.setValues(values);
    Logger.log('🏷️ Custom Rules moved to the v1.1.70 category names: ' + changed);
  }
  return changed;
}

const EXCLUDE_CATEGORY = 'Exclude';

/**
 * v1.1.28: chip colours for categories that aren't in DEFAULT_CATEGORIES —
 * "Exclude" (muted: it's left out of every total) and your own Custom Rules
 * categories, which take the next colour of this rotation in order.
 */
const EXCLUDE_COLOR = ['#F3F4F6', '#4B5563'];   // #6B7280 was 4.39:1 — below WCAG AA (caught by v128 test)
const INVESTMENT_COLOR = ['#ECFDF5', '#065F46'];   // v1.1.70
const CUSTOM_CATEGORY_COLORS = [
  ['#CFFAFE', '#0E7490'], ['#FAE8FF', '#A21CAF'], ['#D1FAE5', '#047857'], ['#FEF3C7', '#92400E'], ['#FCE7F3', '#9D174D']
];

/** Every category with its chip colours, in display order: defaults, Exclude, Investments & savings, then custom ones. */
function categoryPalette(customNames) {
  const list = getCategories().map(name => ({ name: name, bg: DEFAULT_CATEGORIES[name].color[0], fg: DEFAULT_CATEGORIES[name].color[1] }));
  list.push({ name: EXCLUDE_CATEGORY, bg: EXCLUDE_COLOR[0], fg: EXCLUDE_COLOR[1] });
  list.push({ name: INVESTMENT_CATEGORY, bg: INVESTMENT_COLOR[0], fg: INVESTMENT_COLOR[1] });   // v1.1.70
  (customNames || []).filter(n => n && !DEFAULT_CATEGORIES[n] && !isReservedCategory(n))
    .forEach((name, i) => {
      const c = CUSTOM_CATEGORY_COLORS[i % CUSTOM_CATEGORY_COLORS.length];
      list.push({ name: name, bg: c[0], fg: c[1] });
    });
  return list;
}

/**
 * v1.1.19: categories that exist ONLY in Custom Rules (not in
 * DEFAULT_CATEGORIES, and not the reserved "Exclude"). The Dashboard grid
 * used to be built from the defaults alone, so spend in a custom category
 * (say "Kids") never reached any row or the TOTAL — see
 * buildOrRefreshDashboard(). getMergedCategoryRules() was removed in this
 * version (see categorizeTransaction()).
 */
function getCustomCategoryNames(userEmail) {
  const custom = getUserCustomRules(userEmail);
  return Object.keys(custom).filter(c => !DEFAULT_CATEGORIES[c] && !isReservedCategory(c));   // v1.1.70: reserved ones
}
