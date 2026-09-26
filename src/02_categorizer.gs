/**
 * CATEGORIZER MODULE
 * Handles transaction categorization with extensible rules
 */

/**
 * v1.1.5: further category consolidation, per direct request:
 * - 'Housing' → 'Rent' (clearer/shorter).
 * - 'Internet' + 'Mobile Data' merged into one 'Telecommunications' category
 *   (both keyword lists combined).
 * - 'Clothing' and 'Other' removed as separate categories — everything that
 *   would have landed in either now falls into
 *   'Dining/Delivery + Entertainment + Other' instead (Clothing's keywords
 *   were folded in, and the categorizeTransaction() fallback below now
 *   returns that label instead of 'Other'). Note the side effect: a
 *   genuinely unmatched merchant will now show as Dining rather than a
 *   distinct "uncategorized" bucket — that's the tradeoff of this merge.
 * - Added real merchants seen in live data that were falling through to the
 *   old 'Other' bucket: PedidosYa (a delivery app — every PedidosYa* order
 *   is food), Krispy Kreme, IKEA Restaurant.
 * Now 11 categories total (was 13).
 */

/**
 * v1.1.4: category names translated to English (were the only Spanish-named
 * part of an otherwise-English system, e.g. "Vivienda" next to an English
 * "Category" header). Keywords themselves stay in Spanish — they're
 * matching against Spanish bank-email text, which didn't change.
 */
const DEFAULT_CATEGORIES = {
  'Rent': {
    // v1.1.19: removed bare 'RENT' and 'RENTA' — substring matches sent
    // "AUTO RENT A CAR" (verified) and similar to Rent, a FIXED category.
    // A landlord is matched by a Custom Rule (e.g. the landlord's name),
    // which is how real rent payments have been categorized all along.
    keywords: ['ALQUILER', 'ARRENDAMIENTO'],
    icon: '🏠',
    color: ['#EDE9E4', '#57534E']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Gym + Calisthenics': {
    keywords: ['GYM', 'GIMNASIO', 'CALISTENIA', 'FITNESS', 'CROSSFIT', 'PILATES', 'YOGA', 'ENTRENADOR'],
    icon: '💪',
    color: ['#CCFBF1', '#0F766E']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Telecommunications': {
    // v1.1.5: merged Internet + Mobile Data. 'CLARO' bare (not just
    // 'CLARO MOVIL') so "CLARO P REC" (Claro Prepago Recarga, a real
    // merchant descriptor) matches too.
    keywords: ['INTERNET', 'WIFI', 'BANDA ANCHA', 'CLARO INTERNET',
               'DATOS MOVILES', 'RECARGA', 'PLAN MOVIL', 'CLARO MOVIL', 'CLARO', 'PREPAGO'],
    icon: '📶',
    color: ['#F3E8FF', '#7E22CE']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Streaming & Subscriptions': {
    // v1.1.15: BUG FIX — bare 'AMAZON' here caught EVERY Amazon purchase
    // (Amazon.com retail orders, Amazon Marketplace), not just the Prime
    // subscription itself. Narrowed to 'AMAZON PRIME' only; general Amazon
    // purchases now fall through to Dining's bare 'AMAZON' keyword instead
    // (Streaming is checked first, so 'AMAZON PRIME' still correctly
    // intercepts the subscription case before Dining's broader keyword
    // would otherwise catch it too).
    // v1.1.23: 'GOOGLE' — real POPULAR USD rows "Google" / "GOOGLE *Google" (Google One/Play)
    keywords: ['NETFLIX', 'SPOTIFY', 'DISNEY', 'HBO', 'AMAZON PRIME', 'SUSCRIPCION',
               'YOUTUBE PREMIUM', 'OPENAI', 'GOOGLE'],
    icon: '🎬',
    color: ['#E0E7FF', '#3730A3']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Electricity': {
    // v1.1.19: removed bare 'LUZ' — "LUZ DE LUNA CAFE" landed here
    // (verified). The three distributors below cover every real bill.
    keywords: ['EDESUR', 'EDENORTE', 'EDEESTE', 'ELECTRICIDAD'],
    icon: '💡',
    color: ['#FEF9C3', '#854D0E']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Groceries + Barbershop': {
    // v1.1.23: 'POLA' — Supermercados Pola ("SM POLA INDEPENDENCIA"); whole word
    keywords: ['SUPER', 'CARREFOUR', 'JUMBO', 'BRAVO', 'MERCADO', 'FRUTAS', 'BARBERIA', 'SALON',
               'SUPERMERCADOS NACIONAL', 'SM NACIONAL', 'LA SIRENA', 'PLAZA LAMA', 'IBERIA',
               'APREZIO', 'LA FUENTE', 'PRICESMART', 'HIPER OLE', '365 EL CACIQUE', 'HOLA PLAZA', 'POLA'],
    icon: '🛒',
    color: ['#FFF4E5', '#B45309']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Dining/Delivery + Entertainment + Other': {
    // v1.1.5: absorbed Clothing's keywords and the old 'Other' fallback role
    // (see header comment), plus PedidosYa/Krispy Kreme/IKEA Restaurant from
    // real transaction data.
    // v1.1.17: added bare 'UBER EATS' (space, no asterisk) — a real
    // merchant descriptor "UBER EATS-W*UBER EATS- SANTO DOMINGO DOM" didn't
    // match 'UBER*EATS' or 'UBEREATS' and fell through all the way to
    // Transportation's bare 'UBER' keyword instead.
    // v1.1.15: added bare 'AMAZON' (see Streaming & Subscriptions note
    // above — general Amazon purchases live here now, only 'AMAZON PRIME'
    // stays under Streaming).
    // v1.1.15: this category is now checked BEFORE Vehicle Gas (see the
    // reordering note on Vehicle Gas below) specifically so 'BONJOUR' here
    // intercepts "BONJOUR TOTAL ARENOSO" before Vehicle Gas's bare 'TOTAL'
    // keyword ever gets a chance to wrongly match it.
    // v1.1.23: ALISS and the Metro Plaza mall ("ALISS METRO PZA") — 'METRO' alone is Transportation, checked later
    keywords: ['RESTAURANTE', 'PIZZERIA', 'BURGER', 'COMIDA', 'FOOD', 'CAFE', 'DELIVERY',
               'UBER*EATS', 'UBEREATS', 'UBER EATS', 'CINE', 'CINEMA', 'PELICULAS', 'CONCIERTO',
               'ENTRETENIMIENTO', 'BAR', 'SUSHI', 'POLLO', 'HELADOS BON', 'BONJOUR',
               'MCDONALD', 'BURGER KING', 'KFC', 'PIZZA HUT', 'SUBWAY', 'POLLOS VICTORINA',
               'COMEDOR', 'WENDYS', 'JADE TERIYAKI', 'SWEETFROG', 'SWEET FROG',
               'PEDIDOSYA', 'KRISPY KREME', 'IKEA REST', 'AMAZON',
               'ROPA', 'ZAPATOS', 'ADIDAS', 'NIKE', 'PUMA STORE', 'ZARA', 'FASHION', 'BOUTIQUE', 'TIENDA', 'ALISS', 'METRO PZA', 'METRO PLZA', 'METRO PLAZA'],
    icon: '🍽️',
    color: ['#FFE4E6', '#BE123C']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Vehicle Gas': {
    // v1.1.15: BUG FIX/CHANGE — added bare 'TOTAL' and 'NEXT LINCOLN' (a
    // real gas-station brand, confirmed as "NEXT LINCOLN GASOPOLIS" from
    // LAFISE and bare "NEXT LINCOLN" from POPULAR — same business, two
    // different card-descriptor truncations). Bare 'TOTAL' was deliberately
    // left OUT before specifically because "BONJOUR TOTAL ARENOSO" (the
    // convenience-store purchase, → Dining) also contains "TOTAL" as a
    // separate word — now safe to add because Dining/Delivery + ... (with
    // 'BONJOUR') is checked BEFORE this category (see its reordering note),
    // so the Bonjour case is intercepted there first, before ever reaching
    // this keyword. 'TOTALENERGIES' kept too, redundant with bare 'TOTAL'
    // now but harmless.
    keywords: ['SHELL', 'HESS', 'PUMA', 'TEXACO', 'GASOLINA', 'COMBUSTIBLE', 'FUEL',
               'TOTALENERGIES', 'TOTAL', 'NEXT LINCOLN'],
    icon: '⛽',
    color: ['#E2E8F0', '#334155']   // v1.1.28: chip [background, text] in the data sheets
  },
  'Health + Vet + Pharmacy': {
    // v1.1.18: was 'LA CASITA DE BOB' — narrowed to 'LA CASITA DE' (the
    // vet clinic) because POPULAR truncates the merchant name to just "LA
    // CASITA DE" (no "BOB-EP" suffix), which didn't match the longer
    // keyword and fell through to the Dining fallback. The shorter keyword
    // is a substring of the longer one, so it still matches LAFISE/BHD's
    // full "LA CASITA DE BOB-EP..." too — one keyword now covers both.
    // v1.1.23: 'FARMA' (FARMA VALUE…), 'FARM' whole word (FARM CAROL…), labs and clinics (ANALISA, CTRO ESPEC MED)
    keywords: ['FARMACIA', 'FCIA', 'DOCTOR', 'MEDICO', 'HOSPITAL', 'CLINICA', 'SALUD',
               'MEDICINAS', 'VETERINARIA', 'VET', 'MEDICAR', 'GBC', 'LOS HIDALGOS',
               'FARMAX', 'FARMAXTRA', 'CRUZ VERDE', 'FARMATODO', 'FARMAVALUE', 'LA CASITA DE', 'FARMA', 'FARM', 'ANALISA', 'LABORATORIO', 'ESPEC MED'],
    icon: '🏥',
    color: ['#E8F5E9', '#2E7D32']   // v1.1.28: chip [background, text] in the data sheets
  },
  // v1.1.24: 'Education' removed — not in use. Those purchases now fall to the
  // Dining/Other fallback; a Custom Rule with Category "Education" brings it
  // back as its own Dashboard row (custom-only categories get one automatically).
  'Transportation': {
    // v1.1.20: added 'DIDI' — real POPULAR rows in COP ("DiDi CO Ride",
    // "DL*DIDI RIDES", "DLO*Didi") were falling through to the Dining
    // fallback. Whole-word match (4 letters), so the '*' variants match;
    // "DIDI FOOD" still goes to Dining via 'FOOD', which is checked first.
    // v1.1.23: Uber rides billed through PayPal ("PAYPAL *UBERBV", "PAYPAL *UBER BV")
    keywords: ['UBER', 'UBER*RIDES', 'UBER*TRIP', 'DIDI', 'TAXI', 'BUS', 'METRO', 'PARKING', 'GUAGUA', 'MOTOCONCHO', 'UBERBV', 'UBER BV'],
    icon: '🚗',
    color: ['#EAF1FE', '#1D4ED8']   // v1.1.28: chip [background, text] in the data sheets
  }
};

const FALLBACK_CATEGORY = 'Dining/Delivery + Entertainment + Other';

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
 * declaration order (order matters, see the category notes above).
 * v1.1.5: fallback changed from 'Other' to FALLBACK_CATEGORY.
 * v1.1.19: the old `customRules` (merged defaults + Custom Rules) parameter
 * is gone. Every caller already runs findCustomRuleOverride() FIRST, so the
 * merged rules could only ever be reached when no custom rule matched — at
 * which point they were identical to the defaults. Also uses the shared
 * keyword engine above instead of raw substring matching.
 */
function categorizeTransaction(text) {
  if (!text) return FALLBACK_CATEGORY;
  const upperText = normalizeKeyword(text);
  for (let category in DEFAULT_CATEGORIES) {
    if (findMatchingKeyword(upperText, DEFAULT_CATEGORIES[category].keywords)) {
      return category;
    }
  }
  return FALLBACK_CATEGORY;
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
        const category = String(data[i][1] === null || data[i][1] === undefined ? '' : data[i][1]).trim();
        const keyword = normalizeKeyword(data[i][2]);
        if (!category || !keyword) continue;

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

const EXCLUDE_CATEGORY = 'Exclude';

/**
 * v1.1.28: chip colours for categories that aren't in DEFAULT_CATEGORIES —
 * "Exclude" (muted: it's left out of every total) and your own Custom Rules
 * categories, which take the next colour of this rotation in order.
 */
const EXCLUDE_COLOR = ['#F3F4F6', '#4B5563'];   // #6B7280 was 4.39:1 — below WCAG AA (caught by v128 test)
const CUSTOM_CATEGORY_COLORS = [
  ['#CFFAFE', '#0E7490'], ['#FAE8FF', '#A21CAF'], ['#D1FAE5', '#047857'], ['#FEF3C7', '#92400E'], ['#FCE7F3', '#9D174D']
];

/** Every category with its chip colours, in display order: defaults, Exclude, then custom ones. */
function categoryPalette(customNames) {
  const list = getCategories().map(name => ({ name: name, bg: DEFAULT_CATEGORIES[name].color[0], fg: DEFAULT_CATEGORIES[name].color[1] }));
  list.push({ name: EXCLUDE_CATEGORY, bg: EXCLUDE_COLOR[0], fg: EXCLUDE_COLOR[1] });
  (customNames || []).filter(n => n && !DEFAULT_CATEGORIES[n] && n.toUpperCase() !== EXCLUDE_CATEGORY.toUpperCase())
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
 * (say "Pets") never reached any row or the TOTAL — see
 * buildOrRefreshDashboard(). getMergedCategoryRules() was removed in this
 * version (see categorizeTransaction()).
 */
function getCustomCategoryNames(userEmail) {
  const custom = getUserCustomRules(userEmail);
  return Object.keys(custom).filter(c =>
    !DEFAULT_CATEGORIES[c] && c.toUpperCase() !== EXCLUDE_CATEGORY.toUpperCase());
}

