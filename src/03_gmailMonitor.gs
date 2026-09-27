/**
 * GMAIL MONITOR MODULE
 * Searches and parses transaction emails from different banks
 */

/**
 * v1.1.0: senders, subjects, AND real body formats are now all confirmed —
 * the user shared actual .eml files for 6 real transaction emails
 * (Sept 23, 2026), so extraction below (see the bank-specific extractXxx()
 * functions further down) is written and tested directly against real text,
 * not guessed. Summary per bank:
 *   LAFISE   — notificaciones@bancolafise.com
 *              Consumo ("Servicio de Alerta - Nuevo Consumo"): labeled fields,
 *              "Comercio/Ciudad/País:" (merchant) BEFORE "Monto:" (amount,
 *              format "DOP 412.86" — no $ sign, uses the DOP code instead).
 *              Transferencia ("<Nombre>, ¡Transferencia exitosa!"): no real
 *              sample seen yet — falls back to the generic extractor.
 *   BANESCO  — notificaciones@banesco.com.do
 *              Consumo, approved ("Alerta de Consumo Banesco RD"): one
 *              sentence — "...consumo de RD$ 2,640.00, en MERCHANT y su
 *              estado es aprobada." Merchant comes AFTER the amount here,
 *              opposite of LAFISE. A declined variant exists with different
 *              wording ("...ha sido rechazada...") — skipped per-match via a
 *              status check, not by rejecting the whole email.
 *              Transferencia ("Notificación de Transferencia Realizada"):
 *              "Monto: DOP545.00" (no $, no space) ... later ... "Nombre del
 *              Beneficiario: X".
 *   BHD      — Alertas@bhd.com.do
 *              Consumo ("BHD Notificación de Transacciones"): an actual
 *              table — header row (Fecha/Moneda/Monto/Comercio/Estado/Tipo)
 *              then one repeating 6-line block per transaction. This is
 *              where same-day bundling is most likely to actually show up
 *              (multiple row-blocks) — the extractor uses a global regex so
 *              it naturally picks up as many rows as exist.
 *              Transferencia ("Transacciones entre mis productos"): labeled
 *              fields, "Monto:" then later "Beneficiario:".
 *   POPULAR  — notificaciones@popularenlinea.com
 *              Consumo ("Notificación de Consumo"): a compact table (Monto /
 *              Moneda / Fecha / Comercio / Estatus) on one visual row, amount
 *              format "US$14.27".
 *   BDI      — bdi.com.do confirmed, but still no consumption-specific
 *              sample — kept on the old generic extractor (extractAllAmounts)
 *              until a real one shows up.
 *
 * Currency formats actually observed across these 6 samples: "RD$ 2,640.00",
 * "RD 8,315.20" (no $), "DOP 412.86", "DOP545.00" (no space), "RD$ 91.37",
 * "RD" + "$118.40" split across a table cell/newline, "US$14.27". All of
 * every extractXxx() function (and the generic extractAllAmounts() fallback) use \b(?:RD|DOP|USD|US\$|EUR|COP) then an
 * independently-optional \$? — that's what makes both the attached
 * ("RD$91.37") and detached ("RD" ... "$118.40") cases match correctly.
 */
// v1.1.19 (E4): each bank declares its own extractors ({consumo, transfer}),
// so extractTransactionItems() is a lookup instead of an if/else chain and
// adding BDI later means filling in two functions here. The unused
// `amountPattern` field was removed (M13) — nothing had read it since the
// bank-specific extractors replaced it in v1.1.0. Extractor functions are
// plain function declarations further down (hoisted, so referencing them
// here is safe).
const BANK_PATTERNS = {
  LAFISE: {
    name: 'LAFISE',
    // v1.1.5: LAFISE actually sends from TWO different domains — consumption
    // alerts from bancolafise.com, but transfers ("<Name>, ¡Transferencia
    // exitosa!") from a separate digital@notificaciones.lafise.com address,
    // confirmed directly by the user. Neither domain is a substring of the
    // other, so fromDomain is broadened to the shared suffix 'lafise.com'
    // (matches both) and searchQuery now ORs both exact addresses. This is
    // why LAFISE transfers were invisible before — the old search only ever
    // looked for the consumption-alert sender.
    fromDomain: 'lafise.com',
    searchQuery: 'from:notificaciones@bancolafise.com OR from:digital@notificaciones.lafise.com',
    extractors: { consumo: extractLAFISETransactions, transfer: extractLAFISETransferTransactions },
    keywords: ['LAFISE', 'LAFISE BANCO'],
    merchantPattern: /(?:en|en el|en\s+)([^\n]{10,50})/i
  },
  BANESCO: {
    name: 'BANESCO',
    fromDomain: 'banesco.com.do',
    searchQuery: 'from:notificaciones@banesco.com.do',
    extractors: { consumo: extractBANESCOConsumoTransactions, transfer: extractBANESCOTransferTransactions },
    keywords: ['BANESCO', 'BANESCO RD'],
    merchantPattern: /realizada en el\s+([^\n]+?)\s+por\s/i
  },
  BHD: {
    name: 'BHD',
    fromDomain: 'bhd.com.do',
    searchQuery: 'from:Alertas@bhd.com.do',
    extractors: { consumo: extractBHDConsumoTransactions, transfer: extractBHDTransferTransactions },
    keywords: ['BHD', 'BANCO HOTELES DOMINICANA'],
    merchantPattern: /(?:en|en el)([^\n]{10,50})/i
  },
  POPULAR: {
    name: 'POPULAR',
    fromDomain: 'popularenlinea.com',
    searchQuery: 'from:notificaciones@popularenlinea.com',
    extractors: { consumo: extractPOPULARTransactions, transfer: extractPOPULARTransactions },
    keywords: ['POPULAR', 'BANCO POPULAR'],
    merchantPattern: /(?:en|en el)([^\n]{10,50})/i
  },
  BDI: {
    name: 'BDI',
    fromDomain: 'bdi.com.do',
    searchQuery: 'from:bdi.com.do',
    extractors: {}, // no real sample yet — always uses the generic extractAllAmounts() fallback
    keywords: ['BDI', 'BDI DIGITAL', 'BANCO BDI'],
    merchantPattern: /(?:en|en el|hacia|a)([^\n]{10,50})/i
  }
};

/**
 * v1.0.3: Emails to reject outright — marketing/promo sends from the bank,
 * never real consumption notifications. Checked against subject + body.
 * v1.0.8: removed the standalone 'GANA' entry — a real run showed it
 * matching randomly inside long tracking-parameter strings in HTML email
 * boilerplate (e.g. "z=1AvGy9akdi2..."), not real promotional text. Lesson:
 * keep entries here as full words/phrases of 6+ characters where possible;
 * short 4-letter fragments have real collision risk against URL noise.
 * 'PARTICIPA Y GANA' (the full phrase) is unaffected and stays.
 */
const PROMOTIONAL_KEYWORDS = [
  'PROMOCIÓN', 'PROMOCION', 'PROMO', 'OFERTA', 'DESCUENTO', 'NEWSLETTER',
  'BOLETÍN', 'BOLETIN', 'ENTÉRATE', 'ENTERATE', 'NUEVO BENEFICIO',
  'SORTEO', 'PARTICIPA Y GANA', 'DISFRUTA DE', 'CONOCE NUESTRO'
];

/**
 * v1.0.6: real bank notifications that aren't a single transaction — account
 * security (OTP, login, beneficiaries), periodic statements, and surveys/
 * service announcements. v1.0.8: widened past just "account notices" once
 * real samples showed "Estado de cuenta...", "Queremos conocer tu opinión",
 * and event-invite mail ("Acompáñanos en...") all needed the same treatment
 * — none of these represent one transaction, so none should reach
 * extractAmount(). Checked against subject + body, same as
 * PROMOTIONAL_KEYWORDS, for the same reason: a stray RD$ figure inside a
 * statement or survey email would otherwise create a phantom transaction.
 * v1.1.9: widened again — a real POPULAR credit-LIMIT-INCREASE notification
 * ("Actualización de Límite", subject confirmed from a live run) mentions
 * the new/old limit as a dollar figure and was getting parsed as if it were
 * a real purchase (saved rows showed amounts like 50,000/70,000 with
 * merchant "Unknown Merchant" or garbled limit-notice text) — same failure
 * shape as every other non-transactional email that slipped through before
 * its category was covered here.
 */
const NON_TRANSACTIONAL_KEYWORDS = [
  // Account / security
  'OTP', 'CÓDIGO DE VERIFICACIÓN', 'CODIGO DE VERIFICACION', 'CÓDIGO DE SEGURIDAD',
  'CODIGO DE SEGURIDAD', 'ALERTA DE ACCESO', 'ALERTA ACCESO', 'INICIO DE SESIÓN',
  'INICIO DE SESION', 'NUEVO DISPOSITIVO', 'BENEFICIARIO AGREGADO', 'BENEFICIARIO ELIMINADO',
  'CAMBIO DE CONTRASEÑA', 'CAMBIO DE CLAVE', 'RESTABLECER CONTRASEÑA', 'ACTUALIZACIÓN DE DATOS',
  'ACTUALIZACION DE DATOS', 'VERIFICACIÓN DE IDENTIDAD', 'VERIFICACION DE IDENTIDAD',
  'SOLICITUD OPCIONES DE MENÚ', 'SOLICITUD OPCIONES DE MENU',
  // v1.1.23: incoming payroll deposit notices (POPULAR "Notificación Depósito de
  // Nómina") are income, not spending — they were reported as "Could not parse".
  'DEPÓSITO DE NÓMINA', 'DEPOSITO DE NOMINA',
  // Statements / periodic summaries — not a single transaction
  'ESTADO DE CUENTA', 'FONDO DE INVERSIÓN', 'FONDO DE INVERSION',
  // Surveys / service announcements / events
  'QUEREMOS CONOCER TU OPINIÓN', 'QUEREMOS CONOCER TU OPINION', 'ENCUESTA',
  'ACOMPÁÑANOS', 'ACOMPAÑANOS', 'FORMA MÁS CONVENIENTE', 'FORMA MAS CONVENIENTE',
  // Credit limit change notices — mention a dollar figure, but it's a limit, not a purchase
  'ACTUALIZACIÓN DE LÍMITE', 'ACTUALIZACION DE LIMITE', 'AUMENTO DE LÍMITE', 'AUMENTO DE LIMITE',
  'INCREMENTO DE LÍMITE', 'INCREMENTO DE LIMITE', 'NUEVO LÍMITE', 'NUEVO LIMITE',
  'LÍMITE ANTERIOR', 'LIMITE ANTERIOR', 'LÍMITE DE TU TARJETA', 'LIMITE DE TU TARJETA'
];

/**
 * v1.0.8: a real transaction alert can still describe a DECLINED charge —
 * confirmed from a real Banesco sample ("...ha sido rechazada. Motivo de la
 * declinación: TARJETA VENCIDA"). No money moved, so this must not be logged
 * as an expense even though it has a real RD$ amount and a real merchant.
 */
const DECLINED_KEYWORDS = [
  'HA SIDO RECHAZADA', 'FUE RECHAZADA', 'FUE RECHAZADO', 'DECLINADA', 'DECLINACIÓN',
  'DECLINACION', 'TRANSACCIÓN FALLIDA', 'TRANSACCION FALLIDA', 'NO APROBADA', 'NO FUE APROBADA'
];

/**
 * v1.1.4: Type values translated to English throughout the system (sheet
 * names, headers, category names, and these values were a mix of English UI
 * with Spanish data — now consistently English end to end).
 * Also fixes a real bug: bare 'CASHBACK' matched inside "VISA CLASICA
 * SUPERCASHBACK" — a real Banesco CARD PRODUCT NAME (confirmed from a live
 * email), not an actual cashback credit — which mistyped every purchase made
 * with that card as Type=Cashback instead of Transaction, and since Category
 * is only assigned when Type=Transaction, those rows silently got no
 * category at all (and stayed that way even after recategorizing, since
 * recategorization also respects the — wrong — stored Type). Removed the
 * bare keyword; kept only the more specific reward-language ones. No real
 * "you received cashback" sample has been seen yet, so this Type is
 * best-effort until one surfaces — it errs toward under-detecting Cashback
 * rather than risk another silent false-positive like this one.
 */
const TYPE_KEYWORDS = {
  'Card Payment': ['PAGO DE TARJETA', 'PAGO TC', 'PAGO REALIZADO A SU TARJETA', 'PAGO A TARJETA DE CREDITO'],
  'Cashback': ['DEVOLUCION', 'REEMBOLSO', 'REBATE', 'CASHBACK ACREDITADO', 'CASHBACK RECIBIDO'],
  'Transfer': ['TRANSFERENCIA ENVIADA', 'TRANSFERENCIA RECIBIDA', 'PAGOS AL INSTANTE', 'ACH', 'SWIFT'],
};

/**
 * v1.0.9: confirmed real subject lines, checked first (subject alone, exact
 * substring) since they're a much stronger signal than guessing body
 * phrasing. LAFISE's transfer subject includes the customer's name as a
 * prefix ("<Nombre>, ¡Transferencia exitosa!") so this matches on the fixed
 * suffix only. Keys are the English Type values (v1.1.4) — the keywords
 * themselves stay in Spanish since that's the actual language of the bank
 * emails being matched.
 */
const TYPE_SUBJECT_KEYWORDS = {
  'Transfer': [
    '¡TRANSFERENCIA EXITOSA!',              // LAFISE
    'NOTIFICACIÓN DE TRANSFERENCIA REALIZADA', 'NOTIFICACION DE TRANSFERENCIA REALIZADA', // BANESCO
    'TRANSACCIONES ENTRE MIS PRODUCTOS',    // BHD
  ],
  // v1.1.17: BUG FIX — real confirmed LAFISE subject is "¡Realizaste un
  // pago a tu tarjeta LAFISE!" (active/informal phrasing — "you made a
  // payment"), which the old TYPE_KEYWORDS body-guesses ('PAGO REALIZADO A
  // SU TARJETA', passive/formal — "was made to your card") never matched.
  // These rows were getting Type=Transaction with Category defaulting to
  // Dining (the fallback), so paying off the LAFISE card bill was being
  // double-counted as a new expense in the Dashboard on top of the
  // original purchases that made up the balance.
  'Card Payment': [
    '¡REALIZASTE UN PAGO A TU TARJETA',     // LAFISE
  ],
  // v1.1.19 (C3): confirmed CONSUMO subjects. Matching one of these pins the
  // email to Transaction (only a Cashback refinement from the body is still
  // allowed — see detectTransactionType()), so body-keyword guesses like
  // 'ACH' can never re-type a purchase again. This is also what lets
  // "🔁 Recategorize" repair rows that were already mis-typed (the real
  // CACHAREPA CHURCHILL row) — recategorize has the subject, not the body.
  'Transaction': [
    'SERVICIO DE ALERTA - NUEVO CONSUMO',                                   // LAFISE
    'ALERTA DE CONSUMO BANESCO',                                            // BANESCO
    'BHD NOTIFICACIÓN DE TRANSACCIONES', 'BHD NOTIFICACION DE TRANSACCIONES', // BHD
    'NOTIFICACIÓN DE CONSUMO', 'NOTIFICACION DE CONSUMO',                   // POPULAR
  ]
};

/**
 * Returns true if this email looks promotional rather than a real
 * consumption/transaction notification.
 */
function isPromotionalEmail(subject, bodyText) {
  const upper = (subject + ' ' + bodyText).toUpperCase();
  const matched = findMatchingKeyword(upper, PROMOTIONAL_KEYWORDS); // v1.1.19: shared engine (02_categorizer.gs)
  if (matched) {
    Logger.log("  ↳ matched promotional keyword: \"" + matched + "\"");
    return true;
  }
  return false;
}

/**
 * v1.0.8: renamed from isAccountNoticeEmail — scope widened past just
 * account/security notices (see NON_TRANSACTIONAL_KEYWORDS above).
 */
function isNonTransactionalEmail(subject, bodyText) {
  const upper = (subject + ' ' + bodyText).toUpperCase();
  const matched = findMatchingKeyword(upper, NON_TRANSACTIONAL_KEYWORDS); // v1.1.19: shared engine (02_categorizer.gs)
  if (matched) {
    Logger.log("  ↳ matched non-transactional keyword: \"" + matched + "\"");
    return true;
  }
  return false;
}

/**
 * v1.0.8: true if this is a real transaction alert describing a DECLINED /
 * failed charge — no money moved, so it must not be logged as an expense.
 */
function isDeclinedTransactionEmail(subject, bodyText) {
  const upper = (subject + ' ' + bodyText).toUpperCase();
  const matched = findMatchingKeyword(upper, DECLINED_KEYWORDS); // v1.1.19: shared engine (02_categorizer.gs)
  if (matched) {
    Logger.log("  ↳ matched declined-transaction keyword: \"" + matched + "\"");
    return true;
  }
  return false;
}

/**
 * Classifies a transaction email into Transaction / Transfer / Card Payment / Cashback.
 * v1.0.9: checks confirmed subject lines first (TYPE_SUBJECT_KEYWORDS), then
 * falls back to body-keyword guessing (TYPE_KEYWORDS). Defaults to
 * "Transaction" (a card consumption) when nothing else matches.
 */
/**
 * v1.1.19 (C3): subject-only type detection. Returns null when the subject
 * isn't one of the confirmed ones. Used on its own by
 * recategorizeAllTransactions() (04_sheetsWriter.gs), which only has the
 * subject — it used to call detectTransactionType(subject, MERCHANT), i.e.
 * it fed the merchant name in as if it were the email body, so 'ACH' inside
 * "CACHAREPA" re-typed a real purchase as Transfer on every run (verified),
 * and any type that had been decided from the real body got reverted.
 */
function detectTypeFromSubject(subject) {
  const upper = String(subject || '').toUpperCase();
  for (let type in TYPE_SUBJECT_KEYWORDS) {
    if (findMatchingKeyword(upper, TYPE_SUBJECT_KEYWORDS[type])) return type;
  }
  return null;
}

function detectTransactionType(subject, bodyText) {
  const bySubject = detectTypeFromSubject(subject);
  if (bySubject === 'Transfer' || bySubject === 'Card Payment') return bySubject;
  const upper = (String(subject || '') + ' ' + String(bodyText || '')).toUpperCase();
  if (bySubject === 'Transaction') {
    // Confirmed consumo template: the body may only refine it to Cashback
    // (a refund row in the same template) — never to Transfer/Card Payment.
    return findMatchingKeyword(upper, TYPE_KEYWORDS['Cashback']) ? 'Cashback' : 'Transaction';
  }
  // Unknown subject (BDI, new templates): body keywords, whole-word for
  // short ones like 'ACH' (see keywordMatches(), 02_categorizer.gs).
  for (let type in TYPE_KEYWORDS) {
    if (findMatchingKeyword(upper, TYPE_KEYWORDS[type])) return type;
  }
  return 'Transaction';
}


/**
 * v1.1.19 (C1): turns a user-facing inclusive date range into exact
 * [start, endExclusive) bounds in the script's time zone.
 * Why: Gmail's `before:` operator EXCLUDES the date it's given (Google's own
 * example `after:2004/04/16 before:2004/04/18` returns the 16th and 17th).
 * The old query used `before:<end date>`, so "Monitor Gmail Now" never saw
 * today's emails, every date-range run silently dropped its own end date,
 * and the 6 AM run lost the LAST DAY OF EVERY MONTH (the run on the 30th
 * stopped at the 29th; the run on the 1st searched `after:X before:X`, an
 * empty window). Bounds are sent to Gmail as epoch seconds, which Gmail
 * accepts and which also removes any ambiguity about which time zone Gmail
 * uses for plain yyyy/MM/dd dates. Accepts 'yyyy-MM-dd' or 'yyyy/MM/dd'.
 */
function buildRangeBounds(startStr, endStr) {
  const tz = Session.getScriptTimeZone();
  const norm = v => String(v).trim().replace(/\//g, '-');
  const start = Utilities.parseDate(norm(startStr), tz, 'yyyy-MM-dd');
  const endExclusive = Utilities.parseDate(norm(endStr), tz, 'yyyy-MM-dd');
  if (isNaN(start.getTime()) || isNaN(endExclusive.getTime())) {
    throw new Error('Invalid date range: ' + startStr + ' – ' + endStr);
  }
  if (start.getTime() > endExclusive.getTime()) {
    throw new Error('Start date (' + startStr + ') is after end date (' + endStr + ').');
  }
  endExclusive.setDate(endExclusive.getDate() + 1); // include the whole end day
  return { start: start, endExclusive: endExclusive };
}

function toEpochSeconds(date) {
  return Math.floor(date.getTime() / 1000);
}

/**
 * v1.1.19 (M5): paginated Gmail search. The date-range search used to make
 * ONE GmailApp.search(query, 0, 500) call, so a big quarterly backlog lost
 * everything past thread 500 without a word. Pages through results up to
 * `limit` and reports whether the cap was reached so the run summary can say
 * so. (The old rolling-window searchTransactionEmails() was removed here —
 * nothing had called it since v1.1.12.)
 */
const MAX_THREADS_PER_RUN = 2000;

function gmailSearchAll(query, limit) {
  const PAGE = 100;
  let threads = [];
  let start = 0;
  while (threads.length < limit) {
    const want = Math.min(PAGE, limit - threads.length);
    const batch = GmailApp.search(query, start, want);
    threads = threads.concat(batch);
    start += batch.length;
    if (batch.length < want) return { threads: threads, capped: false };
  }
  return { threads: threads, capped: true };
}

/**
 * Diagnostic helper (added v1.0.4, sender logging added v1.0.5). Run this
 * directly from the Apps Script editor — pick debugBankEmailSample in the
 * function dropdown, run it, then View > Logs — any time the monitor comes
 * back with 0 results, or to confirm BHD's/POPULAR's real sender (still
 * unconfirmed as of v1.0.5 — see the BANK_PATTERNS comment above). It
 * searches by bank name ONLY, no domain/keyword restriction, so you see
 * everything, including false positives.
 */
function debugBankEmailSample(daysBack = 30, perBank = 5) {
  const afterDate = Utilities.formatDate(
    new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  for (let bankName in BANK_PATTERNS) {
    const keyword = BANK_PATTERNS[bankName].keywords[0];
    const query = `${keyword} after:${afterDate}`;
    const threads = GmailApp.search(query, 0, perBank);
    Logger.log(`--- ${bankName} (query: "${query}") — ${threads.length} thread(s) ---`);
    threads.forEach(t => {
      const msg = t.getMessages()[0];
      Logger.log(`  [${msg.getDate()}] isUnread=${msg.isUnread()} | From: ${msg.getFrom()} | ${msg.getSubject()}`);
    });
  }
  Logger.log("Done. All 5 banks now have a confirmed sender (see BANK_PATTERNS above); " +
             "re-run this if a new promotional/marketing address starts showing up " +
             "under a bank you track.");
}

/**
 * v1.1.19 (E7): per-run counters, surfaced in the final summary alert —
 * before this, a filtered or unparseable email only ever showed up in the
 * execution log.
 */
function newParseStats() {
  return {
    messagesSeen: 0, outOfRange: 0, notOwnBank: 0, promotional: 0,
    nonTransactional: 0, declined: 0, amountNotFound: 0, parseErrors: 0, placeholders: 0, reversals: 0,
    unrecognized: [], readIds: [],  // v1.1.35: emails for the Unrecognized sheet, and emails read cleanly
    alreadySaved: 0                  // v1.1.36: skipped without reading — already in Transactions
  };
}

/**
 * Extract transactions from email threads.
 * History: v1.0.4 stopped skipping read messages (dedup is the real guard);
 * v1.0.9 made parseEmailMessage() return several items per email; v1.1.5
 * added rawCustomRules so a personal rule can categorize any Type.
 * v1.1.19:
 *  - (M6) takes the run's `range` and skips messages outside it. Gmail
 *    matches THREADS, and it groups every BHD alert with the same subject
 *    into one thread, so one thread can hold months of messages — all of
 *    them were re-parsed on every run, eating into the 6-minute limit.
 *  - returns { transactions, stats, failedThreadIds } — failedThreadIds are
 *    threads with a message that couldn't be parsed, which
 *    markEmailsAsProcessed() now leaves UNREAD so they stay visible (M4).
 *  - the merged `customRules` argument is gone (see categorizeTransaction()).
 */
function extractTransactionsFromThreads(threads, rawCustomRules, range, opts) {
  opts = opts || {};
  const transactions = [];
  const stats = newParseStats();
  const failedThreadIds = new Set();
  // v1.1.36: an email already in Transactions isn't read again (reading is what takes time), and reading stops at
  // opts.deadline — the threads not reached stay unread and are picked up by the next run.
  const skipIds = opts.skipIds || new Set();
  const clock = opts.clock || (() => Date.now());
  const processedThreads = [];
  let stopped = null;

  for (let t = 0; t < threads.length; t++) {
    const thread = threads[t];
    if (opts.deadline && clock() > opts.deadline) {
      stopped = { remaining: threads.length - t, processed: t };
      Logger.log("⏸ Stopped reading at the time budget: " + stopped.remaining + " thread(s) left for the next run");
      break;
    }
    for (let message of thread.getMessages()) {
      if (range) {
        const d = message.getDate();
        if (d < range.start || d >= range.endExclusive) { stats.outOfRange++; continue; }
      }
      stats.messagesSeen++;
      if (skipIds.has(message.getId())) { stats.alreadySaved++; continue; }
      const result = parseEmailMessage(message, rawCustomRules, stats);
      if (result.status === 'failed') failedThreadIds.add(thread.getId());
      transactions.push(...result.items);
      // v1.1.35: what the Unrecognized sheet lists — not saved, or saved with an unreadable merchant
      const unreadable = result.status === 'ok' && result.items.some(it => it.merchant === GARBLED_PLACEHOLDER);
      if (result.status === 'failed' || unreadable) {
        stats.unrecognized.push({ id: message.getId(), date: message.getDate(), bank: result.bank || (result.items[0] || {}).bank || '',
          subject: message.getSubject() || '', reason: unreadable ? 'Saved — merchant unreadable (red row in Transactions)' : result.reason,
          snippet: result.snippet || (unreadable ? (result.items[0].description || '') : '') });
      } else if (result.status === 'ok' || result.status === 'filtered') {
        stats.readIds.push(message.getId());
      }
    }
    processedThreads.push(thread);
  }
  return { transactions: transactions, stats: stats, failedThreadIds: failedThreadIds, processedThreads: processedThreads, stopped: stopped };
}

/**
 * v1.1.19 (M1): true if the email has at least one APPROVED row. The
 * whole-email "declined" filter used to discard a multi-row BHD table
 * entirely if ANY row said "Declinada" — including the approved rows, which
 * the extractor already knows how to keep. Now the email-level filter only
 * fires when nothing in it was approved; row-level status is handled by the
 * extractors (they flag declined rows with `declined: true`).
 * ("NO FUE APROBADA"/"NO APROBADA" are removed first so they don't count.)
 */
function hasApprovedRow(text) {
  const cleaned = String(text || '').toUpperCase().replace(/\bNO\s+(FUE\s+)?APROBAD[AO]\b/g, '');
  return /\bAPROBAD[AO]\b/.test(cleaned);
}

/**
 * v1.1.19 (M3): IsCredit = money coming IN (an abono). It used to be YES
 * whenever the WHOLE email body mentioned "crédito" or "depósito" anywhere,
 * including boilerplate like "tu tarjeta de crédito". Now: Card Payment and
 * Cashback are credits by definition; otherwise only an explicit CRÉDITO /
 * DEPÓSITO inside THIS item's own context counts, and the product name
 * "TARJETA DE CRÉDITO" never does.
 */
function computeIsCredit(type, context) {
  if (type === 'Card Payment' || type === 'Cashback') return true;
  const upper = String(context || '').toUpperCase().replace(/TARJETA\s+DE\s+CR[ÉE]DITO/g, '');
  return /\b(CR[ÉE]DITO|DEP[ÓO]SITO)\b/.test(upper);
}

/**
 * Parse one email into zero or more transaction items.
 * History: v1.0.5 bank by sender domain; v1.0.3/v1.0.8 promotional,
 * non-transactional and declined filters; v1.0.9 array of items; v1.1.0
 * bank+type-specific extractors with the generic fallback; v1.1.5 Custom
 * Rules checked first for every Type; v1.1.13 messageId = Gmail id + index.
 * v1.1.19:
 *  - returns { items, status } — status is 'ok', 'filtered', 'skipped' or
 *    'failed', and `stats` (see newParseStats()) is updated in place.
 *  - (E2) categorizes on the MERCHANT, the same text "🔁 Recategorize" uses.
 *    It used to categorize on the raw match context here and on the
 *    Description there, so the two could disagree and every run rewrote
 *    categories back and forth.
 *  - (M1) declined rows flagged by an extractor are dropped individually;
 *    if every row was declined, the email is skipped WITHOUT falling back to
 *    extractAllAmounts() — that fallback used to save a declined BANESCO
 *    charge anyway, plus its "Dispones de RD$ …" balance as a second row.
 *  - (M3) IsCredit via computeIsCredit().
 */
function parseEmailMessage(message, rawCustomRules, stats) {
  stats = stats || newParseStats();
  try {
    const subject = message.getSubject() || '';
    const plainText = message.getPlainBody() || '';
    const fromAddress = message.getFrom() || ''; // e.g. "LAFISE CF <notificaciones@bancolafise.com>"

    // v1.0.5: bank by the sender's DOMAIN, not by the bank name appearing in
    // the text (Uber receipts mention "LAFISE" as the payment method).
    let bank = null;
    let bankPattern = null;
    for (let bankName in BANK_PATTERNS) {
      const pattern = BANK_PATTERNS[bankName];
      if (pattern.fromDomain && fromAddress.toLowerCase().includes(pattern.fromDomain.toLowerCase())) {
        bank = bankName;
        bankPattern = pattern;
        break;
      }
    }
    if (!bank) {
      stats.notOwnBank++;
      Logger.log("Bank not detected (sender domain not recognized): " + fromAddress + " | " + subject);
      return { items: [], status: 'skipped' };
    }

    if (isPromotionalEmail(subject, plainText)) {
      stats.promotional++;
      Logger.log("Skipped promotional email: " + subject);
      return { items: [], status: 'filtered' };
    }
    if (isNonTransactionalEmail(subject, plainText)) {
      stats.nonTransactional++;
      Logger.log("Skipped non-transactional email: " + subject);
      return { items: [], status: 'filtered' };
    }
    if (isDeclinedTransactionEmail(subject, plainText) && !hasApprovedRow(plainText)) {
      stats.declined++;
      Logger.log("Skipped declined/failed transaction: " + subject);
      return { items: [], status: 'filtered' };
    }

    const date = formatDate(message.getDate());
    const type = detectTransactionType(subject, plainText);

    let items = extractTransactionItems(bank, type, plainText);
    items.forEach(it => { it.merchant = fixStatusAsMerchant(it.merchant, subject, plainText); });   // v1.1.34
    const declinedRows = items.filter(it => it.declined).length;
    items = items.filter(it => !it.declined);
    if (items.length === 0 && declinedRows > 0) {
      stats.declined++;
      Logger.log("Skipped declined transaction (row-level status): " + subject);
      return { items: [], status: 'filtered' };
    }
    if (items.length === 0) {
      items = extractAllAmounts(plainText, bankPattern);
    }
    if (items.length === 0) {
      stats.amountNotFound++;
      const snippet = plainText.substring(0, 700).replace(/\s+/g, ' ').trim();
      Logger.log("Amount not found | " + bank + " | " + subject + " | Body: \"" + snippet + "\"");
      return { items: [], status: 'failed', bank: bank, reason: 'Amount not found', snippet: snippet };
    }

    const results = [];
    const baseMessageId = message.getId(); // v1.1.13: unique per email
    let itemIndex = 0;
    for (let item of items) {
      // v1.0.3: skip zero-amount lines (v1.1.19: and NaN, which used to slip through)
      if (!(item.amount > 0)) {
        Logger.log("Skipped $0.00 line within: " + subject);
        continue;
      }
      // v1.1.23: a reversal is saved as a NEGATIVE amount; its merchant and
      // category are filled in from the original purchase at save time
      // (resolveReversals(), 04_sheetsWriter.gs), so the two net to zero.
      const isReversal = !!item.reversal;
      if (isReversal) stats.reversals++;
      const merchant = isReversal ? REVERSAL_UNMATCHED : (String(item.merchant || '').trim() || 'Unknown Merchant');
      let category = findCustomRuleOverride(merchant, rawCustomRules);
      if (!category && type === 'Transaction' && !isReversal) {   // a reversal takes its original's category
        category = categorizeTransaction(merchant);
      }
      // v1.1.23: paying the card is never spending — explicit "Exclude"
      if (type === 'Card Payment') category = EXCLUDE_CATEGORY;
      if (merchant === GARBLED_PLACEHOLDER) stats.placeholders++;
      const currency = item.currency || 'DOP';

      Logger.log("Parsed OK | " + bank + " | " + type + " | " + (category || '-') + " | " + currency + " " +
                 item.amount + " | merchant=\"" + merchant + "\" | Context: \"" + (item.context || '') + "\"");

      results.push({
        date: date,
        bank: bank,
        merchant: merchant,
        amount: isReversal ? -item.amount : item.amount,
        currency: currency,
        category: category,
        type: type,
        description: merchant,
        reversal: isReversal,
        timeKey: item.timeKey || '',
        txRef: item.ref ? bank + ':' + item.ref : '',
        subject: subject,
        timestamp: new Date().toISOString(),
        messageId: baseMessageId + '_' + (itemIndex++),
        isCredit: isReversal || computeIsCredit(type, item.context),
        isCashback: type === 'Cashback'
      });
    }
    return { items: results, status: 'ok' };
  } catch (error) {
    stats.parseErrors++;
    Logger.log("Error parsing email: " + error + (error && error.stack ? " | " + error.stack : ""));
    let snippet = '';
    try { snippet = String(message.getPlainBody() || '').substring(0, 700).replace(/\s+/g, ' ').trim(); } catch (e) { snippet = ''; }
    return { items: [], status: 'failed', bank: '', reason: 'Could not read: ' + error, snippet: snippet };
  }
}

/**
 * v1.1.0: routes to the right bank+type-specific extractor, all written and
 * tested against real .eml samples (Sept 23, 2026). Each returns an array of
 * {amount, merchant, context}. BDI has no real sample yet, so it goes
 * straight to the generic fallback (extractAllAmounts, further below).
 */
/**
 * v1.1.1: detects USD vs DOP from the matched text itself.
 * v1.1.5: broadened to any 2-3 letter code, to catch a real COP transaction
 * — but that was TOO broad: a live run showed it capturing garbage like
 * "FKR", "WEC", "QAH", "JDH", "UCO", "BGJ" as "currency" — random uppercase
 * fragments from authorization codes, reference numbers, and other noise in
 * the email body that happened to precede a digit somewhere in the matched
 * context window.
 * v1.1.6: BUG FIX — reverted to a fixed, explicit list of the 4 currencies
 * actually relevant here (DOP, USD, EUR, COP), per request, rather than a
 * generic [A-Z]{3} wildcard. Extend this list deliberately if a genuine new
 * currency shows up in real data — don't widen it back to a wildcard.
 */
function detectCurrencyFromMatch(matchText) {
  const m = matchText.match(/\b(RD|US|DOP|USD|EUR|COP)\s*\$?\s*[\d,]/);
  if (!m) return 'DOP';
  if (m[1] === 'RD') return 'DOP';
  if (m[1] === 'US') return 'USD';
  return m[1];
}

/**
 * Dispatches to the bank+type extractor declared in BANK_PATTERNS.
 * History: v1.1.10 fixed LAFISE transfers always using the consumo
 * extractor (they had no extractor of their own, so they fell through to the
 * generic fallback and produced garbled merchants).
 * v1.1.19 (E4): table lookup. Anything that isn't a Transfer (Transaction,
 * Card Payment, Cashback) uses the bank's consumo extractor, as before; a
 * bank with no extractor (BDI) returns [] and parseEmailMessage() falls back
 * to extractAllAmounts().
 */
function extractTransactionItems(bank, type, text) {
  const extractors = (BANK_PATTERNS[bank] && BANK_PATTERNS[bank].extractors) || {};
  const fn = type === 'Transfer' ? extractors.transfer : extractors.consumo;
  return fn ? fn(text) : [];
}

/**
 * LAFISE consumo ("Servicio de Alerta - Nuevo Consumo"). Real sample:
 *   Comercio/Ciudad/País:
 *   UBER*EATS SANTO DOMINGO DOM
 *   ...
 *   Monto:
 *   DOP 412.86
 * Merchant comes BEFORE the amount here — the opposite of every other bank.
 * `[\s\S]*?` (not `.*?`) bridges the Fecha/Marca/Tarjeta/Autorización/
 * Referencia/Tipo fields in between regardless of the newlines they contain.
 */
function extractLAFISETransactions(text) {
  const re = /Comercio\/Ciudad\/País:\s*([^\n]+?)\s*\n[\s\S]*?Monto:\s*(?:RD|DOP|USD|US\$|EUR|COP)?\s*\$?\s*([\d,]+\.?\d*)/gi;
  const results = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    results.push({ merchant: m[1].trim().substring(0, 50), amount: parseFloat(m[2].replace(/,/g, '')), currency: detectCurrencyFromMatch(m[0]), context: m[0].replace(/\s+/g, ' ').trim() });
  }
  return results;
}

/**
 * LAFISE transferencia ("<Nombre>, ¡Transferencia exitosa!").
 * v1.1.10: STOPGAP — no real LAFISE transfer .eml has been seen yet, only
 * the confirmed subject and sender (digital@notificaciones.lafise.com). A
 * live run showed the OLD behavior (silently falling through to the
 * generic backward-looking fallback, since this Type was never dispatched
 * to a dedicated extractor at all — see extractTransactionItems() above)
 * producing garbled mid-sentence merchant text. This deliberately extracts
 * the FIRST currency-prefixed amount in the body ONLY (not every one —
 * safer against accidentally grabbing an unrelated balance/limit figure
 * mentioned elsewhere) and labels the merchant generically rather than
 * guessing at a beneficiary field whose label hasn't been confirmed to
 * exist in this template. Known limitation: won't split a same-day bundle
 * of multiple LAFISE transfers into separate rows (unconfirmed whether
 * LAFISE even bundles transfers the way BHD does). Share a real LAFISE
 * transfer .eml to replace this with a proper extractor, the same way
 * BANESCO/BHD/POPULAR were built from real samples.
 * v1.1.12: tries a handful of common Spanish reference-number labels
 * (Referencia / No. Referencia / Número de confirmación / Confirmación) and
 * appends whatever it finds to the placeholder — "LAFISE Transfer (ref:
 * XXXXX)" instead of the fully generic label, per request, so the row is at
 * least traceable back to the real transfer even without a confirmed
 * beneficiary field. Silently keeps the plain placeholder if none of these
 * labels match (still a guess at the real template).
 */
function extractLAFISETransferTransactions(text) {
  // v1.1.19: first amount that is NOT an available-balance figure.
  const amtRe = /\b(?:RD|DOP|USD|US\$|EUR|COP)\s*\$?\s*([\d,]+\.?\d*)/gi;
  let m = null, cand;
  while ((cand = amtRe.exec(text)) !== null) {
    if (!isBalanceAmount(text, cand.index)) { m = cand; break; }
  }
  if (!m) return [];
  const amount = parseFloat(m[1].replace(/,/g, ''));
  if (amount === 0) return [];

  let merchant = 'LAFISE Transfer (details unconfirmed)';
  const refMatch = text.match(/(?:No\.?\s*Referencia|Referencia|Número de [Cc]onfirmaci[oó]n|Confirmaci[oó]n)\s*[:\-]?\s*([A-Za-z0-9\-\.]{4,30})/i);
  if (refMatch) {
    merchant = `LAFISE Transfer (ref: ${refMatch[1]})`;
  }

  return [{
    merchant: merchant,
    amount: amount,
    currency: detectCurrencyFromMatch(m[0]),
    context: text.substring(0, 400).replace(/\s+/g, ' ').trim(),
    ref: refMatch ? refMatch[1] : ''   // v1.1.23: bank's unique id → duplicate guard
  }];
}

/**
 * BANESCO consumo, approved ("Alerta de Consumo Banesco RD"). Real sample:
 *   "...presenta un consumo de RD$ 2,640.00, en SM BRAVO LA ESPERILLA y su
 *    estado es aprobada."
 * Amount comes BEFORE the merchant here. A declined variant exists with
 * different wording ("...ha sido rechazada...") that this pattern doesn't
 * match at all (no "y su estado es APROBADA/status" phrase in that
 * template) — declined ones are also caught by the whole-email
 * isDeclinedTransactionEmail() check upstream, this is defense-in-depth.
 */
function extractBANESCOConsumoTransactions(text) {
  const re = /consumo de\s*(?:RD|DOP|USD|US\$|EUR|COP)?\s*\$?\s*([\d,]+\.?\d*),?\s*en\s+([^\n]+?)\s+y su estado es\s+(\w+)/gi;
  const results = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    const declined = /RECHAZ|DECLIN/i.test(m[3]); // v1.1.19: flagged, not dropped (M1)
    results.push({ declined: declined, merchant: m[2].trim().substring(0, 50), amount: parseFloat(m[1].replace(/,/g, '')), currency: detectCurrencyFromMatch(m[0]), context: m[0].replace(/\s+/g, ' ').trim() });
  }
  return results;
}

/**
 * BANESCO transferencia ("Notificación de Transferencia Realizada"). Real
 * sample: "Monto: DOP545.00" ... later ... "Nombre del Beneficiario: AURORA
 * BEATRIZ PEREZ FERNANDEZ". Uses the beneficiary's name as "merchant" —
 * Category stays blank for Type=Transferencia regardless (Regla 3), this is
 * just for the description/record.
 */
function extractBANESCOTransferTransactions(text) {
  const re = /Monto:\s*(?:RD|DOP|USD|US\$|EUR|COP)?\s*\$?\s*([\d,]+\.?\d*)[\s\S]*?Nombre del Beneficiario:\s*([^\n]+)/gi;
  const results = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    // v1.1.23: the bank's own unique id ("No. Referencia: E000073.A1195451") — the same
    // transfer notice arrived twice in one thread, with identical content.
    const refM = text.substring(m.index, m.index + 800).match(/No\.?\s*Referencia:\s*([A-Za-z0-9.\-]{4,40})/i);
    results.push({ merchant: m[2].trim().substring(0, 50), amount: parseFloat(m[1].replace(/,/g, '')), currency: detectCurrencyFromMatch(m[0]), context: m[0].replace(/\s+/g, ' ').trim(), ref: refM ? refM[1] : '' });
  }
  return results;
}

/**
 * BHD consumo ("BHD Notificación de Transacciones") — a real table:
 *   Fecha | Moneda | Monto | Comercio | Estado | Tipo   (header, once)
 *   01/09/2026 08:38 am / RD / $118.40 / FCIA MEDICAR... / Aprobada / Compra  (one row per transaction)
 * The `g` flag naturally picks up every repeating 6-line row block, so this
 * is the one most likely to correctly handle the same-day-bundling case the
 * user described, once a real multi-row sample confirms the row layout
 * repeats the same way for row 2, 3, etc.
 */
/**
 * BHD consumo ("BHD Notificación de Transacciones") — a table (Fecha /
 * Moneda / Monto / Comercio / Estado / Tipo), one row per transaction.
 * v1.1.11: BUG FIX — required a literal `\n` between every field, which
 * broke on a real failing sample: "08/04/2026 10:20 am US $7.85 LIBERTARIO
 * COFFEE ROASTER Aprobada Compra" — fields separated by spaces, not
 * newlines (also the first confirmed sample with currency "US" rather than
 * "RD"). Switched every field separator from `\n` to `\s+` (matches either),
 * and anchored the boundary between the merchant and the following fields
 * to the known status words ("Aprobada"/"Rechazada"/"Declinada") instead of
 * a bare `\w+` — needed because a non-greedy merchant capture bordered by
 * `\s+` alone (ambiguous with the spaces INSIDE a multi-word merchant name
 * like "LIBERTARIO COFFEE ROASTER") would otherwise stop too early and
 * misattribute merchant words to Status/Type. Re-verified against both
 * previously-working real samples plus this new one — all three still
 * extract correctly.
 */
const BHD_ROW_STATUSES = 'Aprobada|Rechazada|Declinada|Reversada|Reversado|Reverso';

/**
 * v1.1.23: a REVERSAL row (real .eml, 06/05/2026: the same table row a few
 * seconds later with Estado "Reversada" and an EMPTY Comercio cell) used to
 * match nothing here — the status list didn't include it and the merchant
 * was mandatory — so the generic fallback saved it as a brand-new RD$488
 * "Unknown Merchant" expense. Now: the merchant is optional (and can never
 * swallow a status word, so a one-line multi-row body can't merge rows), and
 * a reversed row comes back with `reversal: true` plus `timeKey` — the
 * table's own date+time, identical on the purchase and its reversal — so
 * saveTransactions() can pair them (see resolveReversals()).
 */
function extractBHDConsumoTransactions(text) {
  const merchant = `((?:(?!\\b(?:${BHD_ROW_STATUSES})\\b)[^\\n])+?)`;
  const re = new RegExp(`(\\d{2}\\/\\d{2}\\/\\d{4}[^\\n]*?)\\s+(RD|DOP|USD?|EUR|COP)\\s+\\$?\\s*([\\d,]+\\.?\\d*)` +
                        `(?:\\s+${merchant})?\\s+(${BHD_ROW_STATUSES})\\s+(\\w+)`, 'gi');
  const results = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    const status = m[5];
    results.push({
      declined: /RECHAZ|DECLIN/i.test(status), // v1.1.19: flagged, not dropped (M1)
      reversal: /REVERS/i.test(status),
      timeKey: m[1].replace(/\s+/g, ' ').trim(),
      merchant: (m[4] || '').trim().substring(0, 50),
      amount: parseFloat(m[3].replace(/,/g, '')),
      currency: detectCurrencyFromMatch(m[2] + ' ' + m[0]),
      context: m[0].replace(/\s+/g, ' ').trim()
    });
  }
  return results;
}

/**
 * BHD transferencia ("Transacciones entre mis productos"). Real single-
 * transaction sample:
 *   Monto:
 *   RD$ 91.37
 *   Beneficiario:
 *   <ACCOUNT HOLDER NAME>
 * v1.1.5: a real multi-transaction BHD transfer email (several transfers
 * bundled same-day, confirmed from a live screenshot) broke this — the
 * garbled merchant values saved (": *DO60BCBH0000...", "ta desde. EC:MQ...")
 * show the [\s\S]*? bridge jumping to the WRONG "Beneficiario:" across
 * transaction boundaries when "Producto origen/destino" fields repeat
 * between entries. Added a sanity check that rejects an obviously-broken
 * capture (looks like an account number / "Producto" label fragment rather
 * than a name) instead of saving garbage — this loses those rows rather
 * than mislabeling them, which is the safer failure mode until a real
 * multi-transaction sample is available to fix the pattern properly.
 */
/**
 * v1.1.10: shared detector for an obviously-broken merchant/beneficiary
 * capture — an account-number fragment, a field label like "Producto
 * destino:", or a stray mid-sentence continuation — rather than a real
 * name. Real merchant/beneficiary text from every bank seen so far is
 * always UPPERCASE (UBER*EATS, SM BRAVO, FCIA MEDICAR, AURORA BEATRIZ
 * PEREZ...), so a capture starting lowercase is itself a strong signal of
 * a broken capture, not just the bank-specific fragments already known.
 * Used both to reject a capture live (BHD/LAFISE transfer extractors) and
 * to clean up already-saved rows (see cleanGarbledMerchants() below).
 * v1.1.18: BUG FIX — removed the 'HOLA ' substring check. It was added to
 * catch the fragment "Hola <name>, Acabas de realiza..." (a greeting
 * leaking into a broken capture), but it was a false-positive magnet: a
 * real BANESCO merchant "HOLA PLAZA LAS AMERICAS" (a real business inside
 * Plaza Las Américas mall) was correctly extracted by
 * extractBANESCOConsumoTransactions(), then NUKED into the placeholder by
 * this check during the very next auto-recategorize pass — which runs at
 * the end of EVERY monitor run, so this silently corrupted the row
 * immediately after it was saved correctly, not just once. Confirmed the
 * original fragment this rule was for is still caught without it — it
 * starts lowercase (`^[a-z]`) AND contains a bracket, both independently
 * sufficient. 'ACABAS DE' alone was already redundant with those too but
 * is left as-is (harmless, no known false-positive risk).
 */
const GARBLED_PLACEHOLDER = '(unparsed — see Email Subject)'; // v1.1.19: moved here from 04_sheetsWriter.gs

const REVERSAL_UNMATCHED = 'Reversal (original purchase not found)'; // v1.1.23

/**
 * v1.1.34: a POPULAR "Código Cash" withdrawal has no merchant column (Monto | Moneda | Fecha | Estatus), so the
 * status word after the date — "Aprobada" — was saved as the merchant (seen in a live Raw_POPULAR). A status word is
 * never a merchant: a Código Cash email gets a clear label, anything else the unreadable-merchant placeholder.
 * The category stays the fallback, Dining/Delivery + Entertainment + Other.
 */
const STATUS_WORD_MERCHANT = /^(APROBAD[AO]|RECHAZAD[AO]|DECLINAD[AO]|REVERSAD[AO]|REVERSO)$/i;
const CODIGO_CASH_MERCHANT = 'Código Cash (cash withdrawal)';
function fixStatusAsMerchant(merchant, subject, text) {
  if (!STATUS_WORD_MERCHANT.test(String(merchant || '').trim())) return merchant;
  return /C[OÓ]DIGO\s*CASH/i.test(String(subject || '') + ' ' + String(text || '')) ? CODIGO_CASH_MERCHANT : GARBLED_PLACEHOLDER;
}

function looksGarbled(text) {
  if (!text) return false;
  const t = String(text).trim();
  if (t.length === 0) return false;
  if (/^[a-z]/.test(t)) return true;       // real bank text is never lowercase-led
  if (/[\[\]]/.test(t)) return true;       // stray bracket from a truncated annotation
  if (/^[:*]/.test(t)) return true;
  const upper = t.toUpperCase();
  if (upper.includes('PRODUCTO') || upper.startsWith('TA DESDE') || /\*DO\d/.test(t) ||
      upper.includes('ACABAS DE')) return true;
  return false;
}

/**
 * BHD transferencia ("Transacciones entre mis productos"). Real confirmed
 * per-transaction field order:
 *   Producto origen: <acct> / Producto destino: <acct> / Descripción: /
 *   Monto: RD$ X / Beneficiario: <name> / Número de confirmación: ... /
 *   Fecha: ... / Tipo de transacción: ...
 * v1.1.15: BUG FIX — the garble rejection used `continue`, silently
 * dropping the item; if that was the only match, `results` came back
 * EMPTY, and parseEmailMessage() falls back to the generic
 * extractAllAmounts() (no garble protection), which re-extracted equally-
 * broken text. Fixed by never discarding — always keep the transaction.
 * v1.1.16: BUG FIX/CHANGE — a live log showed "Monto: *RD$ 2,950.00" with
 * an asterisk directly before the currency (Markdown-style bold from an
 * HTML→text conversion) that the amount regex didn't account for at all,
 * on top of the known multi-transaction-bundling issue. Also, garbled
 * captures strongly suggested some bundled transactions might be INTERNAL
 * transfers between the user's own BHD accounts, with no "Beneficiario:"
 * field at all. Rewritten to: (1) find every "Monto:" occurrence with its
 * own position, tolerating the asterisk; (2) for each, look FORWARD
 * (bounded to before the NEXT "Monto:") for its own "Beneficiario:"; (3) if
 * none is found there, look BACKWARD (bounded to after the PREVIOUS
 * "Monto:") for its own "Producto destino:" account — which structurally
 * comes before "Monto:", not after — and use its last 4 digits as a
 * traceable identifier instead of a name.
 * v1.1.18: CORRECTION from a real single-transaction .eml (self-transfer
 * between the user's own BHD accounts) — the v1.1.16 theory that an
 * internal/self transfer has NO "Beneficiario:" field was wrong. It DOES
 * have one; BHD just fills it with the account holder's own name
 * ("Beneficiario: <ACCOUNT HOLDER NAME>") rather than omitting the field. Verified
 * this real sample extracts correctly as-is (forward-look finds
 * "Beneficiario:" immediately, backward-look branch never triggers) — no
 * code change needed here, just correcting the inaccurate assumption in the
 * comment above so it doesn't mislead future debugging. The backward
 * "Producto destino:" fallback branch stays as defense-in-depth for
 * whatever templates it WAS built from, but its premise is unconfirmed —
 * still no real bundled multi-transaction .eml seen to verify that branch
 * or the field-order-repeats-per-row assumption against.
 */
function extractBHDTransferTransactions(text) {
  const results = [];
  const montoRe = /Monto:\s*\*?\s*(?:RD|DOP|USD|US\$|EUR|COP)?\s*\$?\s*([\d,]+\.?\d*)\*?/gi;
  let m;
  const matches = [];
  while ((m = montoRe.exec(text)) !== null) {
    matches.push({
      amount: parseFloat(m[1].replace(/,/g, '')),
      currency: detectCurrencyFromMatch(m[0]),
      index: m.index,
      end: m.index + m[0].length
    });
  }
  for (let i = 0; i < matches.length; i++) {
    const forwardEnd = (i + 1 < matches.length) ? matches[i + 1].index : text.length;
    const backwardStart = (i > 0) ? matches[i - 1].index : 0;
    const forwardBlock = text.substring(matches[i].end, forwardEnd);
    const backwardBlock = text.substring(backwardStart, matches[i].index);

    let merchant;
    const benefM = forwardBlock.match(/Beneficiario:\s*([^\n]+)/i);
    if (benefM && !looksGarbled(benefM[1].trim())) {
      merchant = benefM[1].trim().substring(0, 50);
    } else {
      // "Producto destino:" comes BEFORE "Monto:" for the same transaction
      // — take the LAST occurrence in the backward block (closest to this
      // transaction's own "Monto:"), not the first, so a bundle doesn't
      // pick up an earlier transaction's account instead of this one's.
      const destMatches = [...backwardBlock.matchAll(/Producto destino:\s*\*?[X\d]*?(\d{4})\D*?(?=\n|Descripci|Monto|$)/gi)];
      const destM = destMatches.length > 0 ? destMatches[destMatches.length - 1] : null;
      merchant = destM ? `BHD Transfer (to account ...${destM[1]})` : GARBLED_PLACEHOLDER;
    }
    const confM = forwardBlock.match(/N[úu]mero de confirmaci[óo]n:\s*([A-Za-z0-9\-]{4,40})/i); // v1.1.23
    results.push({
      merchant: merchant,
      amount: matches[i].amount,
      currency: matches[i].currency,
      context: text.substring(Math.max(0, matches[i].index - 100), Math.min(text.length, matches[i].end + 100)).replace(/\s+/g, ' ').trim(),
      ref: confM ? confM[1] : ''
    });
  }
  return results;
}

/**
 * POPULAR consumo ("Notificación de Consumo") — a compact table on one
 * visual row: Monto | Moneda | Fecha | Comercio | Estatus, e.g.
 *   US$14.27 [tab] Dólar estadounidense [tab] 22/05/2026 [tab] ELDORADO [tab] Aprobada
 * Lower confidence than the others — the real column separators (tabs vs.
 * spaces vs. newlines) may render slightly differently than in the one
 * sample seen; every parse is still logged with its context so this can be
 * tightened from a live run if needed.
 */
/**
 * POPULAR consumo ("Notificación de Consumo") — a compact table on one
 * visual row: Monto | Moneda | Fecha | Comercio | Estatus.
 * v1.1.9: BUG FIX — the currency prefix only accepted RD/DOP/USD/US$; a
 * real POPULAR sample in Colombian Pesos ("COP$33,000.00") never matched at
 * all, so the whole transaction silently disappeared (not just mis-tagged —
 * genuinely never extracted). This is what looked like "transactions within
 * a thread not being returned" — the COP-denominated messages in a thread
 * were failing outright while DOP/USD ones in the same thread succeeded.
 * Added EUR/COP to match the same 4-currency set used everywhere else.
 */
function extractPOPULARTransactions(text) {
  const re = /(?:RD|DOP|USD|US\$|EUR|COP)\s*\$?\s*([\d,]+\.?\d*)[^\n]*?\d{2}\/\d{2}\/\d{4}[^\n]*?[\t ]([A-ZÁÉÍÓÚÑ][^\t\n]{2,40}?)\s*[\t\n]+\s*(\w+)/gi;
  const results = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    const declined = /RECHAZ|DECLIN/i.test(m[3]); // v1.1.19: flagged, not dropped (M1)
    results.push({ declined: declined, merchant: m[2].trim().substring(0, 50), amount: parseFloat(m[1].replace(/,/g, '')), currency: detectCurrencyFromMatch(m[0]), context: m[0].replace(/\s+/g, ' ').trim() });
  }
  return results;
}

/**
 * v1.0.9: generic fallback — finds EVERY currency amount in the body and
 * pairs each with its best-effort PRECEDING context for merchant extraction.
 * Used only for BDI (no real sample yet) and as a safety net when a
 * bank-specific extractor above finds nothing. Note this only looks
 * BACKWARD from each amount, which is wrong for BANESCO/BHD/POPULAR (where
 * merchant comes AFTER the amount) — that's exactly why those three now have
 * dedicated extractors instead of relying on this.
 */
function extractAllAmounts(text, bankPattern) {
  const amountRegex = /\b(?:RD|DOP|USD|US\$|EUR|COP)\s*\$?\s*([\d,]+\.?\d*)/gi;
  const rawMatches = [];
  let m;
  while ((m = amountRegex.exec(text)) !== null) {
    if (isBalanceAmount(text, m.index)) continue; // v1.1.19: "Dispones de RD$ …" is a balance, not a charge
    rawMatches.push({ amount: parseFloat(m[1].replace(/,/g, '')), start: m.index, end: amountRegex.lastIndex });
  }

  const items = [];
  for (let i = 0; i < rawMatches.length; i++) {
    const cur = rawMatches[i];
    const windowStart = i === 0 ? Math.max(0, cur.start - 150) : rawMatches[i - 1].end;
    const context = text.substring(windowStart, cur.end).replace(/\s+/g, ' ').trim();
    const merchant = extractMerchantFromContext(context, bankPattern);
    items.push({ amount: cur.amount, merchant: merchant, currency: detectCurrencyFromMatch(context), context: context });
  }
  return items;
}

/**
 * v1.0.9: merchant extraction scoped to a single transaction's context
 * window (see extractAllAmounts) instead of the whole email body — this
 * also fixes part of the merchant="<first name>" bug, since the customer's name
 * in the greeting is now outside most items' windows (still inside the
 * FIRST item's window if the amount appears early in the body; unconfirmed
 * until a real sample shows exactly where).
 */
function extractMerchantFromContext(context, bankPattern) {
  try {
    const match = context.match(bankPattern.merchantPattern);
    if (match) {
      return match[1].trim().substring(0, 50);
    }
    const words = context.match(/\b[A-Z]{3,}\b/g);
    if (words && words.length > 0) {
      return words[0];
    }
    return "Unknown Merchant";
  } catch (error) {
    Logger.log("Error extracting merchant: " + error);
    return "Unknown";
  }
}

/**
 * Format date consistently
 */
/**
 * v1.1.2: BUG FIX — used to return date.toLocaleDateString('es-DO'), a plain
 * TEXT string. Google Sheets sometimes auto-detects that as a date and
 * sometimes doesn't (depends on the exact string + sheet locale), which is
 * exactly the kind of inconsistency that breaks SUMIFS/QUERY date-range
 * formulas — the Dashboard sheet depends on Date being a real date value in
 * every row. Now returns an actual Date object.
 * v1.1.5: BUG FIX — that Date object was built at MIDNIGHT
 * (new Date(y,m,d)), which is timezone-fragile: the Apps Script PROJECT's
 * time zone (Project Settings) and the SPREADSHEET's own time zone (File >
 * Settings) are two separate settings, and if they don't match, midnight in
 * one can land on the PREVIOUS calendar day when displayed in the other —
 * exactly what caused the Dashboard's month headers to read Dec-25...Nov-26
 * instead of Jan-26...Dec-26 (every date effectively shifted back). Building
 * at NOON instead absorbs any reasonable timezone mismatch (safe up to a
 * 12-hour gap) without crossing a day boundary. The real, permanent fix is
 * aligning the two timezone settings so they match (Project Settings > Time
 * zone = the same zone as the spreadsheet, e.g. America/Santo_Domingo) —
 * this is just a defensive floor under that.
 * NOTE: rows saved before this fix may still have a text date or a
 * midnight-based date — isDuplicate() normalizes both sides through
 * Utilities.formatDate() specifically so old and new rows still compare
 * correctly for duplicate-checking, but a date-range SUMIFS on the
 * Dashboard may still misplace a pre-fix row by one day if the two
 * timezones actually disagree. Re-entering that row's Date cell (or running
 * "🔁 Recategorize Saved Transactions", which does not touch dates, so this
 * still requires manual correction) can fix a specific bad row if noticed.
 */
function formatDate(date) {
  if (typeof date === 'string') {
    date = new Date(date);
  }
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0);
}

/**
 * v1.1.19: true when the amount at `index` is an available balance or limit
 * ("Dispones de RD$ 25,410.08", "Balance disponible", "Límite …"), not a
 * charge. The generic fallback and the LAFISE transfer extractor used to
 * pick these up as transactions.
 */
function isBalanceAmount(text, index) {
  const before = String(text).substring(Math.max(0, index - 40), index).toUpperCase();
  return /(DISPON|SALDO|BALANCE|L[ÍI]MITE)[^\n]{0,25}$/.test(before);
}

/**
 * Mark processed threads as read and label them.
 * v1.1.19 (M4):
 *  - threads with a message that could NOT be parsed (failedThreadIds) are
 *    left UNREAD — marking them read used to hide exactly the emails that
 *    need a look;
 *  - the "Procesado" label is fetched once (it used to be looked up inside
 *    the loop) and created if missing (it was never created, so it was never
 *    applied);
 *  - batched: GmailApp.markThreadsRead() / label.addToThreads() take up to
 *    100 threads per call instead of 2 calls per thread.
 */
const PROCESSED_LABEL = "Procesado";

function markEmailsAsProcessed(threads, failedThreadIds) {
  const failed = failedThreadIds || new Set();
  const done = threads.filter(t => !failed.has(t.getId()));
  let label = null;
  try {
    label = GmailApp.getUserLabelByName(PROCESSED_LABEL) || GmailApp.createLabel(PROCESSED_LABEL);
  } catch (error) {
    Logger.log("Could not get/create the \"" + PROCESSED_LABEL + "\" label: " + error);
  }
  for (let i = 0; i < done.length; i += 100) {
    const chunk = done.slice(i, i + 100);
    try {
      GmailApp.markThreadsRead(chunk);
      if (label) label.addToThreads(chunk);
    } catch (error) {
      Logger.log("Error marking emails: " + error);
    }
  }
  Logger.log("Marked " + done.length + " thread(s) as processed; left " +
             (threads.length - done.length) + " unread because something in them failed to parse.");
  return { marked: done.length, keptUnread: threads.length - done.length };
}

/**
 * Search bank emails in an INCLUSIVE date range ('yyyy-MM-dd' or 'yyyy/MM/dd').
 * v1.1.19 (C1, M5): exact epoch bounds from buildRangeBounds() (the end date
 * is now actually included) and paginated via gmailSearchAll(). Returns
 * { threads, range, capped } — `range` is passed on so messages outside it
 * can be skipped (see extractTransactionsFromThreads()). Search errors are no
 * longer swallowed into an empty result — they propagate to the entry point,
 * which shows them.
 * getUnreadTransactionCount() was removed (M13 — never called).
 */
function searchTransactionEmailsByDateRange(startDate, endDate, banksToTrack) {
  const range = buildRangeBounds(startDate, endDate);
  const selectedBanks = Object.keys(banksToTrack || {}).filter(b => banksToTrack[b] && BANK_PATTERNS[b]);
  const fromClauses = selectedBanks.map(b => BANK_PATTERNS[b].searchQuery).filter(Boolean);
  if (fromClauses.length === 0) {
    Logger.log("No bank in banksToTrack has a confirmed searchQuery — nothing to search.");
    return { threads: [], range: range, capped: false };
  }
  const query = `(${fromClauses.join(" OR ")}) after:${toEpochSeconds(range.start)} before:${toEpochSeconds(range.endExclusive)}`;
  Logger.log("Search query: " + query + "  [" + startDate + " … " + endDate + ", both days included]");
  const result = gmailSearchAll(query, MAX_THREADS_PER_RUN);
  Logger.log("Found " + result.threads.length + " thread(s)" +
             (result.capped ? " — hit the " + MAX_THREADS_PER_RUN + "-thread cap, split the range" : ""));
  return { threads: result.threads, range: range, capped: result.capped };
}
