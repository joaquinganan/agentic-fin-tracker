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
 *              Transferencia ("<Nombre>, ¡Transferencia exitosa!"): HTML only
 *              (empty plain part); real sample seen in v1.1.66, see
 *              extractLAFISETransferTransactions().
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
    // v1.1.51: incoming "Pagos al Instante" transfers come from a third address
    // v1.1.61: two more, from real samples: card payments (bancanet@notificaciones.lafise.com) and a second card-alert
    // template (notificacioneslafisedo@lafise.com.do). Neither was searched, so neither ever reached the sheet.
    searchQuery: 'from:notificaciones@bancolafise.com OR from:digital@notificaciones.lafise.com OR from:PagosAlInstanteMT103@lafise.com' +
      ' OR from:bancanet@notificaciones.lafise.com OR from:notificacioneslafisedo@lafise.com.do',
    extractors: { consumo: extractLAFISEAnyConsumo, transfer: extractLAFISEAnyTransfer, incoming: extractLAFISEIncomingTransactions,
      cardPayment: extractLAFISECardPaymentTransactions },
    keywords: ['LAFISE', 'LAFISE BANCO'],
    merchantPattern: /(?:en|en el|en\s+)([^\n]{10,50})/i
  },
  BANESCO: {
    name: 'BANESCO',
    fromDomain: 'banesco.com.do',
    // v1.1.51: the monthly savings statement (a PDF) — Banesco doesn't notify most incoming transfers
    searchQuery: 'from:notificaciones@banesco.com.do OR from:estadodecuenta@banesco.com.do',
    extractors: { consumo: extractBANESCOConsumoTransactions, transfer: extractBANESCOTransferTransactions,
      incoming: extractBANESCOIncomingTransactions },   // v1.1.62: "Notificación de Transferencia Recibida"
    statement: { subject: /ESTADO DE CUENTA DE AHORROS/i, parse: parseBanescoSavingsStatement },
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
    // v1.1.60: real samples — card purchases ("Notificacion de Consumos") and interbank transfers sent
    // v1.1.61: interbank transfers RECEIVED (same subject as the ones sent; the body says "Recibida"), and strict: an
    // email none of these recognizes goes to Unrecognized instead of being guessed (see STRICT_BANK_NOTE)
    extractors: { consumo: extractBDIConsumoTransactions, transfer: extractBDITransferTransactions, incoming: extractBDIIncomingTransactions,
      cardPayment: extractBDICardPaymentTransactions },
    detectType: (subject, text) => /Interbancaria\s+Recibida/i.test(text) ? 'Incoming' :
      /Tipo de Transacci[óo]n\s+Pago Tarjetas? de Cr[ée]dito/i.test(text) ? 'Card Payment' : null,   // v1.1.62
    // v1.1.62: "… Interbancaria - Completada" confirms a transfer already read from its "Comprobante" (same time)
    ignore: [/Transacci[óo]n Interbancaria\s*-\s*Completada/i],
    strict: true,
    keywords: ['BDI', 'BDI DIGITAL', 'BANCO BDI'],
    merchantPattern: /(?:en|en el|hacia|a)([^\n]{10,50})/i
  },
  // v1.1.60: two more banks, from real samples
  SCOTIABANK: {
    name: 'SCOTIABANK',
    fromDomain: 'scotiabank.com',                        // alertas@scotiabank.com, and *.scotiabank.com.do
    searchQuery: 'from:alertas@scotiabank.com OR from:scotiabank.com.do',
    extractors: { consumo: extractSCOTIABANKConsumoTransactions, incoming: extractSCOTIABANKIncomingTransactions },   // v1.1.61: incoming
    detectType: (subject, text) => /recibido un cr[ée]dito a su cuenta/i.test(text) ? 'Incoming' : null,
    strict: true,                                        // v1.1.61
    keywords: ['SCOTIABANK', 'SCOTIA'],
    merchantPattern: /\ben\s+(.{3,50}?)\s+con su/i
  },
  QIK: {
    name: 'QIK',
    fromDomain: '@qik.',                                 // notificaciones@qik.do, …@qik.com.do
    searchQuery: 'from:qik.do OR from:qik.com.do',
    extractors: { consumo: extractQIKConsumoTransactions },   // v1.1.61: also Código CASH withdrawals
    // v1.1.62: a Código CASH CREATED is not money out yet; the withdrawal comes in its own email ("… utilizado")
    ignore: [/C[óo]digo CASH para ti creado|C[óo]digo CASH para ti ha sido creado/i],
    strict: true,                                        // v1.1.61
    keywords: ['QIK'],
    merchantPattern: /\ben\s+(.{3,50}?)\s+con tu tarjeta/i
  },
  // v1.1.61: from real samples. Two senders: the app's receipts ("Recibo de la transacción": transfers sent and
  // TuEfectivo withdrawals) and notificaciones@ ("Notificaciones Banreservas": transfers received). Neither subject says
  // what happened, so the type is read from the body (detectType).
  BANRESERVAS: {
    name: 'BANRESERVAS',
    fromDomain: 'banreservas.com',
    searchQuery: 'from:notificaciones@banreservas.com OR from:NotificacionesTuBancoApp@banreservas.com',
    extractors: { consumo: extractBANRESERVASWithdrawalTransactions, transfer: extractBANRESERVASTransferTransactions,
      incoming: extractBANRESERVASIncomingTransactions },
    detectType: detectBANRESERVASType,
    strict: true,
    keywords: ['BANRESERVAS'],
    merchantPattern: /Destino:\s*([^,\n]{3,50})/i
  }
};

// v1.1.60: THE list of banks, in tab order (Raw_BDI before Raw_POPULAR, as asked in v1.1.5) — the Setup Wizard, View
// Config and the sheet order all come from it, so a bank is added in one place
const BANK_ORDER = ['LAFISE', 'BANESCO', 'BHD', 'BDI', 'POPULAR', 'SCOTIABANK', 'QIK', 'BANRESERVAS'];

/**
 * v1.1.61: a bank marked `strict` never goes to the generic amount guesser (extractAllAmounts). An email of its own that
 * none of its extractors recognizes (a welcome email, an account-opening notice, a template never seen) goes to
 * Unrecognized, where it can be looked at, instead of being saved as a purchase. Reported on BDI: transfers received
 * and a welcome email listing deposit limits were saved as spending.
 */
const STRICT_BANK_NOTE = 'Format not known yet: nothing guessed. Share its original email (.eml) to add it';

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
  // v1.1.63: account notices with no money in them — LAFISE's Bancanet (user blocked or about to be, temporary password,
  // alias or password changed, duplicate session, unblocked) and BDI's online registration
  'USUARIO DE BANCANET', 'TU USUARIO SE BLOQUEE', 'CONTRASEÑA TEMPORAL', 'CONTRASENA TEMPORAL', 'CONTRASEÑA EN BANCANET',
  'ALIAS DEL USUARIO', 'SESIÓN DUPLICADA', 'SESION DUPLICADA', 'DESBLOQUEO DE USUARIO', 'HACE TIEMPO QUE NO TE VEMOS',
  'SOLICITUD DE REGISTRO', 'CÓDIGO DE ACTIVACIÓN', 'CODIGO DE ACTIVACION', 'FIRMA DE CONTRATO', 'CLICK & SIGN',
  // v1.1.52: bank announcements (LAFISE's Pagos al Instante schedule comes from the incoming-transfers sender)
  'HORARIO TRANSFERENCIAS', 'HORARIOS TRANSFERENCIAS', 'HORARIO DE TRANSFERENCIAS', 'NOTA INFORMATIVA',
  // v1.1.23: incoming payroll deposit notices (POPULAR "Notificación Depósito de
  // Nómina") are income, not spending — they were reported as "Could not parse".
  'DEPÓSITO DE NÓMINA', 'DEPOSITO DE NOMINA',
  // v1.1.62: Banreservas' payroll notice (the salary is set in the Setup Wizard, as with POPULAR's above), its loan
  // reminder, and BHD's purchase-validation code
  'PAGO NÓMINA', 'PAGO NOMINA', 'PAGO DE NÓMINA', 'PAGO DE NOMINA', 'NOTIFICACIÓN DE BALANCES', 'NOTIFICACION DE BALANCES',
  'CÓDIGO DE VALIDACIÓN', 'CODIGO DE VALIDACION',
  // Statements / periodic summaries — not a single transaction
  'ESTADO DE CUENTA', 'FONDO DE INVERSIÓN', 'FONDO DE INVERSION',
  // Surveys / service announcements / events
  'QUEREMOS CONOCER TU OPINIÓN', 'QUEREMOS CONOCER TU OPINION', 'ENCUESTA',
  'ACOMPÁÑANOS', 'ACOMPAÑANOS', 'FORMA MÁS CONVENIENTE', 'FORMA MAS CONVENIENTE',
  // v1.1.61: welcome / account-opening emails (BDI's lists the first deposit and the account's limits as amounts). A real
  // transaction that happens to say "bienvenido" is still read: its bank's extractor finds it (see parseEmailMessage)
  'BIENVENIDO', 'BIENVENIDA', 'PRIMER DEPÓSITO', 'PRIMER DEPOSITO', 'APERTURA DE CUENTA', 'APERTURA DE TU CUENTA',
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
  // v1.1.51: money received. First: its body says "TRANSFERENCIA ... RECIBIDA" and "PAGOS AL INSTANTE", which are
  // Transfer keywords — without this, a received transfer would be read as one sent.
  'Incoming': ['TRANSFERENCIA ENTRANTE', 'PAGO AL INSTANTE RECIBIDO',   // v1.1.61: Scotiabank
    'NOTIFICACIÓN DE TRANSFERENCIA RECIBIDA', 'NOTIFICACION DE TRANSFERENCIA RECIBIDA'],   // v1.1.62: BANESCO
  'Transfer': [
    '¡TRANSFERENCIA EXITOSA!',              // LAFISE
    'NOTIFICACIÓN DE TRANSFERENCIA REALIZADA', 'NOTIFICACION DE TRANSFERENCIA REALIZADA', // BANESCO
    'TRANSACCIONES ENTRE MIS PRODUCTOS',    // BHD
    'TRANSACCIÓN INTERBANCARIA', 'TRANSACCION INTERBANCARIA',   // BDI (v1.1.60)
    'AVISO DE TRANSFERENCIA',                 // LAFISE online banking, to another bank (v1.1.63)
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
    'NOTIFICACIÓN DE PAGO DE TARJETA', 'NOTIFICACION DE PAGO DE TARJETA',   // LAFISE, from bancanet@ (v1.1.61)
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
    'DETALLE DE TRANSACCION TARJETA', 'DETALLE DE TRANSACCIÓN TARJETA',     // LAFISE, second card template (v1.1.61)
    'RETIRO CON CÓDIGO CASH', 'RETIRO CON CODIGO CASH',                     // QIK (v1.1.61)
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
  // v1.1.63: a reply or a forward is a conversation (with the bank's staff, or your own), never an alert
  if (/^\s*(RE|RV|FW|FWD|TR)\s*:/i.test(String(subject || ''))) {
    Logger.log('  ↳ a reply or forward — a conversation, not an alert');
    return true;
  }
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
  // v1.1.63: LAFISE's "Aviso de transferencia" reports its result — "Estado: Error" means no money moved (reported: two
  // failed rent payments would have been saved as transfers)
  const failed = flatText(bodyText).match(/\bEstado:\s*(Error|Rechazad[ao]|Fallid[ao]|Cancelad[ao]|Denegad[ao])\b/i);
  if (failed) {
    Logger.log('  ↳ the transfer failed: "Estado: ' + failed[1] + '"');
    return true;
  }
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
  if (bySubject === 'Transfer' || bySubject === 'Card Payment' || bySubject === 'Incoming') return bySubject;
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
    alreadySaved: 0,                 // v1.1.36: skipped without reading — already in Transactions
    filteredIds: []                  // v1.1.49: read, nothing to save (statements, promotions…) — for the read log
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
      // v1.1.55: statements are read again — each row is recognized by its reference, and a newer version may take more rows
      if (skipIds.has(message.getId()) && !isStatementEmail(message)) { stats.alreadySaved++; continue; }
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
        if (result.status === 'filtered') stats.filteredIds.push(message.getId());
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
 * v1.1.66: the email's text. Some emails carry an EMPTY plain-text part and everything in the HTML (LAFISE's
 * "¡Transferencia exitosa!", seen in a real .eml): getPlainBody() alone gave nothing and the email went to
 * Unrecognized as "Amount not found". Then the HTML is read as text. A plain part with any text is used as before.
 */
function emailPlainText(message) {
  let text = '';
  try { text = String(message.getPlainBody() || ''); } catch (e) { text = ''; }
  if (/\S/.test(text) || typeof message.getBody !== 'function') return text;
  try { return htmlToPlainText(message.getBody()) || text; } catch (e) { return text; }
}

const HTML_ENTITIES = { nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", aacute: 'á', eacute: 'é', iacute: 'í',
  oacute: 'ó', uacute: 'ú', ntilde: 'ñ', uuml: 'ü', Aacute: 'Á', Eacute: 'É', Iacute: 'Í', Oacute: 'Ó', Uacute: 'Ú', Ntilde: 'Ñ',
  Uuml: 'Ü', iexcl: '¡', iquest: '¿', ordm: 'º', ordf: 'ª', deg: '°', copy: '©', reg: '®', euro: '€', middot: '·', ndash: '–',
  mdash: '—', hellip: '…', laquo: '«', raquo: '»' };

/**
 * v1.1.66: HTML → text the way the extractors expect it from Gmail: one line per block or table cell, bold as *text*
 * (as getPlainBody() writes it), entities decoded; styles, scripts and comments dropped; no empty lines. Pure.
 */
function htmlToPlainText(html) {
  const code = n => (n > 0 && n <= 0x10FFFF ? String.fromCodePoint(n) : '');
  return String(html || '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<(style|script|head|title)\b[\s\S]*?<\/\1\s*>/gi, '')
    .replace(/<(b|strong)\b[^>]*>([\s\S]*?)<\/\1\s*>/gi, (m, tag, inner) => {
      const x = inner.replace(/<[^>]*>/g, '').trim();
      return x ? '*' + x + '*' : '';
    })
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/?(p|div|tr|td|th|li|ul|ol|table|tbody|thead|tfoot|h[1-6]|blockquote|center)\b[^>]*>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&#(\d+);/g, (m, n) => code(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (m, n) => code(parseInt(n, 16)))
    .replace(/&([a-z]+);/gi, (m, n) => (HTML_ENTITIES[n] !== undefined ? HTML_ENTITIES[n] : m))
    .split('\n').map(line => line.replace(/[ \t\r\f\v ]+/g, ' ').trim()).filter(Boolean).join('\n');
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
    const plainText = emailPlainText(message);   // v1.1.66: the HTML when the plain-text part is empty
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

    // v1.1.51: a bank statement (PDF): the incoming transfers the bank didn't notify
    if (bankPattern.statement && bankPattern.statement.subject.test(subject)) {
      return parseStatementEmail(message, bank, bankPattern.statement, rawCustomRules, stats);
    }

    // v1.1.62: emails a bank sends that are not a movement of their own (a confirmation of one already notified, a
    // code created but not yet used). Declared per bank; filtered before anything else, so a rescue can't revive them.
    const ignored = (bankPattern.ignore || []).find(rule => rule.test(subject + ' ' + flatText(plainText)));
    if (ignored) {
      stats.nonTransactional++;
      Logger.log("Skipped (" + bank + ": not a movement of its own): " + subject);
      return { items: [], status: 'filtered' };
    }

    if (isPromotionalEmail(subject, plainText)) {
      stats.promotional++;
      Logger.log("Skipped promotional email: " + subject);
      return { items: [], status: 'filtered' };
    }
    if (isNonTransactionalEmail(subject, plainText)) {
      // v1.1.60: a footer like "we'll never ask for your card's security code" isn't a security-code email. When the
      // bank's own extractor finds a real transaction (an amount), the email is read; otherwise it's filtered.
      const bankType = bankTransactionType(bankPattern, subject, plainText);   // v1.1.61: the bank's own reading first
      const real = bankPattern.extractors && Object.keys(bankPattern.extractors).length &&
        extractTransactionItems(bank, bankType, plainText).some(x => x && x.amount > 0);
      if (!real) {
        stats.nonTransactional++;
        Logger.log("Skipped non-transactional email: " + subject);
        return { items: [], status: 'filtered' };
      }
      Logger.log("Kept: a real transaction whose footer only mentions a non-transactional phrase — " + subject);
    }
    if (isDeclinedTransactionEmail(subject, plainText) && !hasApprovedRow(plainText)) {
      stats.declined++;
      Logger.log("Skipped declined/failed transaction: " + subject);
      return { items: [], status: 'filtered' };
    }

    const date = formatDate(message.getDate());
    const type = bankTransactionType(bankPattern, subject, plainText);   // v1.1.61
    const emailType = type;   // v1.1.62: items may override it, one by one

    let items = extractTransactionItems(bank, type, plainText);
    items.forEach(it => { it.merchant = fixStatusAsMerchant(it.merchant, subject, plainText); });   // v1.1.34
    const declinedRows = items.filter(it => it.declined).length;
    items = items.filter(it => !it.declined);
    if (items.length === 0 && declinedRows > 0) {
      stats.declined++;
      Logger.log("Skipped declined transaction (row-level status): " + subject);
      return { items: [], status: 'filtered' };
    }
    if (items.length === 0 && !bankPattern.strict) {   // v1.1.61: a strict bank's unknown email is never guessed
      items = extractAllAmounts(plainText, bankPattern);
    }
    if (items.length === 0) {
      stats.amountNotFound++;
      const snippet = plainText.substring(0, 700).replace(/\s+/g, ' ').trim();
      Logger.log("Amount not found | " + bank + " | " + subject + " | Body: \"" + snippet + "\"");
      return { items: [], status: 'failed', bank: bank, reason: bankPattern.strict ? STRICT_BANK_NOTE : 'Amount not found', snippet: snippet };
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
      // v1.1.62: an item can carry its own type (the tax row of a card payment is a Transfer, not the payment itself)
      const type = item.type || emailType;
      let category = findCustomRuleOverride(merchant, rawCustomRules);
      if (!category && type === 'Transaction' && !isReversal) {   // a reversal takes its original's category
        category = categorizeTransaction(merchant);
      }
      // v1.1.51: money received is saved NEGATIVE, like a reversal — given a category (a Custom Rule, or by hand) it
      // reduces what you spent there; money from your own account is Exclude
      const isIncoming = type === 'Incoming';
      if (item.own && (isIncoming || type === 'Transfer')) category = EXCLUDE_CATEGORY;   // v1.1.63: sent, too
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
        amount: isReversal || isIncoming ? -item.amount : item.amount,
        currency: currency,
        category: category,
        type: type,
        description: item.own && (isIncoming || type === 'Transfer') ? merchant + OWN_ACCOUNT_SUFFIX : merchant,
        reversal: isReversal,
        timeKey: item.timeKey || '',
        txRef: item.ref ? bank + ':' + item.ref : '',
        subject: subject,
        timestamp: new Date().toISOString(),
        messageId: baseMessageId + '_' + (itemIndex++),
        isCredit: isReversal || isIncoming || computeIsCredit(type, item.context),
        isCashback: type === 'Cashback'
      });
    }
    return { items: results, status: 'ok' };
  } catch (error) {
    stats.parseErrors++;
    Logger.log("Error parsing email: " + error + (error && error.stack ? " | " + error.stack : ""));
    let snippet = '';
    try { snippet = String(emailPlainText(message)).substring(0, 700).replace(/\s+/g, ' ').trim(); } catch (e) { snippet = ''; }
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
  // v1.1.63: "US" is dollars only as "US$" (or USD) — a LAFISE authorization code like "US34K7" read as US 12 made a
  // DOP purchase USD (reported, Uber Eats DOP 390)
  const m = matchText.match(/\b(RD|US(?=\s*\$)|DOP|USD|EUR|COP)\s*\$?\s*[\d,]/);
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
  const fn = type === 'Transfer' ? extractors.transfer : type === 'Incoming' ? extractors.incoming :
    (type === 'Card Payment' && extractors.cardPayment) || extractors.consumo;   // v1.1.61: a card payment's own template
  return fn ? fn(text) : [];
}

/**
 * v1.1.61: the type of an email, read by its bank first. Some banks use one subject for different things (Banreservas'
 * "Recibo de la transacción" is a transfer or a withdrawal; BDI's "Comprobante transacción Interbancaria" is sent or
 * received), so a bank can declare detectType(subject, text); when it returns nothing, the shared rules decide.
 */
function bankTransactionType(bankPattern, subject, text) {
  const own = bankPattern && bankPattern.detectType ? bankPattern.detectType(subject, flatText(text)) : null;
  return own || detectTransactionType(subject, text);
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
  const re = /Comercio\/Ciudad\/País:\s*([^\n]+?)\s*\n[\s\S]*?Monto:\s*\**\s*(RD|DOP|USD|US\$|EUR|COP)?\s*\$?\s*([\d,]+\.?\d*)/gi;
  const results = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    // v1.1.62: the currency written next to the amount wins; the whole context (every field between the merchant and
    // the amount) is only the fallback when the amount has none
    const currency = m[2] ? detectCurrencyFromMatch(m[2] + ' 0') : detectCurrencyFromMatch(m[0]);
    results.push({ merchant: m[1].replace(/^\*+|\*+$/g, '').trim().substring(0, 50), amount: parseFloat(m[3].replace(/,/g, '')), currency: currency, context: m[0].replace(/\s+/g, ' ').trim() });
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
 * v1.1.66: a real .eml is in hand: "Acabas de realizar una transferencia de USD 75.50 entre tus cuentas. … De paso te
 * dejamos tu número de referencia: <n>", all in the HTML (the plain-text part is empty — see emailPlainText()). The
 * amount and reference rules above read it; "entre tus cuentas" makes it yours (Exclude). Only a transfer between
 * your own accounts has been seen: one to someone else is read the same way, without a beneficiary.
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
    ref: refMatch ? refMatch[1] : '',   // v1.1.23: bank's unique id → duplicate guard
    own: /\bentre tus cuentas\b/i.test(flatText(text))   // v1.1.66: "…una transferencia de USD 75.50 entre tus cuentas."
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

function markEmailsAsProcessed(threads, failedThreadIds, opts) {
  opts = opts || {};
  const failed = failedThreadIds || new Set();
  let done = threads.filter(t => !failed.has(t.getId()));
  // v1.1.44: every thread in the range used to be marked again — read and labelled — even those marked by earlier
  // runs; over a year-long range that was minutes of Gmail calls (reported: a run killed right after "Parsed 1
  // transactions"). Gmail lists the ones still unread or without the label; only those are marked.
  let alreadyMarked = 0;
  if (opts.query && done.length) {
    try {
      const pending = new Set(gmailSearchAll(opts.query + ' {is:unread -label:' + PROCESSED_LABEL + '}', MAX_THREADS_PER_RUN)
        .threads.map(t => t.getId()));
      const before = done.length;
      done = done.filter(t => pending.has(t.getId()));
      alreadyMarked = before - done.length;
    } catch (error) {
      Logger.log("Could not list the threads still to mark (marking them all): " + error);
    }
  }
  const clock = opts.clock || (() => Date.now());
  let label = null;
  try {
    label = GmailApp.getUserLabelByName(PROCESSED_LABEL) || GmailApp.createLabel(PROCESSED_LABEL);
  } catch (error) {
    Logger.log("Could not get/create the \"" + PROCESSED_LABEL + "\" label: " + error);
  }
  let marked = 0;
  for (let i = 0; i < done.length; i += 100) {
    if (opts.deadline && clock() > opts.deadline) {   // v1.1.44: the rest stays for the next run
      Logger.log("⏸ Stopped marking at the time budget: " + (done.length - i) + " thread(s) left to mark");
      break;
    }
    const chunk = done.slice(i, i + 100);
    marked += chunk.length;
    try {
      GmailApp.markThreadsRead(chunk);
      if (label) label.addToThreads(chunk);
    } catch (error) {
      Logger.log("Error marking emails: " + error);
    }
  }
  const keptUnread = threads.filter(t => failed.has(t.getId())).length;
  Logger.log("Marked " + marked + " thread(s) as processed (" + alreadyMarked + " already were); left " + keptUnread +
             " unread because something in them failed to parse.");
  return { marked: marked, alreadyMarked: alreadyMarked, keptUnread: keptUnread };
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
  return { threads: result.threads, range: range, capped: result.capped, query: query,   // v1.1.44: query, to mark only what's pending
           untracked: findUntrackedBanks(selectedBanks, range) };                        // v1.1.61
}

/**
 * v1.1.61: banks NOT ticked in the Setup Wizard that do have emails in this range (one quick search each, one result
 * at most). Reported: Scotiabank purchases never showed up while the same email was read correctly in memory. A bank
 * added in a new version starts unticked in a setup saved before (v1.1.60), so its emails are never searched; the run
 * summary now says so instead of staying silent. Never stops the run.
 */
function findUntrackedBanks(selectedBanks, range) {
  const out = [];
  BANK_ORDER.filter(b => selectedBanks.indexOf(b) === -1 && BANK_PATTERNS[b] && BANK_PATTERNS[b].searchQuery).forEach(b => {
    try {
      const q = '(' + BANK_PATTERNS[b].searchQuery + ') after:' + toEpochSeconds(range.start) + ' before:' + toEpochSeconds(range.endExclusive);
      if (GmailApp.search(q, 0, 1).length) out.push(b);
    } catch (error) {
      Logger.log('Could not check ' + b + ' (not ticked): ' + error);
    }
  });
  if (out.length) Logger.log('⚠️ Emails found from banks not ticked in the Setup Wizard: ' + out.join(', '));
  return out;
}

/* ======================================================================
 * INCOMING TRANSFERS — v1.1.51
 * Money received: saved as NEGATIVE rows of Type "Incoming". With a category (a Custom Rule on the sender's name, or
 * typed in Incoming Transfers) it reduces what you spent in that category — a roommate's share of the rent lowers
 * Rent. Money from your own account is Exclude. Nothing else is decided for you.
 * ====================================================================== */
const OWN_ACCOUNT_SUFFIX = ' (own account)';

/** Two person names are the same holder when their first two names match (accents, case and signs aside — a
 *  statement writes "Ñ" as a space and cuts long names). Pure. */
function normalizedWords(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/[^A-Z ]/g, ' ').split(/\s+/).filter(Boolean);
}

function sameHolder(a, b) {
  const x = normalizedWords(a), y = normalizedWords(b);
  return x.length >= 2 && y.length >= 2 && x[0] === y[0] && x[1] === y[1];
}

/**
 * LAFISE "TRANSFERENCIA ENTRANTE APLICADA EXITOSAMENTE." (Pagos al Instante). Real sample (values changed):
 *   Nombre del cliente: / <HOLDER> / Número de cuenta: / … / Ordenante: / <SENDER> <cédula> / Monto total: / DOP 4500.00
 */
function extractLAFISEIncomingTransactions(text) {
  const flat = String(text || '').replace(/\s+/g, ' ');
  const m = flat.match(/Monto total:\s*(DOP|USD|RD\$|US\$)?\s*([\d,]+\.\d{2})/i);
  if (!m) return [];
  const sender = ((flat.match(/Ordenante:\s*(.+?)\s*(?:\d{6,}\s*)?Monto total:/i) || [])[1] || '').replace(/\s+\d+$/, '').trim();
  const holder = ((flat.match(/Nombre del cliente:\s*(.+?)\s*N[úu]mero de cuenta:/i) || [])[1] || '').trim();
  const currency = /USD|US\$/i.test(m[1] || '') ? 'USD' : 'DOP';
  return [{ amount: Number(m[2].replace(/,/g, '')), currency: currency, merchant: sender || 'Incoming transfer',
    own: !!(sender && holder && sameHolder(sender, holder)), context: 'incoming transfer' }];
}

/** A statement PDF (an attachment's blob) as text — read by the tracker itself (09_pdfText.gs), no Drive API.
 *  Apps Script gives the bytes signed (-128…127). */
function statementPdfText(blob) {
  const raw = blob.getBytes(), bytes = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) bytes[i] = raw[i] & 255;
  return pdfToText(bytes);
}

/**
 * Banesco savings statement (text of its PDF) → { holder, currency, rows, incoming, problems }. Each row is
 * "dd/mm/yyyy <description> <amount> <balance>"; whether it's a debit or a credit is read from the balance (the column
 * isn't in the text). Everything must add up to the statement's own totals, or nothing is taken from it. Pure.
 */
function parseBanescoSavingsStatement(text) {
  const flat = String(text || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ');
  const num = s => Number(String(s).replace(/,/g, ''));
  const total = re => { const m = flat.match(re); return m ? num(m[1]) : null; };
  const out = { holder: '', currency: /US\$|D[óo]lares|Moneda:\s*USD/i.test(flat) && !/RD\$/.test(flat) ? 'USD' : 'DOP', rows: [], incoming: [], problems: [] };
  const opening = total(/Balance mes anterior:?\s*(-?[\d,]+\.\d{2})/i), closing = total(/Balance al corte:?\s*(-?[\d,]+\.\d{2})/i);
  const credits = total(/Cr[ée]ditos del mes:?\s*([\d,]+\.\d{2})/i), debits = total(/D[ée]bitos del mes:?\s*([\d,]+\.\d{2})/i);
  const holder = flat.match(/(?:Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+\d{4}\s+([A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ .]{4,}?)\s+Cuenta/i)
    || flat.match(/Detalle de tus transacciones\s+([A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ .]{4,}?)\s+Transacciones en/i);
  out.holder = holder ? holder[1].trim() : '';
  if (opening === null || closing === null || credits === null || debits === null) {
    out.problems.push('the statement\'s totals (balances, credits, debits) were not found');
    return out;
  }
  const rowRe = /(\d{2})\/(\d{2})\/(\d{4}) (.+?) (-?[\d,]+\.\d{2}) (-?[\d,]+\.\d{2})(?= \d{2}\/\d{2}\/\d{4}| |$)/g;
  let prev = opening, sumC = 0, sumD = 0, m;
  while ((m = rowRe.exec(flat)) !== null) {
    const amount = num(m[5]), balance = num(m[6]), delta = +(balance - prev).toFixed(2);
    let direction = null;
    if (Math.abs(delta - amount) < 0.01) direction = 'credit';
    else if (Math.abs(delta + amount) < 0.01) direction = 'debit';
    if (!direction) { out.problems.push('row ' + m[1] + '/' + m[2] + ' "' + m[4] + '" does not follow the balance'); prev = balance; continue; }
    if (direction === 'credit') sumC += amount; else sumD += amount;
    out.rows.push({ day: m[3] + '-' + m[2] + '-' + m[1], description: m[4].trim(), amount: amount, balance: balance, direction: direction });
    prev = balance;
  }
  if (!out.rows.length) out.problems.push('no transactions found');
  if (Math.abs(sumC - credits) > 0.01) out.problems.push('credits add up to ' + sumC.toFixed(2) + ', the statement says ' + credits.toFixed(2));
  if (Math.abs(sumD - debits) > 0.01) out.problems.push('debits add up to ' + sumD.toFixed(2) + ', the statement says ' + debits.toFixed(2));
  if (out.rows.length && Math.abs(prev - closing) > 0.01) out.problems.push('the last balance is ' + prev.toFixed(2) + ', the statement says ' + closing.toFixed(2));
  if (out.problems.length) return out;
  // v1.1.55: EVERY credit is money received — descriptions vary ("Ach Ibanking", "Lbtr <name>", deposits, interest…);
  // what each one pays back is the user's call (a Custom Rule on its description, or typed in Incoming Transfers).
  // Yours when it's an LBTR from the holder or its description names the holder.
  const holderKey = out.holder ? normalizedWords(out.holder).slice(0, 2).join(' ') : '';
  out.rows.filter(r => r.direction === 'credit').forEach(r => {
    const lbtr = /^LBTR\b/i.test(r.description), ach = /^ACH\b/i.test(r.description);
    const name = lbtr ? r.description.replace(/^LBTR\s+/i, '').trim() : '';
    const own = !!out.holder && ((lbtr && sameHolder(name, out.holder)) || (holderKey.length > 3 && (' ' + normalizedWords(r.description).join(' ') + ' ').indexOf(' ' + holderKey + ' ') !== -1));
    out.incoming.push({ day: r.day, amount: r.amount, balance: r.balance, currency: out.currency, description: r.description,
      merchant: name || (ach ? 'ACH transfer (sender not in the statement)' : r.description), own: own });
  });
  return out;
}

/** A statement email: its PDF's incoming transfers as Incoming items (one per row, deduplicated per statement row). */
function parseStatementEmail(message, bank, statement, rawCustomRules, stats) {
  const subject = message.getSubject() || '';
  const pdf = (message.getAttachments() || []).find(a => /pdf/i.test(a.getContentType() || '') || /\.pdf$/i.test(a.getName() || ''));
  if (!pdf) return { items: [], status: 'failed', bank: bank, reason: 'Statement without a PDF', snippet: subject };
  let parsed;
  try {
    parsed = statement.parse(statementPdfText(pdf.copyBlob()));
  } catch (error) {
    stats.parseErrors++;
    return { items: [], status: 'failed', bank: bank, reason: 'Statement not read: ' + error, snippet: subject };
  }
  if (parsed.problems.length) {
    stats.parseErrors++;
    return { items: [], status: 'failed', bank: bank, reason: "Statement doesn't add up — nothing taken from it: " + parsed.problems.slice(0, 2).join('; '),
      snippet: subject + ' · ' + parsed.rows.length + ' row(s) read' };
  }
  if (!parsed.incoming.length) return { items: [], status: 'filtered' };   // nothing received by transfer this month
  const baseMessageId = message.getId();
  const items = parsed.incoming.map(x => {
    let category = findCustomRuleOverride(x.merchant, rawCustomRules);
    if (x.own) category = EXCLUDE_CATEGORY;
    const p = x.day.split('-').map(Number);
    return { date: new Date(p[0], p[1] - 1, p[2], 12), bank: bank, merchant: x.merchant, amount: -x.amount, currency: x.currency,
      category: category || '', type: 'Incoming', description: x.own ? x.merchant + OWN_ACCOUNT_SUFFIX : x.merchant,
      reversal: false, timeKey: '', txRef: bank + ':STMT:' + x.day + ':' + x.amount.toFixed(2) + ':' + x.balance.toFixed(2),
      subject: subject, timestamp: new Date().toISOString(), isCredit: true, isCashback: false,
      // v1.1.55: the row itself, not its position — positions moved once every credit was taken
      messageId: baseMessageId + '_' + x.day.replace(/-/g, '') + '-' + Math.round(x.amount * 100) + '-' + Math.round(x.balance * 100) };
  });
  Logger.log('Statement read | ' + bank + ' | ' + parsed.rows.length + ' rows, ' + items.length + ' incoming transfer(s)');
  return { items: items, status: 'ok' };
}

/** v1.1.55: a bank statement email (one this tracker reads: its bank has a statement reader and the subject matches). */
function isStatementEmail(message) {
  const subject = message.getSubject() || '';
  return Object.keys(BANK_PATTERNS).some(b => BANK_PATTERNS[b].statement && BANK_PATTERNS[b].statement.subject.test(subject));
}

/* ======================================================================
 * BDI, SCOTIABANK, QIK — v1.1.60 (real samples; fixtures with invented data)
 * Each reads the text flattened to single spaces, so it doesn't matter how Gmail's plain-text version splits the
 * email's table cells into lines.
 * ====================================================================== */
/**
 * v1.1.62: the text flattened to single spaces, WITHOUT Gmail's bold markers. getPlainBody() writes bold as *text*
 * ("*COMERCIO: *UBER*EATS …", "*RD$ 20.00*"), and that broke every extractor below on real emails while their
 * fixtures (written without the markers) passed: QIK purchases were filtered silently and LAFISE's second template and
 * BDI's transfers received went to Unrecognized. An asterisk that touches a space or the edge of the text is a marker
 * and goes; one inside a word stays (UBER*EATS, PedidosYa*Market, a masked 53****1234).
 */
function flatText(text) {
  return String(text || '').replace(/\s+/g, ' ').replace(/(^|\s)\*+/g, '$1').replace(/\*+(?=\s|$)/g, '').replace(/\s+/g, ' ').trim();
}
function moneyNumber(s) { return parseFloat(String(s).replace(/,/g, '')); }

function moneyCurrency(token, fallback) {
  const t = String(token || '').toUpperCase().replace(/\s+/g, '');
  if (t === 'RD$' || t === 'DOP') return 'DOP';
  if (t === 'US$' || t === 'USD') return 'USD';
  if (t === 'EUR' || t === '€') return 'EUR';
  return fallback || 'DOP';
}

/** BDI "Notificacion de Consumos": a table — Fecha | Moneda | Monto | Comercio | Estado — with one or more rows. */
function extractBDIConsumoTransactions(text) {
  const flat = flatText(text);   // v1.1.62: without Gmail's bold markers
  const re = /(\d{2}\/\d{2}\/\d{2,4}\s+\d{1,2}:\d{2})\s+(RD\$|US\$|DOP|USD|EUR)\s*([\d,]+\.\d{2})\s+(.+?)\s+(APROBADA|RECHAZADA|DECLINADA|REVERSADA|ANULADA)\b/gi;
  const out = [];
  let m;
  while ((m = re.exec(flat)) !== null) {
    if (m[5].toUpperCase() !== 'APROBADA') continue;   // only charges that went through
    out.push({ amount: parseFloat(m[3].replace(/,/g, '')), currency: moneyCurrency(m[2]), merchant: m[4].trim().substring(0, 50),
      timeKey: m[1], context: m[0] });
  }
  return out;
}

/**
 * BDI "Comprobante transacción Interbancaria": a transfer SENT ("[Salida]") to another bank — the amount, the
 * beneficiary's name and, as its own row, the tax and commission when there are any. Anything else (an incoming one
 * would say "[Entrada]") is left for Unrecognized rather than guessed.
 */
function extractBDITransferTransactions(text) {
  const flat = flatText(text);   // v1.1.62: without Gmail's bold markers
  if (!/\[\s*Salida\s*\]/i.test(flat)) return [];
  const amt = flat.match(/Monto\s+(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i);
  if (!amt) return [];
  const currency = moneyCurrency(amt[1]);
  const who = (flat.match(/Beneficiario\s+\**\d*\s*(.+?)\s+Banco Destino/i) || [])[1];
  const bank = (flat.match(/Banco Destino(?:\s*\/\s*C[óo]digo Swift)?\s+(.+?)\s+[A-Z0-9]{8,11}\s+Monto/i) || [])[1];
  const fee = s => { const f = flat.match(s); return f ? parseFloat(f[2].replace(/,/g, '')) : 0; };
  const fees = fee(/Impuesto[^$]*?(RD|US)\$\s*([\d,]+\.\d{2})/i) + fee(/Comisi[óo]n\s+(RD|US)\$\s*([\d,]+\.\d{2})/i);
  const out = [{ amount: parseFloat(amt[2].replace(/,/g, '')), currency: currency,
    merchant: (who || 'Transferencia interbancaria').trim().substring(0, 50), context: (bank ? 'to ' + bank.trim() : 'interbank transfer') }];
  if (fees > 0) out.push({ amount: +fees.toFixed(2), currency: currency, merchant: 'BDI — impuesto y comisión de transferencia', context: 'transfer fees' });
  return out;
}

/** Scotiabank "Autorización …": "por un monto de $25.50 USD en <merchant> con su Tarjeta de Crédito Scotiabank ***1234". */
function extractSCOTIABANKConsumoTransactions(text) {
  const flat = flatText(text);   // v1.1.62: without Gmail's bold markers
  const m = flat.match(/por un monto de\s+(RD\$|US\$|\$)?\s*([\d,]+\.\d{2})\s*(USD|DOP|EUR)?\s+en\s+(.+?)\s+con su\s+Tarjeta/i);
  if (!m) return [];
  const currency = m[3] ? moneyCurrency(m[3]) : (m[1] === 'US$' ? 'USD' : 'DOP');
  return [{ amount: parseFloat(m[2].replace(/,/g, '')), currency: currency, merchant: m[4].trim().substring(0, 50), context: m[0] }];
}


/* ======================================================================
 * v1.1.61: more formats, from real samples shared by another user (the fixtures keep their structure with invented
 * data). Same approach as above: each reads the text flattened to single spaces.
 * ====================================================================== */

/** LAFISE card purchases: the original template ("Comercio/Ciudad/País:"), or the second one below. */
function extractLAFISEAnyConsumo(text) {
  const first = extractLAFISETransactions(text);
  return first.length ? first : extractLAFISECardDetailTransactions(text);
}

/**
 * LAFISE "Detalle de Transaccion Tarjeta de Crédito" (notificacioneslafisedo@lafise.com.do):
 *   COMERCIO: <merchant> MONTO: 1,234.56 PESOS DOMI FECHA: dd/mm/yyyy
 * The currency comes as a word after the amount, cut short ("PESOS DOMI"). A currency word that isn't recognized is not
 * guessed: nothing is returned and the email goes to Unrecognized.
 */
function extractLAFISECardDetailTransactions(text) {
  const flat = flatText(text);
  const m = flat.match(/COMERCIO:\s*(.+?)\s*MONTO:\s*(RD\$|US\$|DOP|USD|EUR)?\s*([\d,]+\.\d{2})\s*(.*?)\s*FECHA:/i);
  if (!m) return [];
  let currency = m[2] ? moneyCurrency(m[2]) : null;
  if (!currency) {
    const word = m[4].toUpperCase();
    if (/^PESOS?\s*DOM|^RD|^DOP/.test(word)) currency = 'DOP';
    else if (/^D[OÓ]LAR|^US/.test(word)) currency = 'USD';
    else if (/^EURO/.test(word)) currency = 'EUR';
    else return [];
  }
  return [{ amount: moneyNumber(m[3]), currency: currency, merchant: m[1].trim().substring(0, 50), context: m[0] }];
}

/**
 * LAFISE "Notificación de pago de tarjeta de crédito" (bancanet@notificaciones.lafise.com): Concepto, Monto, Estado and
 * the card's last digits. Only a payment marked Exitoso counts; any other state is treated as declined.
 */
function extractLAFISECardPaymentTransactions(text) {
  const flat = flatText(text);
  const m = flat.match(/Concepto:\s*(.+?)\s*Monto:\s*(RD\$|US\$|DOP|USD|EUR)\s*([\d,]+\.\d{2})/i);
  if (!m) return [];
  const card = (flat.match(/N[úu]mero de tarjeta de cr[ée]dito:\s*(\S+)/i) || [])[1] || '';
  const state = (flat.match(/Estado:\s*(\S+)/i) || [])[1] || '';
  const ref = (flat.match(/Referencia:\s*(\d{4,})/i) || [])[1] || '';
  return [{ amount: moneyNumber(m[3]), currency: moneyCurrency(m[2]), merchant: (m[1].trim() + (card ? ' ' + card : '')).substring(0, 50),
    declined: !!state && !/^EXITOS[AO]$/i.test(state), ref: ref ? 'PAY:' + ref : '', context: 'card payment' }];
}

/**
 * BDI "Comprobante Transacción Interbancaria Recibida": who paid (Pagado Por), the amount. Yours when the payer and the
 * beneficiary are the same person.
 */
function extractBDIIncomingTransactions(text) {
  const flat = flatText(text);
  if (!/Interbancaria\s+Recibida/i.test(flat)) return [];
  const amt = flat.match(/Monto\s+(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i);
  if (!amt) return [];
  const payer = ((flat.match(/Pagado Por\s+(.+?)\s+Fecha Transacci[óo]n/i) || [])[1] || '').trim();
  const beneficiary = ((flat.match(/Nombre Beneficiario\s+(.+?)\s+No\. Autorizaci[óo]n/i) || [])[1] || '').replace(/[\s.]+$/, '').trim();
  const ref = (flat.match(/No\. Referencia\s+(\d+)/i) || [])[1] || '';
  return [{ amount: moneyNumber(amt[2]), currency: moneyCurrency(amt[1]), merchant: (payer || 'Interbank transfer received').substring(0, 50),
    own: !!(payer && beneficiary && sameHolder(payer, beneficiary)), ref: ref ? 'IN:' + ref : '', context: 'incoming transfer' }];
}

/**
 * Scotiabank "Pago al Instante recibido": "Ha recibido un crédito a su cuenta ***1234 por un Pago al Instante realizado
 * desde Banco X por valor de $1,250.00 DOP." The email names the sending bank, not the person.
 */
function extractSCOTIABANKIncomingTransactions(text) {
  const flat = flatText(text);
  const m = flat.match(/recibido un cr[ée]dito a su cuenta.*?(?:desde\s+(.+?)\s+)?por valor de\s+(RD\$|US\$|\$)?\s*([\d,]+\.\d{2})\s*(USD|DOP|EUR)?/i);
  if (!m) return [];
  const currency = m[4] ? moneyCurrency(m[4]) : (m[2] === 'US$' ? 'USD' : 'DOP');
  const how = /Pago al Instante/i.test(m[0]) ? 'Pago al Instante' : 'Crédito recibido';
  return [{ amount: moneyNumber(m[3]), currency: currency, merchant: (how + (m[1] ? ' desde ' + m[1].trim() : '')).substring(0, 50),
    own: false, context: 'incoming transfer' }];
}

/**
 * QIK. A card purchase ("Se hizo una transacción de RD$ 640.00 en <merchant> con tu tarjeta"), or a Código CASH
 * withdrawal ("El Código CASH … ha sido utilizado con éxito. Monto RD$ 2,000.00 Estatus Exitoso"). The email also shows
 * the card's available balance, never taken for the amount.
 */
function extractQIKConsumoTransactions(text) {
  const flat = flatText(text);
  const m = flat.match(/transacci[óo]n de\s+(RD\$|US\$|DOP|USD|EUR)\s*([\d,]+\.\d{2})\s+en\s+(.+?)\s+con tu tarjeta/i);
  if (m) {
    const when = (flat.match(/Fecha y hora\s+(\d{2}-\d{2}-\d{4}\s+\d{1,2}:\d{2}\s*[AP]M)/i) || [])[1] || '';
    return [{ amount: moneyNumber(m[2]), currency: moneyCurrency(m[1]), merchant: m[3].trim().substring(0, 50), timeKey: when, context: m[0] }];
  }
  if (/C[óo]digo CASH\b.*?utilizado/i.test(flat)) {   // v1.1.61
    const cash = flat.match(/Monto\s+(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i);
    if (!cash) return [];
    const state = (flat.match(/Estatus\s+(\S+)/i) || [])[1] || '';
    return [{ amount: moneyNumber(cash[2]), currency: moneyCurrency(cash[1]), merchant: CODIGO_CASH_MERCHANT,
      declined: !!state && !/^EXITOS[AO]$/i.test(state), context: 'cash withdrawal' }];
  }
  return [];
}

/* ---------- Banreservas (v1.1.61) ---------- */
/** The type from the body: "Transferencia Recibida" is money received; "Transacción: Transferencia …" one sent;
 *  "Transacción: Retiro …" a cash withdrawal. Anything else: not decided here (and, strict, not guessed). */
function detectBANRESERVASType(subject, flat) {
  if (/Transferencia Recibida/i.test(flat)) return 'Incoming';
  if (/Transacci[óo]n:\s*Transferencia/i.test(flat)) return 'Transfer';
  if (/Transacci[óo]n:\s*Retiro/i.test(flat)) return 'Transaction';
  return null;
}

function banreservasAmount(flat) {
  const m = flat.match(/Monto:\s*(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i);
  return m ? { amount: moneyNumber(m[2]), currency: moneyCurrency(m[1]) } : null;
}

/** Tax and commission, when the receipt has them (an empty "Comisión:" is nothing). */
function banreservasFees(flat) {
  const fee = re => { const f = flat.match(re); return f ? moneyNumber(f[1]) : 0; };
  return +(fee(/Comisi[óo]n:\s*(?:RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i) + fee(/Impuestos?:\s*(?:RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i)).toFixed(2);
}

function banreservasRef(flat) {
  return (flat.match(/N[úu]mero de (?:transacci[óo]n|referencia):\s*(\d{6,})/i) || [])[1] || '';
}

/** "Recibo de la transacción" with "Transacción: Transferencia a Tercero": the beneficiary (Destino, before the comma,
 *  without the "SR"/"SRA" the bank adds), and the tax and commission as their own row when not zero. */
function extractBANRESERVASTransferTransactions(text) {
  const flat = flatText(text);
  if (!/Transacci[óo]n:\s*Transferencia/i.test(flat)) return [];
  const amt = banreservasAmount(flat);
  if (!amt) return [];
  const who = ((flat.match(/Destino:\s*([^,]+?)\s*,/i) || [])[1] || '').replace(/^(SR|SRA|SRTA)\.?\s+/i, '').trim();
  const kind = ((flat.match(/Transacci[óo]n:\s*(.+?)\s+Origen:/i) || [])[1] || 'Transferencia').trim();
  const ref = banreservasRef(flat);
  const out = [{ amount: amt.amount, currency: amt.currency, merchant: (who || kind).substring(0, 50), ref: ref, context: kind }];
  const fees = banreservasFees(flat);
  if (fees > 0) out.push({ amount: fees, currency: amt.currency, merchant: 'Banreservas: impuesto y comisión de transferencia', context: 'transfer fees' });
  return out;
}

/** "Recibo de la transacción" with "Transacción: Retiro TuEfectivo" (cash sent to a phone): a cash withdrawal. The
 *  phone number is not kept. */
function extractBANRESERVASWithdrawalTransactions(text) {
  const flat = flatText(text);
  const kind = ((flat.match(/Transacci[óo]n:\s*(Retiro.*?)\s+Origen:/i) || [])[1] || '').trim();
  if (!kind) return [];
  const amt = banreservasAmount(flat);
  if (!amt) return [];
  const out = [{ amount: amt.amount, currency: amt.currency, merchant: (kind + ' (cash withdrawal)').substring(0, 50), ref: banreservasRef(flat), context: kind }];
  const fees = banreservasFees(flat);
  if (fees > 0) out.push({ amount: fees, currency: amt.currency, merchant: 'Banreservas: impuesto y comisión de retiro', context: 'withdrawal fees' });
  return out;
}

/** "Notificaciones Banreservas" / "Transferencia Recibida": who sent it (Origen). The email doesn't name the account
 *  holder, so whether it's your own money is left to you (a Custom Rule with your name → Exclude). */
function extractBANRESERVASIncomingTransactions(text) {
  const flat = flatText(text);
  if (!/Transferencia Recibida/i.test(flat)) return [];
  const amt = banreservasAmount(flat);
  if (!amt) return [];
  const sender = ((flat.match(/Origen:\s*(.+?)\s+Banco Origen:/i) || [])[1] || '').trim();
  return [{ amount: amt.amount, currency: amt.currency, merchant: (sender || 'Transferencia recibida').substring(0, 50), own: false,
    context: 'incoming transfer' }];
}

/* ======================================================================
 * v1.1.62: BDI card payment from the account; BANESCO transfer received.
 * ====================================================================== */
/**
 * BDI "Comprobante de Transacción" with "Tipo de Transacción Pago Tarjetas de Crédito" (real sample, Unrecognized):
 *   Monto Transferido RD$4,000.00 … Impuesto 0.20% RD$8.00 Monto RD$4,008.00
 * The amount paid to the card is a Card Payment (Exclude: it's the purchases already counted, being paid); the tax is
 * a cost of its own, a separate row typed Transfer like the other banks' transfer taxes.
 */
function extractBDICardPaymentTransactions(text) {
  const flat = flatText(text);
  if (!/Pago Tarjetas? de Cr[ée]dito/i.test(flat)) return [];
  const paid = flat.match(/Monto Transferido\s+(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i) || flat.match(/\bMonto\s+(RD\$|US\$|DOP|USD)\s*([\d,]+\.\d{2})/i);
  if (!paid) return [];
  const currency = moneyCurrency(paid[1]);
  const fee = re => { const f = flat.match(re); return f ? moneyNumber(f[1]) : 0; };
  const fees = +(fee(/Impuesto[^$]*?(?:RD|US)\$\s*([\d,]+\.\d{2})/i) + fee(/Comisi[óo]n\s+(?:RD|US)\$\s*([\d,]+\.\d{2})/i)).toFixed(2);
  const ref = (flat.match(/No\.\s*Ref(?:erencia)?\.?\s+(\d{4,})/i) || [])[1] || '';
  const out = [{ amount: moneyNumber(paid[2]), currency: currency, merchant: 'Pago Tarjetas de Crédito (BDI)', ref: ref ? 'PAY:' + ref : '',
    context: 'card payment' }];
  if (fees > 0) out.push({ amount: fees, currency: currency, merchant: 'BDI: impuesto y comisión de pago de tarjeta', type: 'Transfer', context: 'payment fees' });
  return out;
}

/**
 * BANESCO "Notificación de Transferencia Recibida" (reported: saved as a transfer sent, i.e. spending). No sample of
 * its own yet; the sent one ("… Realizada") is labeled fields ("Monto: DOP1,000.00 … Nombre del Beneficiario: X
 * Concepto: LBTR <sender>"), and the received one was read by that same extractor (the amount and the beneficiary
 * came out right), so the same labels are read here. Who sent it: a sender field if there is one, else the name in
 * "Concepto: LBTR <name>"; yours when it's the beneficiary. Without either, the sender is not guessed.
 */
function extractBANESCOIncomingTransactions(text) {
  const raw = String(text || '');
  const flat = flatText(raw);
  const amt = flat.match(/Monto:\s*(RD\$|US\$|DOP|USD|EUR)?\s*([\d,]+\.\d{2})/i);
  if (!amt) return [];
  // each field is on its own line, as in the email sent ("Nombre del Beneficiario: X" ⏎ "Concepto: …")
  const field = re => { const m = raw.match(re); return m ? flatText(m[1]) : ''; };
  const beneficiary = field(/Nombre del Beneficiario:[ \t*]*([^\n]+)/i);
  let sender = field(/(?:^|\n)[ \t*]*(?:Nombre del (?:Ordenante|Originante|Originador|Remitente|Emisor)|Ordenante|Remitente|Originador|Enviad[oa] por):[ \t*]*([^\n]+)/i);
  if (!sender) sender = field(/Concepto:[ \t*]*LBTR\s+([^\n]+)/i);
  const fromBank = field(/Banco (?:Ordenante|Origen|Emisor|Remitente):[ \t*]*([^\n]+)/i);
  const ref = (flat.match(/No\.?\s*Referencia:\s*([A-Za-z0-9.\-]{4,40})/i) || [])[1] || '';
  const merchant = sender || ('Transferencia recibida' + (fromBank ? ' desde ' + fromBank : ''));
  return [{ amount: moneyNumber(amt[2]), currency: amt[1] ? moneyCurrency(amt[1]) : 'DOP', merchant: merchant.substring(0, 50),
    own: !!(sender && beneficiary && sameHolder(sender, beneficiary)), ref: ref ? 'IN:' + ref : '', context: 'incoming transfer' }];
}

/* ======================================================================
 * LAFISE online banking "Aviso de transferencia en banco local" — v1.1.63 (real sample; fixture with invented data)
 *   Cuenta de origen Titular: <HOLDER> … Cuenta destino Titular: <NAME> … Concepto: <text> Monto: 12,500.00 DOP
 *   … Resultado Estado: <status> Referencia: <n>
 * A failed one ("Estado: Error") never gets here: isDeclinedTransactionEmail filters it. To your own name, it's Exclude.
 * ====================================================================== */
function extractLAFISELocalBankTransfer(text) {
  const flat = flatText(text);
  if (!/Aviso de transferencia/i.test(flat)) return [];
  const amt = flat.match(/Monto:\s*([\d,]+\.\d{2})\s*(DOP|USD|EUR)\b/i);
  if (!amt) return [];
  const origin = ((flat.match(/Cuenta de origen\s+Titular:\s*(.+?)\s+N[úu]mero de cuenta:/i) || [])[1] || '').trim();
  const dest = ((flat.match(/Cuenta destino\s+Titular:\s*(.+?)\s+N[úu]mero de cuenta:/i) || [])[1] || '').trim();
  const concept = ((flat.match(/Concepto:\s*(.+?)\s+Monto:/i) || [])[1] || '').trim();
  const ref = (flat.match(/Referencia:\s*(\d{4,})/i) || [])[1];
  return [{ amount: parseFloat(amt[1].replace(/,/g, '')), currency: amt[2].toUpperCase(), merchant: (dest || 'Transferencia').substring(0, 50),
    context: concept ? 'Concepto: ' + concept : 'local bank transfer', own: !!(origin && dest && sameHolder(origin, dest)),
    ref: ref ? 'OUT:' + ref : '' }];
}

/** LAFISE transfers sent: the online-banking notice, or the app's "¡Transferencia exitosa!". */
function extractLAFISEAnyTransfer(text) {
  const notice = extractLAFISELocalBankTransfer(text);
  return notice.length ? notice : extractLAFISETransferTransactions(text);
}

/* ======================================================================
 * UBER RIDES — v1.1.67 (real LAFISE alerts and Uber receipts, Sept–Oct 2026; fixtures with invented data)
 * Uber authorizes an estimate on the card when a ride is requested. When the fare changes it authorizes the final
 * amount and the estimate is released; when the final fare is lower it may charge it with no new alert. A request
 * that never becomes a trip leaves an authorization too. The bank emails every authorization as a purchase, so a ride
 * was saved twice (estimate + final), or at the estimate, and the card statement shows only what Uber charged.
 * Uber's trip receipt ("Your <day> <time> trip with Uber") says what each card was charged and when the ride was
 * requested: the ride alerts around it are checked against it (matchRideReceipts) and the sheet is corrected
 * (reconcileRideReceipts, 04_sheetsWriter.gs). Uber Eats is charged once and is not touched.
 * ====================================================================== */
const RIDE_RECEIPT_QUERY = 'from:noreply@uber.com subject:"trip with Uber"';
const RIDE_RECEIPT_SUBJECT = /trip with Uber/i;
const RIDE_MERCHANT_RE = /\bUBER\s*\*?\s*(?:RIDES?|TRIP)\b|\bUBR\s*\*|PENDING\.UBER/i;
const RECEIPT_MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };

/**
 * The receipt's text → { requested: Date (the time printed at its top, the rider's local time), payments: [{ last4,
 * bank, currency, amount }] }, or null when it isn't a trip receipt. Only card payments whose bank (in parentheses,
 * "Mastercard ••••1234 (LAFISE)") is one the tracker reads; Uber Cash, cash and other cards are left out. Works on the
 * text by lines or flattened to one line. Pure — see tests/.
 */
function parseUberTripReceipt(text) {
  const flat = String(text || '').replace(/[\u00a0\u202f]/g, ' ').replace(/\s+/g, ' ');
  const when = /\b([A-Z][a-z]{2})[a-z]*\.? (\d{1,2}), (\d{4}) ,? ?(\d{1,2}):(\d{2}) ?([AP])\.? ?M\b/i.exec(flat);
  if (!when || RECEIPT_MONTHS[when[1].toLowerCase()] === undefined) return null;
  const hour = Number(when[4]) % 12 + (/p/i.test(when[6]) ? 12 : 0);
  const requested = new Date(Number(when[3]), RECEIPT_MONTHS[when[1].toLowerCase()], Number(when[2]), hour, Number(when[5]));
  const payments = [];
  const re = /(?:•|\*){2,} ?(\d{4}) ?\(([^)]+)\) ?(RD\$|DOP|US\$|USD) ?([\d,]+\.\d{2})/gi;
  let m;
  while ((m = re.exec(flat)) !== null) {
    const label = m[2].toUpperCase().replace(/\s+/g, '');
    const bank = BANK_ORDER.find(b => label.indexOf(b) !== -1);
    if (bank) payments.push({ last4: m[1], bank: bank, currency: moneyCurrency(m[3]), amount: moneyNumber(m[4]) });
  }
  return { requested: requested, payments: payments };
}

/**
 * Ride alerts vs. trip receipts. alerts: [{ id, bank, currency, amount, at (ms, the alert email's time), state: ''
 * | 'charged' | 'hold' (settled by an earlier run) }]; receipts: [{ bank, currency, amount (that card's charge),
 * requested, sent (ms) }]. Returns { decisions: { id: { kind: 'charge' } | { kind: 'hold' } | { kind: 'adjust', amount,
 * was } }, unmatched: [ids] } — decisions only for alerts not settled yet. Pure — see tests/.
 *  1. the charge: an alert of the receipt's amount from the request (−2 min) to the receipt (+15 min), the one
 *     closest to the receipt (the final authorization comes with it);
 *  2. no alert of that amount: the ride's own authorization (the alert closest to the request) was charged at the
 *     receipt's amount — corrected;
 *  3. every other alert from 30 min before the request up to the charge: the estimate, or requests that never
 *     became a trip — holds.
 * An alert near no receipt is left as it is (unmatched): a missing receipt never removes a real charge.
 */
const RIDE_WINDOW = { beforeRequest: 2, afterReceipt: 15, attemptsBefore: 30 };   // minutes
function matchRideReceipts(alerts, receipts) {
  const M = 60000, used = {}, decisions = {};
  const seen = {};
  const list = receipts.filter(r => {   // Uber sends some receipts twice
    const k = r.bank + '|' + r.currency + '|' + Math.round(r.amount * 100) + '|' + r.requested;
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  }).sort((a, b) => a.requested - b.requested);
  const fits = (a, r) => a.bank === r.bank && a.currency === r.currency &&
    a.at >= r.requested - RIDE_WINDOW.beforeRequest * M && a.at <= r.sent + RIDE_WINDOW.afterReceipt * M;
  const charge = [];
  list.forEach((r, i) => {
    let best = null;
    alerts.forEach(a => {
      if (used[a.id] || !fits(a, r) || Math.abs(a.amount - r.amount) >= 0.005) return;
      if (!best || Math.abs(a.at - r.sent) < Math.abs(best.at - r.sent)) best = a;
    });
    if (!best) return;
    used[best.id] = true;
    charge[i] = best;
    if (best.state !== 'charged') decisions[best.id] = { kind: 'charge' };
  });
  list.forEach((r, i) => {
    if (charge[i]) return;
    let best = null;
    alerts.forEach(a => {
      if (used[a.id] || a.state || !fits(a, r)) return;
      if (!best || Math.abs(a.at - r.requested) < Math.abs(best.at - r.requested)) best = a;
    });
    if (!best) return;
    used[best.id] = true;
    charge[i] = best;
    decisions[best.id] = { kind: 'adjust', amount: r.amount, was: best.amount };
  });
  list.forEach((r, i) => {
    if (!charge[i]) return;
    alerts.forEach(a => {
      if (used[a.id] || a.bank !== r.bank || a.currency !== r.currency ||
          a.at < r.requested - RIDE_WINDOW.attemptsBefore * M || a.at > charge[i].at) return;
      used[a.id] = true;
      if (a.state !== 'hold') decisions[a.id] = { kind: 'hold' };
    });
  });
  return { decisions: decisions, unmatched: alerts.filter(a => !a.state && !used[a.id]).map(a => a.id) };
}
