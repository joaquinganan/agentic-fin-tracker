/**
 * INVESTMENTS — v1.1.29
 *
 * The same idea as the bank side: every movement becomes a row of a ledger,
 * whatever its source, and values are computed from it.
 *
 *   Investment Ledger  every movement: Buy / Sell / Dividend / Deposit /
 *                      Withdrawal / Fee, plus Snapshot (positions as shown by
 *                      the broker on a date) and Valuation (the value of an
 *                      account tracked by balance, e.g. a fund or pension)
 *   Holdings           positions computed from the ledger and valued with
 *                      GOOGLEFINANCE, plus the other accounts and a total
 *   Investment Accounts  accounts, and the keyword that identifies the bank
 *                      transfers that fund each one
 *
 * Sources:
 *   · HAPI emails (no-reply@hapi.trade): executed orders (ticker, quantity,
 *     average price, cost and the per-order fee) and dividends (net amount).
 *     "Deposit Completed" has no amount — deposits come from the bank side.
 *   · Bank transfers already in Transactions whose beneficiary contains an
 *     account's deposit keyword → Deposit rows (the transfer itself stays in
 *     whatever category you give it, e.g. "Exclude", so it isn't spending).
 *   · Rows you type: Snapshot, Valuation, anything a broker doesn't email.
 *
 * Positions start from each account's most recent Snapshot and apply the
 * movements dated AFTER that day — so re-reading old emails never counts a
 * trade twice. Earlier movements stay in the ledger as history (dividends,
 * fees and contributions of the year still count them).
 */
const INVESTMENT_LEDGER_SHEET = 'Investment Ledger';
const HOLDINGS_SHEET = 'Holdings';
const INVESTMENT_ACCOUNTS_SHEET = 'Investment Accounts';
const LEDGER_HEADERS = ['Date', 'Account', 'Type', 'Ticker', 'Quantity', 'Price', 'Amount', 'Fee', 'Currency', 'Source', 'Notes', 'Id'];
const LG = { DATE: 0, ACCOUNT: 1, TYPE: 2, TICKER: 3, QTY: 4, PRICE: 5, AMOUNT: 6, FEE: 7, CURRENCY: 8, SOURCE: 9, NOTES: 10, ID: 11 };
const LEDGER_TYPES = ['Buy', 'Sell', 'Dividend', 'Deposit', 'Withdrawal', 'Fee', 'Snapshot', 'Valuation'];
const ACCOUNTS_HEADERS = ['Account', 'Kind', 'Deposit keyword', 'Notes'];
const DEFAULT_INVESTMENT_ACCOUNTS = [
  ['HAPI', 'Broker', 'OUROSR', "Orders and dividends come from HAPI's emails; deposits from bank transfers whose " +
    "beneficiary contains the keyword (HAPI's collection account in the DR)."]
];
const BROKER_PATTERNS = {
  HAPI: { searchQuery: 'from:no-reply@hapi.trade', sender: 'hapi.trade', parse: parseHapiMessage }
};
// GOOGLEFINANCE symbols for tickers it doesn't know as they are. v1.1.30: any crypto pair — 6+ letters ending in USD
// (ETHUSD, SHIBUSD…) — is CURRENCY:<pair>; US stock tickers have at most 5 letters, so none is caught by mistake.
const PRICE_SYMBOLS = {};
function priceSymbol(ticker) {
  if (PRICE_SYMBOLS[ticker]) return PRICE_SYMBOLS[ticker];
  return /^[A-Z]{3,}USD$/.test(ticker) && ticker.length >= 6 ? 'CURRENCY:' + ticker : ticker;
}
/**
 * v1.1.31: GOOGLEFINANCE doesn't price crypto pairs — ETHUSD and SHIBUSD returned errors in a live sheet (a known
 * limitation). Crypto comes from Coinbase's public spot price (no key), fetched by the script on every refresh;
 * stocks and ETFs stay on GOOGLEFINANCE. Crypto prices therefore update when the tracker runs, not continuously.
 */
const CRYPTO_PRICE_URL = pair => 'https://api.coinbase.com/v2/prices/' + pair + '/spot';
function isCryptoPair(ticker) { return /^[A-Z]{3,}USD$/.test(ticker) && ticker.length >= 6; }
// v1.1.31: the sanity check only trusts a RECENT last known price — after months without trades a 50%+ move is real
// (the first version would have frozen such a position at its old price, labelled "check symbol").
const SANITY_MAX_AGE_DAYS = 90;
function daysBetween(fromKey, toKey) {
  const d = k => { const p = String(k).split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]); };
  return Math.round((d(toKey) - d(fromKey)) / 86400000);
}

/** Coinbase spot prices for the crypto pairs among `tickers` → { TICKER: price }; missing on any failure. */
function fetchCryptoPrices(tickers) {
  const pairs = tickers.filter(isCryptoPair).filter((t, i, a) => a.indexOf(t) === i);
  const out = {};
  if (!pairs.length) return out;
  let responses;
  try {
    responses = UrlFetchApp.fetchAll(pairs.map(t => ({ url: CRYPTO_PRICE_URL(t.slice(0, -3) + '-USD'), muteHttpExceptions: true })));
  } catch (error) {
    Logger.log('Crypto prices unavailable: ' + error);
    return out;
  }
  responses.forEach((response, i) => {
    try {
      if (response.getResponseCode() !== 200) return;
      const amount = Number(JSON.parse(response.getContentText()).data.amount);
      if (amount > 0) out[pairs[i]] = amount;
    } catch (error) {
      Logger.log('Crypto price for ' + pairs[i] + ' unreadable: ' + error);
    }
  });
  return out;
}

// prices under a cent (e.g. SHIB) keep 8 decimals instead of showing US$0.00
const USD_PRICE_FORMAT = '[<0.01]"US$"0.00000000;"US$"#,##0.00';
// A live price this far from the last known one is treated as a wrong symbol (e.g. a ticker Google maps to another security)
const PRICE_SANITY = 0.5;

const invNumber = value => {
  if (typeof value === 'number') return value;
  const n = Number(String(value || '').replace(/[,$\s]/g, ''));
  return isNaN(n) ? 0 : n;
};

/** Visible text of an email, whitespace collapsed (plain part, or the HTML stripped). */
function investmentMessageText(message) {
  let text = '';
  try { text = message.getPlainBody() || ''; } catch (error) { text = ''; }
  if (!/[A-Za-z]{3}/.test(text) && typeof message.getBody === 'function') {
    text = String(message.getBody() || '').replace(/<(style|script|head)[\s\S]*?<\/\1>/gi, ' ')
      .replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&');
  }
  return text.replace(/[\u200b\u200c\u034f\u00a0]/g, ' ').replace(/\s+/g, ' ').trim();
}

/**
 * One HAPI email → { kind: 'event', event } | { kind: 'skipped', reason } |
 * { kind: 'unparsed', reason }. An order is only accepted when its own numbers
 * agree (quantity × average price = cost), so a changed template can't write
 * a wrong position silently. Pure — see tests/.
 */
function parseHapiMessage(message) {
  const subject = String(message.getSubject() || '');
  const text = investmentMessageText(message);
  const base = { account: 'HAPI', currency: 'USD', source: 'email', id: 'gmail:' + message.getId(), fee: 0, notes: '' };

  if (/Order Executed/i.test(subject) || /Your order has been executed/i.test(text)) {
    const side = (text.match(/Buy\/Sell:\s*(Buy|Sell)\b/i) || [])[1];
    const ticker = (text.match(/Ticker:\s*([A-Z0-9.\-]{1,12})\b/) || [])[1];
    // the quantity is followed by the order fee, in a row whose label HAPI's template leaves empty
    const qty = text.match(/Quantity(?: Shares)?:\s*([\d,]*\.?\d+)(?:\s*US\$\s*([\d,]*\.?\d+))?/i);
    const price = (text.match(/Average price:\s*US\$\s*([\d,]*\.?\d+)/i) || [])[1];
    const cost = (text.match(/Cost:\s*US\$\s*([\d,]*\.?\d+)/i) || [])[1];
    const status = ((text.match(/Status:\s*(.*?)(?:\s+(?:Your funds|Any questions)\b|$)/i) || [])[1] || '').trim();
    if (!side || !ticker || !qty || !price || !cost) {
      return { kind: 'unparsed', reason: 'order email without side, ticker, quantity, price or cost' };
    }
    if (!/completed/i.test(status)) return { kind: 'unparsed', reason: 'order status "' + status + '"' };
    const q = invNumber(qty[1]), p = invNumber(price), amount = invNumber(cost);
    if (!(q > 0) || !(p > 0) || Math.abs(q * p - amount) > Math.max(0.05, amount * 0.005)) {
      return { kind: 'unparsed', reason: 'quantity × average price does not match the cost' };
    }
    return { kind: 'event', event: Object.assign(base, {
      date: message.getDate(), type: side.charAt(0).toUpperCase() + side.slice(1).toLowerCase(), ticker: ticker,
      qty: q, price: p, amount: amount, fee: qty[2] ? invNumber(qty[2]) : 0
    }) };
  }
  if (/Dividend/i.test(subject)) {
    const ticker = (subject.match(/Dividend from ([A-Z0-9.\-]{1,12})\b/i) ||
      text.match(/dividends? from (?:your shares in )?([A-Z0-9.\-]{1,12})\b/) || [])[1];
    const amount = (text.match(/Net amount received:\s*(?:US)?\$\s*([\d,]*\.?\d+)/i) || [])[1];
    if (!ticker || !amount) return { kind: 'unparsed', reason: 'dividend email without ticker or amount' };
    const paid = text.match(/Payment date:\s*(\d{4})-(\d{2})-(\d{2})/);
    const date = paid ? new Date(Number(paid[1]), Number(paid[2]) - 1, Number(paid[3]), 12) : message.getDate();
    return { kind: 'event', event: Object.assign(base, {
      date: date, type: 'Dividend', ticker: ticker.toUpperCase(), qty: '', price: '', amount: invNumber(amount),
      notes: 'net amount received'
    }) };
  }
  if (/Deposit Completed/i.test(subject)) {
    return { kind: 'skipped', reason: 'deposit confirmation — it has no amount; deposits come from the bank transfer' };
  }
  return { kind: 'skipped', reason: 'not an order or a dividend' };
}

/** Broker emails in the run's range → { events, parsed, skipped, unparsed, threads, failedThreadIds }. */
function captureBrokerEmails(range) {
  const out = { events: [], parsed: 0, skipped: 0, unparsed: 0, threads: [], failedThreadIds: new Set() };
  Object.keys(BROKER_PATTERNS).forEach(broker => {
    const p = BROKER_PATTERNS[broker];
    const query = p.searchQuery + ' after:' + toEpochSeconds(range.start) + ' before:' + toEpochSeconds(range.endExclusive);
    gmailSearchAll(query, MAX_THREADS_PER_RUN).threads.forEach(thread => {
      thread.getMessages().forEach(message => {
        if (String(message.getFrom()).toLowerCase().indexOf(p.sender) === -1) return;
        const when = message.getDate();
        if (when < range.start || when >= range.endExclusive) return;
        let result;
        try { result = p.parse(message); } catch (error) { result = { kind: 'unparsed', reason: String(error) }; }
        if (result.kind === 'event') { out.events.push(result.event); out.parsed++; }
        else if (result.kind === 'skipped') { out.skipped++; }
        else {
          out.unparsed++;
          out.failedThreadIds.add(thread.getId());
          Logger.log('⚠️ ' + broker + ' email not read (' + result.reason + '): "' + message.getSubject() + '" ' + when);
        }
      });
      out.threads.push(thread);
    });
  });
  return out;
}

/** Investment Accounts sheet (created with the defaults) → [{account, kind, keyword}]. */
function readInvestmentAccounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INVESTMENT_ACCOUNTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(INVESTMENT_ACCOUNTS_SHEET);
    sheet.getRange(1, 1, 1, ACCOUNTS_HEADERS.length).setValues([ACCOUNTS_HEADERS]);
    sheet.getRange(2, 1, DEFAULT_INVESTMENT_ACCOUNTS.length, ACCOUNTS_HEADERS.length).setValues(DEFAULT_INVESTMENT_ACCOUNTS);
  }
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 3).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ account: String(r[0]).trim(), kind: String(r[1]).trim(), keyword: String(r[2]).trim() }));
}

/** Bank transfers (Transactions) whose beneficiary contains an account's deposit keyword → Deposit events. Pure. */
function brokerDepositsFromTransactions(values, accounts) {
  const withKeyword = accounts.filter(a => a.keyword);
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if ((r[TX_COL.TYPE] || '') !== 'Transfer' || !r[TX_COL.DATE]) continue;
    const merchant = String(r[TX_COL.MERCHANT] || '').toUpperCase();
    const account = withKeyword.find(a => merchant.indexOf(a.keyword.toUpperCase()) !== -1);
    if (!account) continue;
    const amount = Math.abs(invNumber(r[TX_COL.AMOUNT]));
    if (!amount) continue;
    const ref = r[TX_COL.MESSAGE_ID] || r[TX_COL.TX_REF] || (normalizeDateForCompare(r[TX_COL.DATE]) + ':' + amount);
    out.push({ date: r[TX_COL.DATE], account: account.account, type: 'Deposit', ticker: '', qty: '', price: '',
      amount: amount, fee: '', currency: r[TX_COL.CURRENCY] || 'DOP', source: 'bank',
      notes: (r[TX_COL.BANK] || '') + ' → ' + r[TX_COL.MERCHANT], id: 'bank:' + ref });
  }
  return out;
}

function ledgerRow(e) {
  return [e.date, e.account, e.type, e.ticker || '', e.qty === '' || e.qty === undefined ? '' : e.qty,
    e.price === '' || e.price === undefined ? '' : e.price, e.amount, e.fee || '', e.currency || 'USD',
    e.source || 'manual', e.notes || '', e.id || ''];
}

function getOrCreateLedgerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INVESTMENT_LEDGER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(INVESTMENT_LEDGER_SHEET);
    sheet.getRange(1, 1, 1, LEDGER_HEADERS.length).setValues([LEDGER_HEADERS]);
    sheet.hideColumns(LG.ID + 1);          // internal: email / bank reference, used to skip duplicates
    ensureAutoFilter(sheet, LEDGER_HEADERS.length, 1000);
  }
  return sheet;
}

/** Appends the events not saved yet (by Id), newest first. Rows you typed (no Id) are never touched. */
function saveInvestmentEvents(events) {
  const sheet = getOrCreateLedgerSheet();
  const last = sheet.getLastRow();
  const seen = new Set(last > 1 ? sheet.getRange(2, LG.ID + 1, last - 1, 1).getValues().map(r => String(r[0])).filter(Boolean) : []);
  const fresh = [];
  let duplicates = 0;
  events.forEach(e => {
    if (e.id && seen.has(e.id)) { duplicates++; return; }
    if (e.id) seen.add(e.id);
    fresh.push(ledgerRow(e));
  });
  if (fresh.length) {
    ensureRowCapacity(sheet, last + fresh.length);
    sheet.getRange(last + 1, 1, fresh.length, LEDGER_HEADERS.length).setValues(fresh);
    sortSheetByDateDesc(sheet, LG.DATE + 1);
  }
  return { saved: fresh.length, duplicates: duplicates };
}

/**
 * Ledger → positions, cash, other accounts and the year's figures. Pure — see tests/.
 * Average-cost method: a sale removes cost at the position's average cost and the
 * difference is realized P/L. DOP deposits after a snapshot are converted at opts.usdRate
 * to estimate the broker's cash (the broker converts at its own rate).
 */
function computeHoldings(values, opts) {
  opts = opts || {};
  const year = String(opts.year || new Date().getFullYear());
  const usdRate = Number(opts.usdRate) || 0;
  const rows = values.slice(1).filter(r => r[LG.DATE] && String(r[LG.ACCOUNT]).trim() && LEDGER_TYPES.indexOf(r[LG.TYPE]) !== -1)
    .map(r => ({ r: r, key: normalizeDateForCompare(r[LG.DATE]) }))
    .sort((a, b) => a.key < b.key ? -1 : a.key > b.key ? 1 : 0);
  const snapshotDay = {};
  rows.filter(x => x.r[LG.TYPE] === 'Snapshot').forEach(x => {
    const a = String(x.r[LG.ACCOUNT]).trim();
    if (!snapshotDay[a] || x.key > snapshotDay[a]) snapshotDay[a] = x.key;
  });
  const positions = {}, cash = {}, cashEstimated = {}, valuations = {};
  const totals = { dividendsYtd: 0, feesYtd: 0, realizedYtd: 0, contributionsYtd: { USD: 0, DOP: 0 } };
  const warnings = [];
  const pos = (a, t) => positions[a + '|' + t] ||
    (positions[a + '|' + t] = { account: a, ticker: t, qty: 0, cost: 0, realized: 0, dividends: 0, lastPrice: 0, lastPriceDay: '' });

  rows.forEach(({ r, key }) => {
    const a = String(r[LG.ACCOUNT]).trim(), type = r[LG.TYPE], t = String(r[LG.TICKER] || '').toUpperCase().trim();
    const q = invNumber(r[LG.QTY]), p = invNumber(r[LG.PRICE]), amount = invNumber(r[LG.AMOUNT]), fee = invNumber(r[LG.FEE]);
    const cur = String(r[LG.CURRENCY] || 'USD').toUpperCase();
    const inYear = key.slice(0, 4) === year;
    if (type === 'Valuation') {
      const value = amount || q * p;
      if (!valuations[a] || key >= valuations[a].day) valuations[a] = { account: a, value: value, currency: cur, day: key, units: q, unitPrice: p };
      return;
    }
    // the year's figures count every movement, before or after a snapshot
    if (inYear) {
      if (type === 'Dividend') totals.dividendsYtd += amount;
      if (type === 'Fee') totals.feesYtd += amount;
      if (fee && (type === 'Buy' || type === 'Sell')) totals.feesYtd += fee;
      if (type === 'Deposit') totals.contributionsYtd[cur === 'DOP' ? 'DOP' : 'USD'] += amount;
    }
    if (t && t !== 'CASH' && p > 0 && (type === 'Buy' || type === 'Sell' || type === 'Snapshot')) {
      const x = pos(a, t);
      if (key >= x.lastPriceDay) { x.lastPrice = p; x.lastPriceDay = key; }
    }
    if (type === 'Dividend' && t) pos(a, t).dividends += amount;

    const snap = snapshotDay[a];
    if (type === 'Snapshot') {
      if (key !== snap) return;                  // only the account's latest snapshot is the base
      if (t === 'CASH') cash[a] = (cash[a] || 0) + amount;
      else { const x = pos(a, t); x.qty += q; x.cost += amount; }
      return;
    }
    if (snap && key <= snap) return;             // already inside the snapshot
    const toUsd = value => {
      if (cur !== 'DOP') return value;
      cashEstimated[a] = true;
      return usdRate > 0 ? value / usdRate : 0;
    };
    if (type === 'Buy') {
      const x = pos(a, t); x.qty += q; x.cost += amount;
      cash[a] = (cash[a] || 0) - amount - fee;
    } else if (type === 'Sell') {
      const x = pos(a, t);
      if (x.qty <= 1e-9) {
        warnings.push(a + ' ' + t + ': sale on ' + key + ' with no position — add a Snapshot of this account');
        return;
      }
      const sold = Math.min(q, x.qty), average = x.cost / x.qty;
      if (q > x.qty + 1e-6) warnings.push(a + ' ' + t + ': sold ' + q + ' but held ' + +x.qty.toFixed(6) + ' on ' + key);
      x.qty -= sold; x.cost -= average * sold;
      const gain = amount * (sold / q) - average * sold;
      x.realized += gain;
      if (inYear) totals.realizedYtd += gain;
      cash[a] = (cash[a] || 0) + amount - fee;
    } else if (type === 'Dividend') {
      cash[a] = (cash[a] || 0) + amount;
    } else if (type === 'Deposit') {
      cash[a] = (cash[a] || 0) + toUsd(amount);
    } else if (type === 'Withdrawal') {
      cash[a] = (cash[a] || 0) - toUsd(amount);
    } else if (type === 'Fee') {
      cash[a] = (cash[a] || 0) - amount;
    }
  });

  const open = Object.keys(positions).map(k => positions[k]).filter(x => x.qty > 1e-9)
    .map(x => Object.assign(x, { qty: +x.qty.toFixed(8), cost: +x.cost.toFixed(2), average: x.cost / x.qty }))
    .sort((a, b) => b.cost - a.cost);
  Object.keys(cashEstimated).forEach(a => {
    if (usdRate <= 0) warnings.push(a + ': cash leaves out DOP deposits — set the USD rate on the Dashboard');
  });
  return {
    positions: open,
    closed: Object.keys(positions).map(k => positions[k]).filter(x => x.qty <= 1e-9 && (x.realized || x.dividends)),
    cash: Object.keys(cash).map(a => ({ account: a, amount: +cash[a].toFixed(2), estimated: !!cashEstimated[a] })),
    valuations: Object.keys(valuations).map(a => valuations[a]),
    totals: totals, warnings: warnings, snapshotDay: snapshotDay
  };
}

/**
 * A number as formula text: 10 significant digits, never scientific notation. v1.1.30: the last
 * known price used to be rounded to 6 decimals, which turned a SHIB-sized 0.0000061234 into 0.000006 (−2%).
 */
function formulaNumber(x) {
  const n = Number(Number(x).toPrecision(10));
  return /e/i.test(String(n)) ? n.toFixed(20).replace(/0+$/, '').replace(/\.$/, '') : String(n);
}

/**
 * Price and Price source cells of a position. Stocks/ETFs: GOOGLEFINANCE formulas (live in the sheet). Crypto: the
 * value fetched from Coinbase (opts.fetched). Both fall back to the last known price, and both are compared with it
 * when that price is recent (see PRICE_SANITY, SANITY_MAX_AGE_DAYS).
 */
function holdingPriceFormulas(ticker, lastPrice, opts) {
  opts = opts || {};
  const recent = !(opts.lastPriceDay && opts.today) || daysBetween(opts.lastPriceDay, opts.today) <= SANITY_MAX_AGE_DAYS;
  if (isCryptoPair(ticker)) {
    const fetched = Number(opts.fetched) || 0;
    if (fetched > 0 && !(lastPrice > 0 && recent && Math.abs(fetched / lastPrice - 1) > PRICE_SANITY)) {
      return { price: fetched, source: 'Coinbase' + (opts.fetchedAt ? ' · ' + opts.fetchedAt : '') };
    }
    if (fetched > 0) return { price: lastPrice, source: 'last known — check Coinbase price' };
    return { price: lastPrice > 0 ? lastPrice : 0, source: lastPrice > 0 ? 'last known (Coinbase unavailable)' : 'no price' };
  }
  const symbol = priceSymbol(ticker);
  const live = 'GOOGLEFINANCE("' + symbol + '","price")';
  if (!(lastPrice > 0)) {
    return { price: '=IFERROR(' + live + ',0)', source: '=IF(ISERROR(' + live + '),"no price","live")' };
  }
  const last = formulaNumber(lastPrice);
  if (!recent) {
    return { price: '=IFERROR(' + live + ',' + last + ')', source: '=IF(ISERROR(' + live + '),"last known","live")' };
  }
  const off = 'ABS(' + live + '/' + last + '-1)>' + PRICE_SANITY;
  return {
    price: '=IFERROR(IF(' + off + ',' + last + ',' + live + '),' + last + ')',
    source: '=IFERROR(IF(' + off + ',"last known — check symbol","live"),"last known")'
  };
}

/** Writes the Holdings sheet from computeHoldings(). Values come from formulas, so prices stay live. */
function buildHoldingsSheet(h, prices) {
  prices = prices || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HOLDINGS_SHEET);
  if (sheet) {
    sheet.clear();
    sheet.setConditionalFormatRules([]);
    sheet.getCharts().forEach(c => sheet.removeChart(c));   // sheet.clear() keeps embedded charts
  }
  else sheet = ss.insertSheet(HOLDINGS_SHEET);
  const T = SHEET_THEME, usd = '"US$"#,##0.00', pct = '0.00%';
  const nPos = Math.max(h.positions.length, 1), nCash = h.cash.length, nVal = Math.max(h.valuations.length, 1);
  const R = {};
  R.kpiLabel = 4; R.kpiValue = 5; R.posHead = 8; R.posCols = 9; R.posFirst = 10; R.posLast = R.posFirst + nPos - 1;
  R.cashFirst = R.posLast + 1; R.posTotal = R.cashFirst + nCash; R.valHead = R.posTotal + 3; R.valCols = R.valHead + 1;
  R.valFirst = R.valCols + 1; R.valLast = R.valFirst + nVal - 1; R.grand = R.valLast + 2; R.notes = R.grand + 3;
  ensureRowCapacity(sheet, R.notes + h.warnings.length + 6);

  sheet.getRange('B2').setValue('📈 Investments').setFontSize(18).setFontWeight('bold').setFontColor(T.headerBg);
  sheet.getRange('B3').setValue('Positions from the Investment Ledger · prices from GOOGLEFINANCE (delayed up to 20 min) · built ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')).setFontSize(9).setFontColor('#6B7280');
  if (h.positions.some(p => isCryptoPair(p.ticker))) {
    sheet.getRange('B3').setValue(sheet.getRange('B3').getValue() + ' · crypto from Coinbase when the tracker runs');
  }

  // positions
  sheet.getRange(R.posHead, 2).setValue('Positions').setFontWeight('bold').setFontColor(T.headerBg);
  const heads = ['Account', 'Ticker', 'Shares', 'Avg cost', 'Cost basis', 'Price', 'Market value', 'Unrealized P/L', 'P/L %', 'Weight', 'Dividends', 'Price source'];
  sheet.getRange(R.posCols, 2, 1, heads.length).setValues([heads]);
  const posRows = h.positions.length ? h.positions.map((p, i) => {
    const r = R.posFirst + i, f = holdingPriceFormulas(p.ticker, p.lastPrice, { lastPriceDay: p.lastPriceDay,
      today: prices.today, fetched: (prices.fetched || {})[p.ticker], fetchedAt: prices.fetchedAt });
    return [p.account, p.ticker, p.qty, '=F' + r + '/D' + r, p.cost, f.price, '=D' + r + '*G' + r, '=H' + r + '-F' + r,
      '=IF(F' + r + '>0,I' + r + '/F' + r + ',"")', '=IF($H$' + R.posTotal + '>0,H' + r + '/$H$' + R.posTotal + ',"")', p.dividends, f.source];
  }) : [['No positions yet — add a Snapshot to the Investment Ledger', '', '', '', '', '', '', '', '', '', '', '']];
  sheet.getRange(R.posFirst, 2, posRows.length, heads.length).setValues(posRows);
  if (nCash) {
    sheet.getRange(R.cashFirst, 2, nCash, heads.length).setValues(h.cash.map(c =>
      [c.account, 'Cash' + (c.estimated ? ' (≈ estimated)' : ''), '', '', c.amount, '', c.amount, '', '', '', '', c.estimated ? 'DOP deposits at the Dashboard rate' : 'ledger']));
  }
  sheet.getRange(R.posTotal, 2, 1, heads.length).setValues([['Total', '', '', '',
    '=SUM(F' + R.posFirst + ':F' + (R.posTotal - 1) + ')', '', '=SUM(H' + R.posFirst + ':H' + (R.posTotal - 1) + ')',
    '=SUM(I' + R.posFirst + ':I' + (R.posTotal - 1) + ')', '=IF(F' + R.posTotal + '>0,I' + R.posTotal + '/F' + R.posTotal + ',"")', '',
    '=SUM(L' + R.posFirst + ':L' + (R.posTotal - 1) + ')', '']]).setFontWeight('bold');

  // other accounts (tracked by value)
  sheet.getRange(R.valHead, 2).setValue('Other accounts').setFontWeight('bold').setFontColor(T.headerBg);
  sheet.getRange(R.valCols, 2, 1, 6).setValues([['Account', 'Value', 'Currency', 'As of', 'Units × unit price', 'Value (US$)']]);
  const valRows = h.valuations.length ? h.valuations.map((v, i) => {
    const r = R.valFirst + i;
    return [v.account, v.value, v.currency, v.day, v.units && v.unitPrice ? v.units + ' × ' + v.unitPrice : '',
      '=IF(D' + r + '="DOP",IFERROR(C' + r + '/RATE_USD,""),C' + r + ')'];
  }) : [['Add a Valuation row to the ledger for funds or pensions tracked by balance', '', '', '', '', '']];
  sheet.getRange(R.valFirst, 2, valRows.length, 6).setValues(valRows);

  sheet.getRange(R.grand, 2, 1, 4).setValues([['Total invested (US$)',
    '=H' + R.posTotal + '+SUM(G' + R.valFirst + ':G' + R.valLast + ')', 'DOP-equivalent',
    '=IFERROR(C' + R.grand + '*RATE_USD,"")']]).setFontWeight('bold').setFontSize(12);

  // KPI line
  const kpis = [['Portfolio value', '=C' + R.grand, usd], ['Unrealized P/L', '=I' + R.posTotal, usd],
    ['Dividends ' + new Date().getFullYear(), h.totals.dividendsYtd, usd], ['Fees ' + new Date().getFullYear(), h.totals.feesYtd, usd],
    ['Deposited ' + new Date().getFullYear() + ' (DOP)', h.totals.contributionsYtd.DOP, '"RD$"#,##0.00'],
    ['Deposited ' + new Date().getFullYear() + ' (US$)', h.totals.contributionsYtd.USD, usd]];
  kpis.forEach((k, i) => {
    const c = 2 + i * 2;
    sheet.getRange(R.kpiLabel, c).setValue(k[0]).setFontSize(9).setFontColor('#6B7280');
    sheet.getRange(R.kpiValue, c).setValue(k[1]).setNumberFormat(k[2]).setFontSize(14).setFontWeight('bold').setFontColor(T.headerBg);
  });

  // notes and warnings
  const notes = ['How to use: add a Snapshot of each account (one row per ticker with Quantity, Price and Amount = cost basis, ' +
    'plus a CASH row) — positions start from it and add the movements after that day. Funds and pensions: a Valuation row.']
    .concat(h.warnings.map(w => '⚠️ ' + w));
  sheet.getRange(R.notes, 2, notes.length, 1).setValues(notes.map(n => [n])).setFontSize(9).setFontColor('#6B7280');
  R.notesEnd = R.notes + notes.length - 1;

  // looks
  [R.posCols, R.valCols].forEach(r => sheet.getRange(r, 2, 1, r === R.posCols ? heads.length : 6)
    .setBackground(T.headerBg).setFontColor(T.headerFg).setFontWeight('bold'));
  const nRows = R.posTotal - R.posFirst + 1;
  sheet.getRange(R.posFirst, 4, nRows, 1).setNumberFormat('0.00000');
  sheet.getRange(R.posFirst, 5, nRows, 5).setNumberFormat(usd);
  sheet.getRange(R.posFirst, 5, nRows, 1).setNumberFormat(USD_PRICE_FORMAT);   // Avg cost
  sheet.getRange(R.posFirst, 7, nRows, 1).setNumberFormat(USD_PRICE_FORMAT);   // Price
  sheet.getRange(R.posFirst, 10, nRows, 2).setNumberFormat(pct);
  sheet.getRange(R.posFirst, 12, nRows, 1).setNumberFormat(usd);
  sheet.getRange(R.valFirst, 3, nVal, 1).setNumberFormat('#,##0.00');
  sheet.getRange(R.valFirst, 7, nVal, 1).setNumberFormat(usd);
  sheet.getRange(R.grand, 3).setNumberFormat(usd);
  sheet.getRange(R.grand, 5).setNumberFormat('"RD$"#,##0.00');
  const pl = sheet.getRange(R.posFirst, 9, nRows, 2);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(T.refund).setRanges([pl]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#B91C1C').setRanges([pl]).build()
  ]);
  sheet.setColumnWidth(1, 24);
  sheet.setColumnWidth(2, 150);
  for (let c = 3; c <= 12; c++) sheet.setColumnWidth(c, 112);
  sheet.setColumnWidth(13, 190);
  sheet.setFrozenRows(0);
  sheet.setHiddenGridlines(true);
  return { sheet: sheet, rows: R };
}

function styleLedgerSheet(sheet) {
  styleHeader(sheet, LEDGER_HEADERS.length);
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, LG.DATE + 1, rows, 1).setNumberFormat('yyyy-MM-dd');
  sheet.getRange(2, LG.QTY + 1, rows, 1).setNumberFormat('0.00000###');
  sheet.getRange(2, LG.PRICE + 1, rows, 1).setNumberFormat('#,##0.00######');   // up to 8 decimals (SHIB-sized prices)
  sheet.getRange(2, LG.AMOUNT + 1, rows, 2).setNumberFormat('#,##0.00');
  sheet.getRange(2, LG.TYPE + 1, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(LEDGER_TYPES, true).setAllowInvalid(false).build());
  const colours = { Buy: ['#E8F5E9', '#2E7D32'], Sell: ['#FDECEC', '#B91C1C'], Dividend: ['#F3E8FF', '#7E22CE'],
    Deposit: ['#EAF1FE', '#1D4ED8'], Withdrawal: ['#FFF4E5', '#B45309'], Fee: ['#F3F4F6', '#4B5563'],
    Snapshot: ['#E0F2F1', '#00695C'], Valuation: ['#FEF9C3', '#854D0E'] };
  const typeCol = sheet.getRange(2, LG.TYPE + 1, rows, 1);
  const rules = Object.keys(colours).map(t => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t)
    .setBackground(colours[t][0]).setFontColor(colours[t][1]).setRanges([typeCol]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",ISEVEN(ROW()))')
    .setBackground(SHEET_THEME.stripe).setRanges([sheet.getRange(2, 1, rows, LEDGER_HEADERS.length)]).build());
  sheet.setConditionalFormatRules(rules);
}

/** The investments part of a run: broker emails → ledger, bank deposits → ledger, Holdings rebuilt. */
function runInvestmentsStep(range) {
  const capture = range ? captureBrokerEmails(range) : { events: [], parsed: 0, skipped: 0, unparsed: 0, threads: [], failedThreadIds: new Set() };
  const accounts = readInvestmentAccounts();
  const tx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSACTIONS_SHEET);
  const deposits = tx ? brokerDepositsFromTransactions(tx.getDataRange().getValues(), accounts) : [];
  const saved = saveInvestmentEvents(capture.events.concat(deposits));
  if (capture.threads.length) markEmailsAsProcessed(capture.threads, capture.failedThreadIds);
  const holdings = refreshHoldings();
  return { parsed: capture.parsed, skipped: capture.skipped, unparsed: capture.unparsed, deposits: deposits.length,
    saved: saved.saved, duplicates: saved.duplicates, positions: holdings.positions.length, warnings: holdings.warnings };
}

/** Rebuilds Holdings from the ledger. */
function refreshHoldings() {
  const ledger = getOrCreateLedgerSheet();
  let usdRate = 0;
  try { usdRate = readDashboardRates(SpreadsheetApp.getActiveSpreadsheet()).USD; } catch (error) { usdRate = 0; }
  const h = computeHoldings(ledger.getDataRange().getValues(), { usdRate: usdRate, year: new Date().getFullYear() });
  const now = new Date(), todayKey = normalizeDateForCompare(now);
  const fetched = fetchCryptoPrices(h.positions.map(p => p.ticker));
  const built = buildHoldingsSheet(h, { today: todayKey, fetched: fetched,
    fetchedAt: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') });
  // v1.1.32: read what GOOGLEFINANCE computed, record today's values, then the Performance block
  let marketValues = null;
  if (h.positions.length) {
    SpreadsheetApp.flush();
    marketValues = built.sheet.getRange(built.rows.posFirst, 8, h.positions.length, 1).getValues().map(r => r[0]);
  }
  const valueRows = holdingsValueRows(h, marketValues, { fetched: fetched, usdRate: usdRate });
  recordPortfolioHistory(valueRows, todayKey);
  const current = {};
  valueRows.forEach(r => { current[r.account] = (current[r.account] || 0) + (Number(r.value) || 0); });
  h.returns = computeReturns(ledger.getDataRange().getValues(), current, { usdRate: usdRate, today: todayKey });
  writePerformanceBlock(built.sheet, built.rows.notesEnd + 3, h.returns, readHistory());
  return h;
}

/** Menu "📈 Refresh Investments": deposits from bank transfers + Holdings (emails are read by the regular runs). */
function refreshInvestmentsNow() {
  const config = requireConfig();
  if (!config) return;
  withRunLock(() => {
    try {
      safeToast('Refreshing investments...', '📈 Investments', -1);
      const r = runInvestmentsStep(null);
      formatDataSheets();
      ensureSheetOrder();
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOLDINGS_SHEET);
      if (sheet) SpreadsheetApp.setActiveSheet(sheet);
      safeToast('Done.', '📈 Investments', 3);   // v1.1.33: the progress toast used to stay open
      safeAlert('📈 Investments refreshed\n\n' + r.positions + ' position(s) · ' + r.saved + ' new ledger row(s) from bank deposits' +
        (r.warnings.length ? '\n\n⚠️ ' + r.warnings.join('\n⚠️ ') : ''));
    } catch (error) {
      safeAlert('❌ Could not refresh investments: ' + error);
    }
  });
}

/* ======================================================================
 * PORTFOLIO HISTORY, RETURNS AND REPORT DATA — v1.1.32
 * ====================================================================== */
const HISTORY_SHEET = 'Portfolio History';
const HISTORY_HEADERS = ['Date', 'Account', 'Ticker', 'Quantity', 'Price (US$)', 'Value (US$)'];
const HISTORY_TOTAL = 'TOTAL';
const ANNUALIZE_MIN_DAYS = 180;   // annualizing a few weeks of returns gives absurd numbers

/**
 * Every position, cash balance and balance-tracked account in US$ — the Market value
 * GOOGLEFINANCE already computed in Holdings (marketValues, one per position row), or
 * quantity × (Coinbase price or last known price) when that cell isn't a number yet. Pure.
 */
function holdingsValueRows(h, marketValues, opts) {
  opts = opts || {};
  const fetched = opts.fetched || {}, usdRate = Number(opts.usdRate) || 0;
  const rows = h.positions.map((p, i) => {
    const mv = marketValues ? Number(marketValues[i]) : NaN;
    const price = isFinite(mv) && mv > 0 ? mv / p.qty : (fetched[p.ticker] > 0 ? fetched[p.ticker] : p.lastPrice);
    return { account: p.account, ticker: p.ticker, qty: p.qty, price: price, value: isFinite(mv) && mv > 0 ? mv : p.qty * price };
  });
  h.cash.forEach(c => rows.push({ account: c.account, ticker: 'CASH', qty: '', price: '', value: c.amount }));
  h.valuations.forEach(v => rows.push({ account: v.account, ticker: '', qty: v.units || '', price: '',
    value: v.currency === 'DOP' ? (usdRate > 0 ? v.value / usdRate : 0) : v.value }));
  return rows;
}

/** Writes today's rows (replacing any earlier run of the same day) and a TOTAL row. */
function recordPortfolioHistory(valueRows, todayKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET);
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
  }
  const last = sheet.getLastRow();
  const kept = last > 1 ? sheet.getRange(2, 1, last - 1, HISTORY_HEADERS.length).getValues()
    .filter(r => r[0] && normalizeDateForCompare(r[0]) !== todayKey) : [];
  const parts = todayKey.split('-').map(Number);
  const day = new Date(parts[0], parts[1] - 1, parts[2], 12);
  const total = valueRows.reduce((s, r) => s + (Number(r.value) || 0), 0);
  const today = valueRows.map(r => [day, r.account, r.ticker, r.qty, r.price === '' ? '' : +Number(r.price).toPrecision(10), +Number(r.value).toFixed(2)])
    .concat([[day, HISTORY_TOTAL, '', '', '', +total.toFixed(2)]]);
  const all = today.concat(kept);
  if (last > 1) sheet.getRange(2, 1, last - 1, HISTORY_HEADERS.length).clearContent();
  ensureRowCapacity(sheet, all.length + 1);
  sheet.getRange(2, 1, all.length, HISTORY_HEADERS.length).setValues(all);
  sortSheetByDateDesc(sheet, 1);
  return { rows: today.length, total: total };
}

function readHistory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HISTORY_SHEET);
  return sheet ? sheet.getDataRange().getValues() : [HISTORY_HEADERS];
}

/**
 * Newton's method on Σ cf_i / (1+r)^(t_i/365) = 0. Flows: [{ day: 'yyyy-MM-dd', amount }]
 * (money in = negative). Returns the annual rate, or null when it doesn't converge.
 */
function xirr(flows) {
  if (flows.length < 2 || !flows.some(f => f.amount < 0) || !flows.some(f => f.amount > 0)) return null;
  const t0 = flows[0].day;
  const years = flows.map(f => daysBetween(t0, f.day) / 365);
  let r = 0.1;
  for (let k = 0; k < 100; k++) {
    let f = 0, df = 0;
    flows.forEach((c, i) => { const d = Math.pow(1 + r, years[i]); f += c.amount / d; df -= years[i] * c.amount / (d * (1 + r)); });
    if (Math.abs(df) < 1e-12) return null;
    const next = r - f / df;
    if (!isFinite(next) || next <= -0.9999) return null;
    if (Math.abs(next - r) < 1e-9) return next;
    r = next;
  }
  return null;
}

/**
 * Return since each account started being tracked — its latest Snapshot, first Valuation,
 * or first movement — and for all accounts together. Period return: Modified Dietz
 * (deposits weighted by how long they were invested). Annualized (XIRR) only after
 * ANNUALIZE_MIN_DAYS. DOP flows and values are converted at opts.usdRate. Pure.
 */
function computeReturns(ledgerValues, currentValues, opts) {
  opts = opts || {};
  const usdRate = Number(opts.usdRate) || 0, today = opts.today;
  const toUsd = (amount, cur) => String(cur || 'USD').toUpperCase() === 'DOP' ? (usdRate > 0 ? amount / usdRate : 0) : amount;
  const rows = ledgerValues.slice(1).filter(r => r[LG.DATE] && String(r[LG.ACCOUNT]).trim())
    .map(r => ({ r: r, key: normalizeDateForCompare(r[LG.DATE]), account: String(r[LG.ACCOUNT]).trim() }))
    .sort((a, b) => a.key < b.key ? -1 : a.key > b.key ? 1 : 0);
  const accounts = {};
  Object.keys(currentValues).forEach(a => { accounts[a] = { account: a, start: null, startValue: 0, flows: [] }; });
  const snapDay = {};
  rows.filter(x => x.r[LG.TYPE] === 'Snapshot').forEach(x => { if (!snapDay[x.account] || x.key > snapDay[x.account]) snapDay[x.account] = x.key; });
  Object.keys(accounts).forEach(a => {
    const acc = accounts[a];
    const mine = rows.filter(x => x.account === a);
    if (snapDay[a]) {
      acc.start = snapDay[a];
      mine.filter(x => x.r[LG.TYPE] === 'Snapshot' && x.key === snapDay[a]).forEach(x => {
        const t = String(x.r[LG.TICKER]).toUpperCase().trim();
        acc.startValue += t === 'CASH' ? invNumber(x.r[LG.AMOUNT]) : invNumber(x.r[LG.QTY]) * invNumber(x.r[LG.PRICE]);
      });
    } else {
      const firstVal = mine.find(x => x.r[LG.TYPE] === 'Valuation');
      if (firstVal && !mine.some(x => x.r[LG.TYPE] === 'Deposit' && x.key < firstVal.key)) {
        acc.start = firstVal.key;
        const v = invNumber(firstVal.r[LG.AMOUNT]) || invNumber(firstVal.r[LG.QTY]) * invNumber(firstVal.r[LG.PRICE]);
        acc.startValue = toUsd(v, firstVal.r[LG.CURRENCY]);
      } else if (mine.length) {
        acc.start = mine[0].key;                         // starts from zero; its first deposit is a flow
        acc.startValue = 0;
      }
    }
    if (!acc.start) return;
    mine.forEach(x => {
      const type = x.r[LG.TYPE];
      if (type !== 'Deposit' && type !== 'Withdrawal') return;
      if (acc.startValue > 0 ? x.key <= acc.start : x.key < acc.start) return;
      const amount = toUsd(invNumber(x.r[LG.AMOUNT]), x.r[LG.CURRENCY]);
      acc.flows.push({ day: x.key, amount: type === 'Deposit' ? amount : -amount });
    });
  });
  const measure = (start, startValue, flows, value) => {
    if (!start || !today) return null;
    const T = Math.max(daysBetween(start, today), 0);
    const net = flows.reduce((s, f) => s + f.amount, 0);
    const weighted = flows.reduce((s, f) => s + f.amount * (T > 0 ? (T - daysBetween(start, f.day)) / T : 1), 0);
    const gain = value - startValue - net;
    const base = startValue + weighted;
    const annual = T >= ANNUALIZE_MIN_DAYS
      ? xirr([{ day: start, amount: -startValue }].concat(flows.map(f => ({ day: f.day, amount: -f.amount })))
        .concat([{ day: today, amount: value }]).filter(f => f.amount !== 0)) : null;
    return { start: start, days: T, startValue: startValue, netDeposits: net, value: value, gain: gain,
      periodReturn: base > 0 ? gain / base : null, annualized: annual };
  };
  const list = Object.keys(accounts).map(a => {
    const acc = accounts[a];
    return Object.assign({ account: a }, measure(acc.start, acc.startValue, acc.flows, currentValues[a]) || {});
  }).filter(x => x.start);
  // all accounts together: accounts that start later enter as a flow on their start day
  const started = list.map(x => x.start).sort();
  const start = started[0] || null;
  const flows = [];
  let startValue = 0, value = 0;
  list.forEach(x => {
    value += x.value;
    if (x.start === start) startValue += x.startValue;
    else if (x.startValue) flows.push({ day: x.start, amount: x.startValue });
    accounts[x.account].flows.forEach(f => flows.push(f));
  });
  flows.sort((a, b) => a.day < b.day ? -1 : 1);
  return { accounts: list, total: start ? measure(start, startValue, flows, value) : null };
}

/**
 * What the daily email shows: today's total vs the previous recorded day, the gain in
 * between without deposits, and the positions whose price moved most. Pure.
 */
function investmentsDailyBrief(historyValues, ledgerValues, opts) {
  opts = opts || {};
  const usdRate = Number(opts.usdRate) || 0;
  const rows = historyValues.slice(1).filter(r => r[0]).map(r => ({ day: normalizeDateForCompare(r[0]), account: r[1],
    ticker: String(r[2] || ''), price: invNumber(r[4]), value: invNumber(r[5]) }));
  const days = rows.filter(r => r.account === HISTORY_TOTAL).map(r => r.day).sort();
  if (!days.length) return null;
  const today = days[days.length - 1], prev = days.length > 1 ? days[days.length - 2] : null;
  const totalOn = d => rows.filter(r => r.day === d && r.account === HISTORY_TOTAL).reduce((s, r) => s + r.value, 0);
  const brief = { day: today, total: totalOn(today), totalDop: usdRate > 0 ? totalOn(today) * usdRate : null, prevDay: prev,
    change: null, changePct: null, deposits: 0, movers: [], returns: opts.returns || null };
  if (prev) {
    ledgerValues.slice(1).forEach(r => {
      if (!r[LG.DATE] || (r[LG.TYPE] !== 'Deposit' && r[LG.TYPE] !== 'Withdrawal')) return;
      const k = normalizeDateForCompare(r[LG.DATE]);
      if (k <= prev || k > today) return;
      const cur = String(r[LG.CURRENCY] || 'USD').toUpperCase();
      const amount = cur === 'DOP' ? (usdRate > 0 ? invNumber(r[LG.AMOUNT]) / usdRate : 0) : invNumber(r[LG.AMOUNT]);
      brief.deposits += r[LG.TYPE] === 'Deposit' ? amount : -amount;
    });
    const before = totalOn(prev);
    brief.change = brief.total - before - brief.deposits;
    brief.changePct = before > 0 ? brief.change / before : null;
    const prevPrice = {};
    rows.filter(r => r.day === prev && r.price > 0).forEach(r => { prevPrice[r.account + '|' + r.ticker] = r.price; });
    brief.movers = rows.filter(r => r.day === today && r.price > 0 && prevPrice[r.account + '|' + r.ticker] > 0)
      .map(r => ({ ticker: r.ticker, account: r.account, change: r.price / prevPrice[r.account + '|' + r.ticker] - 1 }))
      .filter(m => Math.abs(m.change) >= 0.0005)
      .sort((a, b) => Math.abs(b.change) - Math.abs(a.change)).slice(0, 3);
  }
  return brief;
}

/** What the monthly email shows for the month of opts.month: value change, deposits, gain, dividends, fees, allocation. Pure. */
function investmentsMonthlyBrief(historyValues, ledgerValues, opts) {
  opts = opts || {};
  const usdRate = Number(opts.usdRate) || 0;
  const monthKey = normalizeDateForCompare(opts.month).slice(0, 7);
  const rows = historyValues.slice(1).filter(r => r[0]).map(r => ({ day: normalizeDateForCompare(r[0]), account: r[1], value: invNumber(r[5]) }));
  const totals = rows.filter(r => r.account === HISTORY_TOTAL).sort((a, b) => a.day < b.day ? -1 : 1);
  const inMonth = totals.filter(r => r.day.slice(0, 7) === monthKey);
  if (!inMonth.length) return null;
  const end = inMonth[inMonth.length - 1];
  const before = totals.filter(r => r.day.slice(0, 7) < monthKey);
  const startRow = before.length ? before[before.length - 1] : inMonth[0];
  const brief = { endDay: end.day, startDay: startRow.day, startValue: startRow.value, endValue: end.value,
    deposits: 0, dividends: 0, fees: 0, allocation: [], returns: opts.returns || null, partial: !before.length };
  ledgerValues.slice(1).forEach(r => {
    if (!r[LG.DATE]) return;
    const k = normalizeDateForCompare(r[LG.DATE]);
    const cur = String(r[LG.CURRENCY] || 'USD').toUpperCase();
    const usd = v => cur === 'DOP' ? (usdRate > 0 ? v / usdRate : 0) : v;
    if (k.slice(0, 7) === monthKey) {
      if (r[LG.TYPE] === 'Dividend') brief.dividends += usd(invNumber(r[LG.AMOUNT]));
      if (r[LG.TYPE] === 'Fee') brief.fees += usd(invNumber(r[LG.AMOUNT]));
      if (r[LG.TYPE] === 'Buy' || r[LG.TYPE] === 'Sell') brief.fees += usd(invNumber(r[LG.FEE]));
    }
    if ((r[LG.TYPE] === 'Deposit' || r[LG.TYPE] === 'Withdrawal') && k > startRow.day && k <= end.day) {
      brief.deposits += (r[LG.TYPE] === 'Deposit' ? 1 : -1) * usd(invNumber(r[LG.AMOUNT]));
    }
  });
  brief.gain = brief.endValue - brief.startValue - brief.deposits;
  brief.gainPct = brief.startValue > 0 ? brief.gain / brief.startValue : null;
  const byAccount = {};
  rows.filter(r => r.day === end.day && r.account !== HISTORY_TOTAL).forEach(r => { byAccount[r.account] = (byAccount[r.account] || 0) + r.value; });
  brief.allocation = Object.keys(byAccount).map(a => ({ account: a, value: byAccount[a], share: end.value > 0 ? byAccount[a] / end.value : 0 }))
    .sort((a, b) => b.value - a.value);
  return brief;
}

/** Holdings: a Performance block and a chart of the total over time (below everything else). */
function writePerformanceBlock(sheet, startRow, returns, historyValues) {
  const T = SHEET_THEME, usd = '"US$"#,##0.00';
  sheet.getRange(startRow, 2).setValue('Performance').setFontWeight('bold').setFontColor(T.headerBg);
  const heads = ['Account', 'Tracked since', 'Start value', 'Net deposits', 'Value now', 'Gain', 'Return', 'Annualized'];
  sheet.getRange(startRow + 1, 2, 1, heads.length).setValues([heads]).setBackground(T.headerBg).setFontColor(T.headerFg).setFontWeight('bold');
  const line = x => [x.account, x.start, x.startValue, x.netDeposits, x.value, x.gain,
    x.periodReturn === null ? '—' : x.periodReturn, x.annualized === null ? 'after ' + ANNUALIZE_MIN_DAYS + ' days' : x.annualized];
  const body = returns.accounts.map(line);
  if (returns.total) body.push(line(Object.assign({ account: 'All accounts' }, returns.total)));
  if (!body.length) body.push(['Performance appears after the first refresh', '', '', '', '', '', '', '']);
  sheet.getRange(startRow + 2, 2, body.length, heads.length).setValues(body);
  sheet.getRange(startRow + 2, 4, body.length, 4).setNumberFormat(usd);
  sheet.getRange(startRow + 2, 8, body.length, 2).setNumberFormat('+0.00%;-0.00%;0.00%');
  if (returns.total) sheet.getRange(startRow + 1 + body.length, 2, 1, heads.length).setFontWeight('bold');
  // total value over time → a small table off to the side, and a line chart of it
  const totals = historyValues.slice(1).filter(r => r[0] && r[1] === HISTORY_TOTAL)
    .map(r => [r[0], invNumber(r[5])]).sort((a, b) => normalizeDateForCompare(a[0]) < normalizeDateForCompare(b[0]) ? -1 : 1);
  const chartRow = startRow + body.length + 4;
  if (totals.length >= 2) {
    const col = 16;   // column P, outside the tables
    sheet.getRange(1, col, 1, 2).setValues([['Date', 'Total (US$)']]);
    ensureRowCapacity(sheet, totals.length + 1);
    sheet.getRange(2, col, totals.length, 2).setValues(totals);
    sheet.getRange(2, col, totals.length, 1).setNumberFormat('yyyy-MM-dd');
    sheet.hideColumns(col, 2);
    sheet.insertChart(sheet.newChart().setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(1, col, totals.length + 1, 2)).setNumHeaders(1)
      .setPosition(chartRow, 2, 0, 0).setOption('title', 'Portfolio value (US$)').setOption('legend', { position: 'none' })
      .setOption('colors', ['#0F766E']).setOption('width', 760).setOption('height', 260).build());
  } else {
    sheet.getRange(chartRow, 2).setValue('The chart appears once there are two days of history.').setFontSize(9).setFontColor('#6B7280');
  }
  return chartRow;
}

/** For the summary emails: the investments brief from Portfolio History and the ledger (null when there's none). */
function investmentsReportData(kind, opts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName(INVESTMENT_LEDGER_SHEET);
  const history = readHistory();
  if (!ledgerSheet || history.length < 2) return null;
  const ledger = ledgerSheet.getDataRange().getValues();
  const usdRate = opts.rates && opts.rates.USD || 0;
  const rows = history.slice(1).filter(r => r[0]);
  const latest = rows.map(r => normalizeDateForCompare(r[0])).sort().pop();
  const current = {};
  rows.filter(r => normalizeDateForCompare(r[0]) === latest && r[1] !== HISTORY_TOTAL)
    .forEach(r => { current[r[1]] = (current[r[1]] || 0) + invNumber(r[5]); });
  const returns = computeReturns(ledger, current, { usdRate: usdRate, today: latest }).total;
  return kind === 'daily'
    ? investmentsDailyBrief(history, ledger, { usdRate: usdRate, returns: returns })
    : investmentsMonthlyBrief(history, ledger, { usdRate: usdRate, returns: returns, month: opts.month });
}

/* ======================================================================
 * FUND / PENSION BALANCES — v1.1.33
 * A dialog that adds a Valuation row, so balances aren't typed by hand into the ledger.
 * ====================================================================== */
const VALUATION_KINDS = ['Fund', 'Pension', 'Other'];

/** null when the entry can be saved, otherwise what's wrong. Pure. */
function validateValuationEntry(e, todayKey) {
  if (!e || !String(e.account || '').trim()) return 'Give the account a name.';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(e.date || ''))) return 'Pick the statement date.';
  if (todayKey && e.date > todayKey) return 'The statement date is in the future.';
  if (['DOP', 'USD'].indexOf(e.currency) === -1) return 'Currency must be DOP or USD.';
  if (e.mode === 'units') {
    if (!(Number(e.units) > 0) || !(Number(e.unitPrice) > 0)) return 'Units and unit price must both be greater than 0.';
  } else if (!(Number(e.amount) > 0)) {
    return 'The balance must be greater than 0.';
  }
  return null;
}

/** Ledger row of a validated entry. Pure. */
function valuationRow(e) {
  const p = e.date.split('-').map(Number);
  const units = e.mode === 'units';
  return [new Date(p[0], p[1] - 1, p[2], 12), String(e.account).trim(), 'Valuation', '', units ? Number(e.units) : '',
    units ? Number(e.unitPrice) : '', units ? '' : Number(e.amount), '', e.currency, 'manual', String(e.notes || '').trim(), ''];
}

function openValuationDialog() {
  const config = getConfig();
  if (!config) {
    SpreadsheetApp.getUi().alert("❌ Setup not completed. Please run Setup Wizard first.");
    return;
  }
  const ledger = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVESTMENT_LEDGER_SHEET);
  const known = [];
  if (ledger && ledger.getLastRow() > 1) {
    ledger.getRange(2, 1, ledger.getLastRow() - 1, LEDGER_HEADERS.length).getValues().forEach(r => {
      const a = String(r[LG.ACCOUNT] || '').trim();
      if (r[LG.TYPE] === 'Valuation' && a && known.indexOf(a) === -1) known.push(a);
    });
  }
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html><html><head><base target="_top"><style>
  body { font-family: -apple-system, 'Segoe UI', Roboto, Arial, sans-serif; margin: 0; padding: 18px 20px; color: #1F2937; font-size: 14px; }
  label { display: block; font-weight: 600; font-size: 13px; margin: 12px 0 5px; }
  input, select { width: 100%; box-sizing: border-box; padding: 9px; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 14px; }
  .row { display: flex; gap: 10px; } .row > div { flex: 1; }
  .modes { display: flex; gap: 16px; margin-top: 4px; } .modes label { font-weight: 400; margin: 0; display: flex; gap: 6px; align-items: center; }
  .modes input { width: auto; }
  .hint { font-size: 12px; color: #6B7280; margin-top: 5px; line-height: 1.4; }
  button { margin-top: 18px; width: 100%; padding: 11px; border: 0; border-radius: 7px; background: #0F766E; color: #fff; font-size: 15px; font-weight: 600; cursor: pointer; }
  button:disabled { opacity: .6; }
  #status { display: none; margin-top: 12px; padding: 10px; border-radius: 6px; font-size: 13px; }
  #status.error { display: block; background: #FDECEC; color: #B91C1C; } #status.info { display: block; background: #EAF1FE; color: #1D4ED8; }
</style></head><body>
  <label for="account">Account</label>
  <input id="account" list="known" placeholder="e.g. Liquidity fund, Pension fund" value="">
  <datalist id="known">${known.map(a => `<option value="${esc(a)}">`).join('')}</datalist>
  <div class="hint">Use the same name every time — Holdings keeps the latest balance of each account.</div>
  <div class="row">
    <div><label for="kind">Kind</label><select id="kind">${VALUATION_KINDS.map(k => `<option>${k}</option>`).join('')}</select></div>
    <div><label for="date">Statement date</label><input id="date" type="date" value="${today}" max="${today}"></div>
  </div>
  <label>Balance as</label>
  <div class="modes">
    <label><input type="radio" name="mode" value="units" checked onchange="mode()"> Units × unit price</label>
    <label><input type="radio" name="mode" value="amount" onchange="mode()"> A balance</label>
  </div>
  <div id="unitsBox" class="row">
    <div><label for="units">Units</label><input id="units" type="number" step="any" min="0"></div>
    <div><label for="unitPrice">Unit price</label><input id="unitPrice" type="number" step="any" min="0"></div>
  </div>
  <div class="hint" id="unitsHint">A fund: units (cuotas) are in your statement; the unit price (valor cuota) in the fund's fact sheet.</div>
  <div id="amountBox" style="display:none"><label for="amount">Balance</label><input id="amount" type="number" step="any" min="0">
    <div class="hint">A pension: the balance of your latest statement.</div></div>
  <div class="row">
    <div><label for="currency">Currency</label><select id="currency"><option>DOP</option><option>USD</option></select></div>
    <div><label for="notes">Notes</label><input id="notes" placeholder="optional"></div>
  </div>
  <button id="save" onclick="save()">Save balance</button>
  <div id="status"></div>
<script>
  function el(id) { return document.getElementById(id); }
  function pick() { return document.querySelector('input[name="mode"]:checked').value; }
  function mode() {
    const u = pick() === 'units';
    el('unitsBox').style.display = u ? 'flex' : 'none'; el('unitsHint').style.display = u ? 'block' : 'none';
    el('amountBox').style.display = u ? 'none' : 'block';
  }
  function fail(m) { el('status').className = 'error'; el('status').textContent = m; }
  function save() {
    const e = { account: el('account').value.trim(), kind: el('kind').value, date: el('date').value, mode: pick(),
      units: el('units').value, unitPrice: el('unitPrice').value, amount: el('amount').value,
      currency: el('currency').value, notes: el('notes').value };
    if (!e.account) return fail('Give the account a name.');
    if (e.mode === 'units' && !(Number(e.units) > 0 && Number(e.unitPrice) > 0)) return fail('Units and unit price must both be greater than 0.');
    if (e.mode === 'amount' && !(Number(e.amount) > 0)) return fail('The balance must be greater than 0.');
    el('save').disabled = true;
    el('status').className = 'info';
    el('status').textContent = '⏳ Saving and updating Holdings — a summary pops up in the sheet. This window will close.';
    google.script.run.withFailureHandler(function(err) { fail('Error: ' + err); el('save').disabled = false; }).addValuationEntry(e);
    setTimeout(function() { google.script.host.close(); }, 1500);
  }
</script></body></html>`).setWidth(460).setHeight(610);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ Fund or pension balance');
}

/** Saves one balance as a Valuation row (a new row every statement — history is kept), then rebuilds Holdings. */
function addValuationEntry(entry) {
  const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const problem = validateValuationEntry(entry, todayKey);
  if (problem) {
    safeAlert('❌ Balance not saved: ' + problem);
    return false;
  }
  return withRunLock(() => {
    try {
      safeToast('Saving the balance...', '📈 Investments', -1);
      const sheet = getOrCreateLedgerSheet();
      const row = valuationRow(entry);
      const last = sheet.getLastRow();
      ensureRowCapacity(sheet, last + 1);
      sheet.getRange(last + 1, 1, 1, LEDGER_HEADERS.length).setValues([row]);
      sortSheetByDateDesc(sheet, LG.DATE + 1);
      // the account is listed in Investment Accounts (no deposit keyword needed for a balance)
      readInvestmentAccounts();
      const accounts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVESTMENT_ACCOUNTS_SHEET);
      const names = accounts.getLastRow() > 1 ? accounts.getRange(2, 1, accounts.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase()) : [];
      if (names.indexOf(row[LG.ACCOUNT].toLowerCase()) === -1) {
        accounts.getRange(accounts.getLastRow() + 1, 1, 1, ACCOUNTS_HEADERS.length)
          .setValues([[row[LG.ACCOUNT], VALUATION_KINDS.indexOf(entry.kind) !== -1 ? entry.kind : 'Other', '', 'Balance from statements (Valuation rows)']]);
      }
      safeToast('Updating Holdings...', '📈 Investments', -1);
      refreshHoldings();
      formatDataSheets();
      ensureSheetOrder();
      safeToast('Done.', '📈 Investments', 3);
      const value = entry.mode === 'units' ? Number(entry.units) * Number(entry.unitPrice) : Number(entry.amount);
      safeAlert('✅ Balance saved\n\n' + row[LG.ACCOUNT] + ' — ' + entry.currency + ' ' +
        value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' as of ' + entry.date +
        (entry.mode === 'units' ? ' (' + entry.units + ' × ' + entry.unitPrice + ')' : '') +
        '\n\nIt now shows in Holdings › Other accounts. Add a new balance with every statement.');
      return true;
    } catch (error) {
      safeAlert('❌ Could not save the balance: ' + error);
      return false;
    }
  });
}
