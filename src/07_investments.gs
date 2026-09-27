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
// GOOGLEFINANCE symbols for tickers it doesn't know as they are
const PRICE_SYMBOLS = { ETHUSD: 'CURRENCY:ETHUSD', BTCUSD: 'CURRENCY:BTCUSD' };
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

/** GOOGLEFINANCE price with a sanity check against the last known price (see PRICE_SANITY). */
function holdingPriceFormulas(ticker, lastPrice) {
  const symbol = PRICE_SYMBOLS[ticker] || ticker;
  const live = 'GOOGLEFINANCE("' + symbol + '","price")';
  if (!(lastPrice > 0)) {
    return { price: '=IFERROR(' + live + ',0)', source: '=IF(ISERROR(' + live + '),"no price","live")' };
  }
  const off = 'ABS(' + live + '/' + lastPrice + '-1)>' + PRICE_SANITY;
  return {
    price: '=IFERROR(IF(' + off + ',' + lastPrice + ',' + live + '),' + lastPrice + ')',
    source: '=IFERROR(IF(' + off + ',"last known — check symbol","live"),"last known")'
  };
}

/** Writes the Holdings sheet from computeHoldings(). Values come from formulas, so prices stay live. */
function buildHoldingsSheet(h) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HOLDINGS_SHEET);
  if (sheet) { sheet.clear(); sheet.setConditionalFormatRules([]); }
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

  // positions
  sheet.getRange(R.posHead, 2).setValue('Positions').setFontWeight('bold').setFontColor(T.headerBg);
  const heads = ['Account', 'Ticker', 'Shares', 'Avg cost', 'Cost basis', 'Price', 'Market value', 'Unrealized P/L', 'P/L %', 'Weight', 'Dividends', 'Price source'];
  sheet.getRange(R.posCols, 2, 1, heads.length).setValues([heads]);
  const posRows = h.positions.length ? h.positions.map((p, i) => {
    const r = R.posFirst + i, f = holdingPriceFormulas(p.ticker, +p.lastPrice.toFixed(6));
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

  // looks
  [R.posCols, R.valCols].forEach(r => sheet.getRange(r, 2, 1, r === R.posCols ? heads.length : 6)
    .setBackground(T.headerBg).setFontColor(T.headerFg).setFontWeight('bold'));
  const nRows = R.posTotal - R.posFirst + 1;
  sheet.getRange(R.posFirst, 4, nRows, 1).setNumberFormat('0.00000');
  sheet.getRange(R.posFirst, 5, nRows, 5).setNumberFormat(usd);
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
  sheet.getRange(2, LG.PRICE + 1, rows, 3).setNumberFormat('#,##0.00###');
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
  buildHoldingsSheet(h);
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
      safeAlert('📈 Investments refreshed\n\n' + r.positions + ' position(s) · ' + r.saved + ' new ledger row(s) from bank deposits' +
        (r.warnings.length ? '\n\n⚠️ ' + r.warnings.join('\n⚠️ ') : ''));
    } catch (error) {
      safeAlert('❌ Could not refresh investments: ' + error);
    }
  });
}
