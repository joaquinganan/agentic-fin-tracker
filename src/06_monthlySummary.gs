/**
 * MONTHLY SUMMARY EMAIL — v1.1.26
 *
 * Sent on the 1st of each month about the PREVIOUS calendar month — the same
 * period as the Dashboard's month columns and as a monthly salary. By the
 * time it runs (the hour chosen in the Setup Wizard), the 6 AM update of the
 * 1st has already swept the whole previous month (see computeDailyWindow()),
 * so the month's last day is included.
 *
 * Same rules as the Dashboard and the daily email: spending = Transaction and
 * Transfer rows with a category other than "Exclude", in DOP-equivalent at
 * the Dashboard's rates. Uses the email kit in 05_dailySummary.gs.
 *
 * Menu "🗓️ Send Monthly Summary Now" sends the last complete month on demand.
 */
const SUMMARY_SUBSCRIPTIONS_CATEGORY = 'Streaming & Subscriptions';

/** Everything the monthly email needs. Pure — see tests/. opts as computeDailySummary() plus optional `month` (a Date in it). */
function computeMonthlySummary(values, opts) {
  const rates = opts.rates;
  const ref = opts.today;
  const month = opts.month
    ? new Date(opts.month.getFullYear(), opts.month.getMonth(), 1, 12)
    : new Date(ref.getFullYear(), ref.getMonth() - 1, 1, 12);
  const monthKeyOf = d => normalizeDateForCompare(d).slice(0, 7);
  const monthKey = monthKeyOf(month);
  const prevMonth = new Date(month.getFullYear(), month.getMonth() - 1, 1, 12);
  const prevKey = monthKeyOf(prevMonth);
  const usualKeys = [1, 2, 3].map(k => monthKeyOf(new Date(month.getFullYear(), month.getMonth() - k, 1, 12)));
  const daysInMonth = new Date(month.getFullYear(), month.getMonth() + 1, 0).getDate();
  const toDop = (amount, cur) => (Number(amount) || 0) * ({ USD: rates.USD, EUR: rates.EUR, COP: rates.COP }[cur] || 1);
  const cardRate = {}, cardName = {};
  (opts.cards || []).forEach(c => { if (c.rate > 0) { cardRate[c.bank] = c.rate; cardName[c.bank] = c.name || ''; } });

  const m = {
    month: month, monthKey: monthKey, prevMonth: prevMonth, daysInMonth: daysInMonth,
    netIncomeDop: Number(opts.netIncomeDop) || 0,
    spent: 0, byCategory: {}, byDay: new Array(daysInMonth + 1).fill(0), byDayToDay: new Array(daysInMonth + 1).fill(0),
    byBank: {}, merchants: {}, subs: {},
    prevSubs: {}, prevByCategory: {}, usualCatSums: {}, monthTotals: {},
    fixed: 0, variable: 0, transfersOut: 0, transfersCount: 0, openCount: 0, openTotal: 0,
    placeholders: 0, unmatchedReversals: 0, purchases: 0, biggest: null,
    missed: { total: 0, byRule: {} }
  };

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[TX_COL.DATE]) continue;
    const key = normalizeDateForCompare(r[TX_COL.DATE]);
    const mk = key.slice(0, 7);
    const type = r[TX_COL.TYPE] || 'Transaction';
    const cat = String(r[TX_COL.CATEGORY] || '');
    const amt = toDop(r[TX_COL.AMOUNT], r[TX_COL.CURRENCY]);
    const merchant = String(r[TX_COL.MERCHANT] || '');
    const bank = String(r[TX_COL.BANK] || '');
    const spend = (type === 'Transaction' || type === 'Transfer') && cat !== '' && cat !== EXCLUDE_CATEGORY;
    if (spend) m.monthTotals[mk] = (m.monthTotals[mk] || 0) + amt;

    if (mk === monthKey) {
      if (spend) {
        m.spent += amt;
        m.byCategory[cat] = (m.byCategory[cat] || 0) + amt;
        m.byDay[Number(key.slice(8, 10))] += amt;
        if (!isBillCategory(cat)) m.byDayToDay[Number(key.slice(8, 10))] += amt;   // calendar colours: no rent/bill spikes
        m.byBank[bank] = (m.byBank[bank] || 0) + amt;
        const g = m.merchants[merchant] || (m.merchants[merchant] = { merchant: merchant, total: 0, count: 0 });
        g.total += amt; g.count++;
        if (FIXED_CATEGORY_NAMES.indexOf(cat) !== -1) m.fixed += amt; else m.variable += amt;
        if (amt > 0 && type === 'Transaction') {
          m.purchases++;
          if (!m.biggest || amt > m.biggest.amount) m.biggest = { merchant: merchant, amount: amt, bank: bank, date: key };
        }
        if (cat === SUMMARY_SUBSCRIPTIONS_CATEGORY) {
          const sg = m.subs[merchant] || (m.subs[merchant] = { merchant: merchant, total: 0, count: 0 });
          sg.total += amt; sg.count++;
        }
        // cashback left on the table: a purchase whose best card (per the terms) is one you have, paid with another
        if (type === 'Transaction' && amt > 0) {
          const rule = CARD_TIP_RULES.find(x => x.test.test(normalizeKeyword(merchant)));
          if (rule && cardRate[rule.bank] && bank !== rule.bank) {
            const k = rule.label + '|' + bank + '|' + rule.bank;
            const e = m.missed.byRule[k] || (m.missed.byRule[k] = { label: rule.label, paidWith: bank,
              best: rule.bank + (cardName[rule.bank] ? ' ' + cardName[rule.bank] : ''), spent: 0, cashback: 0 });
            e.spent += amt; e.cashback += amt * cardRate[rule.bank];
            m.missed.total += amt * cardRate[rule.bank];
          }
        }
      }
      if (type === 'Transfer' && cat !== EXCLUDE_CATEGORY) { m.transfersOut += amt; m.transfersCount++; }
      if (type === 'Transfer' && cat === '') { m.openCount++; m.openTotal += amt; }
      if (merchant === GARBLED_PLACEHOLDER) m.placeholders++;
      if (merchant === REVERSAL_UNMATCHED) m.unmatchedReversals++;
    }
    if (spend && mk === prevKey) {
      m.prevByCategory[cat] = (m.prevByCategory[cat] || 0) + amt;
      if (cat === SUMMARY_SUBSCRIPTIONS_CATEGORY) m.prevSubs[merchant] = true;
    }
    if (spend && usualKeys.indexOf(mk) !== -1) m.usualCatSums[cat] = (m.usualCatSums[cat] || 0) + amt;
  }

  // comparisons
  m.prevTotal = m.monthTotals[prevKey] || 0;
  const usualWithData = usualKeys.filter(k => (m.monthTotals[k] || 0) > 0);
  m.usualTotal = usualWithData.length ? usualWithData.reduce((a, k) => a + m.monthTotals[k], 0) / usualWithData.length : 0;
  m.usualMonths = usualWithData.length;
  m.vsPrev = m.prevTotal > 0 ? m.spent / m.prevTotal - 1 : null;
  m.vsUsual = m.usualTotal > 0 ? m.spent / m.usualTotal - 1 : null;
  m.leftOver = m.netIncomeDop > 0 ? m.netIncomeDop - m.spent : null;
  m.savingsRate = m.netIncomeDop > 0 ? m.leftOver / m.netIncomeDop : null;
  m.incomeUsed = m.netIncomeDop > 0 ? m.spent / m.netIncomeDop : null;

  m.categories = Object.keys(m.byCategory).map(c => ({
    cat: c, amount: m.byCategory[c], share: m.spent > 0 ? m.byCategory[c] / m.spent : 0,
    prev: m.prevByCategory[c] || 0, usual: m.usualMonths ? (m.usualCatSums[c] || 0) / m.usualMonths : 0,
    fixed: FIXED_CATEGORY_NAMES.indexOf(c) !== -1
  })).sort((a, b) => b.amount - a.amount);
  m.topMerchants = Object.keys(m.merchants).map(k => m.merchants[k]).sort((a, b) => b.total - a.total).slice(0, 8);
  m.subscriptions = Object.keys(m.subs).map(k => Object.assign({ recurring: !!m.prevSubs[k] }, m.subs[k]))
    .sort((a, b) => b.total - a.total);
  m.subscriptionsTotal = m.subscriptions.reduce((a, x) => a + x.total, 0);
  m.banks = Object.keys(m.byBank).map(b => ({ bank: b, amount: m.byBank[b] })).sort((a, b) => b.amount - a.amount);
  m.missedList = Object.keys(m.missed.byRule).map(k => m.missed.byRule[k]).sort((a, b) => b.cashback - a.cashback);

  // year to date (same year, up to this month)
  const year = String(month.getFullYear());
  const ytdKeys = Object.keys(m.monthTotals).filter(k => k.slice(0, 4) === year && k <= monthKey && m.monthTotals[k] > 0);
  m.ytdTotal = ytdKeys.reduce((a, k) => a + m.monthTotals[k], 0);
  m.ytdMonths = ytdKeys.length;
  m.ytdAverage = m.ytdMonths ? m.ytdTotal / m.ytdMonths : 0;
  m.ytdSavingsRate = m.netIncomeDop > 0 && m.ytdMonths ? 1 - m.ytdTotal / (m.netIncomeDop * m.ytdMonths) : null;

  m.recommendations = monthlyRecommendations(m);
  m.investments = opts.investments || null;   // v1.1.32
  m.lastRun = opts.lastRun || null;
  m.now = opts.now || ref;
  return m;
}

/** Fixed rules → [{tone, text}]. */
function monthlyRecommendations(m) {
  const out = [];
  const name = SUMMARY_MONTHS_FULL[m.month.getMonth()];
  const prevName = SUMMARY_MONTHS_FULL[m.prevMonth.getMonth()];
  if (m.savingsRate !== null) {
    if (m.savingsRate < 0) out.push({ tone: 'bad', text: 'You spent ' + summaryMoney(-m.leftOver) + ' more than your net income in ' + name + '.' });
    else if (m.savingsRate < 0.1) out.push({ tone: 'warn', text: 'Only ' + summaryPct(m.savingsRate) + ' of your net income was left over (' + summaryMoney(m.leftOver) + ').' });
    else out.push({ tone: 'good', text: 'You kept ' + summaryPct(m.savingsRate) + ' of your net income — ' + summaryMoney(m.leftOver) + '.' });
  }
  if (m.vsPrev !== null && m.vsPrev > 0.15) {
    const driver = m.categories.map(c => ({ cat: c.cat, diff: c.amount - c.prev })).sort((a, b) => b.diff - a.diff)[0];
    out.push({ tone: 'warn', text: 'Spending was up ' + summaryPct(m.vsPrev) + ' vs. ' + prevName +
      (driver && driver.diff > 0 ? ', mostly ' + driver.cat + ' (+' + summaryMoney(driver.diff) + ')' : '') + '.' });
  } else if (m.vsPrev !== null && m.vsPrev < -0.1) {
    out.push({ tone: 'good', text: 'Spending was down ' + summaryPct(m.vsPrev) + ' vs. ' + prevName + '.' });
  }
  m.categories.filter(c => !c.fixed && c.usual > 0 && c.amount > c.usual * 1.15 && c.amount - c.usual > 1000)
    .sort((a, b) => (b.amount - b.usual) - (a.amount - a.usual)).slice(0, 2)
    .forEach(c => out.push({ tone: 'warn', text: c.cat + ': ' + summaryMoney(c.amount) + ', ' +
      summaryPct(c.amount / c.usual - 1) + ' above your usual month (' + summaryMoney(c.usual) + ').' }));
  if (m.subscriptions.length) {
    out.push({ tone: 'info', text: m.subscriptions.length + ' subscription(s) cost ' + summaryMoney(m.subscriptionsTotal) +
      ' this month — about ' + summaryMoney(m.subscriptionsTotal * 12) + ' a year. Worth a quick check that you still use them all.' });
  }
  if (m.missed.total >= 50) {
    const top = m.missedList[0];
    out.push({ tone: 'info', text: 'About ' + summaryMoney(m.missed.total) + ' in cashback was left on the table — e.g. ' +
      top.label.toLowerCase() + ' paid with ' + top.paidWith + ' (' + summaryMoney(top.spent) + ') would earn with your ' + top.best + ' card.' });
  }
  if (m.openCount) {
    out.push({ tone: 'note', text: m.openCount + ' transfer(s) have no category (' + summaryMoney(m.openTotal) +
      ') and aren\'t counted as spending — add a Custom Rule, or "Exclude" for your own accounts.' });
  }
  if (m.placeholders || m.unmatchedReversals) {
    out.push({ tone: 'note', text: 'Data to check: ' + [m.placeholders ? m.placeholders + ' row(s) with an unreadable merchant' : '',
      m.unmatchedReversals ? m.unmatchedReversals + ' reversal(s) without their original purchase' : ''].filter(Boolean).join(' and ') + '.' });
  }
  return out;
}

/** Month calendar (Mon–Sun rows) coloured by the day's spending. */
function ekCalendar(month, byDay) {
  const days = byDay.length - 1;
  const max = Math.max.apply(null, byDay.concat([1]));
  const level = v => v <= 0 ? 0 : Math.min(5, 1 + Math.floor(v / max * 4.999));
  const first = (new Date(month.getFullYear(), month.getMonth(), 1, 12).getDay() + 6) % 7;   // Monday = 0
  let html = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:4px">' +
    '<tr>' + ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].map(d =>
      '<td width="14%" align="center" style="font-size:10px;color:' + EK.faint + '">' + d + '</td>').join('') + '</tr><tr>';
  let col = 0;
  for (let i = 0; i < first; i++, col++) html += '<td></td>';
  for (let d = 1; d <= days; d++, col++) {
    if (col && col % 7 === 0) html += '</tr><tr>';
    const lv = level(byDay[d]);
    html += '<td width="14%" align="center" style="height:34px;border-radius:6px;background:' + EK.heat[lv] + ';font-size:11px;color:' +
      (lv >= 4 ? '#FFFFFF' : EK.muted) + '">' + d + '</td>';
  }
  while (col % 7) { html += '<td></td>'; col++; }
  html += '</tr></table><div style="font-size:11px;color:' + EK.faint + ';padding:4px 4px 0">Less ' +
    EK.heat.map(c => '<span style="display:inline-block;width:12px;height:10px;border-radius:3px;background:' + c +
      ';vertical-align:middle"></span>').join('&nbsp;') + ' More</div>';
  return html;
}

/** Subject, HTML and plain-text bodies of the monthly email. Pure — see tests/. */
function buildMonthlySummaryEmail(m, opts) {
  opts = opts || {};
  const links = opts.links || { dashboard: opts.sheetUrl };
  const name = SUMMARY_MONTHS_FULL[m.month.getMonth()] + ' ' + m.month.getFullYear();
  const prevName = SUMMARY_MONTHS_FULL[m.prevMonth.getMonth()];
  const subject = name + ': ' + summaryMoney(m.spent) + ' spent' +
    (m.savingsRate !== null ? ' · ' + summaryPct(m.savingsRate) + (m.savingsRate < 0 ? ' over income' : ' left over')
      : (m.vsPrev !== null ? ' (' + summarySignedPct(m.vsPrev) + ' vs. ' + prevName + ')' : ''));
  const rows = [], T = [];

  // hero
  const chips = [];
  if (m.vsPrev !== null) chips.push(ekChip(summarySignedPct(m.vsPrev) + ' vs. ' + prevName, m.vsPrev > 0.15 ? 'warn' : m.vsPrev < -0.05 ? 'good' : 'info'));
  if (m.incomeUsed !== null) chips.push(ekChip(summaryPct(m.incomeUsed) + ' of net income', m.incomeUsed > 1 ? 'bad' : m.incomeUsed > 0.9 ? 'warn' : 'good'));
  rows.push('<tr><td style="padding:24px 28px 0"><div style="font-size:13px;color:' + EK.muted + '">Spent in ' +
    SUMMARY_MONTHS_FULL[m.month.getMonth()] + ' · ' + m.purchases + ' purchase(s)</div>' +
    '<div style="font-size:40px;font-weight:800;color:' + EK.navy + ';line-height:1.15;margin:4px 0 10px">' + summaryMoney(m.spent) +
    '</div>' + chips.join('&nbsp; ') + '</td></tr>');
  T.push(name + ': ' + summaryMoney(m.spent) + ' spent · ' + m.purchases + ' purchase(s)');

  // KPIs + income used
  const tiles = [
    { label: 'Left over', value: m.leftOver !== null ? summaryMoneyShort(m.leftOver) : '—',
      sub: m.savingsRate !== null ? summaryPct(m.savingsRate) + ' of net income' : 'set income in the wizard',
      tone: m.savingsRate === null ? null : m.savingsRate < 0 ? 'bad' : m.savingsRate < 0.1 ? 'warn' : 'good' },
    { label: 'vs. ' + prevName, value: m.vsPrev !== null ? summarySignedPct(m.vsPrev) : '—',
      sub: m.prevTotal ? summaryMoneyShort(m.prevTotal) + ' then' : 'no data', tone: m.vsPrev !== null && m.vsPrev > 0.15 ? 'warn' : null },
    { label: 'Usual month', value: m.usualTotal ? summaryMoneyShort(m.usualTotal) : '—',
      sub: m.usualMonths ? 'avg. of ' + m.usualMonths + ' month(s)' : 'no history yet' }
  ];
  let inner = ekKpis(tiles);
  if (m.incomeUsed !== null) {
    inner += '<div style="height:14px;font-size:0">&nbsp;</div>' + ekBar('Net income used', m.incomeUsed,
      m.incomeUsed > 1 ? ekTone('bad').fg : m.incomeUsed > 0.9 ? ekTone('warn').fg : ekTone('good').fg, summaryPct(m.incomeUsed),
      summaryEscape(summaryMoney(m.spent) + ' of ' + summaryMoney(m.netIncomeDop)));
  }
  const fixedShare = m.spent > 0 ? m.fixed / m.spent : 0;
  inner += '<div style="font-size:13px;color:' + EK.text + ';margin:4px 0 6px">Fixed vs. variable</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-radius:999px;overflow:hidden"><tr>' +
    (m.fixed > 0 ? '<td width="' + Math.max(1, Math.round(fixedShare * 100)) + '%" style="background:' + EK.navy + ';height:10px;font-size:0">&nbsp;</td>' : '') +
    (m.variable > 0 ? '<td style="background:#7FA6E8;height:10px;font-size:0">&nbsp;</td>' : '') + '</tr></table>' +
    '<div style="font-size:12px;color:' + EK.muted + ';padding-top:5px"><span style="color:' + EK.navy + '">■</span> Fixed ' +
    summaryEscape(summaryMoney(m.fixed)) + ' (' + summaryPct(fixedShare) + ') &nbsp; <span style="color:#7FA6E8">■</span> Variable ' +
    summaryEscape(summaryMoney(m.variable)) + '</div>';
  rows.push(ekSection('The month in numbers', inner));
  if (m.leftOver !== null) T.push('Left over: ' + summaryMoney(m.leftOver) + ' (' + summaryPct(m.savingsRate) + ' of net income)');
  if (m.vsPrev !== null) T.push('vs. ' + prevName + ': ' + summarySignedPct(m.vsPrev));

  rows.push(ekSection('Day by day · day-to-day spending', ekCalendar(m.month, m.byDayToDay)));

  // categories
  const maxCat = Math.max.apply(null, m.categories.map(c => c.amount).concat([1]));
  rows.push(ekSection('Categories', m.categories.map(c => {
    const parts = [summaryPct(c.share) + ' of the month'];
    if (c.prev > 0) {
      const d = c.amount / c.prev - 1;
      const col = d > 0.15 ? ekTone('warn').fg : d < -0.1 ? ekTone('good').fg : EK.muted;
      parts.push('<span style="color:' + col + '">' + summarySignedPct(d) + ' vs. ' + summaryEscape(prevName) + '</span>');
    } else if (!c.fixed) {
      parts.push('new vs. ' + summaryEscape(prevName));
    }
    return ekBar(c.cat, c.amount / maxCat, c.fixed ? EK.navy : EK.accent, summaryMoney(c.amount), parts.join(' · '));
  }).join('')));
  m.categories.forEach(c => T.push('  • ' + c.cat + ': ' + summaryMoney(c.amount)));

  if (m.topMerchants.length) {
    rows.push(ekSection('Top merchants', ekList(m.topMerchants.map(g => ({
      title: g.merchant, meta: g.count + (g.count === 1 ? ' payment' : ' payments'), right: summaryMoney(g.total) })))));
  }
  if (m.subscriptions.length) {
    rows.push(ekSection('Subscriptions', ekList(m.subscriptions.map(x => ({
      title: x.merchant, meta: x.recurring ? 'also last month' : 'new this month', right: summaryMoney(x.total) }))) +
      '<div style="font-size:12px;color:' + EK.muted + ';padding-top:6px">≈ ' + summaryEscape(summaryMoney(m.subscriptionsTotal * 12)) +
      ' a year at this rate.</div>'));
    T.push('', 'Subscriptions: ' + summaryMoney(m.subscriptionsTotal) + ' (' + m.subscriptions.map(x => x.merchant).join(', ') + ')');
  }

  // banks + cashback left on the table
  if (m.banks.length) {
    const maxBank = Math.max.apply(null, m.banks.map(b => b.amount).concat([1]));
    let bankHtml = m.banks.map(b => ekBar(b.bank, b.amount / maxBank, '#7FA6E8', summaryMoney(b.amount))).join('');
    if (m.missedList.length) {
      bankHtml += '<div style="height:6px;font-size:0">&nbsp;</div>' + ekNotes(m.missedList.slice(0, 3).map(e => ({ tone: 'info',
        text: e.label + ' paid with ' + e.paidWith + ' (' + summaryMoney(e.spent) + '): about ' + summaryMoney(e.cashback) +
          ' back with your ' + e.best + ' card.' })));
    }
    rows.push(ekSection('By bank', bankHtml));
  }

  // transfers + year to date
  rows.push(ekSection('Transfers', '<div style="font-size:14px;color:' + EK.text + '">' + m.transfersCount + ' transfer(s) out · ' +
    summaryEscape(summaryMoney(m.transfersOut)) + (m.openCount ? ' · <span style="color:' + ekTone('warn').fg + '">' + m.openCount +
      ' without a category (' + summaryEscape(summaryMoney(m.openTotal)) + ')</span>' : '') + '</div>'));
  rows.push(ekSection(m.month.getFullYear() + ' so far', ekKpis([
    { label: 'Spent', value: summaryMoneyShort(m.ytdTotal), sub: m.ytdMonths + ' month(s)' },
    { label: 'Per month', value: summaryMoneyShort(m.ytdAverage), sub: 'average' },
    { label: 'Kept', value: m.ytdSavingsRate !== null ? summaryPct(m.ytdSavingsRate) : '—', sub: 'of net income',
      tone: m.ytdSavingsRate === null ? null : m.ytdSavingsRate < 0 ? 'bad' : m.ytdSavingsRate < 0.1 ? 'warn' : 'good' }
  ])));
  T.push('', m.month.getFullYear() + ' so far: ' + summaryMoney(m.ytdTotal) + ' in ' + m.ytdMonths + ' month(s)');

  // investments (v1.1.32)
  if (m.investments && m.investments.none) {   // v1.1.49: the history starts after this month
    rows.push(ekSection('Investments', '<div style="font-size:13px;color:' + EK.muted + '">Your investment history starts on ' +
      summaryEscape(m.investments.firstDay) + ' — this section fills in from that month\'s summary, sent on ' +
      summaryEscape(m.investments.firstReport) + '.</div>'));
    T.push('', 'Investments: history starts on ' + m.investments.firstDay + ' — first monthly figures on ' + m.investments.firstReport);
  } else if (m.investments) {
    const iv = m.investments;
    let inner = ekKpis([
      { label: 'Value at month end', value: summaryUsd(iv.endValue), sub: iv.partial ? 'tracking began this month'
        : iv.added && iv.added.length ? 'added this month: ' + iv.added.join(', ') : 'from ' + summaryUsd(iv.startValue) },
      { label: 'Gain in ' + SUMMARY_MONTHS[m.month.getMonth()], value: summarySignedUsd(iv.gain),
        sub: (iv.gainPct !== null ? summarySignedPctFine(iv.gainPct) + ' · ' : '') + 'deposits left out', tone: iv.gain > 0 ? 'good' : iv.gain < 0 ? 'bad' : null },
      { label: 'Deposited', value: summaryUsd(iv.deposits), sub: 'dividends ' + summaryUsd(iv.dividends) + ' · fees ' + summaryUsd(iv.fees) }
    ]);
    if (iv.allocation.length) {
      inner += '<div style="height:12px;font-size:0">&nbsp;</div>' + iv.allocation.map(a =>
        ekBar(a.account, a.share, '#0F766E', summaryUsd(a.value), summaryEscape(summaryPct(a.share) + ' of the portfolio'))).join('');
    }
    if (iv.returns && iv.returns.periodReturn !== null) {
      inner += '<div style="font-size:12px;color:' + EK.muted + '">Since tracking began (' + summaryEscape(iv.returns.start) + '): ' +
        summaryEscape(summarySignedPctFine(iv.returns.periodReturn) + ' · ' + summarySignedUsd(iv.returns.gain)) +
        (iv.returns.annualized !== null ? ' · ' + summaryEscape(summarySignedPctFine(iv.returns.annualized)) + ' a year' : '') + '</div>';
    }
    rows.push(ekSection('Investments', inner));
    T.push('', 'Investments: ' + summaryUsd(iv.endValue) + ' at month end · gain ' + summarySignedUsd(iv.gain));
  }

  const recs = m.recommendations.length ? m.recommendations : [{ tone: 'good', text: 'Nothing to flag.' }];
  rows.push(ekSection('Recommendations', ekNotes(recs)));
  T.push('', 'Recommendations:', ...recs.map(r => '  • ' + r.text));

  const health = summaryDataHealth(m.lastRun, m.now || new Date());
  rows.push('<tr><td style="padding:20px 28px 0"><div style="font-size:12px;color:' + ekTone(health.tone).fg + ';background:' +
    ekTone(health.tone).bg + ';border-radius:8px;padding:9px 12px">' + summaryEscape(health.text) + '</div></td></tr>');

  const sheetName = opts.sheetName || 'the spreadsheet';
  const buttons = ekButtons([
    { text: 'Open ' + sheetName + ' →', url: links.dashboard, primary: true },
    { text: 'Transactions', url: links.transactions },
    { text: 'Custom Rules', url: m.openCount ? links.rules : null }
  ]);
  if (links.dashboard) T.push('', 'Open ' + sheetName + ': ' + links.dashboard);
  const preheader = summaryMoney(m.spent) + ' spent' + (m.leftOver !== null ? ' · ' + summaryMoney(m.leftOver) + ' left over' : '') +
    (m.vsPrev !== null ? ' · ' + summarySignedPct(m.vsPrev) + ' vs. ' + prevName : '');
  const html = ekShell({
    eyebrow: 'Monthly summary', title: name,
    subtitle: '1–' + m.daysInMonth + ' ' + SUMMARY_MONTHS[m.month.getMonth()] + ' · amounts in DOP-equivalent',
    preheader: preheader, body: rows.join(''), buttons: buttons,
    footer: 'Financial Tracker · amounts converted at the Dashboard\'s exchange rates. ' +
      'Turn this email off in 📊 Tracker › Setup Wizard.'
  });
  return { subject: stripAstral(subject), html: toAsciiHtml(html), text: stripAstral(T.join('\n')), preheader: preheader };
}

/** Builds and sends the monthly email for the month before `now`. Returns { to, subject }. */
function deliverMonthlySummary(now) {
  refreshAutoDeductions();
  const config = getConfig();
  if (!config) throw new Error('Setup not completed.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  const values = sheet ? sheet.getDataRange().getValues() : [[]];
  const rates = readDashboardRates(ss);
  const ref = now || new Date();
  let investments = null;   // v1.1.32
  try {
    investments = investmentsReportData('monthly', { rates: rates, month: new Date(ref.getFullYear(), ref.getMonth() - 1, 1, 12) });
  } catch (error) { Logger.log('Investments brief: ' + error); }
  const summary = computeMonthlySummary(values, {
    today: ref, rates: rates, netIncomeDop: computeNetIncomeDop(config, rates),
    cards: readDashboardCards(ss), lastRun: readLastRun(), now: ref, investments: investments
  });
  const mail = buildMonthlySummaryEmail(summary, { sheetName: ss.getName(), links: summaryLinks(ss) });
  const to = config.notifyEmail || config.email;
  GmailApp.sendEmail(to, mail.subject, mail.text, { htmlBody: mail.html, name: 'Financial Tracker' });
  Logger.log('🗓️ Monthly summary sent to ' + to + ' — ' + mail.subject);
  return { to: to, subject: mail.subject };
}

/** Trigger handler (1st of each month, at the hour chosen in the Setup Wizard). */
function sendMonthlySummary() {
  try {
    const config = getConfig();
    if (!config || !config.notifyMonthly) {
      Logger.log('Monthly summary is turned off — nothing sent.');
      return;
    }
    deliverMonthlySummary();
  } catch (error) {
    Logger.log('❌ Monthly summary failed: ' + error);
  }
}

/** Menu "🗓️ Send Monthly Summary Now" — last complete month, even when the monthly email is off. */
function sendMonthlySummaryNow() {
  const config = requireConfig();
  if (!config) return;
  try {
    const sent = deliverMonthlySummary();
    safeAlert('🗓️ Sent to ' + sent.to + '\n' + sent.subject);
  } catch (error) {
    safeAlert('❌ Could not send the summary: ' + error);
  }
}
