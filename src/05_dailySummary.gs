/**
 * DAILY SUMMARY EMAIL — v1.1.24
 *
 * A short email about the previous day, sent by its own daily trigger at the
 * hour chosen in the Setup Wizard, from the user's own Gmail account, with a
 * link to the spreadsheet. Each section can be turned on or off in the wizard.
 *
 * Numbers follow the Dashboard's rules exactly: spending = Transaction and
 * Transfer rows that have a category other than "Exclude" (Card Payment and
 * Cashback never count; reversals are negative rows and subtract), converted
 * to DOP-equivalent at the Dashboard's exchange rates. Transfers without a
 * category are listed on their own. Recommendations are fixed rules — no
 * guessing — and card tips only mention banks that have a cashback rate in the
 * Dashboard's credit-card table.
 *
 * Menu "📬 Send Daily Summary Now" sends the same email immediately (for a
 * test, or when the trigger is off).
 */
const SUMMARY_SECTIONS_DEFAULT = { totals: true, vsAverage: true, transfers: true, recommendations: true, cashback: true,
  investments: true };   // v1.1.32
const SUMMARY_MIN_HISTORY_DAYS = 7;
// v1.1.26: bills and fixed costs are left out of "day-to-day" comparisons — a rent
// or electricity payment would otherwise make any day look like an alarming spike.
const SUMMARY_BILL_CATEGORIES = ['Rent', 'Gym + Calisthenics', 'Telecommunications', 'Streaming & Subscriptions', 'Electricity'];
function isBillCategory(cat) { return SUMMARY_BILL_CATEGORIES.indexOf(cat) !== -1; }
const SUMMARY_ABOVE_USUAL = 1.05;   // v1.1.26: only flag a category once it's 5%+ over its usual month   // below this, "vs. average" isn't meaningful yet

// Kinds of purchase with an unambiguous best card (see "Which card for what").
const CARD_TIP_RULES = [
  { label: 'Pharmacies', bank: 'BHD', test: /\b(FARMACIAS?|FCIA|FARMA\w*|FARM|MEDICAR|HIDALGOS)\b/ },
  { label: 'Supermarkets', bank: 'BANESCO', test: /\b(SM|SUPERMERCADOS?|SUPER|JUMBO|BRAVO|POLA|PLAZA LAMA|NACIONAL)\b/ },
  { label: 'Electricity and water', bank: 'BANESCO', test: /\b(EDESUR|EDENORTE|EDEESTE|CAASD)\b/ }
];

const SUMMARY_DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const SUMMARY_MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function summaryMoney(n) {
  const v = Math.round(Math.abs(Number(n) || 0)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '−' : '') + 'RD$' + v;
}
function summaryPct(x) { return Math.round(Math.abs(x) * 100) + '%'; }
function summaryEscape(t) {
  return String(t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
/**
 * v1.1.25: GmailApp.sendEmail() breaks characters outside Unicode's Basic
 * Multilingual Plane — every emoji arrived as six "�" (its surrogate pair
 * encoded as two separate 3-byte sequences), while BMP characters such as
 * — → − came through fine. So the HTML body is sent as pure ASCII (each
 * non-ASCII character as a numeric entity, which every client decodes), and
 * the subject and plain-text body, where entities don't exist, carry no
 * emoji at all.
 */
function toAsciiHtml(html) {
  let out = '';
  for (const ch of String(html)) {            // for…of walks code points, keeping surrogate pairs together
    const cp = ch.codePointAt(0);
    out += cp < 128 ? ch : '&#' + cp + ';';
  }
  return out;
}
function stripAstral(text) {
  return String(text).replace(/[\u{10000}-\u{10FFFF}]\uFE0F?[ ]?/gu, '');
}

function keyToDate(key) {
  const m = String(key).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12) : null;
}

/** Net monthly income in DOP-equivalent — same arithmetic as the Dashboard's income box. */
function computeNetIncomeDop(config, rates) {
  const d = config.deductions || {};
  const gross = Number(config.monthlyIncome) || 0;
  const salary = config.incomeCurrency === 'DOP'
    ? gross - (d.ARS || 0) - (d.AFP || 0) - (d.ISR || 0)
    : (gross - (d.ARS || 0) - (d.AFP || 0) - gross * (d.taxRate || 0) / 100) * rates.USD;
  // v1.1.27: other income is added in full, in its own currency
  const other = Number(config.otherIncome) || 0;
  return salary + (config.otherIncomeCurrency === 'USD' ? other * rates.USD : other);
}

/**
 * Everything the email needs about the day before `opts.today`. Pure (no
 * Gmail/Sheets calls) — see tests/.
 * opts: { today: Date, rates: {USD, EUR, COP}, netIncomeDop, cards: [{bank, rate}] }
 */
function computeDailySummary(values, opts) {
  const rates = opts.rates;
  const today = opts.today;
  const day = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1, 12);
  const dayKey = normalizeDateForCompare(day);
  const shiftKey = n => normalizeDateForCompare(new Date(day.getFullYear(), day.getMonth(), day.getDate() + n, 12));
  const monthKey = dayKey.slice(0, 7);
  const prevMonths = [1, 2, 3].map(k => normalizeDateForCompare(new Date(day.getFullYear(), day.getMonth() - k, 1, 12)).slice(0, 7));
  const toDop = (amount, cur) => (Number(amount) || 0) * ({ USD: rates.USD, EUR: rates.EUR, COP: rates.COP }[cur] || 1);

  const s = {
    day: day, dayKey: dayKey, spent: 0, spentDayToDay: 0, purchases: 0, items: [], openTransfers: [], openTransfersTotal: 0,
    windowSpent: 0, mtd: 0, mtdByCategory: {}, prevByCategory: {}, prevMonthsWithData: 0,
    mtdOpenCount: 0, mtdOpenTotal: 0, firstKey: null,
    dayOfMonth: day.getDate(), daysInMonth: new Date(day.getFullYear(), day.getMonth() + 1, 0).getDate(),
    netIncomeDop: Number(opts.netIncomeDop) || 0
  };
  const monthsSeen = {};
  const windowStartKey = shiftKey(-30);
  const weekStartKey = shiftKey(-6);          // v1.1.26: last 7 days incl. yesterday
  const dayTotals = {};

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[TX_COL.DATE]) continue;
    const key = normalizeDateForCompare(r[TX_COL.DATE]);
    const type = r[TX_COL.TYPE] || 'Transaction';
    const cat = String(r[TX_COL.CATEGORY] || '');
    const amt = toDop(r[TX_COL.AMOUNT], r[TX_COL.CURRENCY]);
    const spend = (type === 'Transaction' || type === 'Transfer') && cat !== '' && cat !== EXCLUDE_CATEGORY;
    const open = type === 'Transfer' && cat === '';
    const month = key.slice(0, 7);
    const item = { merchant: String(r[TX_COL.MERCHANT] || ''), bank: String(r[TX_COL.BANK] || ''), category: cat, amount: amt, type: type };

    if (spend && (s.firstKey === null || key < s.firstKey)) s.firstKey = key;
    if (key === dayKey) {
      if (spend) {
        s.spent += amt; s.items.push(item); if (amt > 0) s.purchases++;
        if (!isBillCategory(cat)) s.spentDayToDay += amt;
      }
      if (open) { s.openTransfers.push(item); s.openTransfersTotal += amt; }
    }
    if (spend && !isBillCategory(cat) && key >= windowStartKey && key < dayKey) s.windowSpent += amt;
    if (spend && key >= weekStartKey && key <= dayKey) dayTotals[key] = (dayTotals[key] || 0) + amt;
    if (month === monthKey && key <= dayKey) {
      if (spend) { s.mtd += amt; s.mtdByCategory[cat] = (s.mtdByCategory[cat] || 0) + amt; }
      if (open) { s.mtdOpenCount++; s.mtdOpenTotal += amt; }
    }
    if (spend && prevMonths.indexOf(month) !== -1) {
      s.prevByCategory[cat] = (s.prevByCategory[cat] || 0) + amt;
      monthsSeen[month] = true;
    }
  }

  // Daily average over the 30 days before yesterday — or fewer, if tracking started later.
  const windowStart = s.firstKey && s.firstKey > windowStartKey ? s.firstKey : windowStartKey;
  const windowDays = Math.round((day - keyToDate(windowStart)) / 86400000);
  s.windowDays = Math.max(0, windowDays);
  s.avgDaily = s.windowDays >= SUMMARY_MIN_HISTORY_DAYS ? s.windowSpent / s.windowDays : null;
  // v1.1.26: day-to-day spending only (bills and fixed costs excluded on both sides)
  s.vsAverage = s.avgDaily && s.avgDaily > 0 && s.spentDayToDay > 0 ? s.spentDayToDay / s.avgDaily - 1 : null;
  s.prevMonthsWithData = Object.keys(monthsSeen).length;
  s.usualByCategory = {};
  if (s.prevMonthsWithData) {
    Object.keys(s.prevByCategory).forEach(c => { s.usualByCategory[c] = s.prevByCategory[c] / s.prevMonthsWithData; });
  }
  s.projected = s.dayOfMonth ? s.mtd / s.dayOfMonth * s.daysInMonth : 0;
  s.mtdPct = s.netIncomeDop > 0 ? s.mtd / s.netIncomeDop : null;
  s.projectedPct = s.netIncomeDop > 0 ? s.projected / s.netIncomeDop : null;
  s.items.sort((a, b) => b.amount - a.amount);
  // v1.1.26: extra material for the redesigned email
  s.elapsedPct = s.dayOfMonth / s.daysInMonth;
  s.last7 = [-6, -5, -4, -3, -2, -1, 0].map(n => {
    const date = new Date(day.getFullYear(), day.getMonth(), day.getDate() + n, 12);
    const key = normalizeDateForCompare(date);
    return { key: key, date: date, amount: dayTotals[key] || 0 };
  });
  s.mtdTop = Object.keys(s.mtdByCategory)
    .filter(c => !isBillCategory(c))
    .map(c => ({ cat: c, mtd: s.mtdByCategory[c], usual: s.usualByCategory[c] || 0 }))
    .filter(c => c.mtd > 0).sort((a, b) => b.mtd - a.mtd).slice(0, 5);
  s.mtdBills = Object.keys(s.mtdByCategory).filter(isBillCategory)
    .reduce((sum, c) => sum + s.mtdByCategory[c], 0);
  s.lastRun = opts.lastRun || null;
  s.now = opts.now || today;
  s.investments = opts.investments || null;   // v1.1.32
  s.recommendations = summaryRecommendations(s);
  s.tips = summaryCardTips(s.items, opts.cards || []);
  return s;
}

/** Fixed rules → [{icon, text}]. */
function summaryRecommendations(s) {
  const out = [];
  if (s.spent <= 0) {
    out.push({ icon: '🟢', text: 'No spending logged yesterday.' });
  } else if (s.vsAverage !== null && s.vsAverage > 0.5) {
    const top = s.items.filter(it => !isBillCategory(it.category))[0];
    out.push({ icon: '🟠', text: 'Day-to-day spending was ' + summaryPct(s.vsAverage) + ' above your daily average (' +
      summaryMoney(s.avgDaily) + '). Biggest item: ' + top.merchant + ' (' + summaryMoney(top.amount) + ').' });
  } else if (s.vsAverage !== null && s.vsAverage < -0.3) {
    out.push({ icon: '🟢', text: 'Nice — day-to-day spending was ' + summaryPct(s.vsAverage) + ' below your daily average.' });
  }
  if (s.projectedPct !== null && s.dayOfMonth >= 5) {
    if (s.projectedPct > 1) {
      out.push({ icon: '🔴', text: 'At this pace the month ends around ' + summaryMoney(s.projected) + ' — ' +
        summaryPct(s.projectedPct) + ' of your net income.' });
    } else if (s.projectedPct > 0.85) {
      out.push({ icon: '🟠', text: 'Month pace: about ' + summaryPct(s.projectedPct) + ' of your net income by month-end.' });
    } else {
      out.push({ icon: '🟢', text: 'On track: at this pace the month ends at about ' + summaryPct(s.projectedPct) +
        ' of your net income.' });
    }
  }
  Object.keys(s.mtdByCategory)
    .filter(c => !isBillCategory(c))
    .map(c => ({ cat: c, mtd: s.mtdByCategory[c], usual: s.usualByCategory[c] || 0 }))
    .filter(x => x.usual > 0 && x.mtd > x.usual * SUMMARY_ABOVE_USUAL)
    .sort((a, b) => b.mtd / b.usual - a.mtd / a.usual)
    .slice(0, 2)
    .forEach(x => out.push({ icon: '🟠', text: x.cat + ': ' + summaryMoney(x.mtd) +
      ' so far this month, already above your usual month (' + summaryMoney(x.usual) + ').' }));
  if (s.mtdOpenCount > 0) {
    out.push({ icon: '📝', text: s.mtdOpenCount + ' transfer(s) this month have no category (' + summaryMoney(s.mtdOpenTotal) +
      ') — add a Custom Rule, or use "Exclude" for transfers between your own accounts.' });
  }
  return out;
}

/** Card tips for yesterday's purchases, only for banks with a cashback rate in the card table. */
function summaryCardTips(items, cards) {
  const rateByBank = {}, nameByBank = {};
  cards.forEach(c => { if (c.rate > 0) { rateByBank[c.bank] = c.rate; nameByBank[c.bank] = c.name || ''; } });
  const tips = [];
  const seen = {};
  items.forEach(it => {
    if (it.type !== 'Transaction' || !(it.amount > 0) || seen[it.merchant]) return;
    const text = normalizeKeyword(it.merchant);
    const rule = CARD_TIP_RULES.find(r => r.test.test(text));
    if (!rule || !rateByBank[rule.bank] || it.bank === rule.bank) return;
    seen[it.merchant] = true;
    const rate = rateByBank[rule.bank];
    tips.push({ text: it.merchant + ' (' + summaryMoney(it.amount) + ' on ' + it.bank + '): ' + rule.label.toLowerCase() +
      ' earn ' + summaryPct(rate) + ' with your ' + rule.bank + (nameByBank[rule.bank] ? ' ' + nameByBank[rule.bank] : '') + ' card — about ' + summaryMoney(it.amount * rate) + ' back.' });
  });
  return tips;
}

/**
 * EMAIL KIT — v1.1.26. Small building blocks shared by the daily and the
 * monthly summary. Email clients (Gmail above all) only reliably render
 * TABLE layout with INLINE styles: no flexbox/grid, no <style> classes, no
 * scripts, no external images. Every block returns an HTML string; text
 * passed in is escaped here. The final HTML goes through toAsciiHtml().
 */
const EK = {
  font: "-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif",
  navy: '#1F3864', accent: '#4472C4', text: '#1F2937', muted: '#6B7280', faint: '#9CA3AF',
  line: '#E5E7EB', soft: '#F5F7FB', page: '#EEF1F6', track: '#E8ECF3',
  tones: {
    good: { fg: '#2E7D32', bg: '#E8F5E9', icon: '🟢' },
    warn: { fg: '#B45309', bg: '#FFF4E5', icon: '🟠' },
    bad:  { fg: '#B91C1C', bg: '#FDECEC', icon: '🔴' },
    info: { fg: '#1D4ED8', bg: '#EAF1FE', icon: '🔵' },
    note: { fg: '#4B5563', bg: '#F3F4F6', icon: '📝' }
  },
  heat: ['#F3F4F6', '#DBE6FA', '#B4CCF3', '#7FA6E8', '#4472C4', '#1F3864']
};
const SUMMARY_DAYS_FULL = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const SUMMARY_MONTHS_FULL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
                             'September', 'October', 'November', 'December'];

function ekTone(name) { return EK.tones[name] || EK.tones.note; }

function ekChip(text, tone) {
  const t = ekTone(tone);
  return '<span style="display:inline-block;padding:3px 10px;border-radius:999px;font-size:12px;font-weight:600;' +
    'line-height:18px;background:' + t.bg + ';color:' + t.fg + '">' + summaryEscape(text) + '</span>';
}

/** A section row inside the card: small caps title + content. */
function ekSection(title, inner) {
  return '<tr><td style="padding:22px 28px 0">' +
    '<div style="font-size:11px;font-weight:700;letter-spacing:.8px;text-transform:uppercase;color:' + EK.muted +
    ';margin:0 0 10px">' + summaryEscape(title) + '</div>' + inner + '</td></tr>';
}

/** Tiles in one row: [{label, value, sub, tone}] */
function ekKpis(tiles) {
  const w = Math.floor(100 / tiles.length);
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;' +
    'border-spacing:0"><tr>' + tiles.map((t, i) =>
      '<td width="' + w + '%" valign="top" style="padding:0 ' + (i < tiles.length - 1 ? '8px' : '0') + ' 0 0">' +
      '<div style="background:' + EK.soft + ';border-radius:10px;padding:12px 12px 11px">' +
      '<div style="font-size:11px;color:' + EK.muted + ';text-transform:uppercase;letter-spacing:.5px">' + summaryEscape(t.label) + '</div>' +
      '<div style="font-size:19px;font-weight:700;color:' + (t.tone ? ekTone(t.tone).fg : EK.navy) + ';margin:4px 0 2px;' +
      'white-space:nowrap">' + summaryEscape(t.value) + '</div>' +
      '<div style="font-size:12px;color:' + EK.muted + '">' + summaryEscape(t.sub || '') + '</div></div></td>').join('') +
    '</tr></table>';
}

/** Labelled horizontal bar: fraction 0..1 (clamped), colour, right-hand text. */
function ekBar(label, fraction, color, right, sub) {
  const pct = Math.max(0, Math.min(1, Number(fraction) || 0));
  const fill = pct > 0 ? '<div style="width:' + Math.max(2, Math.round(pct * 100)) + '%;height:8px;background:' + color +
    ';border-radius:999px;font-size:0;line-height:0">&nbsp;</div>' : '';
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 12px">' +
    '<tr><td style="font-size:14px;color:' + EK.text + ';padding:0 0 5px">' + summaryEscape(label) + '</td>' +
    '<td align="right" style="font-size:14px;font-weight:600;color:' + EK.text + ';padding:0 0 5px;white-space:nowrap">' +
    summaryEscape(right) + '</td></tr>' +
    '<tr><td colspan="2" style="background:' + EK.track + ';border-radius:999px;height:8px;font-size:0;line-height:0">' +
    fill + '</td></tr>' + (sub ? '<tr><td colspan="2" style="font-size:12px;color:' + EK.muted + ';padding:4px 0 0">' + sub +
    '</td></tr>' : '') + '</table>';
}

/** Vertical mini columns: [{label, sub, value, highlight}] — max height 70 px. */
function ekColumns(cols) {
  const max = Math.max.apply(null, cols.map(c => c.value).concat([1]));
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' + cols.map(c => {
    const h = c.value > 0 ? Math.max(4, Math.round(c.value / max * 70)) : 2;
    const color = c.highlight ? EK.accent : (c.value > 0 ? '#B4CCF3' : EK.line);
    return '<td valign="bottom" align="center" style="padding:0 3px;height:92px">' +
      '<div style="font-size:10px;color:' + (c.highlight ? EK.navy : EK.faint) + ';font-weight:' + (c.highlight ? 700 : 400) +
      ';margin:0 0 3px;white-space:nowrap">' + summaryEscape(c.sub) + '</div>' +
      '<div style="height:' + h + 'px;background:' + color + ';border-radius:5px 5px 2px 2px;font-size:0;line-height:0">&nbsp;</div></td>';
  }).join('') + '</tr><tr>' + cols.map(c =>
    '<td align="center" style="font-size:11px;padding:6px 0 0;color:' + (c.highlight ? EK.navy : EK.muted) +
    ';font-weight:' + (c.highlight ? 700 : 400) + '">' + summaryEscape(c.label) + '</td>').join('') + '</tr></table>';
}

/** List rows: [{title, meta, right, rightTone}] */
function ekList(rows) {
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + rows.map((r, i) =>
    '<tr><td style="padding:9px 0;' + (i ? 'border-top:1px solid ' + EK.line + ';' : '') + '">' +
    '<div style="font-size:14px;font-weight:600;color:' + EK.text + '">' + summaryEscape(r.title) + '</div>' +
    (r.meta ? '<div style="font-size:12px;color:' + EK.muted + ';margin-top:2px">' + summaryEscape(r.meta) + '</div>' : '') +
    '</td><td align="right" valign="top" style="padding:9px 0 9px 12px;white-space:nowrap;font-size:14px;font-weight:700;color:' +
    (r.rightTone ? ekTone(r.rightTone).fg : EK.text) + ';' + (i ? 'border-top:1px solid ' + EK.line + ';' : '') + '">' +
    summaryEscape(r.right) + '</td></tr>').join('') + '</table>';
}

/** Coloured note cards: [{tone, text}] */
function ekNotes(notes) {
  return notes.map(n => {
    const t = ekTone(n.tone);
    return '<div style="border-left:4px solid ' + t.fg + ';background:' + t.bg + ';padding:10px 12px;border-radius:6px;' +
      'margin:0 0 8px;font-size:14px;line-height:1.45;color:' + EK.text + '">' + t.icon + '&nbsp; ' + summaryEscape(n.text) + '</div>';
  }).join('');
}

/** Buttons: [{text, url, primary}] (entries without a url are skipped). */
function ekButtons(buttons) {
  return buttons.filter(b => b.url).map(b =>
    '<a href="' + summaryEscape(b.url) + '" style="display:inline-block;margin:0 8px 8px 0;padding:11px 16px;border-radius:8px;' +
    'font-size:14px;font-weight:600;text-decoration:none;' + (b.primary
      ? 'background:' + EK.accent + ';color:#FFFFFF;border:1px solid ' + EK.accent
      : 'background:#FFFFFF;color:' + EK.accent + ';border:1px solid #C7D5EE') + '">' + summaryEscape(b.text) + '</a>').join('');
}

/** The whole email: hidden preheader, header band, content rows, footer. */
function ekShell(o) {
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<meta name="color-scheme" content="light only"><title>' + summaryEscape(o.title) + '</title></head>' +
    '<body style="margin:0;padding:0;background:' + EK.page + '">' +
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0;color:' + EK.page + '">' + summaryEscape(o.preheader || '') +
    '&#8199;&#65279;&#847;&#8199;&#65279;&#847;&#8199;&#65279;&#847;&#8199;&#65279;&#847;</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + EK.page + '"><tr>' +
    '<td align="center" style="padding:24px 10px">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;background:#FFFFFF;' +
    'border-radius:14px;overflow:hidden;font-family:' + EK.font + ';color:' + EK.text + '">' +
    '<tr><td style="background:' + EK.navy + ';padding:22px 28px 20px">' +
    '<div style="font-size:12px;color:#B4CCF3;letter-spacing:.6px;text-transform:uppercase">' + summaryEscape(o.eyebrow) + '</div>' +
    '<div style="font-size:22px;font-weight:700;color:#FFFFFF;margin-top:4px">' + summaryEscape(o.title) + '</div>' +
    (o.subtitle ? '<div style="font-size:13px;color:#C9D6EE;margin-top:3px">' + summaryEscape(o.subtitle) + '</div>' : '') +
    '</td></tr>' + o.body +
    '<tr><td style="padding:22px 28px 26px">' + (o.buttons || '') + '</td></tr></table>' +
    '<div style="max-width:600px;font-family:' + EK.font + ';font-size:11px;line-height:1.5;color:' + EK.faint +
    ';padding:14px 12px 0">' + (o.footer || '') + '</div></td></tr></table></body></html>';
}

function summaryLongDate(d) {
  return SUMMARY_DAYS_FULL[d.getDay()] + ' ' + d.getDate() + ' ' + SUMMARY_MONTHS_FULL[d.getMonth()] + ' ' + d.getFullYear();
}
function summaryShortDate(d) {
  return SUMMARY_DAYS[d.getDay()] + ' ' + d.getDate() + ' ' + SUMMARY_MONTHS[d.getMonth()];
}
function summaryTime(d) {
  const h = d.getHours(), m = d.getMinutes();
  return (h % 12 || 12) + ':' + (m < 10 ? '0' : '') + m + (h < 12 ? ' AM' : ' PM');
}
/** Compact money for small spaces: RD$1.2K, RD$45K, RD$1.3M. */
function summaryMoneyShort(n) {
  const a = Math.abs(Number(n) || 0);
  const sign = n < 0 ? '−' : '';
  if (a >= 1e6) return sign + 'RD$' + (a / 1e6).toFixed(1).replace(/\.0$/, '') + 'M';
  if (a >= 1e4) return sign + 'RD$' + Math.round(a / 1e3) + 'K';
  if (a >= 1e3) return sign + 'RD$' + (a / 1e3).toFixed(1).replace(/\.0$/, '') + 'K';
  return sign + 'RD$' + Math.round(a);
}
/** v1.1.32: US$ amounts for the investments section. */
function summaryUsd(n) {
  const v = Number(n) || 0;
  return (v < 0 ? '−' : '') + 'US$' + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function summarySignedUsd(n) { return (n > 0 ? '+' : '') + summaryUsd(n); }
/** Investment moves are mostly under 1% a day: two decimals (spending percentages stay whole numbers). */
function summarySignedPctFine(x) { const v = Number(x) || 0; return (v >= 0 ? '+' : '−') + Math.abs(v * 100).toFixed(2) + '%'; }

function summarySignedPct(x) {
  if (Math.round(Math.abs(x) * 100) === 0) return '±0%';   // v1.1.26: no "−0%"
  return (x >= 0 ? '+' : '−') + summaryPct(x);
}

/**
 * v1.1.26: one line about the morning update (written by runGmailMonitorCore()
 * into Script Properties) — so the email says when the data is stale or when
 * an email couldn't be read. Returns {tone, text}.
 */
function summaryDataHealth(lastRun, now) {
  if (!lastRun || !lastRun.at) {
    return { tone: 'note', text: 'Update status appears here after the next automatic morning update.' };
  }
  const at = new Date(lastRun.at);
  const hours = (now - at) / 3600000;
  const when = (hours < 24 && at.getDate() === now.getDate() ? 'today' : summaryShortDate(at)) + ' at ' + summaryTime(at);
  if (hours > 26) {
    return { tone: 'warn', text: 'No automatic update since ' + when + ' — numbers may be incomplete. Check Extensions › Apps Script › Executions.' };
  }
  let text = 'Data updated ' + when + ' · ' + (lastRun.saved || 0) + ' new transaction(s)';
  if (lastRun.unparsed) text += ' · ' + lastRun.unparsed + ' email(s) couldn\'t be read (left unread in Gmail)';
  if (lastRun.errors) text += ' · ' + lastRun.errors + ' error(s) in the log';
  if (lastRun.unrecognized) text += ' · ' + lastRun.unrecognized + ' email(s) in Unrecognized';   // v1.1.35
  return { tone: lastRun.unparsed || lastRun.errors || lastRun.unrecognized ? 'warn' : 'good', text: text };
}

/** Subject, HTML and plain-text bodies of the daily email. Pure — see tests/. */
function buildDailySummaryEmail(s, opts) {
  opts = opts || {};
  const sec = Object.assign({}, SUMMARY_SECTIONS_DEFAULT, opts.sections || {});
  const links = opts.links || { dashboard: opts.sheetUrl };
  const d = s.day;
  const vs = s.vsAverage === null ? '' : ' (' + summarySignedPct(s.vsAverage) + ' vs. average)';
  const subject = summaryShortDate(d) + ' ' + d.getFullYear() + ': ' + summaryMoney(s.spent) + ' spent' + vs;
  const rows = [], T = [];

  // ---- hero
  if (sec.totals) {
    let chip = '';
    if (s.spent <= 0) chip = ekChip('No spending logged', 'good');
    else if (s.spentDayToDay <= 0) chip = ekChip('Only bills and fixed costs', 'info');
    else if (s.avgDaily === null) chip = ekChip('Average available after ' + SUMMARY_MIN_HISTORY_DAYS + ' days of history', 'note');
    else if (s.vsAverage > 0.5) chip = ekChip('Day-to-day ' + summaryPct(s.vsAverage) + ' above your average', 'warn');
    else if (s.vsAverage < -0.1) chip = ekChip('Day-to-day ' + summaryPct(s.vsAverage) + ' below your average', 'good');
    else chip = ekChip('Day-to-day close to your average', 'info');
    rows.push('<tr><td style="padding:24px 28px 0">' +
      '<div style="font-size:13px;color:' + EK.muted + '">Spent yesterday · ' + s.purchases + ' purchase(s)</div>' +
      '<div style="font-size:40px;font-weight:800;color:' + EK.navy + ';line-height:1.15;margin:4px 0 10px">' +
      summaryMoney(s.spent) + '</div>' + chip + '</td></tr>');
    T.push('Yesterday (' + summaryLongDate(d) + '): ' + summaryMoney(s.spent) + ' · ' + s.purchases + ' purchase(s)' +
      (s.avgDaily !== null ? ' · daily average ' + summaryMoney(s.avgDaily) + vs : ''));
  }

  // ---- month KPIs + pace
  if (sec.vsAverage) {
    const tiles = [{ label: 'This month', value: summaryMoneyShort(s.mtd),
      sub: s.mtdPct !== null ? summaryPct(s.mtdPct) + ' of net income' : 'day ' + s.dayOfMonth + ' of ' + s.daysInMonth }];
    tiles.push({ label: 'Month-end pace', value: s.dayOfMonth >= 5 ? summaryMoneyShort(s.projected) : '—',
      sub: s.dayOfMonth < 5 ? 'from day 5' : (s.projectedPct !== null ? summaryPct(s.projectedPct) + ' of net income' : 'projected'),
      tone: s.projectedPct !== null && s.dayOfMonth >= 5 ? (s.projectedPct > 1 ? 'bad' : s.projectedPct > 0.85 ? 'warn' : null) : null });
    tiles.push({ label: 'Daily average', value: s.avgDaily !== null ? summaryMoneyShort(s.avgDaily) : '—',
      sub: s.avgDaily !== null ? 'day-to-day, ' + s.windowDays + ' days' : 'not enough history yet' });
    let inner = ekKpis(tiles);
    if (s.mtdPct !== null) {
      const used = s.mtdPct, elapsed = s.elapsedPct;
      const color = used > 1 ? ekTone('bad').fg : used > elapsed + 0.05 ? ekTone('warn').fg : ekTone('good').fg;
      inner += '<div style="height:14px;font-size:0">&nbsp;</div>' +
        ekBar('Net income used', used, color, summaryPct(used)) +
        ekBar('Month elapsed', elapsed, '#C7D5EE', summaryPct(elapsed));
    }
    rows.push(ekSection('This month', inner));
    T.push('', 'This month so far: ' + summaryMoney(s.mtd) + (s.mtdPct !== null ? ' (' + summaryPct(s.mtdPct) + ' of net income)' : ''));
    if (s.dayOfMonth >= 5) T.push('At this pace: ' + summaryMoney(s.projected) + (s.projectedPct !== null ? ' (' + summaryPct(s.projectedPct) + ' of net income)' : '') + ' by month-end');
    if (s.avgDaily === null) T.push('Daily average: not enough history yet (needs ' + SUMMARY_MIN_HISTORY_DAYS + ' days).');
  }

  // ---- last 7 days + purchases
  if (sec.totals) {
    rows.push(ekSection('Last 7 days', ekColumns(s.last7.map((x, i) => ({
      label: SUMMARY_DAYS[x.date.getDay()], sub: x.amount ? summaryMoneyShort(x.amount) : '·',
      value: Math.max(0, x.amount), highlight: i === s.last7.length - 1 })))));
    if (s.items.length) {
      const shown = s.items.slice(0, 6);
      let inner = ekList(shown.map(it => ({ title: it.merchant, meta: it.bank + ' · ' + it.category,
        right: summaryMoney(it.amount), rightTone: it.amount < 0 ? 'good' : null })));
      if (s.items.length > shown.length) {
        inner += '<div style="font-size:12px;color:' + EK.muted + ';padding-top:6px">+ ' + (s.items.length - shown.length) + ' more in Transactions</div>';
      }
      rows.push(ekSection('Yesterday\'s purchases', inner));
      shown.forEach(it => T.push('  • ' + it.merchant + ' — ' + it.bank + ' — ' + summaryMoney(it.amount)));
    }
  }

  // ---- categories this month vs usual
  if (sec.vsAverage && s.mtdTop.length) {
    const max = Math.max.apply(null, s.mtdTop.map(c => Math.max(c.mtd, c.usual || 0)).concat([1]));
    rows.push(ekSection('Where this month is going', s.mtdTop.map(c => {
      const over = c.usual > 0 && c.mtd > c.usual * SUMMARY_ABOVE_USUAL;
      const sub = c.usual > 0 ? (over ? '<span style="color:' + ekTone('warn').fg + '">above usual ' +
        summaryEscape(summaryMoneyShort(c.usual)) + '</span>' : 'usual ' + summaryEscape(summaryMoneyShort(c.usual))) : '';
      return ekBar(c.cat, c.mtd / max, over ? ekTone('warn').fg : EK.accent, summaryMoney(c.mtd), sub);
    }).join('') + (s.mtdBills > 0 ? '<div style="font-size:12px;color:' + EK.muted + '">Plus bills and fixed costs: ' +
      summaryEscape(summaryMoney(s.mtdBills)) + ' (rent, gym, phone, subscriptions, electricity).</div>' : '')));
  }

  // ---- transfers to review
  if (sec.transfers) {
    let inner;
    if (s.openTransfers.length) {
      inner = ekList(s.openTransfers.map(t => ({ title: t.merchant, meta: t.bank + ' · no category', right: summaryMoney(t.amount) })));
    } else {
      inner = '<div style="font-size:14px;color:' + EK.muted + '">None yesterday.' +
        (s.mtdOpenCount ? ' ' + s.mtdOpenCount + ' still open this month (' + summaryMoney(s.mtdOpenTotal) + ').' : '') + '</div>';
    }
    rows.push(ekSection('Transfers to review', inner));
    T.push('', 'Transfers to review: ' + (s.openTransfers.length
      ? s.openTransfers.map(t => t.merchant + ' ' + summaryMoney(t.amount)).join(', ') : 'none yesterday'));
  }

  // ---- recommendations + tips
  if (sec.recommendations) {
    const tone = { '🟢': 'good', '🟠': 'warn', '🔴': 'bad', '📝': 'note' };
    const recs = s.recommendations.length ? s.recommendations : [{ icon: '🟢', text: 'Nothing to flag.' }];
    rows.push(ekSection('Recommendations', ekNotes(recs.map(r => ({ tone: tone[r.icon] || 'info', text: r.text })))));
    T.push('', 'Recommendations:', ...recs.map(r => '  • ' + r.text));
  }
  if (sec.cashback) {
    rows.push(ekSection('Card & cashback tips', s.tips.length
      ? ekNotes(s.tips.map(t => ({ tone: 'info', text: t.text })))
      : '<div style="font-size:14px;color:' + EK.muted + '">No card tips for yesterday.</div>'));
    T.push('', 'Card tips:', ...(s.tips.length ? s.tips.map(t => '  • ' + t.text) : ['  • none']));
  }

  // ---- investments (v1.1.32)
  if (sec.investments && s.investments) {
    const iv = s.investments;
    const since = iv.prevDay ? summaryShortDate(keyToDate(iv.prevDay)) : null;
    let inner = ekKpis([
      { label: 'Portfolio', value: summaryUsd(iv.total), sub: iv.totalDop ? summaryMoney(iv.totalDop) : '' },
      { label: since ? 'Since ' + since : 'Change', value: iv.change === null ? '—' : summarySignedUsd(iv.change),
        sub: iv.change === null ? 'first day recorded' : summarySignedPctFine(iv.changePct || 0) + (iv.deposits ? ' · deposits left out' : ''),
        tone: iv.change > 0 ? 'good' : iv.change < 0 ? 'bad' : null },
      { label: 'Since tracking began', value: iv.returns && iv.returns.periodReturn !== null ? summarySignedPctFine(iv.returns.periodReturn) : '—',
        sub: iv.returns ? summarySignedUsd(iv.returns.gain) : '' }
    ]);
    if (iv.movers.length) {
      inner += '<div style="height:10px;font-size:0">&nbsp;</div>' + ekList(iv.movers.map(m => ({
        title: m.ticker, meta: m.account + ' · biggest move' + (since ? ' since ' + since : ''), right: summarySignedPctFine(m.change),
        rightTone: m.change >= 0 ? 'good' : 'bad' })));
    }
    rows.push(ekSection('Investments', inner));
    T.push('', 'Investments: ' + summaryUsd(iv.total) + (iv.change !== null ? ' (' + summarySignedUsd(iv.change) + ' since ' + since + ')' : ''));
  }

  // ---- data health (always)
  const health = summaryDataHealth(s.lastRun, s.now || new Date());
  rows.push('<tr><td style="padding:20px 28px 0"><div style="font-size:12px;color:' + ekTone(health.tone).fg + ';background:' +
    ekTone(health.tone).bg + ';border-radius:8px;padding:9px 12px">' + summaryEscape(health.text) + '</div></td></tr>');
  T.push('', health.text);

  const name = opts.sheetName || 'the spreadsheet';
  const buttons = ekButtons([
    { text: 'Open ' + name + ' →', url: links.dashboard, primary: true },
    { text: 'Transactions', url: links.transactions },
    { text: 'Custom Rules', url: sec.transfers && s.mtdOpenCount ? links.rules : null }
  ]);
  if (links.dashboard) T.push('', 'Open ' + name + ': ' + links.dashboard);

  const preheader = summaryMoney(s.spent) + ' yesterday' + (s.mtdPct !== null
    ? ' · ' + summaryPct(s.mtdPct) + ' of net income used with ' + summaryPct(s.elapsedPct) + ' of the month gone' : '');
  const html = ekShell({
    eyebrow: 'Daily summary', title: summaryLongDate(d), subtitle: 'Amounts in DOP-equivalent',
    preheader: preheader, body: rows.join(''), buttons: buttons,
    footer: 'Financial Tracker · amounts converted at the Dashboard\'s exchange rates. ' +
      'Change what this email includes, or turn it off, in 📊 Tracker › Setup Wizard.'
  });
  return { subject: stripAstral(subject), html: toAsciiHtml(html), text: stripAstral(T.join('\n')), preheader: preheader };
}

/** Exchange rates as set on the Dashboard (defaults before the Dashboard exists). */
function readDashboardRates(ss) {
  const get = (name, key) => {
    const r = ss.getRangeByName(name);
    const v = r ? Number(r.getValue()) : 0;
    return v > 0 ? v : DASH_DEFAULT_RATES[key];
  };
  return { USD: get('RATE_USD', 'usd'), EUR: get('RATE_EUR', 'eur'), COP: get('RATE_COP', 'cop') };
}

/**
 * The user's cards: [{bank, rate, name}]. v1.1.27: from the Setup Wizard
 * (Configuration "cards" + the card catalogue); before the wizard has cards, from
 * an older Dashboard's card table.
 */
function readDashboardCards(ss) {
  const config = getConfig();
  if (config && Array.isArray(config.cards)) {
    return resolveCards(config.cards).map(c => ({ bank: c.bank, rate: c.cashback, name: c.name }));
  }
  const r = ss.getRangeByName('DASH_CARDS');
  if (!r) return [];
  return r.getValues()
    .map(row => ({ bank: String(row[0] || '').trim().toUpperCase(), rate: Number(row[1]) || 0 }))
    .filter(c => c.bank);
}

/** v1.1.26: deep links to the sheets the emails point at (missing sheets are left out). */
function summaryLinks(ss) {
  const url = ss.getUrl();
  const gid = name => { const sh = ss.getSheetByName(name); return sh ? url + '#gid=' + sh.getSheetId() : null; };
  return { dashboard: gid('Dashboard') || url, transactions: gid(TRANSACTIONS_SHEET), transfers: gid('Bank Transfers'),
           rules: gid(CUSTOM_RULES_SHEET) };
}

/** v1.1.26: outcome of the last monitor run, saved by runGmailMonitorCore(). */
const LAST_RUN_PROPERTY = 'FT_LAST_RUN';
function recordLastRun(info) {
  try {
    PropertiesService.getScriptProperties().setProperty(LAST_RUN_PROPERTY, JSON.stringify(info));
  } catch (error) {
    Logger.log('Could not record the run status: ' + error);
  }
}
function readLastRun() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(LAST_RUN_PROPERTY);
    return raw ? JSON.parse(raw) : null;
  } catch (error) {
    return null;
  }
}

/** Builds and sends the email. Returns { to, subject }. */
function deliverDailySummary(now) {
  refreshAutoDeductions();
  const config = getConfig();
  if (!config) throw new Error('Setup not completed.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  const values = sheet ? sheet.getDataRange().getValues() : [[]];
  const rates = readDashboardRates(ss);
  let investments = null;   // v1.1.32
  try { investments = investmentsReportData('daily', { rates: rates }); } catch (error) { Logger.log('Investments brief: ' + error); }
  const summary = computeDailySummary(values, {
    today: now || new Date(), rates: rates, netIncomeDop: computeNetIncomeDop(config, rates), cards: readDashboardCards(ss),
    lastRun: readLastRun(), now: now || new Date(), investments: investments
  });
  const mail = buildDailySummaryEmail(summary, {
    sections: config.notifySections, sheetName: ss.getName(), links: summaryLinks(ss)
  });
  const to = config.notifyEmail || config.email;
  GmailApp.sendEmail(to, mail.subject, mail.text, { htmlBody: mail.html, name: 'Financial Tracker' });
  Logger.log('📬 Daily summary sent to ' + to + ' — ' + mail.subject);
  return { to: to, subject: mail.subject };
}

/** Trigger handler (daily, at the hour chosen in the Setup Wizard). */
function sendDailySummary() {
  try {
    const config = getConfig();
    if (!config || !config.notifyEnabled) {
      Logger.log('Daily summary is turned off — nothing sent.');
      return;
    }
    deliverDailySummary();
  } catch (error) {
    Logger.log('❌ Daily summary failed: ' + error);
  }
}

/** Menu "📬 Send Daily Summary Now" — works even when the daily email is turned off. */
function sendDailySummaryNow() {
  const config = requireConfig();
  if (!config) return;
  try {
    const sent = deliverDailySummary();
    safeAlert('📬 Sent to ' + sent.to + '\n' + sent.subject);
  } catch (error) {
    safeAlert('❌ Could not send the summary: ' + error);
  }
}
