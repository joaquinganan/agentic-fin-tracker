'use strict';
/**
 * Strict in-memory mock of the Apps Script services the system uses.
 * "Strict" means:
 *  - only methods that exist in the real API are implemented — calling
 *    anything else throws (a typo like setFormulae() fails here, not in
 *    production);
 *  - writing outside a sheet's current size throws, like the real
 *    "coordinates of the range are outside the dimensions of the sheet"
 *    error — so a missing ensureRowCapacity() is caught;
 *  - SpreadsheetApp.getUi() throws, as it does in a time-triggered run, so
 *    every alert must go through safeAlert()/safeToast().
 * Formulas are stored as text (nothing is evaluated).
 */

// Formatting calls that only change appearance — accepted and chained.
const STYLE_METHODS = [
  'setNumberFormat', 'setFontWeight', 'setFontSize', 'setFontColor', 'setFontStyle', 'setBackground',
  'setHorizontalAlignment', 'setVerticalAlignment', 'setDataValidation', 'clearDataValidations', 'setWrap',
  'setNote', 'setFontFamily', 'setBorder'
];

const overlaps = (a, b) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
const same = (a, b) => a.r1 === b.r1 && a.r2 === b.r2 && a.c1 === b.c1 && a.c2 === b.c2;

function colToNum(letters) {
  return letters.split('').reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
}

function parseA1(a1) {
  const clean = a1.replace(/\$/g, '');
  const m = clean.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (!m) throw new Error('Mock: unsupported A1 notation ' + a1);
  const r1 = Number(m[2]), c1 = colToNum(m[1]);
  const r2 = m[3] ? Number(m[4]) : r1, c2 = m[3] ? colToNum(m[3]) : c1;
  return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
}

class FakeFilter {
  constructor(range) { this.range = range; this.criteria = {}; }
  getRange() { return this.range; }
  remove() { this.range.sheet.filter = null; }
  getColumnFilterCriteria(col) { return this.criteria[col] || null; }
  setColumnFilterCriteria(col, c) { this.criteria[col] = c; return this; }
}

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    if (row < 1 || col < 1 || numRows < 1 || numCols < 1) throw new Error('Mock: invalid range');
    if (row + numRows - 1 > sheet.maxRows || col + numCols - 1 > sheet.maxCols) {
      throw new Error(`Mock: range R${row}C${col}:${numRows}x${numCols} is outside the dimensions of sheet "${sheet.name}" (${sheet.maxRows}x${sheet.maxCols})`);
    }
    Object.assign(this, { sheet, row, col, numRows, numCols });
    STYLE_METHODS.forEach(m => { this[m] = (...args) => { this.sheet.styles.push({ m, rect: this._rect(), v: args[0], args }); return this; }; });
  }
  _rect() { return { r1: this.row, c1: this.col, r2: this.getLastRow(), c2: this.getLastColumn() }; }
  _addMerge(rect) {
    if (rect.r1 === rect.r2 && rect.c1 === rect.c2) throw new Error('Mock: cannot merge a single cell');
    for (const m of this.sheet.merges) {
      if (same(m, rect)) return;
      if (overlaps(m, rect)) throw new Error('Mock: merge overlaps an existing merged range (Sheets rejects this)');
    }
    this.sheet.merges.push(rect);
  }
  merge() { this._addMerge(this._rect()); return this; }
  mergeAcross() {
    for (let r = this.row; r <= this.getLastRow(); r++) this._addMerge({ r1: r, c1: this.col, r2: r, c2: this.getLastColumn() });
    return this;
  }
  breakApart() { const me = this._rect(); this.sheet.merges = this.sheet.merges.filter(m => !overlaps(m, me)); return this; }
  setBackgrounds(colors) { this._checkShape(colors, 'setBackgrounds'); this.sheet.styles.push({ m: 'setBackgrounds', rect: this._rect(), v: colors }); return this; }
  setFontColors(colors) { this._checkShape(colors, 'setFontColors'); this.sheet.styles.push({ m: 'setFontColors', rect: this._rect(), v: colors }); return this; }
  getSheet() { return this.sheet; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getNumRows() { return this.numRows; }
  getNumColumns() { return this.numCols; }
  getLastRow() { return this.row + this.numRows - 1; }
  getLastColumn() { return this.col + this.numCols - 1; }
  _each(fn) {
    for (let i = 0; i < this.numRows; i++) for (let j = 0; j < this.numCols; j++) fn(this.row + i, this.col + j, i, j);
  }
  getValues() {
    const out = Array.from({ length: this.numRows }, () => new Array(this.numCols).fill(''));
    this._each((r, c, i, j) => { out[i][j] = this.sheet._get(r, c); });
    return out;
  }
  getValue() { return this.sheet._get(this.row, this.col); }
  _checkShape(arr, what) {
    if (!Array.isArray(arr) || arr.length !== this.numRows || arr.some(r => !Array.isArray(r) || r.length !== this.numCols)) {
      throw new Error(`Mock: ${what} shape does not match range ${this.numRows}x${this.numCols}`);
    }
  }
  setValues(values) { this._checkShape(values, 'setValues'); this._each((r, c, i, j) => this.sheet._write(r, c, values[i][j])); return this; }
  setFormulas(formulas) {
    this._checkShape(formulas, 'setFormulas');
    this._each((r, c, i, j) => {
      if (typeof formulas[i][j] !== 'string' || !formulas[i][j].startsWith('=')) throw new Error('Mock: not a formula: ' + formulas[i][j]);
      this.sheet._write(r, c, formulas[i][j]);
    });
    return this;
  }
  /** True when this range IS one merged block — then only its top-left cell holds the value. */
  _isOneMerge() { const me = this._rect(); return this.sheet.merges.some(m => same(m, me)); }
  setValue(v) {
    if (this._isOneMerge()) { this.sheet._write(this.row, this.col, v); return this; }
    this._each((r, c) => this.sheet._write(r, c, v));
    return this;
  }
  setFormula(f) {
    if (typeof f !== 'string' || !f.startsWith('=')) throw new Error('Mock: not a formula: ' + f);
    if (this._isOneMerge()) { this.sheet._write(this.row, this.col, f); return this; }
    this._each((r, c) => this.sheet._write(r, c, f));
    return this;
  }
  clearContent() { this._each((r, c) => this.sheet._set(r, c, '')); return this; }
  createFilter() {
    if (this.sheet.filter) throw new Error('Mock: sheet already has a filter');
    this.sheet.filter = new FakeFilter(this);
    return this.sheet.filter;
  }
  sort(spec) {
    const rows = this.getValues();
    const idx = spec.column - this.col;
    rows.sort((a, b) => {
      const x = a[idx] instanceof Object ? +a[idx] : a[idx];
      const y = b[idx] instanceof Object ? +b[idx] : b[idx];
      return (x < y ? -1 : x > y ? 1 : 0) * (spec.ascending === false ? -1 : 1);
    });
    this.setValues(rows);
    return this;
  }
}

class FakeSheet {
  constructor(ss, name) {
    Object.assign(this, { ss, name, maxRows: 1000, maxCols: 26, cells: new Map(), filter: null, frozenRows: 0,
      hidden: [], hiddenRows: new Set(), merges: [], charts: [], cfRules: [], rowHeights: {}, widths: {}, styles: [],
      id: 1000 + ss.sheets.length });
  }
  /** A write into a cell hidden under a merge would be invisible in Sheets — treat it as a layout bug. */
  _write(r, c, v) {
    if (v !== '' && v !== null && v !== undefined) {
      const covered = this.merges.find(m => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2 && !(r === m.r1 && c === m.c1));
      if (covered) throw new Error(`Mock: wrote "${String(v).slice(0, 40)}" to R${r}C${c}, hidden under merge R${covered.r1}C${covered.c1}`);
    }
    this._set(r, c, v);
  }
  getParent() { return this.ss; }
  hideRows(r, n) { for (let i = 0; i < (n || 1); i++) this.hiddenRows.add(r + i); }
  showRows(r, n) { for (let i = 0; i < (n || 1); i++) this.hiddenRows.delete(r + i); }
  setRowHeight(r, h) { this.rowHeights[r] = h; return this; }
  setRowHeights(r, n, h) { for (let i = 0; i < n; i++) this.rowHeights[r + i] = h; return this; }
  getCharts() { return this.charts.slice(); }
  removeChart(c) { this.charts = this.charts.filter(x => x !== c); }
  insertChart(c) { this.charts.push(c); }
  newChart() {
    const cfg = { ranges: [], options: {} };
    const b = {
      setChartType: t => { cfg.type = t; return b; },
      addRange: r => { cfg.ranges.push(r); return b; },
      setNumHeaders: h => { cfg.numHeaders = h; return b; },
      setPosition: (row, col, x, y) => { cfg.position = { row, col, x, y }; return b; },
      setOption: (k, v) => { cfg.options[k] = v; return b; },
      build: () => ({ cfg })
    };
    return b;
  }
  setConditionalFormatRules(rules) { this.cfRules = rules.slice(); }
  getConditionalFormatRules() { return this.cfRules.slice(); }
  clearConditionalFormatRules() { this.cfRules = []; }
  _key(r, c) { return r + ',' + c; }
  _get(r, c) { const v = this.cells.get(this._key(r, c)); return v === undefined ? '' : v; }
  _set(r, c, v) { if (v === '' || v === null || v === undefined) this.cells.delete(this._key(r, c)); else this.cells.set(this._key(r, c), v); }
  getName() { return this.name; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  insertRowsAfter(after, n) { if (after !== this.maxRows) throw new Error('Mock only supports appending rows at the end'); this.maxRows += n; return this; }
  getLastRow() { let m = 0; for (const k of this.cells.keys()) m = Math.max(m, Number(k.split(',')[0])); return m; }
  getLastColumn() { let m = 0; for (const k of this.cells.keys()) m = Math.max(m, Number(k.split(',')[1])); return m; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { const [r, col, nr, nc] = parseA1(a); return new FakeRange(this, r, col, nr, nc); }
    return new FakeRange(this, a, b, c || 1, d || 1);
  }
  getDataRange() { return new FakeRange(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  appendRow(row) {
    const r = this.getLastRow() + 1;
    if (r > this.maxRows) this.maxRows = r;
    row.forEach((v, i) => this._set(r, i + 1, v));
    return this;
  }
  clear() { this.cells.clear(); return this; }  // like Sheets: merges, charts, rules and validations survive clear()
  clearFormats() { return this; }
  getFilter() { return this.filter; }
  setFrozenRows(n) { this.frozenRows = n; }
  setHiddenGridlines() { return this; }
  getColumnWidth(c) { return this.widths[c] || 100; }
  setColumnWidth(c, w) { this.widths[c] = w; return this; }
  isColumnHiddenByUser(c) { return this.hidden.indexOf(c) !== -1; }
  hideSheet() { this.sheetHidden = true; return this; }
  showSheet() { this.sheetHidden = false; return this; }
  isSheetHidden() { return !!this.sheetHidden; }
  /** Rough stand-in for Sheets' fit-to-content: 7 px per character + 8. */
  autoResizeColumns(start, n) {
    for (let c = start; c < start + n; c++) {
      let longest = 0;
      for (const [k, v] of this.cells) if (Number(k.split(',')[1]) === c) longest = Math.max(longest, String(v).length);
      this.widths[c] = Math.max(21, longest * 7 + 8);
    }
    return this;
  }
  getSheetId() { return this.id; }
  setTabColor() { return this; }
  hideColumns(c, n) { for (let i = 0; i < (n || 1); i++) this.hidden.push(c + i); }   // like Sheets: (column, numColumns)
  /** Test helper: rows 2..lastRow as arrays of `numCols` values. */
  _rows(numCols) {
    const last = this.getLastRow();
    return last < 2 ? [] : this.getRange(2, 1, last - 1, numCols).getValues();
  }
}

class FakeSpreadsheet {
  constructor() { this.sheets = []; this.active = null; this.namedRanges = {}; this.toasts = []; this.toastLog = []; }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  insertSheet(n) {
    if (this.getSheetByName(n)) throw new Error('Mock: sheet exists ' + n);
    const s = new FakeSheet(this, n);
    this.sheets.push(s);
    this.active = s;
    return s;
  }
  getSheets() { return this.sheets.slice(); }
  deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); }
  getActiveSheet() { return this.active || this.sheets[0]; }
  // conservative: activating a hidden tab shows it (as clicking it does), so code that reorders must hide it again
  setActiveSheet(s) { if (s && s.sheetHidden) s.sheetHidden = false; this.activations = (this.activations || 0) + 1; this.active = s; return s; }
  moveActiveSheet(pos) {
    const s = this.getActiveSheet();
    this.sheets = this.sheets.filter(x => x !== s);
    this.sheets.splice(pos - 1, 0, s);
  }
  setNamedRange(name, range) {
    if (this.namedRanges[name]) throw new Error('Mock: named range already exists: ' + name + ' (remove it first)');
    this.namedRanges[name] = range;
  }
  getRangeByName(name) { return this.namedRanges[name] || null; }
  removeNamedRange(name) {
    if (!this.namedRanges[name]) throw new Error('Mock: no named range ' + name);
    delete this.namedRanges[name];
  }
  toast(msg, title, timeout) { this.toasts.push(msg); this.toastLog.push({ msg, title, timeout }); }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/TEST_ID/edit'; }
  getName() { return 'Financial_Tracker_Test'; }
}

function makeServices(options) {
  const opts = options || {};
  const ss = new FakeSpreadsheet();
  const gmail = { threads: opts.threads || [], queries: [], markedRead: [], labels: {}, sent: [] };
  const lock = { held: false, refuse: !!opts.lockBusy };
  const ui = { dialogs: [], alerts: [], menus: [] };
  const triggers = [];
  const props = {};
  const http = { responses: {}, requests: [] };   // UrlFetchApp: http.responses[url] = { code, body }
  const validationBuilder = () => {
    const rule = {};
    const b = { requireValueInList: (list, dropdown) => { rule.list = list.slice(); rule.dropdown = dropdown; return b; },
      requireNumberBetween: () => b, requireNumberGreaterThan: () => b,
      setAllowInvalid: v => { rule.allowInvalid = v; return b; }, setHelpText: t => { rule.helpText = t; return b; }, build: () => rule };
    return b;
  };
  const cfBuilder = () => {
    const rule = { ranges: [] };
    const b = {};
    ['whenNumberLessThan', 'whenNumberGreaterThan', 'whenFormulaSatisfied', 'whenTextEqualTo',
     'setGradientMinpoint', 'setGradientMaxpoint', 'setBackground', 'setFontColor', 'setBold', 'setItalic', 'whenTextStartsWith']
      .forEach(m => { b[m] = v => { rule[m] = v; return b; }; });
    b.setRanges = ranges => { rule.ranges = ranges; return b; };
    b.build = () => rule;
    return b;
  };
  const services = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      setActiveSheet: s => ss.setActiveSheet(s),
      getUi: () => {
        if (!opts.ui) throw new Error('Cannot call SpreadsheetApp.getUi() from this context.');
        // Ui.createMenu(caption) → Menu: addItem(caption, functionName), addSeparator(), addSubMenu(menu), addToUi()
        const createMenu = caption => {
          const menu = { caption, items: [] };
          const m = {
            addItem: (c, fn) => {
              if (typeof fn !== 'string' || !fn) throw new Error('Mock: Menu.addItem needs a function name');
              menu.items.push({ caption: c, fn }); return m;
            },
            addSeparator: () => { menu.items.push({ separator: true }); return m; },
            addSubMenu: sub => { menu.items.push({ sub: sub._menu }); return m; },
            addToUi: () => { ui.menus.push(menu); },
            _menu: menu
          };
          return m;
        };
        // Ui.alert(prompt) or Ui.alert(title, prompt, buttons) → the Button pressed (ui.nextButton, YES by default)
        const Button = { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO', CLOSE: 'CLOSE' };
        const ButtonSet = { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL' };
        const isSet = v => Object.values(ButtonSet).indexOf(v) !== -1;
        const alert = (a, b, c) => {                         // alert(prompt[, buttons]) or alert(title, prompt[, buttons])
          const titled = b !== undefined && !isSet(b);
          ui.alerts.push(titled ? a + '\n' + b : a);
          const set = titled ? c : b;
          return set && set !== ButtonSet.OK ? (ui.nextButton || Button.YES) : Button.OK;
        };
        return { showModalDialog: (html, title) => ui.dialogs.push({ html: html.content, title }), alert, createMenu, Button, ButtonSet,
          showSidebar: html => ui.dialogs.push({ html: html.content, title: html.title, sidebar: true }) };
      },
      newDataValidation: validationBuilder,
      newConditionalFormatRule: cfBuilder,
      flush: () => {},
      BorderStyle: { SOLID: 'SOLID', SOLID_MEDIUM: 'SOLID_MEDIUM', SOLID_THICK: 'SOLID_THICK' }
    },
    Charts: { ChartType: { COLUMN: 'COLUMN', PIE: 'PIE', BAR: 'BAR', LINE: 'LINE' } },
    GmailApp: {
      search(query, start, max) {
        gmail.queries.push(query);
        // like Gmail: "from:" clauses restrict the result to threads from those senders
        const senders = (query.match(/from:\s*([^\s)]+)/gi) || []).map(f => f.replace(/from:\s*/i, '').toLowerCase());
        let matches = senders.length === 0 ? gmail.threads : gmail.threads.filter(t => t.getMessages()
          .some(m => senders.some(s => String(m.getFrom()).toLowerCase().includes(s))));
        // "{is:unread -label:X}" (either): still unread, or without label X
        const pending = query.match(/\{is:unread -label:([^}\s]+)\}/);
        if (pending) {
          const labelled = (gmail.labels[pending[1]] || { threads: [] }).threads;
          matches = matches.filter(t => gmail.markedRead.indexOf(t.getId()) === -1 || labelled.indexOf(t.getId()) === -1);
        }
        return matches.slice(start, start + max);
      },
      markThreadsRead(threads) { gmail.markedRead.push(...threads.map(t => t.getId())); },
      getMessageById(id) {   // v1.1.67; like Gmail, an id it doesn't have is an error
        for (const t of gmail.threads) for (const m of t.getMessages()) if (m.getId() === id) return m;
        throw new Error('Mock: Invalid argument: id ' + id);
      },
      getUserLabelByName: n => gmail.labels[n] || null,
      sendEmail(to, subject, body, options) { gmail.sent.push({ to, subject, body, options }); },
      createLabel(n) {
        gmail.labels[n] = { name: n, threads: [], addToThreads(ts) { this.threads.push(...ts.map(t => t.getId())); } };
        return gmail.labels[n];
      }
    },
    HtmlService: {
      createHtmlOutput: content => { const o = { content, setWidth: () => o, setHeight: () => o, setTitle: t => { o.title = t; return o; } }; return o; }
    },
    ScriptApp: {
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: t => { triggers.splice(triggers.indexOf(t), 1); },
      newTrigger: handler => {
        const t = { handler, getHandlerFunction: () => handler };
        const b = { timeBased: () => b, atHour: h => { t.hour = h; return b; }, everyDays: () => b,
                    onMonthDay: d => { t.monthDay = d; return b; },
                    create: () => { triggers.push(t); return t; } };
        return b;
      }
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: k => (k in props ? props[k] : null),
        setProperty: (k, v) => { props[k] = String(v); },
        deleteProperty: k => { delete props[k]; }
      })
    },
    UrlFetchApp: {
      fetchAll(requests) {
        return requests.map(req => {
          if (typeof req !== 'object' || !req.url) throw new Error('Mock: fetchAll takes request objects with a url');
          http.requests.push(req.url);
          const r = http.responses[req.url];
          if (r instanceof Error) throw r;
          return { getResponseCode: () => (r ? r.code : 404), getContentText: () => (r ? r.body : 'Not found') };
        });
      }
    },
    LockService: {
      getScriptLock: () => ({
        tryLock: () => { if (lock.refuse || lock.held) return false; lock.held = true; return true; },
        releaseLock: () => { lock.held = false; }
      })
    }
  };
  return { services, ss, gmail, lock, ui, triggers, props, http };
}

function fakeThread(id, messages) {
  return { getId: () => id, getMessages: () => messages };
}

module.exports = { makeServices, fakeThread };
