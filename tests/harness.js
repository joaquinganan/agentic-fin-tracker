'use strict';
/**
 * Loads the four .gs files into ONE shared V8 context — the same way Apps
 * Script does (every file shares a single global scope) — with minimal
 * stubs for the Apps Script services the pure logic touches. Anything that
 * needs the real Gmail/Sheets UI is out of scope here and is covered by the
 * manual checklist in tests/README.md.
 */
process.env.TZ = 'America/Santo_Domingo'; // must match appsscript.json "timeZone"

const fs = require('fs');
const path = require('path');
const vm = require('vm');

// GS_ROOT: run the suite against another copy of src/. GS_BUNDLE: against the single-file build (npm run test:bundle).
const ROOT = process.env.GS_BUNDLE ? path.dirname(path.resolve(process.env.GS_BUNDLE)) : (process.env.GS_ROOT || path.join(__dirname, '..', 'src'));
const FILES = process.env.GS_BUNDLE ? [path.basename(process.env.GS_BUNDLE)] : require('../tools/bundle').FILES;   // one list of files
const pad = n => String(n).padStart(2, '0');

/** In-memory stand-in for the handful of Spreadsheet calls the pure paths make. */
function makeSpreadsheet(initialSheets) {
  const data = {};
  Object.keys(initialSheets || {}).forEach(n => { data[n] = initialSheets[n].map(r => r.slice()); });
  const sheet = name => ({
    getName: () => name,
    getDataRange: () => ({ getValues: () => data[name].map(r => r.slice()) }),
    getLastRow: () => data[name].length,
    appendRow: row => { data[name].push(row); }
  });
  return {
    getSheetByName: name => (data[name] ? sheet(name) : null),
    insertSheet: name => { data[name] = []; return sheet(name); },
    _data: data
  };
}

function load(options) {
  const opts = options || {};
  const logs = [];
  const spreadsheet = makeSpreadsheet(opts.sheets);
  const sandbox = {
    console,
    Logger: { log: m => logs.push(String(m)) },
    Session: { getScriptTimeZone: () => 'America/Santo_Domingo', getActiveUser: () => ({ getEmail: () => 'new.user@example.com' }) },
    Utilities: {
      sleep: ms => { sandbox.__slept = (sandbox.__slept || 0) + ms; },   // v1.1.56: no real waiting; the total is kept
      formatDate(d, tz, fmt) {
        return fmt.replace('yyyy', d.getFullYear()).replace('MM', pad(d.getMonth() + 1)).replace('dd', pad(d.getDate()));
      },
      parseDate(str, tz, fmt) {
        const m = String(str).match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (!m) throw new Error('Unparseable date: ' + str);
        return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
      }
    },
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet }
  };
  Object.assign(sandbox, opts.services || {}); // e.g. the strict mock in sheets_mock.js
  const ctx = vm.createContext(sandbox);
  for (const f of FILES) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
  }

  return {
    ctx,
    logs,
    spreadsheet,
    /** Top-level const/let (they live in the script scope, not on the global object). */
    get: name => vm.runInContext(name, ctx),
    /** Date created inside the script realm, so `instanceof Date` behaves as in Apps Script. */
    date: (y, m, d, hh, mm) => vm.runInContext(`new Date(${y}, ${m - 1}, ${d}, ${hh || 12}, ${mm || 0})`, ctx),
    /** Plain JSON copy, so assertions don't trip over cross-realm prototypes. */
    plain: x => JSON.parse(JSON.stringify(x, (k, v) => (v instanceof Object && v.constructor && v.constructor.name === 'Set') ? [...v] : v)),
    fakeMessage({ subject, from, body, date, id }) {
      return {
        getSubject: () => subject,
        getFrom: () => from,
        getPlainBody: () => body,
        getDate: () => date,
        getId: () => id || 'msg-test'
      };
    }
  };
}

function fixture(name) {
  return fs.readFileSync(path.join(__dirname, 'fixtures', name + '.txt'), 'utf8');
}

function manifest() {
  return JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'manifest.json'), 'utf8'));
}

module.exports = { load, fixture, manifest };
