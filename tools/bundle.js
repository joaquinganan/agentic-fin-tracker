#!/usr/bin/env node
/**
 * Builds dist/FinancialTracker.gs — every file of src/ in load order, as ONE file to paste as the only code file of the
 * Apps Script project (v1.1.46). One file can't be half-updated, mixed with another version or pasted into the wrong
 * file (v1.1.33's missing menu was src/01_main.gs pasted twice). Apps Script loads a project's files one after another
 * in one global scope, so the concatenation in the same order behaves the same — the whole test suite runs against it
 * (npm run test:bundle).
 *
 *   node tools/bundle.js          write dist/FinancialTracker.gs
 *   node tools/bundle.js --check  exit 1 if dist/ is not what src/ builds (CI)
 */
'use strict';
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const FILES = ['01_main.gs', '02_categorizer.gs', '03_gmailMonitor.gs', '04_sheetsWriter.gs', '05_dailySummary.gs', '06_monthlySummary.gs', '07_investments.gs'];
const OUT = path.join(ROOT, 'dist', 'FinancialTracker.gs');

function build() {
  const version = (fs.readFileSync(path.join(ROOT, 'src', '01_main.gs'), 'utf8').match(/const SCRIPT_VERSION = "([^"]+)"/) || [])[1];
  if (!version) throw new Error('SCRIPT_VERSION not found in src/01_main.gs');
  const header = [
    '/**',
    ' * Financial Tracker v' + version + ' — https://github.com/joaquinganan/agentic-fin-tracker',
    ' *',
    ' * ONE file: in Extensions › Apps Script, this is the only code file of the project.',
    ' * To update: select everything in this file (Ctrl+A), paste the new version, save (Ctrl+S).',
    ' *',
    ' * Generated from src/ by `npm run bundle` — edit src/, not this file.',
    ' */',
    ''
  ].join('\n');
  return header + FILES.map(f => {
    const code = fs.readFileSync(path.join(ROOT, 'src', f), 'utf8').replace(/\s+$/, '');
    return '\n// ' + '='.repeat(100) + '\n// ' + f + '\n// ' + '='.repeat(100) + '\n\n' + code + '\n';
  }).join('');
}

if (require.main === module) {
  const built = build();
  if (process.argv.includes('--check')) {
    const current = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
    if (current !== built) {
      console.error('dist/FinancialTracker.gs is not what src/ builds — run `npm run bundle` and commit it.');
      process.exit(1);
    }
    console.log('dist/FinancialTracker.gs is up to date.');
  } else {
    fs.mkdirSync(path.dirname(OUT), { recursive: true });
    fs.writeFileSync(OUT, built);
    console.log('Wrote ' + path.relative(ROOT, OUT) + ' (' + Math.round(built.length / 1024) + ' KB, ' + built.split('\n').length + ' lines)');
  }
}
module.exports = { build, FILES, OUT };
