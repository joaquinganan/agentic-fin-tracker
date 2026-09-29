'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const zlib = require('zlib');
const fs = require('fs');
const path = require('path');
const { load } = require('./harness');

test('inflate matches zlib: stored, fixed and dynamic blocks; zlib-wrapped and raw; random, repetitive and large', () => {
  const h = load();
  const random = n => { const b = Buffer.alloc(n); let x = 12345; for (let i = 0; i < n; i++) { x = (x * 1103515245 + 12345) >>> 0; b[i] = x >>> 24; } return b; };
  const text = Buffer.from('BT /F1 9 Tf 1 0 0 -1 40 120 Tm [<0A0B> -8 <0C>] TJ ET\n'.repeat(400));
  const cases = [['empty', Buffer.alloc(0)], ['one byte', Buffer.from('x')], ['text', text], ['random 64 KB', random(65536)], ['1 MB mixed', Buffer.concat([random(300000), text, random(700000)])]];
  for (const [name, data] of cases) {
    for (const level of [0, 1, 6, 9]) {
      const same = (packed, how) => assert.ok(Buffer.from(h.ctx.pdfInflate(new Uint8Array(packed))).equals(data), name + ' · level ' + level + ' · ' + how);
      same(zlib.deflateSync(data, { level }), 'zlib');
      same(zlib.deflateRawSync(data, { level }), 'raw');
    }
  }
});

const SYNTH = path.join(__dirname, 'fixtures', 'statements', 'banesco_savings_synthetic.pdf');

test('a FOP-style statement PDF becomes lines: whole words, the Type0 title, both pages', () => {
  const h = load();
  const text = h.ctx.pdfToText(new Uint8Array(fs.readFileSync(SYNTH)));
  const lines = text.split('\n');
  assert.equal(lines[0], 'Tus Finanzas', 'Type0 font through its ToUnicode map');
  assert.ok(lines.includes('Balance mes anterior: 50,000.00 Débitos del mes: 26,017.03'), 'no spaces inside words (TJ spacing, real glyph widths)');
  assert.ok(lines.includes('03/08/2026 Imp. Art. 12 Ley 288-04 2.00 48,998.00'));
  assert.ok(lines.includes('\f'), 'page break');
  assert.equal(lines.filter(l => /^\d{2}\/\d{2}\/\d{4} /.test(l)).length, 12, 'every row on its own line, both pages');
});

test('PDF values: literal strings with escapes, hex strings, names with #, references', () => {
  const h = load();
  const v = h.plain(h.ctx.pdfParseValue('<< /A (a\\(b\\) \\101\\n) /B <48 65 6C6C6F> /C /Name#20X /D 12 0 R /E [1 -2.5 /F] >>', 0).v);
  assert.deepEqual(v, { A: { s: 'a(b) A\n' }, B: { s: 'Hello', hex: true }, C: { n: 'Name X' }, D: { ref: 12 }, E: [1, -2.5, { n: 'F' }] });
});

test('what the reader does not support says so — the statement goes to Unrecognized with it', () => {
  const h = load();
  assert.throws(() => h.ctx.pdfToText(new Uint8Array(Buffer.from('%PDF-1.5\n1 0 obj << /Type /ObjStm /N 1 >> endobj\n'))), /object streams are not supported/);
  assert.throws(() => h.ctx.pdfToText(new Uint8Array(Buffer.from('%PDF-1.4\nnothing\n%%EOF\n'))), /no document catalog/);
});

test('LAFISE\'s "HORARIO TRANSFERENCIAS PAGOS AL INSTANTE" announcement is filtered, not left in Unrecognized (reported)', () => {
  const h = load();
  const body = 'HORARIOS TRANSFERENCIAS PAGOS AL INSTANTE (LBTR) * BANCO LAFISE * Nota informativa: los pagos se aplican en línea de lunes a viernes.';
  const r = h.plain(h.ctx.parseEmailMessage(h.fakeMessage({ subject: 'HORARIO TRANSFERENCIAS PAGOS AL INSTANTE (LBTR)', from: 'PagosAlInstanteMT103@lafise.com',
    body: body, date: h.date(2026, 9, 18, 11), id: 'notice' }), {}, h.ctx.newParseStats()));
  assert.equal(r.status, 'filtered');
});
