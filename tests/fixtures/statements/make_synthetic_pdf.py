"""Builds banesco_savings_synthetic.pdf: the synthetic statement (banesco_savings_layout.txt, invented data) as a PDF
laid out and encoded the way the real one is (Apache FOP): a Type1 font with a /Differences encoding and /Widths, text
in TJ arrays with spacing inside words, a Type0 title font with a ToUnicode map, Flate-compressed content, two pages.
Run: python3 make_synthetic_pdf.py  (needs reportlab only for Helvetica's glyph widths)."""
import zlib, re
from reportlab.pdfbase.pdfmetrics import stringWidth

GLYPH = {' ': 'space', '-': 'hyphen', ':': 'colon', '.': 'period', ',': 'comma', '/': 'slash', '$': 'dollar', '%': 'percent',
         '°': 'degree', 'é': 'eacute', 'á': 'aacute', 'í': 'iacute', 'ó': 'oacute', 'ú': 'uacute', 'ñ': 'ntilde', 'Ñ': 'Ntilde',
         '0': 'zero', '1': 'one', '2': 'two', '3': 'three', '4': 'four', '5': 'five', '6': 'six', '7': 'seven', '8': 'eight', '9': 'nine'}
lines = [l.rstrip('\n') for l in open('banesco_savings_layout.txt', encoding='utf-8')]
chars = sorted(set(''.join(lines)) - {'\f'})
code = {c: i + 1 for i, c in enumerate(chars)}                      # custom codes, like FOP's subset encoding
names = [GLYPH.get(c, c) for c in chars]
w = lambda c: round(stringWidth(c, 'Helvetica', 1000))

def hexs(text):
    return '<' + ''.join('%02X' % code[c] for c in text) + '>'

def cell(x, y, text, style):
    """One table cell, written the two ways the real file does:
    'td' — 2–3 character pieces, each placed by Td at the real width of the one before (+ small kerning), so a reader
           needs the true glyph widths to see they touch;
    'tj' — words in one TJ with NO space character between them: a -278 (a space's width) moves the next one over, so
           a reader must apply TJ numbers or the words run together."""
    if style == 'tj':
        words = text.split(' ')
        arr = ' -278 '.join(hexs(w) for w in words)
        return 'BT /F1 8 Tf 1 0 0 -1 %.2f %d Tm [%s] TJ ET' % (x, y, arr)
    ops, prev = ['BT /F1 8 Tf 1 0 0 -1 %.2f %d Tm' % (x, y)], None
    for i, p in enumerate(re.findall(r'.{1,3}', text)):
        if prev is not None:
            kern = [0.0, 0.2, -0.15, 0.1][i % 4]                     # points, well under a space (~2.2 pt at 8 pt)
            ops.append('%.3f 0 Td' % (sum(w(c) for c in prev) * 8 / 1000 + kern))
        ops.append(hexs(p) + ' Tj'); prev = p
    ops.append('ET')
    return ' '.join(ops)

def page(chunk, title):
    ops, y = ['q', '1 0 0 -1 0 842 cm'], 40                            # y grows downward, as in the real file
    if title:                                                          # the title: one glyph at a time in a Type0 font
        ops.append('BT 1 0 0 -1 40 %d Tm /F2 18 Tf' % y)
        for i, c in enumerate('Tus Finanzas'):
            ops.append(('0 0 Td' if i == 0 else '%.3f 0 Td' % (w('Tus Finanzas'[i - 1]) * 18 / 1000)) + ' <%04X> Tj' % (i + 1))
        ops.append('ET'); y += 26
    for l in chunk:
        if not l.strip(): continue
        # columns: split where the layout has 3+ spaces, each at its own x (like a table cell)
        x = 40
        style = 'tj' if re.search(r'Balance|Débitos|Créditos|Impuestos|Fecha|Transacción', l) else 'td'
        for m in re.finditer(r'\S(?:.*?\S)?(?=\s{3,}|$)', l):
            ops.append(cell(40 + m.start() * 4.1, y, m.group(0), style))
        y += 12
    ops.append('Q')
    return zlib.compress(('\n'.join(ops)).encode('latin-1'))

split = lines.index('\f')
body = [l for l in lines[1:split]], [l for l in lines[split + 1:]]
objs = {}
objs[1] = b'<< /Type /Catalog /Pages 2 0 R >>'
objs[2] = b'<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 /Resources << /Font << /F1 7 0 R /F2 8 0 R >> >> >>'
for n, content in ((3, page(body[0], True)), (4, page(body[1], False))):
    objs[n] = b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 842] /Resources << /Font << /F1 7 0 R /F2 8 0 R >> >> /Contents %d 0 R >>' % (n + 2)
    objs[n + 2] = (b'<< /Length %d /Filter /FlateDecode >>\nstream\n' % len(content)) + content + b'\nendstream'
objs[7] = ('<< /Type /Font /Subtype /Type1 /BaseFont /HelveticaLTStd-Light /FirstChar 1 /LastChar %d /Widths [%s] '
           '/Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding /Differences [1 %s] >> >>'
           % (len(chars), ' '.join(str(w(c)) for c in chars), ' '.join('/' + n for n in names))).encode('latin-1')
title = 'Tus Finanzas'
cmap = ('/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CMapName /T def 1 begincodespacerange <0000> <FFFF> '
        'endcodespacerange %d beginbfchar %s endbfchar endcmap CMapName currentdict /CMap defineresource pop end end'
        % (len(title), ' '.join('<%04X> <%04X>' % (i + 1, ord(c)) for i, c in enumerate(title)))).encode('latin-1')
objs[8] = b'<< /Type /Font /Subtype /Type0 /BaseFont /EAAAAA+Helvetica /Encoding /Identity-H /DescendantFonts [9 0 R] /ToUnicode 10 0 R >>'
objs[9] = ('<< /Type /Font /Subtype /CIDFontType2 /BaseFont /EAAAAA+Helvetica /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> '
           '/FontDescriptor 11 0 R /DW 1000 /W [1 [%s]] >>'
           % ' '.join(str(w(c)) for c in title)).encode('latin-1')
objs[10] = (b'<< /Length %d >>\nstream\n' % len(cmap)) + cmap + b'\nendstream'
objs[11] = b'<< /Type /FontDescriptor /FontName /EAAAAA+Helvetica /Flags 32 /FontBBox [-166 -225 1000 931] /ItalicAngle 0 /Ascent 718 /Descent -207 /CapHeight 718 /StemV 88 >>'
out, offsets = bytearray(b'%PDF-1.4\n%\xe2\xe3\xcf\xd3\n'), {}
for n in sorted(objs):
    offsets[n] = len(out); out += b'%d 0 obj\n' % n + objs[n] + b'\nendobj\n'
xref = len(out)
out += b'xref\n0 %d\n0000000000 65535 f \n' % (max(objs) + 1) + b''.join(b'%010d 00000 n \n' % offsets[n] for n in sorted(objs))
out += b'trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n' % (max(objs) + 1, xref)
open('banesco_savings_synthetic.pdf', 'wb').write(bytes(out))
print('wrote banesco_savings_synthetic.pdf', len(out), 'bytes')
