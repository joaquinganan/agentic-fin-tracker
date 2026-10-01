/**
 * PDF TEXT — v1.1.52
 *
 * Reads the text of a bank statement PDF inside Apps Script, with no Drive API: v1.1.51 converted the PDF to a
 * Google Doc through Drive, and on real statements that conversion returned text with no totals and no rows (every
 * statement ended in Unrecognized). This reads the PDF itself: finds each page's content, inflates it (FlateDecode),
 * follows the text operators with their positions and rebuilds the lines — the way `pdftotext` does.
 *
 * Scope: what bank statements use — PDF 1.4-style files (objects not packed in object streams), FlateDecode content,
 * simple fonts (WinAnsi, /Differences with glyph names) and Type0 fonts with a ToUnicode map. Anything else throws, and
 * the statement goes to Unrecognized with the reason.
 */

/* ---------- inflate (RFC 1951), for FlateDecode streams ---------- */
const PDF_LENGTH_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const PDF_LENGTH_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const PDF_DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const PDF_DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const PDF_CLEN_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

/** zlib data (2-byte header) or raw deflate → bytes. Pure. */
function pdfInflate(src) {
  let pos = 0, bit = 0, bitCount = 0;
  if (src.length > 2 && (src[0] & 0x0f) === 8 && ((src[0] << 8) | src[1]) % 31 === 0) pos = 2;   // zlib header
  let out = new Uint8Array(Math.max(1024, src.length * 4)), outLen = 0;
  const put = b => {
    if (outLen === out.length) { const bigger = new Uint8Array(out.length * 2); bigger.set(out); out = bigger; }
    out[outLen++] = b;
  };
  const need = n => { while (bitCount < n) { if (pos >= src.length) throw new Error('PDF: compressed data ends early'); bit |= src[pos++] << bitCount; bitCount += 8; } };
  const bits = n => { if (!n) return 0; need(n); const v = bit & ((1 << n) - 1); bit >>>= n; bitCount -= n; return v; };
  // a Huffman table: counts per length and symbols in canonical order
  const build = lengths => {
    const counts = new Uint16Array(16), offs = new Uint16Array(16), symbols = new Uint16Array(lengths.length);
    lengths.forEach(l => counts[l]++);
    counts[0] = 0;
    for (let i = 1; i < 16; i++) offs[i] = offs[i - 1] + counts[i - 1];
    lengths.forEach((l, s) => { if (l) symbols[offs[l]++] = s; });
    return { counts: counts, symbols: symbols };
  };
  const decode = t => {
    let code = 0, first = 0, index = 0;
    for (let len = 1; len < 16; len++) {
      code |= bits(1);
      const count = t.counts[len];
      if (code - count < first) return t.symbols[index + (code - first)];
      index += count; first += count; first <<= 1; code <<= 1;
    }
    throw new Error('PDF: bad compressed data');
  };
  let fixedLit = null, fixedDist = null;
  let final = 0;
  while (!final) {
    final = bits(1);
    const type = bits(2);
    if (type === 0) {                                   // stored
      bit = 0; bitCount = 0;
      const len = src[pos] | (src[pos + 1] << 8); pos += 4;
      for (let i = 0; i < len; i++) put(src[pos++]);
      continue;
    }
    let lit, dist;
    if (type === 1) {                                   // fixed Huffman
      if (!fixedLit) {
        const l = []; for (let i = 0; i < 288; i++) l.push(i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8);
        fixedLit = build(l); fixedDist = build(new Array(30).fill(5));
      }
      lit = fixedLit; dist = fixedDist;
    } else if (type === 2) {                            // dynamic Huffman
      const hlit = bits(5) + 257, hdist = bits(5) + 1, hclen = bits(4) + 4;
      const clen = new Array(19).fill(0);
      for (let i = 0; i < hclen; i++) clen[PDF_CLEN_ORDER[i]] = bits(3);
      const ct = build(clen), lengths = [];
      while (lengths.length < hlit + hdist) {
        const sym = decode(ct);
        if (sym < 16) lengths.push(sym);
        else if (sym === 16) { const prev = lengths[lengths.length - 1], n = bits(2) + 3; for (let i = 0; i < n; i++) lengths.push(prev); }
        else if (sym === 17) { const n = bits(3) + 3; for (let i = 0; i < n; i++) lengths.push(0); }
        else { const n = bits(7) + 11; for (let i = 0; i < n; i++) lengths.push(0); }
      }
      lit = build(lengths.slice(0, hlit)); dist = build(lengths.slice(hlit));
    } else {
      throw new Error('PDF: bad compressed block');
    }
    for (;;) {
      const sym = decode(lit);
      if (sym === 256) break;
      if (sym < 256) { put(sym); continue; }
      const li = sym - 257, len = PDF_LENGTH_BASE[li] + bits(PDF_LENGTH_EXTRA[li]);
      const di = decode(dist), d = PDF_DIST_BASE[di] + bits(PDF_DIST_EXTRA[di]);
      for (let i = 0; i < len; i++) put(out[outLen - d]);
    }
  }
  return out.slice(0, outLen);
}

/* ---------- PDF objects ---------- */
function pdfLatin1(bytes, from, to) {
  let s = '';
  for (let i = from; i < to; i += 8192) s += String.fromCharCode.apply(null, Array.prototype.slice.call(bytes, i, Math.min(to, i + 8192)));
  return s;
}

/** One PDF value from text at pos → { v, pos }. Names are {n}, refs {ref}, strings {s} (bytes as a latin1 string). */
function pdfParseValue(t, pos) {
  const ws = () => { for (;;) { while (pos < t.length && /[\s\0]/.test(t[pos])) pos++; if (t[pos] === '%') { while (pos < t.length && t[pos] !== '\n' && t[pos] !== '\r') pos++; } else break; } };
  ws();
  const c = t[pos];
  if (c === '<' && t[pos + 1] === '<') {
    pos += 2; const d = {};
    for (;;) {
      ws();
      if (t[pos] === '>' && t[pos + 1] === '>') { pos += 2; return { v: d, pos: pos }; }
      const k = pdfParseValue(t, pos); pos = k.pos;
      const v = pdfParseValue(t, pos); pos = v.pos;
      d[k.v.n] = v.v;
    }
  }
  if (c === '[') {
    pos++; const a = [];
    for (;;) { ws(); if (t[pos] === ']') return { v: a, pos: pos + 1 }; const x = pdfParseValue(t, pos); a.push(x.v); pos = x.pos; }
  }
  if (c === '/') {
    let e = pos + 1; while (e < t.length && !/[\s\/\[\]<>()\{\}%]/.test(t[e])) e++;
    return { v: { n: t.slice(pos + 1, e).replace(/#([0-9A-Fa-f]{2})/g, (m, h) => String.fromCharCode(parseInt(h, 16))) }, pos: e };
  }
  if (c === '(') {
    let depth = 1, s = '', i = pos + 1;
    while (i < t.length && depth) {
      const ch = t[i];
      if (ch === '\\') {
        const n = t[i + 1];
        const map = { n: '\n', r: '\r', t: '\t', b: '\b', f: '\f', '(': '(', ')': ')', '\\': '\\' };
        if (map[n] !== undefined) { s += map[n]; i += 2; }
        else if (/[0-7]/.test(n)) { let o = ''; i++; while (o.length < 3 && /[0-7]/.test(t[i])) o += t[i++]; s += String.fromCharCode(parseInt(o, 8) & 255); }
        else if (n === '\r' || n === '\n') { i += 2; if (n === '\r' && t[i] === '\n') i++; }
        else { s += n; i += 2; }
        continue;
      }
      if (ch === '(') depth++;
      if (ch === ')') { depth--; if (!depth) { i++; break; } }
      s += ch; i++;
    }
    return { v: { s: s }, pos: i };
  }
  if (c === '<') {
    const e = t.indexOf('>', pos);
    let h = t.slice(pos + 1, e).replace(/\s+/g, ''); if (h.length % 2) h += '0';
    let s = ''; for (let i = 0; i < h.length; i += 2) s += String.fromCharCode(parseInt(h.substr(i, 2), 16));
    return { v: { s: s, hex: true }, pos: e + 1 };
  }
  const m = /^([+-]?(?:\d+\.?\d*|\.\d+))(?:\s+(\d+)\s+R\b)?/.exec(t.slice(pos, pos + 40));
  if (m) {
    if (m[2] !== undefined && /^\d+$/.test(m[1])) return { v: { ref: Number(m[1]) }, pos: pos + m[0].length };
    return { v: Number(m[1]), pos: pos + m[1].length };
  }
  const w = /^[A-Za-z*'"]+/.exec(t.slice(pos, pos + 20));
  if (w) return { v: w[0] === 'true' ? true : w[0] === 'false' ? false : w[0] === 'null' ? null : { op: w[0] }, pos: pos + w[0].length };
  throw new Error('PDF: cannot read a value at ' + pos);
}

/** Every "n g obj … endobj" of the file → { n: { dict, stream: bytes|null } }. Streams are inflated when FlateDecode. */
function pdfObjects(bytes) {
  const text = pdfLatin1(bytes, 0, bytes.length);
  // v1.1.59: a password-protected statement says so, instead of failing on unreadable content
  if (/\/Encrypt\b/.test(text)) throw new Error('the PDF is password-protected — not supported yet (send one to get it added)');
  if (/\/ObjStm\b/.test(text)) throw new Error('PDF: objects packed in object streams are not supported');
  const objs = {};
  const re = /(\d+)\s+\d+\s+obj\b/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    let parsed;
    try { parsed = pdfParseValue(text, m.index + m[0].length); } catch (error) { continue; }
    const o = { dict: parsed.v, stream: null };
    const after = text.slice(parsed.pos, parsed.pos + 20);
    const sm = /^\s*stream\r?\n/.exec(after);
    if (sm && parsed.v && typeof parsed.v === 'object') {
      const start = parsed.pos + sm[0].length;
      let len = parsed.v.Length;
      if (len && len.ref !== undefined) len = null;   // indirect: found after all objects are read
      o.streamStart = start; o.streamLen = typeof len === 'number' ? len : null;
      re.lastIndex = typeof len === 'number' ? start + len : text.indexOf('endstream', start);
    }
    objs[m[1]] = o;
  }
  Object.keys(objs).forEach(k => {
    const o = objs[k];
    if (o.streamStart === undefined) return;
    let len = o.streamLen;
    if (len === null) {
      const l = o.dict.Length && o.dict.Length.ref !== undefined ? objs[o.dict.Length.ref] : null;
      len = l && typeof l.dict === 'number' ? l.dict : text.indexOf('endstream', o.streamStart) - o.streamStart;
    }
    let data = bytes.slice(o.streamStart, o.streamStart + len);
    const filters = [].concat(o.dict.Filter || []).map(f => f.n);
    if (filters.length && filters.some(f => f !== 'FlateDecode')) { o.unsupported = filters.join(','); return; }
    if (filters.length) data = pdfInflate(data);
    o.stream = data;
  });
  return objs;
}

/* ---------- text: fonts, positions, lines ---------- */
const PDF_WINANSI_HIGH = { 128: 8364, 130: 8218, 131: 402, 132: 8222, 133: 8230, 134: 8224, 135: 8225, 136: 710, 137: 8240, 138: 352, 139: 8249,
  140: 338, 142: 381, 145: 8216, 146: 8217, 147: 8220, 148: 8221, 149: 8226, 150: 8211, 151: 8212, 152: 732, 153: 8482, 154: 353, 155: 8250,
  156: 339, 158: 382, 159: 376 };
const PDF_GLYPHS = { space: ' ', nbspace: ' ', exclam: '!', quotedbl: '"', numbersign: '#', dollar: '$', percent: '%', ampersand: '&',
  quotesingle: "'", quoteright: '\u2019', parenleft: '(', parenright: ')', asterisk: '*', plus: '+', comma: ',', hyphen: '-', minus: '-', period: '.',
  slash: '/', zero: '0', one: '1', two: '2', three: '3', four: '4', five: '5', six: '6', seven: '7', eight: '8', nine: '9', colon: ':',
  semicolon: ';', less: '<', equal: '=', greater: '>', question: '?', at: '@', bracketleft: '[', backslash: '\\', bracketright: ']',
  underscore: '_', grave: '`', braceleft: '{', bar: '|', braceright: '}', degree: '°', ordfeminine: 'ª', ordmasculine: 'º',
  endash: '–', emdash: '—', bullet: '•', quotedblleft: '“', quotedblright: '”', quoteleft: '‘', exclamdown: '¡', questiondown: '¿',
  aacute: 'á', eacute: 'é', iacute: 'í', oacute: 'ó', uacute: 'ú', ntilde: 'ñ', udieresis: 'ü', Aacute: 'Á', Eacute: 'É', Iacute: 'Í',
  Oacute: 'Ó', Uacute: 'Ú', Ntilde: 'Ñ', Udieresis: 'Ü', agrave: 'à', egrave: 'è', ccedilla: 'ç', Ccedilla: 'Ç' };

function pdfGlyphChar(name) {
  if (name.length === 1) return name;
  if (PDF_GLYPHS[name] !== undefined) return PDF_GLYPHS[name];
  const u = /^uni([0-9A-Fa-f]{4})$/.exec(name);
  if (u) { const c = String.fromCharCode(parseInt(u[1], 16)); return c === '\u00a0' ? ' ' : c; }
  return '';
}

/** A font → { decode(bytes) → text, codes(bytes) → [code…], width(code) → glyph width in 1/1000 em }. */
function pdfFont(font, get) {
  const val = x => (x && x.ref !== undefined ? get(x).dict : x);
  const type0 = font.Subtype && font.Subtype.n === 'Type0';
  const widths = {};
  let defaultWidth = 500;
  if (type0) {
    const desc = val((val(font.DescendantFonts) || [])[0]) || {};
    defaultWidth = typeof desc.DW === 'number' ? desc.DW : 1000;
    const w = val(desc.W) || [];
    for (let i = 0; i < w.length;) {
      const first = w[i], next = val(w[i + 1]);
      if (Array.isArray(next)) { next.forEach((x, k) => { widths[first + k] = x; }); i += 2; }
      else { for (let c = first; c <= next; c++) widths[c] = w[i + 2]; i += 3; }
    }
  } else {
    const list = val(font.Widths), first = typeof font.FirstChar === 'number' ? font.FirstChar : 0;
    if (Array.isArray(list)) list.forEach((x, k) => { widths[first + k] = val(x); });
    const d = val(font.FontDescriptor);
    if (d && typeof d.MissingWidth === 'number') defaultWidth = d.MissingWidth;
  }
  const decode = pdfFontDecoder(font, get);
  const codes = s => { const out = []; if (type0) { for (let i = 0; i + 1 < s.length; i += 2) out.push((s.charCodeAt(i) << 8) | s.charCodeAt(i + 1)); }
    else { for (let i = 0; i < s.length; i++) out.push(s.charCodeAt(i)); } return out; };
  return { decode: decode, codes: codes, width: c => (widths[c] !== undefined ? widths[c] : (c === 32 ? 278 : defaultWidth)), type0: type0 };
}

/** A font → decoder of its string bytes. */
function pdfFontDecoder(font, get) {
  const toUni = font.ToUnicode ? get(font.ToUnicode) : null;
  if (toUni && toUni.stream) {
    const cmap = pdfLatin1(toUni.stream, 0, toUni.stream.length), map = {};
    let width = 1;
    const hex = h => parseInt(h, 16);
    const uni = h => { let s = ''; for (let i = 0; i < h.length; i += 4) s += String.fromCharCode(hex(h.substr(i, 4))); return s; };
    (cmap.match(/beginbfchar([\s\S]*?)endbfchar/g) || []).forEach(block => {
      (block.match(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g) || []).forEach(pair => {
        const p = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/.exec(pair); width = p[1].length / 2; map[hex(p[1])] = uni(p[2]);
      });
    });
    (cmap.match(/beginbfrange([\s\S]*?)endbfrange/g) || []).forEach(block => {
      (block.match(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g) || []).forEach(tr => {
        const p = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/.exec(tr); width = p[1].length / 2;
        for (let c = hex(p[1]), k = 0; c <= hex(p[2]); c++, k++) map[c] = String.fromCharCode(hex(p[3]) + k);
      });
    });
    const twoBytes = (font.Subtype && font.Subtype.n === 'Type0') || width === 2;
    return s => {
      let out = '';
      if (twoBytes) { for (let i = 0; i + 1 < s.length; i += 2) out += map[(s.charCodeAt(i) << 8) | s.charCodeAt(i + 1)] || ''; }
      else { for (let i = 0; i < s.length; i++) out += map[s.charCodeAt(i)] !== undefined ? map[s.charCodeAt(i)] : ''; }
      return out;
    };
  }
  let enc = font.Encoding ? (font.Encoding.ref !== undefined ? get(font.Encoding).dict : font.Encoding) : null;
  const diff = {};
  if (enc && enc.Differences) {
    let code = 0;
    enc.Differences.forEach(x => { if (typeof x === 'number') code = x; else if (x && x.n !== undefined) diff[code++] = pdfGlyphChar(x.n); });
  }
  return s => {
    let out = '';
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      if (diff[c] !== undefined) out += diff[c];
      else if (c >= 32 && c < 127) out += String.fromCharCode(c);
      else if (c >= 160) out += c === 160 ? ' ' : String.fromCharCode(c);
      else if (PDF_WINANSI_HIGH[c]) out += String.fromCharCode(PDF_WINANSI_HIGH[c]);
    }
    return out;
  };
}

/** The text pieces of one page's content, with their positions. */
function pdfPagePieces(content, fonts) {
  const t = pdfLatin1(content, 0, content.length);
  const mul = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
    a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
  let ctm = [1, 0, 0, 1, 0, 0], tm = [1, 0, 0, 1, 0, 0], lm = [1, 0, 0, 1, 0, 0], leading = 0, font = null, size = 0;
  let charSpace = 0, wordSpace = 0, hScale = 1;
  const stack = [], pieces = [], ops = [];
  let pos = 0;
  // a string at the current text position: its text, where it starts and how wide it is (device space)
  const show = str => {
    if (!font) return;
    let advance = 0;
    font.codes(str).forEach(c => { advance += (font.width(c) / 1000 * size + charSpace + (c === 32 && !font.type0 ? wordSpace : 0)) * hScale; });
    const m = mul(tm, ctm), devScale = Math.sqrt(m[0] * m[0] + m[1] * m[1]);
    const text = font.decode(str);
    if (text) pieces.push({ x: m[4], y: m[5], text: text, size: Math.abs(size * devScale), width: advance * devScale, order: pieces.length });
    tm = mul([1, 0, 0, 1, advance, 0], tm);
  };
  while (pos < t.length) {
    while (pos < t.length && /[\s\0]/.test(t[pos])) pos++;
    if (pos >= t.length) break;
    if (t[pos] === '%') { while (pos < t.length && t[pos] !== '\n') pos++; continue; }
    if (t.startsWith('BI', pos) && /\s/.test(t[pos + 2] || ' ')) { const e = t.indexOf('EI', pos); pos = e < 0 ? t.length : e + 2; continue; }   // inline image
    let v;
    try { v = pdfParseValue(t, pos); } catch (error) { pos++; continue; }
    pos = v.pos;
    if (!v.v || v.v.op === undefined) { ops.push(v.v); continue; }
    const op = v.v.op, a = ops.splice(0, ops.length);
    if (op === 'q') stack.push(ctm);
    else if (op === 'Q') ctm = stack.pop() || [1, 0, 0, 1, 0, 0];
    else if (op === 'cm' && a.length === 6) ctm = mul(a, ctm);
    else if (op === 'BT') { tm = [1, 0, 0, 1, 0, 0]; lm = tm; }
    else if (op === 'Tf') { font = fonts[a[0] && a[0].n] || null; size = a[1] || size; }
    else if (op === 'Tc') charSpace = a[0] || 0;
    else if (op === 'Tw') wordSpace = a[0] || 0;
    else if (op === 'Tz') hScale = (typeof a[0] === 'number' ? a[0] : 100) / 100;
    else if (op === 'Tm' && a.length === 6) { tm = a.slice(); lm = tm; }
    else if (op === 'Td' || op === 'TD') { if (op === 'TD') leading = -a[1]; lm = mul([1, 0, 0, 1, a[0], a[1]], lm); tm = lm; }
    else if (op === 'TL') leading = a[0];
    else if (op === 'T*') { lm = mul([1, 0, 0, 1, 0, -leading], lm); tm = lm; }
    else if (op === 'Tj') { if (a[0] && a[0].s !== undefined) show(a[0].s); }
    else if (op === "'" || op === '"') {
      if (op === '"') { wordSpace = a[0] || 0; charSpace = a[1] || 0; }
      lm = mul([1, 0, 0, 1, 0, -leading], lm); tm = lm; const s = a[a.length - 1]; if (s && s.s !== undefined) show(s.s);
    }
    else if (op === 'TJ' && Array.isArray(a[0])) {
      a[0].forEach(x => {
        if (x && x.s !== undefined) show(x.s);
        else if (typeof x === 'number') tm = mul([1, 0, 0, 1, -x / 1000 * size * hScale, 0], tm);   // kerning / spacing
      });
    }
  }
  return pieces;
}

/** The PDF's text, page by page, rebuilt into lines (top to bottom, left to right). */
function pdfToText(bytes) {
  const objs = pdfObjects(bytes);
  const get = r => (r && r.ref !== undefined ? objs[r.ref] : null) || { dict: r, stream: null };
  const catalog = Object.keys(objs).map(k => objs[k]).find(o => o.dict && o.dict.Type && o.dict.Type.n === 'Catalog');
  if (!catalog) throw new Error('PDF: no document catalog');
  const pages = [];
  const walk = (node, inherited) => {
    const d = get(node).dict || {};
    const res = d.Resources || inherited;
    if (d.Type && d.Type.n === 'Pages') (d.Kids || []).forEach(k => walk(k, res));
    else pages.push({ dict: d, resources: res });
  };
  walk(catalog.dict.Pages, null);
  const pageTexts = pages.map(p => {
    const res = p.resources && p.resources.ref !== undefined ? get(p.resources).dict : (p.resources || {});
    const fontDict = res.Font && res.Font.ref !== undefined ? get(res.Font).dict : (res.Font || {});
    const fonts = {};
    Object.keys(fontDict || {}).forEach(name => { const f = get(fontDict[name]).dict; if (f) fonts[name] = pdfFont(f, get); });
    const parts = [].concat(p.dict.Contents || []).map(c => get(c));
    parts.forEach(c => { if (c.unsupported) throw new Error('PDF: content compressed with ' + c.unsupported + ' is not supported'); });
    const total = parts.reduce((n, c) => n + (c.stream ? c.stream.length + 1 : 0), 0), content = new Uint8Array(total);
    let at = 0;
    parts.forEach(c => { if (c.stream) { content.set(c.stream, at); at += c.stream.length; content[at++] = 10; } });
    const pieces = pdfPagePieces(content, fonts).sort((a, b) => (b.y - a.y) || (a.x - b.x) || (a.order - b.order));
    const lines = [];
    pieces.forEach(pc => {
      const line = lines.length && Math.abs(lines[lines.length - 1].y - pc.y) <= Math.max(2, pc.size * 0.3) ? lines[lines.length - 1] : null;
      if (line) line.items.push(pc); else lines.push({ y: pc.y, items: [pc] });
    });
    return lines.map(l => l.items.sort((a, b) => (a.x - b.x) || (a.order - b.order)).reduce((s, pc, i, all) => {
      if (!i) return pc.text;
      const prev = all[i - 1], gap = pc.x - (prev.x + prev.width);   // a word space is ~0.25 em; kerning is far less
      return s + (gap > Math.min(prev.size, pc.size) * 0.12 ? ' ' : '') + pc.text;
    }, '').replace(/\s+/g, ' ').trim()).filter(Boolean).join('\n');
  });
  return pageTexts.join('\n\f\n');
}
