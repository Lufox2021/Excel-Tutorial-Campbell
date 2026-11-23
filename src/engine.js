/* Excel Workbench — engine: seeded data, zip/xlsx read & write, Lab A checker, codes.
   No libraries. Runs in the browser (uses DOMParser, DecompressionStream, TextEncoder). */

// ---------- hashing, seeded random numbers ----------
function cyrb53(str, seed = 0) {
  let h1 = 0xdeadbeef ^ seed, h2 = 0x41c6ce57 ^ seed;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507);
  h1 ^= Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507);
  h2 ^= Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return 4294967296 * (2097151 & h2) + (h1 >>> 0);
}
function mulberry32(a) {
  return function () {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
function gauss(r) { let u = 0; while (u === 0) u = r(); const v = r(); return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v); }
const normName = s => String(s == null ? '' : s).trim().replace(/\s+/g, ' ').toLowerCase();

// ---------- completion codes ----------
const CODE_SALT = 'CKA|2026|lufox'; // differs from the general Excel Workbench, so its codes don't work here
function completionCode(name, lab) {
  const h = cyrb53(normName(name) + '|' + lab + '|' + CODE_SALT);
  const s = (h % 2176782336).toString(36).toUpperCase().padStart(6, '0'); // 36^6
  return lab + '-' + s;
}
function verifyCode(name, code) {
  const c = String(code || '').trim().toUpperCase();
  const lab = c.split('-')[0];
  if (!['A', 'B', 'C'].includes(lab)) return false;
  return completionCode(name, lab) === c;
}

// ---------- statistics ----------
const mean = a => a.reduce((s, x) => s + x, 0) / a.length;
function sdS(a) { const m = mean(a); return Math.sqrt(a.reduce((s, x) => s + (x - m) ** 2, 0) / (a.length - 1)); }
function sdP(a) { const m = mean(a); return Math.sqrt(a.reduce((s, x) => s + (x - m) ** 2, 0) / a.length); }
function median(a) { const b = [...a].sort((x, y) => x - y), n = b.length; return n % 2 ? b[(n - 1) / 2] : (b[n / 2 - 1] + b[n / 2]) / 2; }

// ---------- Lab A data set ----------
const LAB_A_SAMPLES = ['A', 'B', 'C', 'D', 'E'];
function labAData(name) {
  const r = mulberry32(cyrb53(normName(name) + '|labA') >>> 0);
  const conc = Math.round((0.0950 + r() * 0.0300) * 10000) / 10000;
  const rows = [];
  for (const s of LAB_A_SAMPLES) {
    const base = 18 + r() * 14, spread = 0.03 + r() * 0.12;
    let vals;
    do { vals = [1, 2, 3].map(() => Math.round((base + spread * gauss(r)) * 100) / 100); }
    while (new Set(vals).size < 3);
    vals.forEach((v, i) => rows.push({ sample: s, trial: i + 1, vol: v }));
  }
  // samples A–D are handed out scrambled; E is typed by hand
  const given = rows.filter(x => x.sample !== 'E');
  for (let i = given.length - 1; i > 0; i--) { const j = Math.floor(r() * (i + 1)); [given[i], given[j]] = [given[j], given[i]]; }
  if (given.every((x, i) => i === 0 || given[i - 1].sample <= x.sample)) given.reverse();
  const typed = rows.filter(x => x.sample === 'E');
  const stats = {};
  for (const s of LAB_A_SAMPLES) {
    const v = rows.filter(x => x.sample === s).map(x => x.vol);
    const n = v.map(x => x / 1000 * conc);            // moles of NaOH (column D)
    const m = mean(n), S = sdS(n), P = sdP(n);
    stats[s] = { mean: m, sdS: S, sdP: P, cv: S / m * 100, cvP: P / m * 100, median: median(n), max: Math.max(...n), min: Math.min(...n),
      vMean: mean(v), vSdS: sdS(v), vMedian: median(v), vMax: Math.max(...v), vMin: Math.min(...v) };
  }
  return { name, conc, rows, given, typed, stats };
}
// The block students paste at A1 / get in the .csv/.xlsx (rows 1–15)
function labABlock(d) {
  const out = [['Name', ''], ['NaOH concentration (M)', d.conc.toFixed(4)], ['Sample', 'Trial', 'Volume (mL)']];
  d.given.forEach(x => out.push([x.sample, String(x.trial), x.vol.toFixed(2)]));
  return out;
}

// ---------- cell-address helpers ----------
function colToNum(c) { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
function numToCol(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function splitRef(ref) { const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref); return m ? { col: colToNum(m[1]), row: +m[2] } : null; }
const addr = (col, row) => numToCol(col) + row;

// ---------- zip writer (stored, no compression) ----------
const CRC_TABLE = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(b) { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC_TABLE[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
function zipStore(files) {
  const enc = new TextEncoder(), parts = [], central = [];
  let offset = 0;
  for (const f of files) {
    const data = enc.encode(f.text), nm = enc.encode(f.name), crc = crc32(data);
    const lh = new Uint8Array(30 + nm.length), dv = new DataView(lh.buffer);
    dv.setUint32(0, 0x04034b50, true); dv.setUint16(4, 20, true); dv.setUint16(8, 0, true);
    dv.setUint16(12, 0x0021, true); dv.setUint32(14, crc, true); dv.setUint32(18, data.length, true);
    dv.setUint32(22, data.length, true); dv.setUint16(26, nm.length, true); lh.set(nm, 30);
    const ch = new Uint8Array(46 + nm.length), cv = new DataView(ch.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true);
    cv.setUint16(14, 0x0021, true); cv.setUint32(16, crc, true); cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true); cv.setUint16(28, nm.length, true); cv.setUint32(42, offset, true); ch.set(nm, 46);
    parts.push(lh, data); central.push(ch); offset += lh.length + data.length;
  }
  const cdSize = central.reduce((s, c) => s + c.length, 0);
  const end = new Uint8Array(22), ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, files.length, true); ev.setUint16(10, files.length, true);
  ev.setUint32(12, cdSize, true); ev.setUint32(16, offset, true);
  return new Blob([...parts, ...central, end], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
const xmlEsc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
// rows: array of arrays of strings; numeric-looking strings become numbers
function makeXlsx(rows, sheetName) {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  let sd = '';
  rows.forEach((row, i) => {
    sd += `<row r="${i + 1}">`;
    row.forEach((v, j) => {
      if (v === '' || v == null) return;
      const ref = addr(j + 1, i + 1);
      sd += /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(v) ? `<c r="${ref}"><v>${v}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t>${xmlEsc(v)}</t></is></c>`;
    });
    sd += '</row>';
  });
  const hdr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  return zipStore([
    { name: '[Content_Types].xml', text: hdr + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' },
    { name: '_rels/.rels', text: hdr + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
    { name: 'xl/workbook.xml', text: hdr + `<workbook xmlns="${NS}" xmlns:r="${RNS}"><sheets><sheet name="${xmlEsc(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', text: hdr + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>' },
    { name: 'xl/worksheets/sheet1.xml', text: hdr + `<worksheet xmlns="${NS}"><sheetData>${sd}</sheetData></worksheet>` }
  ]);
}
function makeCsv(rows) { return rows.map(r => r.map(v => /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v).join(',')).join('\r\n') + '\r\n'; }

// ---------- zip / xlsx reader ----------
function readZip(buf) {
  const dv = new DataView(buf), dec = new TextDecoder();
  let eocd = -1;
  for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('NOTZIP');
  const n = dv.getUint16(eocd + 10, true); let p = dv.getUint32(eocd + 16, true);
  const files = {};
  for (let k = 0; k < n; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('NOTZIP');
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true);
    const name = dec.decode(new Uint8Array(buf, p + 46, nl));
    const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
    files[name] = { method, data: new Uint8Array(buf, start, csize) };
    p += 46 + nl + el + cl;
  }
  return files;
}
async function unzipEntry(f) {
  if (f.method === 0) return f.data;
  if (f.method !== 8) throw new Error('ZIPMETHOD');
  const stream = new Blob([f.data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
const BUILTIN_FMT = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??', 48: '##0.0E+0', 49: '@' };
function fmtInfo(code) {
  code = code || 'General';
  if (/^general$/i.test(code)) return { code, general: true, dec: null, sci: false, pct: false };
  let c = code.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '');
  c = c.split(';')[0];
  const m = /\.([0#?]+)/.exec(c);
  return { code, general: false, dec: m ? m[1].length : 0, sci: /E[+-]/i.test(c), pct: c.includes('%') };
}
function resolvePath(fromFile, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = fromFile.split('/'); parts.pop();
  for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.' && seg !== '') parts.push(seg); }
  return parts.join('/');
}
// Read the parts of a DrawingML chart the labs care about.
function parseChart(doc) {
  const kids = (e, n) => e ? Array.from(e.children).filter(x => x.localName === n) : [];
  const kid = (e, n) => kids(e, n)[0] || null;
  const all = (e, n) => e ? Array.from(e.getElementsByTagNameNS('*', n)) : [];
  const val = e => e ? e.getAttribute('val') : null;
  const yes = e => !!e && (val(e) === '1' || val(e) === 'true' || val(e) === null);
  const space = doc.documentElement, chart = kid(space, 'chart'), plot = kid(chart, 'plotArea');
  const firstSz = el => { const d = all(el, 'defRPr').map(e => e.getAttribute('sz')).filter(Boolean); return d.length ? +d[0] : null; };
  const spaceSz = firstSz(kid(space, 'txPr'));
  // font sizes (in 1/100 pt) that apply to the text of an element
  function sizes(el) {
    if (!el) return [];
    let s = all(el, 'rPr').map(e => e.getAttribute('sz')).filter(Boolean);
    if (!s.length) s = all(el, 'defRPr').map(e => e.getAttribute('sz')).filter(Boolean);
    if (!s.length && spaceSz) s = [spaceSz];
    return s.map(Number);
  }
  function textOf(t) {
    const tx = kid(t, 'tx'); if (!tx) return null;
    const rich = kid(tx, 'rich');
    if (rich) return all(rich, 'p').map(p => all(p, 't').map(x => x.textContent).join('')).join(' ').replace(/\s+/g, ' ').trim();
    const sr = kid(tx, 'strRef'); if (sr) return all(sr, 'v').map(x => x.textContent).join(' ').trim();
    return null;
  }
  const refOf = e => { if (!e) return null; const r = kid(e, 'numRef') || kid(e, 'strRef'); const f = r && kid(r, 'f'); return f ? { f: f.textContent, str: r.localName === 'strRef' } : null; };
  const groups = plot ? Array.from(plot.children).filter(x => /Chart$/.test(x.localName)) : [];
  const axisById = {};
  if (plot) for (const ax of Array.from(plot.children).filter(x => /Ax$/.test(x.localName))) axisById[val(kid(ax, 'axId'))] = ax;
  const series = [];
  for (const g of groups) for (const sEl of kids(g, 'ser')) {
    const tx = kid(sEl, 'tx'), sr = tx && kid(tx, 'strRef');
    const spPr = kid(sEl, 'spPr'), ln = kid(spPr, 'ln'), marker = kid(sEl, 'marker'), tl = kid(sEl, 'trendline');
    const style = val(kid(g, 'scatterStyle'));
    // series colour: marker fill, else line/fill colour; null = Excel's automatic colour for this series index
    const colorOf = sp => {
      if (!sp) return null;
      const f = kid(sp, 'solidFill') || (kid(sp, 'ln') && kid(kid(sp, 'ln'), 'solidFill'));
      const c = f && Array.from(f.children)[0];
      if (!c) return null;
      return c.localName + ':' + (c.getAttribute('val') || c.getAttribute('lastClr') || '') + Array.from(c.children).map(m => '|' + m.localName + '=' + m.getAttribute('val')).join('');
    };
    series.push({
      type: g.localName, idx: val(kid(sEl, 'idx')),
      color: colorOf(kid(marker, 'spPr')) || colorOf(spPr),
      name: tx ? (all(tx, 'v').map(x => x.textContent).join('') || null) : null,
      nameRef: sr && kid(sr, 'f') ? kid(sr, 'f').textContent : null,
      x: refOf(kid(sEl, 'xVal') || kid(sEl, 'cat')), y: refOf(kid(sEl, 'yVal') || kid(sEl, 'val')),
      noLine: ln ? !!kid(ln, 'noFill') : style === 'marker' || style === 'none',
      noMarker: !!marker && val(kid(marker, 'symbol')) === 'none',
      trend: tl ? { type: val(kid(tl, 'trendlineType')), eq: yes(kid(tl, 'dispEq')) && !!kid(tl, 'dispEq'), r2: yes(kid(tl, 'dispRSqr')) && !!kid(tl, 'dispRSqr'), forced: !!kid(tl, 'intercept') } : null
    });
  }
  const axis = id => {
    const ax = axisById[id]; if (!ax) return null;
    const t = kid(ax, 'title'), sc = kid(ax, 'scaling');
    return { kind: ax.localName, deleted: val(kid(ax, 'delete')) === '1', title: t ? { text: textOf(t), sizes: sizes(t) } : null,
      min: sc && kid(sc, 'min') ? +val(kid(sc, 'min')) : null, max: sc && kid(sc, 'max') ? +val(kid(sc, 'max')) : null };
  };
  // chart title split into lines (paragraphs or line breaks), each with the font sizes of its text (null = Excel default)
  function titleLines(t) {
    const tx = kid(t, 'tx'), rich = tx && kid(tx, 'rich'); if (!rich) return [];
    const tSz = firstSz(kid(t, 'txPr')) || firstSz(kid(rich, 'lstStyle')) || spaceSz;
    const out = [];
    for (const p of kids(rich, 'p')) {
      const pPr = kid(p, 'pPr'), d = pPr && kid(pPr, 'defRPr'), pSz = d && d.getAttribute('sz') ? +d.getAttribute('sz') : tSz;
      let cur = { text: '', sizes: [] }; out.push(cur);
      for (const ch of Array.from(p.children)) {
        if (ch.localName === 'r' || ch.localName === 'fld') {
          const rPr = kid(ch, 'rPr'), tt = kid(ch, 't'), sz = rPr && rPr.getAttribute('sz') ? +rPr.getAttribute('sz') : (pSz || null);
          (tt ? tt.textContent : '').split(/\r?\n/).forEach((txt, k) => {
            if (k > 0) { cur = { text: '', sizes: [] }; out.push(cur); }   // a newline inside the text also starts a new line
            if (txt.trim()) cur.sizes.push(sz);
            cur.text += txt;
          });
        } else if (ch.localName === 'br') { cur = { text: '', sizes: [] }; out.push(cur); }
      }
    }
    return out.map(l => ({ text: l.text.replace(/\s+/g, ' ').trim(), sizes: l.sizes })).filter(l => l.text);
  }
  const g0 = groups[0], ids = g0 ? kids(g0, 'axId').map(val) : [];
  const tEl = kid(chart, 'title'), lEl = kid(chart, 'legend');
  return {
    types: groups.map(g => g.localName), series,
    title: tEl ? { text: textOf(tEl), sizes: sizes(tEl), lines: titleLines(tEl) } : null,
    x: axis(ids[0]), y: axis(ids[1]),
    legend: lEl ? { sizes: sizes(lEl) } : null
  };
}
async function readXlsx(buf) {
  const files = readZip(buf);
  const keys = Object.keys(files);
  const get = async path => {
    const k = keys.find(x => x.toLowerCase() === path.toLowerCase());
    return k ? new TextDecoder().decode(await unzipEntry(files[k])) : null;
  };
  const parse = s => new DOMParser().parseFromString(s, 'application/xml');
  const els = (d, n) => Array.from(d.getElementsByTagNameNS('*', n));
  const kid = (e, n) => Array.from(e.children).find(x => x.localName === n);
  const wbText = await get('xl/workbook.xml');
  if (!wbText) throw new Error('NOTXLSX');
  const wb = parse(wbText);
  const rels = {};
  const relText = await get('xl/_rels/workbook.xml.rels');
  if (relText) els(parse(relText), 'Relationship').forEach(r => rels[r.getAttribute('Id')] = r.getAttribute('Target'));
  const sst = [];
  const sstText = await get('xl/sharedStrings.xml');
  if (sstText) els(parse(sstText), 'si').forEach(si => sst.push(els(si, 't').filter(t => t.parentNode.localName !== 'rPh').map(t => t.textContent).join('')));
  const xfFmt = [], fmts = { ...BUILTIN_FMT };
  const stText = await get('xl/styles.xml');
  if (stText) {
    const st = parse(stText);
    els(st, 'numFmt').forEach(f => fmts[+f.getAttribute('numFmtId')] = f.getAttribute('formatCode'));
    const cx = els(st, 'cellXfs')[0];
    if (cx) Array.from(cx.children).filter(x => x.localName === 'xf').forEach(x => xfFmt.push(+(x.getAttribute('numFmtId') || 0)));
  }
  const sheets = [];
  for (const sh of els(wb, 'sheet')) {
    const rid = sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || sh.getAttribute('r:id');
    let t = rels[rid]; if (!t) continue;
    t = t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '');
    const txt = await get(t); if (!txt) continue;
    const doc = parse(txt), cells = {};
    for (const c of els(doc, 'c')) {
      const ref = c.getAttribute('r'); if (!ref) continue;
      const type = c.getAttribute('t') || 'n', s = +(c.getAttribute('s') || 0);
      const fEl = kid(c, 'f'), vEl = kid(c, 'v'), isEl = kid(c, 'is');
      let v = null;
      if (type === 's' && vEl) v = sst[+vEl.textContent] ?? '';
      else if (type === 'inlineStr' && isEl) v = els(isEl, 't').map(x => x.textContent).join('');
      else if (type === 'str' || type === 'e') v = vEl ? vEl.textContent : '';
      else if (type === 'b') v = vEl ? vEl.textContent === '1' : null;
      else if (vEl && vEl.textContent !== '') v = Number(vEl.textContent);
      const numFmtId = xfFmt[s] ?? 0;
      cells[ref] = { ref, v, type, formula: !!fEl, ftext: fEl ? fEl.textContent : '', fmt: fmtInfo(fmts[numFmtId]) };
    }
    // charts placed on this sheet: sheet rels → drawing → drawing rels → chart
    const charts = [];
    const relsOf = async path => { const i = path.lastIndexOf('/'); const x = await get(path.slice(0, i) + '/_rels/' + path.slice(i + 1) + '.rels'); return x ? els(parse(x), 'Relationship') : []; };
    for (const rel of await relsOf(t)) {
      if (!/\/drawing$/.test(rel.getAttribute('Type') || '')) continue;
      const dPath = resolvePath(t, rel.getAttribute('Target'));
      for (const cr of await relsOf(dPath)) {
        if (!/\/chart$/.test(cr.getAttribute('Type') || '')) continue;
        const cx = await get(resolvePath(dPath, cr.getAttribute('Target')));
        if (cx) { try { charts.push(parseChart(parse(cx))); } catch (e) { charts.push({ error: true, types: [], series: [] }); } }
      }
    }
    sheets.push({ name: sh.getAttribute('name'), cells, charts });
  }
  return { sheets, fileNames: keys };
}

// ---------- Lab A checker ----------
const near = (v, e) => typeof v === 'number' && Math.abs(v - e) <= 1e-7 * Math.abs(e) + 1e-12;
const fmtList = arr => arr.length <= 1 ? arr.join('') : arr.slice(0, -1).join(', ') + ' and ' + arr[arr.length - 1];

function checkLabA(wbk) {
  const T = [];
  const task = (id, title, module) => { const t = { id, title, module, ok: false, msg: '' }; T.push(t); return t; };
  const text = c => c && c.v != null ? String(c.v).trim() : '';
  // pick the sheet with the data table
  let sheet = null, hr = 0, col = {};
  for (const sh of wbk.sheets) {
    for (let r = 1; r <= 15 && !sheet; r++) {
      const hdr = {};
      for (let c = 1; c <= 12; c++) {
        const t = text(sh.cells[addr(c, r)]).toLowerCase();
        if (t === 'sample' && !hdr.sample) hdr.sample = c;
        else if (t.startsWith('trial') && !hdr.trial) hdr.trial = c;
        else if (t.startsWith('volume') && !hdr.vol) hdr.vol = c;
        else if (/mol/.test(t) && !/volume/.test(t) && !hdr.mol) hdr.mol = c;
      }
      if (hdr.sample && hdr.trial && hdr.vol) { sheet = sh; hr = r; col = hdr; }
    }
    if (sheet) break;
  }
  const tName = task('name', 'Your name is in cell B1', 1);
  if (!sheet) sheet = wbk.sheets[0] || { cells: {} };
  const cells = sheet.cells;
  const name = text(cells.B1);
  if (!name) { tName.msg = 'Cell B1 is empty. Type your name in B1, exactly as you typed it on the Lab page.'; return { tasks: T, name: '', done: false }; }
  if (/^name$/i.test(name)) { tName.msg = 'B1 says "Name". Your name goes in B1, next to the "Name" label in A1.'; return { tasks: T, name: '', done: false }; }
  tName.ok = true; tName.msg = `Checking the data set for "${name}".`;
  const d = labAData(name);

  const tData = task('data', 'All 15 measurements are in the table', 2);
  if (!hr) { tData.msg = 'I couldn\'t find the data table. Row 3 should have the headers Sample, Trial and Volume (mL), as in the data you pasted.'; return { tasks: T, name, done: false }; }
  // read data rows
  const rows = [];
  for (let r = hr + 1; r < hr + 60; r++) {
    const s = text(cells[addr(col.sample, r)]);
    if (!s) break;
    const tc = cells[addr(col.trial, r)], vc = cells[addr(col.vol, r)];
    rows.push({ r, sample: s.toUpperCase(), trial: tc ? Number(tc.v) : NaN, vol: vc ? vc.v : null, volCell: vc });
  }
  const expKey = x => `${x.sample}|${x.trial}|${Number(x.vol).toFixed(2)}`;
  const pool = {}; d.rows.forEach(x => pool[expKey(x)] = (pool[expKey(x)] || 0) + 1);
  let matched = 0; const rowOK = rows.map(x => { const k = expKey(x); if (typeof x.vol === 'number' && pool[k] > 0) { pool[k]--; matched++; return true; } return false; });
  if (matched === 0 && rows.length) {
    tData.msg = `None of these measurements belong to the data set for "${name}". Make sure B1 has exactly the name you typed on the Lab page (the data set is made from that name).`;
    return { tasks: T, name, done: false };
  }
  const pairs = new Set(rows.map(x => x.sample + '|' + x.trial));
  const missing = d.rows.filter(x => !pairs.has(x.sample + '|' + x.trial)).map(x => `${x.sample} trial ${x.trial}`);
  const eRows = rows.filter(x => x.sample === 'E');
  const textVols = rows.filter(x => x.volCell && typeof x.volCell.v === 'string');
  if (textVols.length) tData.msg = `Some volumes are stored as text, not numbers (row ${textVols.map(x => x.r).join(', ')}). Re-type them without units; units belong only in the header.`;
  else if (!eRows.length) tData.msg = 'Sample E is missing. Type its three trials (shown on the Lab page) into the rows just below the pasted data.';
  else if (missing.length) tData.msg = `Missing: ${fmtList(missing)}.`;
  else if (rows.length !== 15) tData.msg = `The table has ${rows.length} rows; it should have 15 (5 samples × 3 trials). Check for duplicates or a blank row.`;
  else if (rows.some((x, i) => !rowOK[i] && x.sample !== 'E')) { tData.ok = true; tData.msg = '15 rows found.'; }
  else {
    const badE = d.typed.filter(x => !rows.some(y => y.sample === 'E' && y.trial === x.trial && near(y.vol, x.vol)));
    if (badE.length) tData.msg = `Check what you typed for ${fmtList(badE.map(x => 'E trial ' + x.trial))}; it doesn't match the Lab page.`;
    else { tData.ok = true; tData.msg = '15 rows found.'; }
  }

  const tTog = task('together', 'Rows stayed together when you sorted', 3);
  const broken = rows.filter((x, i) => !rowOK[i] && x.sample !== 'E');
  if (!tData.ok && !rows.length) tTog.msg = 'No data to check.';
  else if (broken.length) tTog.msg = `${broken.length} row(s) have a volume that doesn't belong with that sample and trial (e.g. row ${broken[0].r}). This happens when one column is sorted by itself. Undo, or paste the data again, then select the whole table before sorting.`;
  else { tTog.ok = true; tTog.msg = 'Every volume is still beside its own sample and trial.'; }

  const tSort = task('sorted', 'Sorted by Sample, then by Trial', 3);
  const inOrder = rows.every((x, i) => i === 0 || rows[i - 1].sample < x.sample || (rows[i - 1].sample === x.sample && rows[i - 1].trial < x.trial));
  if (inOrder) { tSort.ok = true; tSort.msg = 'A1, A2, A3, B1 … E3.'; }
  else {
    const bySample = rows.every((x, i) => i === 0 || rows[i - 1].sample <= x.sample);
    tSort.msg = bySample ? 'Sorted by Sample, but the trials within a sample are out of order. In Custom Sort, add a second level: Then by Trial, Smallest to Largest.'
      : 'The rows are not in order by Sample. Use Custom Sort on the whole table: Sort by Sample (A to Z), then by Trial (Smallest to Largest).';
  }

  const tMol = task('moles', 'NaOH (mol) calculated with a fixed reference to B2', 4);
  const b2 = cells.B2;
  const conc = b2 && typeof b2.v === 'number' ? b2.v : NaN;
  const molCells = [];
  if (!near(conc, d.conc)) tMol.msg = `B2 should hold the NaOH concentration from the Lab page (${d.conc.toFixed(4)} M). Put it back and don't overwrite it.`;
  else if (!col.mol) tMol.msg = 'Add a column D with the header "NaOH (mol)" and calculate moles in each row: volume in liters × concentration.';
  else {
    const res = rows.map(x => { const c = cells[addr(col.mol, x.r)]; if (c) molCells.push(c); return { x, c, ok: c && typeof x.vol === 'number' && near(c.v, x.vol / 1000 * conc) }; });
    const typedNums = res.filter(o => o.c && !o.c.formula && typeof o.c.v === 'number');
    const bad = res.filter(o => !o.ok);
    if (!res.some(o => o.c)) tMol.msg = 'Column D is empty. In D4 type =C4/1000*$B$2 and fill it down to the last row.';
    else if (typedNums.length) tMol.msg = `Some moles are typed numbers, not formulas (row ${typedNums.map(o => o.x.r).join(', ')}). Use a formula so Excel does the math.`;
    else if (!bad.length) { tMol.ok = true; tMol.msg = 'Every row is right.'; }
    else if (res[0].ok) tMol.msg = `Row ${res[0].x.r} is right but the rows below are not. When you filled down, the reference to B2 moved to B3, B4, …. Make it fixed with dollar signs: $B$2.`;
    else if (res.every(o => o.c && near(o.c.v, o.x.vol * conc))) tMol.msg = 'The values are 1000× too big: change mL to L by dividing the volume by 1000.';
    else tMol.msg = `Row ${bad[0].x.r} is not volume (L) × concentration. The formula in D4 should be =C4/1000*$B$2.`;
  }
  const tMolF = task('molfmt', 'NaOH (mol) shown in scientific notation, 3 decimals', 6);
  if (!molCells.length) tMolF.msg = 'No moles column to check yet.';
  else {
    const off = molCells.filter(c => !(c.fmt.sci && c.fmt.dec === 3));
    if (!off.length) { tMolF.ok = true; tMolF.msg = 'Like 2.515E-03.'; }
    else if (off.every(c => c.fmt.sci)) tMolF.msg = 'Scientific notation is on, but use 3 decimal places (e.g. 2.515E-03). Use Increase/Decrease Decimal.';
    else tMolF.msg = `${off.length} cell(s) in the moles column are not in scientific notation (e.g. ${off[0].ref}). Select D4:D18 → Number Format → Scientific, then set 3 decimals.`;
  }

  // ---------- results table: F3:M8, statistics on the moles in column D ----------
  const HDR = [['Sample', /^sample$/i], ['Mean', /^mean\b/i], ['SD', /^(sd|st\.?\s*dev|standard deviation)\b/i], ['%CV', /cv/i],
    ['Median', /^median\b/i], ['Max', /^max(imum)?\b/i], ['Min', /^min(imum)?\b/i], ['n', /^(n|count)\b/i]];
  const HDR_TEXT = HDR.map(h => h[0]).join(', ');
  const tTab = task('table', 'Results table in F3:M8 (' + HDR_TEXT + '), column E empty', 5);
  const badHdr = HDR.map((h, i) => ({ ref: addr(6 + i, 3), want: h[0], ok: h[1].test(text(cells[addr(6 + i, 3)])) })).filter(h => !h.ok);
  const labelsOK = LAB_A_SAMPLES.every((s, i) => text(cells[addr(6, 4 + i)]).toUpperCase() === s);
  const eUsed = [3, 4, 5, 6, 7, 8].map(r => 'E' + r).filter(r => text(cells[r]) !== '');
  if (badHdr.length && badHdr[0].ref === 'F3') {
    let where = null;
    for (const c of Object.values(cells)) {
      if (!/^sample$/i.test(text(c)) || c.ref === addr(col.sample, hr)) continue;
      const q = splitRef(c.ref);
      if (/^mean/i.test(text(cells[addr(q.col + 1, q.row)]))) { where = c.ref; break; }
    }
    tTab.msg = where ? `Your results table starts at ${where}. Move it so the headers ${HDR_TEXT} are in F3:M3, with column E left empty.`
      : `Put the headers ${HDR_TEXT} in F3:M3, and leave column E empty.`;
  } else if (badHdr.length) tTab.msg = `Headers F3:M3 should read ${HDR_TEXT}. Check ${fmtList(badHdr.map(h => `${h.ref} (${h.want})`))}.`;
  else if (!labelsOK) tTab.msg = 'Put the sample letters A, B, C, D, E in F4:F8, one per row.';
  else if (eUsed.length) tTab.msg = `Leave column E empty (${eUsed.join(', ')} has something in it). The empty column keeps Excel from treating the data and the results as one table when you sort.`;
  else { tTab.ok = true; tTab.msg = 'Headers in F3:M3 and samples A–E in F4:F8.'; }
  if (!tTab.ok) return { tasks: T, name, done: false };

  const rangeOf = c => { const m = c && /\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/.exec(c.ftext || ''); return m ? `${m[1]}${m[2]}:${m[3]}${m[4]}` : null; };
  const molRange = s => { const r = rows.filter(x => x.sample === s).map(x => x.r); if (!r.length) return null; const lo = Math.min(...r), hi = Math.max(...r); return hi - lo === r.length - 1 ? `D${lo}:D${hi}` : null; };
  const rounded = (v, e) => typeof v === 'number' && Math.abs(v - e) <= 0.0006 * Math.abs(e); // a typed, rounded copy
  const GROUP = {
    empty: 'empty. Each needs a formula.',
    notnum: 'showing an error or text instead of a number.',
    typed: 'typed numbers. Use formulas, so the results update if the data change.',
    vol: 'using the volumes in column C. Use the moles in column D (A is D4:D6, B is D7:D9, … E is D16:D18).',
    pop: 'using STDEV.P (population). Three trials are a sample of all the titrations you could run: use STDEV.S (or STDEV).',
    no100: 'missing the × 100. %CV = 100 × SD ÷ mean.',
    follows: 'calculated from a mean or SD in the same row that isn\'t right yet. Fix that first.',
    swap: 'holding the other one: MAX goes in the Max column (K) and MIN in the Min column (L).',
    unsorted: 'not checkable yet: sort the data first, so each sample\'s three rows are together.'
  };
  // diagnose one cell: null when right, else {kind, ref, msg}
  function diag(c, ref, exp, s, what, alt) {
    const rng = molRange(s), K = kind => ({ kind, ref });
    if (!c || c.v === null || c.v === '') return K('empty');
    if (typeof c.v !== 'number') return K('notnum');
    if (near(c.v, exp)) return c.formula ? null : K('typed');
    if (!rng) return K('unsorted');
    if (alt.vol != null && near(c.v, alt.vol)) return K('vol');
    if (alt.pop != null && near(c.v, alt.pop)) return K('pop');
    if (alt.swap != null && near(c.v, alt.swap)) return K('swap');
    if (alt.no100 != null && near(c.v, alt.no100)) return c.fmt.pct ? null : K('no100');
    if (alt.follows) return K('follows');
    const got = rangeOf(c);
    if (c.formula && got && got !== rng) return { kind: 'range', ref, msg: `${ref} uses ${got}, but sample ${s}'s moles are in ${rng}.` };
    if (!c.formula && rounded(c.v, exp)) return K('typed');
    return { kind: 'other', ref, msg: `${ref} isn't the ${what} of sample ${s} (${rng}).` };
  }
  // specs: [{col, what, exp(s), alt(s, i)}]
  const statTask = (id, title, specs) => {
    const t = task(id, title, 5), found = [];
    for (const sp of specs) LAB_A_SAMPLES.forEach((s, i) => { const ref = addr(sp.col, 4 + i); const m = diag(cells[ref], ref, sp.exp(s), s, sp.what, sp.alt ? sp.alt(s, i) : {}); if (m) found.push(m); });
    if (!found.length) { t.ok = true; t.msg = specs.map(sp => `${addr(sp.col, 4)}:${addr(sp.col, 8)}`).join(', ') + ' all right.'; return t; }
    const parts = [];
    for (const k of Object.keys(GROUP)) { const refs = found.filter(f => f.kind === k).map(f => f.ref); if (refs.length) parts.push(`${fmtList(refs)} ${refs.length > 1 ? 'are' : 'is'} ${GROUP[k]}`); }
    found.filter(f => f.msg).forEach(f => parts.push(f.msg));
    t.msg = parts.slice(0, 3).join(' ') + (parts.length > 3 ? ' (and more)' : '');
    return t;
  };
  const st = s => d.stats[s];
  const tMean = statTask('mean', 'Mean of each sample (G4:G8), by formula', [{ col: 7, what: 'mean', exp: s => st(s).mean, alt: s => ({ vol: st(s).vMean }) }]);
  const tSD = statTask('sd', 'Standard deviation of each sample (H4:H8), using STDEV.S', [{ col: 8, what: 'standard deviation', exp: s => st(s).sdS, alt: s => ({ vol: st(s).vSdS, pop: st(s).sdP }) }]);
  statTask('cv', '%CV of each sample (I4:I8), by formula', [{ col: 9, what: '%CV', exp: s => st(s).cv, alt: (s, i) => {
    const bad = [!tMean.ok && !near((cells[addr(7, 4 + i)] || {}).v, st(s).mean) ? 'G' + (4 + i) : null, !tSD.ok && !near((cells[addr(8, 4 + i)] || {}).v, st(s).sdS) ? 'H' + (4 + i) : null].filter(Boolean);
    return { no100: st(s).cv / 100, pop: st(s).cvP, follows: bad.length ? bad.join(' and ') : null };
  } }]);
  statTask('mmm', 'Median, Max and Min of each sample (J4:L8), by formula', [
    { col: 10, what: 'median', exp: s => st(s).median, alt: s => ({ vol: st(s).vMedian }) },
    { col: 11, what: 'maximum', exp: s => st(s).max, alt: s => ({ vol: st(s).vMax, swap: st(s).min }) },
    { col: 12, what: 'minimum', exp: s => st(s).min, alt: s => ({ vol: st(s).vMin, swap: st(s).max }) }]);
  // n: COUNT gives 3 for any three rows, so check the range written in the formula too
  { const t = task('n', 'n (number of trials) of each sample (M4:M8), by COUNT', 5), msgs = [], typedN = [];
    LAB_A_SAMPLES.forEach((s, i) => {
      const ref = addr(13, 4 + i), c = cells[ref], rng = molRange(s);
      if (!c || c.v === null || c.v === '') { msgs.push(`${ref} is empty: use =COUNT(${rng || 'D…'}).`); return; }
      if (typeof c.v !== 'number') { msgs.push(`${ref} shows "${c.v}", not a number.`); return; }
      if (c.v !== 3) { msgs.push(`${ref} counts ${c.v} values; sample ${s} has 3 trials${rng ? ' (' + rng + ')' : ''}.`); return; }
      if (!c.formula) { typedN.push(ref); return; }
      const got = rangeOf(c);
      // only question a plain 3-row block (e.g. =COUNT(D4:D6)); FILTER, COUNTIF, etc. pick the rows themselves
      const g = got && /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(got);
      if (g && rng && +g[4] - +g[2] === 2 && got !== rng && got !== rng.replace(/D/g, 'C')) msgs.push(`${ref} counts ${got}, but sample ${s} is in ${rng}.`);
    });
    if (typedN.length) msgs.unshift(`${fmtList(typedN)} ${typedN.length > 1 ? 'are typed numbers' : 'is a typed number'}. Use =COUNT(D4:D6) for A, =COUNT(D7:D9) for B, and so on, so Excel counts the trials.`);
    if (!msgs.length) { t.ok = true; t.msg = 'M4:M8 all right.'; } else t.msg = msgs.slice(0, 3).join(' ') + (msgs.length > 3 ? ' (and more)' : ''); }
  const tFmt = task('statfmt', 'Results formatted: scientific 3 decimals; %CV 3 decimals; n whole number', 6);
  { const off = [];
    for (let i = 0; i < 5; i++) {
      for (const cn of [7, 8, 10, 11, 12]) { const c = cells[addr(cn, 4 + i)]; if (c && !(c.fmt.sci && c.fmt.dec === 3)) off.push(c.ref); }
      let c = cells[addr(9, 4 + i)]; if (c && !(c.fmt.dec === 3 && !c.fmt.general)) off.push(c.ref);
      c = cells[addr(13, 4 + i)]; if (c && !(c.fmt.general || (c.fmt.dec === 0 && !c.fmt.sci))) off.push(c.ref);
    }
    if (!off.length) { tFmt.ok = true; tFmt.msg = 'Mean, SD, Median, Max, Min like 3.126E-03; %CV with 3 decimals; n as 3.'; }
    else tFmt.msg = `Check the number format of ${fmtList(off.slice(0, 5))}${off.length > 5 ? ' and others' : ''}. Mean, SD, Median, Max, Min: Scientific, 3 decimals (3.126E-03). %CV: 3 decimals (0.194 or 1.938E-01). n: a whole number (3).`;
  }
  const done = T.every(t => t.ok);
  return { tasks: T, name, done, code: done ? completionCode(name, 'A') : null };
}

if (typeof module !== 'undefined') module.exports = { labAData, labABlock, completionCode, verifyCode, normName, makeCsv };
