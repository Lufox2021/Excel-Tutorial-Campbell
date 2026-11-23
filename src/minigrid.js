/* "Try it" mini grid: a tiny formula sandbox (not a fake Excel).
   Supports numbers, + - * / ^, parentheses, cell refs with $, ranges, and
   SUM AVERAGE MEDIAN MIN MAX COUNT STDEV.S STDEV.P STDEV. */

function mgTokenize(s) {
  const out = []; let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (/\s/.test(ch)) { i++; continue; }
    let m;
    if ((m = /^\d+(\.\d+)?([eE][+-]?\d+)?|^\.\d+/.exec(s.slice(i)))) { out.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
    if ((m = /^[A-Za-z][A-Za-z0-9.]*(?=\s*\()/.exec(s.slice(i)))) { out.push({ t: 'fn', v: m[0].toUpperCase() }); i += m[0].length; continue; }
    if ((m = /^\$?[A-Za-z]{1,2}\$?\d+/.exec(s.slice(i)))) { out.push({ t: 'ref', v: m[0].toUpperCase().replace(/\$/g, '') }); i += m[0].length; continue; }
    if ('+-*/^(),:'.includes(ch)) { out.push({ t: ch }); i++; continue; }
    throw new Error(`I don't understand "${s.slice(i)}"`);
  }
  return out;
}
const MG_FN = {
  SUM: a => a.reduce((s, x) => s + x, 0),
  AVERAGE: a => { if (!a.length) throw new Error('#DIV/0!'); return a.reduce((s, x) => s + x, 0) / a.length; },
  MEDIAN: a => median(a), MIN: a => Math.min(...a), MAX: a => Math.max(...a), COUNT: a => a.length,
  'STDEV.S': a => { if (a.length < 2) throw new Error('#DIV/0!'); return sdS(a); }, STDEV: a => { if (a.length < 2) throw new Error('#DIV/0!'); return sdS(a); },
  'STDEV.P': a => sdP(a)
};
const MG_FN2 = {
  SLOPE: (y, x) => linfit(x, y).m, INTERCEPT: (y, x) => linfit(x, y).b, RSQ: (y, x) => linfit(x, y).r2,
  CORREL: (y, x) => { const f = linfit(x, y); return Math.sign(f.m) * Math.sqrt(f.r2); }
};
function mgEval(src, getVal) {
  const tk = mgTokenize(src); let p = 0;
  const peek = () => tk[p], eat = t => { if (!tk[p] || tk[p].t !== t) throw new Error('Check the formula near "' + (tk[p] ? (tk[p].v ?? tk[p].t) : 'the end') + '"'); return tk[p++]; };
  const num = v => { if (Array.isArray(v)) throw new Error('A range like A1:A3 only works inside a function, e.g. =AVERAGE(A1:A3)'); return v; };
  function range(a, b) {
    const A = splitRef(a), B = splitRef(b), out = [];
    for (let r = Math.min(A.row, B.row); r <= Math.max(A.row, B.row); r++)
      for (let c = Math.min(A.col, B.col); c <= Math.max(A.col, B.col); c++) { const v = getVal(addr(c, r)); if (typeof v === 'number') out.push(v); }
    return out;
  }
  function primary() {
    const t = peek(); if (!t) throw new Error('The formula ends too soon');
    if (t.t === 'num') { p++; return t.v; }
    if (t.t === '(') { p++; const v = expr(); eat(')'); return v; }
    if (t.t === 'ref') { p++; if (peek() && peek().t === ':') { p++; const b = eat('ref'); return range(t.v, b.v); } const v = getVal(t.v); if (typeof v === 'string' && v !== '') throw new Error(t.v + ' holds text, not a number'); return typeof v === 'number' ? v : 0; }
    if (t.t === 'fn') {
      p++; const f = MG_FN[t.v], f2 = MG_FN2[t.v]; if (!f && !f2) throw new Error(`#NAME? (${t.v} isn't available here)`);
      eat('('); const raw = [];
      if (peek() && peek().t !== ')') { do { raw.push(expr()); } while (peek() && peek().t === ',' && ++p); }
      eat(')');
      if (f2) { if (raw.length !== 2 || !Array.isArray(raw[0]) || !Array.isArray(raw[1]) || raw[0].length !== raw[1].length || raw[0].length < 2) throw new Error(`${t.v} needs two ranges of the same size: ${t.v}(known_y's, known_x's)`); return f2(raw[0], raw[1]); }
      return f(raw.flatMap(v => Array.isArray(v) ? v : [v]));
    }
    throw new Error('Check the formula near "' + (t.v ?? t.t) + '"');
  }
  function unary() { if (peek() && peek().t === '-') { p++; return -num(unary()); } if (peek() && peek().t === '+') { p++; return num(unary()); } return primary(); }
  function power() { let v = unary(); while (peek() && peek().t === '^') { p++; v = Math.pow(num(v), num(unary())); } return v; }
  function term() { let v = power(); while (peek() && (peek().t === '*' || peek().t === '/')) { const op = tk[p++].t, b = num(power()); v = num(v); if (op === '/' && b === 0) throw new Error('#DIV/0!'); v = op === '*' ? v * b : v / b; } return v; }
  function expr() { let v = term(); while (peek() && (peek().t === '+' || peek().t === '-')) { const op = tk[p++].t, b = num(term()); v = op === '+' ? num(v) + b : num(v) - b; } return v; }
  const v = expr(); if (p < tk.length) throw new Error('Check the formula near "' + (tk[p].v ?? tk[p].t) + '"');
  return num(v);
}
// shift relative row references by dr (what Fill down does)
function mgShift(f, dr) {
  return f.replace(/(\$?)([A-Za-z]{1,2})(\$?)(\d+)(?![\d(A-Za-z])/g, (m, cd, c, rd, r, off, str) => {
    if (/[A-Za-z.]/.test(str[off - 1] || '')) return m; // part of a function name
    return cd + c + rd + (rd ? r : String(+r + dr));
  });
}

/* MiniGrid({mount, cols, rows, cells:{A1:'text'|number|'=f'}, editable:[refs], fillRows:{from,to}, goal(fn)->{ok,msg}, prompt}) */
function MiniGrid(o) {
  const st = { cells: { ...o.cells }, sel: o.editable[0], showF: false };
  const root = document.createElement('div'); root.className = 'mg';
  o.mount.appendChild(root);
  const isF = v => typeof v === 'string' && v.startsWith('=');
  const val = (ref, seen = new Set()) => {
    const raw = st.cells[ref];
    if (!isF(raw)) return raw === undefined || raw === '' ? null : (typeof raw === 'number' || isNaN(Number(raw)) ? raw : Number(raw));
    if (seen.has(ref)) throw new Error('Circular reference');
    seen.add(ref);
    const v = mgEval(raw.slice(1), r => val(r, seen)); seen.delete(ref); return v;
  };
  const show = ref => { try { const v = val(ref); return v === null ? '' : typeof v === 'number' ? (+v.toPrecision(10)).toString() : v; } catch (e) { return '⚠'; } };
  function render() {
    let h = `<div class="mg-prompt">${o.prompt}</div><div class="mg-bar"><span class="mg-name">${st.sel}</span><span class="mg-fx">fx</span><input class="mg-in" spellcheck="false" aria-label="Formula bar"></div><div class="mg-wrap"><table class="mg-t"><tr><th></th>`;
    for (let c = 1; c <= o.cols; c++) h += `<th>${numToCol(c)}</th>`;
    h += '</tr>';
    for (let r = 1; r <= o.rows; r++) {
      h += `<tr><th>${r}</th>`;
      for (let c = 1; c <= o.cols; c++) {
        const ref = addr(c, r), raw = st.cells[ref], ed = o.editable.includes(ref);
        const disp = st.showF && isF(raw) ? raw : show(ref);
        const num = !isF(raw) && typeof val(ref) === 'number' || isF(raw) && !st.showF;
        h += `<td data-ref="${ref}" class="${ed ? 'ed' : ''} ${ref === st.sel ? 'sel' : ''} ${num ? 'num' : ''}">${disp === '⚠' ? '<span class="mg-err">!</span>' : esc(String(disp))}</td>`;
      }
      h += '</tr>';
    }
    h += `</table></div><div class="mg-btns">${o.fill ? '<button class="btn small mg-fill">Fill down ↓</button>' : ''}<button class="btn small ghost mg-showf">${st.showF ? 'Show values' : 'Show formulas'}</button><button class="btn small ghost mg-reset">Reset</button></div><div class="mg-msg"></div><div class="mg-goal"></div>`;
    root.innerHTML = h;
    const inp = root.querySelector('.mg-in');
    inp.value = st.cells[st.sel] ?? ''; inp.readOnly = !o.editable.includes(st.sel);
    inp.addEventListener('keydown', e => { if (e.key === 'Enter') { commit(inp.value); } });
    inp.addEventListener('change', () => commit(inp.value));
    root.querySelectorAll('td').forEach(td => td.addEventListener('click', () => { st.sel = td.dataset.ref; render(); const i = root.querySelector('.mg-in'); if (!i.readOnly) i.focus(); }));
    root.querySelector('.mg-showf').onclick = () => { st.showF = !st.showF; render(); };
    root.querySelector('.mg-reset').onclick = () => { st.cells = { ...o.cells }; st.sel = o.editable[0]; render(); };
    if (o.fill) root.querySelector('.mg-fill').onclick = fillDown;
    status();
  }
  function commit(v) {
    if (!o.editable.includes(st.sel)) return;
    v = v.trim(); st.cells[st.sel] = v === '' ? '' : isF(v) ? v : (isNaN(Number(v)) ? v : Number(v));
    render();
  }
  function fillDown() {
    const src = st.sel, s = splitRef(src), raw = st.cells[src];
    const msg = root.querySelector('.mg-msg');
    if (!o.fill.cols.includes(numToCol(s.col)) || s.row !== o.fill.from) { msg.textContent = `Select ${o.fill.cols.map(c => c + o.fill.from).join(' or ')} first, then Fill down.`; return; }
    if (!isF(raw)) { msg.textContent = 'Type a formula in this cell first (start with =).'; return; }
    for (let r = s.row + 1; r <= o.fill.to; r++) st.cells[addr(s.col, r)] = '=' + mgShift(raw.slice(1), r - s.row);
    render();
    root.querySelector('.mg-msg').textContent = `Copied ${src} down to ${addr(s.col, o.fill.to)}. Click the filled cells (or Show formulas) to see how each reference changed.`;
  }
  function status() {
    const msg = root.querySelector('.mg-msg');
    try { val(st.sel); } catch (e) { msg.textContent = st.sel + ': ' + e.message; }
    if (o.goal) {
      const g = o.goal(ref => { try { return val(ref); } catch (e) { return null; } }, st.cells);
      const el = root.querySelector('.mg-goal'); el.className = 'mg-goal ' + (g.ok ? 'ok' : ''); el.innerHTML = g.msg;
    }
  }
  render();
}
