/* Excel Workbench — page wiring */
const $ = s => document.querySelector(s), $$ = s => Array.from(document.querySelectorAll(s));
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// ---------- tabs and lessons ----------
function showTab(id) {
  $$('nav.tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === id));
  $$('.tab').forEach(t => t.classList.toggle('on', t.id === id));
  window.scrollTo(0, 0);
}
function showLesson(n) {
  $$('#mods button[data-m]').forEach(b => b.classList.toggle('on', b.dataset.m === String(n)));
  $$('[data-lesson]').forEach(a => a.hidden = a.dataset.lesson !== String(n));
  showTab('learn');
}
$$('nav.tabs button').forEach(b => b.onclick = () => showTab(b.dataset.tab));
$$('#mods button[data-m]').forEach(b => b.onclick = () => showLesson(+b.dataset.m));
document.addEventListener('click', e => {
  const g = e.target.closest('[data-go]'); if (g) { e.preventDefault(); showTab(g.dataset.go); }
  const l = e.target.closest('[data-golesson]'); if (l) { e.preventDefault(); showLesson(+l.dataset.golesson); }
});

// ---------- downloads (normal page, with a fallback) ----------
// In the claude.ai preview, downloads go through the viewer's download capability;
// on GitHub Pages (or a local file) a normal browser download is used.
let dlCap = null;
try { if (window.claude && window.claude.use) window.claude.use('downloads').then(d => dlCap = d).catch(() => {}); } catch (e) {}
async function download(blob, filename) {
  if (dlCap) {
    try { await dlCap.save({ filename, data: blob }); return true; }
    catch (e) { return e && e.code === 'declined'; }
  }
  try {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
    return true;
  } catch (e) { return false; }
}

// ---------- Try-it grids ----------
MiniGrid({
  mount: $('#try4'), cols: 5, rows: 7,
  prompt: '<b>Try it.</b> Click <b>B2</b>, type <code>=A2/1000*E2</code> in the formula bar, press Enter, then <b>Fill down</b>. Look at B3–B6. Then change B2 to <code>=A2/1000*$E$2</code> and fill down again.',
  cells: { A1: 'Volume (mL)', B1: 'NaOH (mol)', D1: '', E1: 'Conc (M)', E2: 0.1000, A2: 24.15, A3: 24.32, A4: 24.08, A5: 21.67, A6: 21.80 },
  editable: ['B2', 'B3', 'B4', 'B5', 'B6'], fill: { cols: ['B'], from: 2, to: 6 },
  goal: (v, cells) => {
    const ok = [2, 3, 4, 5, 6].every(r => { const x = v('B' + r); return typeof x === 'number' && Math.abs(x - v('A' + r) / 1000 * 0.1) < 1e-12; });
    if (ok) return { ok: true, msg: '✓ Every row is right: the fixed reference $E$2 stayed on the concentration.' };
    if (typeof v('B2') === 'number' && Math.abs(v('B2') - 0.002415) < 1e-12 && v('B3') === 0) return { ok: false, msg: 'B2 is right, but B3–B6 are 0: after filling down they point to E3, E4, … which are empty. That\'s the missing $.' };
    return { ok: false, msg: 'Goal: B2:B6 = volume in L × the concentration in E2.' };
  }
});
MiniGrid({
  mount: $('#try5'), cols: 4, rows: 7,
  prompt: '<b>Try it.</b> Three trials are in A2:A4. In D2 find the mean, in D3 the sample SD, in D4 the population SD, and in D5 the %CV (use the sample SD).',
  cells: { A1: 'Volume (mL)', A2: 24.15, A3: 24.32, A4: 24.08, C2: 'Mean', C3: 'SD (STDEV.S)', C4: 'SD (STDEV.P)', C5: '%CV' },
  editable: ['D2', 'D3', 'D4', 'D5'],
  goal: v => {
    const a = [24.15, 24.32, 24.08], m = mean(a), S = sdS(a), P = sdP(a), cv = S / m * 100;
    const n = (x, e) => typeof x === 'number' && Math.abs(x - e) < 1e-9;
    const got = [n(v('D2'), m), n(v('D3'), S), n(v('D4'), P), n(v('D5'), cv)];
    if (got.every(Boolean)) return { ok: true, msg: `✓ All four right. Notice STDEV.P (${P.toFixed(4)}) is smaller than STDEV.S (${S.toFixed(4)}).` };
    if (n(v('D5'), cv / 100)) return { ok: false, msg: 'D5 is SD ÷ mean, but %CV also needs × 100.' };
    if (n(v('D5'), P / m * 100)) return { ok: false, msg: 'D5 used the population SD; use D3 (STDEV.S).' };
    const names = ['mean', 'STDEV.S', 'STDEV.P', '%CV'];
    return { ok: false, msg: 'Still to do: ' + names.filter((x, i) => !got[i]).join(', ') + '.' };
  }
});

// ---------- Module 7: Scatter vs Line picture ----------
(function () {
  const V = [0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4], P = V.map(v => 1.5 / v), W = 250, H = 160, L = 34, B = 26, Tm = 10, R = 10;
  const sx = (x, even) => even ? L + (W - L - R) * (V.indexOf(x) + 0.5) / V.length : L + (W - L - R) * x / 4.2;
  const sy = y => Tm + (H - Tm - B) * (1 - y / 3.2);
  function svg(even) {
    let g = `<svg width="${W}" height="${H}" role="img" aria-label="${even ? 'Line chart' : 'Scatter chart'} of pressure vs volume"><line x1="${L}" y1="${H - B}" x2="${W - R}" y2="${H - B}" stroke="#888"/><line x1="${L}" y1="${Tm}" x2="${L}" y2="${H - B}" stroke="#888"/>`;
    const ticks = even ? V : [0, 1, 2, 3, 4];
    ticks.forEach(t => g += `<text x="${sx(t, even)}" y="${H - B + 13}" font-size="9" text-anchor="middle" fill="#5b6470">${even ? t.toFixed(2) : t}</text>`);
    [0, 1, 2, 3].forEach(t => g += `<text x="${L - 5}" y="${sy(t) + 3}" font-size="9" text-anchor="end" fill="#5b6470">${t}</text>`);
    if (even) g += `<polyline fill="none" stroke="#b42318" stroke-width="1.5" points="${V.map((v, i) => sx(v, 1) + ',' + sy(P[i])).join(' ')}"/>`;
    V.forEach((v, i) => g += `<circle cx="${sx(v, even)}" cy="${sy(P[i])}" r="3.5" fill="${even ? '#b42318' : '#1d6f42'}"/>`);
    return g + `<text x="${(L + W) / 2}" y="${H - 2}" font-size="10" text-anchor="middle" fill="#1f2328">Volume (L)</text></svg>`;
  }
  $('#m7pics').innerHTML = `<figure class="mini"><div>${svg(false)}</div><figcaption>Scatter: volumes at their true positions. P ∝ 1/V shows as a smooth curve.</figcaption></figure><figure class="mini bad"><div>${svg(true)}</div><figcaption>Line: volumes spaced evenly as labels. The curve is the wrong shape.</figcaption></figure>`;
})();
MiniGrid({
  mount: $('#try11'), cols: 6, rows: 6,
  prompt: '<b>Try it.</b> Five standards are in A2:B6 and an unknown\'s absorbance in D2. In F2 find the slope, in F3 the intercept, and in F4 the unknown\'s concentration from them.',
  cells: { A1: 'Conc (M)', B1: 'Absorbance', A2: 0.00001, A3: 0.00002, A4: 0.00003, A5: 0.00004, A6: 0.00005, B2: 0.128, B3: 0.251, B4: 0.372, B5: 0.497, B6: 0.618, D1: 'Unknown A', D2: 0.300, E2: 'Slope', E3: 'Intercept', E4: 'Conc (M)' },
  editable: ['F2', 'F3', 'F4'],
  goal: v => {
    const x = [1, 2, 3, 4, 5].map(k => k * 1e-5), y = [0.128, 0.251, 0.372, 0.497, 0.618], f = linfit(x, y), c = (0.3 - f.b) / f.m;
    const n = (a, e) => typeof a === 'number' && Math.abs(a - e) <= 1e-9 * Math.max(1, Math.abs(e));
    if (n(v('F2'), f.m) && n(v('F3'), f.b) && n(v('F4'), c)) return { ok: true, msg: `✓ Slope ${f.m.toFixed(0)} M⁻¹, intercept ${f.b.toFixed(4)}, unknown ${c.toExponential(3)} M.` };
    if (n(v('F2'), f.mSwap)) return { ok: false, msg: 'F2 has the ranges swapped: =SLOPE(known y\'s, known x\'s) = SLOPE(B2:B6,A2:A6).' };
    if (n(v('F4'), 0.3 / f.m)) return { ok: false, msg: 'F4 leaves out the intercept: c = (A − b) ÷ m, so =(D2-F3)/F2.' };
    const todo = [['F2', n(v('F2'), f.m), 'slope'], ['F3', n(v('F3'), f.b), 'intercept'], ['F4', n(v('F4'), c), 'concentration']].filter(t => !t[1]).map(t => `${t[0]} (${t[2]})`);
    return { ok: false, msg: 'Still to do: ' + todo.join(', ') + '. Hint: =SLOPE(B2:B6,A2:A6), =INTERCEPT(B2:B6,A2:A6).' };
  }
});

// ---------- Labs ----------
function tableHTML(rows, startRow) {
  const nc = Math.max(...rows.map(r => r.length));
  let h = '<table class="data"><tr><th class="rh"></th>' + Array.from({ length: nc }, (_, c) => `<th class="rh">${numToCol(c + 1)}</th>`).join('') + '</tr>';
  rows.forEach((r, i) => {
    h += `<tr><td class="rh">${startRow + i}</td>`;
    for (let c = 0; c < nc; c++) { const v = r[c] ?? ''; h += `<td class="${/^-?[\d.]+(E[+-]\d+)?$/i.test(v) ? 'n' : ''}">${esc(v)}</td>`; }
    h += '</tr>';
  });
  return h + '</table>';
}
const LABS = {
  A: { title: 'Lab A · Replicate titrations', modules: '1–6',
    intro: 'You titrated five acid samples (A–E) with NaOH, three trials each. Your job: get the data into Excel, organize it, convert volumes to moles, and summarize each sample\'s precision.',
    data: labAData, block: labABlock,
    extra: d => `<h3>Sample E: type these yourself</h3><p>These trials came in late, from the lab notebook. Type them into rows 16–18 (Sample, Trial, Volume):</p>${tableHTML(d.typed.map(x => [x.sample, String(x.trial), x.vol.toFixed(2)]), 16)}`,
    tasks: [
      [1, 'Paste the data at <b>A1</b>. Type your name in <b>B1</b>, exactly as above.'],
      [1, 'Type Sample E\'s three trials into <b>rows 16–18</b>.'],
      [3, 'Sort the table A3:C18 by <b>Sample</b>, then by <b>Trial</b>, keeping every row together.'],
      [4, 'In <b>D3</b> type the header <code>NaOH (mol)</code>. In D4 calculate moles of NaOH (volume in L × concentration in B2) with a <b>fixed reference</b> to B2, and fill down to D18.'],
      [6, 'Show the moles in <b>scientific notation with 3 decimals</b>.'],
      [5, 'Leave <b>column E empty</b>. Make a results table with the headers <b>Sample, Mean, SD, %CV, Median, Max, Min, n</b> in <b>F3:M3</b> and the samples A–E in <b>F4:F8</b>.'],
      [5, 'In G4:M8 use formulas on the <b>moles in column D</b> for each sample: <b>mean</b> (AVERAGE), <b>SD</b> (STDEV.S), <b>%CV</b> = 100 × SD ÷ mean, <b>median</b> (MEDIAN), <b>max</b> (MAX), <b>min</b> (MIN) and <b>n</b> (COUNT). Sample A is D4:D6, B is D7:D9, and so on.'],
      [6, 'Format the results: Mean, SD, Median, Max and Min in <b>scientific notation with 3 decimals</b>; %CV with <b>3 decimals</b>; n as a <b>whole number</b>. Center the results table.']] },
  B: { title: 'Lab B · Boyle\'s law', modules: '7–9',
    intro: 'You measured the pressure of a fixed amount of gas at eight volumes, once at 298 K and once at 348 K. Your job: plot both data sets on one well-labeled XY chart.',
    data: labBData, block: labBBlock, extra: () => '',
    tasks: [
      [1, 'Paste the data at <b>A1</b>. Type your name in <b>B1</b>, exactly as above.'],
      [7, 'Select <b>A3:C11</b> (headers included) and <span class="path">Insert → Scatter</span> (markers only). You should get <b>two series</b>, with Volume on the x axis.'],
      [8, 'Check that the legend shows the two headers ("P at 298 K (atm)", "P at 348 K (atm)"), not "Series1".'],
      [8, 'Add a <b>chart title</b>, and <b>axis titles</b> with units: <code>Volume (L)</code> and <code>Pressure (atm)</code>. Keep the <b>legend</b>.'],
      [9, 'Set the title, both axis titles and the legend to <b>12 pt</b>.'],
      [11, 'Show the data as <b>markers only</b>, with no lines joining the points.']] },
  C: { title: 'Lab C · Beer\'s law calibration curve', modules: '10–11',
    intro: 'You measured the absorbance of six standard solutions and three unknowns. Your job: make a calibration curve, fit a line, and find each unknown\'s concentration.',
    data: labCData, block: labCBlock, extra: () => '',
    tasks: [
      [1, 'Paste the data at <b>A1</b>. Type your name in <b>B1</b>, exactly as above.'],
      [7, 'Select the standards, <b>A3:B9</b>, and <span class="path">Insert → Scatter</span> (markers only). Don\'t plot the unknowns.'],
      [8, 'Add a <b>chart title</b> and <b>axis titles</b> <code>Concentration (M)</code> and <code>Absorbance</code>. Turn the <b>legend off</b> (one series).'],
      [9, 'Set the title and both axis titles to <b>12 pt</b>. Keep markers only, no connecting lines.'],
      [10, 'Add a <b>linear trendline</b> showing the <b>equation</b> and <b>R²</b> on the chart.'],
      [10, 'In <b>H4:H6</b> type the labels Slope, Intercept, R². In <b>I4:I6</b> calculate them: <code>=SLOPE(B4:B9,A4:A9)</code>, <code>=INTERCEPT(B4:B9,A4:A9)</code>, <code>=RSQ(B4:B9,A4:A9)</code>.'],
      [10, 'In <b>F3</b> type <code>Concentration (M)</code>. In F4 calculate the first unknown\'s concentration, c = (A − b) ÷ m, with <b>fixed references</b> to I4 and I5, and fill down to F6.'],
      [6, 'Format I4, I5 and F4:F6 in <b>scientific notation with 3 decimals</b>, and R² (I6) as a number with <b>4 decimals</b>.']] }
};
let curLab = 'A', labData = null;
function pickLab(l) {
  curLab = l; labData = null;
  $$('#labPick button').forEach(b => b.classList.toggle('on', b.dataset.lab === l));
  $('#labTitle').textContent = LABS[l].title;
  $('#labIntro').innerHTML = LABS[l].intro + ` <span class="muted">(Modules ${LABS[l].modules})</span>`;
  $('#labOut').hidden = true; $('#labNameNote').textContent = ''; $('#labCopyNote').textContent = '';
  if ($('#labName').value.trim()) makeLab();
}
function makeLab() {
  const name = $('#labName').value.trim(), L = LABS[curLab];
  if (!name) { $('#labNameNote').textContent = 'Type your name first.'; return; }
  labData = L.data(name);
  $('#labNameNote').innerHTML = `Data set for <b>${esc(name)}</b>. Type your name in B1 exactly like this; if it doesn't match, the checker will look for a different data set.`;
  $('#labBlock').innerHTML = tableHTML(L.block(labData), 1);
  $('#labExtra').innerHTML = L.extra(labData);
  $('#labTasks').innerHTML = L.tasks.map(([m, t]) => `<li>${t} <a href="#" data-golesson="${m}" class="muted">(Module ${m})</a></li>`).join('');
  $('#labOut').hidden = false; $('#labCopyNote').textContent = '';
  try { localStorage.setItem('ewb-name', name); } catch (e) {}
}
$$('#labPick button').forEach(b => b.onclick = () => pickLab(b.dataset.lab));
document.addEventListener('click', e => { const g = e.target.closest('[data-golab]'); if (g) { e.preventDefault(); showTab('lab'); pickLab(g.dataset.golab); } });
$('#labGo').onclick = makeLab;
$('#labName').addEventListener('keydown', e => { if (e.key === 'Enter') makeLab(); });
try { const n = localStorage.getItem('ewb-name'); if (n) $('#labName').value = n; } catch (e) {}
pickLab('A');
const curBlock = () => LABS[curLab].block(labData);
const fileStem = () => `Lab${curLab}_` + labData.name.replace(/[^A-Za-z0-9]+/g, '_');
$('#labCopy').onclick = async () => {
  const tsv = curBlock().map(r => r.join('\t')).join('\n');
  const note = $('#labCopyNote');
  try { await navigator.clipboard.writeText(tsv); note.textContent = 'Copied. Now click A1 in Excel and paste.'; }
  catch (e) { note.innerHTML = 'Your browser blocked copying. Select all the text below, copy it, and paste at A1:<br><textarea rows="5">' + esc(tsv) + '</textarea>'; }
};
$('#labXlsx').onclick = async () => { if (!await download(makeXlsx(curBlock(), 'Lab ' + curLab), fileStem() + '.xlsx')) $('#labCopyNote').textContent = 'Download blocked here; use Copy data instead.'; };
$('#labCsv').onclick = async () => { if (!await download(new Blob([makeCsv(curBlock())], { type: 'text/csv' }), fileStem() + '.csv')) $('#labCopyNote').textContent = 'Download blocked here; use Copy data instead.'; };

// ---------- Check ----------
async function checkFile(file) {
  const out = $('#checkOut'); out.hidden = false;
  out.innerHTML = '<p class="muted">Reading ' + esc(file.name) + '…</p>';
  if (/\.(csv|xls)$/i.test(file.name)) { out.innerHTML = `<div class="badbox">That's a ${file.name.split('.').pop().toLowerCase()} file. The checker needs the workbook as <b>.xlsx</b> (File → Create a Copy → Download a Copy in Excel for the web).</div>`; return; }
  let wb;
  try { wb = await readXlsx(await file.arrayBuffer()); }
  catch (e) { out.innerHTML = '<div class="badbox">I couldn\'t read this as an Excel workbook (.xlsx). Download it again from Excel with Download a Copy, and upload that file.</div>'; return; }
  const res = checkWorkbook(wb, $('#checkLab').value || null);
  if (res.lab === 'K') { out.innerHTML = '<div class="okbox">This looks like a <b>Key Assessment</b> capstone. Pre-check it on the <a href="#" data-go="ka">Key Assessment</a> tab.</div>'; return; }
  if (!res.lab) { out.innerHTML = '<div class="badbox">I can\'t tell which lab this workbook is for. Choose the lab above and upload it again. (Each lab\'s data should start at A1, with the headers in row 3.)</div>'; return; }
  const all = LAB_TASKS[res.lab], ids = new Set(res.tasks.map(t => t.id));
  let h = `<h3 style="margin-top:0">${esc(LABS[res.lab].title)}</h3>`;
  if (res.done) h += `<div class="okbox"><b>All tasks complete.</b> Your completion code for <b>${esc(res.name)}</b>:<br><span class="code">${res.code}</span><br>Paste this code into the Lab ${res.lab} assignment in Blackboard.</div>`;
  else { const n = res.tasks.filter(t => t.ok).length; h += `<p><b>${n} of ${all.length}</b> tasks done${res.name ? ' for ' + esc(res.name) : ''}. Fix the items marked ✗, download a new copy, and check again.</p>`; }
  h += '<ul class="tasks">';
  for (const t of res.tasks) h += `<li class="${t.ok ? 'ok' : 'no'}"><span class="ic">${t.ok ? '✓' : '✗'}</span><div><div class="t">${esc(t.title)}</div><div class="m">${esc(t.msg)}</div>${t.ok ? '' : `<a href="#" data-golesson="${t.module}">Review Module ${t.module}</a>`}</div></li>`;
  for (const [id, title] of all) if (!ids.has(id)) h += `<li class="skip"><span class="ic">–</span><div><div class="t">${esc(title)}</div><div class="m">Checked after the items above are fixed.</div></div></li>`;
  out.innerHTML = h + '</ul>';
}
$('#file').onchange = e => { if (e.target.files[0]) checkFile(e.target.files[0]); e.target.value = ''; };
const drop = $('#drop');
['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) checkFile(f); });

// ---------- Instructor ----------
$('#vGo').onclick = () => {
  const ok = verifyCode($('#vName').value, $('#vCode').value);
  $('#vOut').innerHTML = !$('#vName').value.trim() || !$('#vCode').value.trim() ? 'Enter a name and a code.' : ok ? '<b style="color:var(--g)">✓ Valid</b> for this name.' : '<b style="color:var(--bad)">✗ Not valid</b> for this name.';
};
$('#vListGo').onclick = () => {
  const lines = $('#vList').value.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  let h = '<table class="data"><tr><th>Name</th><th>Code</th><th>Result</th></tr>';
  for (const l of lines) {
    const m = /^(.*?)[\t,;]\s*([A-Za-z]-[A-Za-z0-9]{6})\s*$/.exec(l);
    if (!m) { h += `<tr><td colspan="2">${esc(l)}</td><td style="color:var(--warn)">couldn't read this line</td></tr>`; continue; }
    const ok = verifyCode(m[1], m[2]);
    h += `<tr><td>${esc(m[1].trim())}</td><td><code>${esc(m[2].toUpperCase())}</code></td><td style="color:${ok ? 'var(--g)' : 'var(--bad)'}">${ok ? '✓ valid' : '✗ not valid'}</td></tr>`;
  }
  $('#vListOut').innerHTML = h + '</table>';
};

// ---------- Key Assessment: Capstones 1 and 2 ----------
$$('#kaPick button').forEach(b => b.onclick = () => {
  $$('#kaPick button').forEach(x => x.classList.toggle('on', x === b));
  $('#cap1').hidden = b.dataset.cap !== '1'; $('#cap2').hidden = b.dataset.cap !== '2';
});
function setupCapstone(p, block, check, tasks, file, okText, sheet) {
  $(`#${p}Block`).innerHTML = tableHTML(block(), 1);
  $(`#${p}Copy`).onclick = async () => {
    const tsv = block().map(r => r.join('\t')).join('\n');
    try { await navigator.clipboard.writeText(tsv); $(`#${p}Note`).textContent = 'Copied. Now click A1 in Excel and paste.'; }
    catch (e) { $(`#${p}Note`).innerHTML = 'Your browser blocked copying. Select all the text below, copy it, and paste at A1:<br><textarea rows="5">' + esc(tsv) + '</textarea>'; }
  };
  $(`#${p}Xlsx`).onclick = async () => { if (!await download(makeXlsx(block(), sheet), file + '.xlsx')) $(`#${p}Note`).textContent = 'Download blocked here; use Copy data instead.'; };
  $(`#${p}Csv`).onclick = async () => { if (!await download(new Blob([makeCsv(block())], { type: 'text/csv' }), file + '.csv')) $(`#${p}Note`).textContent = 'Download blocked here; use Copy data instead.'; };
  try { const n = localStorage.getItem('ewb-name'); if (n) $(`#${p}Name`).value = n; } catch (e) {}
  let last = null;
  async function run(f) {
    last = f;
    const out = $(`#${p}Out`), name = $(`#${p}Name`).value.trim();
    if (name) try { localStorage.setItem('ewb-name', name); } catch (e) {}
    if (/\.(csv|xls)$/i.test(f.name)) { out.innerHTML = '<div class="badbox">The pre-check needs your workbook as <b>.xlsx</b> (File → Create a Copy → Download a Copy in Excel for the web).</div>'; return; }
    let wb;
    try { wb = await readXlsx(await f.arrayBuffer()); }
    catch (e) { out.innerHTML = '<div class="badbox">I couldn\'t read this as an Excel workbook (.xlsx). Download it again from Excel with Download a Copy, and try that file.</div>'; return; }
    const res = check(wb, name), ids = new Set(res.tasks.map(t => t.id));
    let h = res.done ? `<div class="okbox">${okText}</div>` : `<p><b>${res.tasks.filter(t => t.ok).length} of ${tasks.length}</b> checks pass. Fix the items marked ✗, download a new copy, and check again.</p>`;
    h += '<ul class="tasks">';
    // one row per checklist item, in the checklist's order (with group headings for Capstone 2)
    for (const row of tasks) {
      const [id, title, , group] = row, t = res.tasks.find(x => x.id === id);
      if (group) h += `<li class="grp">${esc(group)}</li>`;
      if (t) h += `<li class="${t.ok ? 'ok' : 'no'}"><span class="ic">${t.ok ? '✓' : '✗'}</span><div><div class="t">${esc(t.title)}</div><div class="m">${esc(t.msg)}</div>${t.ok ? '' : `<a href="#" data-golesson="${t.module}">Review Module ${t.module}</a>`}</div></li>`;
      else h += `<li class="skip"><span class="ic">–</span><div><div class="t">${esc(title)}</div><div class="m">Checked once there is a chart.</div></div></li>`;
    }
    out.innerHTML = h + '</ul>';
  }
  $(`#${p}File`).onchange = e => { if (e.target.files[0]) run(e.target.files[0]); e.target.value = ''; };
  $(`#${p}Name`).addEventListener('change', () => { if (last) run(last); });
  const dz = $(`#${p}Drop`);
  ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('over'); }));
  dz.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) run(f); });
}
setupCapstone('ka', kaBlock, checkKeyAssessment, KA_TASKS, 'KeyAssessment_MassVolume',
  '<b>Your plot has all six elements.</b> Now copy it into Word (or a PDF) and attach both files to the Blackboard assignment.', 'Mass vs Volume');
setupCapstone('ka2', ka2Block, checkCapstone2, KA2_TASKS, 'Capstone2_EthanolWater',
  '<b>Both plots and the table have every element.</b> Now copy them into Word (or a PDF) and attach both files to the Blackboard assignment.', 'Ethanol-water');
