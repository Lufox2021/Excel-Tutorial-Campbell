/* Excel Workbench — Labs B (Boyle's law plot) and C (Beer's law calibration), lab detection, task lists. */

// ---------- task lists (titles shown by the checker; order = display order) ----------
const LAB_TASKS = {
  A: [['name', 'Your name is in cell B1', 1], ['data', 'All 15 measurements are in the table', 2], ['together', 'Rows stayed together when you sorted', 3],
    ['sorted', 'Sorted by Sample, then by Trial', 3], ['moles', 'NaOH (mol) calculated with a fixed reference to B2', 4], ['molfmt', 'NaOH (mol) shown in scientific notation, 3 decimals', 6],
    ['table', 'Results table in F3:M8 (Sample, Mean, SD, %CV, Median, Max, Min, n), column E empty', 5], ['mean', 'Mean of each sample (G4:G8), by formula', 5],
    ['sd', 'Standard deviation of each sample (H4:H8), using STDEV.S', 5], ['cv', '%CV of each sample (I4:I8), by formula', 5],
    ['mmm', 'Median, Max and Min of each sample (J4:L8), by formula', 5], ['n', 'n (number of trials) of each sample (M4:M8), by COUNT', 5],
    ['statfmt', 'Results formatted: scientific 3 decimals; %CV 3 decimals; n whole number', 6]],
  B: [['name', 'Your name is in cell B1', 1], ['data', 'Data in A3:C11, unchanged', 2], ['chart', 'A chart on the same sheet as the data', 7],
    ['scatter', 'Chart type is Scatter (XY), not Line', 7], ['series', 'Two series: x = Volume (A4:A11), y = each pressure column', 7],
    ['names', 'Series named from the headers (B3, C3)', 8], ['markers', 'Data shown as markers only, no connecting lines', 11],
    ['title', 'Chart title', 8], ['axtitles', 'Axis titles with units', 8], ['legend', 'Legend shown', 8], ['fonts', 'Title, axis titles and legend at 12 pt', 9]],
  C: [['name', 'Your name is in cell B1', 1], ['data', 'Standards (A3:B9) and unknowns (D3:E6) unchanged', 2], ['chart', 'A Scatter chart of the standards', 7],
    ['series', 'One series: x = Concentration (A4:A9), y = Absorbance (B4:B9)', 7], ['markers', 'Data shown as markers only, no connecting lines', 11],
    ['labels', 'Chart title and axis titles; no legend', 8], ['fonts', 'Title and axis titles at 12 pt', 9],
    ['trend', 'Linear trendline showing the equation and R²', 10], ['fit', 'Slope, Intercept and R² in I4:I6 by formula', 10],
    ['unknowns', 'Unknown concentrations in F4:F6 by formula', 10], ['fmt', 'Formats: scientific 3 decimals; R² 4 decimals', 6]]
};
function taskMaker(lab) {
  const T = [];
  const task = id => { const row = LAB_TASKS[lab].find(t => t[0] === id); const t = { id, title: row[1], module: row[2], ok: false, msg: '' }; T.push(t); return t; };
  return { T, task };
}
const cellText = c => c && c.v != null ? String(c.v).trim() : '';
// 'Lab B'!$A$4:$A$11 -> {col:'A', r1:4, r2:11, text:'A4:A11'} (single column ranges only)
function parseRef(f) {
  if (!f) return null;
  const m = /!?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(f.replace(/^.*!/, '!'));
  if (!m) return null;
  return { col: m[1], col2: m[3] || m[1], r1: +m[2], r2: +(m[4] || m[2]), text: m[3] ? `${m[1]}${m[2]}:${m[3]}${m[4]}` : `${m[1]}${m[2]}` };
}
const pt = sz => (sz / 100).toString().replace(/\.0$/, '') + ' pt';
function sizeProblem(label, sizes, want = 1200) {
  if (!sizes || !sizes.length) return `${label} is still at Excel's default size`;
  const off = [...new Set(sizes.filter(s => s !== want))];
  return off.length ? `${label} is ${off.map(pt).join('/')}` : '';
}
function findChartSheet(wbk, headerTest) {
  for (const sh of wbk.sheets) for (let r = 1; r <= 6; r++) if (headerTest(sh.cells, r)) return { sheet: sh, hr: r };
  return { sheet: wbk.sheets[0] || { cells: {}, charts: [] }, hr: 0 };
}
function allCharts(wbk, sheet) {
  const own = (sheet.charts || []).filter(c => !c.error);
  if (own.length) return { charts: own, elsewhere: false };
  const other = wbk.sheets.flatMap(s => s.charts || []).filter(c => !c.error);
  return { charts: other, elsewhere: other.length > 0 };
}
function chartKind(types) {
  const n = { scatterChart: 'Scatter', lineChart: 'Line', barChart: 'Bar/Column', areaChart: 'Area', pieChart: 'Pie', bubbleChart: 'Bubble', radarChart: 'Radar' };
  return types.map(t => n[t] || t.replace(/Chart$/, '')).join(' + ');
}
// least squares y = m x + b
function linfit(x, y) {
  const n = x.length, mx = mean(x), my = mean(y);
  let sxx = 0, syy = 0, sxy = 0;
  for (let i = 0; i < n; i++) { sxx += (x[i] - mx) ** 2; syy += (y[i] - my) ** 2; sxy += (x[i] - mx) * (y[i] - my); }
  const m = sxy / sxx;
  return { m, b: my - m * mx, r2: sxy * sxy / (sxx * syy), mSwap: sxy / syy, bSwap: mx - (sxy / syy) * my };
}

// ---------- Lab B: Boyle's law at two temperatures ----------
const LAB_B_V = [0.50, 0.75, 1.00, 1.25, 1.50, 2.00, 3.00, 4.00];
const LAB_B_T = [298, 348];
function labBData(name) {
  const r = mulberry32(cyrb53(normName(name) + '|labB') >>> 0);
  const n = 0.040 + r() * 0.040;
  const rows = LAB_B_V.map(v => ({ v, p: LAB_B_T.map(T => Math.round(n * 0.08206 * T / v * (1 + 0.004 * gauss(r)) * 1000) / 1000) }));
  return { name, n, rows };
}
function labBBlock(d) {
  const out = [['Name', ''], ['', ''], ['Volume (L)', `P at ${LAB_B_T[0]} K (atm)`, `P at ${LAB_B_T[1]} K (atm)`]];
  d.rows.forEach(x => out.push([x.v.toFixed(2), x.p[0].toFixed(3), x.p[1].toFixed(3)]));
  return out;
}
function checkLabB(wbk) {
  const { T, task } = taskMaker('B');
  const { sheet } = findChartSheet(wbk, (c, r) => /^volume\s*\(l\)/i.test(cellText(c['A' + r])));
  const cells = sheet.cells;
  const tName = task('name'), name = cellText(cells.B1);
  if (!name || /^name$/i.test(name)) { tName.msg = 'Type your name in B1, exactly as you typed it on the Lab page.'; return { lab: 'B', tasks: T, name: '', done: false }; }
  tName.ok = true; tName.msg = `Checking the data set for "${name}".`;
  const d = labBData(name);
  const tData = task('data');
  const hdr = labBBlock(d)[2], badH = hdr.map((h, j) => cellText(cells[addr(j + 1, 3)]) === h ? null : addr(j + 1, 3)).filter(Boolean);
  const bad = [];
  d.rows.forEach((x, i) => [x.v, x.p[0], x.p[1]].forEach((e, j) => { const c = cells[addr(j + 1, 4 + i)]; if (!c || !near(c.v, e)) bad.push(addr(j + 1, 4 + i)); }));
  if (bad.length === 24) { tData.msg = `A4:C11 doesn't hold the data set for "${name}". Paste the data from the Lab page at A1, and make sure B1 has exactly the name you typed there.`; return { lab: 'B', tasks: T, name, done: false }; }
  if (badH.length) tData.msg = `Header ${fmtList(badH)} was changed. Row 3 should read: ${hdr.join(' | ')}.`;
  else if (bad.length) tData.msg = `${fmtList(bad.slice(0, 6))}${bad.length > 6 ? ' and others' : ''} ${bad.length > 1 ? "don't" : "doesn't"} match the Lab data. Paste the data again at A1.`;
  else { tData.ok = true; tData.msg = '24 values and 3 headers in place.'; }

  const tChart = task('chart'), { charts, elsewhere } = allCharts(wbk, sheet);
  if (!charts.length) { tChart.msg = 'No chart found. Select A3:C11, then Insert → Scatter (markers only).'; return { lab: 'B', tasks: T, name, done: false }; }
  // prefer a scatter chart with the most series
  const ch = charts.slice().sort((a, b) => (b.types.includes('scatterChart') - a.types.includes('scatterChart')) || (b.series.length - a.series.length))[0];
  tChart.ok = true; tChart.msg = elsewhere ? "Found a chart (on a different sheet; that's fine)." : (charts.length > 1 ? `Found ${charts.length} charts; checking the Scatter chart.` : 'Found it.');
  const tSc = task('scatter');
  if (ch.types.length === 1 && ch.types[0] === 'scatterChart') { tSc.ok = true; tSc.msg = 'Scatter (XY).'; }
  else tSc.msg = `This is a ${chartKind(ch.types)} chart. ${ch.types.includes('lineChart') ? 'A Line chart spaces the volumes evenly as labels (0.50, 0.75, … 4.00 all one step apart), which bends the curve. ' : ''}Delete it, select A3:C11 and use Insert → Scatter.`;

  const tSer = task('series'), sMsg = [];
  const want = [{ col: 'B', name: hdr[1], ref: 'B3' }, { col: 'C', name: hdr[2], ref: 'C3' }];
  const ser = ch.series, matched = [];
  const swappedCols = [];
  for (const w of want) {
    const s = ser.find(q => { const y = parseRef(q.y && q.y.f); return y && y.col === w.col; });
    if (!s) {
      const swapped = ser.find(q => { const x = parseRef(q.x && q.x.f); return x && x.col === w.col; });
      if (swapped) { swappedCols.push(w.col); matched.push({ s: swapped, w, swapped: true }); }
      else sMsg.push(`No series plots column ${w.col} (${w.name}). Select Data → Add a series with X values A4:A11 and Y values ${w.col}4:${w.col}11.`);
      continue;
    }
    matched.push({ s, w });
    const y = parseRef(s.y.f), x = s.x && parseRef(s.x.f);
    if (!x) sMsg.push(`The ${w.name} series has no x values, so Excel plotted it against 1, 2, 3, …. Select Data → Edit it and set X values to A4:A11.`);
    else if (x.col !== 'A') sMsg.push(`The ${w.name} series uses ${x.text} for x; it should be A4:A11 (Volume).`);
    else if (x.r1 !== 4 || x.r2 !== 11) sMsg.push(`The ${w.name} series' x values are ${x.text}; they should be A4:A11${x.r1 === 3 ? ' (the header row is not data)' : ''}.`);
    if (y.r1 !== 4 || y.r2 !== 11) sMsg.push(`The ${w.name} series' y values are ${y.text}; they should be ${w.col}4:${w.col}11${y.r1 === 3 ? ' (the header row is not data)' : ''}.`);
  }
  if (swappedCols.length) sMsg.unshift(`The axes are swapped: pressure (${swappedCols.join(' and ')}) is on the x axis and volume on y. Volume (A) goes on x: in Select Data, edit ${swappedCols.length > 1 ? 'each series' : 'the series'} and swap the X and Y ranges.`);
  const empty = ser.filter(q => !q.y && !matched.some(m => m.s === q));
  if (empty.length) sMsg.push(`${empty.length > 1 ? 'Some series have' : 'A series has'} no data. Remove ${empty.length > 1 ? 'them' : 'it'} in Select Data.`);
  const extra = ser.filter(q => q.y && !matched.some(m => m.s === q));
  if (extra.length) {
    const volAsSeries = extra.some(q => { const y = parseRef(q.y && q.y.f); return y && y.col === 'A'; });
    sMsg.push(volAsSeries ? 'Volume (column A) is plotted as its own series. It belongs on the x axis: this happens when the chart isn\'t a Scatter chart, or when A isn\'t used as the X values.' : `There ${extra.length > 1 ? 'are' : 'is'} ${extra.length} extra series. Remove ${extra.length > 1 ? 'them' : 'it'} in Select Data.`);
  }
  if (!sMsg.length) { tSer.ok = true; tSer.msg = 'Both pressure columns against volume.'; } else tSer.msg = sMsg.slice(0, 3).join(' ');

  const tNames = task('names'), nMsg = [];
  for (const { s, w, swapped } of matched) {
    if (swapped) continue;
    const ref = s.nameRef && parseRef(s.nameRef);
    const ok = (ref && ref.text === w.ref) || (s.name && s.name.trim().toLowerCase() === w.name.toLowerCase());
    if (!ok) nMsg.push(s.name || s.nameRef ? `The ${w.col}-column series is named "${s.name || s.nameRef.replace(/^.*!/, '')}"; name it from ${w.ref} ("${w.name}").` : `The ${w.col}-column series has no name, so the legend says "Series${w.col === 'B' ? 1 : 2}". Name it from ${w.ref} ("${w.name}").`);
  }
  if (!matched.length || swappedCols.length) tNames.msg = 'Fix the series first.';
  else if (!nMsg.length) { tNames.ok = true; tNames.msg = 'The legend reads the column headers.'; } else tNames.msg = nMsg.join(' ') + ' (Select Data → Edit → Series name: click the header cell. Including row 3 when you select the data does this for you.)';

  const tMk = task('markers');
  const lined = ser.filter(q => !q.noLine), nomk = ser.filter(q => q.noMarker);
  if (lined.length) tMk.msg = `${lined.length === ser.length ? 'The points are' : 'Some points are'} joined by lines. For measured data, show markers only: pick the Scatter type with markers only, or Format the series → Line → No line.`;
  else if (nomk.length) tMk.msg = 'The markers are turned off, so the data points are invisible. Format the series → Marker → Automatic.';
  else { tMk.ok = true; tMk.msg = 'Markers, no lines.'; }

  labelTasks(task, ch, { title: true, x: [/volume/i, 'Volume (L)'], y: [/pressure/i, 'Pressure (atm)'], legend: true });
  const done = T.every(t => t.ok);
  return { lab: 'B', tasks: T, name, done, code: done ? completionCode(name, 'B') : null };
}
// shared chart-label checks for Labs B and C
function labelTasks(task, ch, o) {
  const isPlaceholder = t => !t || /^(chart|axis) title$/i.test(t);
  const tTitle = task(o.legend ? 'title' : 'labels'), msgs = [];
  if (!ch.title) msgs.push('There\'s no chart title. Add one that says what the graph shows (Chart → Chart Title, or the + / Chart Elements button).');
  else if (isPlaceholder(ch.title.text)) msgs.push(`The chart title still says "${ch.title.text || 'Chart Title'}". Click it and type a real title.`);
  let axMsgs = [];
  for (const [k, label] of [['x', 'horizontal (x)'], ['y', 'vertical (y)']]) {
    const ax = ch[k], [re, example] = o[k];
    if (!ax || !ax.title) axMsgs.push(`The ${label} axis has no title. Add Axis Titles and type "${example}".`);
    else if (isPlaceholder(ax.title.text)) axMsgs.push(`The ${label} axis title still says "Axis Title"; type "${example}".`);
    else if (!re.test(ax.title.text)) axMsgs.push(`The ${label} axis title "${ax.title.text}" should name the quantity: e.g. "${example}".`);
    else if (o[k][2] !== false && !/\(.+\)/.test(ax.title.text)) axMsgs.push(`Add the unit in parentheses to the ${label} axis title: "${example}".`);
  }
  if (o.legend) {
    if (!msgs.length) { tTitle.ok = true; tTitle.msg = `"${ch.title.text}"`; } else tTitle.msg = msgs.join(' ');
    const tAx = task('axtitles');
    if (!axMsgs.length) { tAx.ok = true; tAx.msg = `"${ch.x.title.text}" and "${ch.y.title.text}".`; } else tAx.msg = axMsgs.join(' ');
    const tLeg = task('legend');
    if (ch.legend) { tLeg.ok = true; tLeg.msg = 'Shown, so readers can tell the two temperatures apart.'; } else tLeg.msg = 'Turn the legend on (Chart Elements → Legend): with two series, readers need it to tell them apart.';
  } else {
    const all = msgs.concat(axMsgs);
    if (ch.legend) all.push('Turn the legend off: with only one series it adds nothing.');
    if (!all.length) { tTitle.ok = true; tTitle.msg = `"${ch.title.text}", "${ch.x.title.text}", "${ch.y.title.text}".`; } else tTitle.msg = all.slice(0, 3).join(' ');
  }
  const tF = task('fonts'), fm = [];
  if (ch.title) fm.push(sizeProblem('the title', ch.title.sizes));
  if (ch.x && ch.x.title) fm.push(sizeProblem('the x-axis title', ch.x.title.sizes));
  if (ch.y && ch.y.title) fm.push(sizeProblem('the y-axis title', ch.y.title.sizes));
  if (o.legend && ch.legend) fm.push(sizeProblem('the legend', ch.legend.sizes));
  const fp = fm.filter(Boolean);
  const missing = !ch.title || !(ch.x && ch.x.title) || !(ch.y && ch.y.title) || (o.legend && !ch.legend);
  if (!fp.length && !missing) { tF.ok = true; tF.msg = 'All 12 pt.'; }
  else if (!fp.length) tF.msg = 'Add the missing labels first; then their size is checked.';
  else tF.msg = fp.join('; ').replace(/^./, c => c.toUpperCase()) + '. Select each one and set the font size to 12 (Home → Font Size).';
}

// ---------- Lab C: Beer's law calibration ----------
const LAB_C_STD = [0, 1, 2, 3, 4, 5].map(k => k * 1e-5);
function labCData(name) {
  const r = mulberry32(cyrb53(normName(name) + '|labC') >>> 0);
  const eps = 8000 + r() * 7000, b0 = 0.002 + r() * 0.006;
  const absorb = c => Math.max(0, Math.round((eps * c + b0 + 0.004 * gauss(r)) * 1000) / 1000);
  const stds = LAB_C_STD.map(c => ({ c, a: absorb(c) }));
  const unk = [1, 2, 3].map(k => { const c = (0.8 + r() * 3.8) * 1e-5; return { label: 'U' + k, a: absorb(c) }; });
  const fit = linfit(stds.map(s => s.c), stds.map(s => s.a));
  unk.forEach(u => u.c = (u.a - fit.b) / fit.m);
  return { name, eps, stds, unk, fit };
}
function labCBlock(d) {
  const out = [['Name', ''], ['', ''], ['Concentration (M)', 'Absorbance', '', 'Unknown', 'Absorbance']];
  d.stds.forEach((s, i) => out.push([s.c.toExponential(2).toUpperCase().replace(/E([+-])(\d)$/, 'E$10$2'), s.a.toFixed(3), '', d.unk[i] ? d.unk[i].label : '', d.unk[i] ? d.unk[i].a.toFixed(3) : '']));
  return out;
}
function checkLabC(wbk) {
  const { T, task } = taskMaker('C');
  const { sheet } = findChartSheet(wbk, (c, r) => /^concentration/i.test(cellText(c['A' + r])) && /^absorbance/i.test(cellText(c['B' + r])));
  const cells = sheet.cells;
  const tName = task('name'), name = cellText(cells.B1);
  if (!name || /^name$/i.test(name)) { tName.msg = 'Type your name in B1, exactly as you typed it on the Lab page.'; return { lab: 'C', tasks: T, name: '', done: false }; }
  tName.ok = true; tName.msg = `Checking the data set for "${name}".`;
  const d = labCData(name);
  const tData = task('data'), bad = [];
  d.stds.forEach((s, i) => { if (!cells['A' + (4 + i)] || !near(cells['A' + (4 + i)].v, s.c)) bad.push('A' + (4 + i)); if (!cells['B' + (4 + i)] || !near(cells['B' + (4 + i)].v, s.a)) bad.push('B' + (4 + i)); });
  d.unk.forEach((u, i) => { if (!cells['E' + (4 + i)] || !near(cells['E' + (4 + i)].v, u.a)) bad.push('E' + (4 + i)); });
  if (bad.length >= 12) { tData.msg = `This doesn't hold the data set for "${name}". Paste the data from the Lab page at A1, and make sure B1 has exactly the name you typed there.`; return { lab: 'C', tasks: T, name, done: false }; }
  if (bad.length) tData.msg = `${fmtList(bad.slice(0, 6))} ${bad.length > 1 ? "don't" : "doesn't"} match the Lab data. Paste the data again at A1.`;
  else { tData.ok = true; tData.msg = '6 standards and 3 unknowns in place.'; }

  const tChart = task('chart'), { charts } = allCharts(wbk, sheet);
  if (!charts.length) { tChart.msg = 'No chart found. Select A3:B9, then Insert → Scatter (markers only).'; return { lab: 'C', tasks: T, name, done: false }; }
  const ch = charts.slice().sort((a, b) => b.types.includes('scatterChart') - a.types.includes('scatterChart'))[0];
  if (ch.types.length === 1 && ch.types[0] === 'scatterChart') { tChart.ok = true; tChart.msg = 'Scatter (XY).'; }
  else tChart.msg = `This is a ${chartKind(ch.types)} chart. A calibration curve needs a Scatter chart, so the concentrations sit at their true positions on the x axis. Select A3:B9 → Insert → Scatter.`;

  const tSer = task('series'), sMsg = [];
  const ser = ch.series, s0 = ser.find(q => { const y = parseRef(q.y && q.y.f); return y && y.col === 'B'; });
  if (!s0) {
    const sw = ser.find(q => { const x = parseRef(q.x && q.x.f); return x && x.col === 'B'; });
    sMsg.push(sw ? 'Absorbance is on the x axis. Concentration (A) goes on x and absorbance (B) on y: in Select Data, swap the X and Y ranges. (With the axes swapped, the slope is upside down.)' : 'No series plots Absorbance (B4:B9) against Concentration (A4:A9).');
  } else {
    const x = s0.x && parseRef(s0.x.f), y = parseRef(s0.y.f);
    if (!x) sMsg.push('The series has no x values, so Excel plotted it against 1, 2, 3, …. Select Data → Edit → X values = A4:A9.');
    else if (x.col !== 'A' || x.r1 !== 4 || x.r2 !== 9) sMsg.push(`X values are ${x.text}; they should be A4:A9${x.r1 === 3 ? ' (the header row is not data)' : ''}.`);
    if (y.r1 !== 4 || y.r2 !== 9) sMsg.push(`Y values are ${y.text}; they should be B4:B9${y.r1 === 3 ? ' (the header row is not data)' : ''}.`);
  }
  const others = s0 ? ser.filter(q => q !== s0) : [];
  if (others.length) {
    const unkPlotted = others.some(q => { const y = parseRef(q.y && q.y.f), x = parseRef(q.x && q.x.f); return (y && /[DEF]/.test(y.col)) || (x && /[DEF]/.test(x.col)); });
    sMsg.push(unkPlotted ? 'The unknowns are plotted too. The calibration curve uses the standards only; you find the unknowns from the line, in F4:F6.' : 'Only one series belongs on this chart: remove the extra one in Select Data.');
  }
  if (!sMsg.length) { tSer.ok = true; tSer.msg = 'Absorbance against concentration, standards only.'; } else tSer.msg = sMsg.join(' ');

  const tMk = task('markers');
  if (!s0) tMk.msg = 'Fix the series first.';
  else if (!s0.noLine) tMk.msg = 'The standards are joined by lines. Show markers only; the trendline is the only line a calibration curve needs. Format the series → Line → No line.';
  else if (s0.noMarker) tMk.msg = 'The markers are turned off. Format the series → Marker → Automatic.';
  else { tMk.ok = true; tMk.msg = 'Markers, no lines.'; }

  labelTasks(task, ch, { x: [/conc/i, 'Concentration (M)'], y: [/absorb/i, 'Absorbance', false], legend: false });

  const tTr = task('trend');
  const tl = s0 && s0.trend;
  if (!s0) tTr.msg = 'Fix the series first.';
  else if (!tl) tTr.msg = 'Add a trendline: click a data point → Add Trendline → Linear, then check Display Equation on chart and Display R-squared value on chart.';
  else if (tl.type !== 'linear') tTr.msg = `The trendline is ${tl.type}. Beer's law is a straight line: choose Linear.`;
  else if (tl.forced) tTr.msg = 'The trendline is forced through zero (Set Intercept is checked). Uncheck it: let the fit find its own intercept, as SLOPE and INTERCEPT do.';
  else if (!tl.eq || !tl.r2) tTr.msg = `Show ${!tl.eq && !tl.r2 ? 'the equation and R²' : !tl.eq ? 'the equation' : 'R²'} on the chart: Format Trendline → ${!tl.eq ? 'Display Equation on chart' : ''}${!tl.eq && !tl.r2 ? ' and ' : ''}${!tl.r2 ? 'Display R-squared value on chart' : ''}.`;
  else { tTr.ok = true; tTr.msg = 'Linear, with the equation and R² shown.'; }

  // fit cells I4:I6 with labels in H4:H6
  const tFit = task('fit'), f = d.fit, fMsg = [];
  const lab = ['H4', 'H5', 'H6'].map(r => cellText(cells[r]));
  if (!/slope/i.test(lab[0]) || !/intercept/i.test(lab[1]) || !/r\s*(²|\^?2)|r-?sq/i.test(lab[2])) fMsg.push('Label H4:H6 as Slope, Intercept, R².');
  const fitCell = (ref, exp, what, alts) => {
    const c = cells[ref];
    if (!c || c.v == null || c.v === '') return `${ref} is empty: use =${what}(B4:B9,A4:A9).`;
    if (typeof c.v !== 'number') return `${ref} shows "${c.v}", not a number.`;
    if (near(c.v, exp)) return c.formula ? '' : `${ref} is a typed number; use =${what}(B4:B9,A4:A9).`;
    for (const [v, m] of alts) if (v != null && near(c.v, v)) return `${ref}: ${m}`;
    if (Math.abs(c.v - exp) <= 0.01 * Math.abs(exp)) return `${ref} is close but not exact: it looks copied from the chart's equation. Use =${what}(B4:B9,A4:A9).`;
    return `${ref} isn't the ${what.toLowerCase()} of the standards: =${what}(B4:B9,A4:A9).`;
  };
  const swapMsg = 'the ranges are swapped. The order is known y\'s (absorbance, B4:B9) first, then known x\'s (concentration, A4:A9).';
  fMsg.push(fitCell('I4', f.m, 'SLOPE', [[f.mSwap, swapMsg]]), fitCell('I5', f.b, 'INTERCEPT', [[f.bSwap, swapMsg]]),
    fitCell('I6', f.r2, 'RSQ', [[Math.sqrt(f.r2), 'this is r (CORREL). R² is its square: use =RSQ(B4:B9,A4:A9).']]));
  const fm = fMsg.filter(Boolean);
  if (!fm.length) { tFit.ok = true; tFit.msg = 'Slope, intercept and R² match the standards.'; } else tFit.msg = fm.slice(0, 3).join(' ');

  const tU = task('unknowns'), uMsg = [];
  if (!/conc/i.test(cellText(cells.F3))) uMsg.push('Put the header "Concentration (M)" in F3.');
  const uc = d.unk.map((u, i) => ({ u, ref: 'F' + (4 + i), c: cells['F' + (4 + i)] }));
  const K = {};
  const add = (k, ref) => (K[k] = K[k] || []).push(ref);
  uc.forEach(({ u, ref, c }) => {
    if (!c || c.v == null || c.v === '') return add('empty', ref);
    if (typeof c.v !== 'number') return add('notnum', ref);
    if (near(c.v, u.c)) { if (!c.formula) add('typed', ref); else if (/\d\.\d{3,}/.test(c.ftext) && !/\$?I\$?[45]\b/i.test(c.ftext)) add('chartnums', ref); return; }
    if (!c.formula) return add(Math.abs(c.v - u.c) <= 0.002 * Math.abs(u.c) ? 'typed' : 'other', ref);
    if (!tFit.ok && /\$?I\$?[45]\b/i.test(c.ftext)) return add('fitfirst', ref);
    if (near(c.v, u.a / f.m)) return add('noint', ref);
    if (near(c.v, (u.a + f.b) / f.m)) return add('plusint', ref);
    if (near(c.v, f.m * u.a + f.b) || near(c.v, f.m * u.a)) return add('absorb', ref);
    if (Math.abs(c.v - u.c) <= 0.01 * Math.abs(u.c)) return add('chartnums', ref);
    add('other', ref);
  });
  const UMSG = {
    empty: 'empty.', notnum: 'showing an error or text instead of a number.',
    typed: 'typed numbers. Use a formula: =(E4-$I$5)/$I$4, filled down.',
    chartnums: 'using numbers copied from the chart\'s equation (rounded). Refer to the slope and intercept cells instead: =(E4-$I$5)/$I$4.',
    fitfirst: 'calculated from the slope or intercept in I4:I5, which aren\'t right yet. Fix those first.',
    noint: 'leaving out the intercept: c = (A − b) ÷ m.', plusint: 'adding the intercept; subtract it: c = (A − b) ÷ m.',
    absorb: 'calculating an absorbance, not a concentration. Solve A = m·c + b for c: c = (A − b) ÷ m.',
    other: 'not (A − b) ÷ m for their unknowns.'
  };
  const firstOK = uc[0].c && uc[0].c.formula && near(uc[0].c.v, d.unk[0].c);
  if (firstOK && K.other && K.other.length === 2 && !K.other.includes('F4')) uMsg.push('F4 is right but F5 and F6 are not. When you filled down, the slope and intercept references moved to I5, I6, …. Make them fixed: =(E4-$I$5)/$I$4.');
  else for (const k of Object.keys(UMSG)) if (K[k]) uMsg.push(`${fmtList(K[k])} ${K[k].length > 1 ? 'are' : 'is'} ${UMSG[k]}`);
  if (!uMsg.length) { tU.ok = true; tU.msg = 'All three unknowns right.'; } else tU.msg = uMsg.slice(0, 3).join(' ');

  const tFmt = task('fmt'), off = [];
  for (const r of ['I4', 'I5', 'F4', 'F5', 'F6']) { const c = cells[r]; if (c && typeof c.v === 'number' && !(c.fmt.sci && c.fmt.dec === 3)) off.push(r); }
  { const c = cells.I6; if (c && typeof c.v === 'number' && !(!c.fmt.sci && !c.fmt.general && c.fmt.dec === 4)) off.push('I6'); }
  if (!['I4', 'I5', 'I6', 'F4', 'F5', 'F6'].every(r => cells[r] && typeof cells[r].v === 'number')) tFmt.msg = 'Finish I4:I6 and F4:F6 first; then their format is checked.';
  else if (!off.length) { tFmt.ok = true; tFmt.msg = 'Slope, intercept and concentrations like 1.234E+04; R² like 0.9987.'; }
  else tFmt.msg = `Check the number format of ${fmtList(off)}. Slope, intercept and concentrations: Scientific, 3 decimals. R²: Number, 4 decimals.`;

  const done = T.every(t => t.ok);
  return { lab: 'C', tasks: T, name, done, code: done ? completionCode(name, 'C') : null };
}

// ---------- which lab is this workbook? ----------
function detectLab(wbk) {
  for (const sh of wbk.sheets) if (/^volume\s*\(ml\)/i.test(cellText(sh.cells.A1)) && /^mass|water|ethanol/i.test(cellText(sh.cells.B1))) return 'K';
  for (const sh of wbk.sheets) for (let r = 1; r <= 8; r++) {
    const a = cellText(sh.cells['A' + r]).toLowerCase(), b = cellText(sh.cells['B' + r]).toLowerCase(), c = cellText(sh.cells['C' + r]).toLowerCase();
    if (a === 'sample' && b.startsWith('trial')) return 'A';
    if (a.startsWith('volume (l)') && b.startsWith('p at')) return 'B';
    if (a.startsWith('concentration') && b.startsWith('absorbance')) return 'C';
    if (c.startsWith('volume (ml)')) return 'A';
  }
  return null;
}
function checkWorkbook(wbk, lab) {
  lab = lab || detectLab(wbk);
  if (lab === 'A') { const r = checkLabA(wbk); r.lab = 'A'; return r; }
  if (lab === 'B') return checkLabB(wbk);
  if (lab === 'C') return checkLabC(wbk);
  if (lab === 'K') return { lab: 'K', tasks: [], name: '', done: false };
  return { lab: null, tasks: [], name: '', done: false };
}
