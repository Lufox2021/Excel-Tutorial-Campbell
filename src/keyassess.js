/* Campbell Key Assessment: Mass vs Volume plot (same data for every student). */
const KA = {
  V: [5.0, 10.0, 15.0, 20.0, 25.0, 30.0, 35.0, 40.0],
  M: [4.12, 7.94, 12.34, 16.79, 18.98, 24.36, 27.88, 31.01],   // ethanol, 0.789 g/mL, with 3% CV noise
  hdr: ['Volume (mL)', 'Mass (g)']
};
const kaBlock = () => [KA.hdr].concat(KA.V.map((v, i) => [v.toFixed(1), KA.M[i].toFixed(2)]));
const KA_TASKS = [['chart', 'A chart of the data'], ['title', '1. A descriptive title'], ['subtitle', '2. A subtitle with your name, in a smaller font'],
  ['xaxis', '3. x-axis label with units'], ['yaxis', '4. y-axis label with units'], ['data', '5. Data: mass vs. volume as an XY (Scatter) plot'],
  ['trend', '6. Regression line with its equation and R² value']];
const KA_DEFAULT_TITLE_SZ = 1400; // Excel's default chart-title size
const ptOf = s => (s / 100).toString() + ' pt';

// values of a chart range like 'Sheet1'!$A$2:$A$9 (null if it can't be read)
function rangeValues(wbk, homeSheet, f) {
  if (!f) return null;
  const m = /^(?:'?(.*?)'?!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(f.trim());
  if (!m) return null;
  const sh = (m[1] && wbk.sheets.find(s => s.name === m[1].replace(/''/g, "'"))) || homeSheet;
  const c1 = colToNum(m[2]), c2 = colToNum(m[4] || m[2]), r1 = +m[3], r2 = +(m[5] || m[3]), out = [];
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { const cell = sh.cells[addr(c, r)]; out.push(cell ? cell.v : null); }
  return { values: out, text: (m[4] ? `${m[2]}${m[3]}:${m[4]}${m[5]}` : `${m[2]}${m[3]}`) };
}
const sameSet = (got, want) => got && got.length === want.length && want.every(w => got.some(g => typeof g === 'number' && Math.abs(g - w) < 0.005));

// title, subtitle with name, axis titles: the elements both capstones share
const MASS_VOL_AXES = { x: [/vol/i, /mass/i, 'Volume (mL)'], y: [/mass/i, /vol/i, 'Mass (g)'] };
function plotElements(task, ch, studentName, eg, axes = MASS_VOL_AXES) {
  const placeholder = t => !t || /^(chart|axis) title$/i.test(t.trim());
  const lines = ch.title ? ch.title.lines : [];
  const nm = normName(studentName);
  const hasName = t => nm && normName(t).includes(nm);

  // 1. title: the first line, without the name if the name was typed on it
  const tT = task('title');
  const stripName = t => nm ? t.replace(new RegExp(studentName.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+'), 'i'), '').replace(/^[\s,–—:-]+|[\s,–—:-]+$/g, '') : t;
  const first = lines[0], titleText = first ? stripName(first.text) : '';
  if (!ch.title) tT.msg = `There's no chart title. Add one (Chart → Chart Title) that says what the plot shows, e.g. "${eg}".`;
  else if (!lines.length || placeholder(ch.title.text)) tT.msg = `The title still says "Chart Title". Click it and type a descriptive title, e.g. "${eg}".`;
  else if (!titleText) tT.msg = 'The first line of the title is only your name. Put a descriptive title on the first line and your name on a second, smaller line.';
  else if (titleText.split(/\s+/).length < 3) tT.msg = `"${titleText}" is too short to be descriptive. Say what is plotted, e.g. "${eg}".`;
  else { tT.ok = true; tT.msg = `"${titleText}"`; }

  // 2. subtitle: a later line of the title with the name, smaller than the first line
  const tS = task('subtitle');
  const subIdx = lines.findIndex((l, i) => i > 0 && hasName(l.text));
  if (!nm) tS.msg = 'Type your name above (as in Blackboard) so the subtitle can be checked.';
  else if (!ch.title || !lines.length || placeholder(ch.title.text)) tS.msg = 'Add the chart title first; your name goes on a second line of it.';
  else if (subIdx < 0 && hasName(lines[0].text) && lines.length > 1 && !stripName(lines[0].text)) tS.msg = 'Your name is above the title. Swap them: the descriptive title goes first, and your name on the second, smaller line.';
  else if (subIdx < 0 && hasName(lines[0].text)) tS.msg = 'Your name is on the same line as the title. Put it on its own line: click just before your name and press Enter, then make that line smaller.';
  else if (subIdx < 0) tS.msg = lines.length > 1 ? `The second line of the title, "${lines[1].text}", should include your name, "${studentName.trim()}", as typed above.` : 'Add a subtitle: click at the end of the title, press Enter, type your name, then make that line smaller (select it → Home → Font Size).';
  else {
    const tSizes = lines[0].sizes.map(x => x || KA_DEFAULT_TITLE_SZ), sSizes = lines[subIdx].sizes.map(x => x || KA_DEFAULT_TITLE_SZ);
    const tMin = Math.min(...tSizes), sMax = Math.max(...sSizes);
    if (!(sMax < tMin)) tS.msg = `Your name line is ${ptOf(sMax)} and the title is ${ptOf(tMin)}. Make the name line smaller: select just that line → Home → Font Size (e.g. 2–4 pt smaller than the title).`;
    else { tS.ok = true; tS.msg = `"${lines[subIdx].text}" at ${ptOf(sMax)}, under a ${ptOf(tMin)} title.`; }
  }

  // 3–4. axis titles
  const unit = t => /\(.+\)/.test(t) || /%/.test(t) || /\b(mL|g|L|cm3|cm³|kg|mg|g\/mL)\b/.test(t);
  const axisTask = (id, ax, label, re, other, example) => {
    const t = task(id);
    if (!ax || !ax.title) t.msg = `There's no ${label} axis title. Add Axis Titles and type "${example}".`;
    else if (placeholder(ax.title.text)) t.msg = `The ${label} axis title still says "Axis Title"; type "${example}".`;
    else if (!re.test(ax.title.text) && other.test(ax.title.text)) t.msg = `The ${label} axis title "${ax.title.text}" names the wrong quantity: the ${label} axis is ${example.split(' (')[0].toLowerCase()} (check the data too).`;
    else if (!re.test(ax.title.text)) t.msg = `The ${label} axis title "${ax.title.text}" should name the quantity: e.g. "${example}".`;
    else if (!unit(ax.title.text)) t.msg = `Add the unit to the ${label} axis title: "${example}".`;
    else { t.ok = true; t.msg = `"${ax.title.text}"`; }
  };
  axisTask('xaxis', ch.x, 'x', ...axes.x);
  axisTask('yaxis', ch.y, 'y', ...axes.y);

}

function checkKeyAssessment(wbk, studentName) {
  const T = [];
  const task = id => { const row = KA_TASKS.find(t => t[0] === id); const t = { id, title: row[1], module: 11, ok: false, msg: '' }; T.push(t); return t; };
  const tChart = task('chart');
  let home = wbk.sheets[0], ch = null;
  for (const sh of wbk.sheets) { const c = (sh.charts || []).find(x => !x.error); if (c) { home = sh; ch = c; break; } }
  if (!ch) { tChart.msg = 'No chart found in this workbook. Select the data (both columns, with headers) and Insert → Scatter.'; return { tasks: T, done: false }; }
  tChart.ok = true; tChart.msg = `Found it (on sheet "${home.name}").`;
  plotElements(task, ch, studentName, 'Mass vs. Volume of Ethanol');

  // 5. data
  const tD = task('data'), dm = [];
  if (!(ch.types.length === 1 && ch.types[0] === 'scatterChart')) dm.push(`This is a ${chartKind(ch.types)} chart. Use Scatter (XY)${ch.types.includes('lineChart') ? ': a Line chart treats the volumes as labels 1, 2, 3, …, so its regression equation is wrong' : ''}. Select the data → Insert → Scatter.`);
  const ser = ch.series.filter(s => s.y);
  if (!ser.length) dm.push('The chart has no data series.');
  else {
    if (ser.length > 1) dm.push(`The chart has ${ser.length} series; it should have one (mass vs. volume). Remove the others in Select Data.`);
    const s = ser[0], xv = s.x && rangeValues(wbk, home, s.x.f), yv = rangeValues(wbk, home, s.y.f);
    const xs = xv && xv.values.filter(v => typeof v === 'number'), ys = yv && yv.values.filter(v => typeof v === 'number');
    if (!s.x) dm.push('The series has no x values, so Excel plotted the masses against 1, 2, 3, …. Select Data → Edit → X values = the volume column.');
    else if ((xv && xv.values.some(v => typeof v === 'string')) || yv.values.some(v => typeof v === 'string')) dm.push(`The ${xv && xv.values.some(v => typeof v === 'string') ? 'x' : 'y'} range (${(xv && xv.values.some(v => typeof v === 'string') ? xv : yv).text}) includes the header row, so the points are shifted. In Select Data, start both ranges at row 2.`);
    else if (xs.length !== ys.length) dm.push(`The x range (${xv.text}) and y range (${yv.text}) are different lengths, so the points don't pair up. Use the same rows for both.`);
    else if (sameSet(xs, KA.M) && sameSet(ys, KA.V)) dm.push('The axes are swapped: mass is on x and volume on y. Mass vs. Volume means mass on the y axis: in Select Data, swap the X and Y ranges (and the axis titles).');
    else {
      if (!sameSet(xs, KA.V)) dm.push(`The x values (${xv ? xv.text : '?'}) aren't the 8 volumes from the data.`);
      if (!sameSet(ys, KA.M)) dm.push(`The y values (${yv ? yv.text : '?'}) aren't the 8 masses from the data.`);
      if (sameSet(xs, KA.V) && sameSet(ys, KA.M)) {
        const pairs = xs.map((x, i) => [x, ys[i]]), ok = pairs.every(([x, y]) => KA.V.some((v, k) => Math.abs(v - x) < 0.005 && Math.abs(KA.M[k] - y) < 0.005));
        if (!ok) dm.push('Some masses are no longer beside their volumes (was one column sorted by itself?). Paste the data again.');
      }
    }
  }
  if (!dm.length) { tD.ok = true; tD.msg = 'Mass (y) against volume (x), all 8 points.'; } else tD.msg = dm.slice(0, 3).join(' ');

  // 6. regression line
  const tR = task('trend'), tl = ser[0] && ser[0].trend;
  if (!ser.length) tR.msg = 'Plot the data first.';
  else if (!tl) tR.msg = 'Add the regression line: Chart → Trendline → Linear, then turn on Display Equation on chart and Display R-squared value on chart.';
  else if (tl.type !== 'linear') tR.msg = `The trendline is ${tl.type === 'poly' ? 'polynomial' : tl.type}. Mass is proportional to volume: choose Linear.`;
  else if (tl.forced) tR.msg = 'The line is forced through zero (Set Intercept is checked). Use standard linear regression: uncheck Set Intercept in Format Trendline.';
  else if (!tl.eq || !tl.r2) tR.msg = `Show ${!tl.eq && !tl.r2 ? 'the equation and R²' : !tl.eq ? 'the equation' : 'the R² value'} on the chart: Format Trendline → ${[!tl.eq && 'Display Equation on chart', !tl.r2 && 'Display R-squared value on chart'].filter(Boolean).join(' and ')}.`;
  else { tR.ok = true; tR.msg = 'Standard linear regression, with the equation and R² shown.'; }

  return { tasks: T, done: T.every(t => t.ok) };
}

/* ---------- Capstone 2: five ethanol–water solutions, two plots (same data for every student) ----------
   Densities at 25 °C by mass % ethanol (Engineering ToolBox; 25 and 75 % interpolated), 3% CV noise. */
const KA2 = {
  V: KA.V,
  sols: [
    { pct: 0, label: '0% ethanol (water)', rho: 0.997, M: [4.80, 9.56, 14.41, 20.87, 24.66, 29.34, 34.21, 39.97] },
    { pct: 25, label: '25% ethanol', rho: 0.959, M: [4.82, 8.99, 14.07, 19.34, 24.61, 29.14, 32.84, 38.28] },
    { pct: 50, label: '50% ethanol', rho: 0.910, M: [4.44, 9.03, 13.95, 18.50, 21.67, 26.81, 30.52, 37.75] },
    { pct: 75, label: '75% ethanol', rho: 0.851, M: [4.24, 8.65, 12.88, 17.58, 20.87, 27.25, 29.11, 34.15] },
    { pct: 100, label: '100% ethanol', rho: 0.785, M: [3.94, 7.65, 12.18, 15.97, 20.44, 22.73, 26.98, 32.23] }
  ]
};
KA2.sols.forEach(s => { s.fit = linfit(KA2.V, s.M); });
const ka2Block = () => [['Volume (mL)'].concat(KA2.sols.map(s => s.label))].concat(KA2.V.map((v, i) => [v.toFixed(1)].concat(KA2.sols.map(s => s.M[i].toFixed(2)))));
const KA2_TASKS = [
  ['p1chart', 'Plot 1 · mass vs. volume: a chart of the data', 12, 'Plot 1 · Mass vs. Volume'],
  ['p1title', '1. A descriptive title', 11], ['p1subtitle', '2. A subtitle with your name, in a smaller font', 11],
  ['p1xaxis', '3. x-axis label with units', 11], ['p1yaxis', '4. y-axis label with units', 11],
  ['p1data', '5. Data for all five solutions on one XY (Scatter) plot', 12], ['p1colors', '6. Each solution in its own color, with a legend', 12],
  ['p1trend', '7. A regression line for each solution', 12],
  ['table', '8. Table of % ethanol, slope (density), intercept and R² in H1:K6', 12, 'Table'],
  ['p2chart', 'Plot 2 · density vs. % ethanol: a second chart', 12, 'Plot 2 · Density vs. % Ethanol'],
  ['p2title', '9. A descriptive title', 11], ['p2subtitle', '10. A subtitle with your name, in a smaller font', 11],
  ['p2xaxis', '11. x-axis label with units', 11], ['p2yaxis', '12. y-axis label with units', 11],
  ['p2data', '13. Data: density (y) vs. % ethanol (x), from your table', 12]
];
// which solution (by % ethanol) does a label name?
function pctOf(text) {
  const t = String(text == null ? '' : text).toLowerCase();
  if (/water/.test(t) && !/ethanol/.test(t)) return 0;
  const m = /(\d+)\s*%/.exec(t);
  if (m && KA2.sols.some(s => s.pct === +m[1])) return +m[1];
  if (/ethanol/.test(t) && !/water/.test(t) && !m) return 100;
  return null;
}
const solByPct = p => KA2.sols.find(s => s.pct === p);
const PCT_AXES = { x: [/ethanol|%/i, /density/i, '% Ethanol (by mass)'], y: [/density/i, /ethanol|%/i, 'Density (g/mL)'] };

function checkCapstone2(wbk, studentName) {
  const T = [];
  const task = id => { const row = KA2_TASKS.find(t => t[0] === id); const t = { id, title: row[1], module: row[2], group: row[3] || null, ok: false, msg: '' }; T.push(t); return t; };
  const all = wbk.sheets.flatMap(sh => (sh.charts || []).filter(c => !c.error).map(c => ({ sh, c })));
  const ysOf = (sh, s) => { const r = s.y && rangeValues(wbk, sh, s.y.f); return r ? r.values.filter(v => typeof v === 'number') : []; };
  const xsOf = (sh, s) => { const r = s.x && rangeValues(wbk, sh, s.x.f); return r ? r.values.filter(v => typeof v === 'number') : []; };
  const pctSet = xs => sameSet(xs, KA2.sols.map(s => s.pct)) ? 1 : sameSet(xs.map(x => x * 100), KA2.sols.map(s => s.pct)) ? 100 : 0;
  const massScore = ({ sh, c }) => c.series.filter(s => { const y = ysOf(sh, s), x = xsOf(sh, s); return KA2.sols.some(q => sameSet(y, q.M) || sameSet(x, q.M)); }).length;
  const densScore = ({ sh, c }) => c.series.filter(s => { const y = ysOf(sh, s), x = xsOf(sh, s); return pctSet(x) || pctSet(y) || (y.length === 5 && y.every(v => v > 0.7 && v < 1.1)); }).length;
  const p1 = all.filter(o => massScore(o) > 0).sort((a, b) => massScore(b) - massScore(a))[0] || (all.length === 1 && !densScore(all[0]) ? all[0] : null);
  const p2 = all.filter(o => o !== p1 && densScore(o) > 0)[0] || all.find(o => o !== p1) || null;

  // ---------- Plot 1 ----------
  const t1 = task('p1chart');
  if (!p1) t1.msg = 'No mass-vs-volume chart found. Select A1:F9 (all six columns, with headers) and Insert → Scatter.';
  else {
    t1.ok = true; t1.msg = `Found it (on sheet "${p1.sh.name}").`;
    const { sh, c: ch } = p1;
    plotElements(id => task('p1' + id), ch, studentName, 'Mass vs. Volume of Ethanol–Water Solutions');
    const tD = task('p1data'), dm = [], matched = {}, extra = [];
    if (!(ch.types.length === 1 && ch.types[0] === 'scatterChart')) dm.push(`This is a ${chartKind(ch.types)} chart. Use Scatter (XY)${ch.types.includes('lineChart') ? ': a Line chart treats the volumes as labels 1, 2, 3, …, so its regressions are wrong' : ''}. Select A1:F9 → Insert → Scatter.`);
    for (const s of ch.series.filter(q => q.y)) {
      const xv = s.x && rangeValues(wbk, sh, s.x.f), yv = rangeValues(wbk, sh, s.y.f), xs = xsOf(sh, s), ys = ysOf(sh, s);
      const sol = KA2.sols.find(q => sameSet(ys, q.M)), solX = KA2.sols.find(q => sameSet(xs, q.M));
      if (sol && !matched[sol.pct]) {
        matched[sol.pct] = s;
        if (!s.x) dm.push(`The ${sol.label} series has no x values, so Excel plotted it against 1, 2, 3, …. Select Data → Edit → X values = A2:A9.`);
        else if (xv.values.some(v => typeof v === 'string') || yv.values.some(v => typeof v === 'string')) dm.push(`The ${sol.label} series includes the header row in its range, so its points are shifted. Start its ranges at row 2.`);
        else if (!sameSet(xs, KA2.V)) dm.push(`The ${sol.label} series uses ${xv.text} for x; it should be the volumes, A2:A9.`);
        else if (!xs.every((x, i) => KA2.V.some((v, k) => Math.abs(v - x) < 0.005 && Math.abs(sol.M[k] - ys[i]) < 0.005))) dm.push(`Some ${sol.label} masses are no longer beside their volumes. Paste the data again.`);
      } else if (solX && sameSet(ys, KA2.V) && !matched[solX.pct]) { matched[solX.pct] = s; dm.push(`The ${solX.label} series has its axes swapped (mass on x, volume on y). In Select Data, swap its X and Y ranges.`); }
      else extra.push(s);
    }
    const missing = KA2.sols.filter(q => !matched[q.pct]);
    if (missing.length) dm.push(`Not plotted: ${fmtList(missing.map(q => q.label))}. ${Object.keys(matched).length <= 1 ? 'Select all six columns (A1:F9) before inserting the chart, or u' : 'U'}se Select Data → Add: X values A2:A9, Y values that solution's column.`);
    if (extra.length) dm.push(`${extra.length} extra series ${extra.length > 1 ? "don't" : "doesn't"} match any solution's data (e.g. Volume plotted as a series). Remove ${extra.length > 1 ? 'them' : 'it'} in Select Data.`);
    if (!dm.length) { tD.ok = true; tD.msg = 'All five solutions, mass (y) against volume (x).'; } else tD.msg = dm.slice(0, 3).join(' ');

    const tC = task('p1colors'), cm = [], got = KA2.sols.filter(q => matched[q.pct]);
    if (got.length < 2) cm.push('Plot the solutions first.');
    else {
      const col = q => matched[q.pct].color || ('auto:' + matched[q.pct].idx);
      const groups = {};
      got.forEach(q => (groups[col(q)] = groups[col(q)] || []).push(q.label));
      Object.values(groups).filter(g => g.length > 1).forEach(g => cm.push(`${fmtList(g)} are the same color. Give each solution its own color: click a series → Format Data Series → Marker → Fill (and Border).`));
      if (!ch.legend) cm.push('Turn the legend on (Chart Elements → Legend) so readers can tell which color is which solution.');
      else {
        const bad = got.filter(q => { const s = matched[q.pct]; let n = s.name; if (!n && s.nameRef) { const r = rangeValues(wbk, sh, s.nameRef); n = r && r.values[0]; } return pctOf(n) !== q.pct; });
        if (bad.length) cm.push(`The legend doesn't name ${fmtList(bad.map(q => q.label))} (it shows "Series…" or another name). Select Data → Edit → Series name: click that solution's header cell in row 1.`);
      }
    }
    if (!cm.length) { tC.ok = true; tC.msg = `${got.length} colors, and the legend names each solution.`; } else tC.msg = cm.join(' ');

    const tR = task('p1trend'), rm = [];
    if (!got.length) rm.push('Plot the data first.');
    const noLine = got.filter(q => !matched[q.pct].trend);
    if (noLine.length) rm.push(`${fmtList(noLine.map(q => q.label))} ${noLine.length > 1 ? 'have' : 'has'} no regression line. Click that solution's points → Add Trendline → Linear (one solution at a time).`);
    for (const q of got) {
      const tl = matched[q.pct].trend; if (!tl) continue;
      if (tl.type !== 'linear') rm.push(`The ${q.label} trendline is ${tl.type === 'poly' ? 'polynomial' : tl.type}; choose Linear.`);
      else if (tl.forced) rm.push(`The ${q.label} line is forced through zero; uncheck Set Intercept (standard linear regression).`);
    }
    if (got.length && got.length < 5) rm.push('Plot all five solutions, each with its own line.');
    if (!rm.length) { tR.ok = true; tR.msg = 'Five linear regression lines.'; } else tR.msg = rm.join(' ');
  }

  // ---------- Table H1:K6 ----------
  const tT = task('table'), tm = [];
  const tsh = (p1 || p2 || { sh: wbk.sheets[0] }).sh, cells = tsh.cells, txt = r => cellText(cells[r]);
  if (!(/ethanol|%/i.test(txt('H1')) && /slope|density/i.test(txt('I1')) && /intercept/i.test(txt('J1')) && /r\s*(²|\^?2)|r-?sq/i.test(txt('K1'))))
    tm.push('Put the headers % Ethanol, Slope = Density (g/mL), Intercept (g), R² in H1:K1.');
  const seen = {};
  for (const r of [2, 3, 4, 5, 6]) {
    const hc = cells['H' + r];
    let p = null;
    if (hc && typeof hc.v === 'number') { const v = hc.fmt.pct || (hc.v > 0 && hc.v <= 1 && !Number.isInteger(hc.v)) ? Math.round(hc.v * 100) : hc.v; if (KA2.sols.some(s => s.pct === v)) p = v; }
    if (p === null) { tm.push(hc && typeof hc.v === 'string' && pctOf(hc.v) !== null ? `H${r} is text ("${hc.v}"); type just the number (e.g. ${pctOf(hc.v)}) so Plot 2 can use it.` : `H${r} should be one of the % ethanol values: 0, 25, 50, 75 or 100.`); continue; }
    if (seen[p]) { tm.push(`${p}% ethanol appears twice (H${seen[p]} and H${r}).`); continue; }
    seen[p] = r;
    const sol = solByPct(p), f = sol.fit, others = KA2.sols.filter(q => q !== sol);
    const one = (ref, exp, what, fn, swapV, otherV) => {
      const c = cells[ref];
      if (!c || c.v == null || c.v === '') return `${ref} is empty: =${fn}(…).`;
      if (typeof c.v !== 'number') return `${ref} shows "${c.v}", not a number.`;
      if (near(c.v, exp)) return c.formula ? '' : `${ref} is a typed number; use =${fn}(…).`;
      if (!c.formula && Math.abs(c.v - exp) <= 0.01 * Math.abs(exp) + 1e-4) return `${ref} is a typed, rounded number; use =${fn}(…) so the table keeps full precision.`;
      if (swapV != null && near(c.v, swapV)) return `${ref}: the ranges are swapped. ${fn}(known y's = masses, known x's = volumes $A$2:$A$9).`;
      if (fn === 'RSQ' && near(c.v, Math.sqrt(exp))) return `${ref} is r (CORREL); R² is its square: =RSQ(…).`;
      if (fn === 'SLOPE' && Math.abs(c.v - sol.rho) < 0.0015 && !near(c.v, exp)) return `${ref} looks like a textbook density; use your data's slope: =SLOPE(…).`;
      const o = others.find(q => near(c.v, otherV(q)));
      if (o) return `${ref} is the ${what} for ${o.label}, not ${sol.label}: check which mass column the formula uses.`;
      if (Math.abs(c.v - exp) <= 0.01 * Math.abs(exp) + 1e-4) return `${ref} is close but not exact: it looks copied from the chart. Use =${fn}(…) so the table keeps full precision.`;
      return `${ref} isn't the ${what} for ${sol.label}.`;
    };
    tm.push(one('I' + r, f.m, 'slope', 'SLOPE', f.mSwap, q => q.fit.m), one('J' + r, f.b, 'intercept', 'INTERCEPT', f.bSwap, q => q.fit.b), one('K' + r, f.r2, 'R²', 'RSQ', null, q => q.fit.r2));
  }
  const byRest = new Map();
  for (const m of tm.filter(Boolean)) {
    const g = /^([A-K]\d) ?(:)?(.*)$/.exec(m);
    if (!g) { byRest.set(m, null); continue; }
    const rest = g[3].trim();
    if (!byRest.has(rest)) byRest.set(rest, { refs: [], rest, colon: !!g[2] });
    byRest.get(rest).refs.push(g[1]);
  }
  const tmm = [...byRest].map(([k, v]) => !v ? k : v.refs.length === 1 ? `${v.refs[0]}${v.colon ? ':' : ''} ${v.rest}` : `${fmtList(v.refs)}${v.colon ? ':' : ''} ${v.colon ? v.rest : v.rest.replace(/^is (a |the )?/, 'are ').replace(/number;/, 'numbers;')}`);
  if (!tmm.length) { tT.ok = true; tT.msg = 'Slope (density), intercept and R² for all five solutions.'; } else tT.msg = tmm.slice(0, 3).join(' ') + (tmm.length > 3 ? ' (and more)' : '');

  // ---------- Plot 2 ----------
  const t2 = task('p2chart');
  if (!p2) t2.msg = 'No second chart found. Select H1:I6 (% ethanol and slope = density, with headers) and Insert → Scatter.';
  else {
    t2.ok = true; t2.msg = `Found it (on sheet "${p2.sh.name}").`;
    const { sh, c: ch } = p2;
    plotElements(id => task('p2' + id), ch, studentName, 'Density of Ethanol–Water Solutions vs. % Ethanol', PCT_AXES);
    const tD = task('p2data'), dm = [];
    if (!(ch.types.length === 1 && ch.types[0] === 'scatterChart')) dm.push(`This is a ${chartKind(ch.types)} chart. Use Scatter (XY), so the % ethanol values sit at their true positions. Select H1:I6 → Insert → Scatter.`);
    const ser = ch.series.filter(q => q.y);
    if (!ser.length) dm.push('The chart has no data series.');
    else {
      if (ser.length > 1) dm.push(`The chart has ${ser.length} series; it should have one (density vs. % ethanol). Remove the others in Select Data.`);
      const s = ser[0], xv = s.x && rangeValues(wbk, sh, s.x.f), yv = rangeValues(wbk, sh, s.y.f), xs = xsOf(sh, s), ys = ysOf(sh, s);
      const scale = pctSet(xs);
      if (!s.x) dm.push('The series has no x values, so Excel plotted the densities against 1, 2, 3, …. Select Data → Edit → X values = H2:H6.');
      else if (xv.values.some(v => typeof v === 'string') || yv.values.some(v => typeof v === 'string')) dm.push('A range includes the header row, so the points are shifted. In Select Data, start both ranges at row 2.');
      else if (pctSet(ys) && !scale) dm.push('The axes are swapped: % ethanol is on y. Density goes on y and % ethanol on x: in Select Data, swap the X and Y ranges.');
      else if (!scale) dm.push(`The x values (${xv.text}) aren't the % ethanol values 0, 25, 50, 75, 100 (H2:H6).`);
      else {
        const pairs = xs.map((x, i) => [Math.round(x * scale), ys[i]]);
        const slopesOK = pairs.every(([p, y]) => solByPct(p) && near(y, solByPct(p).fit.m));
        const textbook = pairs.every(([p, y]) => solByPct(p) && Math.abs(y - solByPct(p).rho) < 0.0015);
        if (!slopesOK && !tT.ok && /^I2:I6$/.test(yv.text)) dm.push('Plot 2 uses the slopes in your table (I2:I6), which aren\'t right yet. Fix the table first; the plot updates by itself.');
        else if (!slopesOK) dm.push(textbook ? 'The densities look like textbook values. Plot your own: the slopes in your table (I2:I6).'
          : pairs.every(([p, y]) => solByPct(p) && near(y, solByPct(p).fit.b)) ? 'The y values are the intercepts. Density is the slope: use I2:I6.'
          : `The y values (${yv.text}) aren't each solution's slope (density) from your table. Use I2:I6, with % ethanol in H2:H6 beside them.`);
      }
    }
    if (!dm.length) { tD.ok = true; tD.msg = 'Density (your slopes) against % ethanol, all five solutions.'; } else tD.msg = dm.slice(0, 3).join(' ');
  }
  return { tasks: T, done: T.every(t => t.ok) };
}
