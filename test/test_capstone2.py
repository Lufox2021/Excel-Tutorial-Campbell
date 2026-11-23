"""Capstone 2 (five ethanol-water solutions, Plot 1 + table + Plot 2): make correct and mistaken workbooks,
run the pre-check on each, compare with the expected first failing check. openpyxl writes the files; table
formulas get cached values computed here, as Excel would store them. Run from the repo root."""
import base64, json, math, pathlib, re, shutil, sys, zipfile
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, LineChart, Reference, Series
from openpyxl.chart.trendline import Trendline
from openpyxl.chart.title import Title
from openpyxl.chart.text import RichText, Text
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RegularTextRun
from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test" / "workbooks_c2"
NAME = "Jordan Smith"
V = [5.0, 10.0, 15.0, 20.0, 25.0, 30.0, 35.0, 40.0]
PCT = [0, 25, 50, 75, 100]
LAB = ["0% ethanol (water)", "25% ethanol", "50% ethanol", "75% ethanol", "100% ethanol"]
RHO = [0.997, 0.959, 0.910, 0.851, 0.785]
M = [[4.80, 9.56, 14.41, 20.87, 24.66, 29.34, 34.21, 39.97], [4.82, 8.99, 14.07, 19.34, 24.61, 29.14, 32.84, 38.28],
     [4.44, 9.03, 13.95, 18.50, 21.67, 26.81, 30.52, 37.75], [4.24, 8.65, 12.88, 17.58, 20.87, 27.25, 29.11, 34.15],
     [3.94, 7.65, 12.18, 15.97, 20.44, 22.73, 26.98, 32.23]]
T1, T2 = "Mass vs. Volume of Ethanol–Water Solutions", "Density of Ethanol–Water Solutions vs. % Ethanol"

def para(text, sz):
    return Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=sz) if sz else CharacterProperties()), r=[RegularTextRun(rPr=CharacterProperties(sz=sz) if sz else None, t=text)])
def title(lines): return Title(tx=Text(rich=RichText(p=[para(t, s) for t, s in lines])), overlay=False)

def col_vals(ws, ref):
    m = re.fullmatch(r"\$?([A-Z])\$?(\d+):\$?([A-Z])\$?(\d+)", ref)
    return [ws[f"{m.group(1)}{r}"].value for r in range(int(m.group(2)), int(m.group(4)) + 1)]
def evaluate(ws, f):
    m = re.fullmatch(r"=(SLOPE|INTERCEPT|RSQ|CORREL)\(([^,]+),([^)]+)\)", f)
    y, x = col_vals(ws, m.group(2)), col_vals(ws, m.group(3))
    n = len(x); mx, my = sum(x) / n, sum(y) / n
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y)); sxx = sum((a - mx) ** 2 for a in x); syy = sum((b - my) ** 2 for b in y)
    return {"SLOPE": sxy / sxx, "INTERCEPT": my - sxy / sxx * mx, "RSQ": sxy * sxy / (sxx * syy), "CORREL": sxy / math.sqrt(sxx * syy)}[m.group(1)]

def book(path, *, lines1=((T1, 1400), (NAME, 1100)), lines2=((T2, 1400), (NAME, 1100)), colors=None, legend=True, only=None,
         trend_skip=(), forced=(), table="ok", pct_cells="number", plot2=True, p2_y="I", p2_swap=False, p2_kind="scatter",
         p2_x_title="% Ethanol (by mass)", p2_sheet=False):
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["Volume (mL)"] + LAB)
    for i, v in enumerate(V): ws.append([v] + [m[i] for m in M])
    # Plot 1
    xr = Reference(ws, min_col=1, min_row=2, max_row=9)
    ch = ScatterChart()
    for k in range(5):
        if only is not None and k not in only: continue
        s = Series(Reference(ws, min_col=2 + k, min_row=1, max_row=9), xr, title_from_data=True)
        s.marker.symbol = "circle"; s.graphicalProperties.line.noFill = True
        if colors: s.marker.graphicalProperties.solidFill = colors[k]; s.marker.graphicalProperties.line.solidFill = colors[k]
        if k not in trend_skip: s.trendline = Trendline(trendlineType="linear", **({"intercept": 0} if k in forced else {}))
        ch.series.append(s)
    if lines1: ch.title = title(lines1)
    ch.x_axis.title = title([("Volume (mL)", 1000)]); ch.y_axis.title = title([("Mass (g)", 1000)])
    ch.x_axis.delete = False; ch.y_axis.delete = False
    if not legend: ch.legend = None
    ws.add_chart(ch, "A12")
    # Table H1:K6
    for j, h in enumerate(["% Ethanol", "Slope = Density (g/mL)", "Intercept (g)", "R²"]): ws.cell(1, 8 + j, h)
    cached = {}
    for k in range(5):
        r, col = 2 + k, "BCDEF"[k]
        if pct_cells == "number": ws[f"H{r}"] = PCT[k]
        elif pct_cells == "percent": ws[f"H{r}"] = PCT[k] / 100; ws[f"H{r}"].number_format = "0%"
        else: ws[f"H{r}"] = LAB[k]
        fs = {"I": f"=SLOPE({col}2:{col}9,$A$2:$A$9)", "J": f"=INTERCEPT({col}2:{col}9,$A$2:$A$9)", "K": f"=RSQ({col}2:{col}9,$A$2:$A$9)"}
        if table == "wrongcol" and k == 2: fs = {c: f.replace("D2:D9", "C2:C9") for c, f in fs.items()}
        for c, f in fs.items():
            if table == "typed": ws[f"{c}{r}"] = float(f"{evaluate(ws, f):.4g}")
            else: ws[f"{c}{r}"] = f; cached[f"{c}{r}"] = evaluate(ws, f)
        ws[f"M{r}"] = RHO[k]          # textbook densities, used only by the "textbook" case
    # Plot 2
    if plot2:
        target = wb.create_sheet("Plot 2") if p2_sheet else ws
        xh = Reference(ws, min_col=8, min_row=2, max_row=6)
        yh = Reference(ws, min_col={"I": 9, "J": 10, "M": 13}[p2_y], min_row=2, max_row=6)
        if p2_kind == "line":
            c2 = LineChart(); c2.add_data(Reference(ws, min_col=9, min_row=1, max_row=6), titles_from_data=True); c2.set_categories(xh)
        else:
            c2 = ScatterChart()
            s2 = Series(xh, yh) if p2_swap else Series(yh, xh)
            s2.marker.symbol = "circle"; s2.graphicalProperties.line.noFill = True; c2.series.append(s2)
        if lines2: c2.title = title(lines2)
        c2.x_axis.title = title([(p2_x_title, 1000)]); c2.y_axis.title = title([("Density (g/mL)", 1000)])
        c2.x_axis.delete = False; c2.y_axis.delete = False; c2.legend = None
        target.add_chart(c2, "A2" if p2_sheet else "H12")
    wb.save(path)
    zin = zipfile.ZipFile(path); data = {i.filename: zin.read(i.filename) for i in zin.infolist()}; infos = zin.infolist(); zin.close()
    x = data["xl/worksheets/sheet1.xml"].decode()
    for ref, v in cached.items():
        x = re.sub(rf'(<c r="{ref}"[^>]*>)<f>([^<]*)</f><v\s*/>', rf"\1<f>\2</f><v>{v!r}</v>", x)
        x = re.sub(rf'(<c r="{ref}"[^>]*>)<f>([^<]*)</f><v></v>', rf"\1<f>\2</f><v>{v!r}</v>", x)
    data["xl/worksheets/sheet1.xml"] = x.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos: z.writestr(i, data[i.filename])

CASES = {
    "C2_01_correct_auto_colors": ({}, None),
    "C2_02_correct_own_colors_percent_cells": (dict(colors=("D6EAF8", "85C1E9", "3498DB", "2471A3", "1B4F72"), pct_cells="percent"), None),
    "C2_03_two_same_color": (dict(colors=("1F77B4", "1F77B4", "2CA02C", "D62728", "9467BD")), "p1colors"),
    "C2_04_legend_off": (dict(legend=False), "p1colors"),
    "C2_05_three_plotted": (dict(only=(0, 2, 4)), "p1data"),
    "C2_06_no_trend_on_50": (dict(trend_skip=(2,)), "p1trend"),
    "C2_07_forced_zero_100": (dict(forced=(4,)), "p1trend"),
    "C2_08_table_wrong_column": (dict(table="wrongcol"), "table"),
    "C2_09_table_pct_as_text": (dict(pct_cells="text", plot2=False), "table"),
    "C2_10_table_typed": (dict(table="typed"), "table"),
    "C2_11_no_plot2": (dict(plot2=False), "p2chart"),
    "C2_12_plot2_textbook_densities": (dict(p2_y="M"), "p2data"),
    "C2_13_plot2_axes_swapped": (dict(p2_swap=True), "p2data"),
    "C2_14_plot2_no_subtitle": (dict(lines2=((T2, 1400),)), "p2subtitle"),
    "C2_15_plot2_x_no_units": (dict(p2_x_title="Ethanol"), "p2xaxis"),
    "C2_16_plot2_line_chart": (dict(p2_kind="line"), "p2data"),
    "C2_17_plot1_no_subtitle": (dict(lines1=((T1, 1400),)), "p1subtitle"),
    "C2_18_plot2_intercepts": (dict(p2_y="J"), "p2data"),
    "C2_19_plot2_on_own_sheet": (dict(p2_sheet=True), None),
}

if __name__ == "__main__":
    shutil.rmtree(OUT, ignore_errors=True); OUT.mkdir(parents=True)
    for k, (kw, _) in CASES.items(): book(OUT / f"{k}.xlsx", **kw)
    bad = 0
    with sync_playwright() as p:
        b = p.chromium.launch(); pg = b.new_page(); errs = []; pg.on("pageerror", lambda e: errs.append(str(e)))
        pg.goto((ROOT / "index.html").as_uri())
        for k, (_, want) in CASES.items():
            f = OUT / f"{k}.xlsx"
            res = json.loads(pg.evaluate("""async ([b64, n]) => { const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
                const wb = await readXlsx(bin.buffer); return JSON.stringify(checkCapstone2(wb, n)); }""", [base64.b64encode(f.read_bytes()).decode(), NAME]))
            failed = [t["id"] for t in res["tasks"] if not t["ok"]]
            ok = (res["done"] and want is None) or (not res["done"] and failed[:1] == [want])
            bad += not ok
            print(("PASS " if ok else "FAIL ") + k, "| all elements" if res["done"] else "| failed: " + ",".join(failed))
            for t in res["tasks"]:
                if not t["ok"]: print("      ✗", t["id"], "—", t["msg"])
        print("page errors:", errs); b.close()
    sys.exit(1 if bad or errs else 0)
