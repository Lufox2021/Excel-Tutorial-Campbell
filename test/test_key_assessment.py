"""Key Assessment (Mass vs Volume): make correct and mistaken workbooks, run the pre-check on each,
and compare with the expected result. Run from the repo root:  python3 test/test_key_assessment.py"""
import base64, json, pathlib, re, shutil, subprocess, sys, zipfile
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, LineChart, Reference, Series
from openpyxl.chart.trendline import Trendline
from openpyxl.chart.title import Title
from openpyxl.chart.text import RichText, Text
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RegularTextRun
from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT, RAW = ROOT / "test" / "workbooks_ka", ROOT / "test" / "raw_ka"
NAME = "Jordan Smith"
V = [5.0, 10.0, 15.0, 20.0, 25.0, 30.0, 35.0, 40.0]
M = [4.12, 7.94, 12.34, 16.79, 18.98, 24.36, 27.88, 31.01]

def para(text, sz):
    cp = (lambda: CharacterProperties(sz=sz)) if sz else (lambda: None)
    return Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=sz) if sz else CharacterProperties()), r=[RegularTextRun(rPr=cp(), t=text)])

def title(lines):   # lines: [(text, size or None)]
    return Title(tx=Text(rich=RichText(p=[para(t, s) for t, s in lines])), overlay=False)

def book(*, lines=(("Mass vs. Volume of Ethanol", 1400), (NAME, 1100)), xt="Volume (mL)", yt="Mass (g)", kind="scatter",
         trend="linear", eq=True, r2=True, forced=False, swap=False, header_in_y=False, extra=False):
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["Volume (mL)", "Mass (g)"])
    for v, m in zip(V, M): ws.append([v, m])
    if kind == "line":
        ch = LineChart(); ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=9), titles_from_data=True); ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=9))
        s = ch.series[0]
    else:
        ch = ScatterChart()
        xr, yr = Reference(ws, min_col=1, min_row=2, max_row=9), Reference(ws, min_col=2, min_row=1 if header_in_y else 2, max_row=9)
        s = Series(xr, yr) if swap else Series(yr, xr)
        s.marker.symbol = "circle"; s.graphicalProperties.line.noFill = True
        ch.series.append(s)
        if extra:
            e = Series(Reference(ws, min_col=1, min_row=2, max_row=9), xr); e.marker.symbol = "triangle"; e.graphicalProperties.line.noFill = True; ch.series.append(e)
    if trend:
        kw = dict(trendlineType=trend, dispEq=eq, dispRSqr=r2)
        if trend == "poly": kw["order"] = 2
        if forced: kw["intercept"] = 0
        s.trendline = Trendline(**kw)
    if lines: ch.title = title(lines)
    if xt: ch.x_axis.title = title([(xt, 1000)])
    if yt: ch.y_axis.title = title([(yt, 1000)])
    ch.x_axis.delete = False; ch.y_axis.delete = False; ch.legend = None
    ws.add_chart(ch, "D2")
    return wb

def edit_chart(src, dst, fn):
    zin = zipfile.ZipFile(src); zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        d = zin.read(it.filename)
        if it.filename.startswith("xl/charts/chart"): d = fn(d.decode()).encode()
        zout.writestr(it, d)
    zout.close()

T = "Mass vs. Volume of Ethanol"
CASES = {   # name: (builder kwargs, expected first failing check or None for pass)
    "K01_correct": ({}, None),
    "K02_forced_through_zero": (dict(forced=True), "trend"),
    "K04_no_title": (dict(lines=None), "title"),
    "K05_title_placeholder": (dict(lines=(("Chart Title", None),)), "title"),
    "K06_title_too_short": (dict(lines=(("Density", 1400), (NAME, 1100))), "title"),
    "K07_no_subtitle": (dict(lines=((T, 1400),)), "subtitle"),
    "K08_name_same_line": (dict(lines=((T + " - " + NAME, 1400),)), "subtitle"),
    "K08b_name_first_then_title": (dict(lines=((NAME, 1100), (T, 1400))), "title"),
    "K09_subtitle_same_size": (dict(lines=((T, 1400), (NAME, 1400))), "subtitle"),
    "K10_subtitle_larger": (dict(lines=((T, 1200), (NAME, 1600))), "subtitle"),
    "K11_subtitle_other_name": (dict(lines=((T, 1400), ("Pat Lee", 1100))), "subtitle"),
    "K12_x_no_units": (dict(xt="Volume"), "xaxis"),
    "K13_axis_titles_swapped": (dict(xt="Mass (g)", yt="Volume (mL)"), "xaxis"),
    "K14_data_swapped": (dict(swap=True), "data"),
    "K15_line_chart": (dict(kind="line"), "data"),
    "K16_no_trendline": (dict(trend=None), "trend"),
    "K17_no_equation": (dict(eq=False), "trend"),
    "K18_polynomial": (dict(trend="poly"), "trend"),
    "K19_header_in_y": (dict(header_in_y=True), "data"),
    "K21_title_default_size": (dict(lines=((T, None), (NAME, 1100))), None),   # Excel default title 14 pt > 11 pt
    "K22_extra_series": (dict(extra=True), "data"),
}

if __name__ == "__main__":
    for p in (OUT, RAW): shutil.rmtree(p, ignore_errors=True); p.mkdir(parents=True)
    expect = {}
    for k, (kw, want) in CASES.items():
        book(**kw).save(OUT / f"{k}.xlsx"); expect[k] = want
    # subtitle typed with Shift+Enter: one paragraph with a line break between two runs
    edit_chart(OUT / "K01_correct.xlsx", OUT / "K03_subtitle_line_break.xlsx",
               lambda x: re.sub(r"(<a:r><a:rPr[^>]*/><a:t>Mass vs\. Volume of Ethanol</a:t></a:r>)</a:p><a:p[^>]*><a:pPr>.*?</a:pPr>", r"\1<a:br/>", x, count=1))
    expect["K03_subtitle_line_break"] = None
    edit_chart(OUT / "K01_correct.xlsx", OUT / "K20_no_x_values.xlsx", lambda x: re.sub(r"<xVal>.*?</xVal>", "", x)); expect["K20_no_x_values"] = "data"
    # LibreOffice rewrite of two correct files (a second chart writer)
    for k in ("K01_correct", "K03_subtitle_line_break"): shutil.copy(OUT / f"{k}.xlsx", RAW / f"{k}_lo.xlsx")
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(OUT)] + [str(f) for f in RAW.glob("*.xlsx")], check=True, capture_output=True)
    # LibreOffice flattens a two-line title into one run with a newline and one size, so the smaller name is lost
    expect["K01_correct_lo"] = "subtitle"; expect["K03_subtitle_line_break_lo"] = "subtitle"
    assert "<a:br/>" in zipfile.ZipFile(OUT / "K03_subtitle_line_break.xlsx").read("xl/charts/chart1.xml").decode()

    bad = 0
    with sync_playwright() as p:
        b = p.chromium.launch(); pg = b.new_page(); errs = []; pg.on("pageerror", lambda e: errs.append(str(e)))
        pg.goto((ROOT / "index.html").as_uri())
        for f in sorted(OUT.glob("*.xlsx")):
            res = json.loads(pg.evaluate("""async ([b64, n]) => { const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
                const wb = await readXlsx(bin.buffer); return JSON.stringify(checkKeyAssessment(wb, n)); }""", [base64.b64encode(f.read_bytes()).decode(), NAME]))
            failed = [t["id"] for t in res["tasks"] if not t["ok"]]
            want = expect[f.stem]
            ok = (res["done"] and want is None) or (not res["done"] and failed[:1] == [want])
            bad += not ok
            print(("PASS " if ok else "FAIL ") + f.stem, "| all six" if res["done"] else "| failed: " + ",".join(failed))
            for t in res["tasks"]:
                if not t["ok"]: print("      ✗", t["id"], "—", t["msg"])
        # no name typed: subtitle can't be checked
        res = json.loads(pg.evaluate("""async b64 => { const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0)); const wb = await readXlsx(bin.buffer); return JSON.stringify(checkKeyAssessment(wb, '')); }""", base64.b64encode((OUT / "K01_correct.xlsx").read_bytes()).decode()))
        sub = [t for t in res["tasks"] if t["id"] == "subtitle"][0]
        ok = not sub["ok"] and "Type your name" in sub["msg"]; bad += not ok
        print(("PASS " if ok else "FAIL ") + "no name typed -> asks for the name")
        print("page errors:", errs); b.close()
    sys.exit(1 if bad or errs else 0)
