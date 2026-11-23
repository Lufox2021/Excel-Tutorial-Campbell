"""Make Lab B and Lab C workbooks with real charts (correct, and with typical mistakes).
openpyxl writes them; LibreOffice then recalculates and rewrites them, as a second chart writer.
A few Lab B files are also kept exactly as openpyxl wrote them. Run from the repo root."""
import json, pathlib, subprocess, shutil
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, LineChart, Reference, Series
from openpyxl.chart.trendline import Trendline
from openpyxl.chart.title import Title
from openpyxl.chart.text import RichText, Text
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RegularTextRun
from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).resolve().parent.parent
RAW, OUT = ROOT / "test" / "raw_bc", ROOT / "test" / "workbooks_bc"
NAME = "Jordan Smith"

def page_data():
    with sync_playwright() as p:
        b = p.chromium.launch(); pg = b.new_page(); pg.goto((ROOT / "index.html").as_uri())
        d = {k: json.loads(pg.evaluate(f"n => JSON.stringify({{data: lab{k}Data(n), block: lab{k}Block(lab{k}Data(n))}})", NAME)) for k in "BC"}
        b.close()
    return d

def title(text, sz):
    run_pr = CharacterProperties(sz=sz) if sz else None
    para = Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=sz) if sz else CharacterProperties()), r=[RegularTextRun(rPr=run_pr, t=text)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)

def legend_size(ch, sz):
    if sz: ch.legend.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=sz)), endParaRPr=CharacterProperties(sz=sz))])

def put_block(ws, block, name):
    for r, row in enumerate(block, 1):
        for c, v in enumerate(row, 1):
            if v == "": continue
            try: ws.cell(r, c, float(v))
            except ValueError: ws.cell(r, c, v)
    if name is not None: ws["B1"] = name

# ---------------- Lab B ----------------
def lab_b(d, *, kind="scatter", sz=1200, title_sz=None, chart_title="Pressure vs. volume at two temperatures", xt="Volume (L)", yt="Pressure (atm)",
          legend=True, names=True, lines=False, no_x=False, swap=False, header_in_y=False, one_series=False, name=NAME, change_data=False):
    wb = Workbook(); ws = wb.active; ws.title = "Lab B"
    put_block(ws, d["block"], name)
    if change_data: ws["B5"] = 9.999
    if kind == "line":
        ch = LineChart()
        for c in (2, 3): ch.add_data(Reference(ws, min_col=c, min_row=3, max_row=11), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=4, max_row=11))
    else:
        ch = ScatterChart()
        x = Reference(ws, min_col=1, min_row=4, max_row=11)
        for c in ((2,) if one_series else (2, 3)):
            yref = Reference(ws, min_col=c, min_row=3 if (names and not header_in_y) or header_in_y else 4, max_row=11)
            if swap: s = Series(Reference(ws, min_col=1, min_row=4, max_row=11), Reference(ws, min_col=c, min_row=4, max_row=11), title_from_data=False)
            elif header_in_y: s = Series(yref, None if no_x else x, title_from_data=False)
            else: s = Series(yref, None if no_x else x, title_from_data=names)
            s.marker.symbol = "circle"
            if lines: s.graphicalProperties.line.solidFill = "1F77B4"
            else: s.graphicalProperties.line.noFill = True
            ch.series.append(s)
    if chart_title is not None: ch.title = title(chart_title, title_sz or sz)
    if xt is not None: ch.x_axis.title = title(xt, sz)
    if yt is not None: ch.y_axis.title = title(yt, sz)
    ch.x_axis.delete = False; ch.y_axis.delete = False
    if legend: legend_size(ch, sz)
    else: ch.legend = None
    ws.add_chart(ch, "E3")
    return wb

# ---------------- Lab C ----------------
def lab_c(d, *, sz=1200, chart_title="Calibration curve: absorbance vs. concentration", xt="Concentration (M)", yt="Absorbance", legend=False,
          trend="linear", eq=True, r2=True, forced=False, swap=False, unk_series=False, lines=False, kind="scatter",
          slope="=SLOPE(B4:B9,A4:A9)", icpt="=INTERCEPT(B4:B9,A4:A9)", rsq="=RSQ(B4:B9,A4:A9)", conc="=(E{r}-$I$5)/$I$4",
          fmt=True, name=NAME):
    wb = Workbook(); ws = wb.active; ws.title = "Lab C"
    put_block(ws, d["block"], name)
    ws["H4"], ws["H5"], ws["H6"], ws["F3"] = "Slope", "Intercept", "R²", "Concentration (M)"
    ws["I4"], ws["I5"], ws["I6"] = slope, icpt, rsq
    for i, r in enumerate((4, 5, 6)):
        ws[f"F{r}"] = conc[i] if isinstance(conc, list) else conc.format(r=r, r1=r + 1, r2=r + 2)
    if fmt:
        for ref in ("I4", "I5", "F4", "F5", "F6"): ws[ref].number_format = "0.000E+00"
        ws["I6"].number_format = "0.0000"
    if kind == "line":
        ch = LineChart(); ch.add_data(Reference(ws, min_col=2, min_row=3, max_row=9), titles_from_data=True); ch.set_categories(Reference(ws, min_col=1, min_row=4, max_row=9))
    else:
        ch = ScatterChart()
        x, y = Reference(ws, min_col=1, min_row=4, max_row=9), Reference(ws, min_col=2, min_row=3, max_row=9)
        s = Series(Reference(ws, min_col=1, min_row=4, max_row=9), Reference(ws, min_col=2, min_row=4, max_row=9)) if swap else Series(y, x, title_from_data=True)
        s.marker.symbol = "circle"
        if lines: s.graphicalProperties.line.solidFill = "1F77B4"
        else: s.graphicalProperties.line.noFill = True
        if trend:
            kw = dict(trendlineType=trend, dispEq=eq, dispRSqr=r2)
            if trend == "poly": kw["order"] = 2
            if forced: kw["intercept"] = 0
            s.trendline = Trendline(**kw)
        ch.series.append(s)
        if unk_series:
            u = Series(Reference(ws, min_col=5, min_row=3, max_row=6), Reference(ws, min_col=6, min_row=4, max_row=6), title_from_data=True)
            u.marker.symbol = "triangle"; u.graphicalProperties.line.noFill = True; ch.series.append(u)
    if chart_title is not None: ch.title = title(chart_title, sz)
    if xt is not None: ch.x_axis.title = title(xt, sz)
    if yt is not None: ch.y_axis.title = title(yt, sz)
    ch.x_axis.delete = False; ch.y_axis.delete = False
    if legend: legend_size(ch, sz)
    else: ch.legend = None
    ws.add_chart(ch, "K3")
    return wb

B_CASES = {
    "B01_correct": {}, "B02_line_chart": dict(kind="line"), "B03_no_x_values": dict(no_x=True), "B04_series_unnamed": dict(names=False),
    "B05_lines_joined": dict(lines=True), "B06_no_title": dict(chart_title=None), "B07_axis_title_placeholder": dict(xt="Axis Title"),
    "B08_no_units": dict(yt="Pressure"), "B09_legend_off": dict(legend=False), "B10_default_fonts": dict(sz=None),
    "B11_title_14pt": dict(title_sz=1400), "B12_axes_swapped": dict(swap=True), "B13_header_in_y": dict(header_in_y=True),
    "B14_one_series": dict(one_series=True), "B15_data_changed": dict(change_data=True),
}
C_CASES = {
    "C01_correct": {}, "C02_line_chart": dict(kind="line"), "C03_axes_swapped": dict(swap=True), "C04_unknowns_plotted": dict(unk_series=True),
    "C05_legend_on": dict(legend=True), "C06_no_trendline": dict(trend=None), "C07_trend_no_r2": dict(r2=False), "C08_trend_forced_zero": dict(forced=True),
    "C09_trend_poly": dict(trend="poly"), "C10_slope_swapped": dict(slope="=SLOPE(A4:A9,B4:B9)"), "C11_correl": dict(rsq="=CORREL(B4:B9,A4:A9)"),
    "C12_no_intercept": dict(conc="=E{r}/$I$4"), "C13_missing_dollar": dict(conc="=(E{r}-I{r1})/I{r}"), "C14_no_formatting": dict(fmt=False),
    "C15_lines_joined": dict(lines=True), "C16_default_fonts": dict(sz=None),
}

if __name__ == "__main__":
    d = page_data()
    for p in (RAW, OUT): shutil.rmtree(p, ignore_errors=True); p.mkdir(parents=True)
    for k, kw in B_CASES.items(): lab_b(d["B"], **kw).save(RAW / f"{k}.xlsx")
    for k, kw in C_CASES.items(): lab_c(d["C"], **kw).save(RAW / f"{k}.xlsx")
    # C17: concentrations from numbers copied off the chart label (rounded), C18: typed results
    fit = d["C"]["data"]["fit"]
    m, b = float(f"{fit['m']:.4g}"), float(f"{fit['b']:.4f}")
    lab_c(d["C"], conc=f"=(E{{r}}-{b})/{m}").save(RAW / "C17_numbers_from_chart.xlsx")
    lab_c(d["C"], conc=[float(f"{u['c']:.4g}") for u in d["C"]["data"]["unk"]]).save(RAW / "C18_typed_results.xlsx")
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(OUT)]
                   + [str(f) for f in sorted(RAW.glob("*.xlsx"))], check=True, capture_output=True)
    # keep three Lab B files exactly as openpyxl wrote them (a different chart writer)
    for k in ("B01_correct", "B02_line_chart", "B10_default_fonts"): shutil.copy(RAW / f"{k}.xlsx", OUT / f"{k}_openpyxl.xlsx")
    # Neither LibreOffice nor openpyxl can save a scatter series without x values (both drop the data).
    # Excel keeps it and plots against 1, 2, 3, ...; make that file by removing <xVal> from the correct one.
    import re, zipfile
    zin = zipfile.ZipFile(RAW / "B01_correct.xlsx"); zout = zipfile.ZipFile(OUT / "B03_no_x_values.xlsx", "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename.startswith("xl/charts/chart"): data = re.sub(rb"<xVal>.*?</xVal>", b"", data)
        zout.writestr(it, data)
    zout.close()
    print("made", len(list(OUT.glob("*.xlsx"))), "chart workbooks for", NAME)
