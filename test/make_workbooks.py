"""Make student-style Lab A workbooks (correct and with typical mistakes), then let
LibreOffice recalculate them so they carry cached values like a real Excel file.
Run from the repo root:  python3 test/make_workbooks.py"""
import json, pathlib, subprocess, shutil
from openpyxl import Workbook
from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).resolve().parent.parent
RAW, OUT = ROOT / "test" / "raw", ROOT / "test" / "workbooks"
NAME = "Jordan Smith"

def lab_data(name):
    with sync_playwright() as p:
        b = p.chromium.launch(); pg = b.new_page()
        pg.goto((ROOT / "index.html").as_uri())
        d = json.loads(pg.evaluate("n => JSON.stringify(labAData(n))", name)); b.close()
    return d

def build(d, name_in_b1, *, sort=True, one_col_sort=False, dollar=True, typed_mean=False, pop_sd=False,
          cv_mode="x100", fmt=True, mol_fmt="0.000E+00", stat_fmt="0.000E+00", cv_fmt="0.000", skip_e=False, mL=True,
          table_at_E=False, on_volume=False, wrong_range_E=False, legacy_stdev=False,
          swap_maxmin=False, typed_n=False, wrong_n_range=False, no_extra_cols=False):
    wb = Workbook(); ws = wb.active; ws.title = "Lab A"
    ws["A1"], ws["B1"] = "Name", name_in_b1
    ws["A2"], ws["B2"] = "NaOH concentration (M)", d["conc"]
    ws["A3"], ws["B3"], ws["C3"], ws["D3"] = "Sample", "Trial", "Volume (mL)", "NaOH (mol)"
    rows = list(d["given"]) + ([] if skip_e else list(d["typed"]))
    if sort: rows = sorted(rows, key=lambda x: (x["sample"], x["trial"]))
    vols = [x["vol"] for x in rows]
    if one_col_sort: vols = sorted(vols)   # volumes sorted by themselves: rows broken
    for i, (x, v) in enumerate(zip(rows, vols)):
        r = 4 + i
        ws.cell(r, 1, x["sample"]); ws.cell(r, 2, x["trial"]); ws.cell(r, 3, v)
        f = f"=C{r}/1000*$B$2" if dollar else f"=C{r}/1000*B{r-2}"
        if not mL: f = f"=C{r}*$B$2"
        c = ws.cell(r, 4, f)
        if fmt: c.number_format = mol_fmt
    hdr_col = 5 if table_at_E else 6
    for j, h in enumerate(["Sample", "Mean", "SD", "%CV", "Median", "Max", "Min", "n"]): ws.cell(3, hdr_col + j, h)
    src = "C" if on_volume else "D"
    for i, s in enumerate("ABCDE"):
        r, a = 4 + i, 4 + 3 * i
        rng = f"{src}{a}:{src}{a+2}"
        ws.cell(r, hdr_col, s)
        m = ws.cell(r, hdr_col + 1, float(f"{d['stats'][s]['mean']:.4g}") if typed_mean else f"=AVERAGE({rng})")
        fn = "_xlfn.STDEV.P" if pop_sd else ("STDEV" if legacy_stdev else "_xlfn.STDEV.S")
        G, H = f"{chr(64+hdr_col+1)}{r}", f"{chr(64+hdr_col+2)}{r}"
        sd = ws.cell(r, hdr_col + 2, f"={fn}({'D8:D10' if s == 'E' and wrong_range_E else rng})")  # like the model file
        cv = ws.cell(r, hdr_col + 3, {"x100": f"=100*{H}/{G}", "frac": f"={H}/{G}"}[cv_mode])
        mx, mn = ("MIN", "MAX") if swap_maxmin else ("MAX", "MIN")
        extra = [ws.cell(r, hdr_col + 4, f"=MEDIAN({rng})"), ws.cell(r, hdr_col + 5, f"={mx}({rng})"), ws.cell(r, hdr_col + 6, f"={mn}({rng})")]
        nc = ws.cell(r, hdr_col + 7, 3 if typed_n else f"=COUNT({'D4:D6' if wrong_n_range and s == 'B' else rng})")
        if fmt:
            for c in (m, sd, *extra): c.number_format = stat_fmt
            cv.number_format = cv_fmt
            nc.number_format = "0"
    return wb

CASES = {
    "01_correct": dict(),
    "02_typed_means": dict(typed_mean=True),
    "03_stdev_p": dict(pop_sd=True),
    "04_one_column_sort": dict(one_col_sort=True),
    "05_missing_dollar": dict(dollar=False),
    "06_no_formatting": dict(fmt=False),
    "07_wrong_name": dict(),          # B1 changed below
    "08_cv_no_x100": dict(cv_mode="frac"),
    "09_cv_scientific": dict(cv_fmt="0.000E+00"),   # should pass (as in the instructor's model)
    "10_unsorted": dict(sort=False),
    "11_sample_E_missing": dict(skip_e=True),
    "12_mol_not_scientific": dict(mol_fmt="0.00000"),
    "13_mean_sd_not_scientific": dict(stat_fmt="0.000"),
    "14_mL_not_converted": dict(mL=False),
    "15_blank_name": dict(),
    "16_table_starts_in_E": dict(table_at_E=True),
    "17_stats_on_volume": dict(on_volume=True),
    "18_wrong_range_E": dict(wrong_range_E=True),
    "19_legacy_STDEV": dict(legacy_stdev=True),   # should pass
    "20_max_min_swapped": dict(swap_maxmin=True),
    "21_n_typed": dict(typed_n=True),
    "22_n_wrong_range": dict(wrong_n_range=True),
}

if __name__ == "__main__":
    d = lab_data(NAME)
    for p in (RAW, OUT):
        shutil.rmtree(p, ignore_errors=True); p.mkdir(parents=True)
    for case, kw in CASES.items():
        nm = {"07_wrong_name": "Jordan Smyth", "15_blank_name": None}.get(case, NAME)
        build(d, nm, **kw).save(RAW / f"{case}.xlsx")
    # LibreOffice recalculates and writes cached values, as Excel does
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                    "--outdir", str(OUT)] + [str(f) for f in sorted(RAW.glob("*.xlsx"))],
                   check=True, capture_output=True)
    print("made", len(list(OUT.glob("*.xlsx"))), "workbooks for", NAME)
