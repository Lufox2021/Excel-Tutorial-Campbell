"""Run the Lab A checker on every test workbook and compare with the expected result.
Run from the repo root after make_workbooks.py:  python3 test/check_workbooks.py"""
import base64, json, pathlib, sys
from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).resolve().parent.parent
# case -> (should pass?, the task ids that should fail)
EXPECT = {
    "01_correct": (True, []),
    "02_typed_means": (False, ["mean", "cv"]),
    "03_stdev_p": (False, ["sd", "cv"]),
    "04_one_column_sort": (False, ["together"]),
    "05_missing_dollar": (False, ["moles"]),
    "06_no_formatting": (False, ["molfmt", "statfmt"]),
    "07_wrong_name": (False, ["data"]),
    "08_cv_no_x100": (False, ["cv"]),
    "09_cv_scientific": (True, []),
    "10_unsorted": (False, ["sorted", "mean"]),
    "11_sample_E_missing": (False, ["data"]),
    "12_mol_not_scientific": (False, ["molfmt"]),
    "13_mean_sd_not_scientific": (False, ["statfmt"]),
    "14_mL_not_converted": (False, ["moles"]),
    "15_blank_name": (False, ["name"]),
    "16_table_starts_in_E": (False, ["table"]),
    "17_stats_on_volume": (False, ["mean", "sd"]),
    "18_wrong_range_E": (False, ["sd", "cv"]),
    "19_legacy_STDEV": (True, []),
    "20_max_min_swapped": (False, ["mmm"]),
    "21_n_typed": (False, ["n"]),
    "22_n_wrong_range": (False, ["n"]),
    "real_model_as_sent": (False, ["table"]),   # instructor's Mac Excel file (before Median/Max/Min/n were added)
    "real_model_fixed": (False, ["table"]),
    "filter_method": (True, []),   # =STAT(FILTER($D$4:$D$18,$A$4:$A$18=F4)) filled down
    "countif_n": (True, []),       # n as =COUNTIF($A$4:$A$18,F4)
    # ---- Lab B (charts) ----
    "B01_correct": (True, []), "B01_correct_openpyxl": (True, []),
    "B02_line_chart": (False, ["scatter"]), "B02_line_chart_openpyxl": (False, ["scatter"]),
    "B03_no_x_values": (False, ["series"]), "B04_series_unnamed": (False, ["names"]), "B05_lines_joined": (False, ["markers"]),
    "B06_no_title": (False, ["title"]), "B07_axis_title_placeholder": (False, ["axtitles"]), "B08_no_units": (False, ["axtitles"]),
    "B09_legend_off": (False, ["legend"]), "B10_default_fonts": (False, ["fonts"]), "B10_default_fonts_openpyxl": (False, ["fonts"]),
    "B11_title_14pt": (False, ["fonts"]), "B12_axes_swapped": (False, ["series"]), "B13_header_in_y": (False, ["series"]),
    "B14_one_series": (False, ["series"]), "B15_data_changed": (False, ["data"]),
    # ---- Lab C (calibration) ----
    "C01_correct": (True, []), "C02_line_chart": (False, ["chart"]), "C03_axes_swapped": (False, ["series"]),
    "C04_unknowns_plotted": (False, ["series"]), "C05_legend_on": (False, ["labels"]), "C06_no_trendline": (False, ["trend"]),
    "C07_trend_no_r2": (False, ["trend"]), "C08_trend_forced_zero": (False, ["trend"]), "C09_trend_poly": (False, ["trend"]),
    "C10_slope_swapped": (False, ["fit"]), "C11_correl": (False, ["fit"]), "C12_no_intercept": (False, ["unknowns"]),
    "C13_missing_dollar": (False, ["unknowns"]), "C14_no_formatting": (False, ["fmt"]), "C15_lines_joined": (False, ["markers"]),
    "C16_default_fonts": (False, ["fonts"]), "C17_numbers_from_chart": (False, ["unknowns"]), "C18_typed_results": (False, ["unknowns"]),
}
bad = 0
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page()
    errs = []; pg.on("pageerror", lambda e: errs.append(str(e)))
    pg.goto((ROOT / "index.html").as_uri())
    for f in sorted((ROOT / "test" / "workbooks").glob("*.xlsx")) + sorted((ROOT / "test" / "real").glob("*.xlsx")) + sorted((ROOT / "test" / "workbooks_bc").glob("*.xlsx")):
        b64 = base64.b64encode(f.read_bytes()).decode()
        res = json.loads(pg.evaluate("""async b64 => {
            const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
            const wb = await readXlsx(bin.buffer); return JSON.stringify(checkWorkbook(wb)); }""", b64))
        failed = [t["id"] for t in res["tasks"] if not t["ok"]]
        want_pass, want_fail = EXPECT[f.stem]
        ok = res["done"] == want_pass and all(x in failed for x in want_fail) and (want_pass or failed[:1] == want_fail[:1])
        bad += not ok
        print(("PASS " if ok else "FAIL ") + f"[{res['lab']}] " + f.stem, "| code" if res["done"] else "| failed: " + ",".join(failed))
        for t in res["tasks"]:
            if not t["ok"]: print("      ✗", t["id"], "—", t["msg"])
    print("page errors:", errs)
    b.close()
sys.exit(1 if bad or errs else 0)
