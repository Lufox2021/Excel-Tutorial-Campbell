# Excel Tutorial & Campbell Key Assessment

Campbell University version of the Excel Workbench (the general version is kept in its own repository). Same lessons and labs; its completion codes are separate.

Live tool (once GitHub Pages is on): https://lufox2021.github.io/Excel-Tutorial-Campbell/

Teaches Excel for the web with chemistry data (Modules 1–11, Labs A–C), then checks the student's own .xlsx in the browser and gives a completion code.

- `index.html` – the tool (built from `src/` with `python3 build.py`)
- `src/` – `template.html`, `style.css`, `engine.js` (Lab A data, .xlsx reader incl. charts, Lab A checker, codes), `labs_bc.js` (Labs B and C, lab detection), `keyassess.js` (Capstones 1 and 2), `minigrid.js` (Try-it grids), `ui.js`
- `test/make_workbooks.py`, `test/make_chart_workbooks.py` – student-style workbooks (correct + typical mistakes) for all three labs
- `test/check_workbooks.py` – runs the checker on every workbook and compares with the expected result
- `test/browser_test.py` – clicks through lessons, Try-it grids, Labs, Check, Instructor
- `test/test_key_assessment.py`, `test/test_capstone2.py` – Capstone 1 and 2 cases (make the workbooks and check them)
- `test/real/` – the instructor's model workbook and Excel-style variants
