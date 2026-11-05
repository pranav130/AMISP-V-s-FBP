# AMISP Feeder Mapping Dashboard (UPPCL)

Static dashboard (GitHub Pages ready): `index.html` + `data.js`. No server or build tools needed.

## Files
- `index.html` – the dashboard page.
- `data.js` – all figures, generated from the comparison workbook (Mapping, Sheet1/summary sheet, AMISP feeder master).
- `build_dashboard_data.py` – regenerates `data.js` from the workbook.

## Monthly update
1. Update the workbook (`update_amisp_master.py`), open it in Excel, press Ctrl+Alt+F9, refresh the pivot, save.
2. Regenerate the data (put the summary sheet date as `--asof`):
   ```
   python build_dashboard_data.py "C:\Users\HP\Desktop\AMISP -SHEET-Updation\comparison OTHER THAN KESCO_Nov.xlsx" --asof 03.11.2026
   ```
   (It picks the sheet named like `dd.mm.yyyy` as the status report; use `--summary-sheet` to name it explicitly.)
3. Commit and push `index.html` and `data.js`:
   ```
   git add index.html data.js
   git commit -m "AMISP dashboard data 03.11.2026"
   git push
   ```
4. GitHub → repo Settings → Pages → Deploy from branch `main`, folder `/ (root)` (first time only).

Requires: `pip install pandas openpyxl`
