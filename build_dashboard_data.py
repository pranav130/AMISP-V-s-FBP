"""
Build data.js for the AMISP Feeder Mapping Dashboard from the comparison workbook.

Usage:
    python build_dashboard_data.py "comparison OTHER THAN KESCO_Nov.xlsx" --asof 03.11.2026 [--summary-sheet 03.11.2026]

Open the workbook in Excel once and save it first (so Mapping's formulas are recalculated),
then run this, then push index.html + data.js to GitHub.
Requires: pip install pandas openpyxl
"""
import argparse, json, datetime, re
import pandas as pd, openpyxl

DISCOM = {"PASCHIMANCHAL": "PVVNL", "DAKSHINANCHAL": "DVVNL", "MADHYANCHAL": "MVVNL", "POORVANCHAL": "PuVVNL"}
AMISP = {"INTELLI": "Intelli", "INTELI": "Intelli", "GENUS": "Genus", "GMR": "GMR", "POLARIS": "Polaris"}
MASTER_DISCOM = {"PVVNL": "PVVNL", "DVVNL": "DVVNL", "MVVNL": "MVVNL", "PUVVNL": "PuVVNL"}
ORDER = [("PVVNL", "Intelli"), ("DVVNL", "Genus"), ("DVVNL", "GMR"), ("MVVNL", "Intelli"),
         ("MVVNL", "Polaris"), ("PuVVNL", "Genus"), ("PuVVNL", "GMR")]
CATS = ["Meter No & MF Matched", "Meter No Matched, MF Mismatched", "Meter No Mismatched",
        "Division mismatch", "Subtation mismatch", "Mapping not done"]


def s(v):
    if v is None or (isinstance(v, float) and pd.isna(v)): return ""
    return str(v).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook"); ap.add_argument("--asof", required=True)
    ap.add_argument("--summary-sheet"); ap.add_argument("--out", default="data.js")
    a = ap.parse_args()
    wbv = openpyxl.load_workbook(a.workbook, read_only=True, data_only=True)
    names = wbv.sheetnames
    sum_sheet = a.summary_sheet or next(n for n in names if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", n))

    # ---------- Mapping (per FBP feeder) ----------
    m = pd.read_excel(a.workbook, sheet_name="Mapping", header=1, dtype=str)
    cols = list(m.columns)
    fbp_date = s(wbv["Mapping"]["D1"].value)
    pick = lambda prefix: next(c for c in cols if c.startswith(prefix))
    m = pd.DataFrame({
        "discom": m["Discom"].map(lambda x: DISCOM.get(s(x).upper(), s(x))),
        "amisp": m["AMISP NAME"].map(lambda x: AMISP.get(s(x).upper(), s(x))),
        "zone": m["Zone"].map(s), "circle": m["Circle"].map(s), "division": m["Division"].map(s),
        "substation": m["Substation"].map(s), "feeder": m["Outgoing Feeder"].map(s),
        "type": m["type"].map(lambda x: s(x).capitalize() if s(x).upper() not in ("SS",) else "SS"),
        "meter": m[pick("Meter SlNo")].map(s), "mf": m[pick("OverallMF")].map(s),
        "a_div": m["AMISP Division"].map(s), "a_ss": m["AMISP substation name"].map(s),
        "a_feeder": m["AMISP feeder name"].map(s), "a_meter": m["AMISP Meter No."].map(s),
        "a_mf": m["AMISP MF"].map(s), "remark": m[pick("Remark_")].map(s),
        "match": m["Matching remark"].map(s), "lookup": m["Lookup by"].map(s) if "Lookup by" in cols else m["Lookup By"].map(s),
    })
    m = m[m.feeder != ""]
    m.loc[m.remark == "", "remark"] = "Mapping not done"
    o = m[m.type.str.upper() == "OTHERS"]

    # circles (Others feeders, all matching = manual view)
    circ = []
    for (d, am, z, c), g in o.groupby(["discom", "amisp", "zone", "circle"]):
        r = {"discom": d, "amisp": am, "zone": z, "circle": c, "feeders": len(g),
             "direct": int((g.match == "Directly matched").sum())}
        for k in CATS: r[k] = int((g.remark == k).sum())
        circ.append(r)

    lookup = [{"discom": d, "amisp": am, "lookup": lk or "(blank)", "n": len(g)}
              for (d, am, lk), g in m.groupby(["discom", "amisp", "lookup"])]

    exc_cols = ["discom", "amisp", "zone", "circle", "division", "substation", "feeder", "meter", "mf",
                "a_feeder", "a_meter", "a_mf", "a_div", "a_ss", "remark", "match", "lookup"]
    exc = o[o.remark != "Meter No & MF Matched"][exc_cols]

    # ---------- AMISP feeder master ----------
    am = pd.read_excel(a.workbook, sheet_name="AMISP feeder master", dtype=str, usecols="A:L")
    am = am[am["Device ID"].notna()]
    am["d"] = am["Discom"].str.upper().map(MASTER_DISCOM)
    am["a"] = am["AMISP"].str.upper().map(AMISP)
    am["t"] = am["TYPE"].fillna("(blank)")
    master = [{"discom": d, "amisp": x, "type": t, "n": len(g)} for (d, x, t), g in am.groupby(["d", "a", "t"])]
    master_date = s(am["DATE"].dropna().iloc[0]) if am["DATE"].notna().any() else ""

    # ---------- Summary sheet (03.11.2026 style) ----------
    ws = wbv[sum_sheet]
    grid = {(r, c): ws.cell(r, c).value for r in range(1, 43) for c in range(1, 15)}
    def num(v):
        try: return float(v)
        except (TypeError, ValueError): return None
    summary = []
    for sec, rows in (("Direct", range(6, 19)), ("Manual", range(29, 42))):
        cur = None
        for r in rows:
            name, amisp = s(grid[(r, 1)]), s(grid[(r, 2)])
            if not amisp or "TOTAL" in name.upper(): continue
            vals = {k: num(grid[(r, c)]) for k, c in zip("CDEFIKLM", (3, 4, 5, 6, 9, 11, 12, 13))}
            vals = {k: (int(v) if v is not None else 0) for k, v in vals.items()}
            disc = {"PUVVNL": "PuVVNL"}.get(name.upper(), name.upper())
            row = {"section": sec, "discom": disc, "amisp": AMISP.get(amisp.upper(), amisp),
                   "fbp_total": vals["C"], "mdm_total": vals["D"], "fbp_others": vals["E"], "mdm_others": vals["F"],
                   "mtr_mf": vals["I"], "mf_mm": vals["K"], "mtr_mm": vals["L"], "div_mm": vals["M"]}
            row["matched"] = row["mtr_mf"] + row["mf_mm"] + row["mtr_mm"]
            nval = num(grid[(r, 14)])
            row["not_matched"] = int(nval) if (disc == "KESCO" and nval is not None) else row["mdm_others"] - row["matched"] - row["div_mm"]
            summary.append(row)

    data = {
        "meta": {"asof": a.asof, "summary_sheet": sum_sheet, "fbp_date": fbp_date, "amisp_date": master_date,
                 "built": datetime.datetime.now().strftime("%d.%m.%Y %H:%M"), "workbook": a.workbook.split("\\")[-1].split("/")[-1],
                 "order": ORDER, "cats": CATS},
        "summary": summary, "circles": circ, "lookup": lookup, "master": master,
        "exceptions": {"cols": exc_cols, "rows": exc.values.tolist()},
        "counts": {"mapping_rows": int(len(m)), "others": int(len(o)), "master_rows": int(len(am))},
    }
    with open(a.out, "w", encoding="utf-8") as f:
        f.write("window.DASH_DATA = ")
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";\n")
    print(f"Wrote {a.out}: {len(summary)} summary rows, {len(circ)} circles, {len(exc)} exception feeders")


if __name__ == "__main__":
    main()
