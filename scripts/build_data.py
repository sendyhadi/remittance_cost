"""
Rebuild data/rpw_costs.js from the World Bank Remittance Prices Worldwide (RPW) Excel file.

Usage:
    pip install pandas openpyxl xlsxwriter
    python scripts/build_data.py path/to/rpw_dataset_2011_2025_q3.xlsx

Input (in this repo):
    data/wesp_regions.csv   ISO3 -> UN WESP Annex region, subregion and development status

Method (matches the World Bank's SDG 10.c.1 simple average):
    - sheet "Dataset (from Q2 2016)"
    - keep transparent services only (transparent == "yes", any capitalisation)
    - keep all sending and receiving countries; the viewer filters by WESP region, subregion and status
    - cost columns are coerced to numbers (some quarters are stored as text in the source file)
    - aggregate sums and counts so the viewer can recompute exact simple averages for any filter
"""
import csv
import json
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SHEET = "Dataset (from Q2 2016)"
C200, C500 = "cc1 total cost %", "cc2 total cost %"
PROVIDER_TYPES = ["MTO", "Bank", "Mobile operator", "Post office", "Mixed", "Other"]
INSTRUMENTS = ["Cash", "Bank account", "Card", "Mobile money", "Multiple", "Other", "Not specified"]


def provider_type(x):
    x = str(x).lower()
    if "/" in x:
        return "Mixed"
    if "money transfer" in x:
        return "MTO"
    if "bank" in x and "non" not in x:
        return "Bank"
    if "post" in x:
        return "Post office"
    if "mobile" in x:
        return "Mobile operator"
    return "Other"


def instrument(x):
    if pd.isna(x):
        return "Not specified"
    parts = {p.strip().lower() for p in str(x).split(",") if p.strip()}
    if not parts:
        return "Not specified"
    if len(parts) > 1:
        return "Multiple"
    p = parts.pop()
    if p == "cash":
        return "Cash"
    if p == "bank account transfer":
        return "Bank account"
    if "card" in p:
        return "Card"
    if "mobile" in p:
        return "Mobile money"
    return "Other"


def quarter_range(first, last):
    out = []
    y, q = int(first[:4]), int(first[5])
    while True:
        p = f"{y}_{q}Q"
        out.append(p)
        if p == last:
            return out
        q += 1
        if q == 5:
            y, q = y + 1, 1


STATUS_NAMES = {"DVD": "Developed", "DVG": "Developing", "EIT": "Economies in transition"}
FULL_FILE = "data/rpw_services_2016q2_2025q3.xlsx"


def write_full_excel(d, info):
    """Write the cleaned service-level file behind the 'Download full dataset' button (needs xlsxwriter)."""
    import math
    import xlsxwriter

    d = d.copy()
    for side, col in (("Sender", "source_code"), ("Receiver", "destination_code")):
        m = {c: info(c) for c in d[col].unique()}
        d[f"{side} WESP region"] = d[col].map(lambda c: m[c][0])
        d[f"{side} WESP subregion"] = d[col].map(lambda c: m[c][1])
        d[f"{side} development status"] = d[col].map(lambda c: STATUS_NAMES.get(m[c][2], "Unclassified"))
    d[C200] = d[C200].round(4)
    d[C500] = d[C500].round(4)
    cols = [("period", "Quarter"), ("source_code", "Sender ISO3"), ("source_name", "Sender"),
            ("Sender WESP region", None), ("Sender WESP subregion", None), ("Sender development status", None),
            ("destination_code", "Receiver ISO3"), ("destination_name", "Receiver"),
            ("Receiver WESP region", None), ("Receiver WESP subregion", None), ("Receiver development status", None),
            ("firm", "Provider"), ("firm_type", "Provider type"), ("payment instrument", "Payment instrument"),
            ("access point", "Access point"), ("pickup method", "Pickup method"), ("speed actual", "Speed"),
            ("transparent", "Transparent"), (C200, "Total cost % (USD 200)"), (C500, "Total cost % (USD 500)")]
    out = d[[c for c, _ in cols]]
    wb = xlsxwriter.Workbook(str(ROOT / FULL_FILE))
    hdr = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "#FFFFFF", "text_wrap": True, "valign": "top"})
    ws = wb.add_worksheet("Services")
    ws.write_row(0, 0, [h or c for c, h in cols], hdr)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(out), len(cols) - 1)
    for i, row in enumerate(out.itertuples(index=False), start=1):
        for j, v in enumerate(row):
            if v is None or (isinstance(v, float) and math.isnan(v)):
                continue
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.write_number(i, j, float(v))
            else:
                ws.write_string(i, j, str(v))
    ab = wb.add_worksheet("About")
    ab.set_column(0, 0, 18)
    ab.set_column(1, 1, 110)
    notes = [
        ("Source", "World Bank, Remittance Prices Worldwide (RPW), sheet 'Dataset (from Q2 2016)'. License CC BY 4.0."),
        ("What a row is", "A price quote for one service for a hypothetical USD 200 and USD 500 transfer, collected by "
                          "mystery shopping. Not actual transactions; not volume-weighted."),
        ("Official average", "Simple average of 'Total cost % (USD 200)' across rows with Transparent = yes."),
        ("Cleaning", "Text-stored numbers converted; Transparent flag lower-cased; costs rounded to 4 decimals."),
        ("Regions", "UN WESP Annex. Where a country has no subregion, its region is used."),
    ]
    for i, (k, v) in enumerate(notes):
        ab.write(i, 0, k)
        ab.write(i, 1, v)
    wb.close()
    print(f"Wrote {FULL_FILE}: {len(out)} rows")


def main(xlsx):
    wesp = {r["ISO"]: r for r in csv.DictReader(open(ROOT / "data/wesp_regions.csv", encoding="utf-8"))}

    def info(code):
        w = wesp.get(code)
        if not w or not w["region"]:
            return ["Unclassified", "Unclassified", "Unclassified"]
        return [w["region"], w["subregion"] or w["region"], w["development_status"] or "Unclassified"]

    d = pd.read_excel(xlsx, sheet_name=SHEET)
    d.columns = [str(c).strip() for c in d.columns]
    d = d.loc[:, ~d.columns.str.startswith("Unnamed")]
    for c in (C200, C500):
        d[c] = pd.to_numeric(d[c].astype(str).str.replace(",", ".", regex=False).str.strip(), errors="coerce")
    d["transparent"] = d["transparent"].astype(str).str.strip().str.lower()
    write_full_excel(d, info)
    d = d[d["transparent"] == "yes"].copy()
    d["ft"] = d["firm_type"].map(provider_type)
    d["pi"] = d["payment instrument"].map(instrument)

    periods = quarter_range(sorted(d["period"].unique())[0], sorted(d["period"].unique())[-1])
    pidx = {p: i for i, p in enumerate(periods)}
    names = {}
    for code, name in list(d[["source_code", "source_name"]].drop_duplicates().values) + \
            list(d[["destination_code", "destination_name"]].drop_duplicates().values):
        names[code] = name
    senders = sorted(d["source_code"].unique())
    receivers = sorted(d["destination_code"].unique())
    si = {c: i for i, c in enumerate(senders)}
    di = {c: i for i, c in enumerate(receivers)}

    g = d.groupby(["period", "source_code", "destination_code", "ft", "pi"]).agg(
        n1=(C200, "count"), s1=(C200, "sum"), n2=(C500, "count"), s2=(C500, "sum"), mn=(C200, "min")).reset_index()
    rows = [[pidx[r.period], si[r.source_code], di[r.destination_code], PROVIDER_TYPES.index(r.ft),
             INSTRUMENTS.index(r.pi), int(r.n1), round(float(r.s1), 3), int(r.n2), round(float(r.s2), 3),
             round(float(r.mn), 2) if r.n1 else None] for r in g.itertuples()]

    out = {
        "periods": periods,
        "senders": [[c, names[c]] + info(c) for c in senders],
        "receivers": [[c, names[c]] + info(c) for c in receivers],
        "countryFields": ["code", "name", "region", "subregion", "status"],
        "providerTypes": PROVIDER_TYPES,
        "instruments": INSTRUMENTS,
        "fields": ["period", "sender", "receiver", "providerType", "instrument",
                   "n200", "sum200", "n500", "sum500", "min200"],
        "rows": rows,
        "source": "World Bank, Remittance Prices Worldwide (RPW). License: CC BY 4.0. Regions: UN WESP Annex.",
    }
    js = "window.RPW_DATA=" + json.dumps(out, separators=(",", ":"), ensure_ascii=False) + ";\n"
    (ROOT / "data/rpw_costs.js").write_text(js, encoding="utf-8")
    print(f"Wrote data/rpw_costs.js: {len(rows)} groups, {len(periods)} quarters")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Usage: python scripts/build_data.py path/to/rpw_dataset.xlsx")
    main(sys.argv[1])
