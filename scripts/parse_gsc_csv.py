#!/usr/bin/env python3
"""Converts a Google Search Console "Performance report" CSV export into the
exact shape build_report.py's GSC tab expects. Use this instead of the OAuth
API path when Google Cloud Console access isn't available — no admin
permission needed, just the normal Search Console UI you already use.

How to export (in Search Console):
  1. Open the property (https://search.google.com/search-console)
  2. Performance report
  3. Add a filter: Page > exact URL > the target page for this keyword
  4. Set the date range (e.g. last 6 months)
  5. Make sure the table below is grouped by "Query" (Queries tab)
  6. Click "Export" (top right) > Download CSV

Column names vary by UI language (English "Query"/"Clicks"/... vs French
"Requêtes"/"Clics"/...), so this looks for several known variants,
case-insensitively.

Usage:
  python parse_gsc_csv.py --csv Queries.csv --out gsc.csv
"""

import argparse
import csv
import sys

FIELD_VARIANTS = {
    "query": ["query", "queries", "top queries", "requête", "requêtes", "requetes"],
    "clicks": ["clicks", "clics"],
    "impressions": ["impressions"],
    "ctr": ["ctr"],
    "position": ["position", "avg. position", "average position", "position moyenne"],
}


def find_column(fieldnames_lower, variants):
    for v in variants:
        if v in fieldnames_lower:
            return fieldnames_lower[v]
    return None


def to_float(value):
    if value is None:
        return 0.0
    value = str(value).strip().replace("%", "").replace(",", ".")
    if value in ("", "-"):
        return 0.0
    try:
        return float(value)
    except ValueError:
        return 0.0


def main():
    parser = argparse.ArgumentParser(description="Convert a GSC Performance report CSV export to build_report.py's GSC CSV shape")
    parser.add_argument("--csv", required=True, help="GSC export CSV path")
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    with open(args.csv, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            sys.exit("CSV has no header row — is this a real Search Console export?")
        fieldnames_lower = {name.strip().lower(): name for name in reader.fieldnames}
        col = {key: find_column(fieldnames_lower, variants) for key, variants in FIELD_VARIANTS.items()}
        if not col["query"]:
            sys.exit(f"Could not find a 'Query' column. Columns found: {reader.fieldnames}")

        rows = []
        for raw_row in reader:
            query = (raw_row.get(col["query"]) or "").strip()
            if not query:
                continue
            ctr_raw = raw_row.get(col["ctr"]) if col["ctr"] else None
            ctr = to_float(ctr_raw)
            # GSC UI exports CTR already as a percentage (e.g. "12.3%") — the
            # sheet builder expects a 0-1 fraction like the API does, so
            # normalize back down when the export gave us a percentage.
            if ctr > 1:
                ctr = ctr / 100
            rows.append({
                "query": query,
                "clicks": to_float(raw_row.get(col["clicks"])) if col["clicks"] else 0,
                "impressions": to_float(raw_row.get(col["impressions"])) if col["impressions"] else 0,
                "ctr": ctr,
                "position": to_float(raw_row.get(col["position"])) if col["position"] else 0,
            })

    with open(args.out, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["query", "clicks", "impressions", "ctr", "position"])
        for r in rows:
            writer.writerow([r["query"], r["clicks"], r["impressions"], r["ctr"], r["position"]])

    print(f"Saved: {args.out} — {len(rows)} requêtes converties")
    print("Colonnes détectées :", {k: v for k, v in col.items() if v})


if __name__ == "__main__":
    main()
