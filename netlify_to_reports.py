#!/usr/bin/env python3
"""
netlify_to_reports.py
---------------------
Convert a Netlify Forms CSV export of the "creel-report" form into the reports
JSON that build_survey_workbook.py consumes. Each app submission is stored by
Netlify with the full record in a "payload" column (JSON); this pulls those out.

Usage:
  python netlify_to_reports.py submissions.csv --out reports.json
  python build_survey_workbook.py reports.json --template SURVEY.xlsx --out SURVEY_filled.xlsx
"""
import argparse, csv, json, sys

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("csv_file", help="CSV exported from Netlify Forms")
    ap.add_argument("--out", default="reports.json")
    ap.add_argument("--payload-col", default="payload")
    args = ap.parse_args()

    reports, skipped = [], 0
    with open(args.csv_file, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        cols = reader.fieldnames or []
        if args.payload_col not in cols:
            sys.exit(f"No '{args.payload_col}' column found. CSV columns: {cols}")
        for row in reader:
            raw = (row.get(args.payload_col) or "").strip()
            if not raw:
                skipped += 1; continue
            try:
                reports.append(json.loads(raw))
            except json.JSONDecodeError:
                skipped += 1
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(reports, f, indent=1)
    print(f"Wrote {args.out}: {len(reports)} report(s); skipped {skipped} row(s) with no/invalid payload.")

if __name__ == "__main__":
    main()
