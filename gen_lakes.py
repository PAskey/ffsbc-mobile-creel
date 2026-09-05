#!/usr/bin/env python3
"""
gen_lakes.py — regenerate the app's lake list from SURVEY.xlsx.
The AssessEvent tab is the master list. This reads it and rewrites the LAKES
config in index.html (+ site/index.html), sets DEFAULT_LAKE to the first row,
and bumps the service-worker cache so installed phones pick up the change.

Annual/whenever workflow:
  1) edit the AssessEvent tab in SURVEY.xlsx (add/remove lakes; stocked species
     codes go comma-separated in the `comment` column, e.g. "EB, RB")
  2) run:  python gen_lakes.py
  3) push to GitHub (Netlify redeploys)

Lakes listed here but never given a QR simply sit unused — harmless.
"""
import re, shutil, sys, datetime, openpyxl

TEMPLATE = "../SURVEY.xlsx"
NAME_OVERRIDES = {"MCCONNELL": "McConnell"}   # add special capitalisation here as needed

def disp(wbname):
    w = str(wbname or "").strip()
    return NAME_OVERRIDES.get(w.upper(), " ".join(p.capitalize() for p in w.split()))

def main():
    wb = openpyxl.load_workbook(TEMPLATE, data_only=True)
    ws = wb["AssessEvent"]
    hdr = {str(c.value).strip(): i for i, c in enumerate(ws[1]) if c.value is not None}
    for col in ("waterbody_id", "waterbody_name", "assess_event_name", "comment"):
        if col not in hdr:
            sys.exit(f"AssessEvent is missing a '{col}' column.")
    lakes = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        wid = r[hdr["waterbody_id"]]
        if not wid:
            continue
        stocked = [x.strip().upper() for x in str(r[hdr["comment"]] or "").split(",") if x.strip()]
        lakes.append((str(wid).strip(), disp(r[hdr["waterbody_name"]]),
                      str(r[hdr["assess_event_name"]]).strip(), stocked))
    if not lakes:
        sys.exit("No lakes found in the AssessEvent tab.")

    rows = []
    for wid, name, event, stocked in lakes:
        sj = "[" + ",".join('"%s"' % c for c in stocked) + "]"
        rows.append(f'  "{wid}": {{ name:"{name}", event:"{event}", stocked:{sj} }}')
    obj = "const LAKES = {\n" + ",\n".join(rows) + "\n};"
    default = lakes[0][0]
    stamp = datetime.datetime.utcnow().strftime("%Y%m%d-%H%M%S")

    # index.html: swap the LAKES object and DEFAULT_LAKE
    p = "index.html"; s = open(p, encoding="utf-8").read()
    s2 = re.sub(r"const LAKES = \{[\s\S]*?\n\};", obj, s, count=1)
    s2 = re.sub(r'const DEFAULT_LAKE = "[^"]*";', f'const DEFAULT_LAKE = "{default}";', s2, count=1)
    if s2 == s:
        sys.exit("Could not find the LAKES/DEFAULT_LAKE markers in index.html.")
    open(p, "w", encoding="utf-8").write(s2)

    # sw.js: bump cache so the redeploy reaches installed apps
    swp = "sw.js"; sw = open(swp, encoding="utf-8").read()
    sw2 = re.sub(r'const CACHE = "[^"]*";', f'const CACHE = "ffsbc-creel-{stamp}";', sw, count=1)
    open(swp, "w", encoding="utf-8").write(sw2)

    # sync deploy copies
    shutil.copyfile(p, "site/index.html")
    shutil.copyfile(swp, "site/sw.js")

    print(f"Wrote {len(lakes)} lakes; default = {default} ({lakes[0][1]}).")
    print("Lakes:", ", ".join(f"{n}" for _, n, _, _ in lakes))
    print(f"Service-worker cache -> ffsbc-creel-{stamp}")
    print("Updated index.html + site/. Now push to GitHub to redeploy.")

if __name__ == "__main__":
    main()
