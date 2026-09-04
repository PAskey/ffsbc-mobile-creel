# Marquart Lake Mobile Creel — prototype

A phone-friendly, offline-capable web form (PWA) for anglers to self-report their
catch, plus a script that turns the collected reports into the FFSBC `SURVEY.xlsx`
upload template. Single lake for now (Marquart); adding lakes later is just a new
QR code carrying that lake's constants.

## Data flow

```
Angler scans QR at the lake
        │
        ▼
  index.html (this PWA)  ──►  saves each report to the phone (works with no signal)
        │                          │
        │  when back in signal     ▼
        └──────────────►  central store (Azure/SharePoint/Sheet)  ── you choose
                                   │
                                   ▼
             build_survey_workbook.py  ──►  SURVEY_filled.xlsx  ──►  FFSBC submission site
```

The golden rule: **collect flat, transform later.** The form never tries to write
the multi-sheet template directly — it stores one simple record per report, and the
script assembles the relational workbook when you're ready to upload.

## What's in this folder

| File | What it is |
|---|---|
| `index.html` | The angler form. Self-contained; opens on any phone. |
| `sw.js`, `manifest.webmanifest` | Make it installable + work offline (need to be *hosted*, not opened as a local file). |
| `icon-192.png`, `icon-512.png` | App icons for the "Add to Home Screen" install. |
| `build_survey_workbook.py` | Turns exported reports → filled `SURVEY.xlsx`. |
| `sample_reports.json` | 3 example reports (incl. a "caught nothing" trip). |
| `SURVEY_filled.xlsx` | The template filled from those 3 reports — open it to see the result. |

## The form (index.html)

Auto-filled from the lake (angler never types these): waterbody, event name
`00848LNIC_2026_1_S`, project `FFSBC46_23`, `survey_type=CR`, `trip_complete=COMPLETE`,
`location_description=Mobile_App`, date & time, and a unique `party_id`.

Date fished auto-fills to today but is editable (for trips reported later); it can't
be a future date. Angler enters "how were you fishing?" as three tiered chip controls
— Vessel (Shore/dock, Boat, Single-person craft, Ice — **Ice only appears Nov–Apr**
by the date fished) → Method (Troll and/or Cast; picking **both records `MIX`**;
hidden and set to Ice automatically when the vessel is Ice) → Gear type (Flyfishing /
Lures / Baits, multi-select capped at two) — plus # anglers and # rods (both 1–6),
hours (0.5–12), years fishing (required), a required primary target species (a "no
preference" choice records `TR` = trout general and greys out the secondary), an
optional secondary target, then their catch. The catch list is pre-seeded with the two stocked species
(Rainbow + Eastern Brook) so the angler just enters counts; an "Add another species"
button covers anything unexpected (e.g. an invasive fish). Target dropdowns list the
stocked species first, then all others A–Z. Vessel→fishing_mode, Method→angling_method,
Gear→terminal_gear (mapped to the FFSBC code).

Data-quality guards: number fields strip leading zeros and clamp to their limits;
number caught can't exceed 10× hours fished (clamped as you type, and rejected on
submit if entered out of order); number kept can't exceed number caught.

Each report also carries a **device_id** — a random, non-personal token stored in the
phone's browser the first time the app loads. It lets you spot when reports on
different days likely came from the same device (dedup / repeat-angler signal). It is
NOT written to SURVEY.xlsx (no column for it); it stays in the collected data and the
CSV export for your own analysis. It resets if the browser's storage is cleared, is
per-browser, and can't tell apart two people sharing one phone.

Reports are saved in the browser (`localStorage`) so nothing is lost with no signal.
"Organizer tools" at the bottom exports all stored reports as JSON (feeds the script)
or CSV.

### Try it locally
```bash
cd marquart-creel-pwa
python3 -m http.server 8099
# open http://localhost:8099 on your computer, or your phone on the same wifi
```
(Opening `index.html` as a `file://` won't register the service worker or save
reliably — serve it over http, or host it.)

### Host it (so the QR code works + offline install works)
Any static host works. Given you're on Microsoft/Azure already:
**Azure Static Web Apps (free tier)** — point it at this folder. Also fine: any
web server, GitHub Pages, Netlify. Once hosted at e.g. `https://…/marquart`, make
the QR code point there.

### Wire up the backend (optional, for auto-collection)
In `index.html` set `const ENDPOINT = "..."` to your collection URL (an Azure
Function, a Power Automate/SharePoint endpoint, or a Google Apps Script web app).
Reports POST there and are marked sent; until then they just queue on the phone and
you gather them with the JSON export. Leave `ENDPOINT` empty to run export-only.

### Per-lake QR codes (later)
Each lake gets its own hosted copy (or a `?lake=` parameter) with that lake's
constants in the `EVENT` block, and its own QR. The QR just encodes that lake's URL.

## The transform (build_survey_workbook.py)

```bash
python build_survey_workbook.py sample_reports.json --template SURVEY.xlsx --out SURVEY_filled.xlsx
```
**party_id is assigned here, not in the app.** FFSBC requires party_id and person_id
to be integers 1–500, with the tally restarting each day at a lake. The transform sorts
the reports by date+time and numbers them 1, 2, 3… within each `interview_date` (person_id
is 1 — one self-reporting angler per party). This is why the app can't assign it: two
phones can't coordinate who is "number 7". The app's own `client_ref` and `device_id` are
kept only in the flat data. If a single day ever exceeds 500 reports, the script warns you
to split that day's upload.

Fills **Party** (one row per report), **Angler** (one row per report — person_id,
angling_method, terminal_gear, years_fishing), and **Count** (one row per species,
linked by party_id + person_id). Leaves **IndividualFish** and **ExtendedSurvey** blank. Never
touches SubmissionDetail, AssessEvent, or the reference tabs. Your pre-filled Party
row-2 constants are read and used as defaults. Note: angling_method and terminal_gear
are written to the **Angler** tab (per respondent); Party's copies of those two
columns are left blank — tell me if you'd rather mirror them onto Party too.

## Known limits / production to-dos

- **Template fidelity.** Re-saving with openpyxl drops Excel's *data-validation* and
  *conditional-formatting* extensions on the green tabs, and clears cached formula
  values (AssessEvent columns C & F). Opening the file in Excel once recomputes them
  (already done for the included `SURVEY_filled.xlsx` via a recalc step). If the FFSBC
  site rejects it for a structural change, the robust fix is to (a) do this same fill
  in **R with `openxlsx`** — keeps you in your stack and preserves the template better
  — or (b) inject rows only into the Party/Count sheet XML so every other part of the
  file stays byte-identical. Worth validating one output against
  https://ffsbcsubmit.azurewebsites.net/ before going live.
- **`party_id` (1–500, reset per day) is assigned by the transform; `person_id` is 1.**
  One self-reporting angler per party. Cross-day "same person" tracking is the separate
  `device_id`, which lives only in the collected data / CSV.
- **Junk/spam data.** Public self-report — the transform is the quality gate. Easy to
  add: validate species codes, sanity-cap counts/hours, drop empties.
- **Gear is multi-select but capped at two**, because FFSBC's `terminal_gear` codes
  only cover singles and pairs (FLY, LURE, BAIT, FLY-LURE, BAIT-FLY, BAIT-LR) — there's
  no code for all three. Once two are picked the third is disabled with a note. Every
  record therefore maps to a valid code.

## Prototype status
v0.1 — form, offline queue, export, and template transform all working end-to-end
and verified. Not yet hosted, no live backend wired, not yet validated against the
FFSBC submission site.
