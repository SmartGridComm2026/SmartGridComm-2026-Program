# IEEE SmartGridComm 2026 — Program Page (draft)

Web program for IEEE SmartGridComm 2026, October 26–29, 2026, College Station, Texas.
The layout follows the style of the [2026 IEEE WF-PST program](https://ieee-wfpst.org/program/).

**To review the page, open `index.html` in a browser.** Nothing needs to be installed.

## Files

| File | Purpose |
|---|---|
| `index.html` | The built program page. Generated; do not edit by hand. |
| `template.html` | Layout, styling, and page script. Edit this to change how the page looks. |
| `build_program.py` | Reads the planning spreadsheet and fills the template. Keynotes, panels, breaks, chairs, and times are defined in this file. |
| `requirements.txt` | Python dependency (`openpyxl`). |

## Rebuilding

The planning spreadsheet is **not** in this repository because it contains contact
emails and internal notes. Rebuilding needs a local copy.

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python build_program.py path/to/"IEEESmartGridComm2026 Live Planning Document.xlsx"
```

With no argument, the script looks for the spreadsheet in the parent folder.

### Where the content comes from

| Content | Source |
|---|---|
| Paper sessions: session titles, paper IDs, titles, authors, rooms | Spreadsheet tabs *Web Version Main Program-Sessio* and *Web Version Workshop-Only Sessi* |
| Session chairs | `CHAIRS` in `build_program.py` (copied from the *Session-Paper Map* tab) |
| Keynotes, panels, welcomes, breaks, receptions, times | `DAYS` in `build_program.py` (from the *Web Version Detailed Program* tab) |
| Header, dates, footer text | `CONF` in `template.html` |

## Open questions for reviewers

The spreadsheet disagrees with itself in these places. The page currently uses the
choice shown in **bold**.

1. **Joint Paper Session 1** (Communications & Networking + Power Line Communications) is dated
   Thursday 10:15 in both session tabs, which would double-book Ross. The overview tab puts it on
   **Tuesday 14:45 in Ross**.
2. The second breakout room is "Laurel (CL/24-26)" in the overview tab header and **"Traditions"** everywhere else.
3. Day 3 welcome speaker is listed as **"Dr. Arum Han"** and as "Dr. Hahn".
4. Workshop 1's first paper session is timed 09:00–12:30 but labeled "PM 1". The page follows the
   **AM 1 = keynotes, AM 2 = panel, PM 1 / PM 2 = paper sessions** labels.
5. Day 2 welcome is listed as "Dr. Reddy" and as **"Dr. Davis, Dr. Reddy"** (Dr. Davis's first name is not given).
6. Thursday morning sessions end at **12:30** in the overview and 12:15 for Data Analytics Session 4.

### Still to be announced

Keynote talk titles, panelists for both panels, reception and gala dinner locations,
RELLIS tour, and chairs for seven sessions (Data Analytics 4; Control & Operation 1–5; Joint 2).
