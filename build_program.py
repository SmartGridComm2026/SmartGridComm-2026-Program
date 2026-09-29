"""Build the SmartGridComm 2026 program page from the live planning spreadsheet.

Usage: python build_program.py [planning.xlsx]
  The spreadsheet defaults to the planning document in the parent folder.
  Writes index.html (standalone page to open in a browser) and
  build/artifact.html (the same page without <head>, for publishing as an artifact).

Paper sessions come from the "Web Version" session sheets; the day skeleton
(keynotes, breaks, panels) is defined below from "Web Version Detailed Program".
"""
import json, re, sys
from pathlib import Path
import openpyxl

HERE = Path(__file__).resolve().parent
xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent / "IEEESmartGridComm2026 Live Planning Document.xlsx"
template = HERE / "template.html"
wb = openpyxl.load_workbook(xlsx, data_only=True)

TRACKS = {  # session-name prefix -> (kind, long label, short label)
    "Data Analytics": ("dac", "Data Analytics & Computation", "Data Analytics"),
    "Control & Operation": ("co", "Control & Operation", "Control & Ops"),
    "Communications": ("cn", "Communications & Networking", "Comms & Networking"),
    "Cyber-Physical": ("cps", "Cyber-Physical Security & Privacy", "Security & Privacy"),
    "Joint": ("joint", "Joint Session", "Joint"),
    "Workshop": ("workshop", "Workshop", "Workshop"),
}
# Chairs from "Session-Paper Map" (emails and planning notes removed)
CHAIRS = {
    "DAC 1": "Ying Zhang (Oklahoma State University)",
    "DAC 2": "Md Zahidul Islam (Southern Illinois University)",
    "DAC 3": "Yuzhang Lin (New York University)",
    "CN 1": "Ioannis (Yannis) Zografopoulos · Co-chair: Anurag Srivastava",
    "Joint 1": "Daisuke Mashima · Co-chair: Alexandru Ștefanov",
    "CPS 1": "Daisuke Mashima (Singapore University of Technology and Design)",
    "CPS 2": "Kevin Jin (University of Arkansas)",
    "CPS 3": "Taesic Kim (University of Missouri)",
    "CPS 4": "Mostafa Mohammadpourfard (Texas Tech University)",
}
CODE = {"dac": "DAC", "co": "CO", "cn": "CN", "cps": "CPS", "joint": "Joint"}

def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def parse_papers(text):
    groups, lines, i = [], [clean(l) for l in str(text).split("\n")], 0
    while i < len(lines):
        l = lines[i]
        m = re.match(r"^(\d{10}):\s*(.*)$", l)
        if re.match(r"^(Workshop )?Session Title:", l):
            groups.append({"title": l.split(":", 1)[1].strip(), "papers": []})
        elif m:
            title = m.group(2).strip().strip('"').strip()
            j = i + 1
            while j < len(lines) and not lines[j]:
                j += 1
            authors = lines[j] if j < len(lines) else ""
            if not groups:
                groups.append({"title": "", "papers": []})
            groups[-1]["papers"].append([m.group(1), title, authors])
            i = j
        i += 1
    return [g for g in groups if g["papers"]]

sessions = {}
def load(sheet):
    ws = wb[sheet]
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r[3] or not r[4]:
            continue
        name = clean(r[4])
        prefix = next(k for k in TRACKS if name.startswith(k) or name.startswith("Workshop"))
        kind, long, short = TRACKS[prefix if not name.startswith("Workshop") else "Workshop"]
        num = re.search(r"(?:Session|PM)\s*(\d+)", name)
        key = f"{CODE.get(kind, 'WS')} {num.group(1)}" if kind != "workshop" else name
        sessions[key] = {"kind": kind, "long": long, "short": short, "num": num.group(1) if num else "",
                         "name": name, "room": clean(r[2]), "groups": parse_papers(r[3])}
load("Web Version Main Program-Sessio")
load("Web Version Workshop-Only Sessi")

def paper_session(key, room=None):
    s = sessions[key]
    n = sum(len(g["papers"]) for g in s["groups"])
    if s["kind"] == "joint":
        label = f"Joint Session {s['num']} · " + s["name"].split(" - ", 1)[1].replace("  ", " ")
        label = re.sub(r"\s*\(\d+ papers?\)", "", label)
        title = " · ".join(g["title"] for g in s["groups"])
    else:
        label = f"{s['long']} · Session {s['num']}"
        title = s["groups"][0]["title"]
    return {"kind": s["kind"], "label": label, "room": room or s["room"], "title": title,
            "chair": ("Chair: " + CHAIRS[key]) if key in CHAIRS else "Chair: to be announced",
            "meta": f"{n} papers", "groups": s["groups"],
            "gk": f"{s['short']} {s['num']}" if s["kind"] != "joint" else f"Joint {s['num']}",
            "glance": title}

def ws_groups(name):
    return sessions[name]["groups"]

C, T, R, K = "Century I & II", "Traditions", "Ross", "Corps"
REG = "Pre-Function I & II"
def brk(title, room=None): return {"kind": "break", "title": title, **({"room": room} if room else {})}
def slot(t, items, note=None):
    a, b = [x.strip() for x in re.split(r"[–-]", t)]
    return {"start": a, "end": b, "note": note, "items": items}

W1 = "Cyber-Physical Power System Resilience"
DAYS = [
 {"tab": "Day 1", "short": "Mon · Oct 26", "title": "Day 1 — Monday, October 26",
  "sub": "Workshops & tutorials · Registration opens 8:00 AM · Welcome reception 6:00 PM", "slots": [
  slot("08:00 – 17:00", [{"kind": "special", "label": "Registration", "room": REG, "title": "Main Registration Desk Open"}], "all day"),
  slot("09:00 – 10:30", [
    {"kind": "workshop", "label": "Workshop 1 · Part 1 of 4", "room": C, "title": W1, "meta": "Keynotes: two speakers, 30-minute talk + 15-minute Q&A each", "glance": W1},
    {"kind": "tutorial", "label": "Tutorial 2 · Part 1 of 2", "room": T, "title": "Agentic AI for Active Distribution Networks"},
    {"kind": "tutorial", "label": "Tutorial 3 · Part 1 of 2", "room": R, "title": "LLM-Powered Agentic AI"},
    {"kind": "tutorial", "label": "Tutorial 1 · Part 1 of 2", "room": K, "title": "Quantum Secure 6G"}], "morning"),
  slot("10:30 – 11:00", [brk("Morning Coffee & Networking")]),
  slot("11:00 – 12:30", [
    {"kind": "workshop", "label": "Workshop 1 · Part 2 of 4", "room": C, "title": W1, "meta": "Panel: four speakers, 12-minute talk + 3-minute Q&A each, then a 30-minute discussion"},
    {"kind": "tutorial", "label": "Tutorial 2 · Part 2 of 2", "room": T, "title": "Agentic AI for Active Distribution Networks"},
    {"kind": "tutorial", "label": "Tutorial 3 · Part 2 of 2", "room": R, "title": "LLM-Powered Agentic AI"},
    {"kind": "tutorial", "label": "Tutorial 1 · Part 2 of 2", "room": K, "title": "Quantum Secure 6G"}]),
  slot("12:30 – 13:30", [brk("Lunch, Networking & Industry Exhibits", C)]),
  slot("13:30 – 15:30", [
    {"kind": "workshop", "label": "Workshop 1 · Part 3 of 4", "room": C, "title": "Cybersecurity, AI, and Threat Detection",
     "meta": f"{W1} · Includes Workshop 2, Agentic Energy Systems in Smart Grids (merged)",
     "groups": ws_groups("Workshop 1 - Cyber-Physical Power System Resilience PM 1")},
    {"kind": "workshop", "label": "Workshop 3 · Part 1 of 2", "room": K, "title": "Digital Twin for Smart Grid", "meta": "Keynote / panel"}], "afternoon"),
  slot("15:30 – 16:00", [brk("Afternoon Coffee Break")]),
  slot("16:00 – 17:30", [
    {"kind": "workshop", "label": "Workshop 1 · Part 4 of 4", "room": C, "title": "Power System Resilience and Recovery",
     "meta": W1, "groups": ws_groups("Workshop 1 - Cyber-Physical Power System Resilience PM 2")},
    {"kind": "workshop", "label": "Workshop 3 · Part 2 of 2", "room": K, "title": "Digital Twins", "meta": "Digital Twin for Smart Grid",
     "groups": ws_groups("Workshop 3 - Digital Twin for Smart Grid PM 2")}]),
  slot("18:00 – 20:00", [{"kind": "social", "label": "Networking", "title": "Welcome Reception", "meta": "Location to be announced", "tbd": True}]),
 ]},
 {"tab": "Day 2", "short": "Tue · Oct 27", "title": "Day 2 — Tuesday, October 27",
  "sub": "Opening, keynotes & parallel paper sessions · Gala dinner 7:00 PM", "slots": [
  slot("08:00 – 17:00", [{"kind": "special", "label": "Registration", "room": REG, "title": "Registration Desk Open"}], "all day"),
  slot("08:45 – 09:00", [{"kind": "special", "label": "Welcome", "room": C, "title": "Welcome & Opening Remarks",
     "people": [["Speakers", "Dr. Davis & Dr. Narasimha Reddy", "Texas A&M University"]]}]),
  slot("09:00 – 10:00", [{"kind": "keynote", "label": "Keynote I · Opening Keynote", "room": C, "title": "Dr. Vince Poor",
     "people": [["Affiliation", "Princeton University", ""]], "meta": "Talk title to be announced", "gk": "Keynote I", "glance": "Dr. Vince Poor, Princeton University"}]),
  slot("10:00 – 10:15", [brk("Coffee Break & Poster Session I")]),
  slot("10:15 – 12:15", [paper_session("DAC 1"), paper_session("CO 1"), paper_session("CN 1"), paper_session("CPS 1")], "parallel"),
  slot("12:15 – 13:30", [brk("Lunch", C)]),
  slot("13:30 – 14:30", [{"kind": "keynote", "label": "Keynote II · Afternoon Keynote", "room": C, "title": "Dr. Tom Overbye",
     "people": [["Affiliation", "Texas A&M University", ""]], "meta": "Talk title to be announced", "gk": "Keynote II", "glance": "Dr. Tom Overbye, Texas A&M University"}]),
  slot("14:30 – 14:45", [brk("Afternoon Coffee Break")]),
  slot("14:45 – 17:15", [paper_session("DAC 2"), paper_session("CO 2"), paper_session("Joint 1", R), paper_session("CPS 2")], "parallel"),
  slot("19:00 – 22:00", [{"kind": "social", "label": "Networking", "title": "Conference Prefunction & Gala Dinner", "meta": "Location to be announced", "tbd": True, "glance": "Gala Dinner"}]),
 ]},
 {"tab": "Day 3", "short": "Wed · Oct 28", "title": "Day 3 — Wednesday, October 28",
  "sub": "Keynote, industry panels & parallel paper sessions", "slots": [
  slot("08:30 – 16:00", [{"kind": "special", "label": "Registration", "room": REG, "title": "Registration Desk Open"}], "all day"),
  slot("08:45 – 09:00", [{"kind": "special", "label": "Welcome", "room": C, "title": "Day 3 Welcome", "people": [["Speaker", "Dr. Arum Han", ""]]}]),
  slot("09:00 – 10:00", [{"kind": "keynote", "label": "Keynote III · Technical Keynote", "room": C, "title": "Woody Rickerson",
     "people": [["Affiliation", "ERCOT", ""]], "meta": "Talk title to be announced", "gk": "Keynote III", "glance": "Woody Rickerson, ERCOT"}]),
  slot("10:00 – 10:15", [brk("Coffee Break & Poster Session II")]),
  slot("10:15 – 12:15", [{"kind": "panel", "label": "Panel 1 · Organized by OPAL-RT", "room": C,
     "title": "Understanding the Role of Hardware-in-the-Loop Testing in Supporting Reliable Data Center Interconnections and Grid Operations",
     "meta": "Panelists to be announced", "glance": "Hardware-in-the-loop testing for data center interconnections"}]),
  slot("12:15 – 13:30", [brk("Networking Lunch", C)]),
  slot("13:30 – 14:30", [{"kind": "panel", "label": "Panel 2 · Industry Keynote & Roundtable", "room": C, "title": "G3-Alliance Panel",
     "meta": "Fireside-chat format · Panelists to be announced"}]),
  slot("14:30 – 14:45", [brk("Afternoon Coffee Break")]),
  slot("14:45 – 17:15", [paper_session("DAC 3"), paper_session("CO 3"), paper_session("CO 4"), paper_session("CPS 3")], "parallel"),
 ]},
 {"tab": "Day 4", "short": "Thu · Oct 29", "title": "Day 4 — Thursday, October 29",
  "sub": "Keynote, final paper sessions, awards & closing · Optional tours 2:00 PM", "slots": [
  slot("08:30 – 12:00", [{"kind": "special", "label": "Registration", "room": REG, "title": "Registration Desk Open (morning only)"}], "morning"),
  slot("08:45 – 09:00", [{"kind": "special", "label": "Welcome", "room": C, "title": "Day 4 Welcome"}]),
  slot("09:00 – 10:00", [{"kind": "keynote", "label": "Keynote IV · Morning Keynote", "room": C, "title": "Dr. Veronica Adetola",
     "people": [["Affiliation", "Pacific Northwest National Laboratory (PNNL)", ""]], "meta": "Talk title to be announced", "gk": "Keynote IV", "glance": "Dr. Veronica Adetola, PNNL"}]),
  slot("10:00 – 10:15", [brk("Morning Coffee Break")]),
  slot("10:15 – 12:30", [paper_session("DAC 4"), paper_session("CO 5"), paper_session("Joint 2"), paper_session("CPS 4")], "parallel"),
  slot("12:30 – 14:00", [{"kind": "special", "label": "Lunch & Closing Session", "room": C,
     "title": "Best Paper Awards, Closing & SmartGridComm 2027 Preview", "glance": "Best Paper Awards & Closing"}]),
  slot("14:00 – 15:30", [{"kind": "special", "label": "Optional Tours", "title": "Walking Tour & Technical Tours",
     "meta": "Walking tour confirmed · RELLIS campus tour to be confirmed"}]),
 ]},
]

used = {k for k in sessions if not k.startswith("Workshop")}
placed = sum(1 for d in DAYS for s in d["slots"] for i in s["items"] if i.get("gk", "").split(" ")[0] in
             ("Data", "Control", "Comms", "Security", "Joint"))
papers = sum(len(g["papers"]) for d in DAYS for s in d["slots"] for i in s["items"] for g in i.get("groups", []))
print(f"sessions parsed: {len(sessions)}, paper sessions placed: {placed}/{len(used)}, papers on page: {papers}")

page = template.read_text(encoding="utf-8").replace("/*__DAYS__*/[]", json.dumps(DAYS, ensure_ascii=False, indent=0))
(HERE / "build").mkdir(exist_ok=True)
(HERE / "build" / "artifact.html").write_text(page, encoding="utf-8")
head = ('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n')
style_end = page.index("</style>") + len("</style>")
standalone = head + page[:style_end] + "\n</head>\n<body>\n" + page[style_end:] + "\n</body>\n</html>\n"
(HERE / "index.html").write_text(standalone, encoding="utf-8")
print(f"wrote {HERE / 'index.html'} and build/artifact.html")
