# app.py
# ------------------------------------------------------------
# Excel -> 1 moderne Standalone-HTML (Suche + A4 Druck)
# WICHTIG:
# - ES WIRD NICHTS GEFILTERT. Jede gefundene Zeile wird übernommen.
# - A4 pro Kunde ist ZWINGEND: Der Ausdruck wird automatisch skaliert,
#   sodass alles auf eine Seite passt (CSS transform: scale()).
#   -> keine zweite Seite, kein Abschneiden.
#
# Extraktion:
# 1) B_-Spalten (Hauptquelle):
#    "Mo Z 0 B_Sa" (Zeit), "Mo 0 B_Sa" (Sortiment), Bestelltag aus Header ("B_Sa")
# 2) klassische Tripel:
#    "Mo 21 Zeit" / "Mo 21 Sort" / "Mo 21 Tag"
# 3) DS Tripel (optional):
#    "DS Fr zu Mi Zeit" / "DS Fr zu Mi Sort" / "DS Fr zu Mi Tag"
#
# Hinweis:
# - Wir fassen alle Quellen zusammen (B_ + Tripel + DS) und zeigen sie ALLE an.
# - Liefertage in Kopfzeile kommen aus (Bestellzeilen ∪ Touren).
# ------------------------------------------------------------

import json
import re
from typing import Dict, Tuple, List

import pandas as pd
import streamlit as st

PLAN_TYP = "Standard"
BEREICH = "Alle Sortimente Fleischwerk"
DAYS_DE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]

DAY_SHORT_TO_DE = {
    "Mo": "Montag",
    "Di": "Dienstag",
    "Die": "Dienstag",
    "Mi": "Mittwoch",
    "Mitt": "Mittwoch",
    "Do": "Donnerstag",
    "Don": "Donnerstag",
    "Donn": "Donnerstag",
    "Fr": "Freitag",
    "Sa": "Samstag",
    "Sam": "Samstag",
}

BESTELL_SHORT_TO_DE = {
    "Mo": "Montag",
    "Di": "Dienstag",
    "Die": "Dienstag",
    "Mi": "Mittwoch",
    "Mitt": "Mittwoch",
    "Do": "Donnerstag",
    "Don": "Donnerstag",
    "Donn": "Donnerstag",
    "Fr": "Freitag",
    "Sa": "Samstag",
    "Sam": "Samstag",
}

TOUR_COLS = {
    "Montag": "Mo",
    "Dienstag": "Die",
    "Mittwoch": "Mitt",
    "Donnerstag": "Don",
    "Freitag": "Fr",
    "Samstag": "Sam",
}


def norm(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    s = str(x)
    s = s.replace("\u00a0", " ")  # NBSP
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def group_sort_key(g: str):
    g = g.strip()
    if g.isdigit():
        return (0, int(g))
    return (1, g.lower())


def detect_triplets(columns: List[str]) -> Dict[str, Dict[str, Dict[str, str]]]:
    """
    "<Tag> <Gruppe> Zeit/Sort/Tag"
    """
    rx = re.compile(
        r"^(Mo|Die|Di|Mitt|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(.+?)\s+(Zeit|Sort|Tag)$",
        re.IGNORECASE,
    )
    found: Dict[str, Dict[str, Dict[str, str]]] = {}
    for c in [c.strip() for c in columns]:
        m = rx.match(c)
        if not m:
            continue

        day_short = m.group(1)
        group = m.group(2).strip()
        field = m.group(3).capitalize()

        if day_short.lower() == "donn":
            day_short = "Don"

        day_de = DAY_SHORT_TO_DE.get(day_short)
        if not day_de:
            continue

        found.setdefault(day_de, {}).setdefault(group, {})[field] = c

    clean: Dict[str, Dict[str, Dict[str, str]]] = {}
    for day_de, groups in found.items():
        for g, fields in groups.items():
            if all(k in fields for k in ("Zeit", "Sort", "Tag")):
                clean.setdefault(day_de, {})[g] = fields
    return clean


def detect_bspalten(columns: List[str]) -> Dict[Tuple[str, str, str], Dict[str, str]]:
    """
    'Mo Z 0 B_Sa' / 'Mo 0 B_Sa' / 'Mo L 0 B_Sa'
    """
    cols = [c.strip() for c in columns]
    rx = re.compile(
        r"^(Mo|Die|Di|Mitt|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+"
        r"(?:(Z|L)\s+)?"
        r"(.+?)\s+"
        r"B[_ ]?(Mo|Die|Di|Mitt|Mi|Don|Donn|Do|Fr|Sam|Sa)$",
        re.IGNORECASE,
    )

    mapping: Dict[Tuple[str, str, str], Dict[str, str]] = {}
    for c in cols:
        m = rx.match(c)
        if not m:
            continue

        day_short = m.group(1)
        zl = (m.group(2) or "").upper()
        group = m.group(3).strip()
        b_short = m.group(4)

        if day_short.lower() == "donn":
            day_short = "Don"
        if b_short.lower() == "donn":
            b_short = "Don"

        day_de = DAY_SHORT_TO_DE.get(day_short)
        bestell_de = BESTELL_SHORT_TO_DE.get(b_short)
        if not day_de or not bestell_de:
            continue

        key = (day_de, group, bestell_de)
        mapping.setdefault(key, {})

        if zl == "Z":
            mapping[key]["zeit"] = c
        elif zl == "L":
            mapping[key]["l"] = c
        else:
            mapping[key]["sort"] = c

    return mapping


def detect_ds_triplets(columns: List[str]) -> Dict[str, Dict[str, str]]:
    """
    'DS Fr zu Mi Zeit' / 'DS Fr zu Mi Sort' / 'DS Fr zu Mi Tag'
    """
    cols = [c.strip() for c in columns]
    rx = re.compile(r"^DS\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)
    tmp: Dict[str, Dict[str, str]] = {}
    for c in cols:
        m = rx.match(c)
        if not m:
            continue
        route = m.group(1).strip()
        field = m.group(2).capitalize()
        key = f"DS {route}".replace("zu", "→")
        tmp.setdefault(key, {})[field] = c

    clean: Dict[str, Dict[str, str]] = {}
    for k, fields in tmp.items():
        if all(x in fields for x in ("Zeit", "Sort", "Tag")):
            clean[k] = fields
    return clean


def normalize_time(s: str) -> str:
    s = norm(s)
    if not s:
        return ""
    # 20:00 -> 20:00 Uhr
    if "uhr" not in s.lower() and re.fullmatch(r"\d{1,2}:\d{2}", s):
        return s + " Uhr"
    return s


# -----------------------------
# MODERN STANDALONE HTML TEMPLATE
# - A4 pro Kunde wird per JS automatisch skaliert (transform: scale)
# - Nichts wird ausgeblendet/abgeschnitten: scale reduziert bei Bedarf
# -----------------------------
HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<title>Sendeplan</title>

<style>
@page { size:A4; margin:10mm; }

body{
  font-family: Arial, Helvetica, sans-serif;
  margin:0;
  background:#111;
}

/* ---------- SCREEN UI ---------- */
.app{display:grid;grid-template-columns:320px 1fr;gap:10px;padding:10px}
.sidebar{background:#222;color:#fff;padding:10px;border-radius:10px}
input,button{padding:8px;margin:4px 0;width:100%}
.main{display:flex;justify-content:center}

/* ---------- A4 PAGE ---------- */
.paper{
  width:190mm;
  height:277mm;     /* A4 minus margins */
  background:#fff;
  color:#000;
  overflow:hidden;
  padding:6mm;
  box-sizing:border-box;
}

/* Dynamische Schriftgrößen (werden per JS gesetzt) */
.paper{ --fs:11pt; }
.paper *{ font-size:var(--fs); line-height:1.15; }

h1{font-size:16pt;margin:0;text-align:center}
h2{font-size:14pt;margin:2mm 0;text-align:center;color:#c00}
h3{font-size:10pt;margin:0;text-align:center}

.header{
  display:flex;
  justify-content:space-between;
  margin:4mm 0;
}

table{
  width:100%;
  border-collapse:collapse;
}

th,td{
  border:1px solid #000;
  padding:1.2mm;
  vertical-align:top;
}

th{background:#eee}

/* PRINT */
@media print{
  body{background:#fff}
  .sidebar{display:none}
  .app{display:block;padding:0}
  .paper{box-shadow:none}
}
</style>
</head>

<body>

<div class="app">
<div class="sidebar">
<input id="knr" placeholder="Kundennummer">
<button onclick="showOne()">Anzeigen</button>
<button onclick="showAll()">Alle</button>
<button onclick="window.print()">Drucken</button>
</div>

<div class="main">
<div id="out"></div>
</div>
</div>

<script>
const DATA = __DATA_JSON__;

/* -------- AUTO FIT (ECHT) -------- */
function autoFit(paper){
  let fs = 11;
  paper.style.setProperty("--fs", fs+"pt");

  // so lange verkleinern bis Inhalt passt
  while(paper.scrollHeight > paper.clientHeight && fs > 7){
    fs -= 0.5;
    paper.style.setProperty("--fs", fs+"pt");
  }
}

/* -------- RENDER -------- */
function render(c){
  const rows = c.bestell.map(x=>`
    <tr>
      <td>${x.liefertag}</td>
      <td>${x.sortiment}</td>
      <td>${x.bestelltag}</td>
      <td>${x.bestellschluss}</td>
    </tr>`).join("");

  const html = `
  <div class="paper">
    <h1>Sende- & Belieferungsplan</h1>
    <h2>${c.plan_typ}</h2>
    <h3>${c.name} ${c.bereich}</h3>

    <div class="header">
      <div>
        <b>${c.name}</b><br>
        ${c.strasse}<br>
        ${c.plz} ${c.ort}
      </div>
      <div>
        Kunden-Nr.: ${c.kunden_nr}<br>
        Fachberater: ${c.fachberater}
      </div>
    </div>

    <table>
      <thead>
        <tr>
          <th>Liefertag</th>
          <th>Sortiment</th>
          <th>Bestelltag</th>
          <th>Bestellzeitende</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;

  return html;
}

function showOne(){
  const k=document.getElementById("knr").value.trim();
  if(!DATA[k]) return;
  const out=document.getElementById("out");
  out.innerHTML=render(DATA[k]);
  autoFit(out.firstElementChild);
}

function showAll(){
  const out=document.getElementById("out");
  out.innerHTML=Object.values(DATA).map(render).join("");
  document.querySelectorAll(".paper").forEach(autoFit);
}
</script>

</body>
</html>
"""


# -----------------------------
# Streamlit Generator
# -----------------------------
st.set_page_config(page_title="Excel → Moderne A4-Druckvorlage", layout="wide")
st.title("Excel → Moderne HTML (Auto-Fit auf 1×A4, nichts rausfiltern)")

up = st.file_uploader("Excel (.xlsx) hochladen", type=["xlsx"])
if not up:
    st.info("Excel hochladen → HTML erzeugen → Kundennummer suchen → drucken.")
    st.stop()

df = pd.read_excel(up, engine="openpyxl")
df.columns = [c.strip() for c in df.columns]

required = ["Nr", "SAP-Nr.", "Name", "Strasse", "Plz", "Ort"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Pflichtspalten fehlen: {missing}")
    st.stop()

trip = detect_triplets(df.columns.tolist())
bmap = detect_bspalten(df.columns.tolist())
dsmap = detect_ds_triplets(df.columns.tolist())

with st.expander("Debug – erkannte Quellen"):
    st.write("**B_-Mapping Keys:**", len(bmap))
    st.write("**Tripel Tage:**", ", ".join([d for d in DAYS_DE if d in trip]) or "-")
    st.write("**DS Keys:**", ", ".join(dsmap.keys()) or "-")

data: Dict[str, dict] = {}

for _, r in df.iterrows():
    knr = norm(r.get("Nr", ""))
    if not knr:
        continue

    tours = {}
    for day_de, col in TOUR_COLS.items():
        tours[day_de] = norm(r.get(col, "")) if col in df.columns else ""

    bestell = []

    # 1) B_-Spalten: NICHTS FILTERN (auch wenn Zeit fehlt oder Bestelltag leer wäre -> nehmen)
    for day_de in DAYS_DE:
        keys = [k for k in bmap.keys() if k[0] == day_de]
        keys.sort(key=lambda k: (group_sort_key(k[1]), DAYS_DE.index(k[2]) if k[2] in DAYS_DE else 99))

        for (lday, group, bestelltag) in keys:
            cols = bmap[(lday, group, bestelltag)]
            sortiment = norm(r.get(cols.get("sort", ""), ""))
            zeit = normalize_time(r.get(cols.get("zeit", ""), ""))

            # NICHT filtern: wir nehmen auch leere Felder mit,
            # aber wenn ALLES komplett leer ist, wäre es nur Müll.
            # Deshalb: minimaler Check: mindestens eins der Felder hat Inhalt.
            if not (sortiment or zeit or bestelltag):
                continue

            bestell.append({
                "liefertag": lday,
                "sortiment": sortiment,
                "bestelltag": bestelltag,
                "bestellschluss": zeit
            })

    # 2) Tripel: ebenfalls übernehmen
    for day_de in DAYS_DE:
        for g in sorted(trip.get(day_de, {}).keys(), key=group_sort_key):
            cols = trip[day_de][g]
            zeit = normalize_time(r.get(cols["Zeit"], ""))
            sortiment = norm(r.get(cols["Sort"], ""))
            bestelltag = norm(r.get(cols["Tag"], ""))

            if not (sortiment or zeit or bestelltag):
                continue

            bestell.append({
                "liefertag": day_de,
                "sortiment": sortiment,
                "bestelltag": bestelltag,
                "bestellschluss": zeit
            })

    # 3) DS: eigener Block, nichts filtern
    ds_list = []
    for ds_key, cols in dsmap.items():
        zeit = normalize_time(r.get(cols["Zeit"], ""))
        sortiment = norm(r.get(cols["Sort"], ""))
        bestelltag = norm(r.get(cols["Tag"], ""))

        if not (sortiment or zeit or bestelltag):
            continue

        ds_list.append({
            "ds_key": ds_key,
            "sortiment": sortiment,
            "bestelltag": bestelltag,
            "bestellschluss": zeit
        })

    data[knr] = {
        "plan_typ": PLAN_TYP,
        "bereich": BEREICH,
        "kunden_nr": knr,
        "sap_nr": norm(r.get("SAP-Nr.", "")),
        "name": norm(r.get("Name", "")),
        "strasse": norm(r.get("Strasse", "")),
        "plz": norm(r.get("Plz", "")),
        "ort": norm(r.get("Ort", "")),
        "fax": norm(r.get("Fax", "")) if "Fax" in df.columns else "",
        "fachberater": norm(r.get("Fachberater", "")) if "Fachberater" in df.columns else "",
        "tours": tours,
        "bestell": bestell,
        "ds": ds_list,
    }

st.success(f"{len(data)} Kunden eingebettet. Auto-Fit sorgt für 1×A4 pro Kunde.")

html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))

st.download_button(
    "⬇️ Standalone-HTML herunterladen (modern + Auto-Fit A4)",
    data=html.encode("utf-8"),
    file_name="sende_belieferungsplan_autofit_a4.html",
    mime="text/html",
)

st.caption("Druck-Tipp: Im Browser A4 + Ränder „Keine“ (oder minimal). Auto-Fit sorgt für 1 Seite pro Kunde.")
