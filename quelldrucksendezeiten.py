# app.py
# ------------------------------------------------------------
# Excel -> Standalone-HTML (Suche + A4 Druck, 1 Seite pro Kunde)
# WICHTIG:
# - Es wird NICHTS gefiltert (nur komplett leere Datensätze werden übersprungen).
# - Pro Kunde GENAU 1×A4: Schriftgröße wird automatisch reduziert bis es passt.
# - Deutsche See & alle anderen Sortimente zählen für die Wochentage (Liefertag-Zeile).
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
    s = str(x).replace("\u00a0", " ").strip()
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
# DRUCKSTABILES HTML TEMPLATE
# - Paper = echtes A4 (210x297mm)
# - @page margin 0
# - Auto-Fit reduziert Schrift bis es passt
# - Wochentage kommen aus Bestellzeilen (inkl. Deutsche See)
# -----------------------------
HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Sende- & Belieferungsplan</title>

<style>
  @page { size: A4; margin: 0; }

  :root{
    --bg:#0b1220;
    --panel:rgba(255,255,255,.08);
    --stroke:rgba(255,255,255,.14);
    --text:rgba(255,255,255,.92);
    --muted:rgba(255,255,255,.62);
    --paper:#fff;
    --ink:#0b0f17;
    --sub:#394054;
  }

  *{ box-sizing:border-box; }
  body{
    margin:0;
    font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial, Helvetica, sans-serif;
    background: radial-gradient(900px 520px at 25% 0%, rgba(79,163,255,.18), transparent 55%),
                radial-gradient(900px 520px at 80% 10%, rgba(124,247,194,.14), transparent 55%),
                var(--bg);
    color:var(--text);
  }

  .app{ display:grid; grid-template-columns: 340px 1fr; gap:14px; padding:14px; }
  .sidebar,.main{
    background: linear-gradient(180deg, var(--panel), rgba(255,255,255,.06));
    border:1px solid var(--stroke);
    border-radius:16px;
    overflow:hidden;
    backdrop-filter: blur(14px);
  }
  .sidehead{ padding:14px; border-bottom:1px solid var(--stroke); }
  .sidehead .t{ font-weight:900; letter-spacing:.2px; }
  .sidehead .s{ margin-top:6px; font-size:12px; color:var(--muted); line-height:1.35; }
  .controls{ padding:14px; display:flex; flex-direction:column; gap:10px; }
  .field{
    display:flex; gap:10px; align-items:center;
    padding:10px 12px; border-radius:14px;
    border:1px solid var(--stroke);
    background: rgba(0,0,0,.18);
  }
  input{
    width:100%; border:none; outline:none; background:transparent;
    color:var(--text); font-size:14px;
  }
  input::placeholder{ color: rgba(255,255,255,.45); }
  .btnrow{ display:flex; gap:10px; flex-wrap:wrap; }
  button{
    border:none; cursor:pointer;
    padding:10px 12px; border-radius:14px;
    font-weight:900; font-size:13px; letter-spacing:.2px;
    color:var(--text);
    background: rgba(255,255,255,.08);
    border:1px solid var(--stroke);
  }
  button.primary{
    background: linear-gradient(135deg, rgba(79,163,255,.30), rgba(124,247,194,.18));
    border-color: rgba(124,247,194,.30);
  }
  .hint{ font-size:12px; color:var(--muted); }

  .list{ border-top:1px solid var(--stroke); max-height: calc(100vh - 240px); overflow:auto; }
  .item{ padding:12px 14px; border-bottom:1px solid rgba(255,255,255,.06); cursor:pointer; }
  .item:hover{ background: rgba(255,255,255,.06); }
  .item.active{ background: rgba(79,163,255,.12); border-left:3px solid rgba(79,163,255,.85); padding-left:11px; }
  .k{ font-size:12px; color:var(--muted); font-weight:800; }
  .n{ font-size:13px; font-weight:950; line-height:1.2; }
  .a{ font-size:12px; color:var(--muted); }

  .mainhead{ padding:14px; border-bottom:1px solid var(--stroke); display:flex; justify-content:space-between; gap:10px; }
  .mh1{ font-weight:950; }
  .mh2{ font-size:12px; color:var(--muted); }
  .wrap{ padding:14px; display:flex; justify-content:center; }

  /* --- A4 Paper --- */
  .paper{
    width:210mm;
    height:297mm;
    background:var(--paper);
    color:var(--ink);
    border-radius:14px;
    box-shadow: 0 30px 90px rgba(0,0,0,.45);
    overflow:hidden;
    position:relative;
    --fs: 10.4pt;
  }
  .paper *{ font-size: var(--fs); line-height: 1.15; }

  .inner{
    padding: 12mm 12mm 10mm 12mm; /* Druckrand */
  }

  .ptitle{
    text-align:center;
    font-weight:950;
    font-size: 17pt;
    margin:0;
  }
  .pstd{
    text-align:center;
    font-weight:1000;
    color:#d0192b;
    font-size: 15pt;
    margin: 2mm 0 1.5mm 0;
  }
  .psub{
    text-align:center;
    font-weight:800;
    color:var(--sub);
    margin:0 0 6mm 0;
  }

  .head{
    display:flex; justify-content:space-between; gap:12mm;
    margin-bottom: 5mm;
  }
  .addr .name{ font-weight:950; margin-bottom:1mm; }
  .meta b{ font-weight:950; }

  .lines{ margin: 0 0 5mm 0; }
  .lines b{ font-weight:950; }

  table{ width:100%; border-collapse:collapse; }
  th,td{ border:1px solid #111; padding:1.6mm 1.6mm; vertical-align:top; }
  th{ background:#f3f6fb; font-weight:1000; }
  .cday{ width:17%; font-weight:1000; }
  .csort{ width:52%; }
  .ctag{ width:16%; white-space:nowrap; }
  .ctime{ width:15%; white-space:nowrap; }

  .ds{ margin-top: 5mm; border:1px solid rgba(0,0,0,.22); border-radius:10px; overflow:hidden; }
  .dsh{ background:#f3f6fb; padding:6px 8px; font-weight:1000; }
  .dsb{ padding:6px 8px; }

  @media print{
    body{ background:#fff; }
    .sidebar,.mainhead{ display:none !important; }
    .app{ display:block; padding:0; }
    .wrap{ padding:0; justify-content:flex-start; }
    .paper{
      border-radius:0;
      box-shadow:none;
      page-break-after: always;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
  }

  @media (max-width: 980px){
    .app{ grid-template-columns:1fr; }
    .paper{ width:100%; height:auto; }
  }
</style>
</head>

<body>
<div class="app">
  <div class="sidebar">
    <div class="sidehead">
      <div class="t">Sende- & Belieferungsplan</div>
      <div class="s">Wochentage oben kommen aus Bestellzeilen (inkl. Deutsche See).</div>
    </div>
    <div class="controls">
      <div class="field">
        <input id="knr" placeholder="Kundennummer (z.B. 88130)" inputmode="numeric">
      </div>
      <div class="btnrow">
        <button class="primary" onclick="showOne()">Anzeigen</button>
        <button onclick="showAll()">Alle rendern</button>
        <button onclick="window.print()">Drucken</button>
      </div>
      <div class="hint" id="hint"></div>
    </div>
    <div class="list" id="list"></div>
  </div>

  <div class="main">
    <div class="mainhead">
      <div>
        <div class="mh1" id="mh1">Vorschau</div>
        <div class="mh2" id="mh2">Noch kein Kunde ausgewählt.</div>
      </div>
      <div style="text-align:right">
        <div class="mh1" id="cnt"></div>
        <div class="mh2">Kunden</div>
      </div>
    </div>
    <div class="wrap">
      <div id="out" class="hint">Links Kundennummer eingeben oder aus Liste wählen.</div>
    </div>
  </div>
</div>

<script>
const DATA = __DATA_JSON__;
const ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0)-(Number(b)||0));
const DAYS = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];

function esc(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;");
}

/* Wochentage aus Bestellzeilen (inkl. Deutsche See) + optional Touren */
function buildActiveDays(c){
  const fromBestell = new Set();
  for(const it of (c.bestell || [])){
    const d = String(it.liefertag || "").trim();
    const s = String(it.sortiment || "").trim();
    const t = String(it.bestelltag || "").trim();
    const z = String(it.bestellschluss || "").trim();
    if(d && DAYS.includes(d) && (s || t || z)){
      fromBestell.add(d);
    }
  }
  const fromTours = new Set(DAYS.filter(d => (c.tours && String(c.tours[d]||"").trim() !== "")));
  return DAYS.filter(d => fromBestell.has(d) || fromTours.has(d));
}

function render(c){
  const byDay = {};
  for(const it of (c.bestell||[])){
    if(!byDay[it.liefertag]) byDay[it.liefertag]=[];
    byDay[it.liefertag].push(it);
  }

  const rows = DAYS.map(d=>{
    const arr = byDay[d] || [];
    const s = arr.map(x=>esc(x.sortiment)).join("<br>");
    const t = arr.map(x=>esc(x.bestelltag)).join("<br>");
    const z = arr.map(x=>esc(x.bestellschluss)).join("<br>");
    return `
      <tr>
        <td class="cday">${esc(d)}</td>
        <td class="csort">${s || "&nbsp;"}</td>
        <td class="ctag">${t || "&nbsp;"}</td>
        <td class="ctime">${z || "&nbsp;"}</td>
      </tr>`;
  }).join("");

  const active = buildActiveDays(c);
  const dayLine = active.length ? active.join(" ") : "-";
  const tourLine = active.length
    ? active.map(d => (c.tours && String(c.tours[d]||"").trim()!=="" ? esc(c.tours[d]) : "—")).join(" ")
    : "-";

  let dsHtml = "";
  if(c.ds && c.ds.length){
    dsHtml = `
      <div class="ds">
        <div class="dsh">Durchsteck (DS)</div>
        <div class="dsb">
          ${c.ds.map(it => `<div><b>${esc(it.ds_key)}:</b> ${esc(it.sortiment)} — <b>${esc(it.bestelltag)}</b> · ${esc(it.bestellschluss)}</div>`).join("")}
        </div>
      </div>
    `;
  }

  return `
    <div class="paper">
      <div class="inner">
        <div class="ptitle">Sende- &amp; Belieferungsplan</div>
        <div class="pstd">${esc(c.plan_typ || "Standard")}</div>
        <div class="psub">${esc(c.name)} ${esc(c.bereich || "")}</div>

        <div class="head">
          <div class="addr">
            <div class="name">${esc(c.name)}</div>
            <div>${esc(c.strasse)}</div>
            <div>${esc(c.plz)} ${esc(c.ort)}</div>
          </div>
          <div class="meta">
            <div><b>Kunden-Nr.:</b> ${esc(c.kunden_nr)}</div>
            <div><b>Fachberater:</b> ${esc(c.fachberater || "")}</div>
          </div>
        </div>

        <div class="lines">
          <div><b>Liefertag:</b> ${dayLine}</div>
          <div><b>Tour:</b> ${tourLine}</div>
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

        ${dsHtml}
      </div>
    </div>
  `;
}

/* Auto-Fit: reduziert --fs bis inner passt */
function autoFitPaper(paper){
  if(!paper) return;
  const inner = paper.querySelector(".inner");
  if(!inner) return;

  let fs = 10.4;
  const minFs = 6.8;
  paper.style.setProperty("--fs", fs + "pt");

  for(let pass=0; pass<3; pass++){
    while(inner.scrollHeight > paper.clientHeight && fs > minFs){
      fs -= 0.2;
      paper.style.setProperty("--fs", fs.toFixed(2) + "pt");
    }
  }
}

function autoFitAll(){
  document.querySelectorAll(".paper").forEach(autoFitPaper);
}

function showOne(){
  const k = document.getElementById("knr").value.trim();
  const c = DATA[k];
  const out = document.getElementById("out");
  if(!c){
    out.innerHTML = "<div class='hint'>Kundennummer nicht gefunden.</div>";
    document.getElementById("hint").textContent = `Vorhanden: ${ORDER.length} Kunden`;
    return;
  }
  out.innerHTML = render(c);
  document.getElementById("mh1").textContent = `Vorschau: ${c.kunden_nr}`;
  document.getElementById("mh2").textContent = `${c.name} · ${c.plz} ${c.ort}`;
  document.getElementById("hint").textContent = "Auto-Fit aktiv. Drucken = 1 Seite.";
  setActive(k);
  setTimeout(autoFitAll, 10);
}

function showAll(){
  const out = document.getElementById("out");
  out.innerHTML = ORDER.map(k => render(DATA[k])).join("");
  document.getElementById("mh1").textContent = "Massendruck";
  document.getElementById("mh2").textContent = `${ORDER.length} Kunden gerendert`;
  document.getElementById("hint").textContent = "Auto-Fit aktiv. Jede Seite = 1×A4.";
  setActive(null);
  setTimeout(autoFitAll, 10);
}

function setActive(key){
  document.querySelectorAll(".item").forEach(el=>{
    el.classList.toggle("active", el.dataset.key === key);
  });
}

function buildList(){
  const list = document.getElementById("list");
  list.innerHTML = ORDER.map(k=>{
    const c = DATA[k];
    return `
      <div class="item" data-key="${esc(k)}" onclick="pick('${esc(k)}')">
        <div class="k">Kunden-Nr. ${esc(c.kunden_nr)}</div>
        <div class="n">${esc(c.name)}</div>
        <div class="a">${esc(c.plz)} ${esc(c.ort)}</div>
      </div>
    `;
  }).join("");
}

function pick(k){
  document.getElementById("knr").value = k;
  showOne();
}

document.getElementById("knr").addEventListener("keydown",(e)=>{
  if(e.key==="Enter") showOne();
});

document.getElementById("cnt").textContent = ORDER.length;
buildList();

/* vor dem Drucken nochmal Auto-Fit */
window.addEventListener("beforeprint", () => { autoFitAll(); });
</script>
</body>
</html>
"""


# -----------------------------
# Streamlit Generator
# -----------------------------
st.set_page_config(page_title="Excel → A4-Druckvorlage", layout="wide")
st.title("Excel → HTML (Auto-Fit: 1×A4 pro Kunde, Deutsche See in Wochentagen)")

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

    # 1) B_-Spalten
    for day_de in DAYS_DE:
        keys = [k for k in bmap.keys() if k[0] == day_de]
        keys.sort(key=lambda k: (group_sort_key(k[1]), DAYS_DE.index(k[2]) if k[2] in DAYS_DE else 99))
        for (lday, group, bestelltag) in keys:
            cols = bmap[(lday, group, bestelltag)]
            sortiment = norm(r.get(cols.get("sort", ""), ""))
            zeit = normalize_time(r.get(cols.get("zeit", ""), ""))
            if not (sortiment or zeit or bestelltag):
                continue
            bestell.append({
                "liefertag": lday,
                "sortiment": sortiment,
                "bestelltag": bestelltag,
                "bestellschluss": zeit
            })

    # 2) Tripel
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

    # 3) DS
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

st.success(f"{len(data)} Kunden eingebettet. Wochentage kommen aus Bestellzeilen (inkl. Deutsche See).")

html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))

st.download_button(
    "⬇️ Standalone-HTML herunterladen (A4 safe + Deutsche See in Tagen)",
    data=html.encode("utf-8"),
    file_name="sende_belieferungsplan_A4_safe_deutschesee.html",
    mime="text/html",
)

st.caption("Druck-Tipp: A4 auswählen. Auto-Fit läuft auch vor dem Drucken und erzwingt 1 Seite.")
