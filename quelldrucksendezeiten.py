# app.py
# ------------------------------------------------------------
# Excel -> 1 moderne HTML-Datei (Standalone)
# - In der HTML: Kundennummer suchen/eingeben, Vorschau, A4 drucken
# - Massendruck: "Alle rendern" + Drucken
#
# Extraktion (WICHTIG):
# 1) B_-Spalten (Hauptquelle, damit ALLE Sortimente wie im PDF kommen)
#    z.B. "Mo Z 0 B_Sa" = Zeit, "Mo 0 B_Sa" = Sortiment, Bestelltag aus "B_Sa"
# 2) 21 Tripel (Zusatzquelle): "Mo 21 Zeit/Sort/Tag"
# 3) DS Tripel (optional): "DS Fr zu Mi Zeit/Sort/Tag" -> eigener Block
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
    s = str(x).strip()
    s = re.sub(r"\s+", " ", s)
    # Excel-Nummern (Tour) als 1001.0 -> 1001
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def group_sort_key(g: str):
    g = g.strip()
    if g.isdigit():
        return (0, int(g))
    return (1, g.lower())


def detect_triplets_21(columns: List[str]) -> Dict[str, Dict[str, Dict[str, str]]]:
    """
    Erkennen: "<Tag> <Gruppe> Zeit/Sort/Tag"
    z.B. "Mo 21 Zeit" / "Mo 21 Sort" / "Mo 21 Tag"
    """
    rx = re.compile(
        r"^(Mo|Die|Di|Mitt|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(.+?)\s+(Zeit|Sort|Tag)$",
        re.IGNORECASE,
    )
    found = {}
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

    clean = {}
    for day_de, groups in found.items():
        for g, fields in groups.items():
            if all(k in fields for k in ("Zeit", "Sort", "Tag")):
                clean.setdefault(day_de, {})[g] = fields
    return clean


def detect_bspalten(columns: List[str]) -> Dict[Tuple[str, str, str], Dict[str, str]]:
    """
    Erkennen:
      'Mo Z 0 B_Sa'  -> Zeit
      'Mo 0 B_Sa'    -> Sortiment
      'Mo L 0 B_Sa'  -> ignorieren (optional)
    Bestelltag aus B_Sa/B_Fr/B_Don/... (aus Header!)
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
    Erkennen (optional):
      'DS Fr zu Mi Zeit' / 'DS Fr zu Mi Sort' / 'DS Fr zu Mi Tag'
    Liefertag ist dann z.B. 'DS Fr→Mi'
    """
    cols = [c.strip() for c in columns]
    rx = re.compile(r"^DS\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)

    tmp = {}
    for c in cols:
        m = rx.match(c)
        if not m:
            continue
        route = m.group(1).strip()   # z.B. "Fr zu Mi"
        field = m.group(2).capitalize()
        key = f"DS {route}".replace("zu", "→")
        tmp.setdefault(key, {})[field] = c

    clean = {}
    for k, fields in tmp.items():
        if all(x in fields for x in ("Zeit", "Sort", "Tag")):
            clean[k] = fields
    return clean


# -----------------------------
# MODERN STANDALONE HTML TEMPLATE
# (Keine f-Strings => keine { } Probleme)
# -----------------------------
HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Sende- & Belieferungsplan</title>

<style>
  :root{
    --bg: #0b1220;
    --panel: rgba(255,255,255,0.08);
    --panel2: rgba(255,255,255,0.06);
    --stroke: rgba(255,255,255,0.16);
    --text: rgba(255,255,255,0.92);
    --muted: rgba(255,255,255,0.62);
    --accent: #4fa3ff;
    --accent2: #7cf7c2;
    --danger: #ff5b6e;
    --paper: #ffffff;
    --ink: #0c0f16;
    --ink2: #32384a;
  }

  *{ box-sizing: border-box; }
  body{
    margin:0;
    font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial, "Helvetica Neue", Helvetica, sans-serif;
    background: radial-gradient(1200px 600px at 20% 0%, rgba(79,163,255,0.22), transparent 55%),
                radial-gradient(900px 500px at 80% 10%, rgba(124,247,194,0.18), transparent 55%),
                radial-gradient(900px 700px at 40% 100%, rgba(255,91,110,0.12), transparent 55%),
                var(--bg);
    color: var(--text);
  }

  /* App layout */
  .app{
    min-height: 100vh;
    display: grid;
    grid-template-columns: 360px 1fr;
    gap: 18px;
    padding: 18px;
  }

  .sidebar, .main{
    border: 1px solid var(--stroke);
    background: linear-gradient(180deg, var(--panel), var(--panel2));
    border-radius: 18px;
    backdrop-filter: blur(14px);
    -webkit-backdrop-filter: blur(14px);
    overflow: hidden;
  }

  .side-head{
    padding: 16px 16px 10px 16px;
    border-bottom: 1px solid var(--stroke);
  }
  .title{
    font-size: 16px;
    font-weight: 800;
    letter-spacing: .2px;
    display:flex;
    align-items:center;
    gap:10px;
  }
  .badge{
    font-size: 12px;
    font-weight: 700;
    padding: 4px 10px;
    border-radius: 999px;
    background: rgba(79,163,255,0.16);
    border: 1px solid rgba(79,163,255,0.30);
    color: rgba(255,255,255,0.88);
  }
  .subtitle{
    margin-top: 6px;
    font-size: 12px;
    color: var(--muted);
    line-height: 1.35;
  }

  .search{
    padding: 14px 16px 16px 16px;
    display:flex;
    flex-direction: column;
    gap:10px;
  }

  .field{
    display:flex;
    align-items:center;
    gap:10px;
    padding: 10px 12px;
    border-radius: 14px;
    border: 1px solid var(--stroke);
    background: rgba(0,0,0,0.18);
  }
  .field input{
    width:100%;
    border:none;
    outline:none;
    background: transparent;
    color: var(--text);
    font-size: 14px;
  }
  .field input::placeholder{ color: rgba(255,255,255,0.45); }

  .btnrow{ display:flex; gap:10px; flex-wrap:wrap; }
  .btn{
    border:none;
    outline:none;
    cursor:pointer;
    font-weight: 800;
    letter-spacing: .2px;
    font-size: 13px;
    padding: 10px 12px;
    border-radius: 14px;
    color: var(--text);
    background: rgba(255,255,255,0.08);
    border: 1px solid var(--stroke);
    transition: transform .08s ease, background .12s ease;
  }
  .btn:hover{ background: rgba(255,255,255,0.12); }
  .btn:active{ transform: translateY(1px); }
  .btn.primary{
    background: linear-gradient(135deg, rgba(79,163,255,0.35), rgba(124,247,194,0.22));
    border: 1px solid rgba(124,247,194,0.32);
  }
  .btn.danger{
    background: rgba(255,91,110,0.14);
    border: 1px solid rgba(255,91,110,0.28);
  }

  .list{
    border-top: 1px solid var(--stroke);
    max-height: calc(100vh - 220px);
    overflow:auto;
  }
  .item{
    padding: 12px 16px;
    border-bottom: 1px solid rgba(255,255,255,0.06);
    cursor:pointer;
    display:flex;
    flex-direction: column;
    gap:4px;
  }
  .item:hover{ background: rgba(255,255,255,0.06); }
  .item.active{
    background: rgba(79,163,255,0.14);
    border-left: 3px solid rgba(79,163,255,0.85);
    padding-left: 13px;
  }
  .item .k{
    font-size: 12px;
    color: var(--muted);
    font-weight: 700;
  }
  .item .n{
    font-size: 13px;
    font-weight: 850;
    color: rgba(255,255,255,0.92);
    line-height: 1.25;
  }
  .item .a{
    font-size: 12px;
    color: var(--muted);
    line-height: 1.25;
  }

  .main-head{
    padding: 16px;
    border-bottom: 1px solid var(--stroke);
    display:flex;
    align-items:center;
    justify-content: space-between;
    gap: 12px;
  }
  .main-head .meta{
    display:flex;
    flex-direction: column;
    gap: 4px;
  }
  .main-head .meta .big{
    font-size: 14px;
    font-weight: 900;
    letter-spacing: .2px;
  }
  .main-head .meta .small{
    font-size: 12px;
    color: var(--muted);
  }

  .paper-wrap{
    padding: 16px;
    display:flex;
    justify-content: center;
  }

  /* Paper preview */
  .paper{
    width: 210mm;
    min-height: 297mm;
    background: var(--paper);
    color: var(--ink);
    border-radius: 14px;
    box-shadow: 0 30px 90px rgba(0,0,0,0.45);
    overflow:hidden;
    border: 1px solid rgba(0,0,0,0.10);
  }
  .paper-inner{
    padding: 14mm 14mm 12mm 14mm;
  }

  .p-title{
    text-align:center;
    font-size: 18pt;
    font-weight: 900;
    margin: 0;
    letter-spacing: .2px;
  }
  .p-standard{
    text-align:center;
    font-size: 16pt;
    font-weight: 950;
    color: #d0192b;
    margin: 2mm 0 2mm 0;
  }
  .p-sub{
    text-align:center;
    font-size: 10.5pt;
    margin: 0 0 7mm 0;
    color: var(--ink2);
    font-weight: 700;
  }

  .p-head{
    display:flex;
    justify-content: space-between;
    gap: 14mm;
    margin-bottom: 6mm;
  }
  .p-addr{
    font-size: 10.8pt;
    line-height: 1.35;
  }
  .p-addr .name{
    font-weight: 900;
    margin-bottom: 1mm;
  }
  .p-meta{
    font-size: 10.8pt;
    line-height: 1.35;
    min-width: 70mm;
  }
  .p-meta b{ font-weight: 900; }

  .p-lines{
    font-size: 10.8pt;
    line-height: 1.4;
    margin: 0 0 6mm 0;
    color: var(--ink);
  }
  .p-lines b{ font-weight: 900; }

  table{
    width:100%;
    border-collapse: collapse;
    font-size: 10.3pt;
  }
  thead th{
    text-align:left;
    border: 1px solid #111;
    padding: 2.4mm 2.2mm;
    font-weight: 950;
    background: #f4f6fb;
  }
  tbody td{
    border: 1px solid #111;
    padding: 2.4mm 2.2mm;
    vertical-align: top;
  }
  .col-day{ width: 17%; font-weight: 950; }
  .col-sort{ width: 52%; }
  .col-tag{ width: 16%; white-space: nowrap; }
  .col-zeit{ width: 15%; white-space: nowrap; }

  .ds-block{
    margin-top: 7mm;
    border: 1px solid rgba(0,0,0,0.20);
    border-radius: 10px;
    overflow: hidden;
  }
  .ds-head{
    padding: 8px 10px;
    background: #f4f6fb;
    font-weight: 950;
    font-size: 10.5pt;
    border-bottom: 1px solid rgba(0,0,0,0.20);
  }
  .ds-body{
    padding: 8px 10px;
    font-size: 10.2pt;
    color: var(--ink);
    line-height: 1.45;
  }

  .empty{
    padding: 24px;
    text-align:center;
    color: var(--muted);
    font-size: 13px;
    line-height: 1.4;
  }

  .toast{
    margin-top: 8px;
    color: var(--muted);
    font-size: 12px;
  }

  @media (max-width: 980px){
    .app{ grid-template-columns: 1fr; }
    .paper{ width: 100%; min-height: auto; }
  }

  /* PRINT */
  @media print{
    body{ background: #fff !important; }
    .sidebar, .main-head{ display:none !important; }
    .app{ display:block; padding:0; }
    .main{ border:none; background:transparent; }
    .paper-wrap{ padding:0; }
    .paper{
      width:auto;
      min-height:auto;
      border-radius: 0;
      box-shadow:none;
      border:none;
      page-break-after: always;
    }
  }
</style>
</head>

<body>
<div class="app">
  <div class="sidebar">
    <div class="side-head">
      <div class="title">Sende- & Belieferungsplan <span class="badge">A4</span></div>
      <div class="subtitle">
        Kundennummer eingeben oder aus der Liste wählen.<br>
        Für Massendruck: <b>Alle rendern</b> → <b>Drucken</b>.
      </div>
    </div>

    <div class="search">
      <div class="field">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
          <path d="M21 21l-4.3-4.3m1.3-5.2a7.2 7.2 0 11-14.4 0 7.2 7.2 0 0114.4 0z" stroke="rgba(255,255,255,.7)" stroke-width="2" stroke-linecap="round"/>
        </svg>
        <input id="knr" placeholder="Kundennummer (z.B. 88130)" inputmode="numeric" />
      </div>

      <div class="btnrow">
        <button class="btn primary" onclick="showOne()">Anzeigen</button>
        <button class="btn" onclick="showAll()">Alle rendern</button>
        <button class="btn" onclick="window.print()">Drucken</button>
        <button class="btn danger" onclick="clearView()">Leeren</button>
      </div>
      <div class="toast" id="hint"></div>
    </div>

    <div class="list" id="list"></div>
  </div>

  <div class="main">
    <div class="main-head">
      <div class="meta">
        <div class="big" id="mainTitle">Vorschau</div>
        <div class="small" id="mainSub">Noch kein Kunde ausgewählt.</div>
      </div>
      <div class="meta" style="text-align:right;">
        <div class="big" id="count"></div>
        <div class="small">Kunden im Dokument</div>
      </div>
    </div>

    <div class="paper-wrap">
      <div id="out" class="empty">
        Suche links nach einer Kundennummer oder wähle einen Eintrag aus der Liste.
      </div>
    </div>
  </div>
</div>

<script>
const DATA = __DATA_JSON__;
const ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0)-(Number(b)||0));
let activeKey = null;

function esc(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;");
}

function daysAndTours(c){
  const days = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];
  const active = days.filter(d => (c.tours && String(c.tours[d]||"").trim() !== ""));
  return {
    dayLine: active.length ? active.join(" ") : "-",
    tourLine: active.length ? active.map(d => esc(c.tours[d])).join(" ") : "-"
  };
}

function buildRows(c){
  const days = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];
  const byDay = {};
  for(const it of (c.bestell||[])){
    if(!byDay[it.liefertag]) byDay[it.liefertag] = [];
    byDay[it.liefertag].push(it);
  }

  let rows = "";
  for(const d of days){
    const arr = byDay[d] || [];
    const sortLines = arr.map(x => esc(x.sortiment)).join("<br>");
    const tagLines  = arr.map(x => esc(x.bestelltag)).join("<br>");
    const zeitLines = arr.map(x => esc(x.bestellschluss)).join("<br>");

    rows += `
      <tr>
        <td class="col-day">${esc(d)}</td>
        <td class="col-sort">${sortLines || "&nbsp;"}</td>
        <td class="col-tag">${tagLines || "&nbsp;"}</td>
        <td class="col-zeit">${zeitLines || "&nbsp;"}</td>
      </tr>
    `;
  }
  return rows;
}

function renderPaper(c){
  const {dayLine, tourLine} = daysAndTours(c);

  let dsHtml = "";
  if (c.ds && c.ds.length){
    const lines = c.ds.map(it => `
      <div><b>${esc(it.ds_key)}:</b> ${esc(it.sortiment)} — <b>${esc(it.bestelltag)}</b> · ${esc(it.bestellschluss)}</div>
    `).join("");
    dsHtml = `
      <div class="ds-block">
        <div class="ds-head">Durchsteck (DS)</div>
        <div class="ds-body">${lines}</div>
      </div>
    `;
  }

  return `
  <div class="paper">
    <div class="paper-inner">
      <div class="p-title">Sende- &amp; Belieferungsplan</div>
      <div class="p-standard">${esc(c.plan_typ || "Standard")}</div>
      <div class="p-sub">${esc(c.name)} ${esc(c.bereich || "Alle Sortimente Fleischwerk")}</div>

      <div class="p-head">
        <div class="p-addr">
          <div class="name">${esc(c.name)}</div>
          <div>${esc(c.strasse)}</div>
          <div>${esc(c.plz)} ${esc(c.ort)}</div>
        </div>
        <div class="p-meta">
          <div><b>Kunden-Nr.:</b> ${esc(c.kunden_nr)}</div>
          <div><b>Fachberater:</b> ${esc(c.fachberater || "")}</div>
        </div>
      </div>

      <div class="p-lines">
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
        <tbody>
          ${buildRows(c)}
        </tbody>
      </table>

      ${dsHtml}
    </div>
  </div>`;
}

function setMainHeader(c){
  document.getElementById("mainTitle").textContent = `Vorschau: ${c.kunden_nr}`;
  document.getElementById("mainSub").textContent = `${c.name} · ${c.plz} ${c.ort}`;
}

function setHint(msg){
  document.getElementById("hint").textContent = msg || "";
}

function setActive(key){
  activeKey = key;
  const list = document.getElementById("list").querySelectorAll(".item");
  list.forEach(el => el.classList.toggle("active", el.dataset.key === key));
}

function showOne(){
  const knr = document.getElementById("knr").value.trim();
  const out = document.getElementById("out");
  if(!knr){
    out.className = "empty";
    out.innerHTML = "Bitte eine Kundennummer eingeben.";
    setHint("");
    return;
  }
  const c = DATA[knr];
  if(!c){
    out.className = "empty";
    out.innerHTML = `Kundennummer <b>${esc(knr)}</b> nicht gefunden.`;
    setHint(`Vorhanden: ${ORDER.length} Kunden`);
    return;
  }
  out.className = "";
  out.innerHTML = renderPaper(c);
  setMainHeader(c);
  setHint("Anzeigen bereit. Druck: Button „Drucken“.");
  setActive(knr);
}

function showAll(){
  const out = document.getElementById("out");
  const html = ORDER.map(k => renderPaper(DATA[k])).join("");
  out.className = "";
  out.innerHTML = html;
  document.getElementById("mainTitle").textContent = "Massendruck (alle Kunden)";
  document.getElementById("mainSub").textContent = `${ORDER.length} Seiten gerendert`;
  setHint("Alle Seiten gerendert. Jetzt „Drucken“.");
  setActive(null);
}

function clearView(){
  const out = document.getElementById("out");
  out.className = "empty";
  out.innerHTML = "Suche links nach einer Kundennummer oder wähle einen Eintrag aus der Liste.";
  document.getElementById("mainTitle").textContent = "Vorschau";
  document.getElementById("mainSub").textContent = "Noch kein Kunde ausgewählt.";
  setHint("");
  setActive(null);
}

function buildList(){
  const list = document.getElementById("list");
  list.innerHTML = ORDER.map(k => {
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

document.getElementById("knr").addEventListener("keydown", (e)=>{
  if(e.key === "Enter") showOne();
});

document.getElementById("count").textContent = ORDER.length;
buildList();
</script>
</body>
</html>
"""


# -----------------------------
# Streamlit UI (Generator)
# -----------------------------
st.set_page_config(page_title="Excel → Moderne HTML Druckvorlage", layout="wide")
st.title("Excel → Moderne HTML (Kundennummer eingeben → A4 drucken)")

up = st.file_uploader("Excel (.xlsx) hochladen", type=["xlsx"])
if not up:
    st.info("Lade die Excel hoch. Danach erzeugt die App eine Standalone-HTML (Suche + A4-Druck).")
    st.stop()

df = pd.read_excel(up, engine="openpyxl")
df.columns = [c.strip() for c in df.columns]

required = ["Nr", "SAP-Nr.", "Name", "Strasse", "Plz", "Ort"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Pflichtspalten fehlen: {missing}")
    st.stop()

# Detektoren
trip21 = detect_triplets_21(df.columns.tolist())
bmap = detect_bspalten(df.columns.tolist())
dsmap = detect_ds_triplets(df.columns.tolist())

with st.expander("Debug – erkannte Quellen"):
    st.write("**B_-Spalten Keys (Anzahl):**", len(bmap))
    st.write("**21-Tripel Tage:**", ", ".join([d for d in DAYS_DE if d in trip21]) or "-")
    st.write("**DS-Tripel Keys:**", ", ".join(dsmap.keys()) or "-")

data: Dict[str, dict] = {}

for _, r in df.iterrows():
    knr = norm(r.get("Nr", ""))
    if not knr:
        continue

    # Tours
    tours = {}
    for day_de, col in TOUR_COLS.items():
        tours[day_de] = norm(r.get(col, "")) if col in df.columns else ""

    # Bestellzeilen
    bestell = []

    # 1) B_-Spalten (Hauptdaten)
    for day_de in DAYS_DE:
        keys = [k for k in bmap.keys() if k[0] == day_de]
        # sort: group numeric then order of bestelltag
        keys.sort(key=lambda k: (group_sort_key(k[1]), DAYS_DE.index(k[2]) if k[2] in DAYS_DE else 99))

        for (lday, group, bestelltag) in keys:
            cols = bmap[(lday, group, bestelltag)]
            sortiment = norm(r.get(cols.get("sort", ""), ""))
            zeit = norm(r.get(cols.get("zeit", ""), ""))

            if not sortiment and not zeit:
                continue

            # Zeit normalisieren
            if zeit and "uhr" not in zeit.lower():
                if re.fullmatch(r"\d{1,2}:\d{2}", zeit):
                    zeit = zeit + " Uhr"

            bestell.append({
                "liefertag": lday,
                "sortiment": sortiment,
                "bestelltag": bestelltag,
                "bestellschluss": zeit
            })

    # 2) 21-Tripel (falls zusätzlich gepflegt)
    for day_de in DAYS_DE:
        for g in sorted(trip21.get(day_de, {}).keys(), key=group_sort_key):
            cols = trip21[day_de][g]
            zeit = norm(r.get(cols["Zeit"], ""))
            sortiment = norm(r.get(cols["Sort"], ""))
            bestelltag = norm(r.get(cols["Tag"], ""))

            if sortiment:
                if zeit and "uhr" not in zeit.lower() and re.fullmatch(r"\d{1,2}:\d{2}", zeit):
                    zeit = zeit + " Uhr"
                bestell.append({
                    "liefertag": day_de,
                    "sortiment": sortiment,
                    "bestelltag": bestelltag,
                    "bestellschluss": zeit
                })

    # 3) DS (optional, eigener Block)
    ds_list = []
    for ds_key, cols in dsmap.items():
        zeit = norm(r.get(cols["Zeit"], ""))
        sortiment = norm(r.get(cols["Sort"], ""))
        bestelltag = norm(r.get(cols["Tag"], ""))
        if sortiment or zeit or bestelltag:
            if zeit and "uhr" not in zeit.lower() and re.fullmatch(r"\d{1,2}:\d{2}", zeit):
                zeit = zeit + " Uhr"
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

st.success(f"{len(data)} Kunden eingebettet. HTML kann jetzt erzeugt werden.")

html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))

st.download_button(
    "⬇️ Standalone-HTML herunterladen (modern + A4 Druck)",
    data=html.encode("utf-8"),
    file_name="sende_belieferungsplan_modern.html",
    mime="text/html",
)

st.caption("Tipp: Im Browser beim Drucken A4 + Skalierung 100%. Für Massendruck: in der HTML „Alle rendern“ → Drucken.")
