# app.py
# Excel -> 1 HTML-Datei mit Suchfeld (Kundennummer) -> A4 Drucken
# Layout ist absichtlich "PDF-nah": minimal, klar, tabellarisch.

import json
import re
from typing import Dict  # <-- FIX
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
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s

def detect_triplets(columns):
    """
    erkennt alle Tripel:
      "<Tag> <Gruppe> Zeit"  / "<Tag> <Gruppe> Sort" / "<Tag> <Gruppe> Tag"
    """
    rx = re.compile(r"^(Mo|Die|Di|Mi|Do|Don|Mitt|Fr|Sa|Sam)\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)
    found = {}

    for c in [c.strip() for c in columns]:
        m = rx.match(c)
        if not m:
            continue

        day_short = m.group(1)
        group = m.group(2).strip()
        field = m.group(3).capitalize()

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

def group_sort_key(g: str):
    g = g.strip()
    if g.isdigit():
        return (0, int(g))
    return (1, g.lower())

HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Sende- & Belieferungsplan</title>
<style>
  body { margin:0; font-family: Arial, Helvetica, sans-serif; background:#f0f0f0; }

  .bar {
    position: sticky; top: 0; z-index: 20;
    background:#fff; border-bottom:1px solid #bbb;
    padding: 10px 12px; display:flex; gap:10px; align-items:center; flex-wrap:wrap;
  }
  .bar input{
    padding:10px 12px; font-size:16px; border:1px solid #bbb; border-radius:8px; width: 220px;
  }
  .btn{
    padding:10px 12px; font-size:14px; border:1px solid #bbb; border-radius:8px; background:#fff; cursor:pointer;
  }
  .btn.primary { border-color:#1b66b3; color:#1b66b3; font-weight:700; }
  .muted{ color:#666; }

  .page{
    width: 210mm;
    min-height: 297mm;
    background:#fff;
    margin: 10mm auto;
    padding: 12mm;
    box-sizing:border-box;
    border:1px solid #bbb;
  }

  .h1{ text-align:center; font-size:20pt; font-weight:800; margin:0; }
  .std{ text-align:center; font-size:18pt; font-weight:900; color:#c00; margin:2mm 0 2mm 0; }
  .sub{ text-align:center; font-size:11pt; margin:0 0 6mm 0; color:#111; }

  .head{
    display:flex; justify-content:space-between; gap:12mm;
    margin-bottom: 5mm;
  }
  .addr{ font-size:11pt; line-height:1.25; }
  .meta{ font-size:11pt; line-height:1.35; min-width:70mm; }

  .lines{ margin: 4mm 0 5mm 0; font-size:11pt; line-height:1.35; }
  .lines b{ font-weight:800; }

  table{
    width:100%;
    border-collapse:collapse;
    font-size:10.5pt;
  }
  thead th{
    text-align:left;
    font-weight:800;
    border:1px solid #000;
    padding: 2.5mm 2mm;
  }
  tbody td{
    border:1px solid #000;
    padding: 2.5mm 2mm;
    vertical-align:top;
  }
  .col-day{ width: 17%; font-weight:800; }
  .col-sort{ width: 52%; }
  .col-tag{ width: 16%; white-space:nowrap; }
  .col-zeit{ width: 15%; white-space:nowrap; }

  .err{
    width:210mm; margin:10mm auto; background:#fff;
    border:1px solid #d66; padding:12px; border-radius:8px; color:#900;
  }

  @media print{
    body{ background:#fff; }
    .bar{ display:none; }
    .page{ margin:0; border:none; page-break-after:always; }
  }
</style>
</head>
<body>

<div class="bar">
  <b>Kundennummer:</b>
  <input id="knr" placeholder="z.B. 88130" inputmode="numeric" />
  <button class="btn" onclick="showOne()">Anzeigen</button>
  <button class="btn primary" onclick="window.print()">Drucken</button>
  <button class="btn" onclick="showAll()">Alle drucken</button>
  <span class="muted" id="hint"></span>
</div>

<div id="out"></div>

<script>
const DATA = __DATA_JSON__;

function esc(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;");
}

function buildHeader(c){
  const days = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];
  const active = days.filter(d => (c.tours && String(c.tours[d]||"").trim() !== ""));
  const dayLine = active.length ? active.join(" ") : "-";
  const tourLine = active.length ? active.map(d => esc(c.tours[d])).join(" ") : "-";
  return {dayLine, tourLine};
}

function buildTableRows(c){
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

function renderPage(c){
  const {dayLine, tourLine} = buildHeader(c);

  return `
  <div class="page">
    <div class="h1">Sende- &amp; Belieferungsplan</div>
    <div class="std">${esc(c.plan_typ || "Standard")}</div>
    <div class="sub">${esc(c.name)} ${esc(c.bereich || "Alle Sortimente Fleischwerk")}</div>

    <div class="head">
      <div class="addr">
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
      <tbody>
        ${buildTableRows(c)}
      </tbody>
    </table>
  </div>`;
}

function showOne(){
  const knr = document.getElementById("knr").value.trim();
  const out = document.getElementById("out");
  const hint = document.getElementById("hint");

  if(!knr){
    out.innerHTML = `<div class="err">Bitte Kundennummer eingeben.</div>`;
    hint.textContent = "";
    return;
  }
  const c = DATA[knr];
  if(!c){
    out.innerHTML = `<div class="err">Kundennummer <b>${esc(knr)}</b> nicht gefunden.</div>`;
    hint.textContent = `Vorhanden: ${Object.keys(DATA).length} Kunden`;
    return;
  }
  out.innerHTML = renderPage(c);
  hint.textContent = `${c.name} · ${c.plz} ${c.ort}`;
}

function showAll(){
  const out = document.getElementById("out");
  const hint = document.getElementById("hint");
  const keys = Object.keys(DATA).sort((a,b)=>(Number(a)||0)-(Number(b)||0));
  out.innerHTML = keys.map(k => renderPage(DATA[k])).join("");
  hint.textContent = `Alle Kunden: ${keys.length}`;
}

document.getElementById("knr").addEventListener("keydown",(e)=>{
  if(e.key==="Enter") showOne();
});
</script>
</body>
</html>
"""

st.set_page_config(page_title="Excel → HTML (Vorlage Standard)", layout="wide")
st.title("Excel → HTML (Vorlage „Standard“ wie im PDF)")

up = st.file_uploader("Excel (.xlsx) hochladen", type=["xlsx"])
if not up:
    st.info("Excel hochladen → HTML erzeugen → später Kundennummer eingeben und drucken.")
    st.stop()

df = pd.read_excel(up, engine="openpyxl")
df.columns = [c.strip() for c in df.columns]

required = ["Nr", "SAP-Nr.", "Name", "Strasse", "Plz", "Ort"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Pflichtspalten fehlen: {missing}")
    st.stop()

triplets = detect_triplets(df.columns.tolist())
if not triplets:
    st.warning("Keine Spalten nach Muster '<Tag> <Gruppe> Zeit/Sort/Tag' erkannt. Prüfe Header.")
else:
    days_found = [d for d in DAYS_DE if d in triplets]
    st.success(f"Zeit/Sort/Tag-Tripel erkannt für: {', '.join(days_found)}")

data: Dict[str, dict] = {}

for _, r in df.iterrows():
    knr = norm(r.get("Nr", ""))
    if not knr:
        continue

    tours = {}
    for day_de, col in TOUR_COLS.items():
        tours[day_de] = norm(r.get(col, "")) if col in df.columns else ""

    bestell = []
    for day_de in DAYS_DE:
        groups = sorted(triplets.get(day_de, {}).keys(), key=group_sort_key)
        for g in groups:
            cols = triplets[day_de][g]
            zeit = norm(r.get(cols["Zeit"], ""))
            sortiment = norm(r.get(cols["Sort"], ""))
            bestelltag = norm(r.get(cols["Tag"], ""))

            if sortiment:
                bestell.append({
                    "liefertag": day_de,
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
        "bestell": bestell
    }

st.write(f"**{len(data)}** Kunden eingebettet.")

html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))

st.download_button(
    "⬇️ HTML erzeugen (Kundennummer eingeben → Drucken)",
    data=html.encode("utf-8"),
    file_name="sende_belieferungsplan_standard.html",
    mime="text/html"
)

with st.expander("Debug: erkannte Gruppen je Tag"):
    for d in DAYS_DE:
        groups = sorted(triplets.get(d, {}).keys(), key=group_sort_key)
        st.write(f"**{d}:** {', '.join(groups) if groups else '-'}")
