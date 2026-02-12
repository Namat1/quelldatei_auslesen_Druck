# app.py
# ------------------------------------------------------------
# Excel -> Standalone-HTML (Deutsche See in Wochentage integriert)
# ------------------------------------------------------------

import json
import re
import datetime
from typing import Dict, Tuple, List
import pandas as pd
import streamlit as st

PLAN_TYP = "Standard"
BEREICH = "Alle Sortimente Fleischwerk"
DAYS_DE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]

DAY_SHORT_TO_DE = {
    "Mo": "Montag", "Di": "Dienstag", "Die": "Dienstag",
    "Mi": "Mittwoch", "Mitt": "Mittwoch", "Do": "Donnerstag",
    "Don": "Donnerstag", "Donn": "Donnerstag", "Fr": "Freitag",
    "Sa": "Samstag", "Sam": "Samstag",
}

TOUR_COLS = {
    "Montag": "Mo", "Dienstag": "Die", "Mittwoch": "Mitt",
    "Donnerstag": "Don", "Freitag": "Fr", "Samstag": "Sam",
}

def norm(x) -> str:
    if x is None: return ""
    if isinstance(x, float) and pd.isna(x): return ""
    s = str(x).replace("\u00a0", " ").strip()
    s = re.sub(r"\s+", " ", s)
    if re.fullmatch(r"\d+\.0", s): s = s[:-2]
    return s

def normalize_time(s) -> str:
    if isinstance(s, (datetime.time, pd.Timestamp)):
        return s.strftime("%H:%M") + " Uhr"
    s = norm(s)
    if not s: return ""
    if "uhr" not in s.lower() and re.fullmatch(r"\d{1,2}:\d{2}", s):
        return s + " Uhr"
    return s

def detect_triplets(columns: List[str]) -> Dict[str, Dict[str, Dict[str, str]]]:
    rx = re.compile(r"^(Mo|Die|Di|Mitt|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)
    found = {}
    for c in [c.strip() for c in columns]:
        m = rx.match(c)
        if not m: continue
        day_short, group, field = m.group(1), m.group(2).strip(), m.group(3).capitalize()
        if day_short.lower() == "donn": day_short = "Don"
        day_de = DAY_SHORT_TO_DE.get(day_short)
        if day_de:
            found.setdefault(day_de, {}).setdefault(group, {})[field] = c
    return found

def detect_ds_triplets(columns: List[str]) -> Dict[str, Dict[str, str]]:
    """Erkennt DS Spalten wie 'DS Montag Zeit', 'DS Montag Sort' etc."""
    rx = re.compile(r"^DS\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)
    tmp = {}
    for c in [c.strip() for c in columns]:
        m = rx.match(c)
        if not m: continue
        day_raw, field = m.group(1).strip(), m.group(2).capitalize()
        # Mapping von DS Montag -> Montag
        day_de = DAY_SHORT_TO_DE.get(day_raw) or day_raw
        if day_de in DAYS_DE:
            tmp.setdefault(day_de, {})[field] = c
    return tmp

HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<style>
  @page { size: A4; margin: 0; }
  :root{
    --bg:#0b1220; --panel:rgba(255,255,255,.08); --stroke:rgba(255,255,255,.14);
    --text:rgba(255,255,255,.92); --muted:rgba(255,255,255,.62);
    --paper:#fff; --ink:#0b0f17; --sub:#394054;
  }
  *{ box-sizing:border-box; font-family: sans-serif; }
  body{ margin:0; background: var(--bg); color:var(--text); }
  .app{ display:grid; grid-template-columns: 340px 1fr; height:100vh; padding:15px; gap:15px; }
  .sidebar, .main{ background: var(--panel); border:1px solid var(--stroke); border-radius:12px; overflow:hidden; }
  .list{ height: calc(100vh - 250px); overflow-y:auto; border-top:1px solid var(--stroke); }
  .item{ padding:10px; border-bottom:1px solid rgba(255,255,255,0.1); cursor:pointer; font-size:13px; }
  .item:hover{ background:rgba(255,255,255,0.05); }
  .main{ display:flex; flex-direction:column; }
  .wrap{ flex:1; overflow-y:auto; padding:20px; display:flex; flex-direction:column; align-items:center; }
  
  .paper{
    width:210mm; height:297mm; background:white; color:black; padding:15mm;
    position:relative; box-shadow: 0 0 20px rgba(0,0,0,0.5); page-break-after: always;
    --fs: 10.5pt;
  }
  .paper * { font-size: var(--fs); }
  .ptitle{ text-align:center; font-weight:900; font-size:1.5em; margin:0; }
  .pstd{ text-align:center; color:#d0192b; font-weight:bold; margin:2mm 0; }
  .psub{ text-align:center; color:#666; margin-bottom:10mm; }
  .head{ display:flex; justify-content:space-between; margin-bottom:8mm; }
  table{ width:100%; border-collapse:collapse; }
  th, td{ border:1px solid #333; padding:2mm; text-align:left; }
  th{ background:#f0f0f0; }
  .ds-label { color: #0056b3; font-weight: bold; font-style: italic; }

  @media print{
    .sidebar{ display:none; }
    .app{ display:block; padding:0; }
    .paper{ box-shadow:none; margin:0; }
  }
</style>
</head>
<body>
<div class="app">
  <div class="sidebar">
    <div style="padding:15px"><b>Plan-Manager</b></div>
    <div style="padding:15px; display:flex; flex-direction:column; gap:10px;">
      <input id="knr" placeholder="Kundennummer..." style="padding:8px; border-radius:5px; border:none;">
      <button onclick="showOne()" style="padding:8px; cursor:pointer;">Anzeigen</button>
      <button onclick="showAll()" style="padding:8px; cursor:pointer;">Alle Render</button>
      <button onclick="window.print()" style="padding:8px; background:#28a745; color:white; border:none; border-radius:5px;">Drucken</button>
    </div>
    <div class="list" id="list"></div>
  </div>
  <div class="main"><div class="wrap" id="out"></div></div>
</div>

<script>
const DATA = __DATA_JSON__;
const ORDER = Object.keys(DATA).sort((a,b)=>a-b);
const DAYS = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];

function esc(s){ return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;"); }

function render(c){
  const byDay = {};
  c.bestell.forEach(it => { if(!byDay[it.liefertag]) byDay[it.liefertag]=[]; byDay[it.liefertag].push(it); });

  const rows = DAYS.map(d => {
    const items = byDay[d] || [];
    return `<tr>
      <td style="width:15%"><b>${d}</b></td>
      <td>${items.map(x => (x.is_ds ? '<span class="ds-label">DS: </span>' : '') + esc(x.sortiment)).join("<br>") || "-"}</td>
      <td style="width:15%">${items.map(x => esc(x.bestelltag)).join("<br>") || "-"}</td>
      <td style="width:15%">${items.map(x => esc(x.bestellschluss)).join("<br>") || "-"}</td>
    </tr>`;
  }).join("");

  return `
    <div class="paper">
      <div class="ptitle">Sende- &amp; Belieferungsplan</div>
      <div class="pstd">${esc(c.plan_typ)}</div>
      <div class="psub">${esc(c.name)} | ${esc(c.bereich)}</div>
      <div class="head">
        <div><b>${esc(c.name)}</b><br>${esc(c.strasse)}<br>${esc(c.plz)} ${esc(c.ort)}</div>
        <div style="text-align:right">Kdn-Nr: ${esc(c.kunden_nr)}<br>Tour: ${DAYS.map(d=>(c.tours[d]||"")).filter(x=>x).join("/")}</div>
      </div>
      <table>
        <thead><tr><th>Liefertag</th><th>Sortiment</th><th>Bestelltag</th><th>Schluss</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

function autoFit(){
  document.querySelectorAll(".paper").forEach(p => {
    let fs = 10.5; p.style.setProperty("--fs", fs + "pt");
    while(p.scrollHeight > p.clientHeight && fs > 7){
      fs -= 0.1; p.style.setProperty("--fs", fs + "pt");
    }
  });
}

function showOne(){
  const k = document.getElementById("knr").value;
  if(DATA[k]) { document.getElementById("out").innerHTML = render(DATA[k]); autoFit(); }
}

function showAll(){
  document.getElementById("out").innerHTML = ORDER.map(k=>render(DATA[k])).join("");
  autoFit();
}

document.getElementById("list").innerHTML = ORDER.map(k=>`<div class="item" onclick="document.getElementById('knr').value='${k}';showOne()">${k} - ${DATA[k].name}</div>`).join("");
</script>
</body>
</html>
"""

st.set_page_config(page_title="Plan-Generator", layout="wide")
st.title("Sendeplan Generator (Deutsche See integriert)")

up = st.file_uploader("Excel Datei", type=["xlsx"])
if up:
    df = pd.read_excel(up)
    trip = detect_triplets(df.columns)
    ds_trip = detect_ds_triplets(df.columns)
    
    data = {}
    for _, r in df.iterrows():
        knr = norm(r.get("Nr", ""))
        if not knr: continue
        
        bestell = []
        # 1. Normale Sortimente
        for d_de, groups in trip.items():
            for g, cols in groups.items():
                s, t, z = norm(r.get(cols["Sort"])), norm(r.get(cols["Tag"])), normalize_time(r.get(cols["Zeit"]))
                if s or t or z:
                    bestell.append({"liefertag": d_de, "sortiment": s, "bestelltag": t, "bestellschluss": z, "is_ds": False})
        
        # 2. Deutsche See (DS) integrieren
        for d_de, cols in ds_trip.items():
            s, t, z = norm(r.get(cols["Sort"])), norm(r.get(cols["Tag"])), normalize_time(r.get(cols["Zeit"]))
            if s or t or z:
                bestell.append({"liefertag": d_de, "sortiment": s, "bestelltag": t, "bestellschluss": z, "is_ds": True})

        data[knr] = {
            "plan_typ": PLAN_TYP, "bereich": BEREICH, "kunden_nr": knr,
            "name": norm(r.get("Name", "")), "strasse": norm(r.get("Strasse", "")),
            "plz": norm(r.get("Plz", "")), "ort": norm(r.get("Ort", "")),
            "tours": {d: norm(r.get(c, "")) for d, c in TOUR_COLS.items() if c in df.columns},
            "bestell": bestell
        }

    html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(data, separators=(',', ':')))
    st.download_button("HTML Speichern", data=html, file_name="sendeplan.html", mime="text/html")
