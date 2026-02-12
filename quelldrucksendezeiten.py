# quelldrucksendezeiten.py
# -----------------------------------------------------------------------------
# VERSION: A4-MAX-SPACE & DATEN-CHECK
# -----------------------------------------------------------------------------

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
    "Mi": "Mittwoch", "Mit": "Mittwoch", "Mitt": "Mittwoch",
    "Do": "Donnerstag", "Don": "Donnerstag", "Donn": "Donnerstag",
    "Fr": "Freitag", "Sa": "Samstag", "Sam": "Samstag",
}

TOUR_COLS = {"Montag": "Mo", "Dienstag": "Die", "Mittwoch": "Mitt", "Donnerstag": "Don", "Freitag": "Fr", "Samstag": "Sam"}

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
    if re.fullmatch(r"\d{1,2}:\d{2}", s): return s + " Uhr"
    if re.fullmatch(r"\d{1,2}", s): return s.zfill(2) + ":00 Uhr"
    return s

def group_sort_key(g: str):
    g = str(g).strip()
    if g.isdigit(): return (0, int(g))
    return (1, g.lower())

def detect_bspalten(columns: List[str]):
    rx = re.compile(r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(?:(Z|L)\s+)?(.+?)\s+B[_ ]?(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)$", re.IGNORECASE)
    mapping = {}
    for c in columns:
        m = rx.match(c.strip())
        if m:
            day_de = DAY_SHORT_TO_DE.get(m.group(1))
            zl = (m.group(2) or "").upper()
            group_id = m.group(3).strip()
            bestell_de = DAY_SHORT_TO_DE.get(m.group(4))
            if day_de and bestell_de:
                key = (day_de, group_id, bestell_de)
                mapping.setdefault(key, {})
                if zl == "Z": mapping[key]["zeit"] = c
                elif zl == "L": mapping[key]["l"] = c
                else: mapping[key]["sort"] = c
    return mapping

def detect_triplets(columns: List[str]):
    rx = re.compile(r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)
    found = {}
    for c in columns:
        m = rx.match(c.strip())
        if m:
            day_de = DAY_SHORT_TO_DE.get(m.group(1))
            if day_de:
                found.setdefault(day_de, {}).setdefault(m.group(2).strip(), {})[m.group(3).capitalize()] = c
    return found

def detect_ds_triplets(columns: List[str]):
    rx = re.compile(r"^DS\s+(.+?)\s+zu\s+(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)
    tmp = {}
    for c in columns:
        m = rx.match(c.strip())
        if m:
            day_de = DAY_SHORT_TO_DE.get(m.group(2))
            if day_de:
                key = f"DS {m.group(1)} zu {m.group(2)}"
                tmp.setdefault(day_de, {}).setdefault(key, {})[m.group(3).capitalize()] = c
    return tmp

# --- HTML TEMPLATE FÜR EXTREMEN PLATZBEDARF ---
HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<style>
  @page { size: A4; margin: 0; }
  *{ box-sizing:border-box; font-family: Arial, sans-serif; }
  body{ margin:0; background: #0b1220; color: #fff; }
  .app{ display:grid; grid-template-columns: 350px 1fr; height:100vh; padding:15px; gap:15px; }
  .sidebar, .main{ background: rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.14); border-radius:12px; }
  .list{ height: calc(100vh - 280px); overflow-y:auto; border-top:1px solid rgba(255,255,255,.14); }
  .item{ padding:8px; border-bottom:1px solid rgba(255,255,255,0.05); cursor:pointer; font-size:12px; }
  .wrap{ height: 100%; overflow-y: scroll; padding: 20px; display: flex; flex-direction: column; align-items: center; }

  .paper{
    width: 210mm; height: 296.5mm; background: white; color: black; padding: 10mm;
    box-shadow: 0 0 20px rgba(0,0,0,0.5); display: flex; flex-direction: column;
    --fs: 11pt; 
  }
  .paper * { font-size: var(--fs); line-height: 1.2; }
  .ptitle{ text-align:center; font-weight:900; font-size:1.5em; margin:0; }
  .pstd{ text-align:center; color:#d0192b; font-weight:bold; margin:1mm 0; font-size:1.2em; }
  .psub{ text-align:center; color:#444; margin-bottom:3mm; font-weight:bold; font-size:0.9em; }
  
  .head-box { display:flex; justify-content:space-between; margin-bottom:3mm; border-bottom:1.5px solid #000; padding-bottom:2mm; }
  .head-left { width: 60%; }
  .head-right { width: 35%; text-align: right; }

  .tour-table { width: 100%; border-collapse: collapse; margin-bottom: 3mm; }
  .tour-table th { background: #eee; font-size: 0.7em; padding: 1px; border: 1px solid #000; text-transform: uppercase; }
  .tour-table td { border: 1px solid #000; padding: 3px; text-align: center; font-weight: bold; font-size: 0.9em; }

  table.main-table { width:100%; border-collapse:collapse; table-layout: fixed; border: 1.5px solid #000; }
  table.main-table th { border: 1px solid #000; padding: 4px; background:#f2f2f2; font-weight:bold; text-align:left; font-size: 0.85em; }
  table.main-table td { border: 1px solid #000; padding: 4px; vertical-align: top; font-size: 0.95em; }

  .day-header { background-color: #f0f0f0 !important; font-weight: bold; border-top: 2px solid #000 !important; }

  @media print{
    body { background: white; }
    .sidebar { display:none !important; }
    .app { display:block; padding:0; }
    .wrap { overflow: visible; padding: 0; }
    .paper { box-shadow: none; margin: 0; page-break-after: always; }
  }
</style>
</head>
<body>
<div class="app">
  <div class="sidebar">
    <div style="padding:15px; font-weight:bold;">Sendeplan Generator</div>
    <div style="padding:15px; display:flex; flex-direction:column; gap:10px;">
      <input id="knr" placeholder="Kunden-Nr..." style="width:100%; padding:10px; border-radius:5px;">
      <button onclick="showOne()" style="padding:10px; background:#4fa3ff; color:white; border:none; cursor:pointer;">Anzeigen</button>
      <button onclick="window.print()" style="padding:10px; background:#28a745; color:white; border:none; cursor:pointer;">Drucken</button>
    </div>
    <div class="list" id="list"></div>
  </div>
  <div class="main"><div class="wrap" id="out">Kunden wählen...</div></div>
</div>
<script>
const DATA = __DATA_JSON__;
const ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0)-(Number(b)||0));
const DAYS = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];

function esc(s){ return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;"); }

function render(c){
  let tableRows = "";
  DAYS.forEach(d => {
    const items = (c.bestell || []).filter(it => it.liefertag === d);
    if (items.length > 0) {
      items.forEach((it, idx) => {
        tableRows += `<tr class="${idx === 0 ? 'day-header' : ''}">
          <td style="width:18%">${idx === 0 ? d : ""}</td>
          <td style="width:50%">${esc(it.sortiment)}</td>
          <td style="width:16%">${esc(it.bestelltag)}</td>
          <td style="width:16%">${esc(it.bestellschluss)}</td>
        </tr>`;
      });
    }
  });

  const tourRow = DAYS.map(d => `<td>${esc(c.tours[d] || "—")}</td>`).join("");
  const tourHead = DAYS.map(d => `<th>${d.substring(0,2)}</th>`).join("");

  return `<div class="paper">
    <div class="ptitle">Sende- &amp; Belieferungsplan</div>
    <div class="pstd">${esc(c.plan_typ)}</div>
    <div class="psub">${esc(c.name)} | ${esc(c.bereich)}</div>
    <div class="head-box">
      <div class="head-left"><b>${esc(c.name)}</b><br>${esc(c.strasse)}<br>${esc(c.plz)} ${esc(c.ort)}</div>
      <div class="head-right">Kunden-Nr: <b>${esc(c.kunden_nr)}</b><br>Fachberater: <b>${esc(c.fachberater)}</b></div>
    </div>
    <table class="tour-table"><thead><tr>${tourHead}</tr></thead><tbody><tr>${tourRow}</tr></tbody></table>
    <table class="main-table">
      <thead><tr><th>Liefertag</th><th>Sortiment</th><th>Bestelltag</th><th>Schluss</th></tr></thead>
      <tbody>${tableRows}</tbody>
    </table>
  </div>`;
}

function autoFit(){
  document.querySelectorAll(".paper").forEach(p => {
    let fs = 11; p.style.setProperty("--fs", fs + "pt");
    while(p.scrollHeight > 1115 && fs > 7.5){ fs -= 0.2; p.style.setProperty("--fs", fs + "pt"); }
  });
}

function showOne(){
  const k = document.getElementById("knr").value.trim();
  if(DATA[k]) { document.getElementById("out").innerHTML = render(DATA[k]); autoFit(); }
}

document.getElementById("list").innerHTML = ORDER.map(k=>`<div class="item" onclick="document.getElementById('knr').value='${k}';showOne()"><b>${k}</b> - ${DATA[k].name}</div>`).join("");
</script>
</body>
</html>
"""

# --- STREAMLIT APP ---
st.set_page_config(page_title="Sendeplan A4 Fix", layout="wide")

up = st.file_uploader("Excel Datei laden", type=["xlsx"])
if up:
    df = pd.read_excel(up)
    cols = df.columns.tolist()
    trip = detect_triplets(cols)
    bmap = detect_bspalten(cols)
    ds_trip = detect_ds_triplets(cols)
    
    data = {}
    for _, r in df.iterrows():
        knr = norm(r.get("Nr", ""))
        if not knr: continue
        
        bestell = []
        for d_de in DAYS_DE:
            # Fleisch/Wurst (ID 21)
            if d_de in trip and "21" in trip[d_de]:
                f = trip[d_de]["21"]
                bestell.append({"liefertag": d_de, "sortiment": norm(r.get(f.get("Sort"))), "bestelltag": norm(r.get(f.get("Tag"))), "bestellschluss": normalize_time(r.get(f.get("Zeit"))), "prio": 0})
            
            # B-Spalten (Wiesenhof, Bio, Frischfleisch, Avo, Werbe)
            keys = [k for k in bmap.keys() if k[0] == d_de]
            for k in sorted(keys, key=lambda x: str(x[1])):
                f = bmap[k]
                s = norm(r.get(f.get("sort", "")))
                z = normalize_time(r.get(f.get("zeit", "")))
                if s or z:
                    bestell.append({"liefertag": d_de, "sortiment": s, "bestelltag": k[2], "bestellschluss": z, "prio": 1})
            
            # Deutsche See
            if d_de in ds_trip:
                for key_ds in ds_trip[d_de]:
                    f = ds_trip[d_de][key_ds]
                    bestell.append({"liefertag": d_de, "sortiment": norm(r.get(f.get("Sort"))), "bestelltag": norm(r.get(f.get("Tag"))), "bestellschluss": normalize_time(r.get(f.get("Zeit"))), "prio": 2})

        data[knr] = {
            "plan_typ": PLAN_TYP, "bereich": BEREICH, "kunden_nr": knr,
            "name": norm(r.get("Name", "")), "strasse": norm(r.get("Strasse", "")),
            "plz": norm(r.get("Plz", "")), "ort": norm(r.get("Ort", "")),
            "fachberater": norm(r.get("Fachberater", "")),
            "tours": {d: norm(r.get(TOUR_COLS[d], "")) for d in DAYS_DE},
            "bestell": bestell
        }

    html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(data, separators=(',', ':')))
    st.download_button("Download Sendeplan", data=html, file_name="sendeplan.html", mime="text/html")
