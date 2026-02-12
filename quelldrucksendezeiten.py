# quelldrucksendezeiten.py
# -----------------------------------------------------------------------------
# VERSION: ABSOLUTE PRÄZISION - 100% DATEN-ÜBEREINSTIMMUNG & A4 OPTIMIERT
# -----------------------------------------------------------------------------

import json
import re
import datetime
from typing import Dict, Tuple, List
import pandas as pd
import streamlit as st

# Grundkonfiguration
PLAN_TYP = "Standard"
BEREICH = "Alle Sortimente Fleischwerk"
DAYS_DE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]

# Mapping für Excel-Kürzel
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

# --- Daten-Extraktions-Logik ---

def detect_bspalten(columns: List[str]):
    # Findet ID-basierte Sortimente (0=Avo, 1011=Wiesenhof, 91=Werbemittel, 65=Frischfleisch, 41=Bio)
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
    # Für Standard-Sortimente (ID 21 = Fleisch/Wurst)
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
    # Deutsche See
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

# --- HTML TEMPLATE ---
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
  .item{ padding:10px; border-bottom:1px solid rgba(255,255,255,0.05); cursor:pointer; font-size:13px; }
  .wrap{ height: 100%; overflow-y: scroll; padding: 20px; display: flex; flex-direction: column; align-items: center; }

  .paper{
    width: 210mm; height: 296.5mm; background: white; color: black; padding: 12mm;
    box-shadow: 0 0 20px rgba(0,0,0,0.5); display: flex; flex-direction: column;
    --fs: 11.5pt;
  }
  .paper * { font-size: var(--fs); line-height: 1.4; }
  .ptitle{ text-align:center; font-weight:950; font-size:1.8em; margin:0; }
  .pstd{ text-align:center; color:#d0192b; font-weight:bold; margin:1mm 0; font-size:1.4em; }
  .psub{ text-align:center; color:#444; margin-bottom:5mm; font-weight:bold; }
  
  .head-box { display:flex; justify-content:space-between; margin-bottom:4mm; border-bottom:2px solid #000; padding-bottom:3mm; }
  .head-box b { font-size: 1.1em; }

  .tour-info { margin-bottom: 5mm; }
  .tour-table { width: 100%; border-collapse: collapse; margin-top: 1mm; }
  .tour-table th { background: #eee; font-size: 0.8em; padding: 2px; border: 1px solid #000; text-transform: uppercase; }
  .tour-table td { border: 1px solid #000; padding: 6px; text-align: center; font-weight: bold; font-size: 1em; }

  table.main-table { width:100%; border-collapse:collapse; table-layout: fixed; border: 2px solid #000; }
  table.main-table th { border: 1px solid #000; padding: 8px; background:#f2f2f2; font-weight:bold; text-align:left; }
  table.main-table td { border: 1px solid #000; padding: 8px; vertical-align: top; }

  .day-header { background-color: #e0e0e0 !important; font-weight: 900; border-top: 4px solid #000 !important; }
  .day-label { font-size: 1.2em; }

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
      <button onclick="showOne()" style="padding:10px; background:#4fa3ff; color:white; border:none; cursor:pointer; font-weight:bold;">Anzeigen</button>
      <button onclick="window.print()" style="padding:10px; background:#28a745; color:white; border:none; cursor:pointer; font-weight:bold;">Drucken (A4)</button>
    </div>
    <div class="list" id="list"></div>
  </div>
  <div class="main"><div class="wrap" id="out">Bitte Kunden wählen...</div></div>
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
        const rowClass = idx === 0 ? "day-header" : "";
        tableRows += `<tr class="${rowClass}">
          <td style="width:18%">${idx === 0 ? `<span class="day-label">${d}</span>` : ""}</td>
          <td style="width:47%">${esc(it.sortiment)}</td>
          <td style="width:17%">${esc(it.bestelltag)}</td>
          <td style="width:18%">${esc(it.bestellschluss)}</td>
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
      <div><b>${esc(c.name)}</b><br>${esc(c.strasse)}<br>${esc(c.plz)} ${esc(c.ort)}</div>
      <div style="text-align:right">Kunden-Nr: <b>${esc(c.kunden_nr)}</b><br>Fachberater: <b>${esc(c.fachberater)}</b></div>
    </div>
    <div class="tour-info">
      <table class="tour-table"><thead><tr>${tourHead}</tr></thead><tbody><tr>${tourRow}</tr></tbody></table>
    </div>
    <table class="main-table">
      <thead><tr><th>Liefertag</th><th>Sortiment</th><th>Bestelltag</th><th>Bestellzeitende</th></tr></thead>
      <tbody>${tableRows}</tbody>
    </table>
  </div>`;
}

function autoFit(){
  document.querySelectorAll(".paper").forEach(p => {
    let fs = 11.5; p.style.setProperty("--fs", fs + "pt");
    let safety = 0;
    while(p.scrollHeight > 1115 && fs > 8 && safety < 50){ 
      fs -= 0.2; p.style.setProperty("--fs", fs.toFixed(1) + "pt"); safety++;
    }
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

# --- PYTHON GENERATOR ---
st.set_page_config(page_title="Sendeplan Fix", layout="wide")

up = st.file_uploader("Quelldatei laden", type=["xlsx"])
if up:
    df = pd.read_excel(up)
    cols = df.columns.tolist()
    trip, bmap, ds_trip = detect_triplets(cols), detect_bspalten(cols), detect_ds_triplets(cols)
    
    data = {}
    for _, r in df.iterrows():
        knr = norm(r.get("Nr", ""))
        if not knr: continue
        
        bestell = []
        for d_de in DAYS_DE:
            # 1. Fleisch/Wurst (ID 21)
            if d_de in trip and "21" in trip[d_de]:
                f = trip[d_de]["21"]
                bestell.append({"liefertag": d_de, "sortiment": norm(r.get(f.get("Sort"))), "bestelltag": norm(r.get(f.get("Tag"))), "bestellschluss": normalize_time(r.get(f.get("Zeit"))), "prio": 0})
            
            # 2. B-Spalten (Sortiments-IDs: 0=Avo, 1011=Wiesenhof, 41=Bio, 65=Frischfleisch, 91=Werbemittel)
            keys = [k for k in bmap.keys() if k[0] == d_de]
            # Hier wird nach ID sortiert, damit Wiesenhof immer an der richtigen Stelle kommt
            for k in sorted(keys, key=lambda x: str(x[1])):
                f = bmap[k]
                s = norm(r.get(f.get("sort", "")))
                z = normalize_time(r.get(f.get("zeit", "")))
                if s or z:
                    bestell.append({"liefertag": d_de, "sortiment": s, "bestelltag": k[2], "bestellschluss": z, "prio": 1})
            
            # 3. Deutsche See
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
    st.download_button("Sendeplan herunterladen", data=html, file_name="sendeplan.html", mime="text/html")
