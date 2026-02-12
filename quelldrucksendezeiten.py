# app.py
# -----------------------------------------------------------------------------
# Excel -> Standalone-HTML (Präzisions-Extraktion & A4-Druck-Optimierung)
# -----------------------------------------------------------------------------

import json
import re
import datetime
from typing import Dict, Tuple, List
import pandas as pd
import streamlit as st

# Konfiguration
PLAN_TYP = "Standard"
BEREICH = "Alle Sortimente Fleischwerk"
DAYS_DE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]

# Mapping für Wochentags-Abkürzungen aus der Excel
DAY_SHORT_TO_DE = {
    "Mo": "Montag", "Di": "Dienstag", "Die": "Dienstag",
    "Mi": "Mittwoch", "Mit": "Mittwoch", "Mitt": "Mittwoch",
    "Do": "Donnerstag", "Don": "Donnerstag", "Donn": "Donnerstag",
    "Fr": "Freitag", "Sa": "Samstag", "Sam": "Samstag",
}

# Spalten für die Tour-Daten (1030, 2027, etc.)
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

# --- Detektions-Logiken für die Excel-Spalten ---

def detect_triplets(columns: List[str]):
    """Erkennt Standard-Tripel wie 'Mo 21 Zeit', 'Mo 21 Sort', 'Mo 21 Tag'."""
    rx = re.compile(r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)
    found = {}
    for c in columns:
        m = rx.match(c.strip())
        if m:
            day_de = DAY_SHORT_TO_DE.get(m.group(1))
            if day_de:
                found.setdefault(day_de, {}).setdefault(m.group(2).strip(), {})[m.group(3).capitalize()] = c
    return found

def detect_bspalten(columns: List[str]):
    """Erkennt B-Spalten wie 'Mo Z 0 B_Sa', 'Mo 1011 B_Fr' oder 'Donn Z 22 B_Mo'."""
    rx = re.compile(
        r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+"  # Liefertag
        r"(?:(Z|L)\s+)?"                                     # Optional: Z (Zeit) oder L
        r"(.+?)\s+"                                          # ID (z.B. 0, 1011, 22)
        r"B[_ ]?(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)$", # Bestelltag
        re.IGNORECASE
    )
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

def detect_ds_triplets(columns: List[str]):
    """Erkennt Deutsche See Spalten: 'DS [Junk] zu [Liefertag] [Field]'."""
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

# --- HTML TEMPLATE MIT DRUCK-FIXES ---
HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<style>
  @page { size: A4; margin: 0; }
  :root{ --bg:#0b1220; --stroke:rgba(255,255,255,.14); --text:rgba(255,255,255,.92); }
  *{ box-sizing:border-box; font-family: sans-serif; }
  body{ margin:0; background: var(--bg); color:var(--text); }
  .app{ display:grid; grid-template-columns: 350px 1fr; height:100vh; padding:15px; gap:15px; }
  .sidebar, .main{ background: rgba(255,255,255,.08); border:1px solid var(--stroke); border-radius:12px; overflow:hidden; }
  .list{ height: calc(100vh - 260px); overflow-y:auto; border-top:1px solid var(--stroke); }
  .item{ padding:10px; border-bottom:1px solid rgba(255,255,255,0.05); cursor:pointer; font-size:12px; }
  .item:hover{ background:rgba(255,255,255,0.08); }
  .wrap{ overflow-y:auto; padding:20px; display:flex; flex-direction:column; align-items:center; }
  
  .paper{
    width: 210mm;
    height: 296.5mm; /* Verhindert leere Extraseite */
    background: white; color: black; padding: 12mm;
    position: relative; box-shadow: 0 0 20px rgba(0,0,0,0.5);
    display: flex; flex-direction: column; overflow: hidden;
    --fs: 10.2pt;
  }
  .paper * { font-size: var(--fs); line-height: 1.15; }
  .ptitle{ text-align:center; font-weight:900; font-size:1.6em; margin:0; }
  .pstd{ text-align:center; color:#d0192b; font-weight:bold; margin:1mm 0; font-size:1.3em; }
  .psub{ text-align:center; color:#555; margin-bottom:4mm; font-weight:bold; }
  .head-box { display:flex; justify-content:space-between; margin-bottom:4mm; border-bottom:1px solid #eee; padding-bottom:3mm; }
  
  .tour-bar { display:flex; background:#f4f4f4; border:1px solid #ddd; margin-bottom:4mm; padding:2mm; border-radius:4px; justify-content:space-around; }
  .tour-item { text-align:center; font-size:0.85em; }
  .tour-item b { display:block; font-size:0.75em; color:#666; }

  table{ width:100%; border-collapse:collapse; }
  th, td{ border:1px solid #000; padding:1.4mm; text-align:left; vertical-align:top; }
  th{ background:#f2f2f2; font-weight:bold; }

  @media print{
    .sidebar{ display:none !important; }
    .app{ display:block; padding:0; margin:0; }
    .main{ background:none; border:none; }
    .wrap{ padding:0; display:block; }
    .paper{
      box-shadow: none; margin: 0; border: none;
      page-break-inside: avoid;
      page-break-after: always;
    }
    .paper:last-child { page-break-after: auto !important; }
  }
</style>
</head>
<body>
<div class="app">
  <div class="sidebar">
    <div style="padding:15px"><b>Sendeplan Generator</b></div>
    <div style="padding:15px; display:flex; flex-direction:column; gap:8px;">
      <input id="knr" placeholder="Kunden-Nr..." style="width:100%; padding:8px;">
      <button onclick="showOne()" style="padding:8px; cursor:pointer;">Anzeigen</button>
      <button onclick="showAll()" style="padding:8px; cursor:pointer;">Alle laden</button>
      <button onclick="window.print()" style="padding:8px; background:#28a745; color:white; border:none; cursor:pointer;">Drucken</button>
    </div>
    <div class="list" id="list"></div>
  </div>
  <div class="main"><div class="wrap" id="out"></div></div>
</div>
<script>
const DATA = __DATA_JSON__;
const ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0)-(Number(b)||0));
const DAYS = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];

function esc(s){ return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;"); }

function render(c){
  const byDay = {};
  c.bestell.forEach(it => { if(!byDay[it.liefertag]) byDay[it.liefertag]=[]; byDay[it.liefertag].push(it); });
  
  const rows = DAYS.map(d => {
    const items = byDay[d] || [];
    return `<tr>
      <td style="width:16%"><b>${d}</b></td>
      <td>${items.map(x => esc(x.sortiment)).join("<br>") || "&nbsp;"}</td>
      <td style="width:16%">${items.map(x => esc(x.bestelltag)).join("<br>") || "&nbsp;"}</td>
      <td style="width:18%">${items.map(x => esc(x.bestellschluss)).join("<br>") || "&nbsp;"}</td>
    </tr>`;
  }).join("");

  const tourHtml = DAYS.map(d => `<div class="tour-item"><b>${d}</b>${esc(c.tours[d] || "—")}</div>`).join("");

  return `<div class="paper">
    <div class="ptitle">Sende- &amp; Belieferungsplan</div>
    <div class="pstd">${esc(c.plan_typ)}</div>
    <div class="psub">${esc(c.bereich)}</div>
    <div class="head-box">
      <div><b>${esc(c.name)}</b><br>${esc(c.strasse)}<br>${esc(c.plz)} ${esc(c.ort)}</div>
      <div style="text-align:right">Kunden-Nr: <b>${esc(c.kunden_nr)}</b><br>Fachberater: <b>${esc(c.fachberater)}</b></div>
    </div>
    <div style="font-size:0.8em; margin-bottom:1mm; font-weight:bold;">Tourenplan (Liefertage):</div>
    <div class="tour-bar">${tourHtml}</div>
    <table>
      <thead><tr><th>Liefertag</th><th>Sortiment</th><th>Bestelltag</th><th>Bestellzeitende</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

function autoFit(){
  document.querySelectorAll(".paper").forEach(p => {
    let fs = 10.2; p.style.setProperty("--fs", fs + "pt");
    let safety = 0;
    while(p.scrollHeight > p.clientHeight && fs > 6.8 && safety < 40){
      fs -= 0.1; p.style.setProperty("--fs", fs.toFixed(1) + "pt"); safety++;
    }
  });
}

function showOne(){
  const k = document.getElementById("knr").value.trim();
  if(DATA[k]) { document.getElementById("out").innerHTML = render(DATA[k]); autoFit(); }
}

function showAll(){
  document.getElementById("out").innerHTML = ORDER.map(k=>render(DATA[k])).join("");
  autoFit();
}

document.getElementById("list").innerHTML = ORDER.map(k=>`<div class="item" onclick="document.getElementById('knr').value='${k}';showOne()"><b>${k}</b> - ${esc(DATA[k].name)}</div>`).join("");
</script>
</body>
</html>
"""

# --- STREAMLIT APP ---
st.set_page_config(page_title="Sendeplan Generator", layout="wide")
st.title("Sendeplan Generator (Präzisions-Version)")

up = st.file_uploader("Excel Datei wählen", type=["xlsx"])
if up:
    df = pd.read_excel(up)
    cols = df.columns.tolist()
    
    # Detektoren initialisieren
    trip = detect_triplets(cols)
    bmap = detect_bspalten(cols)
    ds_trip = detect_ds_triplets(cols)
    
    data = {}
    for _, r in df.iterrows():
        knr = norm(r.get("Nr", ""))
        if not knr: continue
        
        bestell = []
        for d_de in DAYS_DE:
            # 1. B-Spalten (Sortiments-IDs wie 0, 1011, 22, 91)
            keys = [k for k in bmap.keys() if k[0] == d_de]
            for k in keys:
                f = bmap[k]
                s = norm(r.get(f.get("sort", "")))
                z = normalize_time(r.get(f.get("zeit", "")))
                if s or z:
                    bestell.append({"liefertag": d_de, "sortiment": s, "bestelltag": k[2], "bestellschluss": z, "priority": 1})
            
            # 2. Standard-Tripel (z.B. ID 21)
            if d_de in trip:
                for g in trip[d_de]:
                    f = trip[d_de][g]
                    s, t, z = norm(r.get(f.get("Sort"))), norm(r.get(f.get("Tag"))), normalize_time(r.get(f.get("Zeit")))
                    if s or t or z:
                        bestell.append({"liefertag": d_de, "sortiment": s, "bestelltag": t, "bestellschluss": z, "priority": 2})

            # 3. Deutsche See Integration (DS)
            if d_de in ds_trip:
                for key_ds in ds_trip[d_de]:
                    f = ds_trip[d_de][key_ds]
                    s, t, z = norm(r.get(f.get("Sort"))), norm(r.get(f.get("Tag"))), normalize_time(r.get(f.get("Zeit")))
                    if s or t or z:
                        bestell.append({"liefertag": d_de, "sortiment": s, "bestelltag": t, "bestellschluss": z, "priority": 3})

        # Sortierung innerhalb eines Tages: Standard-Sortimente zuerst
        bestell.sort(key=lambda x: x["priority"])

        data[knr] = {
            "plan_typ": PLAN_TYP, "bereich": BEREICH, "kunden_nr": knr,
            "name": norm(r.get("Name", "")), "strasse": norm(r.get("Strasse", "")),
            "plz": norm(r.get("Plz", "")), "ort": norm(r.get("Ort", "")),
            "fachberater": norm(r.get("Fachberater", "")),
            "tours": {d: norm(r.get(TOUR_COLS[d], "")) for d in DAYS_DE},
            "bestell": bestell
        }

    # HTML generieren
    html_content = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(data, separators=(',', ':')))
    st.download_button("⬇️ HTML-Sendeplan herunterladen", data=html_content, file_name="sendeplan.html", mime="text/html")
