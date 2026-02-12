# app.py
# ------------------------------------------------------------
# Excel -> Standalone-HTML (Deutsche See integriert in Wochentage)
# Ziel: Pro Kunde GENAU 1x A4 Seite (Schriftgröße passt sich an)
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
    """Erkennt Spalten wie 'Mo Fleisch Zeit', 'Mo Fleisch Sort' etc."""
    rx = re.compile(r"^(Mo|Die|Di|Mitt|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)
    found = {}
    for c in [c.strip() for c in columns]:
        m = rx.match(c)
        if not m: continue
        day_short, group, field = m.group(1), m.group(2).strip(), m.group(3).capitalize()
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
        day_de = DAY_SHORT_TO_DE.get(day_raw) or day_raw
        if day_de in DAYS_DE:
            tmp.setdefault(day_de, {})[field] = c
    return tmp

# --- HTML TEMPLATE MIT AUTO-FIT LOGIK ---
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
  *{ box-sizing:border-box; font-family: ui-sans-serif, system-ui, sans-serif; }
  body{ margin:0; background: var(--bg); color:var(--text); }
  
  /* Sidebar & Main Layout */
  .app{ display:grid; grid-template-columns: 340px 1fr; height:100vh; padding:15px; gap:15px; }
  .sidebar, .main{ background: var(--panel); border:1px solid var(--stroke); border-radius:12px; overflow:hidden; backdrop-filter: blur(10px); }
  .list{ height: calc(100vh - 250px); overflow-y:auto; border-top:1px solid var(--stroke); }
  .item{ padding:10px 15px; border-bottom:1px solid rgba(255,255,255,0.05); cursor:pointer; font-size:13px; }
  .item:hover{ background:rgba(255,255,255,0.08); }
  .main{ display:flex; flex-direction:column; }
  .wrap{ flex:1; overflow-y:auto; padding:20px; display:flex; flex-direction:column; align-items:center; }
  
  /* A4 Papier Simulation & Druck */
  .paper{
    width:210mm; height:297mm; min-height:297mm; background:white; color:black; padding:15mm;
    position:relative; box-shadow: 0 0 30px rgba(0,0,0,0.5); page-break-after: always;
    display: flex; flex-direction: column; --fs: 10.5pt;
  }
  .paper * { font-size: var(--fs); line-height: 1.25; }
  .ptitle{ text-align:center; font-weight:950; font-size:1.6em; margin:0; }
  .pstd{ text-align:center; color:#d0192b; font-weight:bold; margin:2mm 0; font-size:1.2em; }
  .psub{ text-align:center; color:var(--sub); margin-bottom:8mm; font-weight:600; }
  
  .head{ display:flex; justify-content:space-between; margin-bottom:6mm; }
  .head-left b { font-size: 1.1em; }
  
  table{ width:100%; border-collapse:collapse; margin-top: 2mm; }
  th, td{ border:1px solid #000; padding:1.8mm; text-align:left; vertical-align:top; }
  th{ background:#f2f2f2; font-weight:bold; }
  
  .ds-tag { color: #0056b3; font-weight: bold; font-size: 0.85em; vertical-align: middle; }

  @media print{
    body { background: none; }
    .sidebar { display:none !important; }
    .app { display:block; padding:0; }
    .main { border:none; background:none; }
    .wrap { padding:0; overflow:visible; }
    .paper { box-shadow:none; margin:0; border:none; width:210mm; height:297mm; }
  }
</style>
</head>
<body>
<div class="app">
  <div class="sidebar">
    <div style="padding:15px"><b>📄 Sendeplan-Generator</b></div>
    <div style="padding:15px; display:flex; flex-direction:column; gap:10px;">
      <input id="knr" placeholder="Kundennummer eingeben..." style="padding:10px; border-radius:8px; border:1px solid var(--stroke); background:rgba(0,0,0,0.2); color:white;">
      <div style="display:grid; grid-template-columns: 1fr 1fr; gap:8px;">
        <button onclick="showOne()" style="padding:10px; cursor:pointer; border-radius:8px; border:none; background:#4fa3ff; color:white; font-weight:bold;">Anzeigen</button>
        <button onclick="showAll()" style="padding:10px; cursor:pointer; border-radius:8px; border:none; background:rgba(255,255,255,0.1); color:white;">Alle laden</button>
      </div>
      <button onclick="window.print()" style="padding:10px; background:#28a745; color:white; border:none; border-radius:8px; font-weight:bold; cursor:pointer;">Jetzt Drucken (A4)</button>
    </div>
    <div class="list" id="list"></div>
  </div>
  <div class="main">
    <div id="mh_info" style="padding:10px; font-size:12px; color:var(--muted); text-align:center;">Excel hochgeladen. Bitte Kunden wählen.</div>
    <div class="wrap" id="out"></div>
  </div>
</div>

<script>
const DATA = __DATA_JSON__;
const ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0) - (Number(b)||0));
const DAYS = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];

function esc(s){ return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }

function render(c){
  const byDay = {};
  c.bestell.forEach(it => { if(!byDay[it.liefertag]) byDay[it.liefertag]=[]; byDay[it.liefertag].push(it); });

  const rows = DAYS.map(d => {
    const items = byDay[d] || [];
    return `<tr>
      <td style="width:18%"><b>${d}</b></td>
      <td>${items.map(x => (x.is_ds ? '<span class="ds-tag">[DS] </span>' : '') + esc(x.sortiment)).join("<br>") || "-"}</td>
      <td style="width:18%">${items.map(x => esc(x.bestelltag)).join("<br>") || "-"}</td>
      <td style="width:18%">${items.map(x => esc(x.bestellschluss)).join("<br>") || "-"}</td>
    </tr>`;
  }).join("");

  const activeTours = DAYS.map(d => c.tours[d] ? esc(c.tours[d]) : "—").join(" | ");

  return `
    <div class="paper">
      <div class="ptitle">Sende- &amp; Belieferungsplan</div>
      <div class="pstd">${esc(c.plan_typ)}</div>
      <div class="psub">${esc(c.bereich)}</div>
      <div class="head">
        <div class="head-left">
          <b>${esc(c.name)}</b><br>
          ${esc(c.strasse)}<br>
          ${esc(c.plz)} ${esc(c.ort)}
        </div>
        <div style="text-align:right">
          Kunden-Nr: <b>${esc(c.kunden_nr)}</b><br>
          Fachberater: ${esc(c.fachberater || "—")}
        </div>
      </div>
      <div style="margin-bottom:4mm; border-top:1px solid #eee; padding-top:2mm;">
        <b>Wochentag-Touren:</b><br>
        <small style="color:#666">${DAYS.map(d=>d.substring(0,2)).join(" | ")}</small><br>
        ${activeTours}
      </div>
      <table>
        <thead><tr><th>Liefertag</th><th>Sortiment</th><th>Bestelltag</th><th>Bestellschluss</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div style="margin-top:auto; font-size:8pt; color:#999; text-align:center;">
        Erstellt am ${new Date().toLocaleDateString('de-DE')}
      </div>
    </div>`;
}

function autoFit(){
  document.querySelectorAll(".paper").forEach(p => {
    let fs = 10.5; p.style.setProperty("--fs", fs + "pt");
    // Reduziere Schriftgröße, bis alles auf die Seite passt
    let safety = 0;
    while(p.scrollHeight > p.clientHeight + 2 && fs > 6.5 && safety < 40){
      fs -= 0.1; p.style.setProperty("--fs", fs + "pt");
      safety++;
    }
  });
}

function showOne(){
  const k = document.getElementById("knr").value.trim();
  if(!DATA[k]) return;
  document.getElementById("out").innerHTML = render(DATA[k]);
  autoFit();
}

function showAll(){
  document.getElementById("out").innerHTML = ORDER.map(k=>render(DATA[k])).join("");
  autoFit();
}

document.getElementById("list").innerHTML = ORDER.map(k=>`
  <div class="item" onclick="document.getElementById('knr').value='${k}';showOne()">
    <b>${k}</b> - ${esc(DATA[k].name)}
  </div>`).join("");
</script>
</body>
</html>
"""

# --- STREAMLIT GENERATOR ---
st.set_page_config(page_title="Plan-Generator", layout="wide")
st.title("Excel → A4 Druckvorlage (Inkl. Deutsche See)")

up = st.file_uploader("Excel Datei hochladen (.xlsx)", type=["xlsx"])
if up:
    df = pd.read_excel(up)
    trip = detect_triplets(df.columns)
    ds_trip = detect_ds_triplets(df.columns)
    
    data_json = {}
    for _, r in df.iterrows():
        knr = norm(r.get("Nr", ""))
        if not knr: continue
        
        bestell = []
        # 1. Normale Fleischwerk-Sortimente sammeln
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

        # Nach Wochentag sortieren für die Tabelle
        bestell.sort(key=lambda x: DAYS_DE.index(x["liefertag"]))

        data_json[knr] = {
            "plan_typ": PLAN_TYP, "bereich": BEREICH, "kunden_nr": knr,
            "name": norm(r.get("Name", "")), "strasse": norm(r.get("Strasse", "")),
            "plz": norm(r.get("Plz", "")), "ort": norm(r.get("Ort", "")),
            "fachberater": norm(r.get("Fachberater", "")),
            "tours": {d: norm(r.get(c, "")) for d, c in TOUR_COLS.items() if c in df.columns},
            "bestell": bestell
        }

    st.success(f"{len(data_json)} Kunden geladen.")
    
    # JSON in das HTML einbetten
    final_html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(data_json, separators=(',', ':')))
    
    st.download_button(
        label="⬇️ Standalone-HTML Datei speichern",
        data=final_html,
        file_name="sendeplan_A4_komplett.html",
        mime="text/html"
    )

st.info("Hinweis: Nach dem Öffnen der HTML-Datei können Sie über die Suche einzelne Kunden wählen oder per 'Alle laden' den gesamten Massendruck vorbereiten.")
