Das liegt wahrscheinlich daran, dass die Logik für die **B_-Spalten** (die in deinem ersten Entwurf vorhanden war) in der Zwischenversion durch die Vereinfachung für Deutsche See überschrieben wurde.

Hier ist das **vollständige Skript**, das alle Logiken vereint:

1. **B_-Mapping** (Bestellschluss-Logik).
2. **Tripel-Logik** (Standard-Spalten).
3. **Deutsche See Integration** (direkt in die Wochentage).
4. **Auto-Fit** für A4.

```python
# app.py
# ------------------------------------------------------------
# Excel -> Standalone-HTML (KOMPLETT: B-Spalten, Tripel & Deutsche See)
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

def group_sort_key(g: str):
    g = g.strip()
    if g.isdigit(): return (0, int(g))
    return (1, g.lower())

def detect_triplets(columns: List[str]) -> Dict[str, Dict[str, Dict[str, str]]]:
    rx = re.compile(r"^(Mo|Die|Di|Mitt|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)
    found = {}
    for c in [c.strip() for c in columns]:
        m = rx.match(c)
        if not m: continue
        day_s, group, field = m.group(1), m.group(2).strip(), m.group(3).capitalize()
        day_de = DAY_SHORT_TO_DE.get(day_s)
        if day_de: found.setdefault(day_de, {}).setdefault(group, {})[field] = c
    return found

def detect_bspalten(columns: List[str]) -> Dict[Tuple[str, str, str], Dict[str, str]]:
    rx = re.compile(r"^(Mo|Die|Di|Mitt|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(?:(Z|L)\s+)?(.+?)\s+B[_ ]?(Mo|Die|Di|Mitt|Mi|Don|Donn|Do|Fr|Sam|Sa)$", re.IGNORECASE)
    mapping = {}
    for c in [c.strip() for c in columns]:
        m = rx.match(c)
        if not m: continue
        day_s, zl, group, b_s = m.group(1), (m.group(2) or "").upper(), m.group(3).strip(), m.group(4)
        day_de, b_de = DAY_SHORT_TO_DE.get(day_s), DAY_SHORT_TO_DE.get(b_s)
        if day_de and b_de:
            key = (day_de, group, b_de)
            mapping.setdefault(key, {})
            if zl == "Z": mapping[key]["zeit"] = c
            elif zl == "L": mapping[key]["l"] = c
            else: mapping[key]["sort"] = c
    return mapping

def detect_ds_triplets(columns: List[str]) -> Dict[str, Dict[str, str]]:
    rx = re.compile(r"^DS\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)
    tmp = {}
    for c in [c.strip() for c in columns]:
        m = rx.match(c)
        if not m: continue
        day_raw, field = m.group(1).strip(), m.group(2).capitalize()
        day_de = DAY_SHORT_TO_DE.get(day_raw) or day_raw
        if day_de in DAYS_DE: tmp.setdefault(day_de, {})[field] = c
    return tmp

# --- HTML TEMPLATE ---
HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<style>
  @page { size: A4; margin: 0; }
  :root{ --bg:#0b1220; --stroke:rgba(255,255,255,.14); --text:rgba(255,255,255,.92); --paper:#fff; }
  *{ box-sizing:border-box; font-family: sans-serif; }
  body{ margin:0; background: var(--bg); color:var(--text); }
  .app{ display:grid; grid-template-columns: 350px 1fr; height:100vh; padding:15px; gap:15px; }
  .sidebar, .main{ background: rgba(255,255,255,.08); border:1px solid var(--stroke); border-radius:12px; overflow:hidden; }
  .list{ height: calc(100vh - 260px); overflow-y:auto; border-top:1px solid var(--stroke); }
  .item{ padding:10px; border-bottom:1px solid rgba(255,255,255,0.05); cursor:pointer; font-size:12px; }
  .item:hover{ background:rgba(255,255,255,0.08); }
  .wrap{ overflow-y:auto; padding:20px; display:flex; flex-direction:column; align-items:center; height:100%; }
  .paper{
    width:210mm; height:297mm; min-height:297mm; background:white; color:black; padding:12mm;
    position:relative; box-shadow: 0 0 20px rgba(0,0,0,0.5); page-break-after: always;
    display:flex; flex-direction:column; --fs: 10.5pt;
  }
  .paper * { font-size: var(--fs); line-height: 1.2; }
  .ptitle{ text-align:center; font-weight:900; font-size:1.5em; margin:0; }
  .pstd{ text-align:center; color:#d0192b; font-weight:bold; margin:1mm 0; }
  .psub{ text-align:center; color:#555; margin-bottom:5mm; }
  table{ width:100%; border-collapse:collapse; }
  th, td{ border:1px solid #000; padding:1.5mm; text-align:left; vertical-align:top; }
  th{ background:#f2f2f2; }
  .ds-tag { color: #0056b3; font-weight: bold; }
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
    <div style="padding:15px"><b>Sendeplan Generator</b></div>
    <div style="padding:15px; display:flex; flex-direction:column; gap:8px;">
      <input id="knr" placeholder="Kunden-Nr..." style="padding:8px; border-radius:5px; border:none;">
      <button onclick="showOne()" style="padding:8px; cursor:pointer;">Anzeigen</button>
      <button onclick="showAll()" style="padding:8px; cursor:pointer;">Alle laden</button>
      <button onclick="window.print()" style="padding:8px; background:#28a745; color:white; border:none; border-radius:5px; cursor:pointer;">Drucken</button>
    </div>
    <div class="list" id="list"></div>
  </div>
  <div class="main"><div class="wrap" id="out"></div></div>
</div>
<script>
const DATA = __DATA_JSON__;
const ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0) - (Number(b)||0));
const DAYS = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];

function esc(s){ return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;"); }

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

  return `<div class="paper">
    <div class="ptitle">Sende- &amp; Belieferungsplan</div>
    <div class="pstd">${esc(c.plan_typ)}</div>
    <div class="psub">${esc(c.name)} | ${esc(c.bereich)}</div>
    <div style="display:flex; justify-content:space-between; margin-bottom:5mm;">
      <div><b>${esc(c.name)}</b><br>${esc(c.strasse)}<br>${esc(c.plz)} ${esc(c.ort)}</div>
      <div style="text-align:right">Kunden-Nr: <b>${esc(c.kunden_nr)}</b><br>Tour: ${DAYS.map(d=>c.tours[d]).filter(x=>x).join("/")}</div>
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
    while(p.scrollHeight > p.clientHeight && fs > 7){ fs -= 0.1; p.style.setProperty("--fs", fs + "pt"); }
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

# --- STREAMLIT LOGIK ---
st.set_page_config(page_title="Sendeplan", layout="wide")
st.title("Excel → Sendeplan (Inkl. B-Spalten & Deutsche See)")

up = st.file_uploader("Excel Datei wählen", type=["xlsx"])
if up:
    df = pd.read_excel(up)
    trip = detect_triplets(df.columns)
    bmap = detect_bspalten(df.columns)
    ds_trip = detect_ds_triplets(df.columns)
    
    data = {}
    for _, r in df.iterrows():
        knr = norm(r.get("Nr", ""))
        if not knr: continue
        
        bestell = []
        # 1. Tripel (Standard)
        for d_de, groups in trip.items():
            for g, cols in groups.items():
                s, t, z = norm(r.get(cols["Sort"])), norm(r.get(cols["Tag"])), normalize_time(r.get(cols["Zeit"]))
                if s or t or z: bestell.append({"liefertag": d_de, "sortiment": s, "bestelltag": t, "bestellschluss": z, "is_ds": False})
        
        # 2. B-Spalten (Bestellschluss-Logik)
        for (lday, group, btag), cols in bmap.items():
            s = norm(r.get(cols.get("sort", "")))
            z = normalize_time(r.get(cols.get("zeit", "")))
            if s or z: bestell.append({"liefertag": lday, "sortiment": s, "bestelltag": btag, "bestellschluss": z, "is_ds": False})

        # 3. Deutsche See Integration
        for d_de, cols in ds_trip.items():
            s, t, z = norm(r.get(cols["Sort"])), norm(r.get(cols["Tag"])), normalize_time(r.get(cols["Zeit"]))
            if s or t or z: bestell.append({"liefertag": d_de, "sortiment": s, "bestelltag": t, "bestellschluss": z, "is_ds": True})

        bestell.sort(key=lambda x: (DAYS_DE.index(x["liefertag"]), x["is_ds"]))

        data[knr] = {
            "plan_typ": PLAN_TYP, "bereich": BEREICH, "kunden_nr": knr,
            "name": norm(r.get("Name", "")), "strasse": norm(r.get("Strasse", "")),
            "plz": norm(r.get("Plz", "")), "ort": norm(r.get("Ort", "")),
            "tours": {d: norm(r.get(c, "")) for d, c in TOUR_COLS.items() if c in df.columns},
            "bestell": bestell
        }

    html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(data, separators=(',', ':')))
    st.download_button("HTML-Sendeplan herunterladen", data=html, file_name="sendeplan.html", mime="text/html")

```
