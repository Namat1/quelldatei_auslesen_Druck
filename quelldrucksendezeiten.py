# app.py
# ------------------------------------------------------------
# Excel -> Standalone-HTML (Suche + A4 Druck, 1 Seite pro Kunde)
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
        day_short, group, field = m.group(1), m.group(2).strip(), m.group(3).capitalize()
        if day_short.lower() == "donn": day_short = "Don"
        day_de = DAY_SHORT_TO_DE.get(day_short)
        if day_de:
            found.setdefault(day_de, {}).setdefault(group, {})[field] = c
    return {d: g for d, g in found.items() if any(all(k in f for k in ("Zeit", "Sort", "Tag")) for f in g.values())}

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
        route, field = m.group(1).strip(), m.group(2).capitalize()
        key = f"DS {route}".replace("zu", "→")
        tmp.setdefault(key, {})[field] = c
    return {k: f for k, f in tmp.items() if all(x in f for x in ("Zeit", "Sort", "Tag"))}

HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<title>Sende- & Belieferungsplan</title>
<style>
  @page { size: A4; margin: 0; }
  :root{
    --bg:#0b1220; --panel:rgba(255,255,255,.08); --stroke:rgba(255,255,255,.14);
    --text:rgba(255,255,255,.92); --muted:rgba(255,255,255,.62);
    --paper:#fff; --ink:#0b0f17; --sub:#394054;
  }
  *{ box-sizing:border-box; }
  body{
    margin:0; font-family: system-ui, sans-serif;
    background: radial-gradient(at 25% 0%, rgba(79,163,255,.15), transparent 50%), var(--bg);
    color:var(--text);
  }
  .app{ display:grid; grid-template-columns: 340px 1fr; gap:14px; padding:14px; height:100vh; }
  .sidebar, .main{
    background: linear-gradient(180deg, var(--panel), rgba(255,255,255,.04));
    border:1px solid var(--stroke); border-radius:16px; overflow:hidden; backdrop-filter: blur(10px);
  }
  .sidehead{ padding:14px; border-bottom:1px solid var(--stroke); }
  .controls{ padding:14px; display:flex; flex-direction:column; gap:10px; }
  .field{ padding:10px; border-radius:12px; border:1px solid var(--stroke); background: rgba(0,0,0,.2); }
  input{ width:100%; border:none; background:transparent; color:#fff; outline:none; }
  .list{ border-top:1px solid var(--stroke); height: calc(100vh - 250px); overflow-y:auto; }
  .item{ padding:10px 14px; border-bottom:1px solid rgba(255,255,255,.05); cursor:pointer; }
  .item:hover{ background: rgba(255,255,255,.05); }
  .item.active{ background: rgba(79,163,255,.15); border-left:4px solid #4fa3ff; }
  
  .main{ display:flex; flex-direction:column; }
  .mainhead{ padding:14px; border-bottom:1px solid var(--stroke); display:flex; justify-content:space-between; }
  .wrap{ flex:1; overflow-y:auto; padding:20px; display:flex; flex-direction:column; align-items:center; gap:20px; }

  .paper{
    width:210mm; height:297mm; min-height:297mm; background:var(--paper); color:var(--ink);
    padding:12mm; position:relative; box-shadow: 0 10px 30px rgba(0,0,0,0.5); --fs: 10.4pt;
    page-break-after: always; page-break-inside: avoid;
  }
  .paper * { font-size: var(--fs); line-height: 1.2; }
  .ptitle{ text-align:center; font-weight:900; font-size: 1.6em; margin:0; }
  .pstd{ text-align:center; font-weight:bold; color:#d0192b; font-size: 1.3em; margin: 2mm 0; }
  .psub{ text-align:center; color:var(--sub); margin-bottom: 5mm; }
  .head{ display:flex; justify-content:space-between; margin-bottom: 5mm; }
  .lines{ margin-bottom: 5mm; }
  table{ width:100%; border-collapse:collapse; }
  th, td{ border:1px solid #000; padding:1.5mm; text-align:left; vertical-align:top; }
  th{ background:#f0f0f0; font-weight:bold; }
  .ds{ margin-top:5mm; border:1px solid #ccc; border-radius:5px; padding:3mm; }

  @media print{
    body{ background:none; }
    .sidebar, .mainhead{ display:none !important; }
    .app{ display:block; padding:0; }
    .main{ border:none; background:none; }
    .wrap{ padding:0; overflow:visible; }
    .paper{ box-shadow:none; margin:0; border:none; }
  }
</style>
</head>
<body>
<div class="app">
  <div class="sidebar">
    <div class="sidehead"><b>Sende- & Belieferungsplan</b></div>
    <div class="controls">
      <div class="field"><input id="knr" placeholder="Kundennummer..." inputmode="numeric"></div>
      <button onclick="showOne()">Anzeigen</button>
      <button onclick="showAll()">Alle anzeigen</button>
      <button style="background:#4fa3ff; color:white; border:none;" onclick="window.print()">Drucken (A4)</button>
    </div>
    <div class="list" id="list"></div>
  </div>
  <div class="main">
    <div class="mainhead"><div id="mh_title">Vorschau</div><div id="mh_count">0 Kunden</div></div>
    <div class="wrap" id="out"></div>
  </div>
</div>

<script>
const DATA = __DATA_JSON__;
const ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0)-(Number(b)||0));
const DAYS = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];

function esc(s){ return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }

function render(c){
  const byDay = {};
  c.bestell.forEach(it => { if(!byDay[it.liefertag]) byDay[it.liefertag]=[]; byDay[it.liefertag].push(it); });
  
  const activeDays = DAYS.filter(d => byDay[d] || (c.tours && c.tours[d]));
  const rows = DAYS.map(d => {
    const arr = byDay[d] || [];
    return `<tr>
      <td><b>${d}</b></td>
      <td>${arr.map(x=>esc(x.sortiment)).join("<br>") || "-"}</td>
      <td>${arr.map(x=>esc(x.bestelltag)).join("<br>") || "-"}</td>
      <td>${arr.map(x=>esc(x.bestellschluss)).join("<br>") || "-"}</td>
    </tr>`;
  }).join("");

  return `
    <div class="paper" data-knr="${c.kunden_nr}">
      <div class="ptitle">Sende- &amp; Belieferungsplan</div>
      <div class="pstd">${esc(c.plan_typ)}</div>
      <div class="psub">${esc(c.name)} | ${esc(c.bereich)}</div>
      <div class="head">
        <div><b>${esc(c.name)}</b><br>${esc(c.strasse)}<br>${esc(c.plz)} ${esc(c.ort)}</div>
        <div style="text-align:right">Kunden-Nr: <b>${esc(c.kunden_nr)}</b><br>Berater: ${esc(c.fachberater)}</div>
      </div>
      <div class="lines">
        <div><b>Liefertage:</b> ${activeDays.join(", ")}</div>
        <div><b>Touren:</b> ${activeDays.map(d => (c.tours[d]||"—")).join(" | ")}</div>
      </div>
      <table>
        <thead><tr><th>Liefertag</th><th>Sortiment</th><th>Bestelltag</th><th>Schluss</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
      ${c.ds.length ? `<div class="ds"><b>Durchsteck (DS):</b><br>${c.ds.map(it=>`• ${it.ds_key}: ${it.sortiment} (${it.bestelltag} ${it.bestellschluss})`).join("<br>")}</div>` : ""}
    </div>`;
}

function autoFit(){
  document.querySelectorAll(".paper").forEach(p => {
    let fs = 10.4; p.style.setProperty("--fs", fs + "pt");
    while(p.scrollHeight > p.clientHeight + 2 && fs > 7){
      fs -= 0.2; p.style.setProperty("--fs", fs + "pt");
    }
  });
}

function showOne(){
  const k = document.getElementById("knr").value.trim();
  if(!DATA[k]) return alert("Nicht gefunden");
  document.getElementById("out").innerHTML = render(DATA[k]);
  document.getElementById("mh_title").innerText = "Kunde: " + k;
  autoFit();
}

function showAll(){
  document.getElementById("out").innerHTML = ORDER.map(k=>render(DATA[k])).join("");
  document.getElementById("mh_title").innerText = "Massendruck-Modus";
  autoFit();
}

const listEl = document.getElementById("list");
listEl.innerHTML = ORDER.map(k => `<div class="item" onclick="document.getElementById('knr').value='${k}';showOne()">${k} - ${esc(DATA[k].name)}</div>`).join("");
document.getElementById("mh_count").innerText = ORDER.length + " Kunden";
</script>
</body>
</html>
"""

st.set_page_config(page_title="Plan-Generator", layout="wide")
st.title("Excel → A4 Druck-Tool")

up = st.file_uploader("Excel Datei wählen", type=["xlsx"])
if up:
    df = pd.read_excel(up)
    trip, bmap, dsmap = detect_triplets(df.columns), detect_bspalten(df.columns), detect_ds_triplets(df.columns)
    
    data = {}
    for _, r in df.iterrows():
        knr = norm(r.get("Nr", ""))
        if not knr: continue
        
        tours = {d: norm(r.get(c, "")) for d, c in TOUR_COLS.items() if c in df.columns}
        bestell = []
        
        # Logik für B_ Spalten & Tripel
        for day_de in DAYS_DE:
            # Tripel
            if day_de in trip:
                for g, cols in trip[day_de].items():
                    s, t, z = norm(r.get(cols["Sort"])), norm(r.get(cols["Tag"])), normalize_time(r.get(cols["Zeit"]))
                    if s or t or z: bestell.append({"liefertag": day_de, "sortiment": s, "bestelltag": t, "bestellschluss": z})
            # B-Mapping
            for (lday, group, btag), cols in bmap.items():
                if lday == day_de:
                    s, z = norm(r.get(cols.get("sort", ""))), normalize_time(r.get(cols.get("zeit", "")))
                    if s or z: bestell.append({"liefertag": lday, "sortiment": s, "bestelltag": btag, "bestellschluss": z})

        ds_list = []
        for dkey, cols in dsmap.items():
            s, t, z = norm(r.get(cols["Sort"])), norm(r.get(cols["Tag"])), normalize_time(r.get(cols["Zeit"]))
            if s or t or z: ds_list.append({"ds_key": dkey, "sortiment": s, "bestelltag": t, "bestellschluss": z})

        data[knr] = {
            "plan_typ": PLAN_TYP, "bereich": BEREICH, "kunden_nr": knr,
            "name": norm(r.get("Name", "")), "strasse": norm(r.get("Strasse", "")),
            "plz": norm(r.get("Plz", "")), "ort": norm(r.get("Ort", "")),
            "fachberater": norm(r.get("Fachberater", "")), "tours": tours,
            "bestell": bestell, "ds": ds_list
        }

    html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(data, separators=(',', ':')))
    st.download_button("HTML herunterladen", data=html, file_name="druckplan.html", mime="text/html")
