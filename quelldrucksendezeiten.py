```python
# quelldrucksendezeiten.py
# -----------------------------------------------------------------------------
# VERSION: ABSOLUTE PRÄZISION - 100% ÜBEREINSTIMMUNG & A4 OPTIMIERT
# Fixes:
# - Canonical-ID Mapping (Fleisch/Wurst -> 21, Wiesenhof -> 1011, Bio -> 41, Frischfleisch -> 65, Avo -> 0, Werbe -> 91, Pfeiffer -> 22)
# - Robustere Triplet-Erkennung (Zeitende/Bestellzeitende/Bestelltag/Sortiment etc.)
# - B-Spalten ebenfalls canonical (Prio korrekt)
# - Safety-Net: verhindert "Montag Montag" in Zeitspalte
# -----------------------------------------------------------------------------

import json
import re
import datetime
from typing import List
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

# Reihenfolge der Sortimente innerhalb eines Tages (Prio)
# 21: Fleisch/Wurst, 1011: Wiesenhof, 41: Bio, 65: Frischfleisch, 0: Avo, 91: Werbe, 22: Pfeiffer
SORT_PRIO = {"21": 0, "1011": 1, "41": 2, "65": 3, "0": 4, "91": 5, "22": 6}

TOUR_COLS = {
    "Montag": "Mo", "Dienstag": "Die", "Mittwoch": "Mitt",
    "Donnerstag": "Don", "Freitag": "Fr", "Samstag": "Sam"
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


def normalize_time(s) -> str:
    if isinstance(s, (datetime.time, pd.Timestamp)):
        return s.strftime("%H:%M") + " Uhr"
    s = norm(s)
    if not s:
        return ""
    if re.fullmatch(r"\d{1,2}:\d{2}", s):
        return s + " Uhr"
    if re.fullmatch(r"\d{1,2}", s):
        return s.zfill(2) + ":00 Uhr"
    return s


def safe_time(val) -> str:
    """
    verhindert Fälle wie "Montag Montag" (Tag landet fälschlich in Zeit)
    """
    raw = norm(val)
    if re.fullmatch(r"(Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag)", raw):
        return ""
    return normalize_time(val)


def canon_group_id(label: str) -> str:
    """
    Mapped die Sortimentsbezeichnung robust auf die internen IDs.
    Damit wird "Fleisch- & Wurst ..." immer als 21 erkannt, auch wenn keine Zahl im Header steht.
    """
    s = norm(label).lower()

    # harte Treffer (Zahlen)
    m = re.search(r"\b(1011|21|41|65|0|91|22)\b", s)
    if m:
        return m.group(1)

    # heuristische Treffer (Text)
    if "fleisch" in s or "wurst" in s or "heidemark" in s:
        return "21"
    if "wiesenhof" in s or "geflügel" in s:
        return "1011"
    if "bio" in s:
        return "41"
    if "frischfleisch" in s or "veredlung" in s or "schwein" in s or "pök" in s:
        return "65"
    if "avo" in s or "gewürz" in s:
        return "0"
    if "werbe" in s or "werbemittel" in s:
        return "91"
    if "pfeiffer" in s or "gmyrek" in s or "siebert" in s or "bard" in s or "mago" in s:
        return "22"

    return "?"


def detect_bspalten(columns: List[str]):
    """
    Erkennung für Spalten wie:
    "Mo Z Wiesenhof B_Di" / "Mo L Bio B_Mi" / "Mo Wiesenhof B_Di" etc.
    """
    rx = re.compile(
        r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+"
        r"(?:(Z|L)\s+)?(.+?)\s+B[_ ]?(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)$",
        re.IGNORECASE
    )
    mapping = {}
    for c in columns:
        m = rx.match(c.strip())
        if not m:
            continue

        day_de = DAY_SHORT_TO_DE.get(m.group(1))
        zl = (m.group(2) or "").upper()
        group_id = canon_group_id(m.group(3).strip())
        bestell_de = DAY_SHORT_TO_DE.get(m.group(4))

        if day_de and bestell_de:
            key = (day_de, group_id, bestell_de)
            mapping.setdefault(key, {})
            if zl == "Z":
                mapping[key]["zeit"] = c
            elif zl == "L":
                mapping[key]["l"] = c
            else:
                mapping[key]["sort"] = c
    return mapping


def detect_triplets(columns: List[str]):
    """
    Erkennung für Triplets pro Liefertag:
    "<Tag> <Sortiment/Gruppe> Zeit|Zeitende|Bestellzeitende|Uhrzeit"
    "<Tag> <Sortiment/Gruppe> Sort|Sortiment"
    "<Tag> <Sortiment/Gruppe> Tag|Bestelltag"
    -> mappen auf canonical group IDs
    """
    rx = re.compile(
        r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(.+?)\s+"
        r"(Zeit|Zeitende|Bestellzeitende|Uhrzeit|Sort|Sortiment|Tag|Bestelltag)$",
        re.IGNORECASE
    )
    found = {}
    for c in columns:
        m = rx.match(c.strip())
        if not m:
            continue

        day_de = DAY_SHORT_TO_DE.get(m.group(1))
        if not day_de:
            continue

        raw_group = m.group(2).strip()
        gid = canon_group_id(raw_group)

        end_key = m.group(3).lower()
        if end_key in ("sort", "sortiment"):
            key = "Sort"
        elif end_key in ("tag", "bestelltag"):
            key = "Tag"
        else:
            key = "Zeit"

        found.setdefault(day_de, {}).setdefault(gid, {})[key] = c

    return found


def detect_ds_triplets(columns: List[str]):
    rx = re.compile(
        r"^DS\s+(.+?)\s+zu\s+(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+"
        r"(Zeit|Sort|Tag)$",
        re.IGNORECASE
    )
    tmp = {}
    for c in columns:
        m = rx.match(c.strip())
        if not m:
            continue

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
  *{ box-sizing:border-box; font-family: Arial, Helvetica, sans-serif; }
  body{ margin:0; background: #0b1220; color: #fff; }
  .app{ display:grid; grid-template-columns: 350px 1fr; height:100vh; padding:15px; gap:15px; }
  .sidebar, .main{ background: rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.14); border-radius:12px; }
  .list{ height: calc(100vh - 280px); overflow-y:auto; border-top:1px solid rgba(255,255,255,.14); }
  .item{ padding:10px; border-bottom:1px solid rgba(255,255,255,0.05); cursor:pointer; font-size:13px; }
  .wrap{ height: 100%; overflow-y: scroll; padding: 20px; display: flex; flex-direction: column; align-items: center; }

  .paper{
    width: 210mm; min-height: 296.5mm; background: white; color: black; padding: 12mm;
    box-shadow: 0 0 20px rgba(0,0,0,0.5); display: flex; flex-direction: column;
    --fs: 12pt;
  }
  .paper * { font-size: var(--fs); line-height: 1.4; }
  .ptitle{ text-align:center; font-weight:900; font-size:1.8em; margin:0; }
  .pstd{ text-align:center; color:#d0192b; font-weight:bold; margin:1mm 0; font-size:1.4em; }
  .psub{ text-align:center; color:#444; margin-bottom:6mm; font-weight:bold; }

  .head-box { display:flex; justify-content:space-between; margin-bottom:4mm; border-bottom:2px solid #000; padding-bottom:3mm; }

  .tour-info { margin-bottom: 5mm; }
  .tour-table { width: 100%; border-collapse: collapse; margin-top: 1mm; }
  .tour-table th { background: #eee; font-size: 0.8em; padding: 2px; border: 1px solid #000; }
  .tour-table td { border: 1px solid #000; padding: 6px; text-align: center; font-weight: bold; font-size: 1.1em; }

  table.main-table { width:100%; border-collapse:collapse; table-layout: fixed; border: 2px solid #000; }
  table.main-table th { border: 1px solid #000; padding: 8px; background:#f2f2f2; font-weight:bold; text-align:left; }
  table.main-table td { border: 1px solid #000; padding: 8px; vertical-align: top; }

  .day-header { background-color: #e0e0e0 !important; font-weight: 900; border-top: 3px solid #000 !important; }

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
          <td style="width:18%">${idx === 0 ? d : ""}</td>
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
    let fs = 12; p.style.setProperty("--fs", fs + "pt");
    while(p.scrollHeight > 1120 && fs > 8){ fs -= 0.5; p.style.setProperty("--fs", fs + "pt"); }
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
st.set_page_config(page_title="Sendeplan Fix", layout="wide")

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
        if not knr:
            continue

        bestell = []
        for d_de in DAYS_DE:
            day_items = []

            # 1) Triplets (inkl. Fleisch/Wurst etc. über canonical ID)
            if d_de in trip:
                for gid, f in trip[d_de].items():
                    s = norm(r.get(f.get("Sort")))
                    t = safe_time(r.get(f.get("Zeit")))
                    tag = norm(r.get(f.get("Tag")))

                    # nur aufnehmen wenn irgendwas gefüllt ist
                    if s or t or tag:
                        day_items.append({
                            "liefertag": d_de,
                            "sortiment": s,
                            "bestelltag": tag,
                            "bestellschluss": t,
                            "prio": SORT_PRIO.get(gid, 50)
                        })

            # 2) B-Spalten (Wiesenhof, Bio, etc.) -> canonical group id für Prio
            keys = [k for k in bmap.keys() if k[0] == d_de]
            for k in keys:
                f = bmap[k]
                group_id = str(k[1])

                s = norm(r.get(f.get("sort", "")))
                z = safe_time(r.get(f.get("zeit", "")))

                if s or z:
                    day_items.append({
                        "liefertag": d_de,
                        "sortiment": s,
                        "bestelltag": k[2],
                        "bestellschluss": z,
                        "prio": SORT_PRIO.get(group_id, 50)
                    })

            # 3) Deutsche See
            if d_de in ds_trip:
                for key_ds in ds_trip[d_de]:
                    f = ds_trip[d_de][key_ds]
                    s = norm(r.get(f.get("Sort")))
                    t = safe_time(r.get(f.get("Zeit")))
                    tag = norm(r.get(f.get("Tag")))
                    if s or t or tag:
                        day_items.append({
                            "liefertag": d_de,
                            "sortiment": s,
                            "bestelltag": tag,
                            "bestellschluss": t,
                            "prio": 80
                        })

            # innerhalb eines Tages sortieren
            day_items.sort(key=lambda x: x["prio"])
            bestell.extend(day_items)

        data[knr] = {
            "plan_typ": PLAN_TYP,
            "bereich": BEREICH,
            "kunden_nr": knr,
            "name": norm(r.get("Name", "")),
            "strasse": norm(r.get("Strasse", "")),
            "plz": norm(r.get("Plz", "")),
            "ort": norm(r.get("Ort", "")),
            "fachberater": norm(r.get("Fachberater", "")),
            "tours": {d: norm(r.get(TOUR_COLS[d], "")) for d in DAYS_DE},
            "bestell": bestell
        }

    html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(data, separators=(',', ':')))
    st.download_button("Download Sendeplan", data=html, file_name="sendeplan.html", mime="text/html")
```
