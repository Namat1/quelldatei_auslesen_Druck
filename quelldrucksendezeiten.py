# quelldrucksendezeiten_fixed.py
# -----------------------------------------------------------------------------
# VERSION: FIXED - Korrekte Sortiment-Zuordnung basierend auf tatsächlichem Namen
# -----------------------------------------------------------------------------
# Änderungen:
# - Sortimente werden nach ihrem TATSÄCHLICHEN Namen klassifiziert, nicht nach Spaltennummer
# - Bestelltag wird aus L-Spalte gelesen (falls vorhanden), nicht aus Spaltennamen
# - Robustere Handhabung von fehlplatzierten Sortimenten
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
# Basierend auf dem RICHTIGEN PDF (nicht dem falschen!):
# 1. Fleisch- & Wurst (21)
# 2. Geflügel Wiesenhof (1011)
# 3. Bio-Geflügel (41)
# 4. Frischfleisch Veredlung (65)
# 5. Avo-Gewürze (0)
# 6. Avo-Gewürze (0)
# 7. Werbemittel (91)
# WICHTIG: Pfeiffer (22) kommt nach Wiesenhof (1011), vor Bio (41) - siehe Donnerstag im richtigen PDF!
SORT_PRIO = {"21": 0, "1011": 1, "22": 2, "41": 3, "65": 4, "0": 5, "91": 6}

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
    Mapped Sortimentsbezeichnungen robust auf interne IDs.
    WICHTIG: Spezifischere Regeln MÜSSEN vor allgemeineren kommen!
    """
    s = norm(label).lower()

    # harte Treffer (Zahlen)
    m = re.search(r"\b(1011|21|41|65|0|91|22)\b", s)
    if m:
        return m.group(1)

    # heuristische Treffer (Text) - REIHENFOLGE IST KRITISCH!
    # Spezifische Begriffe ZUERST prüfen:
    
    # Bio-Geflügel (41) - sehr spezifisch
    if "bio" in s and "geflügel" in s:
        return "41"
    
    # Wiesenhof/Geflügel (1011) - vor allgemeinem "fleisch"
    if "wiesenhof" in s:
        return "1011"
    if "geflügel" in s:  # nur wenn nicht schon als Bio-Geflügel erkannt
        return "1011"
    
    # Frischfleisch (65) - MUSS vor allgemeinem "fleisch" kommen!
    if "frischfleisch" in s or "veredlung" in s or "schwein" in s or "pök" in s:
        return "65"
    
    # Fleisch/Wurst (21) - allgemeiner Begriff, kommt NACH Frischfleisch
    if "fleisch" in s or "wurst" in s or "heidemark" in s:
        return "21"
    
    # Avo-Gewürze (0)
    if "avo" in s or "gewürz" in s:
        return "0"
    
    # Werbemittel (91)
    if "werbe" in s or "werbemittel" in s:
        return "91"
    
    # Pfeiffer etc. (22)
    if "pfeiffer" in s or "gmyrek" in s or "siebert" in s or "bard" in s or "mago" in s:
        return "22"

    return "?"


def detect_bspalten(columns: List[str]):
    """
    Erkennung für Spalten wie:
    "Mo Z Wiesenhof B_Di" / "Mo L Bio B_Mi" / "Mo Wiesenhof B_Di" etc.
    UND auch Spalten OHNE "B": "Mit Z 41 Mo" (nur Tag ZL Gruppe Tag)
    WICHTIG: Erlaubt optionales Leerzeichen nach B_ (z.B. "Donn 22 B_ Mo")
    
    WICHTIG: Wir extrahieren die Gruppe aus dem Spaltennamen, aber klassifizieren
    später das Sortiment anhand seines tatsächlichen Namens neu!
    """
    # Pattern MIT "B" - erlaubt optionales Leerzeichen nach B_
    rx_b = re.compile(
        r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+"
        r"(?:(Z|L)\s+)?(.+?)\s+B[_ ]\s*(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)$",
        re.IGNORECASE
    )
    # Pattern OHNE "B" (nur für Z/L Spalten!)
    rx_no_b = re.compile(
        r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+"
        r"(Z|L)\s+(.+?)\s+(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)$",
        re.IGNORECASE
    )
    
    mapping = {}
    
    # ERSTE PHASE: Verarbeite Spalten OHNE "B" (diese haben Priorität für Z/L!)
    for c in columns:
        m = rx_no_b.match(c.strip())
        if m:
            day_de = DAY_SHORT_TO_DE.get(m.group(1))
            zl = m.group(2).upper()  # Z oder L ist required hier
            group_text = m.group(3).strip()
            bestell_de_from_name = DAY_SHORT_TO_DE.get(m.group(4))

            if day_de and bestell_de_from_name:
                key = (day_de, group_text, bestell_de_from_name)
                mapping.setdefault(key, {})
                if zl == "Z":
                    mapping[key]["zeit"] = c
                elif zl == "L":
                    mapping[key]["l"] = c
    
    # ZWEITE PHASE: Verarbeite Spalten MIT "B" (überschreiben NICHT existierende Z/L!)
    for c in columns:
        m = rx_b.match(c.strip())
        if m:
            day_de = DAY_SHORT_TO_DE.get(m.group(1))
            zl = (m.group(2) or "").upper()
            group_text = m.group(3).strip()
            bestell_de_from_name = DAY_SHORT_TO_DE.get(m.group(4))

            if day_de and bestell_de_from_name:
                key = (day_de, group_text, bestell_de_from_name)
                mapping.setdefault(key, {})
                if zl == "Z":
                    # NUR setzen wenn noch keine Zeit-Spalte vorhanden!
                    if "zeit" not in mapping[key]:
                        mapping[key]["zeit"] = c
                elif zl == "L":
                    # NUR setzen wenn noch keine L-Spalte vorhanden!
                    if "l" not in mapping[key]:
                        mapping[key]["l"] = c
                else:
                    mapping[key]["sort"] = c
                    mapping[key]["group_text"] = group_text
    
    return mapping


def detect_triplets(columns: List[str]):
    """
    Robustere Triplet-Erkennung:
    "<Tag> <Gruppe> Zeit|Zeitende|Bestellzeitende|Uhrzeit"
    "<Tag> <Gruppe> Sort|Sortiment"
    "<Tag> <Gruppe> Tag|Bestelltag"
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
        # Speichere sowohl den rohen Text als auch die ID
        group_text = raw_group

        end_key = m.group(3).lower()
        if end_key in ("sort", "sortiment"):
            key = "Sort"
        elif end_key in ("tag", "bestelltag"):
            key = "Tag"
        else:
            key = "Zeit"

        found.setdefault(day_de, {}).setdefault(group_text, {})[key] = c

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


# --- HTML TEMPLATE (A4 MIT SCROLLBALKEN - PRINT OPTIMIERT - 4 BEREICHE) ---
HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<style>
  /* Druckseite - WICHTIG: Korrekte A4-Einstellungen */
  @page { 
    size: A4 portrait; 
    margin: 10mm 8mm;
  }

  *{ box-sizing:border-box; font-family: Arial, Helvetica, sans-serif; }
  body{ margin:0; background:#0b1220; color:#fff; }

  /* === BILDSCHIRM-ANSICHT === */
  @media screen {
    .app{ display:grid; grid-template-columns: 350px 1fr; height:100vh; padding:15px; gap:15px; }
    .sidebar, .main{ background: rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.14); border-radius:12px; }
    .list{ height: calc(100vh - 380px); overflow-y:auto; border-top:1px solid rgba(255,255,255,.14); margin-top:10px; }
    .item{ padding:10px; border-bottom:1px solid rgba(255,255,255,0.05); cursor:pointer; font-size:13px; }
    .wrap{ height: 100%; overflow-y: auto; padding: 20px; display: flex; flex-direction: column; align-items: center; }

    /* Papier Container - Fixed A4 Größe für Bildschirm */
    .paper{
      width: 210mm;
      max-width: 210mm;
      height: 297mm;
      background:#fff;
      color:#000;
      box-shadow: 0 0 20px rgba(0,0,0,.5);
      position: relative;
      overflow: hidden;
      margin-bottom: 20px;
    }

    /* Scrollbarer Inhalt */
    .paper-content{
      width: 100%;
      height: 100%;
      padding: 10mm 8mm;
      overflow-y: auto;
      overflow-x: hidden;
    }

    /* Scrollbalken-Styling */
    .paper-content::-webkit-scrollbar { width: 10px; }
    .paper-content::-webkit-scrollbar-track { background: #f1f1f1; }
    .paper-content::-webkit-scrollbar-thumb { background: #888; border-radius: 5px; }
    .paper-content::-webkit-scrollbar-thumb:hover { background: #555; }
  }

  /* === DRUCK-ANSICHT === */
  @media print {
    body{ background:#fff !important; margin: 0; padding: 0; }
    
    .sidebar{ display:none !important; }
    .app{ display:block; padding:0; margin: 0; }
    .wrap{ overflow: visible; padding: 0; margin: 0; }
    
    .paper{
      box-shadow: none;
      margin: 0;
      padding: 0;
      width: 100%;
      max-width: 100%;
      height: auto;
      overflow: visible;
      page-break-after: always;
      background: #fff;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
      color-adjust: exact;
    }
    
    .paper-content{
      overflow: visible;
      height: auto;
      padding: 0;
      margin: 0;
    }
    
    /* WICHTIG: Erzwinge schwarze Farben beim Drucken */
    .paper-content, .paper-content * {
      color: #000 !important;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
      color-adjust: exact !important;
    }
    
    /* Tabellen-Borders schwarz */
    table, th, td {
      border-color: #000 !important;
    }
    
    /* Hintergrundfarben beibehalten */
    .day-header {
      background: #e0e0e0 !important;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    .tour-table th {
      background: #eee !important;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    .main-table th {
      background: #f2f2f2 !important;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    
    /* Roter Titel bleibt rot */
    .pstd {
      color: #d0192b !important;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
  }

  /* === GEMEINSAME STYLES === */
  .paper-content *{ font-size: 9pt; line-height: 1.15; }

  .ptitle{ text-align:center; font-weight:900; font-size:1.5em; margin:0 0 1mm 0; }
  .pstd{ text-align:center; color:#d0192b; font-weight:800; margin:0.5mm 0; font-size:1.15em; }
  .psub{ text-align:center; color:#333; margin: 0 0 2mm 0; font-weight:700; font-size:0.95em; }

  .head-box{
    display:flex;
    justify-content:space-between;
    gap:5mm;
    margin-bottom:2mm;
    border-bottom:1.5px solid #000;
    padding-bottom:1.5mm;
    font-size:8.5pt;
  }

  .tour-info { margin-bottom:2mm; }

  .tour-table { width:100%; border-collapse:collapse; table-layout:fixed; }
  .tour-table th { background:#eee; font-size:0.8em; padding:2px 1px; border:1px solid #000; font-weight:700; }
  .tour-table td { border:1px solid #000; padding:3px 1px; text-align:center; font-weight:800; font-size:0.9em; }

  table.main-table { width:100%; border-collapse:collapse; table-layout:fixed; border:2px solid #000; margin-top:2mm; }
  table.main-table th { border:1px solid #000; padding:3px 4px; background:#f2f2f2; font-weight:800; text-align:left; font-size:0.85em; }
  table.main-table td { border:1px solid #000; padding:3px 4px; vertical-align:top; word-wrap:break-word; overflow-wrap:anywhere; font-size:0.88em; }

  .day-header { background:#e0e0e0 !important; font-weight:900; border-top:2px solid #000 !important; }
  
  /* Bereichs-Buttons */
  .area-buttons {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 8px;
    padding: 15px;
    border-bottom: 1px solid rgba(255,255,255,.14);
  }
  
  .area-btn {
    padding: 12px;
    border: 2px solid rgba(255,255,255,.2);
    background: rgba(255,255,255,.05);
    color: #fff;
    cursor: pointer;
    border-radius: 8px;
    font-weight: 600;
    transition: all 0.2s;
    text-align: center;
  }
  
  .area-btn:hover {
    background: rgba(255,255,255,.12);
    border-color: rgba(255,255,255,.4);
  }
  
  .area-btn.active {
    background: #4fa3ff;
    border-color: #4fa3ff;
  }
  
  /* Verhindere Seitenumbrüche innerhalb von Tabellenzeilen */
  @media print {
    tr { page-break-inside: avoid; }
  }
</style>
</head>
<body>
<div class="app">
  <div class="sidebar">
    <div style="padding:15px; font-weight:bold; font-size:16px;">Sendeplan Generator</div>
    
    <!-- BEREICHS-BUTTONS -->
    <div class="area-buttons">
      <div class="area-btn active" id="btn-direkt" onclick="switchArea('direkt')">Direkt</div>
      <div class="area-btn" id="btn-mk" onclick="switchArea('mk')">MK</div>
      <div class="area-btn" id="btn-nms" onclick="switchArea('nms')">HuPa NMS</div>
      <div class="area-btn" id="btn-malchow" onclick="switchArea('malchow')">HuPa Malchow</div>
    </div>
    
    <div style="padding:15px; display:flex; flex-direction:column; gap:10px;">
      <input id="knr" placeholder="Kunden-Nr..." style="width:100%; padding:10px; border-radius:5px;">
      <button onclick="showOne()" style="padding:10px; background:#4fa3ff; color:white; border:none; cursor:pointer; border-radius:5px;">Anzeigen</button>
      <button onclick="window.print()" style="padding:10px; background:#28a745; color:white; border:none; cursor:pointer; border-radius:5px;">Drucken</button>
      <button onclick="printAll()" style="padding:10px; background:#ff6b35; color:white; border:none; cursor:pointer; font-weight:bold; border-radius:5px;">Alle drucken</button>
    </div>
    <div class="list" id="list"></div>
  </div>

  <div class="main">
    <div class="wrap" id="out">Bitte Bereich und Kunden wählen...</div>
  </div>
</div>

<script>
const ALL_DATA = __DATA_JSON__;
let currentArea = 'direkt';
let DATA = ALL_DATA['direkt'] || {};
let ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0)-(Number(b)||0));
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

  const tourRow  = DAYS.map(d => `<td>${esc(c.tours[d] || "—")}</td>`).join("");
  const tourHead = DAYS.map(d => `<th>${d.substring(0,2)}</th>`).join("");

  return `<div class="paper">
    <div class="paper-content">
      <div class="ptitle">Sende- &amp; Belieferungsplan</div>
      <div class="pstd">${esc(c.plan_typ)}</div>
      <div class="psub">${esc(c.name)} | ${esc(c.bereich)}</div>

      <div class="head-box">
        <div><b>${esc(c.name)}</b><br>${esc(c.strasse)}<br>${esc(c.plz)} ${esc(c.ort)}</div>
        <div style="text-align:right">Kunden-Nr: <b>${esc(c.kunden_nr)}</b><br>Fachberater: <b>${esc(c.fachberater)}</b></div>
      </div>

      <div class="tour-info">
        <table class="tour-table">
          <thead><tr>${tourHead}</tr></thead>
          <tbody><tr>${tourRow}</tr></tbody>
        </table>
      </div>

      <table class="main-table">
        <thead><tr><th>Liefertag</th><th>Sortiment</th><th>Bestelltag</th><th>Bestellzeitende</th></tr></thead>
        <tbody>${tableRows}</tbody>
      </table>
    </div>
  </div>`;
}

function showOne(){
  const k = document.getElementById("knr").value.trim();
  if(DATA[k]) {
    document.getElementById("out").innerHTML = render(DATA[k]);
  }
}

function switchArea(area){
  currentArea = area;
  DATA = ALL_DATA[area] || {};
  ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0)-(Number(b)||0));
  
  // Update Button-Styles
  document.querySelectorAll('.area-btn').forEach(btn => btn.classList.remove('active'));
  document.getElementById(`btn-${area}`).classList.add('active');
  
  // Update Liste
  updateList();
  
  // Clear input und output
  document.getElementById("knr").value = "";
  document.getElementById("out").innerHTML = `Bereich gewechselt zu: ${getAreaName(area)}<br>Bitte Kunden wählen...`;
}

function getAreaName(area){
  const names = {
    'direkt': 'Direkt',
    'mk': 'MK',
    'nms': 'HuPa NMS',
    'malchow': 'HuPa Malchow'
  };
  return names[area] || area;
}

function updateList(){
  document.getElementById("list").innerHTML = ORDER.map(k => {
    const name = (DATA[k] && DATA[k].name) ? DATA[k].name : "";
    return `<div class="item" onclick="document.getElementById('knr').value='${k}';showOne()"><b>${k}</b> - ${esc(name)}</div>`;
  }).join("");
}

function printAll(){
  // Bestätigungsdialog
  const count = ORDER.length;
  const areaName = getAreaName(currentArea);
  if(!confirm(`Möchten Sie wirklich alle ${count} Kunden aus "${areaName}" drucken?`)) {
    return;
  }
  
  // Rendere alle Kunden
  let html = "";
  ORDER.forEach(k => {
    if(DATA[k]) {
      html += render(DATA[k]);
    }
  });
  
  // Zeige alle Kunden im Haupt-Bereich
  document.getElementById("out").innerHTML = html;
  
  // Warte kurz damit das Rendering fertig ist, dann drucke
  setTimeout(() => {
    window.print();
  }, 500);
}

// Initialisiere Liste
updateList();
</script>
</body>
</html>
"""

# --- STREAMLIT APP ---
st.set_page_config(page_title="Sendeplan Generator - 4 Bereiche", layout="wide")

st.title("Sendeplan Generator")
st.write("Verarbeitet 4 Bereiche: Direkt, MK, HuPa NMS, HuPa Malchow")

up = st.file_uploader("Excel Datei laden", type=["xlsx"])
if up:
    # Definiere die 4 Blätter und ihre internen Namen
    SHEETS = {
        'direkt': 'Direkt 1 - 99',
        'mk': 'Hupa MK 882',
        'nms': 'Hupa 2221-4444',
        'malchow': 'Hupa 7773-7779'
    }
    
    all_data = {}
    
    # Verarbeite jedes Blatt
    for area_key, sheet_name in SHEETS.items():
        st.write(f"Verarbeite: **{sheet_name}**...")
        
        try:
            df = pd.read_excel(up, sheet_name=sheet_name)
        except Exception as e:
            st.error(f"Fehler beim Laden von '{sheet_name}': {e}")
            continue
            
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

                # 1) Triplets - WICHTIG: Klassifiziere nach tatsächlichem Sortiment-Namen!
                if d_de in trip:
                    for group_text, f in trip[d_de].items():
                        s = norm(r.get(f.get("Sort")))
                        t = safe_time(r.get(f.get("Zeit")))
                        tag = norm(r.get(f.get("Tag")))

                        if s or t or tag:
                            # Klassifiziere das Sortiment nach seinem NAMEN, nicht nach Spaltennummer!
                            actual_gid = canon_group_id(s)
                            
                            day_items.append({
                                "liefertag": d_de,
                                "sortiment": s,
                                "bestelltag": tag,
                                "bestellschluss": t,
                                "prio": SORT_PRIO.get(actual_gid, 50)
                            })

                # 2) B-Spalten - WICHTIG: Verwende L-Spalte für Bestelltag (wenn vorhanden)!
                keys = [k for k in bmap.keys() if k[0] == d_de]
                for k in keys:
                    f = bmap[k]
                    
                    s = norm(r.get(f.get("sort", "")))
                    z = safe_time(r.get(f.get("zeit", "")))
                    
                    # Verwende L-Spalte für Bestelltag (wenn vorhanden), sonst Spaltennamen
                    l_col = f.get("l")
                    if l_col:
                        tag = norm(r.get(l_col, ""))
                        if not tag:
                            tag = k[2]  # Fallback auf Spaltennamen wenn L-Spalte leer
                    else:
                        tag = k[2]  # Kein L-Spalte -> verwende Spaltennamen

                    if s or z:
                        # Klassifiziere das Sortiment nach seinem NAMEN!
                        actual_gid = canon_group_id(s)
                        
                        day_items.append({
                            "liefertag": d_de,
                            "sortiment": s,
                            "bestelltag": tag,
                            "bestellschluss": z,
                            "prio": SORT_PRIO.get(actual_gid, 50)
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
                                "prio": 5.5  # Nach Avo-Gewürze (5), vor Werbemittel (6)
                            })

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
        
        # Speichere Daten für diesen Bereich
        all_data[area_key] = data
        st.success(f"✓ {sheet_name}: {len(data)} Kunden verarbeitet")
    
    # Generiere HTML mit allen 4 Bereichen
    html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(all_data, ensure_ascii=False, separators=(',', ':')))
    
    st.write("---")
    st.write(f"**Gesamt:** {sum(len(all_data[k]) for k in all_data)} Kunden in {len(all_data)} Bereichen")
    st.download_button("Download Sendeplan (A4)", data=html, file_name="sendeplan_4_bereiche.html", mime="text/html")
