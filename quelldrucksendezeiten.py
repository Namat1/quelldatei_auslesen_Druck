# quelldrucksendezeiten_fixed.py
# -----------------------------------------------------------------------------
# VERSION: FIXED - Korrekte Sortiment-Zuordnung basierend auf tatsächlichem Namen
# + LOGO Upload in Streamlit + Logo im Print oben (Base64 eingebettet)
# + FIX: JS-Syntaxfehler (überflüssige "}" in render()) entfernt
# + CHANGE: In den Tageskarten werden JE Sortiment die einzelnen Bestellzeiten
#          (Bestelltag + Bestellschluss) ausgegeben (keine Zusammenfassung per Set)
# -----------------------------------------------------------------------------

import json
import re
import datetime
import base64
from pathlib import Path
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

    # Bio-Geflügel (41)
    if "bio" in s and "geflügel" in s:
        return "41"

    # Wiesenhof/Geflügel (1011)
    if "wiesenhof" in s:
        return "1011"
    if "geflügel" in s:
        return "1011"

    # Frischfleisch (65)
    if "frischfleisch" in s or "veredlung" in s or "schwein" in s or "pök" in s:
        return "65"

    # Fleisch/Wurst (21)
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
    """
    rx_b = re.compile(
        r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+"
        r"(?:(Z|L)\s+)?(.+?)\s+B[_ ]\s*(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)$",
        re.IGNORECASE
    )
    rx_no_b = re.compile(
        r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+"
        r"(Z|L)\s+(.+?)\s+(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)$",
        re.IGNORECASE
    )

    mapping = {}

    # Phase 1: ohne B
    for c in columns:
        if re.search(r"\sB[_ ]\s*", c, re.IGNORECASE):
            continue

        m = rx_no_b.match(c.strip())
        if m:
            day_de = DAY_SHORT_TO_DE.get(m.group(1))
            zl = m.group(2).upper()
            group_text = m.group(3).strip()
            bestell_de_from_name = DAY_SHORT_TO_DE.get(m.group(4))

            if day_de and bestell_de_from_name:
                key = (day_de, group_text, bestell_de_from_name)
                mapping.setdefault(key, {})
                if zl == "Z":
                    mapping[key]["zeit"] = c
                elif zl == "L":
                    mapping[key]["l"] = c

    # Phase 2: mit B
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
                    if "zeit" not in mapping[key]:
                        mapping[key]["zeit"] = c
                elif zl == "L":
                    if "l" not in mapping[key]:
                        mapping[key]["l"] = c
                else:
                    mapping[key]["sort"] = c
                    mapping[key]["group_text"] = group_text

    return mapping


def detect_triplets(columns: List[str]):
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

        group_text = m.group(2).strip()

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
        r"^DS\s+(.+?)\s+zu\s+(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(Zeit|Sort|Tag)$",
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


def load_logo_data_uri() -> str:
    """
    Lädt Logo von Festplatte (Fallback), z.B. neben dem Script oder /mnt/data.
    """
    candidates = []
    try:
        here = Path(__file__).resolve().parent
        candidates.append(here / "Logo_NORDfrische Center (NFC).png")
    except Exception:
        pass
    candidates.append(Path.cwd() / "Logo_NORDfrische Center (NFC).png")
    candidates.append(Path("/mnt/data/Logo_NORDfrische Center (NFC).png"))

    for p in candidates:
        try:
            if p.exists() and p.is_file():
                b = p.read_bytes()
                return "data:image/png;base64," + base64.b64encode(b).decode("ascii")
        except Exception:
            continue
    return ""


def logo_file_to_data_uri(uploaded_file) -> str:
    """
    Wandelt ein hochgeladenes Streamlit-File (PNG/JPG/SVG) in eine Data-URI um.
    """
    if not uploaded_file:
        return ""
    # robustere MIME-Erkennung (z.B. SVG)
    mime = uploaded_file.type or ("image/svg+xml" if str(uploaded_file.name).lower().endswith(".svg") else "image/png")
    b = uploaded_file.getvalue()
    return f"data:{mime};base64," + base64.b64encode(b).decode("ascii")


# --- HTML TEMPLATE (A4 MIT SCROLLBALKEN - PRINT OPTIMIERT - 4 BEREICHE) ---
HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<style>
  @page { 
    size: A4 portrait; 
    margin: 12mm 10mm;
  }

  *{ box-sizing:border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
  body{ margin:0; background:#1e1e1e; color:#e8eaed; }

  @media screen {
    .app{ display:grid; grid-template-columns: 350px 1fr; height:100vh; padding:15px; gap:15px; }
    .sidebar, .main{ background: #2d2d2d; border:1px solid #3c3c3c; border-radius:12px; box-shadow: 0 2px 8px rgba(0,0,0,0.3); }
    .list{ height: calc(100vh - 380px); overflow-y:auto; border-top:1px solid #3c3c3c; margin-top:10px; }
    .item{ padding:10px; border-bottom:1px solid #3c3c3c; cursor:pointer; font-size:13px; color:#b8b8b8; transition: background 0.2s; }
    .wrap{ height: 100%; overflow-y: auto; padding: 20px; display: flex; flex-direction: column; align-items: center; }

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

    .paper-content{
      width: 100%;
      height: 100%;
      padding: 12mm 10mm;
      overflow-y: auto;
      overflow-x: hidden;
    }

    .paper-content::-webkit-scrollbar { width: 10px; }
    .paper-content::-webkit-scrollbar-track { background: #f1f1f1; }
    .paper-content::-webkit-scrollbar-thumb { background: #888; border-radius: 5px; }
    .paper-content::-webkit-scrollbar-thumb:hover { background: #555; }
    
    .list::-webkit-scrollbar { width: 8px; }
    .list::-webkit-scrollbar-track { background: #252525; }
    .list::-webkit-scrollbar-thumb { background: #4a4a4a; border-radius: 4px; }
    .list::-webkit-scrollbar-thumb:hover { background: #1a73e8; }
  }

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

    .paper-content, .paper-content * {
      color: #000 !important;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
      color-adjust: exact !important;
    }

    img { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
    table, th, td { border-color: transparent !important; }
    
    .paper-content * { font-size: 8.5pt !important; line-height: 1.15 !important; }
    
    .header-section { display: flex !important; justify-content: space-between !important; }
    .logo { height: 14mm !important; }
    .customer-box { background: #f8f9fa !important; border: 1px solid #dadce0 !important; padding: 2mm 3mm !important; }
    .address-box { background: #ffffff !important; border: 1px solid #dadce0 !important; padding: 2mm 3mm !important; }
    
    .main-title { color: #1e3a5f !important; font-size: 1.8em !important; }
    .plan-type { color: #f39c12 !important; font-size: 1.2em !important; }
    
    .tour-table th { background: #1e3a5f !important; color: white !important; padding: 1.5mm 0.5mm !important; }
    .tour-table td { border-right: 1px solid #dadce0 !important; padding: 1.5mm 0.5mm !important; }
    
    .days-grid { gap: 2.5mm !important; }
    
    .day-card { 
      box-shadow: 0 1px 2px rgba(0,0,0,0.1) !important;
      page-break-inside: avoid !important;
      border: 1px solid #e0e0e0 !important;
    }
    
    .day-card-header { padding: 2mm 2.5mm !important; }
    
    .day-card.active .day-card-header { 
      background: #1e73e8 !important; 
      color: white !important;
    }
    
    .day-card.inactive .day-card-header { 
      background: #9aa0a6 !important; 
      color: white !important;
    }
    
    .day-card-body { padding: 2.5mm !important; }
    
    .sortiment-list li {
      font-size: 0.9em !important;
      padding: 1mm 0 !important;
    }

    .li-sub { font-size: 0.82em !important; }
    
    .sortiment-list li:before {
      color: #1e73e8 !important;
    }
    
    .card-info { font-size: 0.85em !important; }
  }

  .paper-content *{ font-size: 8.5pt; line-height: 1.15; }

  /* === HEADER SECTION === */
  .header-section {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 2mm;
    padding-bottom: 1.5mm;
  }

  .header-left {
    display: flex;
    flex-direction: column;
    align-items: flex-start;
  }

  .logo {
    height: 14mm;
    margin-bottom: 0.5mm;
  }

  .logo-subtitle {
    font-size: 0.85em;
    color: #5f6368;
    font-weight: 500;
  }

  .header-right {
    text-align: right;
  }

  .customer-box {
    background: #f8f9fa;
    border: 1px solid #dadce0;
    border-radius: 4px;
    padding: 2mm 3mm;
    font-size: 0.9em;
    line-height: 1.5;
  }

  .customer-box strong {
    font-weight: 600;
    color: #2c3e50;
  }

  /* === TITLE SECTION === */
  .title-section {
    text-align: center;
    margin-bottom: 2mm;
  }

  .main-title {
    font-size: 1.8em;
    font-weight: 900;
    color: #1e3a5f;
    margin: 0 0 0.5mm 0;
  }

  .plan-type {
    font-size: 1.2em;
    color: #f39c12;
    font-weight: 800;
    margin: 0.3mm 0;
  }

  .customer-subtitle {
    font-size: 1em;
    color: #2c3e50;
    font-weight: 600;
    margin-top: 0.5mm;
  }

  /* === ADDRESS BOX === */
  .address-box {
    background: #ffffff;
    border: 1px solid #dadce0;
    border-radius: 4px;
    padding: 2mm 3mm;
    margin-bottom: 2mm;
    font-size: 0.9em;
    line-height: 1.4;
    color: #2c3e50;
  }

  /* === TOUR SECTION === */
  .tour-section {
    margin-bottom: 2mm;
  }

  .tour-table {
    width: 100%;
    border-collapse: collapse;
    background: #f8f9fa;
    border-radius: 4px;
    overflow: hidden;
  }

  .tour-table th {
    background: #1e3a5f;
    color: white;
    padding: 1.5mm 0.5mm;
    font-size: 0.7em;
    font-weight: 700;
    text-align: center;
    border-right: 1px solid rgba(255,255,255,0.2);
  }

  .tour-table th:last-child {
    border-right: none;
  }

  .tour-table td {
    padding: 1.5mm 0.5mm;
    text-align: center;
    font-weight: 700;
    font-size: 0.85em;
    border-right: 1px solid #dadce0;
    color: #2c3e50;
  }

  .tour-table td:last-child {
    border-right: none;
  }

  /* === DAYS GRID === */
  .days-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 2.5mm;
    margin-top: 1.5mm;
  }

  .day-card {
    background: white;
    border-radius: 5px;
    overflow: hidden;
    box-shadow: 0 1px 2px rgba(0,0,0,0.1);
    page-break-inside: avoid;
    border: 1px solid #e0e0e0;
  }

  .day-card-header {
    padding: 2mm 2.5mm;
    font-weight: 700;
    font-size: 1.05em;
    color: white;
  }

  .day-card.active .day-card-header {
    background: #1e73e8;
  }

  .day-card.inactive .day-card-header {
    background: #9aa0a6;
  }

  .day-card-body {
    padding: 2.5mm;
    background: white;
    min-height: 15mm;
  }

  .sortiment-list {
    list-style: none;
    padding: 0;
    margin: 0 0 1.5mm 0;
  }

  .sortiment-list li {
    padding: 1mm 0;
    padding-left: 3.5mm;
    position: relative;
    font-size: 0.9em;
    line-height: 1.25;
    color: #2c3e50;
  }

  .sortiment-list li:before {
    content: "•";
    position: absolute;
    left: 0;
    color: #1e73e8;
    font-weight: bold;
  }

  .li-main{
    display:block;
    font-weight:700;
  }

  .li-sub{
    display:block;
    font-size:0.85em;
    color:#5f6368;
    margin-top:0.5mm;
  }

  .card-info {
    padding: 1mm 0;
    font-size: 0.85em;
    color: #2c3e50;
  }

  .info-label {
    font-weight: 600;
    color: #5f6368;
  }

  .info-value {
    font-weight: 700;
    color: #2c3e50;
  }

  .no-delivery {
    text-align: center;
    color: #9aa0a6;
    font-style: italic;
    margin: 1.5mm 0;
    font-size: 0.9em;
  }

  .area-buttons {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 8px;
    padding: 15px;
    border-bottom: 1px solid #3c3c3c;
  }

  .area-btn {
    padding: 12px;
    border: 2px solid #4a4a4a;
    background: #3a3a3a;
    color: #b8b8b8;
    cursor: pointer;
    border-radius: 8px;
    font-weight: 600;
    transition: all 0.2s;
    text-align: center;
  }

  .area-btn:hover {
    background: #444444;
    border-color: #1a73e8;
    color: #8ab4f8;
  }

  .area-btn.active {
    background: #1a73e8;
    border-color: #1a73e8;
    color: #ffffff;
  }

  .item:hover {
    background: #383838;
  }

  @media print {
    tr { page-break-inside: avoid; }
  }
</style>
</head>
<body>
<div class="app">
  <div class="sidebar">
    <div style="padding:15px; font-weight:bold; font-size:18px; color:#e8eaed; border-bottom:2px solid #3c3c3c; background:#353535;">📊 Sendeplan Generator</div>

    <div class="area-buttons">
      <div class="area-btn active" id="btn-direkt" onclick="switchArea('direkt')">Direkt</div>
      <div class="area-btn" id="btn-mk" onclick="switchArea('mk')">MK</div>
      <div class="area-btn" id="btn-nms" onclick="switchArea('nms')">HuPa NMS</div>
      <div class="area-btn" id="btn-malchow" onclick="switchArea('malchow')">HuPa Malchow</div>
    </div>

    <div style="padding:15px; display:flex; flex-direction:column; gap:10px;">
      <input id="knr" placeholder="Kunden-Nr..." oninput="showOne()" style="width:100%; padding:10px; border-radius:6px; border:2px solid #4a4a4a; font-size:14px; color:#e8eaed; background:#3a3a3a;">
      <button onclick="showOne()" style="padding:10px; background:#1a73e8; color:white; border:none; cursor:pointer; border-radius:6px; font-weight:600; transition: background 0.2s;" onmouseover="this.style.background='#1557b0'" onmouseout="this.style.background='#1a73e8'">Anzeigen</button>
      <button onclick="window.print()" style="padding:10px; background:#0f9d58; color:white; border:none; cursor:pointer; border-radius:6px; font-weight:600; transition: background 0.2s;" onmouseover="this.style.background='#0d7d47'" onmouseout="this.style.background='#0f9d58'">Drucken</button>
      <button onclick="printAll()" style="padding:10px; background:#ea4335; color:white; border:none; cursor:pointer; font-weight:bold; border-radius:6px; transition: background 0.2s;" onmouseover="this.style.background='#c5221f'" onmouseout="this.style.background='#ea4335'">Alle drucken</button>
    </div>
    <div class="list" id="list"></div>
  </div>

  <div class="main">
    <div class="wrap" id="out"><div style="color:#9aa0a6; padding:20px; font-weight:600; text-align:center;">📋 Bitte Bereich und Kunden wählen...</div></div>
  </div>
</div>

<script>
const ALL_DATA = __DATA_JSON__;
const LOGO_SRC = "__LOGO_DATAURI__";
let currentArea = 'direkt';
let DATA = ALL_DATA['direkt'] || {};
let ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0)-(Number(b)||0));
const DAYS = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];

// Debug: Zeige Daten-Status
console.log("ALL_DATA:", ALL_DATA);
console.log("Anzahl Bereiche:", Object.keys(ALL_DATA || {}).length);
console.log("Direkt Kunden:", Object.keys(DATA).length);

// Zeige Hinweis wenn keine Daten
if (Object.keys(ALL_DATA || {}).length === 0) {
  document.getElementById("out").innerHTML = `
    <div style="color:#ea4335; padding:40px; text-align:center; font-size:16px;">
      <h2>⚠️ Keine Daten gefunden!</h2>
      <p>Diese HTML-Datei enthält keine Kundendaten.</p>
      <p><strong>Bitte führen Sie folgende Schritte aus:</strong></p>
      <ol style="text-align:left; display:inline-block; margin-top:20px;">
        <li>Starten Sie Streamlit: <code>streamlit run quelldrucksendezeiten.py</code></li>
        <li>Laden Sie Ihre Excel-Datei hoch</li>
        <li>Warten Sie, bis die Verarbeitung abgeschlossen ist</li>
        <li>Klicken Sie auf "Download Sendeplan (A4)"</li>
        <li>Öffnen Sie die heruntergeladene HTML-Datei</li>
      </ol>
    </div>
  `;
}

function esc(s){ return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;"); }

function render(c){
  // Erstelle Karten für jeden Tag
  let dayCards = "";
  DAYS.forEach(d => {
    const items = (c.bestell || []).filter(it => it.liefertag === d);
    
    if (items.length > 0) {
      // JE Sortiment als eigene Zeile inkl. Bestelltag + Bestellschluss
      const lines = items.map(it => {
        const s = esc(it.sortiment || "—");
        const bt = esc(it.bestelltag || "—");
        const bz = esc(it.bestellschluss || "—");
        return `<li>
                  <span class="li-main">${s}</span>
                  <span class="li-sub">Bestelltag: ${bt} • Schluss: ${bz}</span>
                </li>`;
      }).join("");

      dayCards += `
        <div class="day-card active">
          <div class="day-card-header">${d}</div>
          <div class="day-card-body">
            <ul class="sortiment-list">${lines}</ul>
          </div>
        </div>`;
    } else {
      // Tag ohne Lieferungen - graue Karte
      dayCards += `
        <div class="day-card inactive">
          <div class="day-card-header">${d}</div>
          <div class="day-card-body">
            <p class="no-delivery">Keine Lieferung</p>
          </div>
        </div>`;
    }
  });

  // Tour-Informationen aufbereiten
  const tourItems = DAYS.map(d => {
    const tourNr = (c.tours && c.tours[d]) ? c.tours[d] : "—";
    return `<td>${esc(tourNr)}</td>`;
  }).join("");
  
  const tourHeaders = DAYS.map(d => `<th>${d.substring(0,2)}</th>`).join("");

  // Logo
  const logoHtml = LOGO_SRC ? `<img class="logo" src="${LOGO_SRC}" alt="Logo">` : "";
  
  // Aktuelles Datum
  const heute = new Date();
  const standDatum = heute.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });

  return `<div class="paper">
    <div class="paper-content">
      <div class="header-section">
        <div class="header-left">
          ${logoHtml}
          <div class="logo-subtitle">Das Fleischwerk von EDEKA Nord</div>
        </div>
        <div class="header-right">
          <div class="customer-box">
            <div><strong>Kunden-Nr:</strong> ${esc(c.kunden_nr)}</div>
            <div><strong>Fachberater:</strong> ${esc(c.fachberater)}</div>
            <div><strong>Stand:</strong> ${standDatum}</div>
          </div>
        </div>
      </div>

      <div class="title-section">
        <h1 class="main-title">Sende- &amp; Belieferungsplan</h1>
        <div class="plan-type">${esc(c.plan_typ)}</div>
        <div class="customer-subtitle">${esc(c.name)} | ${esc(c.bereich)}</div>
      </div>

      <div class="address-box">
        <strong>${esc(c.name)}</strong><br>
        ${esc(c.strasse)}<br>
        ${esc(c.plz)} ${esc(c.ort)}
      </div>

      <div class="tour-section">
        <table class="tour-table">
          <thead><tr>${tourHeaders}</tr></thead>
          <tbody><tr>${tourItems}</tr></tbody>
        </table>
      </div>

      <div class="days-grid">
        ${dayCards}
      </div>
    </div>
  </div>`;
}

function findCustomerInAllAreas(knr){
  for(let area in ALL_DATA){
    if(ALL_DATA[area] && ALL_DATA[area][knr]){
      return area;
    }
  }
  return null;
}

function showOne(){
  const k = document.getElementById("knr").value.trim();
  
  if(!k){
    document.getElementById("out").innerHTML = "<div style='color:#9aa0a6; padding:20px; font-weight:500; text-align:center;'>🔍 Bitte Kundennummer eingeben...</div>";
    return;
  }
  
  if(DATA && DATA[k]){
    document.getElementById("out").innerHTML = render(DATA[k]);
    return;
  }
  
  const foundArea = findCustomerInAllAreas(k);
  if(foundArea){
    if(foundArea !== currentArea){
      switchArea(foundArea, true);
    }
    document.getElementById("out").innerHTML = render(ALL_DATA[foundArea][k]);
  } else {
    document.getElementById("out").innerHTML = `<div style="color:#f28b82; padding:20px; font-weight:600; text-align:center;">⚠️ Kunde ${k} nicht gefunden.</div>`;
  }
}

function switchArea(area, preserveInput = false){
  currentArea = area;
  DATA = ALL_DATA[area] || {};
  ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0)-(Number(b)||0));

  document.querySelectorAll('.area-btn').forEach(btn => btn.classList.remove('active'));
  const b = document.getElementById(`btn-${area}`);
  if (b) b.classList.add('active');

  updateList();
  
  if(!preserveInput){
    document.getElementById("knr").value = "";
    document.getElementById("out").innerHTML = `<div style="color:#8ab4f8; padding:20px; font-weight:600; text-align:center;">✓ Bereich gewechselt zu: ${getAreaName(area)}<br><br>Bitte Kunden wählen...</div>`;
  }
}

function getAreaName(area){
  const names = {'direkt':'Direkt','mk':'MK','nms':'HuPa NMS','malchow':'HuPa Malchow'};
  return names[area] || area;
}

function updateList(){
  if (!DATA || Object.keys(DATA).length === 0) {
    document.getElementById("list").innerHTML = `
      <div style="padding:20px; text-align:center; color:#9aa0a6; font-size:13px;">
        <p>Keine Kunden im aktuellen Bereich</p>
      </div>
    `;
    return;
  }
  
  document.getElementById("list").innerHTML = ORDER.map(k => {
    const name = (DATA[k] && DATA[k].name) ? DATA[k].name : "";
    return `<div class="item" onclick="document.getElementById('knr').value='${k}';showOne()"><b style="color:#8ab4f8">${k}</b> <span style="color:#5f6368">•</span> <span style="color:#b8b8b8">${esc(name)}</span></div>`;
  }).join("");
}

function printAll(){
  const dialogHtml = `
    <div id="printDialog" style="position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(0,0,0,0.7); display:flex; align-items:center; justify-content:center; z-index:9999;">
      <div style="background:#2d2d2d; padding:30px; border-radius:12px; max-width:500px; width:90%; border:1px solid #3c3c3c;">
        <h3 style="margin-top:0; color:#e8eaed; font-size:20px;">Drucken nach Liefertag</h3>
        <p style="color:#9aa0a6; margin-bottom:20px;">Wählen Sie den Liefertag aus. Die Kunden werden nach Tournummer sortiert gedruckt.</p>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:20px;">
          <button onclick="printByDeliveryDay('Montag')" style="padding:12px; background:#1a73e8; color:white; border:none; cursor:pointer; border-radius:6px; font-weight:600;">Montag</button>
          <button onclick="printByDeliveryDay('Dienstag')" style="padding:12px; background:#1a73e8; color:white; border:none; cursor:pointer; border-radius:6px; font-weight:600;">Dienstag</button>
          <button onclick="printByDeliveryDay('Mittwoch')" style="padding:12px; background:#1a73e8; color:white; border:none; cursor:pointer; border-radius:6px; font-weight:600;">Mittwoch</button>
          <button onclick="printByDeliveryDay('Donnerstag')" style="padding:12px; background:#1a73e8; color:white; border:none; cursor:pointer; border-radius:6px; font-weight:600;">Donnerstag</button>
          <button onclick="printByDeliveryDay('Freitag')" style="padding:12px; background:#1a73e8; color:white; border:none; cursor:pointer; border-radius:6px; font-weight:600;">Freitag</button>
          <button onclick="printByDeliveryDay('Samstag')" style="padding:12px; background:#1a73e8; color:white; border:none; cursor:pointer; border-radius:6px; font-weight:600;">Samstag</button>
        </div>
        <div style="display:flex; gap:10px;">
          <button onclick="printByDeliveryDay('ALLE')" style="flex:1; padding:12px; background:#0f9d58; color:white; border:none; cursor:pointer; border-radius:6px; font-weight:600;">Alle Tage</button>
          <button onclick="closePrintDialog()" style="flex:1; padding:12px; background:#5f6368; color:white; border:none; cursor:pointer; border-radius:6px; font-weight:600;">Abbrechen</button>
        </div>
      </div>
    </div>
  `;
  document.body.insertAdjacentHTML('beforeend', dialogHtml);
}

function closePrintDialog(){
  const dialog = document.getElementById('printDialog');
  if(dialog) dialog.remove();
}

function printByDeliveryDay(day){
  closePrintDialog();
  
  const areaName = getAreaName(currentArea);
  let customersToPrint = [];
  
  if(day === 'ALLE'){
    customersToPrint = ORDER.map(k => ({key: k, data: DATA[k]})).filter(c => c.data);
  } else {
    ORDER.forEach(k => {
      if(DATA[k] && DATA[k].tours && DATA[k].tours[day]){
        const tourNr = DATA[k].tours[day];
        if(tourNr && tourNr !== "—" && String(tourNr).trim() !== ""){
          customersToPrint.push({ key: k, data: DATA[k], tour: tourNr });
        }
      }
    });

    customersToPrint.sort((a, b) => {
      const tourA = String(a.tour).replace(/\\D/g, '');
      const tourB = String(b.tour).replace(/\\D/g, '');
      return (Number(tourA) || 0) - (Number(tourB) || 0);
    });
  }
  
  if(customersToPrint.length === 0){
    alert(`Keine Kunden mit Lieferung am ${day} gefunden.`);
    return;
  }
  
  const message = day === 'ALLE' 
    ? `Möchten Sie wirklich alle ${customersToPrint.length} Kunden aus "${areaName}" drucken?`
    : `Möchten Sie ${customersToPrint.length} Kunden für ${day} (sortiert nach Tour) drucken?`;
    
  if(!confirm(message)) return;
  
  let html = "";
  customersToPrint.forEach(c => {
    html += render(c.data);
  });
  
  document.getElementById("out").innerHTML = html;
  setTimeout(() => window.print(), 500);
}

updateList();
</script>
</body>
</html>
"""

# --- STREAMLIT APP ---
st.set_page_config(page_title="Sendeplan Generator - 4 Bereiche", layout="wide")
st.title("Sendeplan Generator")
st.write("Verarbeitet 4 Bereiche: Direkt, MK, HuPa NMS, HuPa Malchow")

st.subheader("Logo (optional)")
logo_up = st.file_uploader("Logo oben im Druck (PNG/JPG/SVG)", type=["png", "jpg", "jpeg", "svg"])

# Optional Preview
logo_preview_uri = logo_file_to_data_uri(logo_up) or load_logo_data_uri()
if logo_preview_uri:
    st.image(logo_preview_uri, caption="Verwendetes Logo (Vorschau)", use_container_width=True)
else:
    st.info("Kein Logo gewählt/gefunden. (Upload oder Datei 'Logo_NORDfrische Center (NFC).png')")

st.subheader("Excel")
up = st.file_uploader("Excel Datei laden", type=["xlsx"])

if up:
    SHEETS = {
        'direkt': 'Direkt 1 - 99',
        'mk': 'Hupa MK 882',
        'nms': 'Hupa 2221-4444',
        'malchow': 'Hupa 7773-7779'
    }

    all_data = {}

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

                # 1) Triplets
                if d_de in trip:
                    for group_text, f in trip[d_de].items():
                        s = norm(r.get(f.get("Sort")))
                        t = safe_time(r.get(f.get("Zeit")))
                        tag = norm(r.get(f.get("Tag")))

                        if s or t or tag:
                            actual_gid = canon_group_id(s)
                            day_items.append({
                                "liefertag": d_de,
                                "sortiment": s,
                                "bestelltag": tag,
                                "bestellschluss": t,
                                "prio": SORT_PRIO.get(actual_gid, 50)
                            })

                # 2) B-Spalten
                keys = [k for k in bmap.keys() if k[0] == d_de]
                for k in keys:
                    f = bmap[k]
                    s = norm(r.get(f.get("sort", "")))
                    z = safe_time(r.get(f.get("zeit", "")))

                    l_col = f.get("l")
                    if l_col:
                        tag = norm(r.get(l_col, ""))
                        if not tag:
                            tag = k[2]  # Fallback Spaltennamen
                    else:
                        tag = k[2]

                    if s or z:
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
                                "prio": 5.5  # nach Avo (5), vor Werbemittel (6)
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

        all_data[area_key] = data
        st.success(f"✓ {sheet_name}: {len(data)} Kunden verarbeitet")

    # Debug: Prüfe Datenstruktur
    st.write("**Debug-Info:**")
    st.write(f"- all_data Keys: {list(all_data.keys())}")
    st.write(f"- Direkt Kunden-Anzahl: {len(all_data.get('direkt', {}))}")

    # Erstelle JSON
    try:
        json_data = json.dumps(all_data, ensure_ascii=False, separators=(",", ":"))
        st.write(f"- JSON Größe: {len(json_data)} Zeichen")
        st.write(f"- JSON Start: {json_data[:100]}...")
    except Exception as e:
        st.error(f"Fehler beim JSON-Erstellen: {e}")
        json_data = "{}"

    # Erstelle HTML
    html = HTML_TEMPLATE.replace(
        "__DATA_JSON__", json_data
    ).replace(
        "__LOGO_DATAURI__", logo_preview_uri or ""
    )

    st.write(f"- HTML Größe: {len(html)} Zeichen")

    st.write("---")
    st.write(f"**Gesamt:** {sum(len(all_data[k]) for k in all_data)} Kunden in {len(all_data)} Bereichen")
    st.download_button(
        "Download Sendeplan (A4)",
        data=html,
        file_name="sendeplan_4_bereiche.html",
        mime="text/html"
    )
