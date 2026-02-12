# app.py
# -----------------------------------------------------------------------------
# FINAL VERSION (FIXED):
# - Deutsche See ist automatisch in den Wochentagen enthalten (über bestell[] -> lieftertag)
# - A4-Fit ist jetzt korrekt (keine mm->px Schätzung mehr!): scrollHeight vs clientHeight
# - AutoFit läuft auch VOR dem Drucken (beforeprint) + nach Render
# - Tour-Bar zeigt nur "aktive" Tage (aus Bestellzeilen inkl. Deutsche See + Touren) -> spart Platz, hilft A4
# - Scrollbar bleibt in der Vorschau erhalten
# - Escaping sicherer (& < > ")
# - JSON: ensure_ascii=False (Umlaute bleiben)
# - Triplet-Detection wird NICHT gefiltert (wir nehmen, was wir finden; skip nur komplett leere Zeilen)
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

DAY_SHORT_TO_DE = {
    "Mo": "Montag",
    "Di": "Dienstag",
    "Die": "Dienstag",
    "Mi": "Mittwoch",
    "Mit": "Mittwoch",
    "Mitt": "Mittwoch",
    "Do": "Donnerstag",
    "Don": "Donnerstag",
    "Donn": "Donnerstag",
    "Fr": "Freitag",
    "Sa": "Samstag",
    "Sam": "Samstag",
}

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
    if isinstance(s, datetime.time):
        return s.strftime("%H:%M") + " Uhr"
    if isinstance(s, pd.Timestamp):
        # falls Timestamp wirklich eine Zeit enthält
        try:
            return s.to_pydatetime().strftime("%H:%M") + " Uhr"
        except Exception:
            pass
    s = norm(s)
    if not s:
        return ""
    if "uhr" not in s.lower() and re.fullmatch(r"\d{1,2}:\d{2}", s):
        return s + " Uhr"
    return s


# --- Detektions-Logiken ---

def detect_triplets(columns: List[str]):
    """
    Erkennung klassischer Tripel:
    "<Tag> <Gruppe> Zeit" / "<Tag> <Gruppe> Sort" / "<Tag> <Gruppe> Tag"
    """
    rx = re.compile(
        r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+(.+?)\s+(Zeit|Sort|Tag)$",
        re.IGNORECASE
    )
    found = {}
    for c in columns:
        cc = c.strip()
        m = rx.match(cc)
        if not m:
            continue
        day_de = DAY_SHORT_TO_DE.get(m.group(1))
        if not day_de:
            continue
        grp = m.group(2).strip()
        fld = m.group(3).capitalize()
        found.setdefault(day_de, {}).setdefault(grp, {})[fld] = cc
    return found


def detect_bspalten(columns: List[str]):
    """
    Erkennt B_-Spalten wie:
    "Mo Z 0 B_Sa" / "Mo 0 B_Sa" / "Mo L 0 B_Sa"
    """
    rx = re.compile(
        r"^(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)\s+"
        r"(?:(Z|L)\s+)?"
        r"(.+?)\s+"
        r"B[_ ]?(Mo|Die|Di|Mitt|Mit|Mi|Don|Donn|Do|Fr|Sam|Sa)$",
        re.IGNORECASE
    )
    mapping = {}
    for c in columns:
        cc = c.strip()
        m = rx.match(cc)
        if not m:
            continue
        day_de = DAY_SHORT_TO_DE.get(m.group(1))
        zl = (m.group(2) or "").upper()
        group_id = m.group(3).strip()
        bestell_de = DAY_SHORT_TO_DE.get(m.group(4))
        if not day_de or not bestell_de:
            continue
        key = (day_de, group_id, bestell_de)
        mapping.setdefault(key, {})
        if zl == "Z":
            mapping[key]["zeit"] = cc
        elif zl == "L":
            mapping[key]["l"] = cc
        else:
            mapping[key]["sort"] = cc
    return mapping


def detect_ds_triplets(columns: List[str]):
    """
    DS-Tripel (Deutsche See), z.B.
    "DS Fr zu Mi Zeit" / "DS Fr zu Mi Sort" / "DS Fr zu Mi Tag"
    oder leicht abweichende Schreibweisen – wir halten es tolerant.
    """
    # sehr tolerant: DS <route> (Zeit|Sort|Tag)
    rx = re.compile(r"^DS\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)
    tmp = {}
    for c in columns:
        cc = c.strip()
        m = rx.match(cc)
        if not m:
            continue
        route = m.group(1).strip()
        fld = m.group(2).capitalize()
        tmp.setdefault(route, {})[fld] = cc

    # nur vollständige Tripel
    clean = {}
    for route, fields in tmp.items():
        if all(k in fields for k in ("Zeit", "Sort", "Tag")):
            clean[route] = fields
    return clean


# --- HTML TEMPLATE (Druckstabil + AutoFit korrekt) ---
HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  @page { size: A4; margin: 0; }
  :root{ --bg:#0b1220; --stroke:rgba(255,255,255,.14); --text:rgba(255,255,255,.92); }

  *{ box-sizing:border-box; font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; }
  body{ margin:0; background: var(--bg); color:var(--text); height: 100vh; overflow: hidden; }

  .app{ display:grid; grid-template-columns: 350px 1fr; height:100vh; padding:15px; gap:15px; }
  .sidebar, .main{ background: rgba(255,255,255,.08); border:1px solid var(--stroke); border-radius:12px; overflow:hidden; }

  .list{ height: calc(100vh - 300px); overflow-y:auto; border-top:1px solid var(--stroke); }
  .item{ padding:10px; border-bottom:1px solid rgba(255,255,255,0.06); cursor:pointer; font-size:12px; }
  .item:hover{ background:rgba(255,255,255,0.08); }

  .wrap{
    height: 100%;
    overflow-y: scroll; /* Scrollbar erzwingen */
    padding: 40px 20px;
    display: flex;
    flex-direction: column;
    align-items: center;
    background: #1a2130;
  }

  .paper{
    width: 210mm;
    height: 297mm;
    background: white; color: black;
    padding: 12mm;
    box-shadow: 0 10px 40px rgba(0,0,0,0.8);
    margin-bottom: 30px;
    display: flex; flex-direction: column;
    overflow: hidden; /* wichtig fürs AutoFit */
    --fs: 10.2pt;
  }

  .paper * { font-size: var(--fs); line-height: 1.15; }

  .ptitle{ text-align:center; font-weight:900; font-size:1.55em; margin:0; }
  .pstd{ text-align:center; color:#d0192b; font-weight:900; margin:1mm 0; }
  .psub{ text-align:center; color:#555; margin:0 0 4mm 0; font-weight:900; }

  .head-box { display:flex; justify-content:space-between; gap:10mm; margin-bottom:4mm; border-bottom:1px solid #eee; padding-bottom:3mm; }
  .head-left b{ font-size:1.02em; }
  .head-right{ text-align:right; white-space:nowrap; }

  .meta-line{ margin: 0 0 4mm 0; padding:2.2mm; background:#f4f4f4; border:1px solid #ddd; border-radius:4px; }

  .tour-bar { display:flex; background:#f4f4f4; border:1px solid #ddd; margin-bottom:4mm; padding:2mm; border-radius:4px; justify-content:space-around; gap:2mm; }
  .tour-item { text-align:center; font-size:0.85em; min-width: 22mm; }

  table{ width:100%; border-collapse:collapse; }
  th, td{ border:1px solid #000; padding:1.4mm; text-align:left; vertical-align:top; }
  th{ background:#f2f2f2; font-weight:900; }

  @media print{
    body { overflow: visible; background: white; }
    .sidebar { display:none !important; }
    .app { display:block; padding:0; }
    .wrap { overflow: visible; padding: 0; background: white; }
    .paper { box-shadow: none; margin: 0; page-break-after: always; }
  }
</style>
</head>
<body>
<div class="app">
  <div class="sidebar">
    <div style="padding:15px"><b>Sendeplan Generator</b></div>
    <div style="padding:15px; display:flex; flex-direction:column; gap:8px;">
      <input id="knr" placeholder="Kunden-Nr..." style="width:100%; padding:8px;">
      <button onclick="showOne()" style="padding:8px; cursor:pointer; background:#4fa3ff; color:white; border:none; border-radius:6px; font-weight:900;">Anzeigen</button>
      <button onclick="resetApp()" style="padding:8px; cursor:pointer; background:#ff4f4f; color:white; border:none; border-radius:6px; font-weight:900;">Reset</button>
      <button onclick="showAll()" style="padding:8px; cursor:pointer; border-radius:6px;">Alle laden</button>
      <button onclick="window.print()" style="padding:8px; background:#28a745; color:white; border:none; cursor:pointer; border-radius:6px; font-weight:900;">Drucken</button>
      <div id="hint" style="font-size:12px; opacity:.75; padding-top:4px;"></div>
    </div>
    <div class="list" id="list"></div>
  </div>
  <div class="main"><div class="wrap" id="out">Excel hochladen & Kunden wählen</div></div>
</div>

<script>
const DATA = __DATA_JSON__;
const ORDER = Object.keys(DATA).sort((a,b)=> (Number(a)||0)-(Number(b)||0));
const DAYS = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];

function esc(s){
  return String(s ?? "")
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;");
}

function resetApp() {
  document.getElementById("knr").value = "";
  document.getElementById("out").innerHTML = "App zurückgesetzt. Bitte Kunden wählen.";
  document.getElementById("hint").textContent = "";
}

/* aktive Tage = Bestellzeilen (inkl. Deutsche See) ∪ Touren */
function activeDays(c){
  const set = new Set();
  (c.bestell || []).forEach(it=>{
    const d = String(it.liefertag||"").trim();
    const s = String(it.sortiment||"").trim();
    const t = String(it.bestelltag||"").trim();
    const z = String(it.bestellschluss||"").trim();
    if(d && DAYS.includes(d) && (s || t || z)) set.add(d);
  });
  DAYS.forEach(d=>{
    if(c.tours && String(c.tours[d]||"").trim()!=="") set.add(d);
  });
  return DAYS.filter(d=>set.has(d));
}

function render(c){
  const byDay = {};
  (c.bestell || []).forEach(it => {
    const d = String(it.liefertag || "").trim();
    if(!byDay[d]) byDay[d] = [];
    byDay[d].push(it);
  });

  const rows = DAYS.map(d => {
    const items = byDay[d] || [];
    return `<tr>
      <td style="width:16%"><b>${esc(d)}</b></td>
      <td>${items.map(x => esc(x.sortiment)).join("<br>") || "&nbsp;"}</td>
      <td style="width:16%">${items.map(x => esc(x.bestelltag)).join("<br>") || "&nbsp;"}</td>
      <td style="width:18%">${items.map(x => esc(x.bestellschluss)).join("<br>") || "&nbsp;"}</td>
    </tr>`;
  }).join("");

  const ad = activeDays(c);
  const dayLine = ad.length ? ad.join(" ") : "-";
  const tourHtml = ad.length
    ? ad.map(d => `<div class="tour-item"><b>${esc(d)}</b><br>${esc(c.tours[d] || "—")}</div>`).join("")
    : DAYS.map(d => `<div class="tour-item"><b>${esc(d)}</b><br>${esc(c.tours[d] || "—")}</div>`).join("");

  return `<div class="paper">
    <div class="ptitle">Sende- &amp; Belieferungsplan</div>
    <div class="pstd">${esc(c.plan_typ)}</div>
    <div class="psub">${esc(c.bereich)}</div>

    <div class="head-box">
      <div class="head-left"><b>${esc(c.name)}</b><br>${esc(c.strasse)}<br>${esc(c.plz)} ${esc(c.ort)}</div>
      <div class="head-right">Kunden-Nr.: <b>${esc(c.kunden_nr)}</b><br>Fachberater: <b>${esc(c.fachberater || "")}</b></div>
    </div>

    <div class="meta-line"><b>Liefertag:</b> ${esc(dayLine)}</div>

    <div class="tour-bar">${tourHtml}</div>

    <table>
      <thead><tr><th>Liefertag</th><th>Sortiment</th><th>Bestelltag</th><th>Bestellzeitende</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

/* AutoFit: korrekt per clientHeight (keine px-Schätzung) */
function autoFit(){
  document.querySelectorAll(".paper").forEach(p => {
    let fs = 10.2;
    const minFs = 6.5;
    p.style.setProperty("--fs", fs + "pt");

    // einige Passes, weil Layout nach Schriftänderung nachzieht
    for(let pass=0; pass<3; pass++){
      let safety = 0;
      while(p.scrollHeight > p.clientHeight && fs > minFs && safety < 120){
        fs -= 0.1;
        p.style.setProperty("--fs", fs.toFixed(1) + "pt");
        safety++;
      }
    }
  });
}

function showOne(){
  const k = document.getElementById("knr").value.trim();
  if(DATA[k]) {
    document.getElementById("out").innerHTML = render(DATA[k]);
    document.getElementById("hint").textContent = "AutoFit aktiv: 1×A4 pro Kunde (auch beim Drucken).";
    setTimeout(autoFit, 20);
  } else {
    document.getElementById("out").innerHTML = "Kundennummer nicht gefunden.";
    document.getElementById("hint").textContent = `Vorhanden: ${ORDER.length} Kunden`;
  }
}

function showAll(){
  document.getElementById("out").innerHTML = ORDER.map(k=>render(DATA[k])).join("");
  document.getElementById("hint").textContent = "Massendruck: jede Seite wird automatisch passend skaliert.";
  setTimeout(autoFit, 20);
}

document.getElementById("list").innerHTML = ORDER.map(k=>
  `<div class="item" onclick="document.getElementById('knr').value='${k}';showOne()"><b>${k}</b> - ${esc(DATA[k].name)}</div>`
).join("");

/* Wichtig: vor dem Drucken nochmal fitten */
window.addEventListener("beforeprint", () => { autoFit(); });
</script>
</body>
</html>
"""

# --- STREAMLIT LOGIK ---
st.set_page_config(page_title="Sendeplan Generator", layout="wide")

up = st.file_uploader("Excel Datei wählen", type=["xlsx"])
if up:
    df = pd.read_excel(up, engine="openpyxl")
    cols = [c.strip() for c in df.columns.tolist()]

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
            # 1) B-Spalten (IDs wie 0, 1011, 22, 65, 91, ...)
            keys = [k for k in bmap.keys() if k[0] == d_de]
            for k in sorted(keys, key=lambda x: (str(x[1]))):  # stabile Sortierung
                f = bmap[k]
                s = norm(r.get(f.get("sort", ""), ""))
                z = normalize_time(r.get(f.get("zeit", ""), ""))
                # NICHT filtern – nur komplett leer skippen
                if s or z or k[2]:
                    bestell.append({
                        "liefertag": d_de,
                        "sortiment": s,
                        "bestelltag": k[2],
                        "bestellschluss": z,
                        "priority": 1
                    })

            # 2) Standard-Tripel (z.B. "Mo 21 Zeit/Sort/Tag")
            if d_de in trip:
                for g, f in trip[d_de].items():
                    s = norm(r.get(f.get("Sort", ""), ""))
                    t = norm(r.get(f.get("Tag", ""), ""))
                    z = normalize_time(r.get(f.get("Zeit", ""), ""))
                    if s or t or z:
                        bestell.append({
                            "liefertag": d_de,
                            "sortiment": s,
                            "bestelltag": t,
                            "bestellschluss": z,
                            "priority": 2
                        })

        # 3) Deutsche See Tripel (wenn vorhanden): wir hängen sie an den Tag aus "Tag"-Feld,
        # falls da ein Wochentag steht, sonst lassen wir es im aktuellen d_de nicht raten.
        # In deinem bisherigen Excel ist DS i.d.R. schon passend je Tag getrennt – wenn nicht,
        # steht in Tag oft sowas wie "Montag". Dann mappen wir das.
        for route, f in ds_trip.items():
            s = norm(r.get(f.get("Sort", ""), ""))
            t = norm(r.get(f.get("Tag", ""), ""))
            z = normalize_time(r.get(f.get("Zeit", ""), ""))
            if not (s or t or z):
                continue

            # versuche, den Liefertag aus t zu erkennen (Montag/Dienstag/...)
            d_guess = t.strip()
            d_use = d_guess if d_guess in DAYS_DE else ""  # wenn nicht erkennbar, lassen wir leer
            # wenn leer, dann zur Sicherheit NICHT verlieren: wir hängen an Montag (oder du wünschst "unbekannt")
            if not d_use:
                d_use = "Montag"

            bestell.append({
                "liefertag": d_use,
                "sortiment": s,
                "bestelltag": t,
                "bestellschluss": z,
                "priority": 3
            })

        bestell.sort(key=lambda x: x["priority"])

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
            "bestell": bestell,
        }

    html = HTML_TEMPLATE.replace(
        "__DATA_JSON__",
        json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    )

    st.download_button(
        "HTML herunterladen",
        data=html.encode("utf-8"),
        file_name="sendeplan_final.html",
        mime="text/html"
    )
else:
    st.info("Bitte Excel hochladen.")
