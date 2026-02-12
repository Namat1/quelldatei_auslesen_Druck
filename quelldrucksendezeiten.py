# app.py
# Streamlit: Excel einlesen -> EINE HTML-Datei erzeugen:
# - In der HTML gibst du eine Kundennummer ein
# - Karte wird angezeigt
# - Drucken (A4) (Browser: Strg+P)
#
# Unterstützt:
# - Touren aus Spalten: Mo / Die / Mitt / Don / Fr / Sam
# - Bestelldaten aus ALLEN Tripel-Gruppen nach Muster:
#     "<Tag> <Gruppe> Zeit" + "<Tag> <Gruppe> Sort" + "<Tag> <Gruppe> Tag"
#   z.B. "Mo 21 Zeit", "Mo 21 Sort", "Mo 21 Tag"
#        "Mo 0 Zeit"  (falls vorhanden) ...
#
# Hinweis: Deine Excel hat viele zusätzliche Spalten (Z/L/Bezug), die hier ignoriert werden.

import io
import json
import re
import zipfile
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


# -----------------------------
# Konfiguration
# -----------------------------
PLAN_TYP = "Standard"
BEREICH = "Alle Sortimente Fleischwerk"

DAY_SHORT_TO_DE = {
    "Mo": "Montag",
    "Di": "Dienstag",
    "Die": "Dienstag",
    "Mi": "Mittwoch",
    "Mitt": "Mittwoch",
    "Do": "Donnerstag",
    "Don": "Donnerstag",
    "Fr": "Freitag",
    "Sa": "Samstag",
    "Sam": "Samstag",
}
DAYS_DE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]


def norm(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    s = str(x).strip()
    s = re.sub(r"\s+", " ", s)
    # Excel-Zahlen (Touren) kommen manchmal als 1001.0
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


# -----------------------------
# Tripel-Erkennung (Zeit/Sort/Tag)
# -----------------------------
def detect_triplets(columns: List[str]) -> Dict[str, Dict[str, Dict[str, str]]]:
    """
    Findet alle Tripel nach Schema:
      "<Tag> <Gruppe> Zeit", "<Tag> <Gruppe> Sort", "<Tag> <Gruppe> Tag"

    Rückgabe:
      triplets[day_de][group_key] = {"Zeit": col, "Sort": col, "Tag": col}

    group_key ist string (z.B. "21", "0", "1011", "65", ...)
    """
    # normalisierte Spalten für Matching
    cols = [c.strip() for c in columns]

    # Regex: Tag (Mo|Die|Di|Mi|Do|Fr|Sa) + Gruppe (beliebig, aber ohne "Zeit/Sort/Tag") + Feldname
    # Beispiele:
    # "Mo 21 Zeit"
    # "Don 1011 Sort"
    # "Die 0 Tag"
    rx = re.compile(r"^(Mo|Die|Di|Mi|Do|Don|Mitt|Fr|Sa|Sam)\s+(.+?)\s+(Zeit|Sort|Tag)$", re.IGNORECASE)

    triplets: Dict[str, Dict[str, Dict[str, str]]] = {}

    for c in cols:
        m = rx.match(c)
        if not m:
            continue

        day_short = m.group(1)
        group = m.group(2).strip()
        field = m.group(3).capitalize()

        day_de = DAY_SHORT_TO_DE.get(day_short, "")
        if not day_de:
            continue

        triplets.setdefault(day_de, {}).setdefault(group, {})[field] = c

    # Nur vollständige Tripel behalten
    clean: Dict[str, Dict[str, Dict[str, str]]] = {}
    for day_de, groups in triplets.items():
        for group, fields in groups.items():
            if all(k in fields for k in ("Zeit", "Sort", "Tag")):
                clean.setdefault(day_de, {})[group] = fields

    return clean


def sort_group_key(g: str):
    """
    Sortiert Gruppen sinnvoll:
    - Zahlen zuerst numerisch (0, 21, 65, 91, 1011)
    - sonst alphabetisch
    """
    gs = g.strip()
    if gs.isdigit():
        return (0, int(gs))
    return (1, gs.lower())


# -----------------------------
# HTML Template (standalone)
# -----------------------------
HTML_TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Fleischwerk – Sende- & Belieferungsplan</title>
<style>
  :root {{
    --border:#777;
    --muted:#666;
    --accent:#1b66b3;
  }}
  body {{
    margin:0;
    font-family: Arial, Helvetica, sans-serif;
    background:#f3f3f3;
  }}

  /* Topbar (nicht drucken) */
  .topbar {{
    position: sticky;
    top: 0;
    z-index: 10;
    background: #fff;
    border-bottom: 1px solid #bbb;
    padding: 10px 12px;
    display:flex;
    gap:10px;
    align-items:center;
    flex-wrap:wrap;
  }}
  .topbar input {{
    padding:10px 12px;
    font-size:16px;
    border:1px solid #bbb;
    border-radius:10px;
    width: 220px;
  }}
  .btn {{
    padding:10px 12px;
    font-size:14px;
    border:1px solid #bbb;
    border-radius:10px;
    background:#fff;
    cursor:pointer;
  }}
  .btn.primary {{
    border-color: var(--accent);
    color: var(--accent);
    font-weight:700;
  }}
  .muted {{ color:var(--muted); }}

  .wrap {{ padding: 12px; }}

  /* A4 Seite */
  .page {{
    width: 210mm;
    min-height: 297mm;
    background: #fff;
    margin: 10mm auto;
    padding: 12mm;
    box-sizing: border-box;
    border: 1px solid #bbb;
  }}

  /* PDF-ähnlicher Kopf */
  .headtitle {{
    text-align:center;
    font-size:22pt;
    font-weight:800;
    margin-top:2mm;
  }}
  .headsub {{
    text-align:center;
    font-size:20pt;
    font-weight:900;
    color:#c00;
    margin-top:1mm;
  }}
  .headsmall {{
    text-align:center;
    font-size:11pt;
    color:#444;
    margin-top:1mm;
    margin-bottom:6mm;
  }}

  .topinfo {{
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    gap:10mm;
    margin-bottom:6mm;
  }}
  .addrname {{ font-weight:800; }}
  .meta2 {{
    font-size:10.5pt;
    line-height:1.4;
    min-width:70mm;
  }}

  .linebox {{
    margin: 4mm 0 6mm 0;
    font-size:10.5pt;
  }}
  .line {{ margin: 1mm 0; }}

  /* große Tabelle wie im PDF */
  .bigtable {{
    width:100%;
    border-collapse:collapse;
    font-size:10pt;
  }}
  .bigtable th, .bigtable td {{
    border:1px solid var(--border);
    padding:3mm 2.5mm;
    vertical-align:top;
  }}
  .bigtable th {{
    background:#f2f2f2;
    font-weight:800;
  }}
  .col-day {{ width:18%; font-weight:800; }}
  .col-sort {{ width:50%; }}
  .col-tag  {{ width:16%; }}
  .col-zeit {{ width:16%; text-align:left; white-space:nowrap; }}

  .error {{
    max-width: 210mm;
    margin: 10mm auto;
    padding: 14px;
    border:1px solid #f0c;
    border-radius:10px;
    background:#fff;
    color:#900;
  }}

  @media print {{
    body {{ background:#fff; }}
    .topbar {{ display:none; }}
    .page {{
      margin:0;
      border:none;
      width:auto;
      min-height:auto;
      page-break-after:always;
    }}
  }}
</style>
</head>
<body>

<div class="topbar">
  <b>Kundennummer:</b>
  <input id="knr" placeholder="z.B. 41391" inputmode="numeric" />
  <button class="btn" onclick="loadCard()">Anzeigen</button>
  <button class="btn primary" onclick="window.print()">Drucken</button>
  <button class="btn" onclick="renderAll()">Alle drucken</button>
  <span class="muted" id="hint"></span>
</div>

<div class="wrap">
  <div id="out"></div>
</div>

<script>
const DATA = __DATA_JSON__;

function esc(s) {
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;");
}

function renderCard(c) {
  const days = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];

  // Liefertage (nur wo Tour vorhanden)
  const activeDays = days.filter(d => (c.tours && String(c.tours[d] || "").trim() !== ""));
  const dayLine = activeDays.length ? activeDays.join("  ") : "-";
  const tourLine = activeDays.length ? activeDays.map(d => esc(c.tours[d])).join("  ") : "-";

  // Bestell-Daten nach Liefertag gruppieren
  const byDay = {};
  for (const it of (c.bestell || [])) {
    if (!byDay[it.liefertag]) byDay[it.liefertag] = [];
    byDay[it.liefertag].push(it);
  }

  // Große Tabelle wie im PDF: pro Liefertag eine Tabellenzeile;
  // in den Zellen stehen mehrere Zeilen via <br>
  let bodyRows = "";
  for (const d of days) {
    const arr = byDay[d] || [];
    const sortLines = arr.map(x => esc(x.sortiment)).join("<br>");
    const tagLines  = arr.map(x => esc(x.bestelltag)).join("<br>");
    const zeitLines = arr.map(x => esc(x.bestellschluss)).join("<br>");

    bodyRows += `
      <tr>
        <td class="col-day">${esc(d)}</td>
        <td class="col-sort">${sortLines || "&nbsp;"}</td>
        <td class="col-tag">${tagLines || "&nbsp;"}</td>
        <td class="col-zeit">${zeitLines || "&nbsp;"}</td>
      </tr>
    `;
  }

  return `
    <div class="page">
      <div class="headtitle">Sende- &amp; Belieferungsplan</div>
      <div class="headsub">${esc(c.plan_typ || "Standard")}</div>
      <div class="headsmall">${esc(c.bereich || "Alle Sortimente Fleischwerk")}</div>

      <div class="topinfo">
        <div class="addr">
          <div class="addrname">${esc(c.name)}</div>
          <div>${esc(c.strasse)}</div>
          <div>${esc(c.plz)} ${esc(c.ort)}</div>
        </div>
        <div class="meta2">
          <div><b>Kunden-Nr.:</b> ${esc(c.kunden_nr)}</div>
          <div><b>Fachberater:</b> ${esc(c.fachberater || "")}</div>
          ${c.fax ? `<div><b>Fax:</b> ${esc(c.fax)}</div>` : ""}
          ${c.sap_nr ? `<div><b>SAP-Nr.:</b> ${esc(c.sap_nr)}</div>` : ""}
        </div>
      </div>

      <div class="linebox">
        <div class="line"><b>Liefertag:</b> ${dayLine}</div>
        <div class="line"><b>Tour:</b> ${tourLine}</div>
      </div>

      <table class="bigtable">
        <thead>
          <tr>
            <th>Liefertag</th>
            <th>Sortiment</th>
            <th>Bestelltag</th>
            <th>Bestellzeitende</th>
          </tr>
        </thead>
        <tbody>
          ${bodyRows}
        </tbody>
      </table>
    </div>
  `;
}

function loadCard() {
  const knr = document.getElementById("knr").value.trim();
  const out = document.getElementById("out");
  const hint = document.getElementById("hint");

  if (!knr) {
    out.innerHTML = `<div class="error">Bitte Kundennummer eingeben.</div>`;
    hint.textContent = "";
    return;
  }

  const c = DATA[knr];
  if (!c) {
    out.innerHTML = `<div class="error">Kundennummer <b>${esc(knr)}</b> nicht gefunden.</div>`;
    hint.textContent = `Vorhanden: ${Object.keys(DATA).length} Kunden`;
    return;
  }

  out.innerHTML = renderCard(c);
  hint.textContent = `${c.name} · ${c.plz} ${c.ort}`;
}

function renderAll() {
  const out = document.getElementById("out");
  const hint = document.getElementById("hint");

  const keys = Object.keys(DATA).sort((a,b) => (Number(a)||0) - (Number(b)||0));
  if (!keys.length) {
    out.innerHTML = `<div class="error">Keine Daten vorhanden.</div>`;
    hint.textContent = "";
    return;
  }

  out.innerHTML = keys.map(k => renderCard(DATA[k])).join("");
  hint.textContent = `Alle Kunden gerendert: ${keys.length}`;
}

document.getElementById("knr").addEventListener("keydown", (e) => {
  if (e.key === "Enter") loadCard();
});
</script>

</body>
</html>
"""


# -----------------------------
# Streamlit App
# -----------------------------
st.set_page_config(page_title="Excel → HTML (Kundennummer eingeben → drucken)", layout="wide")
st.title("Excel → HTML (Kundennummer eingeben → A4 drucken)")

up = st.file_uploader("Excel (.xlsx) hochladen", type=["xlsx"])
if not up:
    st.info("Lade deine Excel hoch. Danach bekommst du 1 HTML-Datei: Kundennummer eingeben → Anzeigen → Drucken.")
    st.stop()

df = pd.read_excel(up, engine="openpyxl")
df.columns = [c.strip() for c in df.columns]

# Pflichtspalten
required = ["Nr", "SAP-Nr.", "Name", "Strasse", "Plz", "Ort"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Pflichtspalten fehlen: {missing}")
    st.stop()

# Tourspalten (wenn vorhanden)
tour_cols = {"Montag": "Mo", "Dienstag": "Die", "Mittwoch": "Mitt", "Donnerstag": "Don", "Freitag": "Fr", "Samstag": "Sam"}

# Tripel finden (alle Gruppen)
triplets = detect_triplets(df.columns.tolist())
found_days = [d for d in DAYS_DE if d in triplets]
st.success(f"Tripel-Gruppen erkannt für: {', '.join(found_days) if found_days else '(keine)'}")

with st.expander("Erkannte Gruppen (pro Tag)"):
    if not triplets:
        st.write("Keine Zeit/Sort/Tag-Tripel erkannt.")
    else:
        for d in DAYS_DE:
            groups = sorted(triplets.get(d, {}).keys(), key=sort_group_key)
            st.write(f"**{d}:** {', '.join(groups) if groups else '-'}")

# Datenobjekt für HTML bauen
data: Dict[str, dict] = {}

for _, r in df.iterrows():
    knr = norm(r.get("Nr", ""))
    if not knr:
        continue

    tours = {}
    for day_de, col in tour_cols.items():
        tours[day_de] = norm(r.get(col, "")) if col in df.columns else ""

    bestell_rows = []
    # pro Tag alle Gruppen in sauberer Reihenfolge
    for day_de in DAYS_DE:
        groups = sorted(triplets.get(day_de, {}).keys(), key=sort_group_key)
        for g in groups:
            cols = triplets[day_de][g]
            zeit = norm(r.get(cols["Zeit"], ""))
            sortiment = norm(r.get(cols["Sort"], ""))
            bestelltag = norm(r.get(cols["Tag"], ""))

            # nur wenn wirklich Inhalt da ist
            if sortiment:
                bestell_rows.append({
                    "liefertag": day_de,
                    "sortiment": sortiment,
                    "bestelltag": bestelltag,
                    "bestellschluss": zeit
                })

    data[knr] = {
        "plan_typ": PLAN_TYP,
        "bereich": BEREICH,
        "kunden_nr": knr,
        "sap_nr": norm(r.get("SAP-Nr.", "")),
        "name": norm(r.get("Name", "")),
        "strasse": norm(r.get("Strasse", "")),
        "plz": norm(r.get("Plz", "")),
        "ort": norm(r.get("Ort", "")),
        "fax": norm(r.get("Fax", "")) if "Fax" in df.columns else "",
        "fachberater": norm(r.get("Fachberater", "")) if "Fachberater" in df.columns else "",
        "tours": tours,
        "bestell": bestell_rows,
    }

st.write(f"**{len(data)}** Kunden in die HTML eingebettet.")

# HTML erzeugen (ohne f-string Probleme!)
data_json = json.dumps(data, ensure_ascii=False)
html_out = HTML_TEMPLATE.replace("__DATA_JSON__", data_json)

st.download_button(
    "⬇️ HTML erzeugen (Kundennummer eingeben → drucken)",
    data=html_out.encode("utf-8"),
    file_name="kundenkarte_a4.html",
    mime="text/html",
)

with st.expander("Optional: ZIP mit Einzel-HTMLs"):
    st.caption("Falls du lieber pro Kunde eine separate HTML-Datei willst (Einzeldruck ohne Eingabe).")

    if st.button("ZIP erstellen"):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for knr, obj in data.items():
                single_data = {knr: obj}
                single_json = json.dumps(single_data, ensure_ascii=False)
                single_html = HTML_TEMPLATE.replace("__DATA_JSON__", single_json)
                z.writestr(f"A4_{knr}.html", single_html)

        st.download_button(
            "⬇️ ZIP herunterladen",
            data=buf.getvalue(),
            file_name="A4_Einzelkarten.zip",
            mime="application/zip",
        )
