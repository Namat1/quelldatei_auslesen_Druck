import json
import re
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Kundenkarte HTML Generator", layout="wide")
st.title("Excel → 1x HTML (Kundennummer eingeben → A4 drucken)")

st.write("Lade deine Excel hoch. Danach erzeugt die App eine **einzige HTML-Datei**, "
         "in der du später nur noch die **Kundennummer** eingibst und druckst.")

up = st.file_uploader("Excel (.xlsx)", type=["xlsx"])
if not up:
    st.stop()

df = pd.read_excel(up, engine="openpyxl")
df.columns = [c.strip() for c in df.columns]

required = ["Nr", "SAP-Nr.", "Name", "Strasse", "Plz", "Ort", "Fax", "Fachberater"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Pflichtspalten fehlen: {missing}")
    st.stop()

# Tour columns
tour_cols = {"Montag":"Mo", "Dienstag":"Die", "Mittwoch":"Mitt", "Donnerstag":"Don", "Freitag":"Fr", "Samstag":"Sam"}

# 21 triplets
def find_triplet(day_short):
    # supports Di/Die
    prefixes = [day_short]
    if day_short == "Di":
        prefixes = ["Di", "Die"]
    zeit = sort = tag = None
    for p in prefixes:
        for c in df.columns:
            cl = c.lower().strip()
            if re.fullmatch(rf"{p.lower()}\s*21\s*zeit", cl):
                zeit = c
            if re.fullmatch(rf"{p.lower()}\s*21\s*sort", cl):
                sort = c
            if re.fullmatch(rf"{p.lower()}\s*21\s*tag", cl):
                tag = c
    return zeit, sort, tag

triplets = {
    "Montag": find_triplet("Mo"),
    "Dienstag": find_triplet("Di"),
    "Mittwoch": find_triplet("Mi"),
    "Donnerstag": find_triplet("Do"),
    "Freitag": find_triplet("Fr"),
    "Samstag": find_triplet("Sa"),
}

# Build data dict for HTML (hardcoded JSON)
def norm(x):
    x = "" if pd.isna(x) else str(x)
    return re.sub(r"\s+", " ", x.strip())

data = {}
for _, r in df.iterrows():
    knr = norm(r["Nr"])
    if not knr:
        continue

    tours = {}
    for day, col in tour_cols.items():
        tours[day] = norm(r[col]) if col in df.columns else ""

    bestell = []
    for day, (zc, sc, tc) in triplets.items():
        if not (zc and sc and tc):
            continue
        zeit = norm(r.get(zc, ""))
        sort = norm(r.get(sc, ""))
        tag = norm(r.get(tc, ""))
        if sort:
            bestell.append({
                "liefertag": day,
                "sortiment": sort,
                "bestelltag": tag,
                "bestellschluss": zeit
            })

    data[knr] = {
        "kunden_nr": knr,
        "sap_nr": norm(r["SAP-Nr."]),
        "name": norm(r["Name"]),
        "strasse": norm(r["Strasse"]),
        "plz": norm(r["Plz"]),
        "ort": norm(r["Ort"]),
        "fax": norm(r["Fax"]),
        "fachberater": norm(r["Fachberater"]),
        "tours": tours,
        "bestell": bestell
    }

st.success(f"{len(data)} Kunden in HTML eingebettet.")

html = f"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Fleischwerk – Kundenkarte A4</title>
<style>
  :root {{
    --border:#bbb; --muted:#666; --accent:#1b66b3;
  }}
  body {{ margin:0; font-family: Arial, Helvetica, sans-serif; background:#f3f3f3; }}
  .topbar {{
    position: sticky; top: 0; z-index: 10;
    background: #fff; border-bottom: 1px solid var(--border);
    padding: 10px 12px;
    display:flex; gap:10px; align-items:center; flex-wrap:wrap;
  }}
  .topbar input {{
    padding:10px 12px; font-size:16px; border:1px solid var(--border);
    border-radius:10px; width: 220px;
  }}
  .btn {{
    padding:10px 12px; font-size:14px; border:1px solid var(--border);
    border-radius:10px; background:#fff; cursor:pointer;
  }}
  .btn.primary {{
    border-color: var(--accent); color: var(--accent); font-weight:700;
  }}
  .wrap {{ padding: 12px; }}
  .page {{
    width:210mm; min-height:297mm; background:#fff;
    margin: 10mm auto; padding: 12mm; box-sizing:border-box;
    border:1px solid var(--border);
  }}
  .header {{
    display:flex; justify-content:space-between; gap:12mm;
    border-bottom:2px solid var(--accent);
    padding-bottom:6mm; margin-bottom:6mm;
  }}
  .title {{ font-size:16pt; font-weight:700; line-height:1.1; }}
  .subtitle {{ margin-top:2mm; color:var(--muted); font-size:10.5pt; line-height:1.2; }}
  .meta {{ font-size:10.5pt; line-height:1.4; min-width:70mm; }}
  .grid {{ display:grid; grid-template-columns:1fr 1fr; gap:6mm; margin-bottom:6mm; }}
  .box {{ border:1px solid var(--border); border-radius:10px; padding:4mm; }}
  .boxtitle {{ font-weight:700; color:var(--accent); margin-bottom:3mm; }}
  table {{ width:100%; border-collapse:collapse; font-size:10pt; }}
  th, td {{ border-bottom:1px solid #e5e5e5; padding:2.2mm 2mm; vertical-align:top; }}
  th {{ text-align:left; font-weight:700; background:#f7f7f7; }}
  .right {{ text-align:right; white-space:nowrap; }}
  .muted {{ color:var(--muted); }}
  .dayblock {{ margin-top: 5mm; }}
  .daytitle {{ font-weight:800; margin-bottom:2mm; }}
  .items .sortiment {{ width:64%; }}
  .items .bestelltag {{ width:18%; white-space:nowrap; }}
  .items .zeit {{ width:18%; white-space:nowrap; text-align:right; }}
  .error {{
    max-width: 210mm; margin: 10mm auto; padding: 14px;
    border:1px solid #f0c; border-radius:10px; background:#fff;
    color:#900;
  }}

  @media print {{
    body {{ background:#fff; }}
    .topbar {{ display:none; }}
    .page {{ margin:0; border:none; width:auto; min-height:auto; page-break-after:always; }}
  }}
</style>
</head>
<body>
<div class="topbar">
  <b>Kundennummer:</b>
  <input id="knr" placeholder="z.B. 41391" inputmode="numeric" />
  <button class="btn" onclick="loadCard()">Anzeigen</button>
  <button class="btn primary" onclick="window.print()">Drucken</button>
  <span class="muted" id="hint"></span>
</div>

<div class="wrap">
  <div id="out"></div>
</div>

<script>
const DATA = {json.dumps(data, ensure_ascii=False)};

function esc(s) {{
  return String(s ?? "").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;");
}}

function renderCard(c) {{
  const days = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];

  // Tours
  let tourRows = "";
  for (const d of days) {{
    const t = (c.tours && c.tours[d]) ? String(c.tours[d]).trim() : "";
    if (t) tourRows += `<tr><td>${{esc(d)}}</td><td class="right">${{esc(t)}}</td></tr>`;
  }}
  if (!tourRows) tourRows = `<tr><td colspan="2" class="muted">Keine Tourdaten</td></tr>`;

  // Bestell by delivery day
  const byDay = {{}};
  for (const it of (c.bestell || [])) {{
    if (!byDay[it.liefertag]) byDay[it.liefertag] = [];
    byDay[it.liefertag].push(it);
  }}

  let blocks = "";
  for (const d of days) {{
    const arr = byDay[d];
    if (!arr || arr.length === 0) continue;
    let rows = "";
    for (const it of arr) {{
      rows += `<tr>
        <td class="sortiment">${{esc(it.sortiment)}}</td>
        <td class="bestelltag">${{esc(it.bestelltag)}}</td>
        <td class="zeit">${{esc(it.bestellschluss)}}</td>
      </tr>`;
    }}
    blocks += `
      <div class="dayblock">
        <div class="daytitle">${{esc(d)}}</div>
        <table class="items">
          <thead><tr><th>Sortiment</th><th>Bestelltag</th><th>Bestellschluss</th></tr></thead>
          <tbody>${{rows}}</tbody>
        </table>
      </div>
    `;
  }}
  if (!blocks) blocks = `<div class="muted">Keine Bestelldaten gefunden.</div>`;

  return `
  <div class="page">
    <div class="header">
      <div>
        <div class="title">${{esc(c.name)}}</div>
        <div class="subtitle">${{esc(c.strasse)}}<br>${{esc(c.plz)}} ${{esc(c.ort)}}</div>
      </div>
      <div class="meta">
        <div><b>Kunden-Nr.:</b> ${{esc(c.kunden_nr)}}</div>
        <div><b>SAP-Nr.:</b> ${{esc(c.sap_nr)}}</div>
        ${{c.fachberater ? `<div><b>Fachberater:</b> ${{esc(c.fachberater)}}</div>` : ""}}
        ${{c.fax ? `<div><b>Fax:</b> ${{esc(c.fax)}}</div>` : ""}}
      </div>
    </div>

    <div class="grid">
      <div class="box">
        <div class="boxtitle">Touren (Liefertage)</div>
        <table>
          <thead><tr><th>Tag</th><th class="right">Tour</th></tr></thead>
          <tbody>${{tourRows}}</tbody>
        </table>
      </div>
      <div class="box">
        <div class="boxtitle">Hinweise</div>
        <div class="muted" style="font-size:10pt;line-height:1.35;">
          Drucken: A4 · Skalierung 100% · (optional) Hintergrundgrafiken an.
        </div>
      </div>
    </div>

    <div class="content">
      ${{blocks}}
    </div>

    <div class="footer muted" style="margin-top:8mm;font-size:9pt;border-top:1px solid #eee;padding-top:3mm;">
      Fleischwerk-Kundenkarte – aus Excel generiert
    </div>
  </div>`;
}}

function loadCard() {{
  const knr = document.getElementById("knr").value.trim();
  const out = document.getElementById("out");
  const hint = document.getElementById("hint");

  if (!knr) {{
    out.innerHTML = `<div class="error">Bitte Kundennummer eingeben.</div>`;
    hint.textContent = "";
    return;
  }}
  const c = DATA[knr];
  if (!c) {{
    out.innerHTML = `<div class="error">Kundennummer <b>${{esc(knr)}}</b> nicht gefunden.</div>`;
    hint.textContent = `Vorhanden: ${{Object.keys(DATA).length}} Kunden`;
    return;
  }}
  out.innerHTML = renderCard(c);
  hint.textContent = `${{c.name}} · ${{c.plz}} ${{c.ort}}`;
}}

document.getElementById("knr").addEventListener("keydown", (e) => {{
  if (e.key === "Enter") loadCard();
}});
</script>
</body>
</html>
"""

st.download_button(
    "⬇️ HTML-Datei erzeugen (Kundennummer eingeben → drucken)",
    data=html.encode("utf-8"),
    file_name="kundenkarte_a4.html",
    mime="text/html"
)

st.caption("Hinweis: In dieser Version werden die '21 Zeit/Sort/Tag'-Spalten genutzt. "
           "Wenn du alle Gruppen (0/1011/22/41/65/91/DS) wie im PDF willst, erweitern wir den Parser.")
