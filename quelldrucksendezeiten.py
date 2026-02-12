import io
import re
import zipfile
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

# -----------------------------
# Helpers / Mapping
# -----------------------------
DAY_MAP = {
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

ORDER_DAYS = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip())

def find_col(cols: List[str], target: str) -> str:
    """
    Find column ignoring case and extra spaces.
    """
    t = target.lower().replace(" ", "")
    for c in cols:
        cc = c.lower().replace(" ", "")
        if cc == t:
            return c
    raise KeyError(f"Spalte nicht gefunden: {target}")

def get_triplet_cols(cols: List[str]) -> Dict[str, Tuple[str, str, str]]:
    """
    Finds triplets like:
      'Mo 21 Zeit', 'Mo 21 Sort', 'Mo 21 Tag'
    Returns mapping delivery_day_de -> (zeit_col, sort_col, tag_col)
    """
    # normalize for scan
    cols_norm = {c: c.lower().strip() for c in cols}

    triplets = {}
    # we check prefixes for each weekday variant
    variants = [
        ("Mo", "Montag"),
        ("Di", "Dienstag"),
        ("Mi", "Mittwoch"),
        ("Do", "Donnerstag"),
        ("Fr", "Freitag"),
        ("Sa", "Samstag"),
    ]

    for short, day_de in variants:
        # some files use "Die 21 ..." instead of "Di 21 ..."
        possible_prefixes = [short, "Die"] if short == "Di" else [short]

        zeit_col = sort_col = tag_col = None

        for p in possible_prefixes:
            # match like "Mo 21 Zeit" (allow multiple spaces)
            for c in cols:
                cl = c.lower()
                if re.fullmatch(rf"{p.lower()}\s*21\s*zeit", cl.replace("  ", " ").strip()):
                    zeit_col = c
                if re.fullmatch(rf"{p.lower()}\s*21\s*sort", cl.replace("  ", " ").strip()):
                    sort_col = c
                if re.fullmatch(rf"{p.lower()}\s*21\s*tag", cl.replace("  ", " ").strip()):
                    tag_col = c

        if zeit_col and sort_col and tag_col:
            triplets[day_de] = (zeit_col, sort_col, tag_col)

    return triplets


# -----------------------------
# HTML template (A4)
# -----------------------------
def render_card_html(row: pd.Series,
                     triplets: Dict[str, Tuple[str, str, str]],
                     tour_cols: Dict[str, str]) -> str:
    kunden_nr = norm(row.get("Nr", ""))
    sap = norm(row.get("SAP-Nr.", ""))
    name = norm(row.get("Name", ""))
    strasse = norm(row.get("Strasse", ""))
    plz = norm(row.get("Plz", ""))
    ort = norm(row.get("Ort", ""))
    fachberater = norm(row.get("Fachberater", ""))
    fax = norm(row.get("Fax", ""))

    # Tour table
    tour_rows = ""
    for day_de in ORDER_DAYS:
        col = tour_cols.get(day_de)
        if not col:
            continue
        val = norm(row.get(col, ""))
        if val and val.lower() != "nan":
            tour_rows += f"<tr><td>{day_de}</td><td class='right'>{val}</td></tr>"
    if not tour_rows:
        tour_rows = "<tr><td colspan='2' class='muted'>Keine Tourdaten</td></tr>"

    # Items by delivery day
    day_sections = ""
    for day_de in ORDER_DAYS:
        if day_de not in triplets:
            continue
        zeit_col, sort_col, tag_col = triplets[day_de]
        zeit = norm(row.get(zeit_col, ""))
        sortiment = norm(row.get(sort_col, ""))
        bestelltag = norm(row.get(tag_col, ""))

        # skip empty triplets
        if not (zeit or sortiment or bestelltag) or sortiment.lower() == "nan":
            continue

        day_sections += f"""
        <div class="dayblock">
          <div class="daytitle">{day_de}</div>
          <table class="items">
            <thead><tr><th>Sortiment</th><th>Bestelltag</th><th>Bestellschluss</th></tr></thead>
            <tbody>
              <tr>
                <td class="sortiment">{sortiment}</td>
                <td class="bestelltag">{bestelltag}</td>
                <td class="zeit">{zeit}</td>
              </tr>
            </tbody>
          </table>
        </div>
        """

    if not day_sections:
        day_sections = "<div class='muted'>Keine 21er Bestelldaten gefunden (Zeit/Sort/Tag).</div>"

    return f"""
    <div class="page">
      <div class="header">
        <div>
          <div class="title">{name}</div>
          <div class="subtitle">{strasse}<br>{plz} {ort}</div>
        </div>
        <div class="meta">
          <div><b>Kunden-Nr.:</b> {kunden_nr}</div>
          <div><b>SAP-Nr.:</b> {sap}</div>
          {f"<div><b>Fachberater:</b> {fachberater}</div>" if fachberater else ""}
          {f"<div><b>Fax:</b> {fax}</div>" if fax else ""}
        </div>
      </div>

      <div class="grid">
        <div class="box">
          <div class="boxtitle">Touren (Liefertage)</div>
          <table class="touren">
            <thead><tr><th>Tag</th><th class="right">Tour</th></tr></thead>
            <tbody>{tour_rows}</tbody>
          </table>
        </div>

        <div class="box">
          <div class="boxtitle">Druck</div>
          <div class="hint">
            Browser → Drucken → A4 → Skalierung 100%.<br>
            Sammeldruck: Sammel-HTML öffnen → „Alle Seiten“.
          </div>
        </div>
      </div>

      <div class="content">
        {day_sections}
      </div>

      <div class="footer">
        Fleischwerk-Stammkarte – automatisch aus Excel (21 Zeit/Sort/Tag)
      </div>
    </div>
    """

def wrap_document(pages_html: str, title: str = "A4 Druck") -> str:
    return f"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
  :root {{
    --border:#bbb;
    --muted:#666;
    --accent:#1b66b3;
  }}
  body {{
    margin:0;
    font-family: Arial, Helvetica, sans-serif;
    background:#f3f3f3;
  }}
  .page {{
    width:210mm;
    min-height:297mm;
    background:#fff;
    margin:10mm auto;
    padding:12mm;
    box-sizing:border-box;
    border:1px solid var(--border);
  }}
  .header {{
    display:flex;
    justify-content:space-between;
    gap:12mm;
    border-bottom:2px solid var(--accent);
    padding-bottom:6mm;
    margin-bottom:6mm;
  }}
  .title {{ font-size:16pt; font-weight:700; line-height:1.1; }}
  .subtitle {{ margin-top:2mm; color:var(--muted); font-size:10.5pt; line-height:1.2; }}
  .meta {{ font-size:10.5pt; line-height:1.4; min-width:70mm; }}
  .grid {{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:6mm;
    margin-bottom:6mm;
  }}
  .box {{
    border:1px solid var(--border);
    border-radius:8px;
    padding:4mm;
  }}
  .boxtitle {{ font-weight:700; color:var(--accent); margin-bottom:3mm; }}
  table {{ width:100%; border-collapse:collapse; font-size:10pt; }}
  th, td {{ border-bottom:1px solid #e5e5e5; padding:2.2mm 2mm; vertical-align:top; }}
  th {{ text-align:left; font-weight:700; background:#f7f7f7; }}
  .right {{ text-align:right; white-space:nowrap; }}
  .items .sortiment {{ width:64%; }}
  .items .bestelltag {{ width:18%; white-space:nowrap; }}
  .items .zeit {{ width:18%; white-space:nowrap; text-align:right; }}
  .dayblock {{ margin-top:5mm; }}
  .daytitle {{ font-weight:800; margin-bottom:2mm; }}
  .muted {{ color:var(--muted); }}
  .hint {{ color:var(--muted); font-size:10pt; line-height:1.35; }}
  .footer {{ margin-top:8mm; color:var(--muted); font-size:9pt; border-top:1px solid #eee; padding-top:3mm; }}

  @media print {{
    body {{ background:#fff; }}
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
{pages_html}
</body>
</html>
"""


# -----------------------------
# Streamlit App
# -----------------------------
st.set_page_config(page_title="Fleischwerk A4 Druckkarten (breite Excel)", layout="wide")
st.title("Fleischwerk – A4 Karten aus breiter Excel (Massendruck & Einzeldruck)")

up = st.file_uploader("Excel hochladen (.xlsx)", type=["xlsx"])

if not up:
    st.info("Lade deine Excel hoch. Das Script nutzt: Stamm (Nr/SAP/Name/Adresse), Touren (Mo..Sa) und 21 Zeit/Sort/Tag.")
    st.stop()

df = pd.read_excel(up, engine="openpyxl")
df.columns = [c.strip() for c in df.columns]

# Pflichtspalten check
required = ["Nr", "SAP-Nr.", "Name", "Strasse", "Plz", "Ort"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Pflichtspalten fehlen: {missing}")
    st.stop()

# Tour cols map (Mo/Die/Mitt/Don/Fr/Sam)
tour_cols = {}
tour_map = {"Montag":"Mo", "Dienstag":"Die", "Mittwoch":"Mitt", "Donnerstag":"Don", "Freitag":"Fr", "Samstag":"Sam"}
for day_de, colname in tour_map.items():
    if colname in df.columns:
        tour_cols[day_de] = colname

triplets = get_triplet_cols(df.columns.tolist())
if not triplets:
    st.warning("Keine Tripel 'Mo 21 Zeit / Sort / Tag' gefunden. Prüfe die Header exakt.")
else:
    st.success(f"Gefundene 21er-Tripel: {', '.join(triplets.keys())}")

with st.expander("Vorschau Daten"):
    st.dataframe(df.head(20), use_container_width=True)

# Auswahl Einzeldruck
df["__label"] = df["Nr"].astype(str).str.strip() + " – " + df["Name"].astype(str).str.strip()
labels = df["__label"].tolist()
sel = st.selectbox("Einzeldruck: Markt wählen", labels, index=0)
row = df.loc[df["__label"] == sel].iloc[0]

single_html = wrap_document(render_card_html(row, triplets, tour_cols), title=f"A4 {row['Nr']}")

st.download_button(
    "⬇️ Einzel-HTML herunterladen",
    data=single_html.encode("utf-8"),
    file_name=f"A4_{row['Nr']}.html",
    mime="text/html"
)

# Massendruck
pages = ""
for _, r in df.iterrows():
    pages += render_card_html(r, triplets, tour_cols)

batch_html = wrap_document(pages, title="A4 Sammeldruck")
st.download_button(
    "⬇️ Sammel-HTML (Massendruck) herunterladen",
    data=batch_html.encode("utf-8"),
    file_name="A4_Sammeldruck.html",
    mime="text/html"
)

# ZIP (optional)
buf = io.BytesIO()
with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
    for _, r in df.iterrows():
        html = wrap_document(render_card_html(r, triplets, tour_cols), title=f"A4 {r['Nr']}")
        z.writestr(f"A4_{r['Nr']}.html", html)

st.download_button(
    "⬇️ ZIP mit allen Einzel-HTMLs",
    data=buf.getvalue(),
    file_name="A4_Einzelkarten.zip",
    mime="application/zip"
)

with st.expander("Vorschau (Einzel)"):
    st.components.v1.html(single_html, height=820, scrolling=True)
