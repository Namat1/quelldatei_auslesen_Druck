import re
import io
import zipfile
from dataclasses import dataclass, asdict
from typing import List, Dict, Optional

import pandas as pd
import streamlit as st

# -----------------------------
# Datenmodell
# -----------------------------
WEEKDAYS_DE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]

@dataclass
class Item:
    lieftertag: str
    sortiment: str
    bestelltag: str
    bestellschluss: str

@dataclass
class CustomerCard:
    kunden_nr: str
    sap_nr: str
    name: str
    strasse: str
    plz: str
    ort: str
    fachberater: str = ""
    fax: str = ""
    touren: Dict[str, str] = None  # {"Montag":"1030",...}
    items: List[Item] = None


# -----------------------------
# Parser: "Standard-Text" (PDF-Auszug)
# -----------------------------
def parse_standard_text(text: str) -> CustomerCard:
    """
    Erwartet den Stil, den du gepostet hast:
    - Kopf: Name/Adresse/KundenNr/Fachberater/Liefertage/Tour
    - Danach Blöcke je Liefertag: "Montag ... Sortiment ... Bestelltag ... Bestellschluss ..."
    """
    # Normalize
    t = re.sub(r"\r\n", "\n", text).strip()
    lines = [ln.strip() for ln in t.split("\n") if ln.strip()]

    # Helper
    def find_line(prefix: str) -> Optional[str]:
        for ln in lines:
            if ln.lower().startswith(prefix.lower()):
                return ln
        return None

    # Kopf rausziehen (robust / heuristisch)
    # Name: erste Zeile, die nicht "Standard" ist und nicht Adresse/Meta startet
    name = ""
    start_idx = 0
    if lines and lines[0].lower().startswith("standard"):
        start_idx = 1

    # Oft: "P.+R. STRUVE ... Alle Sortimente Fleischwerk"
    # Wir nehmen die erste sinnvolle Zeile als Name
    for i in range(start_idx, min(start_idx + 5, len(lines))):
        if not any(lines[i].lower().startswith(x) for x in ["kunden-nr", "fachberater", "liefertag", "tour", "bestell"]):
            name = lines[i]
            break

    # Adresse: nächste 2 Zeilen, die wie Straße/PLZ/Ort aussehen
    strasse, plz, ort = "", "", ""
    for ln in lines:
        if re.search(r"\b\d{5}\b", ln) and any(cityword.isalpha() for cityword in ln.replace("-", " ").split()):
            # "22177 HAMBURG"
            m = re.search(r"(\d{5})\s+(.+)$", ln)
            if m:
                plz = m.group(1)
                ort = m.group(2).strip()
    # Straße: Zeile, die eine Hausnummer enthält, aber keine PLZ
    for ln in lines:
        if re.search(r"\d", ln) and not re.search(r"\b\d{5}\b", ln) and any(k in ln.upper() for k in ["STR", "PLATZ", "CHAUSSEE", "WEG", "RING", "ALLEE", "STRASSE", "CH."]):
            strasse = ln
            break

    kunden_nr = ""
    ln_kn = find_line("Kunden-Nr")
    if ln_kn:
        m = re.search(r"(\d+)", ln_kn)
        if m:
            kunden_nr = m.group(1)

    fachberater = ""
    ln_fb = find_line("Fachberater")
    if ln_fb:
        fachberater = ln_fb.split(":", 1)[-1].strip()

    # Liefertage (optional)
    # Touren
    touren = {wd: "" for wd in WEEKDAYS_DE}
    ln_tour = find_line("Tour")
    if ln_tour:
        # "Tour: 1030 2027 3032 4032 5031"
        nums = re.findall(r"\b\d{3,5}\b", ln_tour)
        # Liefertagzeile gibt Reihenfolge
        ln_lt = find_line("Liefertag")
        if ln_lt and ":" in ln_lt:
            # "Liefertag: Montag Dienstag ..."
            day_part = ln_lt.split(":", 1)[-1]
            days = [d for d in WEEKDAYS_DE if d in day_part]
        else:
            # Default: Mo-Fr
            days = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"]
        for d, n in zip(days, nums):
            touren[d] = n

    # Items: wir parsen Liefertag-Blöcke
    items: List[Item] = []
    # Wir bauen einen String ab "Montag ..." (erste vorkommende Wochentag-Zeile)
    full = "\n".join(lines)

    # Split by weekday headings in content:
    # Ersetzt "Montag " am Zeilenanfang/mitten im Text in Marker
    pattern = r"(?m)^(Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag)\b"
    parts = re.split(pattern, full)
    # re.split liefert: [prefix, day1, block1, day2, block2, ...]
    if len(parts) >= 3:
        prefix = parts[0]
        day_blocks = parts[1:]
        for i in range(0, len(day_blocks) - 1, 2):
            day = day_blocks[i].strip()
            block = day_blocks[i + 1].strip()

            # In jedem Block: wiederholte Sequenz:
            # Sortiment (Text) + Bestelltag (Wochentag) + Bestellschluss (Zeit)
            # In deinem Beispiel steht pro Sortiment: Bestelltag dann Zeit in separaten Zeilen.
            # Wir extrahieren alle Zeiten und alle Bestelltag-Keywords und sortiment-Strings dazwischen.
            # Heuristik: Wir suchen Muster "... <Bestelltag> ... <HH:MM Uhr>"
            # Sortiment ist dann der Text davor bis zum vorherigen Treffer.
            block_lines = [ln.strip() for ln in block.split("\n") if ln.strip()]

            # Wir laufen Zeilenweise und merken "aktuelles Sortiment", bis ein Bestelltag+Zeit kommt
            current_sortiment_lines = []
            pending_bestelltag = None

            def flush_if_possible(bt: Optional[str], time_str: Optional[str]):
                nonlocal current_sortiment_lines
                if bt and time_str and current_sortiment_lines:
                    sortiment = " ".join(current_sortiment_lines).strip()
                    # Säubern
                    sortiment = re.sub(r"\s+", " ", sortiment)
                    items.append(Item(lieftertag=day, sortiment=sortiment, bestelltag=bt, bestellschluss=time_str))
                    current_sortiment_lines = []

            for ln in block_lines:
                # Bestelltag?
                bt = None
                for wd in WEEKDAYS_DE:
                    if ln == wd:
                        bt = wd
                        break

                tm = None
                mtime = re.search(r"\b(\d{1,2}:\d{2})\s*Uhr\b", ln)
                if mtime:
                    tm = f"{mtime.group(1)} Uhr"

                if bt:
                    pending_bestelltag = bt
                    continue

                if tm and pending_bestelltag:
                    flush_if_possible(pending_bestelltag, tm)
                    pending_bestelltag = None
                    continue

                # sonst: gehört zum Sortiment
                # (wir ignorieren Überschriftzeilen wie "Liefertag Sortiment Bestelltag Bestellzeitende")
                if "liefertag" in ln.lower() and "sortiment" in ln.lower():
                    continue
                current_sortiment_lines.append(ln)

    # SAP/Fax sind im Standardtext oft nicht enthalten -> bleiben leer (kannst du ergänzen)
    return CustomerCard(
        kunden_nr=kunden_nr,
        sap_nr="",
        name=name,
        strasse=strasse,
        plz=plz,
        ort=ort,
        fachberater=fachberater,
        fax="",
        touren=touren,
        items=items,
    )


# -----------------------------
# Excel-Modus (normiertes Format)
# -----------------------------
def load_normalized_excel(file) -> pd.DataFrame:
    """
    Erwartetes Format (Sheet egal):
    Spalten:
      kunden_nr, sap_nr, name, strasse, plz, ort, fachberater, fax,
      tour_mo, tour_di, tour_mi, tour_do, tour_fr, tour_sa,
      lieftertag, sortiment, bestelltag, bestellschluss
    Pro (Kunde, Liefertag, Sortiment) eine Zeile.
    """
    df = pd.read_excel(file, engine="openpyxl")
    # Normalize column names
    df.columns = [c.strip().lower() for c in df.columns]
    return df


def cards_from_normalized_df(df: pd.DataFrame) -> List[CustomerCard]:
    cards: List[CustomerCard] = []
    key_cols = ["kunden_nr", "name"]
    for col in key_cols:
        if col not in df.columns:
            raise ValueError(f"Spalte fehlt: {col}")

    group_cols = ["kunden_nr"]
    for kunden_nr, g in df.groupby(group_cols):
        row0 = g.iloc[0]
        touren = {
            "Montag": str(row0.get("tour_mo", "") or ""),
            "Dienstag": str(row0.get("tour_di", "") or ""),
            "Mittwoch": str(row0.get("tour_mi", "") or ""),
            "Donnerstag": str(row0.get("tour_do", "") or ""),
            "Freitag": str(row0.get("tour_fr", "") or ""),
            "Samstag": str(row0.get("tour_sa", "") or ""),
        }
        items = []
        for _, r in g.iterrows():
            if pd.isna(r.get("lieftertag")) or pd.isna(r.get("sortiment")):
                continue
            items.append(Item(
                lieftertag=str(r.get("lieftertag")),
                sortiment=str(r.get("sortiment")),
                bestelltag=str(r.get("bestelltag", "")),
                bestellschluss=str(r.get("bestellschluss", "")),
            ))

        cards.append(CustomerCard(
            kunden_nr=str(kunden_nr),
            sap_nr=str(row0.get("sap_nr", "") or ""),
            name=str(row0.get("name", "") or ""),
            strasse=str(row0.get("strasse", "") or ""),
            plz=str(row0.get("plz", "") or ""),
            ort=str(row0.get("ort", "") or ""),
            fachberater=str(row0.get("fachberater", "") or ""),
            fax=str(row0.get("fax", "") or ""),
            touren=touren,
            items=items
        ))
    return cards


# -----------------------------
# HTML Rendering (A4 Print)
# -----------------------------
def render_card_html(card: CustomerCard) -> str:
    tour_rows = ""
    if card.touren:
        for day in ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]:
            tour = card.touren.get(day, "")
            if tour and tour != "nan":
                tour_rows += f"<tr><td>{day}</td><td>{tour}</td></tr>"

    # Items grouped by delivery day
    by_day: Dict[str, List[Item]] = {}
    for it in card.items or []:
        by_day.setdefault(it.lieftertag, []).append(it)

    sections = ""
    for day in WEEKDAYS_DE:
        if day not in by_day:
            continue
        rows = ""
        for it in by_day[day]:
            rows += (
                "<tr>"
                f"<td class='sortiment'>{it.sortiment}</td>"
                f"<td class='bestelltag'>{it.bestelltag}</td>"
                f"<td class='zeit'>{it.bestellschluss}</td>"
                "</tr>"
            )
        sections += f"""
        <div class="dayblock">
          <div class="daytitle">{day}</div>
          <table class="items">
            <thead><tr><th>Sortiment</th><th>Bestelltag</th><th>Bestellschluss</th></tr></thead>
            <tbody>{rows}</tbody>
          </table>
        </div>
        """

    return f"""
    <div class="page">
      <div class="header">
        <div>
          <div class="title">{card.name}</div>
          <div class="subtitle">{card.strasse}<br>{card.plz} {card.ort}</div>
        </div>
        <div class="meta">
          <div><b>Kunden-Nr.:</b> {card.kunden_nr}</div>
          {f"<div><b>SAP-Nr.:</b> {card.sap_nr}</div>" if card.sap_nr else ""}
          {f"<div><b>Fachberater:</b> {card.fachberater}</div>" if card.fachberater else ""}
          {f"<div><b>Fax:</b> {card.fax}</div>" if card.fax else ""}
        </div>
      </div>

      <div class="grid">
        <div class="box">
          <div class="boxtitle">Touren</div>
          <table class="touren">
            <thead><tr><th>Tag</th><th>Tour</th></tr></thead>
            <tbody>{tour_rows or "<tr><td colspan='2' class='muted'>Keine Tourdaten</td></tr>"}</tbody>
          </table>
        </div>

        <div class="box">
          <div class="boxtitle">Hinweis</div>
          <div class="hint">
            A4-Druck: Im Browser „Drucken“ → Skalierung 100% → Ränder „Standard“. <br>
            Sammeldruck: Sammel-HTML öffnen → Drucken → „Alle Seiten“.
          </div>
        </div>
      </div>

      <div class="content">
        {sections or "<div class='muted'>Keine Sortiments-/Bestelldaten gefunden.</div>"}
      </div>

      <div class="footer">
        generiert aus Excel/PDF-Text → HTML A4
      </div>
    </div>
    """


def wrap_document(pages_html: str, doc_title: str = "A4 Druck") -> str:
    # A4 print CSS with page breaks
    return f"""<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{doc_title}</title>
<style>
  :root {{
    --border:#bbb;
    --bg:#fff;
    --muted:#666;
    --accent:#1b66b3;
  }}
  body {{
    margin: 0;
    font-family: Arial, Helvetica, sans-serif;
    background: #f3f3f3;
  }}
  .page {{
    width: 210mm;
    min-height: 297mm;
    background: var(--bg);
    margin: 10mm auto;
    padding: 12mm;
    box-sizing: border-box;
    border: 1px solid var(--border);
  }}
  .header {{
    display:flex;
    justify-content: space-between;
    gap: 12mm;
    border-bottom: 2px solid var(--accent);
    padding-bottom: 6mm;
    margin-bottom: 6mm;
  }}
  .title {{
    font-size: 16pt;
    font-weight: 700;
    line-height: 1.1;
  }}
  .subtitle {{
    margin-top: 2mm;
    color: var(--muted);
    font-size: 10.5pt;
    line-height: 1.2;
  }}
  .meta {{
    font-size: 10.5pt;
    line-height: 1.4;
    min-width: 70mm;
  }}
  .grid {{
    display:grid;
    grid-template-columns: 1fr 1fr;
    gap: 6mm;
    margin-bottom: 6mm;
  }}
  .box {{
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 4mm;
  }}
  .boxtitle {{
    font-weight: 700;
    color: var(--accent);
    margin-bottom: 3mm;
  }}
  table {{
    width:100%;
    border-collapse: collapse;
    font-size: 10pt;
  }}
  th, td {{
    border-bottom: 1px solid #e5e5e5;
    padding: 2.2mm 2mm;
    vertical-align: top;
  }}
  th {{
    text-align: left;
    font-weight: 700;
    background: #f7f7f7;
  }}
  .items .sortiment {{
    width: 64%;
  }}
  .items .bestelltag {{
    width: 18%;
    white-space: nowrap;
  }}
  .items .zeit {{
    width: 18%;
    white-space: nowrap;
    text-align: right;
  }}
  .dayblock {{
    margin-top: 5mm;
  }}
  .daytitle {{
    font-weight: 800;
    margin-bottom: 2mm;
    color:#111;
  }}
  .muted {{
    color: var(--muted);
  }}
  .hint {{
    color: var(--muted);
    font-size: 10pt;
    line-height: 1.35;
  }}
  .footer {{
    margin-top: 8mm;
    color: var(--muted);
    font-size: 9pt;
    border-top: 1px solid #eee;
    padding-top: 3mm;
  }}

  /* PRINT */
  @media print {{
    body {{ background: #fff; }}
    .page {{
      margin: 0;
      border: none;
      border-radius: 0;
      width: auto;
      min-height: auto;
      page-break-after: always;
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
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Fleischwerk A4 Druckkarten", layout="wide")
st.title("Fleischwerk – A4 Druckkarten (Einzeln & Massendruck)")

mode = st.radio("Eingabequelle", ["PDF-Text (Standard) einfügen", "Excel (normiertes Format) hochladen"], horizontal=True)

cards: List[CustomerCard] = []

if mode == "PDF-Text (Standard) einfügen":
    st.info("Füge hier den 'Standard ...' Block (aus deinem PDF-Text) ein. Pro Markt ein Block.")
    raw = st.text_area("Standard-Text", height=320, placeholder="Standard\nP.+R. ...\n...")
    if raw.strip():
        # Optional: mehrere Blöcke durch Trennlinie
        blocks = re.split(r"\n\s*---+\s*\n", raw.strip())
        for b in blocks:
            try:
                cards.append(parse_standard_text(b))
            except Exception as e:
                st.error(f"Fehler beim Parsen eines Blocks: {e}")

else:
    st.info("Erwartet ein normiertes Excel-Format (pro Sortiment eine Zeile).")
    up = st.file_uploader("Excel-Datei (.xlsx)", type=["xlsx"])
    if up:
        try:
            df = load_normalized_excel(up)
            cards = cards_from_normalized_df(df)
            st.success(f"{len(cards)} Kunde(n) geladen.")
            with st.expander("Vorschau Tabelle"):
                st.dataframe(df, use_container_width=True)
        except Exception as e:
            st.error(str(e))

if cards:
    st.subheader("Ausgabe")

    # Auswahl für Einzeldruck
    names = [f"{c.kunden_nr} – {c.name}" for c in cards]
    sel = st.selectbox("Einzeldruck: Markt auswählen", names, index=0)
    idx = names.index(sel)
    single_card = cards[idx]

    # Render single
    single_html = wrap_document(render_card_html(single_card), doc_title=f"A4 – {single_card.kunden_nr}")
    st.download_button(
        "⬇️ Einzel-HTML (A4) herunterladen",
        data=single_html.encode("utf-8"),
        file_name=f"A4_{single_card.kunden_nr}.html",
        mime="text/html"
    )

    # Render batch
    pages = "".join([render_card_html(c) for c in cards])
    batch_html = wrap_document(pages, doc_title="A4 – Sammeldruck")
    st.download_button(
        "⬇️ Sammel-HTML (Massendruck) herunterladen",
        data=batch_html.encode("utf-8"),
        file_name="A4_Sammeldruck.html",
        mime="text/html"
    )

    # ZIP of individual HTMLs
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for c in cards:
            html = wrap_document(render_card_html(c), doc_title=f"A4 – {c.kunden_nr}")
            z.writestr(f"A4_{c.kunden_nr}.html", html)
    st.download_button(
        "⬇️ ZIP mit allen Einzel-HTMLs",
        data=buf.getvalue(),
        file_name="A4_Einzelkarten.zip",
        mime="application/zip"
    )

    with st.expander("Einzel-Vorschau (HTML)"):
        st.components.v1.html(single_html, height=800, scrolling=True)

else:
    st.warning("Noch keine Daten geladen. Bitte oben Text einfügen oder Excel hochladen.")
