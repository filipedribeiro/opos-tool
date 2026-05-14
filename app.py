import streamlit as st
import pandas as pd
import pdfplumber
import io
from rapidfuzz import fuzz
from datetime import datetime

# ─────────────────────────────────────────────
#  SEITENKONFIGURATION
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="OPOS-Abstimmung",
    page_icon="📊",
    layout="wide"
)

st.title("📊 OPOS-Abstimmungs-Tool")
st.caption("Automatischer Abgleich von OPOS-Listen mit gebuchten Buchungen")

# ─────────────────────────────────────────────
#  HILFSFUNKTIONEN
# ─────────────────────────────────────────────

def parse_pdf(uploaded_file):
    """Liest die OPOS-PDF aus und gibt einen DataFrame zurück."""
    tables_found = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table and len(table) > 1:
                    tables_found.append(table)

    if not tables_found:
        return None, "Keine Tabellen in der PDF gefunden."

    # Richtige Tabelle finden: die, die typische OPOS-Spalten enthält
    opos_keywords = ["rechnungs", "datum", "betrag", "soll", "haben", "beleg"]
    best_table = None
    best_score = -1

    for table in tables_found:
        if not table:
            continue
        # Header-Zeile suchen – nicht nur erste Zeile prüfen
        for row_idx, row in enumerate(table[:5]):  # erste 5 Zeilen prüfen
            row_text = " ".join(
                str(c).lower() for c in row if c
            )
            score = sum(1 for kw in opos_keywords if kw in row_text)
            if score > best_score:
                best_score = score
                best_table = (table, row_idx)

    if not best_table or best_score < 2:
        return None, "Keine OPOS-Tabelle erkannt. Bitte Spaltenbezeichnungen prüfen."

    table, header_idx = best_table

    # Header aus der erkannten Zeile
    header = [
        str(c).strip() if c else f"Spalte_{i}"
        for i, c in enumerate(table[header_idx])
    ]
    col_count = len(header)

    # Duplikate im Header vermeiden
    seen = {}
    clean_header = []
    for h in header:
        if h in seen:
            seen[h] += 1
            clean_header.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            clean_header.append(h)

    # Datenzeilen ab der Zeile NACH dem Header
    rows = []
    for row in table[header_idx + 1:]:
        if not any(cell for cell in row if cell):
            continue
        row = list(row)
        if len(row) < col_count:
            row += [None] * (col_count - len(row))
        elif len(row) > col_count:
            row = row[:col_count]
        rows.append(row)

    if not rows:
        return None, "Tabelle gefunden aber keine Datenzeilen."

    df = pd.DataFrame(rows, columns=clean_header)
    df = df.dropna(how="all")
    return df, None

def parse_excel(uploaded_file):
    """Liest die Excel- oder CSV-Datei ein."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        # Erstmal mit Semikolon versuchen (deutsches Excel-Format)
        try:
            df = pd.read_csv(uploaded_file, sep=";", encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(uploaded_file, sep=",", encoding="utf-8-sig")
    else:
        df = pd.read_excel(uploaded_file)

    df.columns = df.columns.str.strip()
    return df, None


def normalize_rechnr(s):
    """Bereinigt eine Rechnungsnummer für den Vergleich."""
    if pd.isna(s):
        return ""
    return str(s).strip().upper().replace(" ", "")


def fuzzy_match(r1, r2, threshold):
    """Gibt True zurück wenn zwei Rechnungsnummern ähnlich genug sind."""
    r1 = normalize_rechnr(r1)
    r2 = normalize_rechnr(r2)
    if not r1 or not r2:
        return False
    return fuzz.ratio(r1, r2) >= threshold


def to_float(val):
    """Wandelt einen Wert sicher in eine Zahl um (z.B. '1.234,56' → 1234.56)."""
    if pd.isna(val) or str(val).strip() in ("", "-", "0", "0,00", "0.00"):
        return 0.0
    s = str(val).strip()
    # Deutsches Format: Punkt als Tausender, Komma als Dezimal
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_datum(val):
    """Versucht verschiedene Datumsformate zu parsen."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def abgleichen(df_opos, df_excel, cfg, stichtag):
    """
    Kernfunktion: Vergleicht OPOS-Liste mit Excel-Buchungen.
    Berücksichtigt Soll/Haben-Spiegelung und Teilbuchungen.
    """
    col_o_rechnr = cfg["opos_rechnr"]
    col_o_datum  = cfg["opos_datum"]
    col_o_soll   = cfg["opos_soll"]
    col_o_haben  = cfg["opos_haben"]

    col_x_rechnr = cfg["excel_rechnr"]
    col_x_datum  = cfg["excel_datum"]
    col_x_soll   = cfg["excel_soll"]
    col_x_haben  = cfg["excel_haben"]

    fuzzy_thresh = cfg["fuzzy"]
    tol          = cfg["toleranz"]
    spiegelung   = cfg["spiegelung"]

    # ── Beträge in Zahlen umwandeln ──
    df_opos = df_opos.copy()
    df_opos["_soll"]  = df_opos[col_o_soll].apply(to_float)
    df_opos["_haben"] = df_opos[col_o_haben].apply(to_float)
    df_opos["_datum"] = df_opos[col_o_datum].apply(parse_datum)
    df_opos["_rechnr"] = df_opos[col_o_rechnr].apply(normalize_rechnr)

    df_excel = df_excel.copy()
    df_excel["_soll"]  = df_excel[col_x_soll].apply(to_float)
    df_excel["_haben"] = df_excel[col_x_haben].apply(to_float)
    df_excel["_datum"] = df_excel[col_x_datum].apply(parse_datum)
    df_excel["_rechnr"] = df_excel[col_x_rechnr].apply(normalize_rechnr)

    # ── Stichtagsfilter auf OPOS anwenden ──
    if stichtag:
        df_opos = df_opos[df_opos["_datum"].apply(
            lambda d: d is not None and d <= stichtag
        )]

    # ── Excel-Buchungen nach Rechnungsnummer gruppieren ──
    excel_gruppen = {}
    for _, row in df_excel.iterrows():
        key = row["_rechnr"]
        if key not in excel_gruppen:
            excel_gruppen[key] = []
        excel_gruppen[key].append(row)

    matched = []
    missing = []

    for _, op in df_opos.iterrows():
        op_rechnr = op["_rechnr"]
        op_soll   = op["_soll"]
        op_haben  = op["_haben"]

        # 1. Exakter Treffer suchen
        ex_rows = excel_gruppen.get(op_rechnr)

        # 2. Fuzzy-Treffer falls kein exakter gefunden
        if not ex_rows:
            for key, rows in excel_gruppen.items():
                if fuzzy_match(op_rechnr, key, fuzzy_thresh):
                    ex_rows = rows
                    break

        if not ex_rows:
            missing.append({
                "Rechnungsnr. (OPOS)": op[col_o_rechnr],
                "Datum":               op[col_o_datum],
                "Buchungstext":        op.get(cfg.get("opos_text", ""), ""),
                "Betrag Soll":         op_soll if op_soll else "–",
                "Betrag Haben":        op_haben if op_haben else "–",
                "Grund":               "Nicht in Excel gefunden"
            })
            continue

        # ── Summen aus Excel berechnen ──
        ex_sum_soll  = sum(r["_soll"]  for r in ex_rows)
        ex_sum_haben = sum(r["_haben"] for r in ex_rows)
        anzahl       = len(ex_rows)

        # ── Spiegelungslogik ──
        # OPOS Soll  ↔ Excel Haben
        # OPOS Haben ↔ Excel Soll
        if spiegelung:
            diff_soll  = abs(op_soll  - ex_sum_haben)
            diff_haben = abs(op_haben - ex_sum_soll)
        else:
            diff_soll  = abs(op_soll  - ex_sum_soll)
            diff_haben = abs(op_haben - ex_sum_haben)

        ok = diff_soll <= tol and diff_haben <= tol

        if anzahl > 1:
            typ = "Teilbuchungen"
        elif ok:
            typ = "Direkt"
        else:
            typ = "Differenz"

        matched.append({
            "Rechnungsnr. (OPOS)": op[col_o_rechnr],
            "Datum":               op[col_o_datum],
            "Buchungstext":        op.get(cfg.get("opos_text", ""), ""),
            "OPOS Betrag Soll":    op_soll  if op_soll  else "–",
            "OPOS Betrag Haben":   op_haben if op_haben else "–",
            "Excel Umsatz Haben":  ex_sum_haben if spiegelung else ex_sum_soll,
            "Excel Umsatz Soll":   ex_sum_soll  if spiegelung else ex_sum_haben,
            "Differenz":           round(max(diff_soll, diff_haben), 2),
            "Anzahl Teilbuchungen":anzahl,
            "Typ":                 typ,
            "OK":                  ok
        })

    # ── Salden berechnen ──
    saldo_opos  = df_opos["_soll"].sum()  - df_opos["_haben"].sum()
    if spiegelung:
        saldo_excel = df_excel["_haben"].sum() - df_excel["_soll"].sum()
    else:
        saldo_excel = df_excel["_soll"].sum()  - df_excel["_haben"].sum()

    return {
        "matched":    pd.DataFrame(matched),
        "missing":    pd.DataFrame(missing),
        "saldo_opos": saldo_opos,
        "saldo_excel": saldo_excel,
        "opos_count": len(df_opos),
    }


# ─────────────────────────────────────────────
#  SIDEBAR – EINSTELLUNGEN
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Einstellungen")

    st.subheader("📄 Spalten in der OPOS-PDF")
    opos_rechnr = st.text_input("Rechnungsnummer / Belegnummer", value="Belegnummer")
    opos_datum  = st.text_input("Buchungsdatum",                  value="Belegdatum")
    opos_text   = st.text_input("Buchungstext",                   value="Buchungstext")
    opos_soll   = st.text_input("Betrag Soll",                    value="Betrag Soll")
    opos_haben  = st.text_input("Betrag Haben",                   value="Betrag Haben")

    st.subheader("📊 Spalten in der Excel-Datei")
    excel_rechnr = st.text_input("Rechnungsnummer / Buchungsfeld", value="Buchungsfeld 1")
    excel_datum  = st.text_input("Buchungsdatum ",                  value="Buchungsdatum")
    excel_text   = st.text_input("Buchungstext ",                   value="Buchungstext")
    excel_soll   = st.text_input("Umsatz Soll",                    value="Umsatz Soll")
    excel_haben  = st.text_input("Umsatz Haben",                   value="Umsatz Haben")

    st.subheader("🔄 Abgleichsoptionen")
    spiegelung = st.checkbox("Soll/Haben-Spiegelung aktiv", value=True,
                             help="Hauptverband und Landesverband buchen spiegelverkehrt")
    fuzzy      = st.slider("Fuzzy-Matching Schwelle (%)", 60, 100, 85,
                           help="Wie ähnlich müssen Rechnungsnummern sein? 85% = empfohlen")
    toleranz   = st.number_input("Betragstoleranz (€)", value=0.01, step=0.01,
                                 help="Erlaubte Rundungsdifferenz beim Betragsvergleich")

    stichtag_input = st.date_input("Stichtag für Abstimmung",
                                   value=datetime(2026, 4, 30).date())

# ─────────────────────────────────────────────
#  HAUPTBEREICH – UPLOAD
# ─────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 OPOS-Liste (PDF)")
    pdf_file = st.file_uploader(
        "PDF vom Hauptverband hochladen",
        type=["pdf"],
        help="Die monatliche OPOS-Liste als PDF"
    )
    if pdf_file:
        st.success(f"✓ {pdf_file.name} hochgeladen")

with col2:
    st.subheader("📊 Gebuchte Buchungen (Excel)")
    excel_file = st.file_uploader(
        "Excel- oder CSV-Datei hochladen",
        type=["xlsx", "xls", "csv"],
        help="Interne Buchungsliste des Landesverbands"
    )
    if excel_file:
        st.success(f"✓ {excel_file.name} hochgeladen")

# ─────────────────────────────────────────────
#  ANALYSE STARTEN
# ─────────────────────────────────────────────
st.divider()

if st.button("🚀 Analyse starten", type="primary",
             disabled=not (pdf_file and excel_file)):

    cfg = {
        "opos_rechnr": opos_rechnr,
        "opos_datum":  opos_datum,
        "opos_text":   opos_text,
        "opos_soll":   opos_soll,
        "opos_haben":  opos_haben,
        "excel_rechnr":excel_rechnr,
        "excel_datum": excel_datum,
        "excel_text":  excel_text,
        "excel_soll":  excel_soll,
        "excel_haben": excel_haben,
        "spiegelung":  spiegelung,
        "fuzzy":       fuzzy,
        "toleranz":    toleranz,
    }

    with st.spinner("PDF wird analysiert..."):
        df_opos, err = parse_pdf(pdf_file)

    if err:
        st.error(f"Fehler beim PDF-Lesen: {err}")
        st.stop()

    # Spalten prüfen
    fehlende = [c for c in [opos_rechnr, opos_datum, opos_soll, opos_haben]
                if c not in df_opos.columns]
    if fehlende:
        st.error(f"Diese Spalten wurden im PDF nicht gefunden: {fehlende}")
        st.info(f"Gefundene Spalten: {list(df_opos.columns)}")
        st.stop()

    with st.spinner("Excel wird eingelesen..."):
        df_excel, err = parse_excel(excel_file)

    if err:
        st.error(f"Fehler beim Excel-Lesen: {err}")
        st.stop()

    fehlende = [c for c in [excel_rechnr, excel_datum, excel_soll, excel_haben]
                if c not in df_excel.columns]
    if fehlende:
        st.error(f"Diese Spalten wurden in der Excel-Datei nicht gefunden: {fehlende}")
        st.info(f"Gefundene Spalten: {list(df_excel.columns)}")
        st.stop()

    with st.spinner("Abgleich läuft..."):
        ergebnis = abgleichen(df_opos, df_excel, cfg, stichtag_input)

    # ── ERGEBNISSE ANZEIGEN ──
    st.success("✅ Analyse abgeschlossen!")

    # Kennzahlen
    diff = ergebnis["saldo_excel"] - ergebnis["saldo_opos"]
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("OPOS-Positionen",  ergebnis["opos_count"])
    k2.metric("Abgeglichen",      len(ergebnis["matched"]))
    k3.metric("Fehlend",          len(ergebnis["missing"]),
              delta=f"-{len(ergebnis['missing'])}" if len(ergebnis["missing"]) else None,
              delta_color="inverse")
    k4.metric("Saldo OPOS",       f"{ergebnis['saldo_opos']:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
    k5.metric("Saldo Excel",      f"{ergebnis['saldo_excel']:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
    k6.metric("Differenz",        f"{diff:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."),
              delta="OK" if abs(diff) < 0.01 else "Differenz vorhanden",
              delta_color="normal" if abs(diff) < 0.01 else "inverse")

    st.divider()

    # Tabs für Ergebnisse
    tab1, tab2, tab3 = st.tabs([
        f"❌ Fehlende Buchungen ({len(ergebnis['missing'])})",
        f"✅ Abgeglichene Buchungen ({len(ergebnis['matched'])})",
        "📋 Abstimmungsübersicht"
    ])

    with tab1:
        if ergebnis["missing"].empty:
            st.success("🎉 Alle Buchungen wurden gefunden – keine fehlenden Positionen!")
        else:
            st.warning(f"{len(ergebnis['missing'])} Buchung(en) aus der OPOS-Liste fehlen in Excel:")
            st.dataframe(ergebnis["missing"], use_container_width=True)

            # Export
            csv = ergebnis["missing"].to_csv(index=False, sep=";").encode("utf-8-sig")
            st.download_button(
                "⬇️ Fehlende Buchungen als CSV exportieren",
                csv, "fehlende_buchungen.csv", "text/csv"
            )

    with tab2:
        if ergebnis["matched"].empty:
            st.info("Keine Buchungen konnten abgeglichen werden.")
        else:
            # Farbliche Markierung: Differenzen hervorheben
            def highlight_diff(row):
                if not row["OK"]:
                    return ["background-color: #fff3cd"] * len(row)
                elif row["Typ"] == "Teilbuchungen":
                    return ["background-color: #d1ecf1"] * len(row)
                return [""] * len(row)

            styled = ergebnis["matched"].drop(columns=["OK"]).style.apply(
                highlight_diff, axis=1
            )
            st.dataframe(styled, use_container_width=True)

            csv = ergebnis["matched"].to_csv(index=False, sep=";").encode("utf-8-sig")
            st.download_button(
                "⬇️ Abgeglichene Buchungen als CSV exportieren",
                csv, "abgeglichene_buchungen.csv", "text/csv"
            )

    with tab3:
        st.subheader(f"Abstimmungsübersicht zum {stichtag_input.strftime('%d.%m.%Y')}")

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Positionen**")
            st.write(f"OPOS-Positionen gesamt: **{ergebnis['opos_count']}**")
            st.write(f"Abgeglichen: **{len(ergebnis['matched'])}**")
            st.write(f"Fehlend in Excel: **{len(ergebnis['missing'])}**")
            if not ergebnis["matched"].empty:
                teil = len(ergebnis["matched"][ergebnis["matched"]["Typ"] == "Teilbuchungen"])
                st.write(f"Mit Teilbuchungen: **{teil}**")

        with col_b:
            st.markdown("**Salden**")
            st.write(f"Saldo laut OPOS: **{ergebnis['saldo_opos']:,.2f} €**")
            st.write(f"Saldo laut Excel: **{ergebnis['saldo_excel']:,.2f} €**")
            if abs(diff) < 0.01:
                st.success(f"✅ Differenz: 0,00 € – vollständig abgestimmt!")
            else:
                st.warning(f"⚠️ Differenz: {diff:,.2f} € – Buchungen prüfen")

        st.divider()
        st.subheader("📌 Spiegelungslogik")
        st.info(
            "**Betrag Soll** (OPOS/Hauptverband) ↔ **Umsatz Haben** (Excel/Landesverband)\n\n"
            "**Betrag Haben** (OPOS/Hauptverband) ↔ **Umsatz Soll** (Excel/Landesverband)"
            if spiegelung else
            "Spiegelung ist deaktiviert – Soll wird mit Soll verglichen."
        )

elif not pdf_file or not excel_file:
    st.info("👆 Bitte beide Dateien hochladen um die Analyse zu starten.")
