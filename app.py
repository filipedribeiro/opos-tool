import streamlit as st
import pandas as pd
import pdfplumber
from rapidfuzz import fuzz
from datetime import datetime
from collections import defaultdict

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

def to_float(val):
    """
    Wandelt einen Wert sicher in eine Zahl um.
    Unterstützt deutsches Format: '8.789,00' → 8789.0
    Unterstützt amerikanisches Format: '8789.00' → 8789.0
    """
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    # Buchstaben am Ende entfernen (z.B. '8.789,00S' → '8.789,00')
    s = s.rstrip("SHsh")
    if s in ("", "-", "–", "0", "0,00", "0.00"):
        return 0.0
    # Leerzeichen entfernen (z.B. '8 789,00')
    s = s.replace(" ", "")
    # Deutsches Format: Punkt als Tausender, Komma als Dezimal
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_datum(val):
    """Versucht verschiedene Datumsformate zu parsen."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def normalize_rechnr(s):
    """Bereinigt eine Rechnungsnummer für den Vergleich."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return str(s).strip().upper().replace(" ", "")


def fuzzy_match(r1, r2, threshold):
    """Gibt True zurück wenn zwei Rechnungsnummern ähnlich genug sind."""
    r1 = normalize_rechnr(r1)
    r2 = normalize_rechnr(r2)
    if not r1 or not r2:
        return False
    return fuzz.ratio(r1, r2) >= threshold


def parse_pdf(uploaded_file):
    """
    Liest OPOS-PDF über Textpositionen aus.
    Funktioniert auch bei engen Spalten und Briefkopf-Dokumenten.
    """
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                x_tolerance=3,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
            )
            if not words:
                continue

            # Zeilen nach Y-Position gruppieren
            zeilen = defaultdict(list)
            for w in words:
                y = round(w["top"] / 4) * 4
                zeilen[y].append(w)

            sorted_y = sorted(zeilen.keys())
            zeilen_liste = [
                sorted(zeilen[y], key=lambda w: w["x0"])
                for y in sorted_y
            ]

            # Header-Zeile anhand OPOS-Schlüsselwörter finden
            opos_keywords = ["rechnungs", "datum", "betrag", "soll", "haben", "beleg"]
            header_idx = None
            best_score = 0
            for i, zeile in enumerate(zeilen_liste):
                text = " ".join(w["text"].lower() for w in zeile)
                score = sum(1 for kw in opos_keywords if kw in text)
                if score > best_score:
                    best_score = score
                    header_idx = i

            if header_idx is None or best_score < 2:
                continue

            # Spalten aus Header ableiten
            header_words = zeilen_liste[header_idx]
            spalten = [(w["x0"], w["text"]) for w in header_words]

            def zuordnen(wort_x, sp):
                best = 0
                for idx, (sx, _) in enumerate(sp):
                    if wort_x >= sx - 8:
                        best = idx
                return best

            # Datenzeilen einlesen
            rows = []
            for zeile in zeilen_liste[header_idx + 1:]:
                if not zeile:
                    continue
                row = [""] * len(spalten)
                for w in zeile:
                    col_idx = zuordnen(w["x0"], spalten)
                    row[col_idx] = (row[col_idx] + " " + w["text"]).strip()
                if any(row):
                    rows.append(row)

            if not rows:
                continue

            header = [s[1] for s in spalten]
            # Doppelte Spaltennamen automatisch umbenennen
            seen = {}
            unique_header = []
            for h in header:
                if h in seen:
                    seen[h] += 1
                    unique_header.append(f"{h}_{seen[h]}")
                else:
                    seen[h] = 0
                    unique_header.append(h)
            df = pd.DataFrame(rows, columns=unique_header)
            df = df.dropna(how="all")
            df = df[df.apply(lambda r: any(str(v).strip() for v in r), axis=1)]
            return df, None

    return None, "Keine OPOS-Tabelle in der PDF gefunden."


def parse_excel(uploaded_file):
    """Liest die Excel- oder CSV-Datei ein."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file, sep=";", encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(uploaded_file, sep=",", encoding="utf-8-sig")
    else:
        df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()
    return df, None


def fmt_eur(val):
    """Formatiert eine Zahl als deutschen Euro-Betrag."""
    try:
        return f"{float(val):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)


def abgleichen(df_opos, df_excel, cfg, stichtag):
    """
    Kernfunktion: Vergleicht OPOS-Liste mit Excel-Buchungen.
    Berücksichtigt Soll/Haben-Spiegelung und Teilbuchungen.
    """
    col_o_rechnr = cfg["opos_rechnr"]
    col_o_datum  = cfg["opos_datum"]
    col_o_soll   = cfg["opos_soll"]
    col_o_haben  = cfg["opos_haben"]
    col_o_text   = cfg["opos_text"]

    col_x_rechnr = cfg["excel_rechnr"]
    col_x_datum  = cfg["excel_datum"]
    col_x_soll   = cfg["excel_soll"]
    col_x_haben  = cfg["excel_haben"]

    fuzzy_thresh = cfg["fuzzy"]
    tol          = cfg["toleranz"]
    spiegelung   = cfg["spiegelung"]

    # ── OPOS vorbereiten ──
    df_opos = df_opos.copy()
    df_opos["_soll"]   = df_opos[col_o_soll].apply(to_float)
    df_opos["_haben"]  = df_opos[col_o_haben].apply(to_float)
    df_opos["_datum"]  = df_opos[col_o_datum].apply(parse_datum)
    df_opos["_rechnr"] = df_opos[col_o_rechnr].apply(normalize_rechnr)

    # ── Excel vorbereiten ──
    df_excel = df_excel.copy()
    df_excel["_soll"]   = df_excel[col_x_soll].apply(to_float)
    df_excel["_haben"]  = df_excel[col_x_haben].apply(to_float)
    df_excel["_datum"]  = df_excel[col_x_datum].apply(parse_datum)
    df_excel["_rechnr"] = df_excel[col_x_rechnr].apply(normalize_rechnr)

    # ── Stichtagsfilter ──
    if stichtag:
        df_opos = df_opos[df_opos["_datum"].apply(
            lambda d: d is not None and d <= stichtag
        )]

    # ── Excel nach Rechnungsnummer gruppieren ──
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
        op_text   = (
            str(op.get(col_o_text, "")).strip()
            if col_o_text in df_opos.columns else ""
        )

        # 1. Exakter Treffer
        ex_rows = excel_gruppen.get(op_rechnr)

        # 2. Fuzzy-Treffer
        if not ex_rows:
            for key, rows in excel_gruppen.items():
                if fuzzy_match(op_rechnr, key, fuzzy_thresh):
                    ex_rows = rows
                    break

        if not ex_rows:
            missing.append({
                "Rechnungsnr.": op[col_o_rechnr],
                "Datum":        op[col_o_datum],
                "Buchungstext": op_text,
                "Betrag Soll":  fmt_eur(op_soll)  if op_soll  else "–",
                "Betrag Haben": fmt_eur(op_haben) if op_haben else "–",
                "Grund":        "Nicht in Excel gefunden"
            })
            continue

        # ── Summen berechnen ──
        ex_sum_soll  = sum(r["_soll"]  for r in ex_rows)
        ex_sum_haben = sum(r["_haben"] for r in ex_rows)
        anzahl       = len(ex_rows)

        # ── Spiegelungslogik ──
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
            "Rechnungsnr.":        op[col_o_rechnr],
            "Datum":               op[col_o_datum],
            "Buchungstext":        op_text,
            "OPOS Betrag Soll":    fmt_eur(op_soll)      if op_soll  else "–",
            "OPOS Betrag Haben":   fmt_eur(op_haben)     if op_haben else "–",
            "Excel Umsatz Haben":  fmt_eur(ex_sum_haben) if spiegelung else fmt_eur(ex_sum_soll),
            "Excel Umsatz Soll":   fmt_eur(ex_sum_soll)  if spiegelung else fmt_eur(ex_sum_haben),
            "Differenz":           fmt_eur(round(max(diff_soll, diff_haben), 2)),
            "Anzahl Teilbuchungen":anzahl,
            "Typ":                 typ,
            "OK":                  ok
        })

    # ── Salden ──
    saldo_opos  = df_opos["_soll"].sum()  - df_opos["_haben"].sum()
    if spiegelung:
        saldo_excel = df_excel["_haben"].sum() - df_excel["_soll"].sum()
    else:
        saldo_excel = df_excel["_soll"].sum()  - df_excel["_haben"].sum()

    return {
        "matched":     pd.DataFrame(matched),
        "missing":     pd.DataFrame(missing),
        "saldo_opos":  saldo_opos,
        "saldo_excel": saldo_excel,
        "opos_count":  len(df_opos),
    }


# ─────────────────────────────────────────────
#  SIDEBAR – EINSTELLUNGEN
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Einstellungen")

    st.subheader("📄 Spalten in der OPOS-PDF")
    opos_rechnr = st.text_input("Rechnungsnummer",  value="Rechnungs-Nr.")
    opos_datum  = st.text_input("Buchungsdatum",    value="Datum")
    opos_text   = st.text_input("Buchungstext",     value="Buchungstext")
    opos_soll   = st.text_input("Betrag Soll",      value="Soll")
    opos_haben  = st.text_input("Betrag Haben",     value="Haben")

    st.subheader("📊 Spalten in der Excel-Datei")
    excel_rechnr = st.text_input("Rechnungsnummer / Buchungsfeld", value="Belegfeld1")
    excel_datum  = st.text_input("Buchungsdatum ",                  value="Datum")
    excel_text   = st.text_input("Buchungstext ",                   value="Buchungstext")
    excel_soll   = st.text_input("Umsatz Soll",                    value="Umsatz Soll")
    excel_haben  = st.text_input("Umsatz Haben",                   value="Umsatz Haben")

    st.subheader("🔄 Abgleichsoptionen")
    spiegelung = st.checkbox(
        "Soll/Haben-Spiegelung aktiv", value=True,
        help="Hauptverband und Landesverband buchen spiegelverkehrt"
    )
    fuzzy = st.slider(
        "Fuzzy-Matching Schwelle (%)", 60, 100, 85,
        help="Wie ähnlich müssen Rechnungsnummern sein? 85% = empfohlen"
    )
    toleranz = st.number_input(
        "Betragstoleranz (€)", value=0.01, step=0.01,
        help="Erlaubte Rundungsdifferenz beim Betragsvergleich"
    )
    stichtag_input = st.date_input(
        "Stichtag für Abstimmung",
        value=datetime(2026, 4, 30).date()
    )


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

st.divider()

# ─────────────────────────────────────────────
#  ANALYSE STARTEN
# ─────────────────────────────────────────────
if st.button("🚀 Analyse starten", type="primary",
             disabled=not (pdf_file and excel_file)):

    cfg = {
        "opos_rechnr":  opos_rechnr,
        "opos_datum":   opos_datum,
        "opos_text":    opos_text,
        "opos_soll":    opos_soll,
        "opos_haben":   opos_haben,
        "excel_rechnr": excel_rechnr,
        "excel_datum":  excel_datum,
        "excel_text":   excel_text,
        "excel_soll":   excel_soll,
        "excel_haben":  excel_haben,
        "spiegelung":   spiegelung,
        "fuzzy":        fuzzy,
        "toleranz":     toleranz,
    }

    # ── PDF einlesen ──
    with st.spinner("📄 PDF wird analysiert..."):
        df_opos, err = parse_pdf(pdf_file)

    if err:
        st.error(f"Fehler beim PDF-Lesen: {err}")
        st.stop()

    # Spalten prüfen
    fehlende = [c for c in [opos_rechnr, opos_datum, opos_soll, opos_haben]
                if c not in df_opos.columns]
    if fehlende:
        st.error(f"Diese Spalten wurden im PDF nicht gefunden: {fehlende}")
        with st.expander("Gefundene Spalten – zum Anpassen in der Seitenleiste"):
            st.write(list(df_opos.columns))
        with st.expander("Vorschau der eingelesenen PDF-Daten"):
            st.dataframe(df_opos.head())
        st.stop()

    # ── Excel einlesen ──
    with st.spinner("📊 Excel wird eingelesen..."):
        df_excel, err = parse_excel(excel_file)

    if err:
        st.error(f"Fehler beim Excel-Lesen: {err}")
        st.stop()

    fehlende = [c for c in [excel_rechnr, excel_datum, excel_soll, excel_haben]
                if c not in df_excel.columns]
    if fehlende:
        st.error(f"Diese Spalten wurden in der Excel-Datei nicht gefunden: {fehlende}")
        with st.expander("Gefundene Spalten – zum Anpassen in der Seitenleiste"):
            st.write(list(df_excel.columns))
        st.stop()

    # ── Abgleich ──
    with st.spinner("🔄 Abgleich läuft..."):
        ergebnis = abgleichen(df_opos, df_excel, cfg, stichtag_input)

    st.success("✅ Analyse abgeschlossen!")

    # ── Kennzahlen ──
    diff = ergebnis["saldo_excel"] - ergebnis["saldo_opos"]
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("OPOS-Positionen", ergebnis["opos_count"])
    k2.metric("Abgeglichen",     len(ergebnis["matched"]))
    k3.metric("Fehlend",         len(ergebnis["missing"]))
    k4.metric("Saldo OPOS",      fmt_eur(ergebnis["saldo_opos"]))
    k5.metric("Saldo Excel",     fmt_eur(ergebnis["saldo_excel"]))
    k6.metric("Differenz",       fmt_eur(diff))

    st.divider()

    # ── Tabs ──
    tab1, tab2, tab3 = st.tabs([
        f"❌ Fehlende Buchungen ({len(ergebnis['missing'])})",
        f"✅ Abgeglichene Buchungen ({len(ergebnis['matched'])})",
        "📋 Abstimmungsübersicht"
    ])

    with tab1:
        if ergebnis["missing"].empty:
            st.success("🎉 Alle Buchungen gefunden – keine fehlenden Positionen!")
        else:
            st.warning(
                f"{len(ergebnis['missing'])} Buchung(en) aus der OPOS-Liste "
                f"fehlen in Excel:"
            )
            st.dataframe(ergebnis["missing"], use_container_width=True)
            csv = ergebnis["missing"].to_csv(
                index=False, sep=";"
            ).encode("utf-8-sig")
            st.download_button(
                "⬇️ Fehlende Buchungen als CSV exportieren",
                csv, "fehlende_buchungen.csv", "text/csv"
            )

    with tab2:
        if ergebnis["matched"].empty:
            st.info("Keine Buchungen konnten abgeglichen werden.")
        else:
            st.dataframe(
                ergebnis["matched"].drop(columns=["OK"], errors="ignore"),
                use_container_width=True
            )
            csv = ergebnis["matched"].to_csv(
                index=False, sep=";"
            ).encode("utf-8-sig")
            st.download_button(
                "⬇️ Abgeglichene Buchungen als CSV exportieren",
                csv, "abgeglichene_buchungen.csv", "text/csv"
            )

    with tab3:
        st.subheader(
            f"Abstimmungsübersicht zum {stichtag_input.strftime('%d.%m.%Y')}"
        )
        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown("**Positionen**")
            st.write(f"OPOS-Positionen gesamt: **{ergebnis['opos_count']}**")
            st.write(f"Abgeglichen: **{len(ergebnis['matched'])}**")
            st.write(f"Fehlend in Excel: **{len(ergebnis['missing'])}**")
            if not ergebnis["matched"].empty:
                teil = len(
                    ergebnis["matched"][
                        ergebnis["matched"]["Typ"] == "Teilbuchungen"
                    ]
                )
                st.write(f"Mit Teilbuchungen: **{teil}**")

        with col_b:
            st.markdown("**Salden**")
            st.write(f"Saldo laut OPOS: **{fmt_eur(ergebnis['saldo_opos'])}**")
            st.write(f"Saldo laut Excel: **{fmt_eur(ergebnis['saldo_excel'])}**")
            if abs(diff) < 0.01:
                st.success("✅ Differenz: 0,00 € – vollständig abgestimmt!")
            else:
                st.warning(f"⚠️ Differenz: {fmt_eur(diff)} – Buchungen prüfen")

        st.divider()
        st.subheader("📌 Spiegelungslogik")
        if spiegelung:
            st.info(
                "**Betrag Soll** (OPOS/Hauptverband) "
                "↔ **Umsatz Haben** (Excel/Landesverband)\n\n"
                "**Betrag Haben** (OPOS/Hauptverband) "
                "↔ **Umsatz Soll** (Excel/Landesverband)"
            )
        else:
            st.info(
                "Spiegelung deaktiviert – Soll wird mit Soll verglichen."
            )

        # Rohdaten zur Kontrolle
        with st.expander("🔍 Vorschau PDF-Rohdaten (erste 10 Zeilen)"):
            st.dataframe(df_opos.head(10))
        with st.expander("🔍 Vorschau Excel-Rohdaten (erste 10 Zeilen)"):
            st.dataframe(df_excel.head(10))

elif not pdf_file or not excel_file:
    st.info("👆 Bitte beide Dateien hochladen um die Analyse zu starten.")
