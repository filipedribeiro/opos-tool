import streamlit as st
import pandas as pd
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
st.caption("Automatischer Abgleich von OPOS-Liste mit gebuchten Buchungen")


# ─────────────────────────────────────────────
#  HILFSFUNKTIONEN
# ─────────────────────────────────────────────

def to_float(val):
    """
    Wandelt einen Wert sicher in float um.
    Unterstützt deutsches Format: '8.789,00' → 8789.0
    """
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        if pd.isna(val):
            return 0.0
        return float(val)
    s = str(val).strip()
    if s in ("", "-", "–", "0", "0,00", "0.00", "None", "nan"):
        return 0.0
    # Buchstaben am Ende entfernen (z.B. '8.789,00S' → '8.789,00')
    s = s.rstrip("SHsh ")
    # Leerzeichen als Tausendertrennzeichen entfernen
    s = s.replace(" ", "")
    # Deutsches Format: '8.789,00' → '8789.00'
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_datum(val):
    """Parst Datum in verschiedenen Formaten inkl. pandas Timestamp."""
    if val is None:
        return None
    if hasattr(val, "date"):
        return val.date()
    if isinstance(val, float) and pd.isna(val):
        return None
    s = str(val).strip()
    if " " in s:
        s = s.split(" ")[0]
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
    """True wenn zwei Rechnungsnummern ähnlich genug sind."""
    r1 = normalize_rechnr(r1)
    r2 = normalize_rechnr(r2)
    if not r1 or not r2:
        return False
    return fuzz.ratio(r1, r2) >= threshold


def fmt_eur(val):
    """Formatiert als deutschen Euro-Betrag: 1234.5 → '1.234,50 €'"""
    try:
        f = float(val)
        if f == 0.0:
            return "–"
        return f"{f:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)


def lade_excel(uploaded_file):
    """Liest Excel oder CSV ein."""
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".csv"):
            try:
                df = pd.read_csv(uploaded_file, sep=";", encoding="utf-8-sig")
            except Exception:
                df = pd.read_csv(uploaded_file, sep=",", encoding="utf-8-sig")
        else:
            df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()
        return df, None
    except Exception as e:
        return None, str(e)


def abgleichen(df_opos, df_excel, cfg, von_datum, bis_datum):
    """
    Kernfunktion: Vergleicht OPOS-Liste mit Excel-Buchungen.
    - Datumsfilter: Von/Bis auf OPOS anwenden
    - Soll/Haben-Spiegelung: OPOS Betrag Soll ↔ Excel Umsatz Haben
    - Teilbuchungen: mehrere Excel-Zeilen mit gleicher Rechnungsnr. werden summiert
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

    # ── Datumsfilter auf OPOS anwenden ──
    # Sicherstellen dass alle Daten als date verglichen werden
    def datum_ok(d, von, bis):
        if d is None:
            return False
        # datetime zu date konvertieren falls nötig
        if hasattr(d, "date"):
            d = d.date()
        if hasattr(von, "date"):
            von = von.date()
        if hasattr(bis, "date"):
            bis = bis.date()
        if von and bis:
            return von <= d <= bis
        elif bis:
            return d <= bis
        elif von:
            return d >= von
        return True

    df_opos = df_opos[df_opos["_datum"].apply(
        lambda d: datum_ok(d, von_datum, bis_datum)
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
                "Rechnungsnr. (OPOS)": op[col_o_rechnr],
                "Datum":               op[col_o_datum],
                "Buchungstext":        op_text,
                "Betrag Soll":         fmt_eur(op_soll),
                "Betrag Haben":        fmt_eur(op_haben),
                "Grund":               "Nicht in Excel gefunden"
            })
            continue

        # ── Summen berechnen ──
        ex_sum_soll  = sum(r["_soll"]  for r in ex_rows)
        ex_sum_haben = sum(r["_haben"] for r in ex_rows)
        anzahl       = len(ex_rows)

        # ── Spiegelungslogik ──
        # OPOS Betrag Soll  ↔ Excel Umsatz Haben
        # OPOS Betrag Haben ↔ Excel Umsatz Soll
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
            "Rechnungsnr. (OPOS)":  op[col_o_rechnr],
            "Datum":                op[col_o_datum],
            "Buchungstext":         op_text,
            "OPOS Betrag Soll":     fmt_eur(op_soll),
            "OPOS Betrag Haben":    fmt_eur(op_haben),
            "Excel Umsatz Haben":   fmt_eur(ex_sum_haben) if spiegelung else fmt_eur(ex_sum_soll),
            "Excel Umsatz Soll":    fmt_eur(ex_sum_soll)  if spiegelung else fmt_eur(ex_sum_haben),
            "Differenz":            fmt_eur(round(max(diff_soll, diff_haben), 2)),
            "Anzahl Teilbuchungen": anzahl,
            "Typ":                  typ,
            "OK":                   ok
        })

    # ── Salden berechnen ──
    saldo_opos = df_opos["_soll"].sum() - df_opos["_haben"].sum()
    if spiegelung:
        saldo_excel = df_excel["_haben"].sum() - df_excel["_soll"].sum()
    else:
        saldo_excel = df_excel["_soll"].sum() - df_excel["_haben"].sum()

    return {
        "matched":     pd.DataFrame(matched),
        "missing":     pd.DataFrame(missing),
        "saldo_opos":  saldo_opos,
        "saldo_excel": saldo_excel,
        "opos_count":  len(df_opos),
        "df_opos":     df_opos,
        "df_excel":    df_excel,
    }


# ─────────────────────────────────────────────
#  SIDEBAR – EINSTELLUNGEN
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Einstellungen")

    # ── Datumsfilter ──
    st.subheader("📅 Datumsfilter")
    von_datum = st.date_input(
        "Von Datum (optional)",
        value=None,
        help="Buchungen VOR diesem Datum werden ignoriert. Leer lassen = kein Von-Filter"
    )
    bis_datum = st.date_input(
        "Bis Datum / Stichtag",
        value=datetime(2026, 4, 30).date(),
        help="Buchungen NACH diesem Datum werden ignoriert"
    )
    if von_datum and bis_datum and von_datum > bis_datum:
        st.error("⚠️ Von-Datum darf nicht nach dem Bis-Datum liegen!")

    # ── OPOS Spalten ──
    st.subheader("📄 Spalten in der OPOS-Datei")
    opos_rechnr = st.text_input("Rechnungsnummer",  value="Rechnungs-Nr.")
    opos_datum  = st.text_input("Buchungsdatum",    value="Datum")
    opos_text   = st.text_input("Buchungstext",     value="Buchungstext")
    opos_soll   = st.text_input("Betrag Soll",      value="Betrag Soll")
    opos_haben  = st.text_input("Betrag Haben",     value="Betrag Haben")

    # ── Excel Spalten ──
    st.subheader("📊 Spalten in der Buchungs-Excel")
    excel_rechnr = st.text_input("Rechnungsnummer / Buchungsfeld", value="Belegfeld1")
    excel_datum  = st.text_input("Buchungsdatum ",                  value="Datum")
    excel_text   = st.text_input("Buchungstext ",                   value="Buchungstext")
    excel_soll   = st.text_input("Umsatz Soll",                    value="Umsatz Soll")
    excel_haben  = st.text_input("Umsatz Haben",                   value="Umsatz Haben")

    # ── Abgleichsoptionen ──
    st.subheader("🔄 Abgleichsoptionen")
    spiegelung = st.checkbox(
        "Soll/Haben-Spiegelung aktiv", value=True,
        help=(
            "Hauptverband und Landesverband buchen spiegelverkehrt:\n"
            "OPOS Betrag Soll ↔ Excel Umsatz Haben\n"
            "OPOS Betrag Haben ↔ Excel Umsatz Soll"
        )
    )
    fuzzy = st.slider(
        "Fuzzy-Matching Schwelle (%)", 60, 100, 85,
        help="Wie ähnlich müssen Rechnungsnummern sein? 85% = empfohlen"
    )
    toleranz = st.number_input(
        "Betragstoleranz (€)", value=0.01, step=0.01,
        help="Erlaubte Rundungsdifferenz beim Betragsvergleich"
    )


# ─────────────────────────────────────────────
#  HAUPTBEREICH – UPLOAD
# ─────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 OPOS-Liste (Excel/CSV)")
    st.caption("Vom Hauptverband – konvertiert aus PDF")
    opos_file = st.file_uploader(
        "OPOS-Datei hochladen",
        type=["xlsx", "xls", "csv"],
        key="opos",
        help="Die OPOS-Liste als Excel oder CSV"
    )
    if opos_file:
        st.success(f"✓ {opos_file.name} hochgeladen")

with col2:
    st.subheader("📊 Gebuchte Buchungen (Excel/CSV)")
    st.caption("Interne Buchungsliste des Landesverbands")
    excel_file = st.file_uploader(
        "Buchungs-Datei hochladen",
        type=["xlsx", "xls", "csv"],
        key="excel",
        help="Interne Buchungsliste des Landesverbands"
    )
    if excel_file:
        st.success(f"✓ {excel_file.name} hochgeladen")

# Datumsfilter-Info anzeigen
if von_datum and bis_datum:
    st.info(
        f"📅 Datumsfilter aktiv: "
        f"**{von_datum.strftime('%d.%m.%Y')}** bis "
        f"**{bis_datum.strftime('%d.%m.%Y')}**"
    )
elif bis_datum:
    st.info(
        f"📅 Stichtag aktiv: bis **{bis_datum.strftime('%d.%m.%Y')}**"
    )

st.divider()

# ─────────────────────────────────────────────
#  ANALYSE STARTEN
# ─────────────────────────────────────────────
datum_ok = not (von_datum and bis_datum and von_datum > bis_datum)

if st.button(
    "🚀 Analyse starten", type="primary",
    disabled=not (opos_file and excel_file and datum_ok)
):
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

    # ── OPOS einlesen ──
    with st.spinner("📄 OPOS-Datei wird eingelesen..."):
        df_opos, err = lade_excel(opos_file)
    if err:
        st.error(f"Fehler beim Lesen der OPOS-Datei: {err}")
        st.stop()

    fehlende = [c for c in [opos_rechnr, opos_datum, opos_soll, opos_haben]
                if c not in df_opos.columns]
    if fehlende:
        st.error(f"Diese Spalten wurden in der OPOS-Datei nicht gefunden: {fehlende}")
        with st.expander("📋 Gefundene Spalten in der OPOS-Datei"):
            st.write(list(df_opos.columns))
        with st.expander("🔍 Vorschau OPOS-Datei (erste 5 Zeilen)"):
            st.dataframe(df_opos.head())
        st.stop()

    # ── Excel einlesen ──
    with st.spinner("📊 Buchungs-Excel wird eingelesen..."):
        df_excel, err = lade_excel(excel_file)
    if err:
        st.error(f"Fehler beim Lesen der Buchungs-Excel: {err}")
        st.stop()

    fehlende = [c for c in [excel_rechnr, excel_datum, excel_soll, excel_haben]
                if c not in df_excel.columns]
    if fehlende:
        st.error(f"Diese Spalten wurden in der Buchungs-Excel nicht gefunden: {fehlende}")
        with st.expander("📋 Gefundene Spalten in der Buchungs-Excel"):
            st.write(list(df_excel.columns))
        with st.expander("🔍 Vorschau Buchungs-Excel (erste 5 Zeilen)"):
            st.dataframe(df_excel.head())
        st.stop()

    # ── Abgleich ──
    with st.spinner("🔄 Abgleich läuft..."):
        ergebnis = abgleichen(df_opos, df_excel, cfg, von_datum, bis_datum)

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
                f"fehlen in den gebuchten Buchungen:"
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
        st.subheader("Abstimmungsübersicht")

        # Datumsfilter-Zusammenfassung
        if von_datum and bis_datum:
            st.info(
                f"📅 Zeitraum: **{von_datum.strftime('%d.%m.%Y')}** "
                f"bis **{bis_datum.strftime('%d.%m.%Y')}**"
            )
        elif bis_datum:
            st.info(f"📅 Stichtag: **{bis_datum.strftime('%d.%m.%Y')}**")

        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown("**Positionen**")
            st.write(f"OPOS-Positionen gesamt: **{ergebnis['opos_count']}**")
            st.write(f"Abgeglichen: **{len(ergebnis['matched'])}**")
            st.write(f"Fehlend: **{len(ergebnis['missing'])}**")
            if not ergebnis["matched"].empty:
                teil = len(ergebnis["matched"][
                    ergebnis["matched"]["Typ"] == "Teilbuchungen"
                ])
                diff_anz = len(ergebnis["matched"][
                    ergebnis["matched"]["Typ"] == "Differenz"
                ])
                st.write(f"Mit Teilbuchungen: **{teil}**")
                st.write(f"Mit Betragsdifferenz: **{diff_anz}**")

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
                "**OPOS Betrag Soll** (Hauptverband) "
                "↔ **Excel Umsatz Haben** (Landesverband)\n\n"
                "**OPOS Betrag Haben** (Hauptverband) "
                "↔ **Excel Umsatz Soll** (Landesverband)"
            )
        else:
            st.info("Spiegelung deaktiviert – Soll wird mit Soll verglichen.")

        with st.expander("🔍 Vorschau OPOS-Rohdaten (erste 10 Zeilen)"):
            st.dataframe(ergebnis["df_opos"].head(10))
        with st.expander("🔍 Vorschau Excel-Rohdaten (erste 10 Zeilen)"):
            st.dataframe(ergebnis["df_excel"].head(10))

elif not opos_file or not excel_file:
    st.info("👆 Bitte beide Dateien hochladen um die Analyse zu starten.")
