import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
from datetime import datetime, date
from itertools import combinations

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
    """Wandelt Wert sicher in float. Deutsches Format: '8.789,00' → 8789.0"""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        if pd.isna(val):
            return 0.0
        return float(val)
    s = str(val).strip()
    if s in ("", "-", "–", "0", "0,00", "0.00", "None", "nan"):
        return 0.0
    s = s.rstrip("SHsh ")
    s = s.replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_datum(val):
    """Parst Datum sicher zu datetime.date. Behebt Zeitzonenproblem."""
    if val is None:
        return None
    if hasattr(val, "year") and hasattr(val, "month") and hasattr(val, "day"):
        try:
            return date(int(val.year), int(val.month), int(val.day))
        except Exception:
            pass
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
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


def datum_in_bereich(d, von, bis):
    """Prüft ob Datum im Bereich [von, bis] liegt (beide inklusiv)."""
    if d is None:
        return False
    if isinstance(d, datetime):
        d = date(d.year, d.month, d.day)
    if not isinstance(d, date):
        return False
    if von and bis:
        return von <= d <= bis
    elif bis:
        return d <= bis
    elif von:
        return d >= von
    return True


def normalize_rechnr(s):
    """Bereinigt Rechnungsnummer für Vergleich."""
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
    """Formatiert als deutschen Euro-Betrag."""
    try:
        f = float(val)
        if f == 0.0:
            return "0,00 €"
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


def ist_31_dezember(d):
    """True wenn Datum ein 31. Dezember ist."""
    if d is None:
        return False
    return d.month == 12 and d.day == 31


def finde_erklaerende_buchungen(differenz, offene_buchungen, tol=0.01):
    """
    Sucht Kombinationen von offenen Buchungen die die Differenz erklären.
    Gibt bis zu 5 mögliche Kombinationen zurück.
    Performant: max. 4er-Kombinationen, max. 500 Buchungen.
    """
    if abs(differenz) < tol:
        return []

    ergebnisse = []
    buchungen = offene_buchungen[:500]  # Performance-Limit

    for anzahl in range(1, 5):  # 1er bis 4er Kombinationen
        for kombi in combinations(range(len(buchungen)), anzahl):
            summe = sum(buchungen[i]["betrag"] for i in kombi)
            if abs(summe - abs(differenz)) <= tol:
                ergebnisse.append([buchungen[i] for i in kombi])
            if len(ergebnisse) >= 5:
                return ergebnisse

    return ergebnisse


def abgleichen(df_opos, df_excel, cfg, von_datum, bis_datum,
               extra_31_12=None):
    """
    Kernfunktion: Vollständiger Buchungsabgleich mit Saldenberechnung.
    - Datumsfilter nur auf OPOS
    - Soll/Haben-Spiegelung
    - Teilbuchungen
    - Optionale 31.12.-Buchungen
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
    df_opos["_is31"]   = df_opos["_datum"].apply(ist_31_dezember)

    # ── Excel vorbereiten ──
    df_excel = df_excel.copy()
    df_excel["_soll"]   = df_excel[col_x_soll].apply(to_float)
    df_excel["_haben"]  = df_excel[col_x_haben].apply(to_float)
    df_excel["_datum"]  = df_excel[col_x_datum].apply(parse_datum)
    df_excel["_rechnr"] = df_excel[col_x_rechnr].apply(normalize_rechnr)

    # ── 31.12.-Buchungen aus OPOS identifizieren (vor Filter) ──
    df_31_12 = df_opos[df_opos["_is31"]].copy()

    # ── Datumsfilter auf OPOS (ohne 31.12. die extra ausgewählt werden) ──
    df_opos_filtered = df_opos[df_opos["_datum"].apply(
        lambda d: datum_in_bereich(d, von_datum, bis_datum) and not ist_31_dezember(d)
    )].copy()

    # ── Ausgewählte 31.12.-Buchungen hinzufügen ──
    if extra_31_12 is not None and len(extra_31_12) > 0:
        extra_rows = df_31_12[df_31_12[col_o_rechnr].astype(str).isin(
            [str(r) for r in extra_31_12]
        )]
        df_opos_filtered = pd.concat([df_opos_filtered, extra_rows], ignore_index=True)

    # ── Excel nach Rechnungsnummer gruppieren ──
    excel_gruppen = {}
    for _, row in df_excel.iterrows():
        key = row["_rechnr"]
        if key not in excel_gruppen:
            excel_gruppen[key] = []
        excel_gruppen[key].append(row)

    matched = []
    missing = []

    for _, op in df_opos_filtered.iterrows():
        op_rechnr = op["_rechnr"]
        op_soll   = op["_soll"]
        op_haben  = op["_haben"]
        op_text   = (
            str(op.get(col_o_text, "")).strip()
            if col_o_text in df_opos_filtered.columns else ""
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
                "_soll":               op_soll,
                "_haben":              op_haben,
                "Grund":               "Nicht in Excel gefunden"
            })
            continue

        ex_sum_soll  = sum(r["_soll"]  for r in ex_rows)
        ex_sum_haben = sum(r["_haben"] for r in ex_rows)
        anzahl       = len(ex_rows)

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
    # OPOS-Saldo
    opos_summe_soll  = df_opos_filtered["_soll"].sum()
    opos_summe_haben = df_opos_filtered["_haben"].sum()
    saldo_opos       = opos_summe_soll - opos_summe_haben

    # Excel-Saldo (gespiegelt: Haben - Soll)
    excel_summe_soll  = df_excel["_soll"].sum()
    excel_summe_haben = df_excel["_haben"].sum()
    if spiegelung:
        saldo_excel = excel_summe_haben - excel_summe_soll
    else:
        saldo_excel = excel_summe_soll - excel_summe_haben

    differenz = saldo_excel - saldo_opos

    # ── Offene Buchungen für Differenzanalyse vorbereiten ──
    offene_fuer_analyse = []
    for m in missing:
        betrag = m["_soll"] if m["_soll"] > 0 else m["_haben"]
        if betrag > 0:
            offene_fuer_analyse.append({
                "rechnr":    m["Rechnungsnr. (OPOS)"],
                "datum":     m["Datum"],
                "text":      m["Buchungstext"],
                "betrag":    betrag,
                "soll":      m["_soll"],
                "haben":     m["_haben"],
            })

    return {
        "matched":           pd.DataFrame(matched),
        "missing":           pd.DataFrame(missing).drop(
            columns=["_soll", "_haben"], errors="ignore"
        ),
        "missing_raw":       missing,
        "opos_summe_soll":   opos_summe_soll,
        "opos_summe_haben":  opos_summe_haben,
        "saldo_opos":        saldo_opos,
        "excel_summe_soll":  excel_summe_soll,
        "excel_summe_haben": excel_summe_haben,
        "saldo_excel":       saldo_excel,
        "differenz":         differenz,
        "opos_count":        len(df_opos_filtered),
        "df_opos":           df_opos_filtered,
        "df_excel":          df_excel,
        "df_31_12":          df_31_12,
        "offene_analyse":    offene_fuer_analyse,
        "col_o_rechnr":      col_o_rechnr,
        "col_o_text":        col_o_text,
    }


# ─────────────────────────────────────────────
#  SIDEBAR – EINSTELLUNGEN
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Einstellungen")

    st.subheader("📅 Datumsfilter")
    von_datum = st.date_input(
        "Von Datum (optional)",
        value=None,
        help="Buchungen VOR diesem Datum werden ignoriert"
    )
    bis_datum = st.date_input(
        "Bis Datum / Stichtag",
        value=datetime(2026, 4, 30).date(),
        help="Buchungen NACH diesem Datum werden ignoriert"
    )
    if von_datum and bis_datum and von_datum > bis_datum:
        st.error("⚠️ Von-Datum darf nicht nach dem Bis-Datum liegen!")

    st.subheader("📄 Spalten OPOS-Datei")
    opos_rechnr = st.text_input("Rechnungsnummer",  value="Rechnungs-Nr.")
    opos_datum  = st.text_input("Buchungsdatum",    value="Datum")
    opos_text   = st.text_input("Buchungstext",     value="Buchungstext")
    opos_soll   = st.text_input("Betrag Soll",      value="Betrag Soll")
    opos_haben  = st.text_input("Betrag Haben",     value="Betrag Haben")

    st.subheader("📊 Spalten Buchungs-Excel")
    excel_rechnr = st.text_input("Rechnungsnummer / Buchungsfeld", value="Belegfeld1")
    excel_datum  = st.text_input("Buchungsdatum ",                  value="Datum")
    excel_text   = st.text_input("Buchungstext ",                   value="Buchungstext")
    excel_soll   = st.text_input("Umsatz Soll",                    value="Umsatz Soll")
    excel_haben  = st.text_input("Umsatz Haben",                   value="Umsatz Haben")

    st.subheader("🔄 Abgleichsoptionen")
    spiegelung = st.checkbox(
        "Soll/Haben-Spiegelung aktiv", value=True,
        help="OPOS Betrag Soll ↔ Excel Umsatz Haben"
    )
    fuzzy = st.slider("Fuzzy-Matching (%)", 60, 100, 85)
    toleranz = st.number_input("Betragstoleranz (€)", value=0.01, step=0.01)


# ─────────────────────────────────────────────
#  HAUPTBEREICH – UPLOAD
# ─────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 OPOS-Liste (Excel/CSV)")
    st.caption("Vom Hauptverband")
    opos_file = st.file_uploader(
        "OPOS-Datei hochladen",
        type=["xlsx", "xls", "csv"],
        key="opos"
    )
    if opos_file:
        st.success(f"✓ {opos_file.name}")

with col2:
    st.subheader("📊 Gebuchte Buchungen (Excel/CSV)")
    st.caption("Interne Liste Landesverband")
    excel_file = st.file_uploader(
        "Buchungs-Datei hochladen",
        type=["xlsx", "xls", "csv"],
        key="excel"
    )
    if excel_file:
        st.success(f"✓ {excel_file.name}")

if von_datum and bis_datum:
    st.info(f"📅 Zeitraum: **{von_datum.strftime('%d.%m.%Y')}** bis **{bis_datum.strftime('%d.%m.%Y')}** (inklusiv)")
elif bis_datum:
    st.info(f"📅 Stichtag: bis **{bis_datum.strftime('%d.%m.%Y')}** (inklusiv)")

st.divider()

# ─────────────────────────────────────────────
#  SESSION STATE
# ─────────────────────────────────────────────
if "ergebnis" not in st.session_state:
    st.session_state.ergebnis = None
if "extra_31_12" not in st.session_state:
    st.session_state.extra_31_12 = []
if "df_opos_raw" not in st.session_state:
    st.session_state.df_opos_raw = None
if "df_excel_raw" not in st.session_state:
    st.session_state.df_excel_raw = None


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
    st.session_state.cfg = cfg
    st.session_state.von_datum = von_datum
    st.session_state.bis_datum = bis_datum
    st.session_state.extra_31_12 = []

    with st.spinner("📄 OPOS wird eingelesen..."):
        df_opos, err = lade_excel(opos_file)
    if err:
        st.error(f"Fehler OPOS: {err}")
        st.stop()

    fehlende = [c for c in [opos_rechnr, opos_datum, opos_soll, opos_haben]
                if c not in df_opos.columns]
    if fehlende:
        st.error(f"Spalten nicht gefunden in OPOS: {fehlende}")
        with st.expander("Gefundene Spalten"):
            st.write(list(df_opos.columns))
        st.stop()

    with st.spinner("📊 Excel wird eingelesen..."):
        df_excel, err = lade_excel(excel_file)
    if err:
        st.error(f"Fehler Excel: {err}")
        st.stop()

    fehlende = [c for c in [excel_rechnr, excel_datum, excel_soll, excel_haben]
                if c not in df_excel.columns]
    if fehlende:
        st.error(f"Spalten nicht gefunden in Excel: {fehlende}")
        with st.expander("Gefundene Spalten"):
            st.write(list(df_excel.columns))
        st.stop()

    st.session_state.df_opos_raw  = df_opos
    st.session_state.df_excel_raw = df_excel

    with st.spinner("🔄 Abgleich läuft..."):
        st.session_state.ergebnis = abgleichen(
            df_opos, df_excel, cfg, von_datum, bis_datum, []
        )

# ─────────────────────────────────────────────
#  ERGEBNISSE ANZEIGEN
# ─────────────────────────────────────────────
if st.session_state.ergebnis:
    ergebnis = st.session_state.ergebnis
    cfg      = st.session_state.cfg
    st.success("✅ Analyse abgeschlossen!")

    # ══════════════════════════════════════════
    #  SALDENÜBERSICHT
    # ══════════════════════════════════════════
    st.subheader("📊 Saldenübersicht")

    col_o, col_x, col_d = st.columns(3)

    with col_o:
        st.markdown("**OPOS – Hauptverband**")
        st.metric("Summe Betrag Soll",  fmt_eur(ergebnis["opos_summe_soll"]))
        st.metric("Summe Betrag Haben", fmt_eur(ergebnis["opos_summe_haben"]))
        st.metric("Saldo (Soll – Haben)", fmt_eur(ergebnis["saldo_opos"]))

    with col_x:
        st.markdown("**Excel – Landesverband**")
        st.metric("Summe Umsatz Soll",  fmt_eur(ergebnis["excel_summe_soll"]))
        st.metric("Summe Umsatz Haben", fmt_eur(ergebnis["excel_summe_haben"]))
        if cfg["spiegelung"]:
            st.metric("Saldo (Haben – Soll)", fmt_eur(ergebnis["saldo_excel"]))
        else:
            st.metric("Saldo (Soll – Haben)", fmt_eur(ergebnis["saldo_excel"]))

    with col_d:
        st.markdown("**Differenz**")
        diff = ergebnis["differenz"]
        if abs(diff) < cfg["toleranz"]:
            st.success("✅ Kein Unterschied")
            st.metric("Differenz", "0,00 €")
        else:
            richtung = "Excel höher als OPOS" if diff > 0 else "OPOS höher als Excel"
            st.warning(f"⚠️ {richtung}")
            st.metric("Differenz", fmt_eur(abs(diff)))
        st.metric("Abgeglichen",    len(ergebnis["matched"]))
        st.metric("Fehlend",        len(ergebnis["missing"]))

    st.divider()

    # ══════════════════════════════════════════
    #  31.12.-BUCHUNGEN AUSWAHL
    # ══════════════════════════════════════════
    df_31 = ergebnis["df_31_12"]
    col_rechnr = ergebnis["col_o_rechnr"]
    col_text   = ergebnis["col_o_text"]

    if len(df_31) > 0:
        st.subheader("📅 Optionale 31.12.-Buchungen")
        st.info(
            f"Es wurden **{len(df_31)} Buchungen** mit Datum 31. Dezember gefunden. "
            f"Diese liegen außerhalb des gewählten Zeitraums aber können optional "
            f"in die Abstimmung einbezogen werden."
        )

        auswahl = []
        for _, row in df_31.iterrows():
            rechnr  = str(row[col_rechnr])
            text    = str(row.get(col_text, "")) if col_text in df_31.columns else ""
            soll    = row["_soll"]
            haben   = row["_haben"]
            betrag  = fmt_eur(soll) if soll > 0 else fmt_eur(haben)
            label   = f"**{rechnr}** – {text[:50]} – {betrag}"
            checked = rechnr in st.session_state.extra_31_12
            if st.checkbox(label, value=checked, key=f"cb_31_{rechnr}"):
                auswahl.append(rechnr)

        if auswahl != st.session_state.extra_31_12:
            st.session_state.extra_31_12 = auswahl
            # Neu berechnen mit Auswahl
            with st.spinner("🔄 Salden werden neu berechnet..."):
                st.session_state.ergebnis = abgleichen(
                    st.session_state.df_opos_raw,
                    st.session_state.df_excel_raw,
                    cfg,
                    st.session_state.von_datum,
                    st.session_state.bis_datum,
                    auswahl
                )
            st.rerun()

        if auswahl:
            st.success(f"✓ {len(auswahl)} Buchung(en) zusätzlich einbezogen")

        st.divider()

    # ══════════════════════════════════════════
    #  DIFFERENZANALYSE
    # ══════════════════════════════════════════
    diff = ergebnis["differenz"]
    if abs(diff) >= cfg["toleranz"] and ergebnis["offene_analyse"]:
        st.subheader("🔍 Differenzanalyse")
        st.warning(
            f"Differenz von **{fmt_eur(abs(diff))}** erkannt. "
            f"Das System sucht nach Buchungen die diese Differenz erklären..."
        )

        with st.spinner("Kombinationen werden berechnet..."):
            kombinationen = finde_erklaerende_buchungen(
                diff, ergebnis["offene_analyse"], cfg["toleranz"]
            )

        if kombinationen:
            st.success(
                f"✅ **{len(kombinationen)} mögliche Erklärung(en)** gefunden!"
            )
            for i, kombi in enumerate(kombinationen, 1):
                with st.expander(
                    f"Möglichkeit {i}: "
                    f"{' + '.join(fmt_eur(b['betrag']) for b in kombi)} "
                    f"= {fmt_eur(sum(b['betrag'] for b in kombi))}"
                ):
                    for b in kombi:
                        st.write(
                            f"📌 **{b['rechnr']}** – {b['text']} – "
                            f"{fmt_eur(b['betrag'])}"
                        )
                    st.info(
                        "Diese Buchung(en) erklären die Differenz vollständig. "
                        "Bitte in Excel nachtragen."
                    )
        else:
            st.info(
                "Keine einzelne Buchung oder einfache Kombination gefunden "
                "die die Differenz exakt erklärt. Bitte manuell prüfen."
            )
        st.divider()

    # ══════════════════════════════════════════
    #  TABS: ERGEBNISSE
    # ══════════════════════════════════════════
    tab1, tab2, tab3 = st.tabs([
        f"❌ Fehlende Buchungen ({len(ergebnis['missing'])})",
        f"✅ Abgeglichene Buchungen ({len(ergebnis['matched'])})",
        "📋 Details & Rohdaten"
    ])

    with tab1:
        if ergebnis["missing"].empty:
            st.success("🎉 Alle Buchungen gefunden!")
        else:
            st.warning(f"{len(ergebnis['missing'])} Buchung(en) fehlen in Excel:")
            st.dataframe(ergebnis["missing"], use_container_width=True)
            csv = ergebnis["missing"].to_csv(
                index=False, sep=";"
            ).encode("utf-8-sig")
            st.download_button(
                "⬇️ Fehlende Buchungen exportieren",
                csv, "fehlende_buchungen.csv", "text/csv"
            )

    with tab2:
        if ergebnis["matched"].empty:
            st.info("Keine Buchungen abgeglichen.")
        else:
            st.dataframe(
                ergebnis["matched"].drop(columns=["OK"], errors="ignore"),
                use_container_width=True
            )
            csv = ergebnis["matched"].to_csv(
                index=False, sep=";"
            ).encode("utf-8-sig")
            st.download_button(
                "⬇️ Abgeglichene Buchungen exportieren",
                csv, "abgeglichene_buchungen.csv", "text/csv"
            )

    with tab3:
        st.subheader("📌 Spiegelungslogik")
        if cfg["spiegelung"]:
            st.info(
                "**OPOS Betrag Soll** ↔ **Excel Umsatz Haben**\n\n"
                "**OPOS Betrag Haben** ↔ **Excel Umsatz Soll**"
            )
        else:
            st.info("Spiegelung deaktiviert.")

        with st.expander("🔍 Vorschau OPOS-Rohdaten (erste 10 Zeilen)"):
            st.dataframe(ergebnis["df_opos"].head(10))
        with st.expander("🔍 Vorschau Excel-Rohdaten (erste 10 Zeilen)"):
            st.dataframe(ergebnis["df_excel"].head(10))

elif not opos_file or not excel_file:
    st.info("👆 Bitte beide Dateien hochladen um die Analyse zu starten.")
