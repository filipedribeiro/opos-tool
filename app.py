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
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return str(s).strip().upper().replace(" ", "")


def fuzzy_match(r1, r2, threshold):
    r1 = normalize_rechnr(r1)
    r2 = normalize_rechnr(r2)
    if not r1 or not r2:
        return False
    return fuzz.ratio(r1, r2) >= threshold


def fmt_eur(val):
    try:
        f = float(val)
        if f == 0.0:
            return "0,00 €"
        return f"{f:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)


def lade_excel(uploaded_file):
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


def opos_soll_haben(betrag, kennzeichen):
    kz = str(kennzeichen).strip().upper() if kennzeichen else ""
    b  = abs(to_float(betrag))
    if kz.startswith("S"):
        return b, 0.0
    elif kz.startswith("H"):
        return 0.0, b
    else:
        raw = to_float(betrag)
        if raw >= 0:
            return raw, 0.0
        else:
            return 0.0, abs(raw)


def finde_kombination(ziel_betrag, buchungen, tol=0.01, max_results=5):
    """
    Sucht Kombinationen die einen Zielbetrag ergeben.
    Prüft zuerst ob alle Buchungen zusammen die Differenz erklären,
    dann 1er bis 4er Kombinationen.
    """
    if abs(ziel_betrag) < tol:
        return []

    buch = buchungen[:500]
    ergebnisse = []

    # Zuerst: Summe ALLER Buchungen prüfen
    summe_alle = sum(b["betrag"] for b in buch)
    if abs(summe_alle - abs(ziel_betrag)) <= tol:
        return [buch]  # Alle zusammen erklären die Differenz

    # Dann: 1er bis 4er Kombinationen
    for anzahl in range(1, 5):
        for kombi in combinations(range(len(buch)), anzahl):
            summe = sum(buch[i]["betrag"] for i in kombi)
            if abs(summe - abs(ziel_betrag)) <= tol:
                ergebnisse.append([buch[i] for i in kombi])
            if len(ergebnisse) >= max_results:
                return ergebnisse

    return ergebnisse


def finde_erklaerende_buchungen_getrennt(diff_a, diff_b, missing_raw, tol=0.01):
    """
    Sucht getrennt nach Soll- und Haben-Buchungen.
    diff_a = Soll OPOS - Haben Excel → fehlende S-Buchungen
    diff_b = Haben OPOS - Soll Excel → fehlende H-Buchungen
    """
    soll_buchungen  = []
    haben_buchungen = []

    for m in missing_raw:
        soll   = m["_soll"]
        haben  = m["_haben"]
        rechnr = str(m["Rechnungsnr. (OPOS)"])
        text   = m["Buchungstext"]
        datum  = str(m["Datum"])[:10]

        if soll > 0:
            soll_buchungen.append({
                "rechnr": rechnr, "text": text, "datum": datum,
                "betrag": soll, "kz": "S"
            })
        if haben > 0:
            haben_buchungen.append({
                "rechnr": rechnr, "text": text, "datum": datum,
                "betrag": haben, "kz": "H"
            })

    kombi_soll  = finde_kombination(diff_a, soll_buchungen,  tol) if abs(diff_a) > tol else []
    kombi_haben = finde_kombination(diff_b, haben_buchungen, tol) if abs(diff_b) > tol else []

    return kombi_soll, kombi_haben, soll_buchungen, haben_buchungen


def berechne_salden(ergebnis, simulierte_rechnr, cfg):
    """
    Berechnet Salden neu unter Berücksichtigung simulierter Buchungen.
    Simulierte Buchungen = fehlende Buchungen die als 'gebucht' markiert wurden.
    """
    tol        = cfg["toleranz"]
    spiegelung = cfg["spiegelung"]

    # Basis-Salden
    opos_summe_soll  = ergebnis["opos_summe_soll"]
    opos_summe_haben = ergebnis["opos_summe_haben"]
    excel_summe_soll  = ergebnis["excel_summe_soll"]
    excel_summe_haben = ergebnis["excel_summe_haben"]

    # Simulierte Buchungen zur Excel-Seite addieren (gespiegelt!)
    for m in ergebnis["missing_raw"]:
        rechnr = str(m["Rechnungsnr. (OPOS)"])
        if rechnr in simulierte_rechnr:
            soll  = m["_soll"]
            haben = m["_haben"]
            if spiegelung:
                # OPOS Soll → Excel Haben, OPOS Haben → Excel Soll
                excel_summe_haben += soll
                excel_summe_soll  += haben
            else:
                excel_summe_soll  += soll
                excel_summe_haben += haben

    saldo_opos  = opos_summe_soll  - opos_summe_haben
    saldo_excel = -(excel_summe_haben - excel_summe_soll)

    diff_a    = opos_summe_soll  - excel_summe_haben
    diff_b    = opos_summe_haben - excel_summe_soll
    differenz = diff_a - diff_b

    return {
        "opos_summe_soll":   opos_summe_soll,
        "opos_summe_haben":  opos_summe_haben,
        "saldo_opos":        saldo_opos,
        "excel_summe_soll":  excel_summe_soll,
        "excel_summe_haben": excel_summe_haben,
        "saldo_excel":       saldo_excel,
        "diff_a":            diff_a,
        "diff_b":            diff_b,
        "differenz_ab":      differenz,
    }


def abgleichen(df_opos, df_excel, cfg, von_datum, bis_datum,
               opt_von, opt_bis, extra_auswahl=None):
    col_o_rechnr = cfg["opos_rechnr"]
    col_o_datum  = cfg["opos_datum"]
    col_o_saldo  = cfg["opos_saldo"]
    col_o_kz     = cfg["opos_kz"]
    col_o_text   = cfg["opos_text"]

    col_x_rechnr = cfg["excel_rechnr"]
    col_x_datum  = cfg["excel_datum"]
    col_x_soll   = cfg["excel_soll"]
    col_x_haben  = cfg["excel_haben"]

    fuzzy_thresh = cfg["fuzzy"]
    tol          = cfg["toleranz"]
    spiegelung   = cfg["spiegelung"]

    hat_opt = opt_von is not None and opt_bis is not None

    # ── OPOS vorbereiten ──
    df_opos = df_opos.copy()
    df_opos["_datum"]  = df_opos[col_o_datum].apply(parse_datum)
    df_opos["_rechnr"] = df_opos[col_o_rechnr].apply(normalize_rechnr)
    df_opos["_soll"]   = df_opos.apply(
        lambda r: opos_soll_haben(r[col_o_saldo], r[col_o_kz])[0], axis=1
    )
    df_opos["_haben"]  = df_opos.apply(
        lambda r: opos_soll_haben(r[col_o_saldo], r[col_o_kz])[1], axis=1
    )

    def ist_in_opt(d):
        if not hat_opt:
            return False
        return datum_in_bereich(d, opt_von, opt_bis)

    df_opos["_ist_opt"] = df_opos["_datum"].apply(ist_in_opt)
    df_optional = df_opos[df_opos["_ist_opt"]].copy()

    df_opos_filtered = df_opos[
        df_opos["_datum"].apply(
            lambda d: datum_in_bereich(d, von_datum, bis_datum)
        ) & ~df_opos["_ist_opt"]
    ].copy()

    if extra_auswahl is not None and len(extra_auswahl) > 0:
        extra_rows = df_optional[
            df_optional[col_o_rechnr].astype(str).isin(
                [str(r) for r in extra_auswahl]
            )
        ]
        df_opos_filtered = pd.concat(
            [df_opos_filtered, extra_rows], ignore_index=True
        )

    # ── Excel vorbereiten ──
    df_excel = df_excel.copy()
    df_excel["_soll"]   = df_excel[col_x_soll].apply(to_float)
    df_excel["_haben"]  = df_excel[col_x_haben].apply(to_float)
    df_excel["_datum"]  = df_excel[col_x_datum].apply(parse_datum)
    df_excel["_rechnr"] = df_excel[col_x_rechnr].apply(normalize_rechnr)

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
        op_kz = str(op.get(col_o_kz, "")).strip()

        ex_rows = excel_gruppen.get(op_rechnr)
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
                "Kennzeichen":         op_kz,
                "Betrag (Saldo)":      fmt_eur(op[col_o_saldo]),
                "→ Soll":              fmt_eur(op_soll),
                "→ Haben":             fmt_eur(op_haben),
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
            "Kennzeichen":          op_kz,
            "OPOS Betrag (Saldo)":  fmt_eur(op[col_o_saldo]),
            "OPOS → Soll":          fmt_eur(op_soll),
            "OPOS → Haben":         fmt_eur(op_haben),
            "Excel Umsatz Haben":   fmt_eur(ex_sum_haben) if spiegelung else fmt_eur(ex_sum_soll),
            "Excel Umsatz Soll":    fmt_eur(ex_sum_soll)  if spiegelung else fmt_eur(ex_sum_haben),
            "Differenz":            fmt_eur(round(max(diff_soll, diff_haben), 2)),
            "Anzahl Teilbuchungen": anzahl,
            "Typ":                  typ,
            "OK":                   ok
        })

    opos_summe_soll  = df_opos_filtered["_soll"].sum()
    opos_summe_haben = df_opos_filtered["_haben"].sum()
    saldo_opos       = opos_summe_soll - opos_summe_haben

    excel_summe_soll  = df_excel["_soll"].sum()
    excel_summe_haben = df_excel["_haben"].sum()
    saldo_excel       = -(excel_summe_haben - excel_summe_soll)

    diff_a    = opos_summe_soll  - excel_summe_haben
    diff_b    = opos_summe_haben - excel_summe_soll
    differenz = diff_a - diff_b

    offene_fuer_analyse = []
    for m in missing:
        betrag = m["_soll"] if m["_soll"] > 0 else m["_haben"]
        if betrag > 0:
            offene_fuer_analyse.append({
                "rechnr": m["Rechnungsnr. (OPOS)"],
                "datum":  m["Datum"],
                "text":   m["Buchungstext"],
                "betrag": betrag,
                "soll":   m["_soll"],
                "haben":  m["_haben"],
            })

    return {
        "matched":           pd.DataFrame(matched),
        "missing":           pd.DataFrame(missing),
        "missing_raw":       missing,
        "opos_summe_soll":   opos_summe_soll,
        "opos_summe_haben":  opos_summe_haben,
        "saldo_opos":        saldo_opos,
        "excel_summe_soll":  excel_summe_soll,
        "excel_summe_haben": excel_summe_haben,
        "saldo_excel":       saldo_excel,
        "diff_a":            diff_a,
        "diff_b":            diff_b,
        "differenz_ab":      differenz,
        "opos_count":        len(df_opos_filtered),
        "df_opos":           df_opos_filtered,
        "df_excel":          df_excel,
        "df_optional":       df_optional,
        "offene_analyse":    offene_fuer_analyse,
        "col_o_rechnr":      col_o_rechnr,
        "col_o_text":        col_o_text,
        "col_o_saldo":       col_o_saldo,
        "col_o_kz":          col_o_kz,
    }


# ─────────────────────────────────────────────
#  SESSION STATE
# ─────────────────────────────────────────────
for key, val in {
    "ergebnis":        None,
    "extra_auswahl":   [],
    "sim_auswahl":     [],   # simulierte fehlende Buchungen
    "df_opos_raw":     None,
    "df_excel_raw":    None,
    "cfg":             None,
    "_von_datum":      None,
    "_bis_datum":      None,
    "_opt_von":        None,
    "_opt_bis":        None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = val


# ─────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Einstellungen")

    st.subheader("📅 Hauptfilter")
    von_datum = st.date_input(
        "Von Datum (optional)", value=None,
        help="Buchungen VOR diesem Datum werden ignoriert (inklusiv)"
    )
    bis_datum = st.date_input(
        "Bis Datum / Stichtag",
        value=datetime(2026, 4, 30).date(),
        help="Buchungen NACH diesem Datum werden ignoriert (inklusiv)"
    )
    if von_datum and bis_datum and von_datum > bis_datum:
        st.error("⚠️ Von-Datum darf nicht nach Bis-Datum liegen!")

    st.subheader("📅 Optionaler Zeitraum")
    st.caption(
        "Buchungen in diesem Zeitraum werden separat angezeigt "
        "und können einzeln einbezogen werden. Beide Daten inklusiv."
    )
    opt_von = st.date_input(
        "Von (optional)", value=None, key="widget_opt_von",
        help="z.B. 30.12.2025 – inklusiv"
    )
    opt_bis = st.date_input(
        "Bis (optional)", value=None, key="widget_opt_bis",
        help="z.B. 31.12.2025 – inklusiv"
    )
    if opt_von and opt_bis and opt_von > opt_bis:
        st.error("⚠️ Opt. Von darf nicht nach Opt. Bis liegen!")

    st.subheader("📄 Spalten OPOS-Datei")
    opos_rechnr = st.text_input("Rechnungsnummer",   value="Rechnungs-Nr.")
    opos_datum  = st.text_input("Buchungsdatum",     value="Datum")
    opos_text   = st.text_input("Buchungstext",      value="Buchungstext")
    opos_saldo  = st.text_input("Betrag (Saldo)",    value="Saldo")
    opos_kz     = st.text_input("Kennzeichen (S/H)", value="Unnamed: 9")
    st.caption("S = Soll, H = Haben")

    st.subheader("📊 Spalten Buchungs-Excel")
    excel_rechnr = st.text_input("Rechnungsnummer / Buchungsfeld", value="Belegfeld1")
    excel_datum  = st.text_input("Buchungsdatum ",                  value="Datum")
    excel_text   = st.text_input("Buchungstext ",                   value="Buchungstext")
    excel_soll   = st.text_input("Umsatz Soll",                    value="Umsatz Soll")
    excel_haben  = st.text_input("Umsatz Haben",                   value="Umsatz Haben")

    st.subheader("🔄 Abgleichsoptionen")
    spiegelung = st.checkbox(
        "Soll/Haben-Spiegelung aktiv", value=True,
        help="OPOS S ↔ Excel Umsatz Haben"
    )
    fuzzy    = st.slider("Fuzzy-Matching (%)", 60, 100, 85)
    toleranz = st.number_input("Betragstoleranz (€)", value=0.01, step=0.01)


# ─────────────────────────────────────────────
#  UPLOAD
# ─────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.subheader("📄 OPOS-Liste (Excel/CSV)")
    st.caption("Vom Hauptverband")
    opos_file = st.file_uploader(
        "OPOS-Datei hochladen", type=["xlsx", "xls", "csv"], key="opos"
    )
    if opos_file:
        st.success(f"✓ {opos_file.name}")

with col2:
    st.subheader("📊 Gebuchte Buchungen (Excel/CSV)")
    st.caption("Interne Liste Landesverband")
    excel_file = st.file_uploader(
        "Buchungs-Datei hochladen", type=["xlsx", "xls", "csv"], key="excel"
    )
    if excel_file:
        st.success(f"✓ {excel_file.name}")

if von_datum and bis_datum:
    st.info(f"📅 Hauptfilter: **{von_datum.strftime('%d.%m.%Y')}** bis **{bis_datum.strftime('%d.%m.%Y')}** (inklusiv)")
elif bis_datum:
    st.info(f"📅 Stichtag: bis **{bis_datum.strftime('%d.%m.%Y')}** (inklusiv)")
if opt_von and opt_bis:
    st.info(f"📅 Optionaler Zeitraum: **{opt_von.strftime('%d.%m.%Y')}** bis **{opt_bis.strftime('%d.%m.%Y')}** (inklusiv)")

st.divider()


# ─────────────────────────────────────────────
#  ANALYSE STARTEN
# ─────────────────────────────────────────────
datum_ok = (
    not (von_datum and bis_datum and von_datum > bis_datum) and
    not (opt_von and opt_bis and opt_von > opt_bis)
)

if st.button(
    "🚀 Analyse starten", type="primary",
    disabled=not (opos_file and excel_file and datum_ok)
):
    cfg = {
        "opos_rechnr": opos_rechnr, "opos_datum": opos_datum,
        "opos_text": opos_text, "opos_saldo": opos_saldo, "opos_kz": opos_kz,
        "excel_rechnr": excel_rechnr, "excel_datum": excel_datum,
        "excel_text": excel_text, "excel_soll": excel_soll,
        "excel_haben": excel_haben, "spiegelung": spiegelung,
        "fuzzy": fuzzy, "toleranz": toleranz,
    }

    with st.spinner("📄 OPOS wird eingelesen..."):
        df_opos, err = lade_excel(opos_file)
    if err:
        st.error(f"Fehler OPOS: {err}"); st.stop()

    fehlende = [c for c in [opos_rechnr, opos_datum, opos_saldo, opos_kz]
                if c not in df_opos.columns]
    if fehlende:
        st.error(f"Spalten nicht gefunden in OPOS: {fehlende}")
        with st.expander("📋 Gefundene Spalten"):
            st.write(list(df_opos.columns))
        st.stop()

    with st.spinner("📊 Excel wird eingelesen..."):
        df_excel, err = lade_excel(excel_file)
    if err:
        st.error(f"Fehler Excel: {err}"); st.stop()

    fehlende = [c for c in [excel_rechnr, excel_datum, excel_soll, excel_haben]
                if c not in df_excel.columns]
    if fehlende:
        st.error(f"Spalten nicht gefunden in Excel: {fehlende}")
        with st.expander("📋 Gefundene Spalten"):
            st.write(list(df_excel.columns))
        st.stop()

    st.session_state["df_opos_raw"]   = df_opos
    st.session_state["df_excel_raw"]  = df_excel
    st.session_state["cfg"]           = cfg
    st.session_state["_von_datum"]    = von_datum
    st.session_state["_bis_datum"]    = bis_datum
    st.session_state["_opt_von"]      = opt_von
    st.session_state["_opt_bis"]      = opt_bis
    st.session_state["extra_auswahl"] = []
    st.session_state["sim_auswahl"]   = []

    with st.spinner("🔄 Abgleich läuft..."):
        st.session_state["ergebnis"] = abgleichen(
            df_opos, df_excel, cfg,
            von_datum, bis_datum,
            opt_von, opt_bis, []
        )


# ─────────────────────────────────────────────
#  ERGEBNISSE
# ─────────────────────────────────────────────
if st.session_state["ergebnis"]:
    ergebnis = st.session_state["ergebnis"]
    cfg      = st.session_state["cfg"]

    # Salden mit simulierten Buchungen berechnen
    sim_auswahl = st.session_state["sim_auswahl"]
    salden = berechne_salden(ergebnis, set(sim_auswahl), cfg)

    st.success("✅ Analyse abgeschlossen!")

    # ══════════════════════════════════════════
    #  SALDENÜBERSICHT
    # ══════════════════════════════════════════
    st.subheader("📊 Saldenübersicht")

    if sim_auswahl:
        st.info(
            f"📌 Simulation aktiv: **{len(sim_auswahl)} fehlende Buchung(en)** "
            f"werden als gebucht behandelt."
        )

    col_o, col_x = st.columns(2)
    with col_o:
        st.markdown("**OPOS – Hauptverband**")
        c1, c2, c3 = st.columns(3)
        c1.metric("Summe Soll",  fmt_eur(salden["opos_summe_soll"]))
        c2.metric("Summe Haben", fmt_eur(salden["opos_summe_haben"]))
        c3.metric("Saldo",       fmt_eur(salden["saldo_opos"]))

    with col_x:
        st.markdown("**Excel – Landesverband**")
        c1, c2, c3 = st.columns(3)
        c1.metric("Summe Soll",  fmt_eur(salden["excel_summe_soll"]))
        c2.metric("Summe Haben", fmt_eur(salden["excel_summe_haben"]))
        c3.metric("Saldo",       fmt_eur(salden["saldo_excel"]))

    st.divider()

    st.markdown("**🔄 Spiegelungsabgleich**")
    st.caption("Soll OPOS entspricht Haben Excel · Haben OPOS entspricht Soll Excel")

    ca, cb, cdiff = st.columns(3)
    diff_a    = salden["diff_a"]
    diff_b    = salden["diff_b"]
    differenz = salden["differenz_ab"]

    ca.metric("(a) Soll OPOS − Haben Excel", fmt_eur(diff_a))
    cb.metric("(b) Haben OPOS − Soll Excel", fmt_eur(diff_b))

    with cdiff:
        if abs(differenz) < cfg["toleranz"]:
            st.success("✅ Vollständig abgestimmt")
            st.metric("Differenz (a) − (b)", "0,00 €")
        else:
            richtung = "OPOS höher" if differenz > 0 else "Excel höher"
            st.warning(f"⚠️ {richtung}")
            st.metric("Differenz (a) − (b)", fmt_eur(differenz))

    cm, cf = st.columns(2)
    cm.metric("Abgeglichen", len(ergebnis["matched"]))
    cf.metric("Fehlend",     len(ergebnis["missing_raw"]))

    st.divider()

    # ══════════════════════════════════════════
    #  OPTIONALE BUCHUNGEN (Zeitraum)
    # ══════════════════════════════════════════
    df_opt     = ergebnis["df_optional"]
    col_rechnr = ergebnis["col_o_rechnr"]
    col_text   = ergebnis["col_o_text"]
    col_saldo  = ergebnis["col_o_saldo"]
    col_kz     = ergebnis["col_o_kz"]

    if len(df_opt) > 0:
        _opt_von = st.session_state["_opt_von"]
        _opt_bis = st.session_state["_opt_bis"]
        zeitraum_str = ""
        if _opt_von and _opt_bis:
            zeitraum_str = (
                f" ({_opt_von.strftime('%d.%m.%Y')} – "
                f"{_opt_bis.strftime('%d.%m.%Y')}, inklusiv)"
            )

        st.subheader("📅 Optionale Buchungen")
        st.info(
            f"**{len(df_opt)} Buchungen** im optionalen Zeitraum"
            f"{zeitraum_str}. Einzeln einbeziehen:"
        )

        auswahl = []
        for _, row in df_opt.iterrows():
            rechnr  = str(row[col_rechnr])
            text    = str(row.get(col_text, "")) if col_text in df_opt.columns else ""
            soll    = row["_soll"]
            haben   = row["_haben"]
            betrag  = fmt_eur(soll) if soll > 0 else fmt_eur(haben)
            kz      = str(row.get(col_kz, "")).strip()
            d       = row["_datum"]
            datum_s = d.strftime("%d.%m.%Y") if d else "?"
            label   = f"**{rechnr}** · {datum_s} · {text[:45]} · **{betrag}** ({kz})"
            checked = rechnr in st.session_state["extra_auswahl"]
            if st.checkbox(label, value=checked, key=f"cb_opt_{rechnr}"):
                auswahl.append(rechnr)

        if sorted(auswahl) != sorted(st.session_state["extra_auswahl"]):
            st.session_state["extra_auswahl"] = auswahl
            with st.spinner("🔄 Wird neu berechnet..."):
                st.session_state["ergebnis"] = abgleichen(
                    st.session_state["df_opos_raw"],
                    st.session_state["df_excel_raw"],
                    cfg,
                    st.session_state["_von_datum"],
                    st.session_state["_bis_datum"],
                    st.session_state["_opt_von"],
                    st.session_state["_opt_bis"],
                    auswahl
                )
                st.session_state["sim_auswahl"] = []
            st.rerun()

        if auswahl:
            st.success(f"✓ {len(auswahl)} von {len(df_opt)} Buchung(en) einbezogen")

        st.divider()

    # ══════════════════════════════════════════
    #  DIFFERENZANALYSE (getrennt Soll/Haben)
    # ══════════════════════════════════════════
    if abs(differenz) >= cfg["toleranz"] and ergebnis["missing_raw"]:
        st.subheader("🔍 Differenzanalyse")
        st.warning(
            f"Differenz von **{fmt_eur(abs(differenz))}** erkannt. "
            f"Suche getrennt nach Soll- und Haben-Buchungen..."
        )

        with st.spinner("Kombinationen werden berechnet..."):
            kombi_soll, kombi_haben, soll_buch, haben_buch = (
                finde_erklaerende_buchungen_getrennt(
                    diff_a, diff_b,
                    ergebnis["missing_raw"],
                    cfg["toleranz"]
                )
            )

        col_s, col_h = st.columns(2)

        with col_s:
            st.markdown(f"**Soll-Seite** · (a) = {fmt_eur(diff_a)}")
            if abs(diff_a) < cfg["toleranz"]:
                st.success("✅ Soll-Seite bereits ausgeglichen")
            elif kombi_soll:
                st.success(f"✅ {len(kombi_soll)} Möglichkeit(en) gefunden")
                for i, kombi in enumerate(kombi_soll, 1):
                    summe = sum(b["betrag"] for b in kombi)
                    with st.expander(
                        f"S-Möglichkeit {i}: "
                        f"{' + '.join(fmt_eur(b['betrag']) for b in kombi)} "
                        f"= {fmt_eur(summe)}"
                    ):
                        for b in kombi:
                            st.write(
                                f"📌 **{b['rechnr']}** · {b['datum']} · "
                                f"{b['text']} · {fmt_eur(b['betrag'])} (S)"
                            )
                        st.info("Diese S-Buchung(en) gleichen die Soll-Differenz aus.")
            else:
                st.info(f"Keine Kombination aus {len(soll_buch)} S-Buchungen gefunden.")

        with col_h:
            st.markdown(f"**Haben-Seite** · (b) = {fmt_eur(diff_b)}")
            if abs(diff_b) < cfg["toleranz"]:
                st.success("✅ Haben-Seite bereits ausgeglichen")
            elif kombi_haben:
                st.success(f"✅ {len(kombi_haben)} Möglichkeit(en) gefunden")
                for i, kombi in enumerate(kombi_haben, 1):
                    summe = sum(b["betrag"] for b in kombi)
                    with st.expander(
                        f"H-Möglichkeit {i}: "
                        f"{' + '.join(fmt_eur(b['betrag']) for b in kombi)} "
                        f"= {fmt_eur(summe)}"
                    ):
                        for b in kombi:
                            st.write(
                                f"📌 **{b['rechnr']}** · {b['datum']} · "
                                f"{b['text']} · {fmt_eur(b['betrag'])} (H)"
                            )
                        st.info("Diese H-Buchung(en) gleichen die Haben-Differenz aus.")
            else:
                st.info(f"Keine Kombination aus {len(haben_buch)} H-Buchungen gefunden.")

        st.divider()

    # ══════════════════════════════════════════
    #  TABS
    # ══════════════════════════════════════════
    tab1, tab2, tab3 = st.tabs([
        f"❌ Fehlende Buchungen ({len(ergebnis['missing_raw'])})",
        f"✅ Abgeglichene Buchungen ({len(ergebnis['matched'])})",
        "📋 Details & Rohdaten"
    ])

    with tab1:
        missing_raw = ergebnis["missing_raw"]
        if not missing_raw:
            st.success("🎉 Alle Buchungen gefunden!")
        else:
            st.warning(f"{len(missing_raw)} Buchung(en) fehlen in Excel:")
            st.caption(
                "Simuliere das Nachtragen einzelner Buchungen – "
                "die Salden werden sofort aktualisiert."
            )

            sim_neu = []
            for m in missing_raw:
                rechnr  = str(m["Rechnungsnr. (OPOS)"])
                datum   = str(m["Datum"])
                text    = m["Buchungstext"]
                kz      = m["Kennzeichen"]
                betrag  = m["Betrag (Saldo)"]
                soll    = fmt_eur(m["_soll"])
                haben   = fmt_eur(m["_haben"])
                label   = (
                    f"**{rechnr}** · {datum[:10]} · "
                    f"{text[:45]} · **{betrag}** ({kz})"
                )
                checked = rechnr in st.session_state["sim_auswahl"]
                if st.checkbox(label, value=checked, key=f"cb_sim_{rechnr}"):
                    sim_neu.append(rechnr)

            if sorted(sim_neu) != sorted(st.session_state["sim_auswahl"]):
                st.session_state["sim_auswahl"] = sim_neu
                st.rerun()

            if sim_neu:
                st.success(
                    f"✓ **{len(sim_neu)} Buchung(en)** simuliert als gebucht – "
                    f"Salden oben wurden aktualisiert."
                )

            # Export (ohne simulierte Buchungen – nur echte Fehlende)
            df_missing_export = ergebnis["missing"].drop(
                columns=["_soll", "_haben"], errors="ignore"
            )
            csv = df_missing_export.to_csv(
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
        st.info(
            "**OPOS Kennzeichen S** ↔ **Excel Umsatz Haben**\n\n"
            "**OPOS Kennzeichen H** ↔ **Excel Umsatz Soll**\n\n"
            "Differenz = (Soll OPOS − Haben Excel) − (Haben OPOS − Soll Excel)\n\n"
            "Simulation: Fehlende Buchungen werden gespiegelt zur Excel-Seite addiert"
        )
        with st.expander("🔍 Vorschau OPOS (erste 10 Zeilen)"):
            st.dataframe(ergebnis["df_opos"].head(10))
        with st.expander("🔍 Vorschau Excel (erste 10 Zeilen)"):
            st.dataframe(ergebnis["df_excel"].head(10))

elif not opos_file or not excel_file:
    st.info("👆 Bitte beide Dateien hochladen um die Analyse zu starten.")
