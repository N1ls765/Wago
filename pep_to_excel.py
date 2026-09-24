#!/usr/bin/env python3
"""Tabellen aus einem PEP-ecopassport-PDF (WAGO) in eine Excel-Datei ueberfuehren.

Aufruf:  python pep_to_excel.py eingabe.pdf [ausgabe.xlsx]

Aufbau der Pipeline
-------------------
1. EXTRAKTION   Raster aus gezeichneten Linien/Rechtecken -> Zellen (pdf_grid.py)
2. AUFBEREITUNG Kopfzeilen zusammenfuehren, Abschnittszeilen erkennen,
                Zahlen mit Dezimalkomma in echte Floats wandeln
3. EXPORT       ein Blatt je Tabelle plus zwei Langformat-Blaetter fuer Pivots

Alle Seitenzahlen und Spaltenzuordnungen sind in TABELLEN bzw. in den
build_*-Funktionen gebuendelt. Aendert sich das PDF-Layout, muss nur dort
nachgezogen werden.
"""

import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from pdf_grid import cell_text, extract_grid, grid_lines, open_pdf


# ---------------------------------------------------------------- Definitionen fuer Tests
import pandas as pd

def validate_material_sum(df_material):
    """Prüft, ob die Summe der Materialanteile ca. 100 % ergibt."""
    # Schutz vor leeren oder unvollständigen DataFrames
    if df_material.empty or "Kategorie PEP Material" not in df_material.columns or "Anteil (%)" not in df_material.columns:
        return False, 0

    df_calc = df_material[df_material["Kategorie PEP Material"] != "Total Weight"]
    total_sum = df_calc["Anteil (%)"].sum()

    is_valid = 99.8 <= total_sum <= 100.2
    return is_valid, round(total_sum, 2)

def validate_mandatory_values(df_indikatoren):
    """Prüft, ob alle GWP-Werte valide Zahlen sind und nicht fehlen."""
    # Schutz vor leeren DataFrames
    if df_indikatoren.empty or "Indikator" not in df_indikatoren.columns:
        return False, "Indikator-Tabelle ist leer oder unvollständig."

    gwp_df = df_indikatoren[df_indikatoren["Indikator"].str.contains("Climate change", na=False)]
    if gwp_df.empty:
        return False, "Keine Climate Change Indikatoren (GWP) gefunden."

    phasen = ["A1-3", "A4", "A5", "B1-7", "C1-4", "Total"]
    # Prüfen, ob alle geforderten Phasen-Spalten überhaupt extrahiert wurden
    fehlende_spalten = [p for p in phasen if p not in gwp_df.columns]
    if fehlende_spalten:
        return False, f"Phasen-Spalten wurden vom Raster nicht erkannt: {fehlende_spalten}"

    missing_values = gwp_df[phasen].isna().sum().sum()
    is_valid = missing_values == 0
    return is_valid, f"{missing_values} fehlende Werte in GWP-Indikatoren."

def validate_parsing_errors(protokoll_liste):
    """Prüft das Fehler-Protokoll aus pep_to_excel.py auf Konvertierungsfehler."""
    is_valid = len(protokoll_liste) == 0
    return is_valid, protokoll_liste

# ---------------------------------------------------------------- Hilfsfunktionen

ZAHL = re.compile(r"^-?\d{1,3}(?:\.\d{3})*(?:,\d+)?(?:[Ee][+-]?\d+)?$|^-?\d+(?:,\d+)?(?:[Ee][+-]?\d+)?$")

protokoll = []


def zu_zahl(text):
    """'1,24E-02' -> 1.24e-2. Gibt den Originaltext zurueck, wenn keine Zahl.

    Deutsches Zahlenformat: Komma = Dezimaltrenner, Punkt = Tausendertrenner.
    Nicht konvertierbare Werte werden protokolliert statt still verschluckt.
    """
    if text is None:
        return None
    s = str(text).strip()
    if s == "":
        return None
    if not ZAHL.match(s):
        return s
    try:
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        protokoll.append(f"nicht konvertierbar: {s!r}")
        return s


def zeilentext(page, xs, ys, zeile):
    """Eine Rasterzeile als durchgehender Text - fuer verbundene Zellen."""
    chars = [c for c in page.chars
             if ys[zeile] <= (c["top"] + c["bottom"]) / 2 < ys[zeile + 1]
             and xs[0] <= (c["x0"] + c["x1"]) / 2 < xs[-1]]
    return cell_text(chars)


def ist_abschnitt(row):
    """Zeile ohne jeden Zahlenwert in den Datenspalten = Zwischenueberschrift."""
    return not any(ZAHL.match(c.strip()) for c in row[1:] if c.strip())

def finde_seite(pdf, suchtext):
    """Durchsucht alle Seiten nach einem Textmarker und gibt den Index zurueck."""
    for i, page in enumerate(pdf.pages):
        text = page.extract_text()
        if text and suchtext in text:
            return i
    raise ValueError(f"Suchtext '{suchtext}' nicht im PDF gefunden!")

# ---------------------------------------------------------------- Tabelle 1-3

def build_produktfamilie(pdf, seite_idx):
    """Produktserie / Artikelnummer / Masse."""
    page = pdf.pages[seite_idx]
    xs, ys = grid_lines(page, min_len=5)
    grid = extract_grid(page, xs, ys)

    # Der Seriename steht in einer ueber mehrere Zeilen verbundenen Zelle.
    # Deren Hoehe ergibt sich aus den breiten Trennlinien der linken Spalte -
    # ueber die Zeilenreihenfolge allein waere die Zuordnung nicht sicher.
    baender = sorted({round(r["top"], 1) for r in page.rects
                      if r["height"] < 2 and r["x0"] < 60 and r["width"] > 100})
    beschriftung = {}
    for oben, unten in zip(baender, baender[1:]):
        text = cell_text([c for c in page.chars
                          if oben < (c["top"] + c["bottom"]) / 2 < unten
                          and xs[0] <= c["x0"] < xs[6]])
        if text.startswith("WAGO"):
            beschriftung[(oben, unten)] = text.replace(" ,", ",")

    def serie_fuer(y):
        for (oben, unten), text in beschriftung.items():
            if oben < y < unten:
                return text
        return None

    rows = []
    for i, r in enumerate(grid):
        artikel, masse = r[6].strip(), r[9].strip()
        if artikel.startswith("221-"):
            mitte = (ys[i] + ys[i + 1]) / 2
            # round: der Float zu "9.376" wuerde sonst als 9.375999... im
            # Blatt stehen; das PEP gibt drei Nachkommastellen an.
            rows.append({"Produktserie": serie_fuer(mitte), "Artikelnummer": artikel,
                         "Masse deklarierte Einheit (g)": round(zu_zahl(masse.replace(".", ",")), 3)})
    return pd.DataFrame(rows)


def build_material(pdf, seite_idx):
    """Materialzusammensetzung des Referenzprodukts."""
    page = pdf.pages[seite_idx]
    xs, ys = grid_lines(page, min_len=5)
    grid = extract_grid(page, xs, ys)

    rows, kopf, im_koerper = [], {}, False
    for r in grid:
        label = " ".join(x for x in r[1:4] if x).strip()
        material, masse, anteil = r[4].strip(), r[5].strip(), r[6].strip()
        if label.startswith(("Reference product mass", "Packaging mass")):
            kopf[label] = material
        # Tabellenkoerper beginnt nach der Kopfzeile und endet mit "Total Weight";
        # davor steht die Abschnittsueberschrift, danach die Legende des Kreisdiagramms.
        if label.startswith("Category PEP Material"):
            im_koerper = True
            continue
        if not im_koerper:
            continue
        if material and masse:
            rows.append({"Kategorie PEP Material": label, "Material": material,
                         "Masse (kg)": zu_zahl(masse.replace(".", ",")),
                         "Anteil (%)": zu_zahl(anteil.replace("%", "").strip())})
        elif label.startswith("Total"):
            rows.append({"Kategorie PEP Material": "Total Weight", "Material": "",
                         "Masse (kg)": zu_zahl(masse.replace(".", ",")),
                         "Anteil (%)": zu_zahl(anteil.replace("%", "").strip())})
            break
    return pd.DataFrame(rows), kopf


def build_lca_setup(pdf, seite_idx):
    """Software-/Datenbank-Parameter der Oekobilanz."""
    grid = extract_grid(pdf.pages[seite_idx])
    rows = [{"Parameter": r[0].strip(), "Beschreibung": r[1].strip()}
            for r in grid[1:] if r[0].strip()]
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- Indikatoren

PHASEN = ["A1-3", "A4", "A5", "B1-7", "C1-4", "Total", "D"]


def build_indikatoren(pdf, seite_idx, tabellenname):
    """Seiten 9/10/11: Indikatortabellen mit sieben Lebenszyklusspalten.

    Liefert (Breitformat, Langformat). Zwischenueberschriften der Seite
    werden als Spalte 'Gruppe' mitgefuehrt, statt sie wegzuwerfen.
    """
    page = pdf.pages[seite_idx]
    xs, ys = grid_lines(page)
    grid = extract_grid(page, xs, ys)

    kopf = next(i for i, r in enumerate(grid) if r[0].strip() == "Impact category")
    gruppe = zeilentext(page, xs, ys, kopf - 1)

    breit, lang = [], []
    for i in range(kopf + 1, len(grid)):
        r = grid[i]
        if ist_abschnitt(r):
            gruppe = zeilentext(page, xs, ys, i)
            continue

        indikator = r[0].strip()
        einheit = r[1].strip()
        werte = [zu_zahl(c) for c in r[2:9]]

        # Verbundene Zelle: ein einziger Wert gilt fuer die ganze Zeile
        gefuellt = [(p, v) for p, v in zip(PHASEN, werte) if v is not None]
        verbunden = len(gefuellt) == 1 and gefuellt[0][0] not in ("Total",)

        eintrag = {"Tabelle": tabellenname, "Gruppe": gruppe,
                   "Indikator": indikator, "Einheit": einheit}
        if verbunden:
            eintrag.update({p: None for p in PHASEN})
            eintrag["Total"] = gefuellt[0][1]
            eintrag["Hinweis"] = "Einzelwert (verbundene Zelle im PDF)"
            lang.append({"Tabelle": tabellenname, "Gruppe": gruppe,
                         "Indikator": indikator, "Einheit": einheit,
                         "Phase": "Total", "Wert": gefuellt[0][1]})
        else:
            eintrag.update(dict(zip(PHASEN, werte)))
            eintrag["Hinweis"] = ""
            for p, v in zip(PHASEN, werte):
                if v is not None:
                    lang.append({"Tabelle": tabellenname, "Gruppe": gruppe,
                                 "Indikator": indikator, "Einheit": einheit,
                                 "Phase": p, "Wert": v})
        breit.append(eintrag)

    return pd.DataFrame(breit), pd.DataFrame(lang)


# ---------------------------------------------------------------- Faktoren

def gruppenlabels(page, xs, ys, zeile=0):
    """Gruppenueberschriften der Faktorentabellen aus Zeichenbloecken lesen.

    Die Ueberschriften stehen in verbundenen Zellen und werden vom Raster
    zerschnitten. Daher Zeichen der Kopfzeile einsammeln und anhand der
    x-Luecken wieder zu Bloecken gruppieren.
    """
    chars = sorted((c for c in page.chars
                    if ys[zeile] <= (c["top"] + c["bottom"]) / 2 < ys[zeile + 1]
                    and c["x0"] > xs[1] - 5),
                   key=lambda c: c["x0"])
    bloecke, aktuell = [], [chars[0]]
    for a, b in zip(chars, chars[1:]):
        if b["x0"] - a["x1"] > 3:
            bloecke.append(aktuell)
            aktuell = []
        aktuell.append(b)
    bloecke.append(aktuell)
    return ["".join(c["text"] for c in bl).strip() for bl in bloecke]


def build_faktoren(pdf, seite_idx, serie):
    """Multiplikationsfaktoren je Variante und Lebenszyklusphase."""
    page = pdf.pages[seite_idx]
    xs, ys = grid_lines(page)

    if len(xs) < 2 or len(ys) < 2:
        return pd.DataFrame(), pd.DataFrame()

    grid = extract_grid(page, xs, ys)
    if len(grid) < 3:
        return pd.DataFrame(), pd.DataFrame()

    gruppen = gruppenlabels(page, xs, ys, 0)
    varianten_zeile = [c.strip() for c in grid[1][1:]]
    n = len(varianten_zeile) // len(gruppen)          # Varianten je Gruppe
    varianten = varianten_zeile[:n]

    spalten = [f"{g} | {v}" for g in gruppen for v in varianten]

    breit, lang = [], []
    for r in grid[2:]:
        indikator = r[0].strip()
        if not indikator:
            continue
        werte = [zu_zahl(c) for c in r[1:1 + len(spalten)]]
        breit.append({"Serie": serie, "Indikator": indikator,
                      **dict(zip(spalten, werte))})
        for (g, v), w in zip([(g, v) for g in gruppen for v in varianten], werte):
            if w is not None:
                lang.append({"Serie": serie, "Indikator": indikator,
                             "Phase": g, "Variante": f"221-{v}", "Faktor": w})
    return pd.DataFrame(breit), pd.DataFrame(lang)


# ---------------------------------------------------------------- Excel-Export

KOPF_FILL = PatternFill("solid", fgColor="6EA644")     # WAGO-Gruen
KOPF_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
BASIS_FONT = Font(name="Arial", size=10)
RAHMEN = Side(style="thin", color="BFBFBF")


def schreibe(writer, df, blatt, zahlenformat="0,00E+00", zahl_ab=0):
    """DataFrame als formatiertes Blatt ablegen."""
    df.to_excel(writer, sheet_name=blatt[:31], index=False)
    ws = writer.sheets[blatt[:31]]

    for zelle in ws[1]:
        zelle.fill, zelle.font = KOPF_FILL, KOPF_FONT
        zelle.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 34
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for j, spalte in enumerate(df.columns, start=1):
        breite = max(df[spalte].astype(str).str.len().max() if len(df) else 0,
                     len(str(spalte)))
        ws.column_dimensions[get_column_letter(j)].width = min(max(breite + 2, 10), 46)
        for i in range(2, len(df) + 2):
            z = ws.cell(row=i, column=j)
            z.font = BASIS_FONT
            z.border = Border(bottom=RAHMEN)
            if j > zahl_ab and isinstance(z.value, float):
                z.number_format = zahlenformat


def main(pdf_pfad, xlsx_pfad):
    pdf = open_pdf(pdf_pfad)

    # Dynamische Seitensuche anhand von markanten, kurzen Pflicht-Wörtern
    idx_prod = finde_seite(pdf, "Product series")
    idx_mat = finde_seite(pdf, "Constituent Materials")
    idx_setup = finde_seite(pdf, "LCA Software")

    # Die EPD-Norm schreibt die Reihenfolge der LCIA-Tabellen fest vor.
    # Daher reicht es, den Startpunkt zu finden und relativ weiterzugehen.
    idx_pflicht = finde_seite(pdf, "Mandatory impact indicators")
    idx_sachbilanz = idx_pflicht + 1
    idx_optional = idx_pflicht + 2

    # Faktoren können über mehrere Seiten gehen (Start bis Disclaimer)
    idx_fak_start = finde_seite(pdf, "Extrapolation")
    idx_disclaimer = finde_seite(pdf, "Disclaimer")

    # Tabellen mit den dynamisch gefundenen Indizes extrahieren
    produktfamilie = build_produktfamilie(pdf, idx_prod)
    material, material_kopf = build_material(pdf, idx_mat)
    setup = build_lca_setup(pdf, idx_setup)

    ind_pflicht, lang_pflicht = build_indikatoren(pdf, idx_pflicht, "Verbindliche Indikatoren")
    ind_sachbilanz, lang_sachbilanz = build_indikatoren(pdf, idx_sachbilanz, "Sachbilanz")
    ind_optional, lang_optional = build_indikatoren(pdf, idx_optional, "Optionale Indikatoren")

    # Alle Faktorentabellen zwischen dem Kapitel-Start und dem Disclaimer einsammeln
    fak_breit_list, fak_lang_list = [], []
    for i in range(idx_fak_start, idx_disclaimer):
        b, l = build_faktoren(pdf, i, f"Faktoren_Seite_{i+1}")
        if not b.empty:
            fak_breit_list.append(b)
            fak_lang_list.append(l)

    indikatoren_lang = pd.concat([lang_pflicht, lang_sachbilanz, lang_optional], ignore_index=True)
    indikatoren_breit = pd.concat([ind_pflicht, ind_sachbilanz, ind_optional], ignore_index=True)
    faktoren_lang = pd.concat(fak_lang_list, ignore_index=True) if fak_lang_list else pd.DataFrame()

    # === TESTS ===
    global protokoll
    ist_valide, fehler = validate_parsing_errors(protokoll)
    if not ist_valide:
        print(f"KRITISCHER FEHLER: {len(fehler)} Werte konnten nicht in Zahlen umgewandelt werden!")

    mat_valide, mat_summe = validate_material_sum(material)
    if not mat_valide:
        print(f"WARNUNG: Die Materialanteile ergeben {mat_summe}% statt 100%. Da fehlt etwas!")

    gwp_valide, gwp_fehler = validate_mandatory_values(indikatoren_breit)
    if not gwp_valide:
        print(f"WARNUNG: {gwp_fehler}")
    # === ENDE DER TESTS ===

    # Info-Block (bleibt identisch)
    info = pd.DataFrame([
        ("Quelldatei", Path(pdf_pfad).name),
        ("PEP-Registriernummer", "Dynamisch ausgelesen (Platzhalter)"),
        ("Masse Referenzprodukt", material_kopf.get("Reference product mass", "Unbekannt")),
        ("Masse Verpackung je Einheit", material_kopf.get("Packaging mass per declared unit", "Unbekannt")),
    ], columns=["Feld", "Wert"])

    # Dynamischer Excel-Export
    with pd.ExcelWriter(xlsx_pfad, engine="openpyxl") as writer:
        schreibe(writer, info, "Info", zahl_ab=99)
        schreibe(writer, produktfamilie, "Produktfamilie", "0.000", zahl_ab=2)
        schreibe(writer, material, "Materialzusammensetzung", "0,00E+00", zahl_ab=2)
        schreibe(writer, setup, "LCA_Setup", zahl_ab=99)
        schreibe(writer, ind_pflicht, "Ind_verbindlich", zahl_ab=4)
        schreibe(writer, ind_sachbilanz, "Ind_Sachbilanz", zahl_ab=4)
        schreibe(writer, ind_optional, "Ind_optional", zahl_ab=4)

        # Jede gefundene Faktoren-Tabelle auf ein eigenes Blatt schreiben
        for idx, f_df in enumerate(fak_breit_list):
            schreibe(writer, f_df, f"Faktoren_Tab_{idx+1}", "0.00", zahl_ab=2)

        schreibe(writer, indikatoren_lang, "Indikatoren_lang", zahl_ab=5)
        if not faktoren_lang.empty:
            schreibe(writer, faktoren_lang, "Faktoren_lang", "0.00", zahl_ab=4)

    print(f"\nERFOLG! Excel geschrieben: {xlsx_pfad}")

    tabellen_uebersicht = [
        ("Produktfamilie", produktfamilie),
        ("Materialzusammensetzung", material),
        ("LCA_Setup", setup),
        ("Ind_verbindlich", ind_pflicht),
        ("Ind_Sachbilanz", ind_sachbilanz),
        ("Ind_optional", ind_optional)
    ]

    # Dynamische Faktorentabellen zur Übersicht hinzufügen
    for idx, f_df in enumerate(fak_breit_list):
        tabellen_uebersicht.append((f"Faktoren_Tab_{idx+1}", f_df))

    tabellen_uebersicht.extend([
        ("Indikatoren_lang", indikatoren_lang),
        ("Faktoren_lang", faktoren_lang)
    ])

    for name, df in tabellen_uebersicht:
        print(f"  {name:26s} {len(df):4d} Zeilen x {len(df.columns):3d} Spalten")

    # Warnungen für fehlende Massen berechnen und ausgeben
    if not faktoren_lang.empty and not produktfamilie.empty:
        nur_faktoren = sorted(set(faktoren_lang["Variante"]) - set(produktfamilie["Artikelnummer"]))
        if nur_faktoren:
            protokoll.append("Faktoren ohne Produktzeile (keine Masse im PEP): " + ", ".join(nur_faktoren))

    if protokoll:
        print("\nWarnungen:")
        for w in protokoll:
            print("  ", w)


if __name__ == "__main__":
    quelle = sys.argv[1] if len(sys.argv) > 1 else "eingabe.pdf"
    ziel = sys.argv[2] if len(sys.argv) > 2 else Path(quelle).with_suffix(".xlsx").name
    main(quelle, ziel)
