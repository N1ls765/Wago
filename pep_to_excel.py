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


# ---------------------------------------------------------------- Tabelle 1-3

def build_produktfamilie(pdf):
    """Seite 5 (PDF-Index 3): Produktserie / Artikelnummer / Masse."""
    page = pdf.pages[3]
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


def build_material(pdf):
    """Seite 6 (PDF-Index 4): Materialzusammensetzung des Referenzprodukts."""
    page = pdf.pages[4]
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


def build_lca_setup(pdf):
    """Seite 9 (PDF-Index 7): Software-/Datenbank-Parameter der Oekobilanz."""
    grid = extract_grid(pdf.pages[7])
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
    """Seiten 14/15: Multiplikationsfaktoren je Variante und Lebenszyklusphase."""
    page = pdf.pages[seite_idx]
    xs, ys = grid_lines(page)
    grid = extract_grid(page, xs, ys)

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

    produktfamilie = build_produktfamilie(pdf)
    material, material_kopf = build_material(pdf)
    setup = build_lca_setup(pdf)

    ind_pflicht, lang_pflicht = build_indikatoren(pdf, 8, "Verbindliche Indikatoren")
    ind_sachbilanz, lang_sachbilanz = build_indikatoren(pdf, 9, "Sachbilanz")
    ind_optional, lang_optional = build_indikatoren(pdf, 10, "Optionale Indikatoren")

    fak_green, lang_green = build_faktoren(pdf, 13, "Green Range")
    fak_std, lang_std = build_faktoren(pdf, 14, "Original Series")

    indikatoren_lang = pd.concat([lang_pflicht, lang_sachbilanz, lang_optional],
                                 ignore_index=True)
    faktoren_lang = pd.concat([lang_green, lang_std], ignore_index=True)

    info = pd.DataFrame([
        ("Quelldatei", Path(pdf_pfad).name),
        ("PEP-Registriernummer", "WAGO-00001-V01.01-EN"),
        ("Produkt", "WAGO 221 Series / Splicing Connector with Levers"),
        ("Referenzprodukt", "221-422"),
        ("Ausgabedatum", "02-2026"),
        ("Gueltigkeit", "5 Jahre"),
        ("Referenzlebensdauer", "30 Jahre bei 70 % Nutzungsrate, 1 A"),
        # Die folgenden Felder liest excel_to_aasx.py fuer Pflichtelemente des
        # Teilmodells. Leer gelassene Felder fallen dort auf Standardwerte
        # zurueck bzw. werden als "im PEP nicht enthalten" berichtet.
        ("Funktionelle Einheit",
         "Connect power transmission cables together or to equipment; "
         "1 A over 30 years at a 70 % use rate, one unit incl. packaging"),
        ("EPD-Ersteller", "WAGO Kontakttechnik GmbH & Co. KG"),
        ("Hersteller-Kontakt", "Alexander Flekler, alexander.flekler@wago.com"),
        ("Programmbetreiber-Website", "www.pep-ecopassport.org"),
        ("Programmbetreiber E-Mail", ""),
        ("Deklarations-URL", ""),
        ("Masse Referenzprodukt",
         material_kopf.get("Reference product mass", "1.60E-03 kg")),
        ("Masse Verpackung je Einheit",
         material_kopf.get("Packaging mass per declared unit", "3.58E-04 kg")),
        ("", ""),
        ("Blatt", "Inhalt (Seite im PDF)"),
        ("Produktfamilie", "Varianten und Massen (S. 5)"),
        ("Materialzusammensetzung", "Werkstoffe des Referenzprodukts (S. 6)"),
        ("LCA_Setup", "Software, Datenbank, Methode (S. 9)"),
        ("Ind_verbindlich", "Verbindliche Wirkungsindikatoren (S. 9)"),
        ("Ind_Sachbilanz", "Sachbilanz-/Ressourcenindikatoren (S. 10)"),
        ("Ind_optional", "Optionale Wirkungsindikatoren (S. 11)"),
        ("Faktoren_GreenRange", "Multiplikationsfaktoren Green Range (S. 14)"),
        ("Faktoren_Original", "Multiplikationsfaktoren Originalserie (S. 15)"),
        ("Indikatoren_lang", "Alle Indikatoren im Langformat - fuer Pivot"),
        ("Faktoren_lang", "Alle Faktoren im Langformat - fuer Pivot"),
        ("", ""),
        ("Hinweis Faktoren",
         "Wirkung einer Variante = Wert des Referenzprodukts 221-422 x Faktor"),
        ("Hinweis Nutzungsphase",
         "B6-Ergebnisse mit dem Quadrat des tatsaechlichen Stroms skalieren"),
    ], columns=["Feld", "Wert"])

    with pd.ExcelWriter(xlsx_pfad, engine="openpyxl") as writer:
        schreibe(writer, info, "Info", zahl_ab=99)
        schreibe(writer, produktfamilie, "Produktfamilie", "0.000", zahl_ab=2)
        schreibe(writer, material, "Materialzusammensetzung", "0,00E+00", zahl_ab=2)
        schreibe(writer, setup, "LCA_Setup", zahl_ab=99)
        schreibe(writer, ind_pflicht, "Ind_verbindlich", zahl_ab=4)
        schreibe(writer, ind_sachbilanz, "Ind_Sachbilanz", zahl_ab=4)
        schreibe(writer, ind_optional, "Ind_optional", zahl_ab=4)
        schreibe(writer, fak_green, "Faktoren_GreenRange", "0.00", zahl_ab=2)
        schreibe(writer, fak_std, "Faktoren_Original", "0.00", zahl_ab=2)
        schreibe(writer, indikatoren_lang, "Indikatoren_lang", zahl_ab=5)
        schreibe(writer, faktoren_lang, "Faktoren_lang", "0.00", zahl_ab=4)

    # Kontrollausgabe: so faellt sofort auf, wenn die Erkennung verrutscht
    print(f"geschrieben: {xlsx_pfad}")
    for name, df in [("Produktfamilie", produktfamilie),
                     ("Materialzusammensetzung", material),
                     ("LCA_Setup", setup),
                     ("Ind_verbindlich", ind_pflicht),
                     ("Ind_Sachbilanz", ind_sachbilanz),
                     ("Ind_optional", ind_optional),
                     ("Faktoren_GreenRange", fak_green),
                     ("Faktoren_Original", fak_std),
                     ("Indikatoren_lang", indikatoren_lang),
                     ("Faktoren_lang", faktoren_lang)]:
        print(f"  {name:26s} {len(df):4d} Zeilen x {len(df.columns):3d} Spalten")
    # Varianten mit Faktoren, aber ohne Massenangabe im PEP (221-2421 steht
    # nur in der Faktorentabelle S. 14) - fuer die spaetere AASX-Erzeugung
    # fehlt damit die deklarierte Masse.
    nur_faktoren = sorted(set(faktoren_lang["Variante"]) - set(produktfamilie["Artikelnummer"]))
    if nur_faktoren:
        protokoll.append("Faktoren ohne Produktzeile (keine Masse im PEP): "
                         + ", ".join(nur_faktoren))
    if protokoll:
        print("\nWarnungen:")
        for w in protokoll:
            print("  ", w)


if __name__ == "__main__":
    quelle = sys.argv[1] if len(sys.argv) > 1 else "eingabe.pdf"
    ziel = sys.argv[2] if len(sys.argv) > 2 else Path(quelle).with_suffix(".xlsx").name
    main(quelle, ziel)
