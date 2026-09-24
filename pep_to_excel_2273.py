#!/usr/bin/env python3
"""Tabellen aus dem PEP WAGO-00002 (2273-Serie) in eine Excel-Datei ueberfuehren.

Aufruf:  python pep_to_excel_2273.py WAGO-00002-V01.01-EN.pdf [WAGO_PEP_2273.xlsx]

Gegenstueck zu pep_to_excel.py (PEP WAGO-00001, 221-Serie). Die Pipeline ist
dieselbe:
1. EXTRAKTION   Raster aus gezeichneten Linien/Rechtecken -> Zellen (pdf_grid.py)
2. AUFBEREITUNG Kopfzeilen zusammenfuehren, Abschnittszeilen erkennen,
                Zahlen mit Dezimalkomma in echte Floats wandeln
3. EXPORT       ein Blatt je Tabelle plus zwei Langformat-Blaetter fuer Pivots

Die Indikatortabellen (S. 9-11) und die LCA-Setup-Tabelle (S. 8) sind im
00002-PEP identisch aufgebaut und werden mit den Funktionen aus
pep_to_excel.py gelesen. Abweichend vom 00001-PEP und daher hier neu:
  - Produkt- und Materialtabelle haben andere Spaltenpositionen; die Spalten
    werden deshalb ueber ihre Kopfzeile gesucht statt ueber feste Indizes
  - es gibt nur EINE Faktorentabelle (S. 14) statt zwei (Green Range /
    Original Series), mit vier Varianten 2273-203/-204/-205/-208
  - in der Faktorentabelle stehen die ersten beiden Gruppenueberschriften so
    eng beieinander, dass die Luecken-Heuristik aus pep_to_excel.py sie zu
    einer verschmilzt; die Gruppen werden hier ueber die Spaltenbreite
    (Anzahl Varianten je Gruppe) bestimmt
  - Stammdaten (Registriernummer, Ausgabedatum, Verifizierer ...) werden aus
    dem PDF-Text gelesen statt fest eingetragen

Das Blatt "Info" enthaelt zusaetzlich die Felder, die excel_to_aasx_2273.py
fuer das (neue) EPD-Type-III-Teilmodell braucht.
"""

import re
import sys
from pathlib import Path

import pandas as pd

from pdf_grid import cell_text, extract_grid, grid_lines, open_pdf
from pep_to_excel import build_indikatoren, build_lca_setup, protokoll, schreibe, zu_zahl

SERIE = "2273"

# PDF-Seitenindex (0-basiert) je Tabelle. Aendert sich das Layout, nur hier nachziehen.
SEITE_PRODUKTE = 3        # PDF S. 4
SEITE_MATERIAL = 4        # PDF S. 5
SEITE_SETUP = 7           # PDF S. 8
SEITE_IND_PFLICHT = 8     # PDF S. 9
SEITE_IND_SACHBILANZ = 9  # PDF S. 10
SEITE_IND_OPTIONAL = 10   # PDF S. 11
SEITE_FAKTOREN = 13       # PDF S. 14


# ---------------------------------------------------------------- Hilfsfunktionen

def spalte_mit(zeile, *texte):
    """Index der ersten Zelle, die mit einem der Texte beginnt (sonst None)."""
    for i, c in enumerate(zeile):
        if any(c.strip().startswith(t) for t in texte):
            return i
    return None


def seitentext(pdf, idx):
    return cell_text(pdf.pages[idx].chars, join="\n")


def suche(text, muster, standard, feld):
    """Regex-Treffer (Gruppe 1) oder Standardwert; Fehlschlag wird protokolliert."""
    m = re.search(muster, text)
    if m:
        return m.group(1).strip()
    protokoll.append(f"{feld}: nicht im PDF-Text gefunden, Standardwert {standard!r} verwendet")
    return standard


# ---------------------------------------------------------------- Stammdaten

def build_stammdaten(pdf):
    """Kopfangaben aus dem Fliesstext (Titelseite, Kap. 1, 2, 4)."""
    titel = seitentext(pdf, 0)
    kap1 = seitentext(pdf, 2)
    kap2 = seitentext(pdf, SEITE_PRODUKTE)
    kap4 = seitentext(pdf, 5) + "\n" + seitentext(pdf, 6)

    # Im PDF steht "WAGO-00002-V.01.01-EN"; Dateiname und 00001-Excel nutzen
    # die Schreibweise ohne Punkt nach dem V.
    reg = suche(titel, r"PEP Registration No\.:\s*(WAGO-\d+-V\.?\d+\.\d+-EN)",
                "WAGO-00002-V.01.01-EN", "PEP-Registriernummer").replace("-V.", "-V")
    return {
        "PEP-Registriernummer": reg,
        "Produkt": suche(titel, r"\n(\d{4} / [^\n]*Series)\n",
                         "2273 / PUSH WIRE® Splicing Connector Series", "Produkt"),
        "Ausgabedatum": suche(titel, r"Date of issue:\s*(\d{2}-\d{4})", "02-2026", "Ausgabedatum"),
        "Gueltigkeit": suche(titel, r"Validity period:\s*([^\n]+)", "5 years", "Gueltigkeit"),
        "Verifizierer-Akkreditierung": suche(titel, r"Verifier accreditation number:\s*(\S+)",
                                             "VH44", "Verifizierer-Akkreditierung"),
        "Funktionelle Einheit": suche(" ".join(kap1.split("\n")),
                                      r"The functional unit of this study is:\s*(.*?standards\.)",
                                      "", "Funktionelle Einheit"),
        "Referenzprodukt": suche(kap2, r"\(type (\d{4}-\d{3})\)", "2273-202", "Referenzprodukt"),
        "Kontaktwiderstand (Mikroohm)": suche(kap4, r"resistance of ([\d.]+) micro Ohms",
                                              "480.381944", "Kontaktwiderstand"),
        "Energie Modul B6 (J)": suche(kap4, r"amounts to (\d+) J", "318136", "Energie B6"),
    }


# ---------------------------------------------------------------- Tabellen

def build_produktfamilie(pdf):
    """PDF S. 4: Produktserie / Artikelnummer / Masse der deklarierten Einheit."""
    page = pdf.pages[SEITE_PRODUKTE]
    xs, ys = grid_lines(page, min_len=5)
    grid = extract_grid(page, xs, ys)

    kopf = next(i for i, r in enumerate(grid) if spalte_mit(r, "Item number") is not None)
    c_art = spalte_mit(grid[kopf], "Item number")
    c_masse = spalte_mit(grid[kopf], "Mass of declared unit")

    # Seriename in einer ueber mehrere Zeilen verbundenen Zelle: deren Hoehe
    # ergibt sich aus den breiten Trennlinien der linken Spalte (wie 00001).
    baender = sorted({round(r["top"], 1) for r in page.rects
                      if r["height"] < 2 and r["x0"] < 60 and r["width"] > 100})
    beschriftung = {}
    for oben, unten in zip(baender, baender[1:]):
        text = cell_text([c for c in page.chars
                          if oben < (c["top"] + c["bottom"]) / 2 < unten
                          and xs[0] <= c["x0"] < xs[c_art]])
        if text.startswith("WAGO"):
            beschriftung[(oben, unten)] = text

    def serie_fuer(y):
        for (oben, unten), text in beschriftung.items():
            if oben < y < unten:
                return text
        return None

    rows = []
    for i, r in enumerate(grid[kopf + 1:], start=kopf + 1):
        artikel, masse = r[c_art].strip(), r[c_masse].strip()
        if re.fullmatch(rf"{SERIE}-\d+", artikel):
            mitte = (ys[i] + ys[i + 1]) / 2
            rows.append({"Produktserie": serie_fuer(mitte), "Artikelnummer": artikel,
                         "Masse deklarierte Einheit (g)": round(zu_zahl(masse.replace(".", ",")), 3)})
    return pd.DataFrame(rows)


def build_material(pdf):
    """PDF S. 5: Materialzusammensetzung des Referenzprodukts."""
    page = pdf.pages[SEITE_MATERIAL]
    xs, ys = grid_lines(page, min_len=5)
    grid = extract_grid(page, xs, ys)

    kopfzeile = next(i for i, r in enumerate(grid) if spalte_mit(r, "Materials") is not None)
    c_mat = spalte_mit(grid[kopfzeile], "Materials")
    c_masse = spalte_mit(grid[kopfzeile], "Mass (kg)")
    c_anteil = spalte_mit(grid[kopfzeile], "Percentage")

    def label(r):
        return " ".join(x for x in r[1:c_mat] if x).strip()

    # Kopfangaben oberhalb der Tabelle: Wert steht in der ersten Zelle hinter
    # der Beschriftung (hier in der Spalte "Materials").
    kopf = {}
    for r in grid[:kopfzeile]:
        text = label(r)
        if text.startswith(("Reference product mass", "Packaging mass")) and r[c_mat].strip():
            kopf[" ".join(text.split())] = r[c_mat].strip()

    rows = []
    for r in grid[kopfzeile + 1:]:
        kategorie, material = label(r), r[c_mat].strip()
        masse = zu_zahl(r[c_masse].strip())
        anteil = zu_zahl(r[c_anteil].replace("%", "").strip())
        if kategorie.startswith("Total"):
            rows.append({"Kategorie PEP Material": "Total Weight", "Material": "",
                         "Masse (kg)": masse, "Anteil (%)": anteil})
            break     # danach folgt nur noch die Legende des Kreisdiagramms
        if material and masse is not None:
            rows.append({"Kategorie PEP Material": kategorie, "Material": material,
                         "Masse (kg)": masse, "Anteil (%)": anteil})
    return pd.DataFrame(rows), kopf


def build_faktoren(pdf):
    """PDF S. 14: Multiplikationsfaktoren je Variante und Lebenszyklusphase.

    Die Gruppenueberschriften stehen in verbundenen Zellen, die breiter
    gesetzt sind als ihr Spaltenraster - einzelne Zeichen ragen in die
    Nachbarspalte. Weder Rasterzuordnung noch Luecken-Heuristik trennen sie
    sauber. Die Kopfzeile wird deshalb als Text gelesen und an den
    Modulcodes (A1-3_, A4_, ..., Total, D_) getrennt; die Varianten-Zeile
    (203 204 205 208 203 ...) liefert die Anzahl Spalten je Gruppe.
    """
    page = pdf.pages[SEITE_FAKTOREN]
    xs, ys = grid_lines(page)
    grid = extract_grid(page, xs, ys)

    varianten_zeile = [c.strip() for c in grid[1][1:]]
    varianten = list(dict.fromkeys(v for v in varianten_zeile if v))
    n_gruppen = len(varianten_zeile) // len(varianten)

    kopf = cell_text([c for c in page.chars
                      if ys[0] <= (c["top"] + c["bottom"]) / 2 < ys[1]])
    kopf = kopf.replace("Impact category", "", 1)
    gruppen = [g.strip() for g in re.split(r"(?=(?:[A-D]\d(?:-\d)?|D)_|Total)", kopf) if g.strip()]
    if len(gruppen) != n_gruppen:
        raise SystemExit(f"Faktorentabelle: {len(gruppen)} Ueberschriften {gruppen}, "
                         f"aber {n_gruppen} Spaltengruppen")

    spalten = [f"{g} | {v}" for g in gruppen for v in varianten]
    breit, lang = [], []
    for r in grid[2:]:
        indikator = r[0].strip()
        if not indikator:
            continue
        werte = [zu_zahl(c) for c in r[1:1 + len(spalten)]]
        breit.append({"Serie": SERIE, "Indikator": indikator, **dict(zip(spalten, werte))})
        for (g, v), w in zip([(g, v) for g in gruppen for v in varianten], werte):
            if w is not None:
                lang.append({"Serie": SERIE, "Indikator": indikator,
                             "Phase": g, "Variante": f"{SERIE}-{v}", "Faktor": w})
    return pd.DataFrame(breit), pd.DataFrame(lang)


# ---------------------------------------------------------------- Hauptprogramm

def main(pdf_pfad, xlsx_pfad):
    pdf = open_pdf(pdf_pfad)

    stamm = build_stammdaten(pdf)
    produktfamilie = build_produktfamilie(pdf)
    material, material_kopf = build_material(pdf)
    setup = build_lca_setup(pdf)

    ind_pflicht, lang_pflicht = build_indikatoren(pdf, SEITE_IND_PFLICHT, "Verbindliche Indikatoren")
    ind_sachbilanz, lang_sachbilanz = build_indikatoren(pdf, SEITE_IND_SACHBILANZ, "Sachbilanz")
    ind_optional, lang_optional = build_indikatoren(pdf, SEITE_IND_OPTIONAL, "Optionale Indikatoren")

    faktoren, faktoren_lang = build_faktoren(pdf)

    indikatoren_lang = pd.concat([lang_pflicht, lang_sachbilanz, lang_optional],
                                 ignore_index=True)

    ref = stamm["Referenzprodukt"]
    info = pd.DataFrame([
        ("Quelldatei", Path(pdf_pfad).name),
        ("PEP-Registriernummer", stamm["PEP-Registriernummer"]),
        ("Produkt", stamm["Produkt"]),
        ("Referenzprodukt", ref),
        ("Ausgabedatum", stamm["Ausgabedatum"]),
        ("Gueltigkeit", stamm["Gueltigkeit"]),
        ("Referenzlebensdauer", "30 Jahre bei 70 % Nutzungsrate, 1 A"),
        # Die folgenden Felder liest excel_to_aasx_2273.py. Leer gelassene
        # Felder fallen dort auf Standardwerte zurueck bzw. werden als "im PEP
        # nicht enthalten" berichtet.
        ("Funktionelle Einheit", stamm["Funktionelle Einheit"]),
        ("Produktbeschreibung",
         "Compact push-wire splicing connector for solid copper conductors "
         "0.5-2.5 mm², rated 450 V / up to 24 A; the last digit of the item "
         "number gives the number of clamping units"),
        ("EPD-Ersteller", "WAGO Kontakttechnik GmbH & Co. KG"),
        ("Hersteller-Kontakt", "Alexander Flekler, alexander.flekler@wago.com"),
        ("Programmbetreiber-Website", "www.pep-ecopassport.org"),
        ("Programmbetreiber E-Mail", ""),
        ("Deklarations-URL", ""),
        ("Verifizierer-Akkreditierung", stamm["Verifizierer-Akkreditierung"]),
        ("Verifizierung", "extern"),
        ("Referenzjahr Daten", "2024"),
        ("Kontaktwiderstand (Mikroohm)", stamm["Kontaktwiderstand (Mikroohm)"]),
        ("Energie Modul B6 (J)", stamm["Energie Modul B6 (J)"]),
        ("Masse Referenzprodukt",
         material_kopf.get("Reference product mass", "0.71E-03 kg")),
        ("Masse Verpackung je Einheit",
         material_kopf.get("Packaging mass per declared unit", "2.92E-04 kg")),
        ("", ""),
        ("Blatt", "Inhalt (Seite im PDF)"),
        ("Produktfamilie", "Varianten und Massen (S. 4)"),
        ("Materialzusammensetzung", "Werkstoffe des Referenzprodukts (S. 5)"),
        ("LCA_Setup", "Software, Datenbank, Methode (S. 8)"),
        ("Ind_verbindlich", "Verbindliche Wirkungsindikatoren (S. 9)"),
        ("Ind_Sachbilanz", "Sachbilanz-/Ressourcenindikatoren (S. 10)"),
        ("Ind_optional", "Optionale Wirkungsindikatoren (S. 11)"),
        ("Faktoren", "Multiplikationsfaktoren 2273-Serie (S. 14)"),
        ("Indikatoren_lang", "Alle Indikatoren im Langformat - fuer Pivot"),
        ("Faktoren_lang", "Alle Faktoren im Langformat - fuer Pivot"),
        ("", ""),
        ("Hinweis Faktoren",
         f"Wirkung einer Variante = Wert des Referenzprodukts {ref} x Faktor"),
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
        schreibe(writer, faktoren, "Faktoren", "0.00", zahl_ab=2)
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
                     ("Faktoren", faktoren),
                     ("Indikatoren_lang", indikatoren_lang),
                     ("Faktoren_lang", faktoren_lang)]:
        print(f"  {name:26s} {len(df):4d} Zeilen x {len(df.columns):3d} Spalten")
    print("  Faktorgruppen: " + ", ".join(faktoren_lang["Phase"].unique()))

    # Plausibilitaet: jede Variante mit Faktoren braucht eine Massenangabe,
    # jede Nicht-Referenz-Variante Faktoren.
    mit_faktor = set(faktoren_lang["Variante"])
    mit_masse = set(produktfamilie["Artikelnummer"])
    if mit_faktor - mit_masse:
        protokoll.append("Faktoren ohne Produktzeile (keine Masse im PEP): "
                         + ", ".join(sorted(mit_faktor - mit_masse)))
    if mit_masse - mit_faktor - {ref}:
        protokoll.append("Produktzeilen ohne Faktoren: "
                         + ", ".join(sorted(mit_masse - mit_faktor - {ref})))
    if protokoll:
        print("\nWarnungen:")
        for w in protokoll:
            print("  ", w)


if __name__ == "__main__":
    quelle = sys.argv[1] if len(sys.argv) > 1 else "WAGO-00002-V01.01-EN.pdf"
    ziel = sys.argv[2] if len(sys.argv) > 2 else "WAGO_PEP_2273.xlsx"
    main(quelle, ziel)
