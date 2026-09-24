#!/usr/bin/env python3
"""Tabellen aus einem PEP-ecopassport-PDF (WAGO) in eine Excel-Datei ueberfuehren.

Aufruf:  python pep_to_excel.py eingabe.pdf [ausgabe.xlsx]

Aufbau der Pipeline
-------------------
1. EXTRAKTION   Raster aus gezeichneten Linien/Rechtecken -> Zellen (pdf_grid.py)
2. AUFBEREITUNG Kopfzeilen zusammenfuehren, Abschnittszeilen erkennen,
                Zahlen mit Dezimalkomma in echte Floats wandeln
3. EXPORT       ein Blatt je Tabelle plus zwei Langformat-Blaetter fuer Pivots

Alle Seitenzahlen und Spaltenzuordnungen sind in den build_*-Funktionen
gebuendelt. Aendert sich das PDF-Layout, muss nur dort nachgezogen werden.
Seitenangaben beziehen sich auf die gedruckte Seitenzahl im PEP (= PDF-Seite);
der PDF-Index in pdfplumber ist jeweils um eins kleiner.

Die Kopfdaten (Registriernummer, Ausgabedatum, Verifizierer, PCR/PSR,
funktionelle Einheit ...) liest build_kopfdaten() aus dem Fliesstext. Wird
ein Feld nicht gefunden, greift der bisherige Standardwert und es gibt eine
Warnung - so faellt ein geaendertes PDF auf, statt still falsche Werte zu
uebernehmen.
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
    """Seite 4 (PDF-Index 3): Produktserie / Artikelnummer / Masse."""
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
    """Seite 5 (PDF-Index 4): Materialzusammensetzung des Referenzprodukts."""
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
    """Seite 8 (PDF-Index 7): Software-/Datenbank-Parameter der Oekobilanz."""
    grid = extract_grid(pdf.pages[7])
    rows = [{"Parameter": r[0].strip(), "Beschreibung": r[1].strip()}
            for r in grid[1:] if r[0].strip()]
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- Kopfdaten

# Feld -> (Regex, Standardwert). Die Regex laeuft ueber den Text aller Seiten;
# Gruppe 1 ist der Wert. Standardwerte = Stand WAGO-00001-V01.01-EN.
KOPFDATEN = {
    "PEP-Registriernummer": (r"PEP Registration No\.?:\s*([A-Z0-9][\w.-]+)", "WAGO-00001-V01.01-EN"),
    "Ausgabedatum": (r"Date of issue:\s*(\d{2}-\d{4})", "02-2026"),
    "Gueltigkeit (Jahre)": (r"Validity period:\s*(\d+)\s*years", "5"),
    "Verifizierer-Akkreditierung": (r"Verifier accreditation number:\s*(\S+)", "VH44"),
    "PCR": (r"based on the (PCR-[\w-]+ \d{2} \d{2})", "PCR-ed4-EN-2021 09 06"),
    "PSR": (r"and (PSR-[\w-]+ \d{2} \d{2})", "PSR-0001-ed4-EN-2022 11 16"),
    "Referenzprodukt": (r"product with code (\d{3}-\d+)", "221-422"),
    "Referenzjahr Primaerdaten": (r"refers to the (\d{4}) calendar year", "2024"),
    "Referenzlebensdauer (Jahre)": (r"reference lifetime\s+\(RTL\) is (\d+) years", "30"),
    "Programmbetreiber-Website": (r"Information and reference documents:\s*(\S+)", "www.pep-ecopassport.org"),
    "Produkt": (r"^(.*\bSeries /.*)$", "221 Series / Splicing Connector with Levers"),
    "EPD-Ersteller": (r"Company information\s*\n(.+)", "WAGO Kontakttechnik GmbH & Co. KG"),
    "Hersteller-Adresse": (r"Company information\s*\n.+\n(.+)", "Hansastra\u00dfe 27, 32423 Minden, Germany"),
    "Hersteller-Kontakt": (r"Contact person:\s*(.+)", "Alexander Flekler, alexander.flekler@wago.com"),
    "Normen": (r"PEPs are compliant with (.+)", "XP C08-100-1:2016 and EN 50693:2019 or NF E38-500 :2022"),
    "Funktionelle Einheit": (r"The functional unit of this study is:\s*(.+?)\s*Lifetime and use rate",
                             "The function of the product is to connect power transmission cables together, "
                             "or to connect them to equipment, for one unit and its packaging, under operating "
                             "conditions identical to those of the cable: 1 ampere (A) over 30 years, with a use "
                             "rate of 70%, in accordance with the applicable standards."),
}


def verifizierung(page):
    """'extern' / 'intern' aus dem Ankreuzfeld auf dem Deckblatt.

    Das Kreuz ist ein einzelnes 'x' rechts neben 'Internal:' bzw.
    'External:'. Es gilt die Beschriftung, die auf derselben Zeile links vom
    Kreuz am naechsten steht.
    """
    worte = page.extract_words()
    labels = [w for w in worte if w["text"] in ("Internal:", "External:")]
    for k in (w for w in worte if w["text"] in ("x", "X", "\u2713", "\u2714")):
        links = [l for l in labels if abs(l["top"] - k["top"]) <= 15 and l["x0"] <= k["x0"]]
        if links:
            return "extern" if max(links, key=lambda l: l["x0"])["text"] == "External:" else "intern"
    return None


def build_kopfdaten(pdf):
    """Kopfdaten des PEP aus dem Fliesstext aller Seiten."""
    text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    daten = {}
    for feld, (muster, standard) in KOPFDATEN.items():
        m = re.search(muster, text, re.S if feld == "Funktionelle Einheit" else re.M)
        if m:
            daten[feld] = " ".join(m.group(1).split())
        else:
            daten[feld] = standard
            protokoll.append(f"Kopfdaten: '{feld}' nicht im PDF gefunden, Standardwert benutzt")
    daten["Produkt"] = daten["Produkt"] if daten["Produkt"].startswith("WAGO") else "WAGO " + daten["Produkt"]
    daten["Verifizierung"] = verifizierung(pdf.pages[0]) or "extern"
    # "XP C08-100-1:2016 and EN 50693:2019 or NF E38-500 :2022" -> Liste,
    # ergaenzt um ISO 14025 (Deckblatt, "Document complies with ...").
    normen = [n.strip().replace(" :", ":") for n in re.split(r"\s+(?:and|or)\s+", daten["Normen"])]
    if re.search(r"complies with ISO 14025:2006", text):
        normen.insert(0, "ISO 14025:2006")
    daten["Normen"] = "; ".join(normen)
    # Das Deckblatt nennt "PCR-4-ed4-EN-2021 09 06", der Fliesstext (S. 3, 6,
    # 7) durchgaengig "PCR-ed4-EN-2021 09 06". Uebernommen wird der Fliesstext.
    deck = re.search(r"Drafting rules:\s*\W?(PCR-[\w-]+ \d{2} \d{2})", text)
    if deck and deck.group(1) != daten["PCR"]:
        protokoll.append(f"PDF uneinheitlich: Deckblatt nennt {deck.group(1)!r}, "
                         f"Fliesstext {daten['PCR']!r} - uebernommen: Fliesstext")
    return daten


# ---------------------------------------------------------------- Indikatoren

PHASEN = ["A1-3", "A4", "A5", "B1-7", "C1-4", "Total", "D"]


def build_indikatoren(pdf, seite_idx, tabellenname):
    """Seiten 9/10/11 (PDF-Index 8/9/10): Indikatortabellen mit sieben Lebenszyklusspalten.

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
    """Seiten 14/15 (PDF-Index 13/14): Multiplikationsfaktoren je Variante und Phase.

    Achtung WAGO-00001: Die Green-Range-Tabelle (S. 14) enthaelt eine vierte
    Variantenspalte "2421", die im PDF weggeschnitten und damit unsichtbar
    ist - im Text steht sie aber (S. 4). In der Produkttabelle fehlt sie
    (keine Masse), der Text auf S. 13 nennt stattdessen "221-2431". Die
    Werte werden uebernommen, in Faktoren_lang aber mit einem Hinweis
    versehen (siehe main).
    """
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

    kopf = build_kopfdaten(pdf)

    # Varianten mit Faktoren, aber ohne Zeile in der Produkttabelle (bei
    # WAGO-00001: 221-2421, Spalte im PDF ausgeblendet, siehe build_faktoren).
    nur_faktoren = sorted(set(faktoren_lang["Variante"]) - set(produktfamilie["Artikelnummer"]))
    faktoren_lang["Hinweis"] = [
        "nicht in der Produkttabelle (S. 4): keine Masse im PEP, kein Teilmodell"
        if v in nur_faktoren else "" for v in faktoren_lang["Variante"]]

    info = pd.DataFrame([
        ("Quelldatei", Path(pdf_pfad).name),
        ("PEP-Registriernummer", kopf["PEP-Registriernummer"]),
        ("Produkt", kopf["Produkt"]),
        ("Referenzprodukt", kopf["Referenzprodukt"]),
        ("Ausgabedatum", kopf["Ausgabedatum"]),
        ("Gueltigkeit", f"{kopf['Gueltigkeit (Jahre)']} Jahre"),
        ("Referenzlebensdauer", f"{kopf['Referenzlebensdauer (Jahre)']} Jahre bei 70 % Nutzungsrate, 1 A"),
        # Die folgenden Felder liest excel_to_aasx.py fuer das Teilmodell.
        # Leer gelassene Felder werden dort weggelassen und im Bericht genannt.
        ("Funktionelle Einheit", kopf["Funktionelle Einheit"]),
        ("EPD-Ersteller", kopf["EPD-Ersteller"]),
        ("Hersteller-Adresse", kopf["Hersteller-Adresse"]),
        ("Hersteller-Kontakt", kopf["Hersteller-Kontakt"]),
        ("Programmbetreiber", "Association P.E.P. (PEP ecopassport)"),
        ("Programmbetreiber-Website", kopf["Programmbetreiber-Website"]),
        ("Programmbetreiber E-Mail", ""),
        ("Deklarations-URL", ""),
        ("Verifizierung", kopf["Verifizierung"]),
        ("Verifizierer-Akkreditierung", kopf["Verifizierer-Akkreditierung"]),
        ("PCR", kopf["PCR"]),
        ("PSR", kopf["PSR"]),
        ("Normen", kopf["Normen"]),
        ("Referenzjahr Primaerdaten", kopf["Referenzjahr Primaerdaten"]),
        ("Referenzlebensdauer (Jahre)", kopf["Referenzlebensdauer (Jahre)"]),
        ("Masse Referenzprodukt",
         material_kopf.get("Reference product mass", "1.60E-03 kg")),
        ("Masse Verpackung je Einheit",
         material_kopf.get("Packaging mass per declared unit", "3.58E-04 kg")),
        ("", ""),
        ("Blatt", "Inhalt (Seite im PDF)"),
        ("Produktfamilie", "Varianten und Massen (S. 4)"),
        ("Materialzusammensetzung", "Werkstoffe des Referenzprodukts (S. 5)"),
        ("LCA_Setup", "Software, Datenbank, Methode (S. 8)"),
        ("Ind_verbindlich", "Verbindliche Wirkungsindikatoren (S. 9)"),
        ("Ind_Sachbilanz", "Sachbilanz-/Ressourcenindikatoren (S. 10)"),
        ("Ind_optional", "Optionale Wirkungsindikatoren (S. 11)"),
        ("Faktoren_GreenRange", "Multiplikationsfaktoren Green Range (S. 14)"),
        ("Faktoren_Original", "Multiplikationsfaktoren Originalserie (S. 15)"),
        ("Indikatoren_lang", "Alle Indikatoren im Langformat - fuer Pivot"),
        ("Faktoren_lang", "Alle Faktoren im Langformat - fuer Pivot"),
        ("", ""),
        ("Hinweis Faktoren",
         "Wirkung einer Variante = Wert des Referenzprodukts " + kopf["Referenzprodukt"] + " x Faktor"),
        ("Hinweis Nutzungsphase",
         "B6-Ergebnisse mit dem Quadrat des tatsaechlichen Stroms skalieren"),
        ("Hinweis Lebenszyklus",
         "Spalte B1-7 enthaelt nur B6 (B1-B5, B7 nicht anwendbar, S. 7); "
         "Spalte C1-4 enthaelt C2-C4 (C1 ausgeschlossen, S. 7/12)"),
    ] + ([("Hinweis Varianten",
           ", ".join(nur_faktoren) + ": nur in der Faktorentabelle, nicht in der "
           "Produkttabelle - keine Masse im PEP, daher kein Teilmodell")]
         if nur_faktoren else []),
        columns=["Feld", "Wert"])

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
    if nur_faktoren:
        protokoll.append("Faktoren ohne Produktzeile (keine Masse im PEP): "
                         + ", ".join(nur_faktoren))
    # Artikelnummern, die im Text genannt, aber in keiner Tabelle stehen
    # (WAGO-00001 S. 13: "221-2431" - vermutlich Tippfehler fuer 221-2421).
    serie = kopf["Referenzprodukt"].split("-")[0]
    text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    bekannt = set(produktfamilie["Artikelnummer"]) | set(faktoren_lang["Variante"])
    fremd = sorted(set(re.findall(rf"\b{serie}-\d{{3,4}}\b", text)) - bekannt)
    if fremd:
        protokoll.append("Im Text genannt, aber in keiner Tabelle: " + ", ".join(fremd))
    if protokoll:
        print("\nWarnungen:")
        for w in protokoll:
            print("  ", w)


if __name__ == "__main__":
    quelle = sys.argv[1] if len(sys.argv) > 1 else "eingabe.pdf"
    ziel = sys.argv[2] if len(sys.argv) > 2 else Path(quelle).with_suffix(".xlsx").name
    main(quelle, ziel)
