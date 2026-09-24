"""Gemeinsame Bausteine: WAGO-PEP (PEP ecopassport, PDF) -> Standard-Excel (Vorlage v1).

Die PEP-Skripte (pep_221.py, pep_2273.py, ...) enthalten nur, was eine PEP
besonders macht (Texte, Klemmstellen-Regel, Sonderfaelle). Alles, was jede
WAGO-PEP gleich braucht, steht hier einmal:

  1. PDF lesen        Tabellen ueber ihren Inhalt finden (nicht ueber feste
                      Seitenzahlen), Produkte, Indikatoren, Faktoren, Kopfdaten
  2. Zuordnen         PEP-Spalte -> stageCode, PEP-Name -> indicatorCode
  3. Pruefen          Vollstaendigkeit, Total = Summe der Module, Kopfdaten
  4. Schreiben        Standard-Excel ueber epd_excel.schreibe_arbeitsmappe,
                      danach epd_excel.pruefe

Grundsatz: lieber abbrechen als falsch schreiben. Kopfdaten, die im PDF nicht
gefunden werden, fuehren zu einem Fehler (kein stiller Standardwert). Ein
Wert kann dann im PEP-Skript unter "kopfdaten" fest vorgegeben werden.

Neue PEP: Ordner epd/<serie>/ anlegen, PDF hineinlegen, epd/2273/pep_2273.py
dorthin kopieren und umbenennen, KONFIG anpassen, ausfuehren, Meldungen lesen.
"""

import datetime as dt
import re
import zipfile
from pathlib import Path

from pdf_grid import cell_text, extract_grid, grid_lines, open_pdf

import epd_excel

PROJEKT = Path(__file__).resolve().parent.parent          # skripte/.. = Projektordner
AASX_TEMPLATE = PROJEKT / "vorlagen" / "epd-type-iii-submodel-template.aasx"

# --------------------------------------------------------------- Zuordnungen (PEP ecopassport)

# Spalten der Indikatortabellen (Reihenfolge wie im PDF)
PEP_SPALTEN = ["A1-3", "A4", "A5", "B1-7", "C1-4", "Total", "D"]
# PEP-Spalte -> stageCode. PEP Kap. 4.4/4.5: B1-7 enthaelt nur B6, C1-4 nur C2-C4.
PHASEN = {"A1-3": "A1-A3", "A4": "A4", "A5": "A5", "B1-7": "B6", "C1-4": "C2-C4", "D": "D", "Total": "Total"}
# Modulspalten, deren Summe die Spalte Total ergibt (D liegt ausserhalb)
SUMME_AUS = ["A1-3", "A4", "A5", "B1-7", "C1-4"]
# Gruppenueberschrift der Faktorentabellen (PEP Kap. 7) -> PEP-Spalte
FAKTOR_GRUPPEN = {"A1-3": "A1-3", "A4": "A4", "A5": "A5", "B6": "B1-7", "C1-4": "C1-4", "D": "D", "Total": "Total"}
# Zeilen der Sachbilanz, die keine Indikatoren sind -> Blatt Produkte
BIOGEN = {"Biogenic carbon content of the product": "biogen_produkt",
          "Biogenic carbon content of the associated packaging": "biogen_verpackung"}

PASSIV = ("Not applicable: passive component requiring no maintenance, repair, replacement, refurbishment "
          "or water during its lifetime (PEP section 4.4).")

ZAHL = re.compile(r"^-?\d{1,3}(?:\.\d{3})*(?:,\d+)?(?:[Ee][+-]?\d+)?$|^-?\d+(?:,\d+)?(?:[Ee][+-]?\d+)?$")

# Kopfdaten: Feld -> (Regex ueber den Text aller Seiten, Flags). Gruppe 1 = Wert.
KOPFDATEN = {
    "registriernummer": (r"PEP Registration No\.?:\s*(WAGO-\d+-V\.?\d+(?:\.\d+)+-[A-Z]{2})", re.M),
    "produktname": (r"^\s*(\d{3,4}\b[^\n]*\bSeries\b[^\n]*?)\s*$", re.M),
    "ausgabedatum": (r"Date of issue:\s*(\d{2}-\d{4})", re.M),
    "gueltigkeit_jahre": (r"Validity period:\s*(\d+)\s*years", re.M),
    "verifizierer_akkreditierung": (r"Verifier accreditation number:\s*(\S+)", re.M),
    "pcr": (r"based on the (PCR-[\w-]+ \d{2} \d{2})", re.M),
    "psr": (r"and (PSR-[\w-]+ \d{2} \d{2})", re.M),
    "referenzprodukt": (r"\(type (\d{3,4}-\d+)\)", re.M),
    "lca_referenzjahr": (r"refers to the (\d{4}) calendar year", re.M),
    "referenzlebensdauer_jahre": (r"(?:over|of) (\d+) years", re.M),
    "hersteller": (r"Company information\s*\n(.+)", re.M),
    "normen": (r"PEPs are compliant with (.+)", re.M),
    "funktionelle_einheit": (r"The functional unit of this study is:\s*(.+?standards\.)", re.S),
    "widerstand": (r"electrical resistance of ([\d.,]+ ?(?:mΩ|m?[Oo]hms?|micro Ohms))", re.M),
    "energie_b6": (r"amounts to ([\d.,]+) J", re.M),
    "produktmasse": (r"Reference product mass\s+(-?[\d.,]+(?:E[+-]?\d+)?)\s*kg", re.M),
}


# --------------------------------------------------------------- PDF-Werkzeuge

def zu_zahl(text):
    """'1,24E-02' -> 0.0124; leer -> None; keine Zahl -> Text unveraendert."""
    s = str(text or "").strip()
    if s == "":
        return None
    if not ZAHL.match(s):
        return s
    return float(s.replace(".", "").replace(",", "."))


def zeilentext(page, xs, ys, zeile):
    chars = [c for c in page.chars if ys[zeile] <= (c["top"] + c["bottom"]) / 2 < ys[zeile + 1]
             and xs[0] <= (c["x0"] + c["x1"]) / 2 < xs[-1]]
    return cell_text(chars)


def ist_abschnitt(row):
    """Zeile ohne Zahlenwert in den Datenspalten = Zwischenueberschrift."""
    return not any(ZAHL.match(c.strip()) for c in row[1:] if c.strip())


def spalte_mit(zeile, *texte):
    for i, c in enumerate(zeile):
        if any(c.strip().startswith(t) for t in texte):
            return i
    return None


def seitentext_roh(page):
    """Zeichen der Seite ohne Wortlogik (die Tabellen-Zeichen sind 'schraeg' codiert)."""
    return cell_text(page.chars)


# --------------------------------------------------------------- Tabellen finden

def finde_seiten(pdf, vorgabe=None):
    """{produkte, setup, indikatoren: [..], faktoren: [..]} als 0-basierte Seitenindizes.

    Erkennung ueber den Inhalt, damit verschobene Seiten nicht stoeren:
      Faktoren     Kopfzeile enthaelt '_Manufacturing Stages'
      Indikatoren  Tabellenkopf 'Impact category' (und keine Faktorentabelle)
      Produkte     Spalte 'Item number'
      LCA-Setup    Zeile 'LCA Software'
    vorgabe: {name: gedruckte Seitenzahl(en)} ueberschreibt die Erkennung.
    """
    seiten = {"produkte": None, "setup": None, "indikatoren": [], "faktoren": []}
    for i, page in enumerate(pdf.pages):
        roh = seitentext_roh(page)
        if "Impactcategory" not in roh.replace(" ", "") and "Item number" not in roh \
                and "LCA Software" not in roh:
            continue
        if "_Manufacturing" in roh.replace(" ", ""):
            seiten["faktoren"].append(i)
            continue
        grid = extract_grid(page, *grid_lines(page, min_len=5))
        zellen = [c.strip() for r in grid for c in r]
        if any(z == "Impact category" for z in zellen):
            seiten["indikatoren"].append(i)
        elif any(z.startswith("Item number") for z in zellen) and seiten["produkte"] is None:
            seiten["produkte"] = i
        elif any(z == "LCA Software" for z in zellen) and seiten["setup"] is None:
            seiten["setup"] = i
    for name, wert in (vorgabe or {}).items():
        seiten[name] = [s - 1 for s in wert] if isinstance(wert, (list, tuple)) else wert - 1
    return seiten


# --------------------------------------------------------------- Tabellen lesen

def lies_produkte(pdf, idx, serie):
    """Produkttabelle: Serie / Artikelnummer / Masse der deklarierten Einheit (g)."""
    page = pdf.pages[idx]
    xs, ys = grid_lines(page, min_len=5)
    grid = extract_grid(page, xs, ys)
    kopf = next(i for i, r in enumerate(grid) if spalte_mit(r, "Item number") is not None)
    c_art = spalte_mit(grid[kopf], "Item number")
    c_masse = spalte_mit(grid[kopf], "Mass of declared unit")

    # Serienname in einer ueber mehrere Zeilen verbundenen Zelle; deren Grenzen
    # ergeben sich aus den breiten Trennlinien der linken Spalte.
    baender = sorted({round(r["top"], 1) for r in page.rects
                      if r["height"] < 2 and r["x0"] < 60 and r["width"] > 100})
    beschriftung = {}
    for oben, unten in zip(baender, baender[1:]):
        text = cell_text([c for c in page.chars if oben < (c["top"] + c["bottom"]) / 2 < unten
                          and xs[0] <= c["x0"] < xs[c_art]])
        if text.startswith("WAGO"):
            beschriftung[(oben, unten)] = text.replace(" ,", ",")

    zeilen = []
    for i, r in enumerate(grid[kopf + 1:], start=kopf + 1):
        artikel, masse = r[c_art].strip(), r[c_masse].strip()
        if re.fullmatch(rf"{re.escape(serie)}-\d+", artikel):
            mitte = (ys[i] + ys[i + 1]) / 2
            name = next((t for (o, u), t in beschriftung.items() if o < mitte < u), None)
            zeilen.append({"serie": name, "artikel": artikel,
                           "masse_g": round(zu_zahl(masse.replace(".", ",")), 3)})
    return zeilen


def lies_setup(pdf, idx):
    grid = extract_grid(pdf.pages[idx])
    return {r[0].strip(): r[1].strip() for r in grid[1:] if r[0].strip()}


def lies_indikatoren(pdf, idx):
    """Indikatortabelle -> [{indikator, einheit, spalte, wert}] (nur gefuellte Zellen)."""
    page = pdf.pages[idx]
    xs, ys = grid_lines(page)
    grid = extract_grid(page, xs, ys)
    kopf = next(i for i, r in enumerate(grid) if r[0].strip() == "Impact category")
    out = []
    for r in grid[kopf + 1:]:
        if ist_abschnitt(r):
            continue
        werte = [zu_zahl(c) for c in r[2:9]]
        gefuellt = [(p, v) for p, v in zip(PEP_SPALTEN, werte) if v is not None]
        if len(gefuellt) == 1:                      # verbundene Zelle: Einzelwert = Total
            gefuellt = [("Total", gefuellt[0][1])]
        for p, v in gefuellt:
            if isinstance(v, str):
                raise ValueError(f"S. {idx + 1}: '{v}' in {r[0].strip()} / {p} ist keine Zahl")
            out.append({"indikator": r[0].strip(), "einheit": r[1].strip(), "spalte": p, "wert": v})
    return out


def lies_faktoren(pdf, idx, serie):
    """Faktorentabelle -> [{indikator, spalte, variante, faktor}].

    Die Gruppenueberschriften (A1-3_Manufacturing Stages ...) stehen in
    verbundenen Zellen, die breiter gesetzt sind als ihr Raster. Die Kopfzeile
    wird deshalb als Text gelesen und an den Modulcodes getrennt; die
    Variantenzeile (203 204 205 208 203 ...) liefert die Spalten je Gruppe.
    """
    page = pdf.pages[idx]
    xs, ys = grid_lines(page)
    grid = extract_grid(page, xs, ys)
    k = next(i for i, r in enumerate(grid) if "".join(r).replace(" ", "").startswith("Impactcategory")
             or r[0].replace(" ", "") == "Impactcategory")
    varianten_zeile = [c.strip() for c in grid[k + 1][1:]]
    varianten = list(dict.fromkeys(v for v in varianten_zeile if v))
    kopf = cell_text([c for c in page.chars if ys[k] <= (c["top"] + c["bottom"]) / 2 < ys[k + 1]])
    kopf = re.sub(r"^\s*Impact\s*category", "", kopf.replace("Im p a c t", "Impact"))
    gruppen = [g.strip() for g in re.split(r"(?=(?:[A-D]\d(?:-\d)?|D)_|Total)", kopf) if g.strip()]
    n_gruppen = len([v for v in varianten_zeile if v]) // len(varianten)
    if len(gruppen) != n_gruppen:
        raise ValueError(f"S. {idx + 1}: {len(gruppen)} Gruppenueberschriften {gruppen}, "
                         f"aber {n_gruppen} Spaltengruppen")
    spalten = []
    for g in gruppen:
        code = g.split("_")[0]
        if code not in FAKTOR_GRUPPEN:
            raise ValueError(f"S. {idx + 1}: unbekannte Faktorgruppe '{g}'")
        spalten += [(FAKTOR_GRUPPEN[code], v) for v in varianten]
    out = []
    for r in grid[k + 2:]:
        indikator = r[0].strip()
        if not indikator:
            continue
        for (spalte, v), c in zip(spalten, r[1:1 + len(spalten)]):
            f = zu_zahl(c)
            if f is None:
                continue
            if isinstance(f, str):
                raise ValueError(f"S. {idx + 1}: Faktor '{f}' fuer {indikator} ist keine Zahl")
            out.append({"indikator": indikator, "spalte": spalte, "variante": f"{serie}-{v}", "faktor": f})
    return out


def verifizierung(page):
    """'extern' / 'intern' aus dem Ankreuzfeld auf dem Deckblatt (None = kein Kreuz gefunden)."""
    worte = page.extract_words()
    labels = [w for w in worte if w["text"] in ("Internal:", "External:")]
    for k in (w for w in worte if w["text"] in ("x", "X", "✓", "✔")):
        links = [lb for lb in labels if abs(lb["top"] - k["top"]) <= 15 and lb["x0"] <= k["x0"]]
        if links:
            return "extern" if max(links, key=lambda lb: lb["x0"])["text"] == "External:" else "intern"
    return None


def lies_kopfdaten(pdf, vorgaben):
    """Kopfdaten aus dem Fliesstext. vorgaben (aus dem PEP-Skript) gehen vor.

    Rueckgabe: (daten, fehler). Nicht gefundene Felder ohne Vorgabe sind Fehler.
    """
    text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    daten, fehler = {}, []
    for feld, (muster, flags) in KOPFDATEN.items():
        if feld in vorgaben:
            daten[feld] = vorgaben[feld]
            continue
        m = re.search(muster, text, flags)
        if m:
            daten[feld] = " ".join(m.group(1).split())
        else:
            fehler.append(f"Kopfdaten: '{feld}' nicht im PDF gefunden - im PEP-Skript unter 'kopfdaten' vorgeben")
    daten["verifizierung"] = vorgaben.get("verifizierung") or verifizierung(pdf.pages[0])
    if daten["verifizierung"] is None:
        fehler.append("Kopfdaten: Ankreuzfeld Internal/External nicht erkannt - 'verifizierung' vorgeben")
    if "registriernummer" in daten:
        daten["registriernummer"] = daten["registriernummer"].replace("-V.", "-V")
    if "normen" in daten and isinstance(daten["normen"], str):
        normen = [n.strip().replace(" :", ":") for n in re.split(r"\s+(?:and|or)\s+", daten["normen"])]
        if re.search(r"complies with ISO 14025:2006", text):
            normen.insert(0, "ISO 14025:2006")
        daten["normen"] = normen
    deck = re.search(r"Drafting rules:\s*\W?(PCR-[\w-]+ \d{2} \d{2})", text)
    hinweise = []
    if deck and daten.get("pcr") and deck.group(1) != daten["pcr"]:
        hinweise.append(f"PDF uneinheitlich: Deckblatt nennt {deck.group(1)!r}, Fliesstext {daten['pcr']!r} "
                        "- uebernommen: Fliesstext")
    return daten, fehler, hinweise, text


# --------------------------------------------------------------- Pruefen

def pruefe_extraktion(indikatoren, faktoren, produkte, ref, katalog, text, serie):
    """Plausibilitaet der gelesenen Tabellen. -> (fehler, warnungen)"""
    fehler, warnungen = [], []
    namen = {i["indikator"] for i in indikatoren}
    unbekannt = sorted(namen - set(katalog) - set(BIOGEN))
    if unbekannt:
        fehler.append("Indikatoren nicht im Katalog (epd_excel.STANDARD_INDIKATOREN, Spalte quelle): "
                      + "; ".join(unbekannt))
    fehlt = sorted(set(katalog) - namen)
    if fehlt:
        warnungen.append("Katalog-Indikatoren, die die PEP nicht enthaelt: " + "; ".join(fehlt))

    # Total = Summe der Module A1-3 ... C1-4 (bei Rundung auf 3 Stellen)
    werte = {(i["indikator"], i["spalte"]): i["wert"] for i in indikatoren}
    for name in sorted(namen - set(BIOGEN)):
        teile = [werte.get((name, s)) for s in SUMME_AUS]
        total = werte.get((name, "Total"))
        if total is None or any(t is None for t in teile):
            fehler.append(f"{name}: Werte unvollstaendig ({', '.join(s for s, t in zip(SUMME_AUS, teile) if t is None)}"
                          f"{', Total' if total is None else ''} fehlt)")
            continue
        summe = sum(teile)
        toleranz = max(0.02 * abs(total), 0.02 * max(abs(t) for t in teile), 1e-12)
        if abs(summe - total) > toleranz:
            fehler.append(f"{name}: Summe der Module {summe:.3g} != Total {total:.3g} - Spalte verrutscht?")

    # Faktoren vollstaendig fuer jede Variante mit Produktzeile
    artikel = [p["artikel"] for p in produkte]
    if ref not in artikel:
        fehler.append(f"Referenzprodukt {ref} nicht in der Produkttabelle")
    mit_faktor = {f["variante"] for f in faktoren}
    for a in artikel:
        if a == ref:
            continue
        if a not in mit_faktor:
            fehler.append(f"{a}: keine Faktoren in der Faktorentabelle")
            continue
        vorhanden = {(f["indikator"], f["spalte"]) for f in faktoren if f["variante"] == a}
        luecken = [f"{n}/{s}" for (n, s) in werte if (n, s) not in vorhanden and n not in BIOGEN]
        if luecken:
            fehler.append(f"{a}: {len(luecken)} Faktoren fehlen, z. B. {', '.join(luecken[:4])}")
    ohne = sorted(mit_faktor - set(artikel))
    if ohne:
        warnungen.append(f"Faktoren ohne Zeile in der Produkttabelle (keine Masse, kein Teilmodell): {', '.join(ohne)}")
    im_text = set(re.findall(rf"\b{re.escape(serie)}-\d{{3,4}}\b", text))
    fremd = sorted(im_text - set(artikel) - mit_faktor)
    if fremd:
        warnungen.append(f"Im Text genannt, aber in keiner Tabelle: {', '.join(fremd)}")
    return fehler, warnungen


# --------------------------------------------------------------- Standard-Excel

def bild_aus_template(name, ordner, template=AASX_TEMPLATE):
    """Produktbild in produktbilder/ des EPD-Ordners (fehlt es dort, aus dem AASX-Template holen)."""
    ziel = Path(ordner) / "produktbilder" / name
    if not ziel.is_file():
        with zipfile.ZipFile(template) as z:
            part = next((n for n in z.namelist() if n.endswith("/" + name)), None)
            if part is None:
                return None
            ziel.parent.mkdir(exist_ok=True)
            ziel.write_bytes(z.read(part))
    return f"produktbilder/{name}"


def einheit_normal(text):
    return (text.replace("mΩ", "mOhm").replace("micro Ohms", "micro-ohm").replace("mOhms", "mOhm"))


def pep_zu_vorlage(k, ordner):
    """Eine PEP (Konfiguration k aus dem PEP-Skript) in die Standard-Excel schreiben. -> True = ok

    ordner: EPD-Ordner (epd/<serie>/) mit dem PDF; dorthin wird auch die Excel geschrieben.
    """
    ordner = Path(ordner)
    pdf_pfad = ordner / k["pdf"]
    pdf = open_pdf(pdf_pfad)
    print(f"{k['pdf']} -> {k['ziel']}")

    kopf, fehler, warnungen, text = lies_kopfdaten(pdf, k.get("kopfdaten", {}))
    ref = kopf.get("referenzprodukt", "")
    serie = ref.split("-")[0]
    seiten = finde_seiten(pdf, k.get("seiten"))
    print("  Seiten: " + ", ".join(f"{n} {[s + 1 for s in v] if isinstance(v, list) else (v + 1 if v is not None else '-')}"
                                   for n, v in seiten.items()))
    for name in ("produkte", "setup"):
        if seiten[name] is None:
            fehler.append(f"Tabelle '{name}' nicht gefunden - im PEP-Skript unter 'seiten' vorgeben")
    if not seiten["indikatoren"]:
        fehler.append("keine Indikatortabelle gefunden")
    if fehler:
        epd_excel.ausgeben(fehler, warnungen)
        return False

    produkte = lies_produkte(pdf, seiten["produkte"], serie)
    setup = lies_setup(pdf, seiten["setup"])
    indikatoren = [i for s in seiten["indikatoren"] for i in lies_indikatoren(pdf, s)]
    faktoren = [f for s in seiten["faktoren"] for f in lies_faktoren(pdf, s, serie)]
    katalog = {z[6]: z[0] for z in epd_excel.STANDARD_INDIKATOREN}
    f2, w2 = pruefe_extraktion(indikatoren, faktoren, produkte, ref, katalog, text, serie)
    fehler += f2
    warnungen += w2
    print(f"  gelesen: {len(produkte)} Produkte, {len(indikatoren)} Indikatorwerte, {len(faktoren)} Faktoren")
    if fehler:
        epd_excel.ausgeben(fehler, warnungen)
        print("  Abbruch - keine Excel geschrieben.")
        return False

    # -- Blatt EPD
    monat, jahr = kopf["ausgabedatum"].split("-")
    ausgabe = dt.date(int(jahr), int(monat), 1)
    datenbank = setup.get("LCA database", "ecoinvent")
    software = f"{setup.get('LCA Software', 'openLCA')}, {setup.get('Impact assessment method', 'EN 15804+A2')}"
    reg = kopf["registriernummer"]
    epd = {
        "vorlagen_version": epd_excel.VORLAGEN_VERSION,
        "id_basis": "https://wago.com/ids", "submodel_kennung": "EPD", "dateiname_praefix": "wago",
        "signifikante_stellen": 3,
        "registriernummer": reg,
        "dokument_domaene": "PEP ecopassport",
        "programm_version": re.search(r"V(\d+(?:\.\d+)+)", reg).group(1),
        "ausgabedatum": ausgabe,
        "gueltig_bis": ausgabe.replace(year=ausgabe.year + int(kopf["gueltigkeit_jahre"])),
        "datumsgenauigkeit": "Monat",                      # PEP nennt nur Monat/Jahr
        "sprachen": reg.rsplit("-", 1)[1].lower(),
        "quelldatenformat": "PEP ecopassport PDF",
        "deklarations_url": k.get("deklarations_url"),
        "epd_dokument": k["pdf"],
        "programmbetreiber": "Association P.E.P. (PEP ecopassport)",
        "programmbetreiber_email": k.get("programmbetreiber_email"),
        "verifizierungsart": "ThirdPartyVerification" if kopf["verifizierung"] == "extern" else "InternalVerification",
        "verifizierer_akkreditierung": kopf["verifizierer_akkreditierung"],
        "epd_ersteller": kopf["hersteller"],
        "hersteller": kopf["hersteller"],
        "produktname": kopf["produktname"],
        "referenzlebensdauer_jahre": int(kopf["referenzlebensdauer_jahre"]),
        "geltungsbereich": "001",                          # Nutzung weltweit (globaler Strommix in B6)
        "herstellungsbeschreibung": k["herstellung"],
        "scope_typ": "Class" if len(produkte) > 1 else "Instance",
        "scope_beschreibung": ("Product family PEP {registriernummer} ({produktname}) covering {produkte}; "
                               "reference product {referenzprodukt}. This submodel describes {artikel}: {herkunft}. "
                               "Declared modules: A1-A3, A4, A5, B6, C2-C4, D."),
        "lca_referenzjahr": int(kopf["lca_referenzjahr"]),
        "zeitliche_repraesentativitaet": ("Primary production data refer to calendar year {lca_referenzjahr}; "
                                          f"background data from {datenbank}."),
        "geografischer_scope": "GLO",
        "geografische_beschreibung": ("Primary production data are representative of WAGO's manufacturing "
                                      f"facilities in Germany. Background data are based on {datenbank} and "
                                      "represent average conditions at European or global level; the use stage "
                                      "(B6) is modelled with a global electricity dataset."),
        "technologie_beschreibung": ("Actual manufacturing technologies of the declared product (injection "
                                     "moulding, metal forming and assembly); background processes represent "
                                     f"average state-of-the-art technologies as modelled in {datenbank} ({software})."),
        "energiemodell": "DE",
        "fe_menge": 1, "fe_einheit": "A", "fe_beschreibung": kopf["funktionelle_einheit"],
        "de_menge": 1, "de_einheit": "item",
        "de_beschreibung": "1 item {artikel} including its packaging; mass {masse_kg} kg per declared unit.",
        "recyclinganteil_prozent": None,
    }

    # -- Produkte
    werte = {(i["indikator"], i["spalte"]): i["wert"] for i in indikatoren}
    fak = {(f["variante"], f["indikator"], f["spalte"]): f["faktor"] for f in faktoren}
    zeilen_p = []
    for p in produkte:
        a = p["artikel"]
        n = k.get("klemmstellen", {}).get(a, int(a[-1]))
        z = {"artikel": a, "referenz": "ja" if a == ref else "nein",
             "herkunft": "Deklariert" if a == ref else "Faktor", "serie": p["serie"],
             "beschreibung": f"{p['serie']}, {n} clamping units. {k['produkttext']}",
             "masse_kg": round(p["masse_g"] / 1000, 9),
             "bild": bild_aus_template(k["bilder"][a], ordner) if a in k.get("bilder", {}) else None}
        for name, schl in BIOGEN.items():
            basis = werte.get((name, "Total"))
            if basis is None:
                continue
            if a == ref:
                z[schl] = basis
            elif (a, name, "Total") in fak:                # PEP Kap. 7 nennt auch hierfuer Faktoren
                z[schl] = float(f"{basis * fak[(a, name, 'Total')]:.3g}")
        if a == ref:
            z["produktmasse_kg"] = zu_zahl(kopf["produktmasse"].replace(".", ","))
        zeilen_p.append(z)
    artikel = {z["artikel"] for z in zeilen_p}

    # -- Module (PEP Kap. 4.4 / 4.5)
    b6 = (f"contact resistance {einheit_normal(kopf['widerstand'])}, {kopf['energie_b6']} J of low-voltage "
          "electricity over the reference lifetime")
    module = [
        {"stageCode": "A1-A3", "stageStatus": "Assessed", "hinweis": "PEP column 'A1-3'."},
        {"stageCode": "A4", "stageStatus": "Assessed"},
        {"stageCode": "A5", "stageStatus": "Assessed"},
    ] + [{"stageCode": b, "stageStatus": "NotAssessed", "hinweis": PASSIV} for b in ("B1", "B2", "B3", "B4", "B5")] + [
        {"stageCode": "B6", "stageStatus": "Assessed",
         "hinweis": f"PEP column 'B1-7' contains module B6 only (Joule losses; {b6}). Scale the results by the "
                    "square of the actual current."},
        {"stageCode": "B7", "stageStatus": "NotAssessed", "hinweis": PASSIV},
        {"stageCode": "C1", "stageStatus": "NotAssessed",
         "hinweis": "Module C1 (de-installation) is not applicable and excluded from the system boundary "
                    "(PEP section 4.5)."},
        {"stageCode": "C2-C4", "stageStatus": "Assessed",
         "hinweis": "PEP column 'C1-4', covering modules C2 (transport), C3 (waste treatment) and C4 (disposal); "
                    "no split per module in the PEP."},
        {"stageCode": "D", "stageStatus": "Assessed"},
    ]

    # -- Ergebnisse (Referenzprodukt) und Faktoren (Varianten mit Produktzeile)
    ergebnisse = [{"artikel": ref, "code": katalog[i["indikator"]], "stage": PHASEN[i["spalte"]],
                   "wert": i["wert"], "einheit_quelle": i["einheit"]}
                  for i in indikatoren if i["indikator"] not in BIOGEN]
    faktoren_x = [{"artikel": f["variante"], "code": katalog[f["indikator"]], "stage": PHASEN[f["spalte"]],
                   "faktor": f["faktor"]}
                  for f in faktoren if f["indikator"] not in BIOGEN and f["variante"] in artikel]

    normen = kopf["normen"] + [n for n in k.get("normen_zusatz", []) if n not in kopf["normen"]]
    listen = {
        "Produkte": zeilen_p, "Module": module,
        "Indikatoren": [dict(zip(("code", "kategorie", "name", "einheit", "basiseinheit", "status", "quelle"), z))
                        for z in epd_excel.STANDARD_INDIKATOREN],
        "Ergebnisse": ergebnisse, "Faktoren": faktoren_x,
        "Standorte": [{"name": k["standort"], "typ": "Fixed", "betreiber": kopf["hersteller"], "land": "DE"}],
        "Normen_PCR": ([{"art": "Norm", "bezeichnung": n} for n in normen]
                       + [{"art": "PCR", "bezeichnung": kopf["pcr"]}, {"art": "PSR", "bezeichnung": kopf["psr"]}]),
    }
    ziel = ordner / k["ziel"]
    epd_excel.schreibe_arbeitsmappe(ziel, epd, listen)

    # Selbstkontrolle: die geschriebene Excel muss die Pruefung der Vorlage bestehen
    _, f3, w3 = epd_excel.pruefe(epd_excel.lese_arbeitsmappe(ziel), basis_ordner=ordner)
    print(f"  geschrieben: {len(zeilen_p)} Produkte, {len(ergebnisse)} Werte, {len(faktoren_x)} Faktoren")
    epd_excel.ausgeben(f3, warnungen + w3)
    try:
        anzeige = ziel.resolve().relative_to(PROJEKT)
    except ValueError:
        anzeige = ziel
    print("  " + ("NICHT in Ordnung" if f3 else
                  f"in Ordnung - bitte Excel sichten, dann: python skripte/excel_to_aasx.py {anzeige.as_posix()}"))
    return not f3
