#!/usr/bin/env python3
"""Fuellt das Teilmodell "EPD Type III" (Template-Stand September 2026) aus der Excel.

Aufruf:
    python excel_to_aasx.py WAGO_PEP_221.xlsx --out ./aasx

    # alle Varianten der Produktfamilie, hochgerechnet ueber die
    # Multiplikationsfaktoren aus Blatt "Faktoren_lang":
    python excel_to_aasx.py WAGO_PEP_221.xlsx --alle-varianten --out ./aasx

    Optionen: --template (Standard epd-type-iii-submodel-template.aasx),
    --pdf (Quell-PDF, Standard: "Quelldatei" aus dem Info-Blatt), --bild
    (Produktbild), --mit-total, --ohne-idshort-in-listen, --basis-id,
    --submodel-kennung, --behalte-leere, --behalte-leere-pflicht

Neues Template (September 2026)
-------------------------------
Gegenueber dem alten Template hat sich die Struktur grundlegend geaendert,
deshalb lief das alte Skript in "kein Bauplan fuer lciaResults gefunden":
    lciaResults (lciaIndicator/characterizationUnit/lifeCyclePhaseValues)
        -> EnvironmentalResults (resultCategory/indicatorCode/unit/stageValues)
    representativenessReference + systemBoundary -> LCAMethodology (+ DeclaredStages)
    declaredUnit (klein)        -> DeclaredUnit + FunctionalUnit
    biogenicCarbonContent       -> ContentDeclaration/biogenicCarbonContent
    neu: EPDScope, verificationType/verificationDate
    Anhaenge liegen unter /aasx/files/ statt /aasx/documents/
Das Template ist ausserdem mit Beispielwerten gefuellt (teils als
"[ILLUSTRATIVE EXAMPLE]" markiert, z. B. ein fiktives Vertragswerk). Diese
Werte werden deshalb zuerst komplett geleert; ins Ergebnis kommt nur, was
aus der Excel (= dem PEP) stammt.

Ablauf
------
1. Template-AASX lesen, data.xml mit lxml parsen
2. Bauplaene fuer Listeneintraege sichern, dann alle Beispielwerte leeren
3. Felder und Listen aus der Excel befuellen
4. Leere Elemente entfernen, Pflichtfelder ohne Wert im Bericht nennen
5. Konformitaetspass: Template-Fehler beheben, die aas-core3 meldet
   (doppelte ConceptDescription-Ids, ungueltige idShorts, falsch typisierte
   Listen, fehlende Einheiten/Definitionen, Key-Typen, ungenutzte CDs)
6. Mit aas-core3 einlesen, pruefen und in XSD-Reihenfolge neu schreiben
7. AASX schreiben: OPC-Teile der Vorlage, neue data.xml, PDF/Bild unter
   /aasx/files/ samt Relationships und Content-Types

Pruefung: aas-core3 (pip install aas-core3.0) wird fuer Schritt 6 benutzt.
Erwartet sind ausschliesslich AASd-120-Meldungen (idShort an Listenkindern,
WAGO-Konvention). Alles andere wird als FEHLER ausgegeben, Exit-Code 1.

Identifier (Abstimmung mit WAGO, September 2026)
------------------------------------------------
    Teilmodell   https://wago.com/ids/submodel/221-422/EPD
    AAS          https://wago.com/ids/aas/221-422
    Asset        https://wago.com/ids/asset/221-422
Anpassbar ueber --basis-id / --submodel-kennung bzw. ID_MUSTER.

Fachwissen, das nicht aus der Excel kommt, steht gebuendelt in INDIKATOREN
(PDF-Name -> Code/Kategorie), STUFEN (PEP-Spalte -> stageCode) und TEXTE.
"""

import argparse
import copy
import re
import sys
import zipfile
from decimal import Decimal
from pathlib import Path

import pandas as pd
from lxml import etree

try:
    from aas_core3 import verification as aas_verification
    from aas_core3 import xmlization as aas_xml
except ImportError:                                   # pragma: no cover
    aas_verification = aas_xml = None

NS = "https://admin-shell.io/aas/3/0"
Q = lambda tag: f"{{{NS}}}{tag}"

# --------------------------------------------------------------- Einstellungen

# Das PEP nennt nur Monat und Jahr ("02-2026"). True = xs:date mit Tag 01
# plus Qualifier DatePrecision; False waere nicht xs:date-konform.
DATUM_ISO = True

# AAS v3 verbietet idShort bei direkten Kindern einer SubmodelElementList
# (AASd-120). WAGO vergibt sie trotzdem (GWP_total, sv_A1_A3, stage_B6 ...),
# damit die Eintraege ueber den idShortPath der SmartData-AAS-Env-API lesbar
# adressierbar sind. --ohne-idshort-in-listen erzeugt die strikte Form.
IDSHORT_IN_LISTEN = True

ID_MUSTER = {
    "submodel": "{basis}/submodel/{artikel}/{kennung}",
    "aas": "{basis}/aas/{artikel}",
    "asset": "{basis}/asset/{artikel}",
}
BASIS_ID = "https://wago.com/ids"
SUBMODEL_KENNUNG = "EPD"

# Pflichtfelder (SMT/Cardinality One), fuer die das PEP keinen Wert liefert,
# werden entfernt und im Bericht genannt (--behalte-leere-pflicht: stehen lassen).
LEERE_PFLICHT_ENTFERNEN = True

# EPDScope/scopeType, Vokabular des Templates: Instance | Class. Das PEP
# deklariert eine Produktfamilie (S. 13 "homogeneous environmental family"),
# daher Class. Das Template-Beispiel nutzt Instance - bei Bedarf umstellen.
EPD_SCOPE_TYP = "Class"

# Signifikante Stellen fuer Indikatorwerte. Das PEP gibt drei an; hochgerechnete
# Werte (Referenz x Faktor) werden auf dieselbe Genauigkeit gerundet.
STELLEN = 3

# Einheiten fuer ConceptDescriptions mit Datentyp *_MEASURE ohne Einheit
# (AASc-3a-009). Indikator-CDs bekommen ihre Einheit aus der Excel.
CD_EINHEITEN = {
    "https://admin-shell.io/idta/EPD/MassPerDeclaredUnit/1/0": "kg",
    "https://admin-shell.io/idta/EPD/ReferenceServiceLife/1/0": "a",
    "https://admin-shell.io/idta/EPD/BiogenicCarbon/Product/1/0": "kg C",
    "https://admin-shell.io/idta/EPD/BiogenicCarbon/Packaging/1/0": "kg C",
    "https://wg-epd.example.com/draft/v1/ProductMass": "kg",
    "https://wg-epd.example.com/draft/v1/RecycledContentPercentage": "%",
    "https://wg-epd.example.com/draft/v1/Quantity": "1",
    "https://wg-epd.example.com/draft/v1/IndicatorValue":
        "as 'unit' of the enclosing EnvironmentalResult",
}

# --------------------------------------------------------------- Fachmapping

# PDF-Indikatorname -> (indicatorCode, resultCategory, Status im PEP)
# Codes nach EN 15804+A2 in der Schreibweise des Template-Vokabulars
# (CD IndicatorCode: "GWP-total", "Total-PE" ...). PERM, PERT, PENRM, PENRT,
# SM, RSF, NRSF und FW fehlen im Vokabular; das Template erlaubt zusaetzliche
# Codes "where required by the applicable PCR" - sie werden im Bericht genannt.
# Der Status (S. 9/10 verbindlich, S. 11 optional) steht als Qualifier
# PEP/IndicatorStatus am Eintrag.
INDIKATOREN = {
    # verbindliche Wirkungsindikatoren (PDF S. 9)
    "Climate change - total": ("GWP-total", "ImpactIndicator", "mandatory"),
    "Climate change - fossil fuels": ("GWP-fossil", "ImpactIndicator", "mandatory"),
    "Climate change - biogenics": ("GWP-biogenic", "ImpactIndicator", "mandatory"),
    "Climate change - land use and land use transformation": ("GWP-luluc", "ImpactIndicator", "mandatory"),
    "Ozone depletion": ("ODP", "ImpactIndicator", "mandatory"),
    "Acidification (AP)": ("AP", "ImpactIndicator", "mandatory"),
    "Freshwater eutrophication": ("EP-freshwater", "ImpactIndicator", "mandatory"),
    "Marine aquatic eutrophication": ("EP-marine", "ImpactIndicator", "mandatory"),
    "Terrestrial eutrophication": ("EP-terrestrial", "ImpactIndicator", "mandatory"),
    "Photochemical ozone formation": ("POCP", "ImpactIndicator", "mandatory"),
    "Abiotic resource depletion - elements or resource depletion - metals and minerals":
        ("ADPE", "ImpactIndicator", "mandatory"),
    "Abiotic resource depletion - fossil fuels or resource depletion - fossils":
        ("ADPF", "ImpactIndicator", "mandatory"),
    "Water requirement": ("WDP", "ImpactIndicator", "mandatory"),
    # optionale Wirkungsindikatoren (PDF S. 11)
    "Total use of primary energy during the life cycle": ("Total-PE", "ImpactIndicator", "optional"),
    "Emission of fine particles": ("PM", "ImpactIndicator", "optional"),
    "Ionizing radiation, human health": ("IR", "ImpactIndicator", "optional"),
    "Ecotoxicity (fresh water)": ("ETP-fw", "ImpactIndicator", "optional"),
    "Human toxicity, carcinogenic effects": ("HTP-c", "ImpactIndicator", "optional"),
    "Human toxicity, non-carcinogenic effects": ("HTP-nc", "ImpactIndicator", "optional"),
    "Impacts related to land use/soil quality": ("SQP", "ImpactIndicator", "optional"),
    # Ressourceneinsatz (PDF S. 10)
    "Use of renewable primary energy, excluding renewable primary energy resources used as raw materials":
        ("PERE", "ResourceUse", "mandatory"),
    "Use of renewable primary energy resources used as raw materials": ("PERM", "ResourceUse", "mandatory"),
    "Total use of renewable primary energy resources": ("PERT", "ResourceUse", "mandatory"),
    "Use of non-renewable primary energy, excluding non-renewable primary energy resources used as raw materials":
        ("PENRE", "ResourceUse", "mandatory"),
    "Use of non-renewable primary energy resources used as raw materials": ("PENRM", "ResourceUse", "mandatory"),
    "Total use of non-renewable primary energy resources": ("PENRT", "ResourceUse", "mandatory"),
    "Use of secondary materials": ("SM", "ResourceUse", "mandatory"),
    "Use of renewable secondary fuels": ("RSF", "ResourceUse", "mandatory"),
    "Use of non-renewable secondary fuels": ("NRSF", "ResourceUse", "mandatory"),
    "Net use of fresh water": ("FW", "ResourceUse", "mandatory"),
    # Abfallkategorien (PDF S. 10)
    "Hazardous waste disposed of": ("HWD", "Waste", "mandatory"),
    "Non-hazardous waste disposed of": ("NHWD", "Waste", "mandatory"),
    "Radioactive waste disposed of": ("RWD", "Waste", "mandatory"),
    # Output-Stroeme (PDF S. 10)
    "Components for reuse": ("CRU", "OutputFlow", "mandatory"),
    "Materials for recycling": ("MFR", "OutputFlow", "mandatory"),
    "Materials for energy recovery": ("MER", "OutputFlow", "mandatory"),
    "Exported energy": ("EE", "OutputFlow", "mandatory"),
}

# Diese beiden Zeilen der Sachbilanz sind keine Indikatoren, sondern Felder
# unter ContentDeclaration/biogenicCarbonContent.
BIOGEN = {
    "Biogenic carbon content of the product": "biogenicCarbonContentProduct",
    "Biogenic carbon content of the associated packaging": "biogenicCarbonContentPackaging",
}

# Phasenbezeichnung im Faktorenblatt -> Spalte der Indikatortabellen
FAKTOR_PHASE = {
    "A1-3_Manufacturing Stages": "A1-3",
    "A4_Distribution phases": "A4",
    "A5_Installation phases": "A5",
    "B6_Use stage": "B1-7",
    "C1-4_End-of-life stages": "C1-4",
    "D_Benefits and loads": "D",
    "Total": "Total",
}

# Lebenszyklusstufen im Teilmodell: (stageCode, PEP-Spalte oder None).
# None = im PEP nicht deklariert -> valueStatus NotDeclared ohne Wert
# (Guide Kap. 10: nie 0 fuer "nicht deklariert"), stageStatus NotAssessed.
#  - B: Das PEP weist eine Spalte "B1-7" aus, laut S. 7 ist davon nur B6
#    relevant (Joule-Verluste); B1-B5 und B7 sind nicht anwendbar.
#  - C: Die Spalte "C1-4" enthaelt C2-C4, C1 ist ausgeschlossen (S. 7, S. 12).
#    Eine Aufteilung auf C2/C3/C4 gibt das PEP nicht her, deshalb bleibt es
#    - wie bei A1-A3 - ein aggregierter Code. "C1-C4" steht nicht im Vokabular
#    des Templates, das zusaetzliche Codes aber ausdruecklich zulaesst.
STUFEN = [
    ("A1-A3", "A1-3"),
    ("A4", "A4"),
    ("A5", "A5"),
    ("B1", None), ("B2", None), ("B3", None), ("B4", None), ("B5", None),
    ("B6", "B1-7"),
    ("B7", None),
    ("C1-C4", "C1-4"),
    ("D", "D"),
]
STUFEN_HINWEIS = {
    "B6": "PEP column 'B1-7'. Only module B6 (energy losses during operation, 0.55 mOhm) is "
          "relevant; B1-B5 and B7 are not applicable to this passive component (PEP section 4.4). "
          "Scale by the square of the actual current.",
    "C1-C4": "PEP column 'C1-4', covering modules C2-C4. C1 (de-installation) is excluded from "
             "the system boundary / declared as zero, manual process without energy use "
             "(PEP sections 4.5 and 6).",
}
PEP_SPALTEN = ["A1-3", "A4", "A5", "B1-7", "C1-4", "Total", "D"]

# Beschreibende Texte aus dem PEP (S. 6 und 8). {jahr} und {datenbank}
# werden aus der Excel gefuellt.
TEXTE = {
    "manufacturingDescription":
        "Production and packaging at WAGO manufacturing facilities in Germany (injection moulding, "
        "metal forming and assembly). Primary data (energy, material losses, internal logistics, "
        "regranulation) refer to calendar year {jahr}.",
    "timeRepresentativenessDescription":
        "Primary production data refer to calendar year {jahr}; background data from {datenbank}.",
    "geographicalScope": "GLO",
    "geographicalDescription":
        "Primary production data are representative of WAGO's manufacturing facilities in Germany. "
        "Background data are based on {datenbank} and represent average conditions at European or "
        "global level; the use stage is modelled with a global electricity dataset.",
    "technologyDescription":
        "Actual manufacturing technologies of the declared product (injection moulding, metal forming "
        "and assembly); background processes represent average state-of-the-art technologies as "
        "modelled in {datenbank}.",
    "lcaEnergyModel":
        "Manufacturing: average electricity mix of WAGO production facilities in Germany "
        "(calendar year {jahr}). Use stage (B6): global low-voltage electricity dataset ({datenbank}).",
}

# Semantic-Ids fuer Elemente, die das Template nur als leere Liste mitbringt
SEM = {
    "documentId": "https://admin-shell.io/vdi/2770/1/0/DocumentId",
    "documentDomainId": "https://admin-shell.io/vdi/2770/1/0/DocumentDomainId",
    "documentIdentifier": "0173-1#02-ABI500#001",
    "documentIsPrimary": "https://admin-shell.io/idta/EPD/DocumentIsPrimary/1/0",
}

# --------------------------------------------------------------- XML-Werkzeug

IDSHORT_MUSTER = re.compile(r"^[a-zA-Z][a-zA-Z0-9_]*$")
SME = ("property", "multiLanguageProperty", "file", "range", "blob", "referenceElement",
       "submodelElementCollection", "submodelElementList", "relationshipElement",
       "annotatedRelationshipElement", "entity", "operation", "capability", "basicEventElement")

warnungen = []


def bereinige_idshort(text):
    """idShort auf das Muster [a-zA-Z][a-zA-Z0-9_]* bringen ('sv_A1-A3' -> 'sv_A1_A3')."""
    if text is None:
        return None
    neu = re.sub(r"[^a-zA-Z0-9_]", "_", text.strip())
    if not neu or not neu[0].isalpha():
        neu = "x" + neu
    return neu


def idshorts_bereinigen(baum):
    n = 0
    for el in baum.iter(Q("idShort")):
        if el.text and not IDSHORT_MUSTER.match(el.text):
            el.text = bereinige_idshort(el.text)
            n += 1
    return n


def lname(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def inhalt(el):
    """Container der Kinder: aas:value (SMC/SML) bzw. aas:submodelElements."""
    w = el.find(Q("value"))
    return w if w is not None else el.find(Q("submodelElements"))


def kind(el, idshort):
    """Direktes Kind nach idShort (Gross-/Kleinschreibung egal) oder '[i]'."""
    if el is None:
        return None
    werte = inhalt(el)
    if werte is None:
        return None
    if idshort.startswith("[") and idshort.endswith("]"):
        i = int(idshort[1:-1])
        return werte[i] if i < len(werte) else None
    gesucht = bereinige_idshort(idshort).lower()
    for c in werte:
        if (bereinige_idshort(c.findtext(Q("idShort"))) or "").lower() == gesucht:
            return c
    return None


def finde(wurzel, pfad):
    """Element ueber einen idShort-Pfad wie 'manufacturerProduct/productName'."""
    el = wurzel
    for teil in pfad.split("/"):
        el = kind(el, teil)
        if el is None:
            return None
    return el


fehlende_pfade = []


def pfad(sm, p):
    """Wie finde(), merkt sich aber Pfade, die das Template nicht (mehr) hat."""
    el = finde(sm, p)
    if el is None and p not in fehlende_pfade:
        fehlende_pfade.append(p)
    return el


def setze(el, wert):
    """aas:value eines Property / MultiLanguageProperty / File setzen."""
    if el is None or wert is None or (isinstance(wert, float) and pd.isna(wert)) or str(wert).strip() == "":
        return False
    alt = el.find(Q("value"))
    if alt is not None:
        el.remove(alt)
    neu = etree.SubElement(el, Q("value"))
    if lname(el) == "multiLanguageProperty":
        ls = etree.SubElement(neu, Q("langStringTextType"))
        etree.SubElement(ls, Q("language")).text = "en"
        etree.SubElement(ls, Q("text")).text = str(wert)
    else:
        neu.text = str(wert)
    return True


def idshort_setzen(el, text):
    ids = el.find(Q("idShort"))
    if ids is None:
        ids = etree.Element(Q("idShort"))
        el.insert(0, ids)
    ids.text = text


def liste_leeren(el):
    werte = el.find(Q("value"))
    if werte is None:
        werte = etree.SubElement(el, Q("value"))
    for c in list(werte):
        werte.remove(c)
    return werte


def werte_entfernen(el):
    """Alle Nutzwerte unterhalb von el entfernen, Listen leeren, Struktur behalten."""
    for sub in list(el.iter()):
        t = lname(sub)
        if t in ("property", "multiLanguageProperty", "file", "range"):
            for tag in ("value", "min", "max", "valueId"):
                alt = sub.find(Q(tag))
                if alt is not None:
                    sub.remove(alt)
        elif t == "submodelElementList" and sub is not el:
            w = sub.find(Q("value"))
            if w is not None:
                for c in list(w):
                    w.remove(c)


def erster_eintrag(liste):
    """Tiefe Kopie des ersten Listeneintrags (Bauplan) oder None."""
    if liste is None:
        return None
    w = liste.find(Q("value"))
    if w is None or not len(w):
        return None
    return copy.deepcopy(w[0])


def qualifier_anhaengen(el, typ, text):
    """Zusatzangabe maschinenlesbar anhaengen (AASd-021: je Typ ein Qualifier)."""
    if el is None:
        return
    qs = el.find(Q("qualifiers"))
    if qs is None:
        qs = etree.SubElement(el, Q("qualifiers"))
    for alt in qs.findall(Q("qualifier")):
        if alt.findtext(Q("type")) == typ:
            qs.remove(alt)
    q = etree.SubElement(qs, Q("qualifier"))
    etree.SubElement(q, Q("type")).text = typ
    etree.SubElement(q, Q("valueType")).text = "xs:string"
    etree.SubElement(q, Q("value")).text = text


def referenz(tag, wert, typ="ExternalReference", key_typ="GlobalReference"):
    r = etree.Element(Q(tag))
    etree.SubElement(r, Q("type")).text = typ
    k = etree.SubElement(etree.SubElement(r, Q("keys")), Q("key"))
    etree.SubElement(k, Q("type")).text = key_typ
    etree.SubElement(k, Q("value")).text = wert
    return r


def neues_property(idshort, sem, value_type="xs:string", wert=None, kardinalitaet=None):
    el = etree.Element(Q("property"))
    if idshort:
        etree.SubElement(el, Q("idShort")).text = idshort
    if sem:
        el.append(referenz("semanticId", sem))
    if kardinalitaet:
        qualifier_anhaengen(el, "SMT/Cardinality", kardinalitaet)
    etree.SubElement(el, Q("valueType")).text = value_type
    setze(el, wert)
    return el


def neue_smc(idshort, sem, kinder):
    el = etree.Element(Q("submodelElementCollection"))
    if idshort:
        etree.SubElement(el, Q("idShort")).text = idshort
    if sem:
        el.append(referenz("semanticId", sem))
    w = etree.SubElement(el, Q("value"))
    for k in kinder:
        w.append(k)
    return el


def listen_property(liste, wert, idshort, bauplan=None):
    """Property-Eintrag fuer eine SML anlegen (Bauplan oder neu nach semanticIdListElement)."""
    if bauplan is not None and lname(bauplan) == "property":
        el = copy.deepcopy(bauplan)
        idshort_setzen(el, idshort)
        setze(el, wert)
        return el
    sem = liste.findtext(f"{Q('semanticIdListElement')}/{Q('keys')}/{Q('key')}/{Q('value')}")
    vt = liste.findtext(Q("valueTypeListElement")) or "xs:string"
    return neues_property(idshort, sem, vt, wert)


def wert_status(sv, status):
    setze(kind(sv, "valueStatus"), status)


def zahl(v, stellen=STELLEN):
    """xs:decimal: ausgeschrieben, ohne Exponent (0.0124), 'stellen' signifikant."""
    if v is None or pd.isna(v):
        return None
    return f"{Decimal(f'{float(v):.{stellen - 1}E}').normalize():f}"


def double(v, stellen=STELLEN):
    """xs:double, 'stellen' signifikante Stellen ('0.0124', '4.73e-07')."""
    if v is None or pd.isna(v):
        return None
    s = f"{float(v):.{stellen}g}"
    return "0" if s in ("-0", "0") else s


# --------------------------------------------------------------- Excel lesen

def info_text(info, feld, standard=""):
    """Info-Wert als Text; leere Zellen (NaN) fallen auf den Standard zurueck."""
    v = info.get(feld)
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        return standard
    return str(v).strip()


def lade(xlsx):
    x = pd.ExcelFile(xlsx)
    info_df = pd.read_excel(x, "Info").dropna(subset=["Feld"])
    return {
        "info": info_df.set_index("Feld")["Wert"].to_dict(),
        "produkte": pd.read_excel(x, "Produktfamilie", dtype={"Artikelnummer": str}),
        "setup": pd.read_excel(x, "LCA_Setup").set_index("Parameter")["Beschreibung"].to_dict(),
        "indikatoren": pd.read_excel(x, "Indikatoren_lang"),
        "faktoren": pd.read_excel(x, "Faktoren_lang", dtype={"Variante": str}),
    }


def pruefe_mapping(daten):
    """Indikatoren der Excel ohne Mapping (und umgekehrt) melden statt still verlieren."""
    in_excel = set(daten["indikatoren"]["Indikator"])
    ohne = sorted(in_excel - set(INDIKATOREN) - set(BIOGEN))
    if ohne:
        warnungen.append("Indikatoren ohne Eintrag in INDIKATOREN (nicht uebernommen): " + "; ".join(ohne))
    fehlt = sorted(set(INDIKATOREN) - in_excel)
    if fehlt:
        warnungen.append("In INDIKATOREN, aber nicht in der Excel: " + "; ".join(fehlt))


def werte_fuer_variante(daten, artikel, referenz_artikel):
    """{(Indikator, PEP-Spalte): Wert} fuer eine Variante.

    Referenzprodukt: Werte direkt aus dem PEP. Andere Varianten:
    Referenzwert x Multiplikationsfaktor (PEP Abschnitt 7).
    """
    il = daten["indikatoren"]
    basis = {(r["Indikator"], r["Phase"]): r["Wert"] for _, r in il.iterrows()
             if r["Phase"] in PEP_SPALTEN}
    einheiten = dict(zip(il["Indikator"], il["Einheit"]))
    if artikel == referenz_artikel:
        return basis, einheiten

    fl = daten["faktoren"]
    teil = fl[fl["Variante"] == artikel]
    if teil.empty:
        raise SystemExit(f"keine Multiplikationsfaktoren fuer {artikel} im Blatt Faktoren_lang")
    faktoren = {(r["Indikator"], FAKTOR_PHASE[r["Phase"]]): r["Faktor"]
                for _, r in teil.iterrows() if r["Phase"] in FAKTOR_PHASE}
    skaliert = {}
    for schluessel, wert in basis.items():
        f = faktoren.get(schluessel)
        if f is None:
            warnungen.append(f"{artikel}: kein Faktor fuer {schluessel[0]!r} / {schluessel[1]}")
            continue
        skaliert[schluessel] = float(f"{wert * f:.{STELLEN}g}")
    return skaliert, einheiten


def masse_kg(text):
    """'1.60E-03 kg' -> 0.0016"""
    m = re.search(r"-?\d+(?:[.,]\d+)?(?:[Ee][+-]?\d+)?", str(text or ""))
    return float(m.group(0).replace(",", ".")) if m else None


# --------------------------------------------------------------- Befuellen

def bauplaene_sichern(sm):
    """Bauplaene fuer Listeneintraege kopieren, bevor die Beispielwerte geleert werden."""
    er = pfad(sm, "EnvironmentalResults")
    if er is None:
        raise SystemExit("Template ohne 'EnvironmentalResults' - ist das die richtige Vorlage?")
    plan_er = erster_eintrag(er)
    if plan_er is None:
        raise SystemExit("Template: 'EnvironmentalResults' ist leer, kein Bauplan fuer einen Indikator")
    plan_sv = erster_eintrag(kind(plan_er, "stageValues"))
    if plan_sv is None:
        raise SystemExit("Template: kein Bauplan fuer 'stageValues' im ersten EnvironmentalResult")
    plaene = {
        "result": plan_er,
        "stage": plan_sv,
        "stufe": erster_eintrag(finde(sm, "LCAMethodology/DeclaredStages")),
    }
    for listenpfad in ("lca_StandardsPcr/standardsCompliance", "lca_StandardsPcr/productCategoryRules",
                       "EPDScope/coveredProductReferences", "identificationPublication/languages"):
        plaene[listenpfad] = erster_eintrag(finde(sm, listenpfad))
    for p in plaene.values():
        if p is not None:
            werte_entfernen(p)
            for q in p.iter(Q("qualifiers")):              # Beispiel-Qualifier behalten nur Kardinalitaet
                for x in list(q):
                    if x.findtext(Q("type")) != "SMT/Cardinality":
                        q.remove(x)
    return plaene


def fuelle_liste(sm, listenpfad, werte, plaene, praefix):
    liste = pfad(sm, listenpfad)
    if liste is None:
        return
    container = liste_leeren(liste)
    for i, w in enumerate(werte):
        container.append(listen_property(liste, w, f"{praefix}_{i}", plaene.get(listenpfad)))


def befuelle(sm, daten, artikel, offene, dateien, mit_total=False):
    info, setup = daten["info"], daten["setup"]
    ref = info_text(info, "Referenzprodukt", "221-422")
    produkte = daten["produkte"].set_index("Artikelnummer")
    if artikel not in produkte.index:
        raise SystemExit(f"{artikel} steht nicht im Blatt Produktfamilie "
                         f"(vorhanden: {', '.join(produkte.index)})")
    zeile = produkte.loc[artikel]
    masse = float(zeile["Masse deklarierte Einheit (g)"]) / 1000.0
    reg = info_text(info, "PEP-Registriernummer")
    jahr = info_text(info, "Referenzjahr Primaerdaten", "2024")
    datenbank = setup.get("LCA database", "ecoinvent")
    texte = {k: v.format(jahr=jahr, datenbank=datenbank) for k, v in TEXTE.items()}

    ausgabe = info_text(info, "Ausgabedatum", "02-2026")              # "MM-YYYY"
    monat, jahr_ausgabe = ausgabe.split("-")
    gueltig_jahre = int(re.sub(r"\D", "", info_text(info, "Gueltigkeit", "5")) or 5)
    ablauf_jahr = str(int(jahr_ausgabe) + gueltig_jahre)
    if DATUM_ISO:
        ausgabe, ablauf = f"{jahr_ausgabe}-{monat}-01", f"{ablauf_jahr}-{monat}-01"
    else:
        ablauf = f"{monat}-{ablauf_jahr}"

    plaene = bauplaene_sichern(sm)
    werte_entfernen(sm)                                   # Beispielwerte des Templates raus

    # -- Teilmodell-Kopf
    dn = sm.find(f"{Q('displayName')}/{Q('langStringNameType')}/{Q('text')}")
    if dn is not None:
        dn.text = f"WAGO {artikel} EPD Type III"
    de = sm.find(f"{Q('description')}/{Q('langStringTextType')}/{Q('text')}")
    if de is not None:
        de.text = f"Environmental Product Declaration (EPD Type III) for WAGO {artikel}, based on PEP {reg}"
    admin = sm.find(Q("administration"))
    sem = sm.findtext(f"{Q('semanticId')}/{Q('keys')}/{Q('key')}/{Q('value')}")
    if admin is not None and sem and admin.find(Q("templateId")) is not None:
        admin.find(Q("templateId")).text = sem

    # -- Identifikation & Veroeffentlichung
    docs = pfad(sm, "identificationPublication/documentIds")
    if docs is not None:
        liste_leeren(docs).append(neue_smc("documentId_0", SEM["documentId"], [
            neues_property("documentDomainId", SEM["documentDomainId"], "xs:string", "PEP ecopassport", "One"),
            neues_property("documentIdentifier", SEM["documentIdentifier"], "xs:string", reg, "One"),
            neues_property("documentIsPrimary", SEM["documentIsPrimary"], "xs:boolean", "true", "ZeroToOne"),
        ]))
    version = re.search(r"V(\d+(?:\.\d+)+)", reg)
    setze(pfad(sm, "identificationPublication/programOperatorVersion"), version.group(1) if version else None)
    setze(pfad(sm, "identificationPublication/dateOfIssue"), ausgabe)
    setze(pfad(sm, "identificationPublication/validUntil"), ablauf)
    if DATUM_ISO:
        for feld in ("dateOfIssue", "validUntil"):
            qualifier_anhaengen(finde(sm, f"identificationPublication/{feld}"), "DatePrecision",
                                "month (day of month not stated in the PEP)")
    sprache = re.search(r"-([A-Z]{2})$", reg)
    fuelle_liste(sm, "identificationPublication/languages",
                 [sprache.group(1).lower() if sprache else "en"], plaene, "language")
    setze(pfad(sm, "identificationPublication/sourceDataFormat"), "PEP ecopassport PDF")
    setze(pfad(sm, "identificationPublication/declarationUrl"), info_text(info, "Deklarations-URL"))

    # -- Programmbetreiber & Verifizierung
    setze(pfad(sm, "programOperatorVerification/programOperator/Company"),
          info_text(info, "Programmbetreiber", "Association P.E.P. (PEP ecopassport)"))
    setze(pfad(sm, "programOperatorVerification/programOperator/EmailAddress/EmailAddress"),
          info_text(info, "Programmbetreiber E-Mail"))
    setze(pfad(sm, "programOperatorVerification/thirdPartyVerifier/verifierAccreditationId"),
          info_text(info, "Verifizierer-Akkreditierung", "VH44"))
    setze(pfad(sm, "programOperatorVerification/epdDeveloper/Company"),
          info_text(info, "EPD-Ersteller", "WAGO Kontakttechnik GmbH & Co. KG"))
    setze(pfad(sm, "programOperatorVerification/verificationType"),
          "Internal" if info_text(info, "Verifizierung", "extern").lower().startswith("int") else "External")

    # -- Hersteller & Produkt
    hersteller = info_text(info, "EPD-Ersteller", "WAGO Kontakttechnik GmbH & Co. KG")
    setze(pfad(sm, "manufacturerProduct/manufacturerName"), hersteller)
    setze(pfad(sm, "manufacturerProduct/productName"), info_text(info, "Produkt"))
    setze(pfad(sm, "manufacturerProduct/productArticleNumberOfManufacturer"), artikel)
    setze(pfad(sm, "manufacturerProduct/productDescription"), f"{zeile['Produktserie']}, item {artikel}")
    if dateien.get("bild"):
        setze(pfad(sm, "manufacturerProduct/productImage"), dateien["bild"])
    setze(pfad(sm, "manufacturerProduct/massPerDeclaredUnit"), zahl(masse, 4))
    setze(pfad(sm, "manufacturerProduct/referenceServiceLife"),
          info_text(info, "Referenzlebensdauer (Jahre)", "30"))
    setze(pfad(sm, "manufacturerProduct/manufacturingDescription"), texte["manufacturingDescription"])

    # -- LCA-Methodik
    setze(pfad(sm, "LCAMethodology/lcaReferenceYear"), jahr)
    for feld in ("timeRepresentativenessDescription", "geographicalScope", "geographicalDescription",
                 "technologyDescription", "lcaEnergyModel"):
        setze(pfad(sm, f"LCAMethodology/{feld}"), texte[feld])
    stufen = pfad(sm, "LCAMethodology/DeclaredStages")
    if stufen is not None and plaene["stufe"] is not None:
        c = liste_leeren(stufen)
        for code, spalte in STUFEN:
            e = copy.deepcopy(plaene["stufe"])
            idshort_setzen(e, bereinige_idshort(f"stage_{code}"))
            setze(kind(e, "stageCode"), code)
            setze(kind(e, "stageStatus"), "Assessed" if spalte else "NotAssessed")
            c.append(e)

    # -- Normen, PCR/PSR
    normen = [n.strip() for n in info_text(info, "Normen",
              "ISO 14025:2006; XP C08-100-1:2016; EN 50693:2019; NF E38-500:2022").split(";") if n.strip()]
    fuelle_liste(sm, "lca_StandardsPcr/standardsCompliance", normen, plaene, "standard")
    regeln = [info_text(info, "PCR", "PCR-ed4-EN-2021 09 06"), info_text(info, "PSR", "PSR-0001-ed4-EN-2022 11 16")]
    fuelle_liste(sm, "lca_StandardsPcr/productCategoryRules", regeln, plaene, "rule")

    # -- Quelldokument
    if dateien.get("pdf"):
        el = pfad(sm, "epdDocument")
        setze(el, dateien["pdf"])
        if el is not None and el.find(Q("contentType")) is None:
            etree.SubElement(el, Q("contentType")).text = "application/pdf"

    # -- Umweltindikatoren
    werte, einheiten = werte_fuer_variante(daten, artikel, ref)
    n = fuelle_ergebnisse(sm, plaene, werte, einheiten, mit_total)
    if artikel != ref:
        qualifier_anhaengen(finde(sm, "EnvironmentalResults"), "PEP/ValueSource",
                            f"Extrapolated: values of reference product {ref} multiplied by the PEP "
                            f"multiplication factors for {artikel} (PEP section 7), rounded to "
                            f"{STELLEN} significant digits.")

    # -- Geltungsbereich
    setze(pfad(sm, "EPDScope/scopeType"), EPD_SCOPE_TYP)
    fuelle_liste(sm, "EPDScope/coveredProductReferences", list(produkte.index), plaene, "product")
    herkunft = ("values declared directly in the PEP" if artikel == ref else
                f"values derived from reference product {ref} with the PEP multiplication factors")
    setze(pfad(sm, "EPDScope/scopeDescription"),
          f"Product family PEP {reg} ({info_text(info, 'Produkt')}), reference product {ref}. "
          f"This submodel describes variant {artikel}: {herkunft}.")

    # -- Inhaltsdeklaration
    if artikel == ref:
        setze(pfad(sm, "ContentDeclaration/productMass"),
              zahl(masse_kg(info_text(info, "Masse Referenzprodukt")), 3))
    il = daten["indikatoren"]
    for name, ziel in BIOGEN.items():
        treffer = il[(il["Indikator"] == name) & (il["Phase"] == "Total")]
        if treffer.empty:
            continue
        wert = float(treffer["Wert"].iloc[0])
        if artikel != ref:
            fl = daten["faktoren"]
            f = fl[(fl["Variante"] == artikel) & (fl["Indikator"] == name) & (fl["Phase"] == "Total")]
            if f.empty:
                warnungen.append(f"{artikel}: kein Total-Faktor fuer {name!r}")
                continue
            wert *= float(f["Faktor"].iloc[0])
        setze(pfad(sm, f"ContentDeclaration/biogenicCarbonContent/{ziel}"), zahl(wert))

    # -- Funktionelle und deklarierte Einheit
    setze(pfad(sm, "FunctionalUnit/quantity"), "1")
    setze(pfad(sm, "FunctionalUnit/unit"), "A")
    setze(pfad(sm, "FunctionalUnit/description"), info_text(info, "Funktionelle Einheit"))
    setze(pfad(sm, "DeclaredUnit/quantity"), "1")
    setze(pfad(sm, "DeclaredUnit/unit"), "piece")
    setze(pfad(sm, "DeclaredUnit/description"),
          f"1 piece of {artikel} including its packaging; mass {zahl(masse, 4)} kg per declared unit.")

    pruefe(sm, offene)
    return n


def fuelle_ergebnisse(sm, plaene, werte, einheiten, mit_total=False):
    container = liste_leeren(finde(sm, "EnvironmentalResults"))
    stufen = list(STUFEN) + ([("Total", "Total")] if mit_total else [])
    geschrieben = 0
    for name, (code, kategorie, status) in INDIKATOREN.items():
        if not any(s and (name, s) in werte for _, s in stufen):
            continue
        e = copy.deepcopy(plaene["result"])
        idshort_setzen(e, bereinige_idshort(code))
        # Beschreibung des Bauplans ("Global Warming Potential total.") passt
        # nur zum ersten Indikator - durch den Indikator ersetzen.
        for t in e.findall(f"{Q('description')}/{Q('langStringTextType')}/{Q('text')}"):
            t.text = f"{name} ({code})"
        for t in e.findall(f"{Q('displayName')}/{Q('langStringNameType')}/{Q('text')}"):
            t.text = code
        qualifier_anhaengen(e, "PEP/IndicatorStatus", status)
        setze(kind(e, "resultCategory"), kategorie)
        setze(kind(e, "indicatorCode"), code)
        setze(kind(e, "indicatorName"), name)
        setze(kind(e, "unit"), einheiten.get(name))
        sv_liste = kind(e, "stageValues")
        c = liste_leeren(sv_liste)
        for stage, spalte in stufen:
            sv = copy.deepcopy(plaene["stage"])
            idshort_setzen(sv, bereinige_idshort(f"sv_{stage}"))
            setze(kind(sv, "stageCode"), stage)
            wert = werte.get((name, spalte)) if spalte else None
            if wert is not None and not pd.isna(wert):
                wert_status(sv, "Declared")
                setze(kind(sv, "value"), double(wert))
            else:
                wert_status(sv, "NotDeclared")
                v = kind(sv, "value")                     # Guide Kap. 10: kein Wert statt 0
                if v is not None:
                    v.getparent().remove(v)
            if stage in STUFEN_HINWEIS:
                qualifier_anhaengen(sv, "PEP/StageMapping", STUFEN_HINWEIS[stage])
            c.append(sv)
        container.append(e)
        geschrieben += 1
    return geschrieben


def kardinalitaet(el):
    for q in el.findall(f"{Q('qualifiers')}/{Q('qualifier')}"):
        if q.findtext(Q("type")) == "SMT/Cardinality":
            return q.findtext(Q("value"))
    return None


PFLICHT = ("One", "OneToMany")


def hat_inhalt(el):
    tag = lname(el)
    wert = el.find(Q("value"))
    if tag in ("submodelElementCollection", "submodelElementList"):
        return wert is not None and len(wert) > 0
    if tag == "multiLanguageProperty":
        return wert is not None and bool(
            (wert.findtext(f"{Q('langStringTextType')}/{Q('text')}") or "").strip())
    return wert is not None and bool((wert.text or "").strip())


def irgendein_wert(el):
    """Hat el selbst oder irgendein Nachfahre einen Nutzwert?"""
    return any(lname(x) in ("property", "multiLanguageProperty", "file", "range") and hat_inhalt(x)
               for x in el.iter())


def pruefe(sm, offene):
    """Pflichtfelder ohne Wert sammeln (Cardinality One / OneToMany).

    Pflicht ist relativ zum Elternelement: ein Pflichtkind einer optionalen
    Collection, die insgesamt leer bleibt (programOperator/EmailAddress),
    wird nicht gemeldet - die Collection faellt als Ganzes weg.
    """
    def lauf(el, p):
        c = inhalt(el)
        if c is None:
            return
        for k in c:
            tag = lname(k)
            if tag not in SME:
                continue
            kp = f"{p}/{k.findtext(Q('idShort')) or '[]'}"
            pflicht = kardinalitaet(k) in PFLICHT
            if tag in ("submodelElementCollection", "submodelElementList"):
                if not irgendein_wert(k):
                    if pflicht:
                        offene.append(kp)
                else:
                    lauf(k, kp)
            elif pflicht and not hat_inhalt(k):
                offene.append(kp)
    lauf(sm, "")


def aufraeumen(baum):
    """Leere Elemente entfernen (Pflichtfelder nur mit LEERE_PFLICHT_ENTFERNEN).

    Mehrere Durchlaeufe, weil eine Collection erst leer wird, nachdem ihre
    Kinder entfernt wurden (programOperator/EmailAddress ...).
    """
    entfernt = 0
    while True:
        weg = []
        for el in baum.iter():
            if lname(el) not in ("property", "multiLanguageProperty", "file", "range",
                                 "submodelElementCollection", "submodelElementList"):
                continue
            eltern = el.getparent()
            if eltern is None or lname(eltern) not in ("value", "submodelElements"):
                continue
            if not hat_inhalt(el) and (kardinalitaet(el) not in PFLICHT or LEERE_PFLICHT_ENTFERNEN):
                weg.append((eltern, el))
        if not weg:
            return entfernt
        for eltern, el in weg:
            eltern.remove(el)
            entfernt += 1


# --------------------------------------------------------------- Konformitaet

def _lang(el, tag, text, sprache="en"):
    ls = etree.SubElement(el, Q(tag))
    etree.SubElement(ls, Q("language")).text = sprache
    etree.SubElement(ls, Q("text")).text = text
    return ls


def _element_beschreibungen(baum):
    """semanticId -> (idShort, englische Beschreibung) aus den Submodel-Elementen."""
    out = {}
    for el in baum.iter():
        if lname(el) not in ("submodel",) + SME:
            continue
        sem = el.findtext(f"{Q('semanticId')}/{Q('keys')}/{Q('key')}/{Q('value')}")
        if not sem or sem in out:
            continue
        beschr = None
        for ls in el.findall(f"{Q('description')}/{Q('langStringTextType')}"):
            if ls.findtext(Q("language")) == "en":
                beschr = ls.findtext(Q("text"))
        out[sem] = (el.findtext(Q("idShort")), beschr)
    return out


def _kurzname(text, maxlen=18):
    text = (text or "").strip()
    return text if len(text) <= maxlen else text[:maxlen].rstrip("._- ")


def listen_typen(baum):
    """AASd-108/109: typeValueListElement / valueTypeListElement an den Inhalt anpassen.

    Das Template deklariert standardsCompliance und productCategoryRules als
    Liste von Collections, befuellt sie aber mit Properties; bei
    applicableJurisdictions fehlt der valueTypeListElement.
    """
    typname = {"property": "Property", "multiLanguageProperty": "MultiLanguageProperty",
               "submodelElementCollection": "SubmodelElementCollection", "file": "File",
               "submodelElementList": "SubmodelElementList", "range": "Range",
               "referenceElement": "ReferenceElement", "blob": "Blob", "entity": "Entity"}
    n = 0
    for lst in baum.iter(Q("submodelElementList")):
        w = lst.find(Q("value"))
        kinder = [c for c in (w if w is not None else []) if lname(c) in SME]
        if not kinder:
            continue
        arten = {lname(c) for c in kinder}
        if len(arten) != 1:
            warnungen.append(f"Liste {lst.findtext(Q('idShort'))}: gemischte Elementtypen {arten}")
            continue
        art = typname.get(arten.pop())
        tvle = lst.find(Q("typeValueListElement"))
        if tvle is not None and tvle.text != art:
            tvle.text = art
            n += 1
        vtle = lst.find(Q("valueTypeListElement"))
        if art in ("Property", "Range"):
            vts = {c.findtext(Q("valueType")) for c in kinder}
            if vtle is None:
                vtle = etree.Element(Q("valueTypeListElement"))
                tvle.addnext(vtle)
                n += 1
            vtle.text = vts.pop() if len(vts) == 1 else "xs:string"
        elif vtle is not None:
            lst.remove(vtle)
            n += 1
    return n


def konformitaet(baum, einheiten_je_code):
    """Verstoesse gegen AAS-v3-Constraints beheben, die aus dem Template stammen."""
    bericht = []

    n = idshorts_bereinigen(baum)
    if n:
        bericht.append(f"{n} ungueltige idShorts bereinigt (z. B. 'sv_A1-A3' -> 'sv_A1_A3')")

    # -- AASd-120
    if not IDSHORT_IN_LISTEN:
        n = 0
        for lst in baum.iter(Q("submodelElementList")):
            for c in (lst.find(Q("value")) if lst.find(Q("value")) is not None else []):
                ids = c.find(Q("idShort"))
                if ids is not None:
                    c.remove(ids)
                    n += 1
        if n:
            bericht.append(f"AASd-120: {n} idShorts an Listenkindern entfernt")

    # -- AASd-108/109
    n = listen_typen(baum)
    if n:
        bericht.append(f"AASd-108/109: {n} Typangaben von Listen an den Inhalt angepasst")

    # -- AASd-122 / AASd-124: ExternalReference nur mit GlobalReference-Keys
    n = 0
    for ref in baum.iter():
        if lname(ref) not in ("semanticId", "reference", "supplementalSemanticId", "valueId",
                              "dataSpecification", "isCaseOf", "semanticIdListElement"):
            continue
        if ref.findtext(Q("type")) != "ExternalReference":
            continue
        for key in ref.findall(f"{Q('keys')}/{Q('key')}"):
            t = key.find(Q("type"))
            if t is not None and t.text not in ("GlobalReference", "FragmentReference"):
                t.text = "GlobalReference"
                n += 1
    if n:
        bericht.append(f"AASd-122/124: {n} Key-Typen in ExternalReferences auf GlobalReference gesetzt")

    # -- ConceptDescriptions
    cds_el = baum.find(Q("conceptDescriptions"))
    if cds_el is None:
        cds_el = etree.SubElement(baum, Q("conceptDescriptions"))

    # doppelte Ids (Template: 29 Indikator-CDs doppelt) - erste behalten
    gesehen, n = set(), 0
    for cd in list(cds_el):
        i = cd.findtext(Q("id"))
        if i in gesehen:
            cds_el.remove(cd)
            n += 1
        gesehen.add(i)
    if n:
        bericht.append(f"{n} doppelte ConceptDescription-Ids entfernt")

    # ungenutzte CDs (Altlasten http://example.com/..., eclass.example.com/...)
    benutzt = set()
    for el in baum.iter(Q("key")):
        if any(lname(a) == "dataSpecification" for a in el.iterancestors()):
            continue
        benutzt.add(el.findtext(Q("value")))
    n = 0
    for cd in list(cds_el):
        if cd.findtext(Q("id")) not in benutzt:
            cds_el.remove(cd)
            n += 1
    if n:
        bericht.append(f"{n} nicht referenzierte ConceptDescriptions entfernt")

    beschr = _element_beschreibungen(baum)
    vorhanden = {cd.findtext(Q("id")): cd for cd in cds_el.findall(Q("conceptDescription"))}
    n_def = n_kurz = n_unit = 0
    for cd_id, cd in vorhanden.items():
        iec = cd.find(f"{Q('embeddedDataSpecifications')}/{Q('embeddedDataSpecification')}"
                      f"/{Q('dataSpecificationContent')}/{Q('dataSpecificationIec61360')}")
        if iec is None:
            continue
        for ls in iec.findall(f"{Q('shortName')}/{Q('langStringShortNameTypeIec61360')}"):
            t = ls.find(Q("text"))
            if t is not None and t.text and len(t.text) > 18:
                t.text = _kurzname(t.text)
                n_kurz += 1
        dtyp = iec.findtext(Q("dataType")) or ""
        if dtyp.endswith("_MEASURE") and not iec.findtext(Q("unit")) and iec.find(Q("unitId")) is None:
            code = cd_id.rsplit("/Indicator/", 1)[-1] if "/Indicator/" in cd_id else None
            einheit = CD_EINHEITEN.get(cd_id) or (einheiten_je_code.get(code) if code else None)
            if einheit is None:
                bericht.append(f"AASc-3a-009: keine Einheit fuer {cd_id} bekannt - CD_EINHEITEN ergaenzen")
            else:
                u = etree.Element(Q("unit"))
                u.text = einheit
                anker = iec.find(Q("shortName"))
                anker = anker if anker is not None else iec.find(Q("preferredName"))
                anker.addnext(u)
                n_unit += 1
        defs = iec.find(Q("definition"))
        hat_en = defs is not None and any(
            ls.findtext(Q("language")) == "en" and (ls.findtext(Q("text")) or "").strip()
            for ls in defs.findall(Q("langStringDefinitionTypeIec61360")))
        if not hat_en:
            idshort, text = beschr.get(cd_id, (None, None))
            if not text:
                pn = iec.find(f"{Q('preferredName')}/{Q('langStringPreferredNameTypeIec61360')}")
                text = (pn.findtext(Q("text")) if pn is not None else None) or idshort or cd_id
            if defs is None:
                defs = etree.SubElement(iec, Q("definition"))
            _lang(defs, "langStringDefinitionTypeIec61360", text)
            n_def += 1
    if n_kurz:
        bericht.append(f"IEC 61360: {n_kurz} shortNames auf 18 Zeichen gekuerzt")
    if n_unit:
        bericht.append(f"AASc-3a-009: {n_unit} Einheiten in ConceptDescriptions ergaenzt")
    if n_def:
        bericht.append(f"AASc-3a-008: {n_def} englische Definitionen ergaenzt")

    # fehlende CDs fuer verwendete semanticIds
    n = 0
    for sem, (idshort, text) in beschr.items():
        if sem in vorhanden or idshort is None:
            continue
        cd = etree.SubElement(cds_el, Q("conceptDescription"))
        etree.SubElement(cd, Q("idShort")).text = bereinige_idshort(idshort)
        etree.SubElement(cd, Q("id")).text = sem
        ed = etree.SubElement(etree.SubElement(cd, Q("embeddedDataSpecifications")),
                              Q("embeddedDataSpecification"))
        ed.append(referenz("dataSpecification",
                           "https://admin-shell.io/DataSpecificationTemplates/DataSpecificationIec61360/3/0"))
        iec = etree.SubElement(etree.SubElement(ed, Q("dataSpecificationContent")),
                               Q("dataSpecificationIec61360"))
        _lang(etree.SubElement(iec, Q("preferredName")), "langStringPreferredNameTypeIec61360", idshort)
        _lang(etree.SubElement(iec, Q("shortName")), "langStringShortNameTypeIec61360", _kurzname(idshort))
        etree.SubElement(iec, Q("dataType")).text = "STRING"
        _lang(etree.SubElement(iec, Q("definition")), "langStringDefinitionTypeIec61360", text or idshort)
        vorhanden[sem] = cd
        n += 1
    if n:
        bericht.append(f"{n} ConceptDescriptions fuer semanticIds ohne CD ergaenzt")
    return bericht


def normalisieren(baum):
    """Mit aas-core3 einlesen, pruefen und in Schema-Reihenfolge neu schreiben.

    Das Template haelt die XSD-Reihenfolge nicht ein (z. B. idShort vor
    category, description hinter value). aas-core3 liest das tolerant und
    schreibt schemakonform. Rueckgabe: (XML-Bytes, Liste der Befunde).
    """
    roh = etree.tostring(baum, encoding="unicode")
    if aas_xml is None:
        warnungen.append("aas-core3 nicht installiert (pip install aas-core3.0): keine Pruefung, "
                         "keine Normalisierung der XSD-Reihenfolge")
        return etree.tostring(baum, xml_declaration=True, encoding="utf-8", pretty_print=True), None
    env = aas_xml.environment_from_str(roh)
    befunde = [f"{f.path}: {f.cause}" for f in aas_verification.verify(env)]
    neu = etree.fromstring(aas_xml.to_str(env).encode("utf-8"))
    return etree.tostring(neu, xml_declaration=True, encoding="utf-8", pretty_print=True), befunde


# --------------------------------------------------------------- AASX-Ein/Ausgabe

def lade_template(aasx_pfad):
    with zipfile.ZipFile(aasx_pfad) as z:
        spec = spec_part(z)
        baum = etree.fromstring(z.read(spec))
    sm = baum.find(f"{Q('submodels')}/{Q('submodel')}")
    if sm is None:
        raise SystemExit(f"{aasx_pfad}: kein Submodel in {spec}")
    return baum, sm


def spec_part(z):
    """Part-Name der data.xml aus aasx-origin.rels (Standard aasx/data.xml)."""
    for name in z.namelist():
        if name.endswith("aasx-origin.rels"):
            m = re.search(r'Type="[^"]*aas-spec"[^>]*Target="/?([^"]+)"', z.read(name).decode("utf-8"))
            if not m:
                m = re.search(r'Target="/?([^"]+)"[^>]*Type="[^"]*aas-spec"', z.read(name).decode("utf-8"))
            if m:
                return m.group(1)
    return "aasx/data.xml"


REL_NS = "http://admin-shell.io/aasx/relationships/"


def paketname(datei):
    return re.sub(r"[^A-Za-z0-9._-]", "_", Path(datei).name)


def content_types_ergaenzen(xml, endungen):
    typen = {"pdf": "application/pdf", "png": "image/png", "jpg": "image/jpeg",
             "jpeg": "image/jpeg", "xml": "application/xml"}
    text = xml.decode("utf-8")
    for e in sorted(endungen):
        if re.search(rf'Extension="{e}"', text, re.I):
            continue
        text = text.replace("</Types>", f'<Default Extension="{e}" '
                                         f'ContentType="{typen.get(e, "application/octet-stream")}"/></Types>')
    return text.encode("utf-8")


def schreibe_aasx(vorlage, xml, ziel, dateien):
    """AASX schreiben. dateien: {Part-Name ohne fuehrenden Slash: Bytes}.

    Uebernommen werden die OPC-Teile der Vorlage ([Content_Types].xml, Rels,
    aasx-origin). Nicht uebernommen: data.xml, deren Rels und alle Dateien der
    Vorlage unter aasx/files bzw. aasx/documents - was gebraucht wird, steht in
    'dateien'. Die aas-suppl-Relationships gehen vom aas-spec-Part aus (Part 5).

    Das neue Template benutzt die veraltete Relationship-Form
    "http://www.admin-shell.io/aasx/relationships/...". Eclipse BaSyx findet
    damit kein AAS im Paket. Alle Rels werden deshalb auf die Form aus
    IDTA Part 5 ("http://admin-shell.io/aasx/relationships/...") umgestellt;
    der AASX Package Explorer liest beide.
    """
    with zipfile.ZipFile(vorlage) as quelle:
        spec = spec_part(quelle)
        spec_rels = str(Path(spec).parent / "_rels" / (Path(spec).name + ".rels")).replace("\\", "/")
        rel_ns = REL_NS
        with zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as z:
            for i in quelle.infolist():
                n = i.filename
                if n in (spec, spec_rels) or n.startswith(("aasx/files/", "aasx/documents/")):
                    continue
                roh = quelle.read(n)
                if n == "[Content_Types].xml":
                    roh = content_types_ergaenzen(roh, {Path(d).suffix.lstrip(".").lower() for d in dateien})
                elif n.endswith(".rels"):
                    roh = re.sub(rb"https?://(?:www\.)?admin-shell\.io/aasx/relationships/",
                                 REL_NS.encode(), roh)
                z.writestr(i, roh)
            z.writestr(spec, xml)
            if dateien:
                rels = ['<?xml version="1.0" encoding="utf-8"?>'
                        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
                for i, part in enumerate(dateien, start=1):
                    rels.append(f'<Relationship Type="{rel_ns}aas-suppl" Target="/{part}" Id="Rsuppl{i}"/>')
                rels.append("</Relationships>")
                z.writestr(spec_rels, "".join(rels))
            for part, daten in dateien.items():
                z.writestr(part, daten)


def identifier(artikel, basis=None, kennung=None):
    felder = {"basis": (basis or BASIS_ID).rstrip("/"), "artikel": artikel,
              "kennung": kennung or SUBMODEL_KENNUNG}
    return {rolle: muster.format(**felder) for rolle, muster in ID_MUSTER.items()}


def ids_setzen(baum, ids):
    sm = baum.find(f"{Q('submodels')}/{Q('submodel')}")
    alte_sm_id = sm.findtext(Q("id"))
    sm.find(Q("id")).text = ids["submodel"]
    for shell in baum.iter(Q("assetAdministrationShell")):
        shell.find(Q("id")).text = ids["aas"]
        ai = shell.find(Q("assetInformation"))
        if ai is not None and ai.find(Q("globalAssetId")) is not None:
            ai.find(Q("globalAssetId")).text = ids["asset"]
        for key in shell.iter(Q("key")):
            if key.findtext(Q("type")) == "Submodel" and key.findtext(Q("value")) == alte_sm_id:
                key.find(Q("value")).text = ids["submodel"]


def template_datei(vorlage, sm, feld):
    """(Part-Name, Bytes) der Datei, auf die ein File-Element der Vorlage zeigt."""
    wert = (finde(sm, feld).findtext(Q("value")) or "") if finde(sm, feld) is not None else ""
    part = wert.lstrip("/")
    with zipfile.ZipFile(vorlage) as z:
        if part and part in z.namelist():
            return part, z.read(part)
    return None, None


def normiert(text):
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def finde_pdf(args, daten):
    """Quell-PDF: --pdf, sonst 'Quelldatei' aus dem Info-Blatt neben Excel/aktuellem Ordner."""
    if args.pdf:
        p = Path(args.pdf)
        if not p.is_file():
            raise SystemExit(f"--pdf {args.pdf}: Datei nicht gefunden")
        return p
    name = info_text(daten["info"], "Quelldatei")
    for ordner in (Path(args.excel).resolve().parent, Path.cwd()):
        if name and (ordner / name).is_file():
            return ordner / name
    return None


# --------------------------------------------------------------- Hauptprogramm

def main():
    global IDSHORT_IN_LISTEN, LEERE_PFLICHT_ENTFERNEN
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel")
    ap.add_argument("--template", default="epd-type-iii-submodel-template.aasx")
    ap.add_argument("--referenz", default=None,
                    help="(veraltet, wird ignoriert) - das neue Template bringt alle Bauplaene mit")
    ap.add_argument("--variante", default=None, help="Artikelnummer (Standard: Referenzprodukt)")
    ap.add_argument("--alle-varianten", action="store_true")
    ap.add_argument("--out", default=".")
    ap.add_argument("--basis-id", default=BASIS_ID)
    ap.add_argument("--submodel-kennung", default=SUBMODEL_KENNUNG)
    ap.add_argument("--pdf", default=None, help="Quell-PDF fuer epdDocument (Standard: Info/Quelldatei)")
    ap.add_argument("--bild", default=None, help="Produktbild fuer productImage")
    ap.add_argument("--behalte-leere", action="store_true", help="leere optionale Elemente belassen")
    ap.add_argument("--mit-total", action="store_true", help="zusaetzlich stageCode 'Total' aus dem PEP")
    ap.add_argument("--ohne-idshort-in-listen", action="store_true",
                    help="idShorts an Listenkindern entfernen (strikt AASd-120-konform)")
    ap.add_argument("--behalte-leere-pflicht", action="store_true",
                    help="leere Pflichtfelder (Cardinality One) nicht entfernen")
    args = ap.parse_args()

    IDSHORT_IN_LISTEN = not args.ohne_idshort_in_listen
    LEERE_PFLICHT_ENTFERNEN = not args.behalte_leere_pflicht
    if args.referenz:
        print("Hinweis: --referenz wird mit dem neuen Template nicht mehr gebraucht und ignoriert.")
    if not Path(args.template).is_file():
        raise SystemExit(f"Template {args.template} nicht gefunden")

    daten = lade(args.excel)
    pruefe_mapping(daten)
    ref = info_text(daten["info"], "Referenzprodukt", "221-422")
    einheiten_je_code = {INDIKATOREN[n][0]: e for n, e in
                         zip(daten["indikatoren"]["Indikator"], daten["indikatoren"]["Einheit"])
                         if n in INDIKATOREN}
    nicht_im_vokabular = []

    varianten = list(daten["produkte"]["Artikelnummer"]) if args.alle_varianten else [args.variante or ref]
    nur_faktoren = sorted(set(daten["faktoren"]["Variante"]) - set(daten["produkte"]["Artikelnummer"]))
    if nur_faktoren and args.alle_varianten:
        print("Hinweis: Faktoren ohne Produktzeile (keine Masse im PEP), nicht erzeugt: " + ", ".join(nur_faktoren))

    pdf = finde_pdf(args, daten)
    ziel_ordner = Path(args.out)
    ziel_ordner.mkdir(parents=True, exist_ok=True)
    fehler_gesamt = 0

    for artikel in varianten:
        fehlende_pfade.clear()
        baum, sm = lade_template(args.template)
        dateien, verweise = {}, {}

        # Quell-PDF: angegebene Datei, sonst das PDF der Vorlage, sofern es
        # dieselbe Registriernummer traegt (wago-00001-v01-01-en.pdf).
        if pdf is not None:
            part = f"aasx/files/{paketname(pdf)}"
            dateien[part], verweise["pdf"] = pdf.read_bytes(), "/" + part
        else:
            part, roh = template_datei(args.template, sm, "epdDocument")
            reg = info_text(daten["info"], "PEP-Registriernummer")
            if part and normiert(reg) and normiert(reg) in normiert(Path(part).stem):
                dateien[part], verweise["pdf"] = roh, "/" + part
            else:
                warnungen.append("kein Quell-PDF gefunden - epdDocument entfaellt (--pdf angeben)")
        # Produktbild: --bild, sonst das Bild der Vorlage, wenn es zur Variante gehoert
        if args.bild:
            part = f"aasx/files/{paketname(args.bild)}"
            dateien[part], verweise["bild"] = Path(args.bild).read_bytes(), "/" + part
        else:
            part, roh = template_datei(args.template, sm, "manufacturerProduct/productImage")
            if part and normiert(artikel) in normiert(Path(part).stem):
                dateien[part], verweise["bild"] = roh, "/" + part

        offene = []
        n = befuelle(sm, daten, artikel, offene, verweise, args.mit_total)
        if not args.behalte_leere:
            aufraeumen(baum)
        ids = identifier(artikel, args.basis_id, args.submodel_kennung)
        ids_setzen(baum, ids)
        bericht = konformitaet(baum, einheiten_je_code)
        xml, befunde = normalisieren(baum)

        ziel = ziel_ordner / f"wago-{artikel.lower()}-epd-submodel.aasx"
        schreibe_aasx(args.template, xml, ziel, dateien)

        quelle = "PEP-Werte" if artikel == ref else f"hochgerechnet aus {ref}"
        print(f"{ziel.name:44s} {n:3d} Indikatoren  ({quelle})")
        print(f"   Submodel-Id {ids['submodel']}  |  AAS-Id {ids['aas']}")
        print(f"   Anhaenge: {', '.join('/' + d for d in dateien) or '-'}")
        if fehlende_pfade:
            print(f"   Im Template nicht gefunden ({len(fehlende_pfade)}), nicht befuellt:")
            for p in fehlende_pfade:
                print(f"      {p}")
        if offene:
            was = "entfernt" if LEERE_PFLICHT_ENTFERNEN else "leer belassen"
            print(f"   Pflichtfelder ohne Wert im PEP ({len(offene)}, {was}):")
            for p in offene:
                print(f"      {p}")
        for zeile_b in bericht:
            print(f"   {zeile_b}")
        if befunde is not None:
            erwartet = [b for b in befunde if "AASd-120" in b]
            rest = [b for b in befunde if "AASd-120" not in b]
            print(f"   aas-core3: {len(erwartet)} x AASd-120 (WAGO-Konvention), {len(rest)} weitere Befunde")
            for b in rest:
                print(f"      FEHLER {b}")
            fehler_gesamt += len(rest)
        codes = {INDIKATOREN[k][0] for k in INDIKATOREN}
        vokabular = set()
        for cd in etree.fromstring(xml).iter(Q("conceptDescription")):
            if cd.findtext(Q("id")) == "https://wg-epd.example.com/draft/v1/IndicatorCode":
                vokabular = {v.text for v in cd.iter(Q("value")) if lname(v.getparent()) == "valueReferencePair"}
        if vokabular:
            nicht_im_vokabular = sorted(codes - vokabular)

    if nicht_im_vokabular:
        print("\nHinweis: indicatorCodes ausserhalb des Template-Vokabulars (vom Template zugelassen): "
              + ", ".join(nicht_im_vokabular))
    if warnungen:
        print("\nWarnungen:")
        for w in dict.fromkeys(warnungen):
            print("  ", w)
    if fehler_gesamt:
        print(f"\n{fehler_gesamt} Befunde von aas-core3 - Ergebnis NICHT konform.")
        sys.exit(1)


if __name__ == "__main__":
    main()
