#!/usr/bin/env python3
"""Fuellt das IDTA-Teilmodell "EPD Type III" aus der zuvor erzeugten Excel-Datei.

Aufruf:
    python excel_to_aasx.py WAGO_PEP_221.xlsx \\
        --template epd-type-iii-submodel-template.aasx \\
        --referenz wago-00001-v01-01-en-epd-submodel-instance.aasx \\
        --out ./aasx

    # alle neun Varianten der Produktfamilie auf einmal, hochgerechnet
    # ueber die Multiplikationsfaktoren aus Blatt "Faktoren_lang":
    python excel_to_aasx.py WAGO_PEP_221.xlsx --alle-varianten --out ./aasx

Ablauf
------
1. Template-AASX entpacken, aasx/data.xml mit lxml parsen
2. Skalare Felder ueber ihren idShort-Pfad setzen
3. Wiederholte Listeneintraege (LCIA-Indikatoren, Normen, PCR) aus einem
   Bauplan klonen - Bauplan aus dem Template, ersatzweise aus der Referenz
4. Konformitaetspass (konformitaet): Template-Altlasten beheben, die
   aas-core3 als Verstoss meldet - ungueltige idShorts, Key-Typen in
   ExternalReferences, ConceptDescriptions ohne Definition/Einheit, zu
   lange shortNames, fehlende CDs
5. data.xml zurueck ins AASX schreiben, uebrige OPC-Teile unveraendert
   (Dokumente und data.xml.rels der Vorlage werden ersetzt, nicht dupliziert)
6. Pruefbericht: welche Pflichtfelder (SMT/Cardinality = One) liefert das
   PEP nicht - diese werden entfernt und ausgewiesen

Identifier (Abstimmung mit WAGO, September 2026)
------------------------------------------------
Die Identifier folgen dem Schema der bestehenden WAGO-AAS-Infrastruktur
(SmartData AAS Env / Registry, c1.api.wago.com), Artikelnummer = Produkt-Id:
    Teilmodell   https://wago.com/ids/submodel/221-422/EPD
    AAS          https://wago.com/ids/aas/221-422
    Asset        https://wago.com/ids/asset/221-422
Das Teilmodell-Schema ist mit WAGO abgestimmt. AAS- und Asset-Id sind analog
gebildet; sollten die vorhandenen Shells anders heissen, genuegt --basis-id
bzw. eine Anpassung von ID_MUSTER.

idShorts an Listenkindern (documentId_01, GWP_total, A1_3, standard_0 ...)
werden nach WAGO-Konvention vergeben, siehe IDSHORT_IN_LISTEN.

Ergebnis mit aas-core3.0 (pip install aas-core3.0) pruefen:
    from aas_core3 import xmlization, verification
    env = xmlization.environment_from_str(open("aasx/data.xml").read())
    print(list(verification.verify(env)))
Erwartet sind ausschliesslich AASd-120-Meldungen (idShort an Listenkindern),
die sind mit WAGO abgestimmt. Alles andere waere ein Fehler.

Das Mapping PDF-Indikatorname -> EN-15804-Code steht in INDIKATOREN. Das ist
die einzige Stelle mit Fachwissen, das nicht aus der Excel kommt.
"""

import argparse
import re
import zipfile
from decimal import Decimal
from pathlib import Path

import pandas as pd
from lxml import etree

NS = "https://admin-shell.io/aas/3/0"
NSMAP = {"aas": NS}
Q = lambda tag: f"{{{NS}}}{tag}"

# --------------------------------------------------------------- Einstellungen

# Die vorhandene WAGO-Instanz schreibt "02-2026" in xs:date-Felder. Das ist
# kein gueltiges xs:date. Auf True erzeugt das Skript ISO-Daten (2026-02-01),
# auf False das Format der Bestandsinstanz.
DATUM_ISO = True

# xs:decimal kennt keine Exponentialschreibweise. True = Format der
# Bestandsinstanz ("1.24E-02"), False = datentypkonform ("0.0124").
DEZIMAL_EXPONENT = False

# AAS v3 verbietet idShort bei direkten Kindern einer SubmodelElementList
# (AASd-120). WAGO vergibt sie in seiner AAS-Infrastruktur trotzdem
# (documentId_01, GWP_total, A1_3, standard_0 ...), weil die Eintraege dann
# ueber den idShortPath der SmartData-AAS-Env-API lesbar adressierbar sind
# (lciaResults.GWP_total statt lciaResults[0]) und Viewer sie nicht nur
# durchnummeriert anzeigen. Abgestimmt mit WAGO: idShorts werden gesetzt.
# --ohne-idshort-in-listen erzeugt die strikt spezifikationskonforme Form.
# Der Indikator bleibt in beiden Faellen ueber lciaIndicator/indicatorCode
# identifizierbar.
IDSHORT_IN_LISTEN = True

# Identifier-Schema der WAGO-AAS-Infrastruktur (Artikelnummer = Produkt-Id).
# {basis} wird aus --basis-id gefuellt, {artikel} ist die Artikelnummer,
# {kennung} der Teilmodell-Kurzname aus --submodel-kennung (Standard "EPD",
# analog zu ".../PCF" fuer den Product Carbon Footprint).
ID_MUSTER = {
    "submodel": "{basis}/submodel/{artikel}/{kennung}",
    "aas": "{basis}/aas/{artikel}",
    "asset": "{basis}/asset/{artikel}",
}
BASIS_ID = "https://wago.com/ids"
SUBMODEL_KENNUNG = "EPD"

# Pflichtfelder (SMT/Cardinality = One), fuer die das PEP keinen Wert liefert
# (z. B. pictogramSource), werden standardmaessig entfernt und im Bericht
# ausgewiesen. Ein leeres Element wuerde in Auswertungen wie ein Wert wirken.
# --behalte-leere-pflicht laesst sie stehen.
LEERE_PFLICHT_ENTFERNEN = True

# Einheiten fuer ConceptDescriptions mit Datentyp *_MEASURE, denen im
# Template die Einheit fehlt (AASc-3a-009). Schluessel = CD-Id.
CD_EINHEITEN = {
    "https://admin-shell.io/idta/EPD/MassPerDeclaredUnit/1/0": "kg",
    "https://admin-shell.io/idta/EPD/ReferenceServiceLife/1/0": "year",
    "https://admin-shell.io/idta/EPD/BiogenicCarbon/Product/1/0": "kg C",
    "https://admin-shell.io/idta/EPD/BiogenicCarbon/Packaging/1/0": "kg C",
    "https://admin-shell.io/idta/EPD/QuantityOfReferenceUnit/1/0": "1",
    "0173-1#02-AAO561#004": "1",
    # Modulwert eines LCIA-Ergebnisses: die Einheit haengt vom Indikator ab
    # und steht in characterizationUnit/unit des jeweiligen Eintrags.
    "0173-1#02-ABG858#003": "as characterizationUnit/unit of the enclosing LCIA result",
}

# --------------------------------------------------------------- Fachmapping

# PDF-Indikatorname -> (EN-15804-Code, indicatorGroup, Status im PEP)
# Status "mandatory"/"optional" entspricht der Tabellenzuordnung im PEP
# (S. 9/10 verbindlich, S. 11 optional) und wird als Qualifier
# PEP/IndicatorStatus an den Listeneintrag geschrieben - sonst ginge die
# Unterscheidung verloren.
# Die 20 Codes der Wirkungsindikatoren sind aus der WAGO-Bestandsinstanz
# uebernommen. Die 17 Sachbilanzindikatoren ergaenzt das Skript nach den
# Kuerzeln aus EN 15804+A2 - in der Bestandsinstanz fehlen sie bisher.
INDIKATOREN = {
    # verbindliche Wirkungsindikatoren (PDF S. 9)
    "Climate change - total": ("GWP-total", "environmentalImpact", "mandatory"),
    "Climate change - fossil fuels": ("GWP-fossil", "environmentalImpact", "mandatory"),
    "Climate change - biogenics": ("GWP-biogenic", "environmentalImpact", "mandatory"),
    "Climate change - land use and land use transformation": ("GWP-luluc", "environmentalImpact", "mandatory"),
    "Ozone depletion": ("ODP", "environmentalImpact", "mandatory"),
    "Acidification (AP)": ("AP", "environmentalImpact", "mandatory"),
    "Freshwater eutrophication": ("EP-freshwater", "environmentalImpact", "mandatory"),
    "Marine aquatic eutrophication": ("EP-marine", "environmentalImpact", "mandatory"),
    "Terrestrial eutrophication": ("EP-terrestrial", "environmentalImpact", "mandatory"),
    "Photochemical ozone formation": ("POCP", "environmentalImpact", "mandatory"),
    "Abiotic resource depletion - elements or resource depletion - metals and minerals":
        ("ADPE", "environmentalImpact", "mandatory"),
    "Abiotic resource depletion - fossil fuels or resource depletion - fossils":
        ("ADPF", "environmentalImpact", "mandatory"),
    "Water requirement": ("WDP", "environmentalImpact", "mandatory"),
    # optionale Wirkungsindikatoren (PDF S. 11)
    "Total use of primary energy during the life cycle": ("Total-PE", "environmentalImpact", "optional"),
    "Emission of fine particles": ("PM", "environmentalImpact", "optional"),
    "Ionizing radiation, human health": ("IR", "environmentalImpact", "optional"),
    "Ecotoxicity (fresh water)": ("ETP-fw", "environmentalImpact", "optional"),
    "Human toxicity, carcinogenic effects": ("HTP-c", "environmentalImpact", "optional"),
    "Human toxicity, non-carcinogenic effects": ("HTP-nc", "environmentalImpact", "optional"),
    "Impacts related to land use/soil quality": ("SQP", "environmentalImpact", "optional"),
    # Ressourceneinsatz (PDF S. 10)
    "Use of renewable primary energy, excluding renewable primary energy resources used as raw materials":
        ("PERE", "resourceUse", "mandatory"),
    "Use of renewable primary energy resources used as raw materials": ("PERM", "resourceUse", "mandatory"),
    "Total use of renewable primary energy resources": ("PERT", "resourceUse", "mandatory"),
    "Use of non-renewable primary energy, excluding non-renewable primary energy resources used as raw materials":
        ("PENRE", "resourceUse", "mandatory"),
    "Use of non-renewable primary energy resources used as raw materials": ("PENRM", "resourceUse", "mandatory"),
    "Total use of non-renewable primary energy resources": ("PENRT", "resourceUse", "mandatory"),
    "Use of secondary materials": ("SM", "resourceUse", "mandatory"),
    "Use of renewable secondary fuels": ("RSF", "resourceUse", "mandatory"),
    "Use of non-renewable secondary fuels": ("NRSF", "resourceUse", "mandatory"),
    "Net use of fresh water": ("FW", "resourceUse", "mandatory"),
    # Abfall- und Ausgangsstroeme (PDF S. 10)
    "Hazardous waste disposed of": ("HWD", "wasteAndOutputFlow", "mandatory"),
    "Non-hazardous waste disposed of": ("NHWD", "wasteAndOutputFlow", "mandatory"),
    "Radioactive waste disposed of": ("RWD", "wasteAndOutputFlow", "mandatory"),
    "Components for reuse": ("CRU", "wasteAndOutputFlow", "mandatory"),
    "Materials for recycling": ("MFR", "wasteAndOutputFlow", "mandatory"),
    "Materials for energy recovery": ("MER", "wasteAndOutputFlow", "mandatory"),
    "Exported energy": ("EE", "wasteAndOutputFlow", "mandatory"),
}

# Diese beiden Zeilen der Sachbilanz sind keine LCIA-Indikatoren, sondern
# eigene Felder unter manufacturerProduct/biogenicCarbonContent.
BIOGEN = {
    "Biogenic carbon content of the product": "biogenicCarbonContentProduct",
    "Biogenic carbon content of the associated packaging": "biogenicCarbonContentPackaging",
}

# Phasenbezeichnung im Faktorenblatt -> Lebenszyklusmodul im Teilmodell
FAKTOR_PHASE = {
    "A1-3_Manufacturing Stages": "A1-3",
    "A4_Distribution phases": "A4",
    "A5_Installation phases": "A5",
    "B6_Use stage": "B1-7",
    "C1-4_End-of-life stages": "C1-4",
    "D_Benefits and loads": "D",
}
# "Total" wird bewusst nicht geschrieben: das Teilmodell deklariert Module,
# die Summe bildet das auswertende System selbst.
MODULE = ["A1-3", "A4", "A5", "B1-7", "C1-4", "D"]


# --------------------------------------------------------------- XML-Werkzeug

IDSHORT_MUSTER = re.compile(r"^[a-zA-Z][a-zA-Z0-9_]*$")


def bereinige_idshort(text):
    """idShort auf das erlaubte Muster [a-zA-Z][a-zA-Z0-9_]* bringen.

    Das Template enthaelt 'referenceFlow.id' und 'referenceFlow.amount' -
    Punkte sind in AAS v3 nicht erlaubt. Sie werden zu Unterstrichen; die
    Semantik bleibt ueber die semanticId erhalten. Pfade in diesem Skript
    duerfen weiter die Template-Schreibweise benutzen, kind() normalisiert
    beide Seiten.
    """
    if text is None:
        return None
    neu = re.sub(r"[^a-zA-Z0-9_]", "_", text.strip())
    if not neu or not neu[0].isalpha():
        neu = "x" + neu
    return neu


def idshorts_bereinigen(baum):
    """Alle ungueltigen idShorts im Baum korrigieren, Anzahl zurueckgeben."""
    n = 0
    for el in baum.iter(Q("idShort")):
        if el.text and not IDSHORT_MUSTER.match(el.text):
            el.text = bereinige_idshort(el.text)
            n += 1
    return n


def kind(el, idshort):
    """Direktes Kind eines Wertcontainers nach idShort oder Index.

    Collections und Lists halten ihre Kinder in aas:value, das Submodel
    selbst in aas:submodelElements - beides wird hier abgedeckt. Fuer
    Listenkinder ohne idShort ist die Schreibweise '[0]', '[1]' ... erlaubt.
    """
    werte = el.find(Q("value"))
    if werte is None:
        werte = el.find(Q("submodelElements"))
    if werte is None:
        return None
    if idshort.startswith("[") and idshort.endswith("]"):
        i = int(idshort[1:-1])
        return werte[i] if i < len(werte) else None
    gesucht = bereinige_idshort(idshort)
    for c in werte:
        if bereinige_idshort(c.findtext(Q("idShort"))) == gesucht:
            return c
    # Listenkinder ohne idShort (spezifikationskonform, z. B. wenn ein von
    # diesem Skript erzeugtes Paket als Vorlage dient): Template-Namen wie
    # 'documentId_01' oder 'standard_0' auf ihre Position abbilden.
    if (etree.QName(el).localname == "submodelElementList"
            and not any(c.find(Q("idShort")) is not None for c in werte)):
        m = re.search(r"_(\d+)$", idshort)
        i = int(m.group(1)) if m else 0
        if m and idshort.endswith("_01"):          # 1-basiert im Template
            i = 0
        return werte[i] if i < len(werte) else None
    return None


def idshort_setzen(el, text):
    """idShort eines Elements setzen oder anlegen (steht laut XSD an erster Stelle)."""
    ids = el.find(Q("idShort"))
    if ids is None:
        ids = etree.Element(Q("idShort"))
        el.insert(0, ids)
    ids.text = text


def finde(wurzel, pfad):
    """Element ueber einen idShort-Pfad wie 'manufacturerProduct/productName'."""
    el = wurzel
    for teil in pfad.split("/"):
        if el is None:
            return None
        el = kind(el, teil)
    return el


def setze(el, wert):
    """aas:value eines Property / MultiLanguageProperty / File setzen.

    Das XSD schreibt eine feste Reihenfolge vor; aas:value steht hinter
    aas:valueType und vor aas:valueId. Ein vorhandener Wert wird ersetzt.
    """
    if el is None or wert is None or wert == "":
        return False
    tag = etree.QName(el).localname
    alt = el.find(Q("value"))
    if alt is not None:
        el.remove(alt)

    neu = etree.SubElement(el, Q("value"))
    if tag == "multiLanguageProperty":
        ls = etree.SubElement(neu, Q("langStringTextType"))
        etree.SubElement(ls, Q("language")).text = "en"
        etree.SubElement(ls, Q("text")).text = str(wert)
    else:
        neu.text = str(wert)

    valueid = el.find(Q("valueId"))
    if valueid is not None:                       # aas:value muss davor stehen
        el.remove(neu)
        el.insert(list(el).index(valueid), neu)
    return True


def liste_leeren(el):
    werte = el.find(Q("value"))
    if werte is None:
        werte = etree.SubElement(el, Q("value"))
    for c in list(werte):
        werte.remove(c)
    return werte


def bauplan(template_sm, referenz_sm, pfad):
    """Blaupause fuer einen Listeneintrag holen.

    Das Template laesst manche Listen leer (standardsCompliance,
    productCategoryRules). Dann dient die bereits befuellte Bestandsinstanz
    als Muster - so bleiben semanticIds und Qualifier erhalten, statt dass
    das Skript sie neu erfindet.
    """
    for sm in (template_sm, referenz_sm):
        if sm is None:
            continue
        el = finde(sm, pfad)
        if el is not None:
            werte = el.find(Q("value"))
            if werte is not None and len(werte):
                kopie = etree.fromstring(etree.tostring(werte[0]))
                leeren(kopie)
                return kopie
    return None


def leeren(el):
    """Alle Nutzwerte eines geklonten Bauplans entfernen, Struktur behalten."""
    for sub in el.iter():
        tag = etree.QName(sub).localname
        if tag in ("property", "multiLanguageProperty", "file"):
            alt = sub.find(Q("value"))
            if alt is not None:
                sub.remove(alt)


def zahl(v, stellen=2):
    """Zahlenwert fuer ein xs:decimal-Feld.

    xs:decimal erlaubt keine Exponentialschreibweise. Die Bestandsinstanz
    schreibt trotzdem "1.24E-02" - aas-core3 meldet das als Verstoss gegen
    den deklarierten Datentyp. Standardmaessig schreibt dieses Skript daher
    ausgeschrieben (0.0124); DEZIMAL_EXPONENT = True stellt auf das Format
    der Bestandsinstanz zurueck, falls die Gegenstelle das so erwartet.

    Die Rundung auf drei signifikante Stellen passiert vor der Umwandlung,
    damit keine Float-Artefakte in die ausgeschriebene Form durchschlagen.
    """
    if v is None or pd.isna(v):
        return None
    wissenschaftlich = f"{float(v):.{stellen}E}"
    if DEZIMAL_EXPONENT:
        return wissenschaftlich
    return f"{Decimal(wissenschaftlich):f}"


# --------------------------------------------------------------- Excel lesen

def info_text(info, feld, standard=""):
    """Info-Wert als Text; leere Zellen (NaN) fallen auf den Standard zurueck.

    Ohne diese Huerde wuerde eine leere Excel-Zelle als Text "nan" im
    Teilmodell landen.
    """
    v = info.get(feld)
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        return standard
    return str(v).strip()


def lade(xlsx):
    x = pd.ExcelFile(xlsx)
    info = pd.read_excel(x, "Info").set_index("Feld")["Wert"].to_dict()
    return {
        "info": info,
        "produkte": pd.read_excel(x, "Produktfamilie"),
        "material": pd.read_excel(x, "Materialzusammensetzung"),
        "setup": pd.read_excel(x, "LCA_Setup").set_index("Parameter")["Beschreibung"].to_dict(),
        "indikatoren": pd.read_excel(x, "Indikatoren_lang"),
        "faktoren": pd.read_excel(x, "Faktoren_lang"),
    }


def werte_fuer_variante(daten, artikel):
    """{(Indikator, Modul): Wert} fuer eine Variante der Produktfamilie.

    Fuer das Referenzprodukt 221-422 sind das die Werte direkt aus dem PEP.
    Fuer alle anderen Varianten wird der Referenzwert mit dem
    Multiplikationsfaktor aus Abschnitt 7 des PEP hochgerechnet.
    """
    il = daten["indikatoren"]
    basis = {}
    for _, r in il.iterrows():
        if r["Phase"] in MODULE or r["Phase"] == "Total":
            basis[(r["Indikator"], r["Phase"])] = r["Wert"]
    einheiten = dict(zip(il["Indikator"], il["Einheit"]))

    if artikel == "221-422":
        return basis, einheiten, {}

    fl = daten["faktoren"]
    teil = fl[fl["Variante"] == artikel]
    if teil.empty:
        raise SystemExit(f"keine Extrapolationsfaktoren fuer {artikel} im Blatt Faktoren_lang")

    faktoren = {(r["Indikator"], FAKTOR_PHASE[r["Phase"]]): r["Faktor"]
                for _, r in teil.iterrows() if r["Phase"] in FAKTOR_PHASE}
    total = {r["Indikator"]: r["Faktor"] for _, r in teil.iterrows() if r["Phase"] == "Total"}
    for ind, f in total.items():
        faktoren[(ind, "Total")] = f

    skaliert = {}
    for (ind, modul), wert in basis.items():
        f = faktoren.get((ind, modul))
        if f is not None:
            skaliert[(ind, modul)] = wert * f
    return skaliert, einheiten, total


# --------------------------------------------------------------- Befuellen

def befuelle(sm, referenz_sm, daten, artikel, offene, pdf_anhang=None, mit_total=False):
    """Ein Teilmodell mit den Werten einer Variante fuellen."""
    info, setup = daten["info"], daten["setup"]
    produkte = daten["produkte"].set_index("Artikelnummer")
    zeile = produkte.loc[artikel]
    masse_kg = float(zeile["Masse deklarierte Einheit (g)"]) / 1000.0

    ausgabe = info_text(info, "Ausgabedatum", "02-2026")          # "02-2026"
    monat, jahr = ausgabe.split("-")
    gueltig = str(int(jahr) + 5)
    if DATUM_ISO:
        ausgabe, ablauf = f"{jahr}-{monat}-01", f"{gueltig}-{monat}-01"
    else:
        ablauf = f"{monat}-{gueltig}"

    # -- Identifikation & Veroeffentlichung
    setze(finde(sm, "identificationPublication/documentIds/documentId_01/documentDomainId"),
          "PEP ecopassport")
    setze(finde(sm, "identificationPublication/documentIds/documentId_01/documentIdentifier"),
          info_text(info, "PEP-Registriernummer"))
    setze(finde(sm, "identificationPublication/documentIds/documentId_01/documentIsPrimary"), "true")
    setze(finde(sm, "identificationPublication/programOperatorVersion"), "V01.01-EN")
    setze(finde(sm, "identificationPublication/dateOfIssue"), ausgabe)
    setze(finde(sm, "identificationPublication/validUntil"), ablauf)
    if DATUM_ISO:
        # Das PEP nennt nur Monat und Jahr ("02-2026"); der Tag ist ergaenzt.
        # Das steht maschinenlesbar am Datum, statt still als exakter Tag zu gelten.
        for feld in ("dateOfIssue", "validUntil"):
            qualifier_anhaengen(finde(sm, f"identificationPublication/{feld}"),
                                "DatePrecision", "month (day of month not stated in the PEP)")
    setze(finde(sm, "identificationPublication/languages/language_01"), "en")
    setze(finde(sm, "identificationPublication/sourceDataFormat"), "PEP Ecopassport PDF")

    # -- Programmbetreiber & Verifizierung
    setze(finde(sm, "programOperatorVerification/programOperator/Company"),
          "Association P.E.P. (PEP ecopassport)")
    # Das PEP nennt keine E-Mail des Programmbetreibers, nur die Website.
    # Kommt spaeter ein Wert in die Info-Tabelle, wird er hier uebernommen.
    setze(finde(sm, "programOperatorVerification/programOperator/EmailAddress/EmailAddress"),
          info_text(info, "Programmbetreiber E-Mail"))
    setze(finde(sm, "programOperatorVerification/thirdPartyVerifier/verifierAccreditationId"),
          "VH44")
    # Das PEP nennt keinen externen LCA-Dienstleister; erklaerende
    # Organisation und Ansprechpartner sind WAGO. Ueberschreibbar ueber
    # Info-Feld "EPD-Ersteller".
    setze(finde(sm, "programOperatorVerification/epdDeveloper/Company"),
          info_text(info, "EPD-Ersteller", "WAGO Kontakttechnik GmbH & Co. KG"))

    # -- Hersteller & Produkt
    setze(finde(sm, "manufacturerProduct/manufacturerName"), "WAGO Kontakttechnik GmbH & Co. KG")
    setze(finde(sm, "manufacturerProduct/productName"), info_text(info, "Produkt"))
    setze(finde(sm, "manufacturerProduct/productArticleNumberOfManufacturer"), artikel)
    setze(finde(sm, "manufacturerProduct/productDescription"), str(zeile["Produktserie"]))
    setze(finde(sm, "manufacturerProduct/declaredUnit/quantity"), "1.0")
    setze(finde(sm, "manufacturerProduct/declaredUnit/unit"), "piece")
    setze(finde(sm, "manufacturerProduct/massPerDeclaredUnit"), zahl(masse_kg, 3))
    setze(finde(sm, "manufacturerProduct/referenceServiceLife"), "30")

    # -- Repraesentativitaet
    # Technischer Zweck = funktionelle Einheit aus Abschnitt 1 des PEP.
    setze(finde(sm, "representativenessReference/technicalPurpose"),
          info_text(info, "Funktionelle Einheit",
                    "Connect power transmission cables together or to equipment; "
                    "1 A over 30 years at a 70 % use rate, one unit incl. packaging"))
    # Pfad in Template-Schreibweise; kind() normalisiert den Punkt.
    setze(finde(sm, "representativenessReference/referenceFlow.id"), artikel)
    setze(finde(sm, "representativenessReference/referenceFlow.amount"), "1.0")
    setze(finde(sm, "representativenessReference/referenceYear"), "2024")
    setze(finde(sm, "representativenessReference/dataSetValidUntil"), ablauf)
    setze(finde(sm, "representativenessReference/timeRepresentativenessDescription"),
          f"Primary production data refer to calendar year 2024; declaration valid until {ablauf}")
    setze(finde(sm, "representativenessReference/location"), "DE")
    setze(finde(sm, "representativenessReference/geographicalDescription"),
          "Primary data from WAGO manufacturing facilities in Germany; "
          "background data at European or global level (ecoinvent)")
    # Technologiebeschreibung, nicht Wirkungsabschaetzungsmethode - letztere
    # gehoert nach lciaIndicator/indicatorStandard und steht dort auch.
    setze(finde(sm, "representativenessReference/technologyDescription"),
          "Injection moulding, metal forming and assembly operations; background "
          "processes represent average state-of-the-art technologies as modelled "
          "in " + setup.get("LCA database", "ecoinvent"))
    setze(finde(sm, "representativenessReference/energyModel"),
          "Average electricity mix of WAGO production facilities in Germany, calendar year 2024")

    # -- Deklarierte Einheit & Systemgrenze
    setze(finde(sm, "declaredUnit/declaredQuantity"), "1.0")
    setze(finde(sm, "declaredUnit/unitOfMeasure"), "piece")
    setze(finde(sm, "declaredUnit/massPerDeclaredUnit"), zahl(masse_kg, 3))
    setze(finde(sm, "declaredUnit/referenceServiceLife"), "30")
    for feld in ("includesA1_A3", "includesA4_A5", "includesB1_B7", "includesC1_C4", "includesD"):
        setze(finde(sm, f"systemBoundary/{feld}"), "true")
    # DE und Global widersprechen sich nicht: Primaerdaten stammen aus
    # deutschen Werken, die Nutzungsphase ist weltweit modelliert. Das steht
    # sonst nur implizit in zwei getrennten Feldern.
    setze(finde(sm, "systemBoundary/geographicScope"),
          "Global (use stage, worldwide electricity dataset); "
          "primary manufacturing data: DE")
    setze(finde(sm, "systemBoundary/technologyCoverage"),
          "Injection moulding, metal forming and assembly; background processes "
          "as modelled in " + setup.get("LCA database", "ecoinvent"))

    # includesC1_C4 ist ein einzelnes Boolean fuer die ganze Modulgruppe, das
    # PEP deckt aber nur C2-C4 ab. Das Feld ist ZeroToOne, es waeren also drei
    # Antworten moeglich: true ueberzeichnet C1, false verschweigt C2-C4,
    # Weglassen verliert beides. Gewaehlt: true, weil die Modulgruppe
    # bilanziert ist, plus ein Qualifier mit dem genauen Umfang - anders als
    # Freitext in einem fremden Feld ist der maschinell auswertbar und haengt
    # an genau der Aussage, die er einschraenkt.
    c14 = finde(sm, "systemBoundary/includesC1_C4")
    qualifier_anhaengen(c14, "ScopeRestriction",
                        "Modules C2-C4 only. C1 (de-installation) is excluded from "
                        "the system boundary and declared as zero: de-installation "
                        "is a manual process requiring no energy "
                        "(PCR-ed4-EN-2021 09 06).")
    # Gleiches Muster fuer die Nutzungsphase: bilanziert ist nur B6
    # (Joule-Verluste); B1-B5 und B7 sind fuer das passive Bauteil nicht
    # anwendbar (PEP Abschnitt 4.4).
    qualifier_anhaengen(finde(sm, "systemBoundary/includesB1_B7"), "ScopeRestriction",
                        "Module B6 only (energy losses during operation, 0.55 mOhm, "
                        "scale by the square of the actual current). B1-B5 and B7 are "
                        "not applicable to a passive component (PSR-0001-ed4-EN-2022 11 16).")

    # administration/templateId zeigt im Template auf example.com. Hier den
    # Identifier des IDTA-Teilmodells eintragen, aus dem die Instanz stammt.
    admin = sm.find(Q("administration"))
    sem = sm.find(f"{Q('semanticId')}/{Q('keys')}/{Q('key')}/{Q('value')}")
    if admin is not None and sem is not None and admin.find(Q("templateId")) is not None:
        admin.find(Q("templateId")).text = sem.text

    # -- Verweis auf das Quelldokument
    # Der Wert muss ein Pfad innerhalb des AASX-Pakets sein, sonst laeuft die
    # Referenz ins Leere. Die Datei wird in schreibe_aasx() mitverpackt.
    if pdf_anhang:
        el = finde(sm, "epdDocument")
        setze(el, anhang_pfad(pdf_anhang))
        if el is not None and el.find(Q("contentType")) is None:
            ct = etree.Element(Q("contentType"))
            ct.text = "application/pdf"
            el.insert(list(el).index(el.find(Q("value"))), ct)
    setze(finde(sm, "identificationPublication/declarationUrl"),
          info_text(info, "Deklarations-URL"))

    # -- Normen und PCR/PSR
    fuelle_normen(sm, referenz_sm)

    # -- LCIA-Ergebnisse
    werte, einheiten, total = werte_fuer_variante(daten, artikel)
    n = fuelle_lcia(sm, referenz_sm, werte, einheiten, setup, mit_total)

    # -- biogener Kohlenstoff (eigene Felder, kein LCIA-Indikator)
    il = daten["indikatoren"]
    for name, ziel in BIOGEN.items():
        treffer = il[(il["Indikator"] == name) & (il["Phase"] == "Total")]
        if treffer.empty:
            continue
        wert = float(treffer["Wert"].iloc[0]) * total.get(name, 1.0)
        setze(finde(sm, f"manufacturerProduct/biogenicCarbonContent/{ziel}"), zahl(wert))

    pruefe(sm, offene)
    return n


def fuelle_normen(sm, referenz_sm):
    normen = ["ISO 14025:2006", "EN 50693:2019", "XP C08-100-1:2016", "NF E38-500 :2022"]
    el = finde(sm, "lca_StandardsPcr/standardsCompliance")
    plan = bauplan(sm, referenz_sm, "lca_StandardsPcr/standardsCompliance")
    if el is not None and plan is not None:
        werte = liste_leeren(el)
        for i, norm in enumerate(normen):
            eintrag = etree.fromstring(etree.tostring(plan))
            idshort_setzen(eintrag, f"standard_{i}")
            setze(kind(eintrag, "standardDesignation"), norm)
            werte.append(eintrag)

    regeln = [
        ("PCR", "PCR-ed4-EN-2021 09 06",
         "PEP ecopassport Product Category Rules (PCR)", "ed4", "PEP ecopassport"),
        ("PSR", "PSR-0001-ed4-EN-2022 11 16",
         "PEP ecopassport Product Specific Rules (PSR)", "ed4", "PEP ecopassport"),
    ]
    el = finde(sm, "lca_StandardsPcr/productCategoryRules")
    plan = bauplan(sm, referenz_sm, "lca_StandardsPcr/productCategoryRules")
    if el is not None and plan is not None:
        werte = liste_leeren(el)
        for i, (typ, ident, name, version, hrsg) in enumerate(regeln):
            eintrag = etree.fromstring(etree.tostring(plan))
            idshort_setzen(eintrag, f"productCategoryRule_{i}")
            setze(kind(eintrag, "ruleType"), typ)
            setze(kind(eintrag, "ruleIdentifier"), ident)
            setze(kind(eintrag, "ruleName"), name)
            setze(kind(eintrag, "ruleVersion"), version)
            setze(kind(eintrag, "publisherName"), hrsg)
            werte.append(eintrag)


def fuelle_lcia(sm, referenz_sm, werte, einheiten, setup, mit_total=False):
    el = finde(sm, "lciaResults")
    plan = bauplan(sm, referenz_sm, "lciaResults")
    if el is None or plan is None:
        raise SystemExit("kein Bauplan fuer lciaResults gefunden")

    phasenliste = kind(plan, "lifeCyclePhaseValues")
    phasenplan = None
    if phasenliste is not None:
        pv = phasenliste.find(Q("value"))
        if pv is not None and len(pv):
            phasenplan = etree.fromstring(etree.tostring(pv[0]))
            leeren(phasenplan)
    if phasenplan is None:
        raise SystemExit("kein Bauplan fuer lifeCyclePhaseValues gefunden")

    container = liste_leeren(el)
    geschrieben = 0
    for name, (code, gruppe, status) in INDIKATOREN.items():
        module = [(m, werte[(name, m)]) for m in MODULE if (name, m) in werte]
        if not module:
            continue
        # "Total" ist kein EN-15804-Modulcode und daher standardmaessig nicht
        # gesetzt. --mit-total uebernimmt den ausgewiesenen Summenwert aus dem
        # PEP. Nicht selbst aufsummieren: der PEP-Total laeuft ueber A1-3 bis
        # C1-4 und schliesst Modul D aus (Lasten und Nutzen jenseits der
        # Systemgrenze). Ausserdem wuerde das Aufaddieren von auf drei
        # signifikante Stellen gerundeten Modulwerten die letzte Stelle
        # verfaelschen (8,55E-02 statt der ausgewiesenen 8,56E-02).
        if mit_total and (name, "Total") in werte:
            module.append(("Total", werte[(name, "Total")]))

        eintrag = etree.fromstring(etree.tostring(plan))
        # idShort vorerst immer setzen (macht die Bauplan-Pfade eindeutig);
        # konformitaet() entfernt ihn spaeter, sofern IDSHORT_IN_LISTEN False ist.
        idshort_setzen(eintrag, code.replace("-", "_"))

        qualifier_anhaengen(eintrag, "PEP/IndicatorStatus", status)

        ind = kind(eintrag, "lciaIndicator")
        setze(kind(ind, "indicatorCode"), code)
        setze(kind(ind, "indicatorName"), name)
        setze(kind(ind, "indicatorGroup"), gruppe)
        setze(kind(ind, "indicatorStandard"), setup.get("Impact assessment method", "EN 15804+A2"))

        cu = kind(eintrag, "characterizationUnit")
        setze(kind(cu, "unit"), einheiten.get(name, ""))

        setze(kind(eintrag, "referenceUnitForCalculation"), "piece")
        setze(kind(eintrag, "quantityOfReferenceUnit"), "1.0")

        pl = kind(eintrag, "lifeCyclePhaseValues")
        pw = liste_leeren(pl)
        for modul, wert in module:
            pe = etree.fromstring(etree.tostring(phasenplan))
            # WAGO-Konvention wie beim Indikator: Modulcode mit Unterstrich
            # (A1-3 -> A1_3). Der Template-Bauplan hat keinen idShort, die
            # Bestandsinstanz schon; hier wird er in jedem Fall gesetzt und
            # von konformitaet() nur entfernt, wenn IDSHORT_IN_LISTEN aus ist.
            idshort_setzen(pe, modul.replace("-", "_"))
            setze(kind(pe, "lifeCyclePhase"), modul)
            setze(kind(pe, "value"), zahl(wert))
            pw.append(pe)

        container.append(eintrag)
        geschrieben += 1
    return geschrieben


def pruefe(sm, offene):
    """Pflichtfelder ohne Wert sammeln (SMT/Cardinality = One bzw. OneToMany)."""
    def container(el):
        return el.find(Q("value")) if el.find(Q("value")) is not None else el.find(Q("submodelElements"))

    def lauf(el, pfad):
        inhalt = container(el)
        if inhalt is None:
            return
        for c in inhalt:
            tag = etree.QName(c).localname
            if tag not in ("submodelElementCollection", "submodelElementList",
                           "property", "multiLanguageProperty", "file"):
                continue
            ids = c.findtext(Q("idShort")) or "[]"
            p = f"{pfad}/{ids}"
            card = c.findtext(f"{Q('qualifiers')}/{Q('qualifier')}/{Q('value')}")
            pflicht = card in ("One", "OneToMany")
            if tag in ("submodelElementCollection", "submodelElementList"):
                inner = c.find(Q("value"))
                if pflicht and (inner is None or not len(inner)):
                    offene.append(p)
                else:
                    lauf(c, p)
            else:
                wert = c.find(Q("value"))
                if tag == "multiLanguageProperty":
                    leer = wert is None or not (wert.findtext(f"{Q('langStringTextType')}/{Q('text')}") or "").strip()
                else:
                    leer = wert is None or not (wert.text or "").strip()
                if pflicht and leer:
                    offene.append(p)

    lauf(sm, "")


# --------------------------------------------------------------- AASX-Ein/Ausgabe

def lade_submodel(aasx_pfad):
    with zipfile.ZipFile(aasx_pfad) as z:
        xml = z.read("aasx/data.xml")
    baum = etree.fromstring(xml)
    idshorts_bereinigen(baum)
    sm = baum.find(f"{Q('submodels')}/{Q('submodel')}")
    return baum, sm


SUPPL_REL = "http://admin-shell.io/aasx/relationships/aas-suppl"

# Part-Namen im Paket. Der fuehrende Slash entspricht der OPC-Schreibweise und
# dem Target der Relationship; einzelne Werkzeuge erwarten den relativen Pfad.
ANHANG_ABSOLUT = True


def anhang_pfad(datei):
    p = f"aasx/documents/{Path(datei).name}"
    return "/" + p if ANHANG_ABSOLUT else p


def content_types_ergaenzen(xml, endungen):
    """Fehlende Default-Eintraege in [Content_Types].xml nachtragen."""
    typen = {"pdf": "application/pdf", "png": "image/png", "jpg": "image/jpeg",
             "jpeg": "image/jpeg"}
    text = xml.decode("utf-8")
    for e in endungen:
        if f'Extension="{e}"' in text:
            continue
        eintrag = f'  <Default Extension="{e}" ContentType="{typen.get(e, "application/octet-stream")}" />\n'
        text = text.replace("</Types>", eintrag + "</Types>")
    return text.encode("utf-8")


def rels_ergaenzen(xml, ziele):
    """Beziehungen von aasx-origin auf die eingebetteten Dateien setzen.

    Vorhandene aas-suppl-Beziehungen auf /aasx/documents/... werden entfernt,
    weil die Dokumente der Vorlage nicht mit ins Ziel kopiert werden.
    """
    text = xml.decode("utf-8")
    text = re.sub(r'\s*<Relationship [^>]*Type="' + re.escape(SUPPL_REL)
                  + r'"[^>]*Target="/?aasx/documents/[^"]*"[^>]*/>', "", text)
    n = text.count("<Relationship ")
    for i, ziel in enumerate(ziele, start=1):
        if ziel in text:
            continue
        text = text.replace("</Relationships>",
                            f'  <Relationship Type="{SUPPL_REL}" Target="{ziel}" '
                            f'Id="Rsup{n + i}" />\n</Relationships>')
    return text.encode("utf-8")


def identifier(artikel, basis=None, kennung=None):
    """Teilmodell-, AAS- und Asset-Id nach dem WAGO-Schema (ID_MUSTER)."""
    felder = {"basis": (basis or BASIS_ID).rstrip("/"),
              "artikel": artikel,
              "kennung": kennung or SUBMODEL_KENNUNG}
    return {rolle: muster.format(**felder) for rolle, muster in ID_MUSTER.items()}


def schreibe_aasx(vorlage, baum, ziel, submodel_id, shell_id, asset_id, anhaenge=(),
                  aufraeumen_an=True):
    for el in baum.iter(Q("submodel")):
        el.find(Q("id")).text = submodel_id
        break
    for el in baum.iter(Q("assetAdministrationShell")):
        el.find(Q("id")).text = shell_id
        ai = el.find(Q("assetInformation"))
        if ai is not None and ai.find(Q("globalAssetId")) is not None:
            ai.find(Q("globalAssetId")).text = asset_id
        for key in el.iter(Q("key")):
            if key.findtext(Q("type")) == "Submodel":
                key.find(Q("value")).text = submodel_id
        break

    if aufraeumen_an:
        aufraeumen(baum)
    bericht = konformitaet(baum)
    xml = etree.tostring(baum, xml_declaration=True, encoding="utf-8", pretty_print=False)

    pfade = [anhang_pfad(a) for a in anhaenge]
    ersetzt = {"aasx/data.xml"}
    if anhaenge:
        # Wird eine bereits befuellte Instanz als Vorlage benutzt, liegen dort
        # schon data.xml.rels und ein PDF. Beide werden neu erzeugt - sonst
        # entstehen doppelte Zip-Eintraege bzw. zwei PDFs im Paket.
        ersetzt.add("aasx/_rels/data.xml.rels")
        with zipfile.ZipFile(vorlage) as quelle:
            ersetzt.update(i.filename for i in quelle.infolist()
                           if i.filename.startswith("aasx/documents/"))
    # OPC-Teile ([Content_Types].xml, Rels, aasx-origin) uebernehmen, um
    # Anhaenge ergaenzt: ohne Eintrag in Content-Types und Relationships ist
    # eine mitgelieferte Datei fuer AASX-Leser nicht sichtbar.
    with zipfile.ZipFile(vorlage) as quelle, \
            zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as z:
        for i in quelle.infolist():
            if i.filename in ersetzt:
                continue
            roh = quelle.read(i.filename)
            if anhaenge and i.filename == "[Content_Types].xml":
                roh = content_types_ergaenzen(
                    roh, {Path(a).suffix.lstrip(".").lower() for a in anhaenge})
            if anhaenge and i.filename == "aasx/_rels/aasx-origin.rels":
                roh = rels_ergaenzen(roh, pfade)
            z.writestr(i, roh)
        z.writestr("aasx/data.xml", xml)
        if anhaenge:
            # IDTA Part 5 laesst Supplementary-Relationships vom aas-spec-Part
            # ausgehen, andere Implementierungen suchen sie an aasx-origin.
            # Beide Beziehungen zu setzen ist OPC-konform und erspart die Wette,
            # welche Lesart das Zielwerkzeug hat.
            z.writestr("aasx/_rels/data.xml.rels", rels_ergaenzen(
                b'<?xml version="1.0" encoding="utf-8"?>\n'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/'
                b'package/2006/relationships">\n</Relationships>\n', pfade))
        for a in anhaenge:
            z.writestr(f"aasx/documents/{Path(a).name}", Path(a).read_bytes())
    return bericht


PFLICHT = ("One", "OneToMany")


def kardinalitaet(el):
    for q in el.findall(f"{Q('qualifiers')}/{Q('qualifier')}"):
        if q.findtext(Q("type")) == "SMT/Cardinality":
            return q.findtext(Q("value"))
    return None


def hat_inhalt(el):
    tag = etree.QName(el).localname
    wert = el.find(Q("value"))
    if tag in ("submodelElementCollection", "submodelElementList"):
        return wert is not None and len(wert) > 0
    if tag == "multiLanguageProperty":
        return wert is not None and bool(
            (wert.findtext(f"{Q('langStringTextType')}/{Q('text')}") or "").strip())
    return wert is not None and bool((wert.text or "").strip())


def aufraeumen(baum):
    """Leere optionale Elemente entfernen.

    Ein Element ohne aas:value ist zwar gueltig, viele Leser geben dafuer
    aber den Platzhalter ihrer Sprache aus (bei aas-core3 Python None), was
    in Auswertungen wie ein echter Wert aussieht. Pflichtfelder (Cardinality
    One / OneToMany) bleiben stehen - sie fehlen wirklich und sollen im
    Pruefbericht auffallen, statt lautlos zu verschwinden.

    Mehrere Durchlaeufe, weil eine Collection erst leer wird, nachdem ihre
    Kinder entfernt wurden (characterizationUnit/baseUnit/unitId).
    """
    entfernt = 0
    while True:
        weg = []
        for el in baum.iter():
            if etree.QName(el).localname not in (
                    "property", "multiLanguageProperty", "file",
                    "submodelElementCollection", "submodelElementList"):
                continue
            eltern = el.getparent()
            if eltern is None or etree.QName(eltern).localname not in (
                    "value", "submodelElements"):
                continue
            if not hat_inhalt(el) and (kardinalitaet(el) not in PFLICHT
                                       or LEERE_PFLICHT_ENTFERNEN):
                weg.append((eltern, el))
        if not weg:
            return entfernt
        for eltern, el in weg:
            eltern.remove(el)
            entfernt += 1


# --------------------------------------------------------------- Konformitaet

GLOBAL_KEYTYPES = {"GlobalReference"}


def _lang(el, tag, text, sprache="en"):
    """LangString-Element mit Sprache/Text anlegen."""
    ls = etree.SubElement(el, Q(tag))
    etree.SubElement(ls, Q("language")).text = sprache
    etree.SubElement(ls, Q("text")).text = text
    return ls


def _element_beschreibungen(baum):
    """semanticId -> (idShort, englische Beschreibung) aus den Submodel-Elementen."""
    out = {}
    for el in baum.iter():
        if etree.QName(el).localname not in (
                "submodel", "submodelElementCollection", "submodelElementList",
                "property", "multiLanguageProperty", "file", "range", "referenceElement"):
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
    """shortName auf die von IEC 61360 erlaubten 18 Zeichen kuerzen."""
    text = (text or "").strip()
    return text if len(text) <= maxlen else text[:maxlen].rstrip("._- ")


def konformitaet(baum):
    """Verstoesse gegen AAS-v3-Constraints beheben, die aus dem Template stammen.

    Rueckgabe: Liste von Berichtszeilen. Behoben werden:
      AASd-120  idShort an direkten Kindern einer SubmodelElementList
      AASd-122/124  ExternalReference mit Key-Typ ausserhalb GlobalReference
      AASc-3a-008  ConceptDescription ohne englische Definition
      AASc-3a-009  *_MEASURE-ConceptDescription ohne Einheit
      shortName laenger als 18 Zeichen
      semanticIds ohne zugehoerige ConceptDescription (minimale CD ergaenzen)
    """
    bericht = []

    # -- AASd-120
    if not IDSHORT_IN_LISTEN:
        n = 0
        for lst in baum.iter(Q("submodelElementList")):
            werte = lst.find(Q("value"))
            if werte is None:
                continue
            for c in werte:
                ids = c.find(Q("idShort"))
                if ids is not None:
                    c.remove(ids)
                    n += 1
        if n:
            bericht.append(f"AASd-120: {n} idShorts an Listenkindern entfernt")
    else:
        n = sum(1 for lst in baum.iter(Q("submodelElementList"))
                for c in (lst.find(Q("value")) if lst.find(Q("value")) is not None else [])
                if c.find(Q("idShort")) is not None)
        bericht.append(f"AASd-120: {n} idShorts an Listenkindern beibehalten (WAGO-Konvention)")

    # -- AASd-122 / AASd-124
    n = 0
    for ref in baum.iter():
        if etree.QName(ref).localname not in ("semanticId", "reference", "supplementalSemanticId",
                                              "valueId", "dataSpecification", "isCaseOf"):
            continue
        if ref.findtext(Q("type")) != "ExternalReference":
            continue
        for key in ref.findall(f"{Q('keys')}/{Q('key')}"):
            t = key.find(Q("type"))
            if t is not None and t.text not in GLOBAL_KEYTYPES:
                t.text = "GlobalReference"
                n += 1
    if n:
        bericht.append(f"AASd-122/124: {n} Key-Typen in ExternalReferences auf GlobalReference gesetzt")

    # -- ConceptDescriptions
    beschr = _element_beschreibungen(baum)
    cds_el = baum.find(Q("conceptDescriptions"))
    if cds_el is None:
        cds_el = etree.SubElement(baum, Q("conceptDescriptions"))
    vorhanden = {cd.findtext(Q("id")): cd for cd in cds_el.findall(Q("conceptDescription"))}

    n_def = n_kurz = n_unit = 0
    for cd_id, cd in vorhanden.items():
        inhalt = cd.find(f"{Q('embeddedDataSpecifications')}/{Q('embeddedDataSpecification')}"
                         f"/{Q('dataSpecificationContent')}/{Q('dataSpecificationIec61360')}")
        if inhalt is None:
            continue
        # shortName <= 18 Zeichen
        for ls in inhalt.findall(f"{Q('shortName')}/{Q('langStringShortNameTypeIec61360')}"):
            t = ls.find(Q("text"))
            if t is not None and t.text and len(t.text) > 18:
                t.text = _kurzname(t.text)
                n_kurz += 1
        # Einheit fuer *_MEASURE
        dtyp = inhalt.findtext(Q("dataType")) or ""
        if dtyp.endswith("_MEASURE") and inhalt.find(Q("unit")) is None:
            einheit = CD_EINHEITEN.get(cd_id)
            if einheit is None:
                bericht.append(f"AASc-3a-009: keine Einheit fuer {cd_id} bekannt - CD_EINHEITEN ergaenzen")
            else:
                u = etree.Element(Q("unit"))
                u.text = einheit
                # XSD-Reihenfolge: preferredName, shortName, unit, unitId, ...
                sn = inhalt.find(Q("shortName"))
                anker = sn if sn is not None else inhalt.find(Q("preferredName"))
                inhalt.insert(list(inhalt).index(anker) + 1, u)
                n_unit += 1
        # englische Definition
        defs = inhalt.find(Q("definition"))
        hat_en = defs is not None and any(
            ls.findtext(Q("language")) == "en" and (ls.findtext(Q("text")) or "").strip()
            for ls in defs.findall(Q("langStringDefinitionTypeIec61360")))
        if not hat_en:
            idshort, text = beschr.get(cd_id, (None, None))
            if not text:
                pn = inhalt.find(f"{Q('preferredName')}/{Q('langStringPreferredNameTypeIec61360')}")
                text = (pn.findtext(Q("text")) if pn is not None else None) or idshort or cd_id
            if defs is None:
                defs = etree.Element(Q("definition"))
                # XSD-Reihenfolge: ... sourceOfDefinition? dataType, definition, valueFormat ...
                dt = inhalt.find(Q("dataType"))
                anker = [Q("valueFormat"), Q("valueList"), Q("value"), Q("levelType")]
                pos = len(inhalt)
                for tag in anker:
                    e = inhalt.find(tag)
                    if e is not None:
                        pos = min(pos, list(inhalt).index(e))
                if dt is not None:
                    pos = list(inhalt).index(dt) + 1
                inhalt.insert(pos, defs)
            _lang(defs, "langStringDefinitionTypeIec61360", text)
            n_def += 1
    if n_kurz:
        bericht.append(f"IEC 61360: {n_kurz} shortNames auf 18 Zeichen gekuerzt")
    if n_unit:
        bericht.append(f"AASc-3a-009: {n_unit} Einheiten in ConceptDescriptions ergaenzt")
    if n_def:
        bericht.append(f"AASc-3a-008: {n_def} englische Definitionen aus Elementbeschreibungen ergaenzt")

    # -- fehlende ConceptDescriptions fuer verwendete semanticIds
    n = 0
    for sem, (idshort, text) in beschr.items():
        if sem in vorhanden or idshort is None:
            continue
        cd = etree.SubElement(cds_el, Q("conceptDescription"))
        etree.SubElement(cd, Q("idShort")).text = bereinige_idshort(idshort)
        etree.SubElement(cd, Q("id")).text = sem
        eds = etree.SubElement(cd, Q("embeddedDataSpecifications"))
        ed = etree.SubElement(eds, Q("embeddedDataSpecification"))
        ds = etree.SubElement(ed, Q("dataSpecification"))
        etree.SubElement(ds, Q("type")).text = "ExternalReference"
        keys = etree.SubElement(ds, Q("keys"))
        key = etree.SubElement(keys, Q("key"))
        etree.SubElement(key, Q("type")).text = "GlobalReference"
        etree.SubElement(key, Q("value")).text = \
            "https://admin-shell.io/DataSpecificationTemplates/DataSpecificationIec61360/3/0"
        dsc = etree.SubElement(ed, Q("dataSpecificationContent"))
        iec = etree.SubElement(dsc, Q("dataSpecificationIec61360"))
        pn = etree.SubElement(iec, Q("preferredName"))
        _lang(pn, "langStringPreferredNameTypeIec61360", idshort)
        sn = etree.SubElement(iec, Q("shortName"))
        _lang(sn, "langStringShortNameTypeIec61360", _kurzname(idshort))
        etree.SubElement(iec, Q("dataType")).text = "STRING"
        d = etree.SubElement(iec, Q("definition"))
        _lang(d, "langStringDefinitionTypeIec61360", text or idshort)
        vorhanden[sem] = cd
        n += 1
    if n:
        bericht.append(f"{n} ConceptDescriptions fuer semanticIds ohne CD ergaenzt")
    return bericht


def qualifier_anhaengen(el, typ, text):
    """Zusatzangabe maschinenlesbar an ein Element haengen."""
    if el is None:
        return
    qs = el.find(Q("qualifiers"))
    if qs is None:
        qs = etree.Element(Q("qualifiers"))
        el.insert(list(el).index(el.find(Q("valueType"))) if el.find(Q("valueType")) is not None
                  else len(el), qs)
    # AASd-021: je Typ nur ein Qualifier - vorhandenen ersetzen statt anhaengen
    for alt in qs.findall(Q("qualifier")):
        if alt.findtext(Q("type")) == typ:
            qs.remove(alt)
    q = etree.SubElement(qs, Q("qualifier"))
    etree.SubElement(q, Q("type")).text = typ
    etree.SubElement(q, Q("valueType")).text = "xs:string"
    etree.SubElement(q, Q("value")).text = text


# --------------------------------------------------------------- Hauptprogramm

def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel")
    ap.add_argument("--template", default="epd-type-iii-submodel-template.aasx")
    ap.add_argument("--referenz", default=None,
                    help="befuellte Bestandsinstanz, liefert Bauplaene fuer leere Listen")
    ap.add_argument("--variante", default="221-422")
    ap.add_argument("--alle-varianten", action="store_true")
    ap.add_argument("--out", default=".")
    ap.add_argument("--basis-id", default=BASIS_ID,
                    help="Praefix der WAGO-Identifier (Standard: %(default)s)")
    ap.add_argument("--submodel-kennung", default=SUBMODEL_KENNUNG,
                    help="Kurzname des Teilmodells in der Submodel-Id, "
                         "z. B. https://wago.com/ids/submodel/221-422/EPD (Standard: %(default)s)")
    ap.add_argument("--pdf", default=None,
                    help="Quell-PDF, wird als Anhang ins AASX-Paket gelegt")
    ap.add_argument("--anhang-relativ", action="store_true",
                    help="Anhangpfad ohne fuehrenden Slash (aasx/documents/...)")
    ap.add_argument("--behalte-leere", action="store_true",
                    help="leere optionale Elemente im Teilmodell belassen")
    ap.add_argument("--mit-total", action="store_true",
                    help="zusaetzlich zur Modulaufteilung die Summe als Phase 'Total'")
    ap.add_argument("--ohne-idshort-in-listen", action="store_true",
                    help="idShorts an Listenkindern entfernen (strikt AASd-120-konform; "
                         "Standard ist die WAGO-Konvention mit idShorts)")
    ap.add_argument("--behalte-leere-pflicht", action="store_true",
                    help="leere Pflichtfelder (Cardinality One) nicht entfernen")
    args = ap.parse_args()

    global ANHANG_ABSOLUT, IDSHORT_IN_LISTEN, LEERE_PFLICHT_ENTFERNEN
    ANHANG_ABSOLUT = not args.anhang_relativ
    IDSHORT_IN_LISTEN = not args.ohne_idshort_in_listen
    LEERE_PFLICHT_ENTFERNEN = not args.behalte_leere_pflicht

    daten = lade(args.excel)
    referenz_sm = lade_submodel(args.referenz)[1] if args.referenz else None

    varianten = (list(daten["produkte"]["Artikelnummer"]) if args.alle_varianten
                 else [args.variante])

    # Varianten, fuer die das PEP Faktoren, aber keine Masse angibt (221-2421
    # steht in der Faktorentabelle S. 14, nicht aber in der Produkttabelle
    # S. 5). Ohne Masse kann kein vollstaendiges Teilmodell entstehen.
    nur_faktoren = sorted(set(daten["faktoren"]["Variante"]) - set(daten["produkte"]["Artikelnummer"]))
    if nur_faktoren:
        print("Hinweis: Faktoren ohne Produktzeile (keine Masse im PEP), nicht erzeugt: "
              + ", ".join(nur_faktoren))
    ziel_ordner = Path(args.out)
    ziel_ordner.mkdir(parents=True, exist_ok=True)

    for artikel in varianten:
        baum, sm = lade_submodel(args.template)
        offene = []
        n = befuelle(sm, referenz_sm, daten, artikel, offene, args.pdf, args.mit_total)

        kennung = artikel.lower()
        ziel = ziel_ordner / f"wago-{kennung}-epd-submodel.aasx"
        ids = identifier(artikel, args.basis_id, args.submodel_kennung)
        bericht = schreibe_aasx(args.template, baum, ziel,
                      submodel_id=ids["submodel"],
                      shell_id=ids["aas"],
                      asset_id=ids["asset"],
                      anhaenge=[args.pdf] if args.pdf else (),
                      aufraeumen_an=not args.behalte_leere)

        quelle = "PEP-Werte" if artikel == "221-422" else "extrapoliert aus 221-422"
        print(f"{ziel.name:44s} {n:3d} LCIA-Indikatoren  ({quelle})")
        print(f"   Submodel-Id {ids['submodel']}  |  AAS-Id {ids['aas']}")
        if offene:
            was = "entfernt" if LEERE_PFLICHT_ENTFERNEN else "leer belassen"
            print(f"   Pflichtfelder ohne Wert im PEP ({len(offene)}, {was}):")
            for p in offene:
                print(f"      {p}")
        for zeile_b in bericht:
            print(f"   {zeile_b}")


if __name__ == "__main__":
    main()
