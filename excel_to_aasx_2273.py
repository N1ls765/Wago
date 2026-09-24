#!/usr/bin/env python3
"""Fuellt das (neue) EPD-Type-III-Teilmodell aus der Excel-Datei der 2273-Serie.

Aufruf:
    python excel_to_aasx_2273.py WAGO_PEP_2273.xlsx \\
        --pdf WAGO-00002-V01.01-EN.pdf --out ./aasx_2273

    # alle fuenf Varianten der Produktfamilie, hochgerechnet ueber die
    # Multiplikationsfaktoren aus Blatt "Faktoren_lang":
    python excel_to_aasx_2273.py WAGO_PEP_2273.xlsx --alle-varianten \\
        --pdf WAGO-00002-V01.01-EN.pdf --out ./aasx_2273

Gegenstueck zu excel_to_aasx.py, aber fuer das Teilmodell, das seit Commit
18c6789 als epd-type-iii-submodel-template.aasx im Repo liegt
(EnvironmentalResults / stageValues, LCAMethodology, EPDScope,
ContentDeclaration ... - siehe weitereInformationenWago/
EPD-Type-III-Submodel-User-Guide.md). Die generischen XML-, Paket- und
Konformitaetswerkzeuge werden aus excel_to_aasx.py importiert.

Ablauf
------
1. Template-AASX entpacken, aasx/data.xml parsen
2. Baupläne fuer alle Listeneintraege sichern (Template, ersatzweise die
   Referenzinstanz), danach ALLE Werte des Templates leeren. Das Template ist
   eine vorbefuellte 221-422-Instanz; ohne diesen Schritt blieben
   221-Werte (Produktbild, Deklarations-URL, Standorte ...) unbemerkt stehen.
3. Felder und Listen mit den 2273-Werten neu befuellen
4. Listentypen, Einheiten und idShorts konform machen (aas-core3)
5. data.xml ins AASX schreiben. Die Dateien der Vorlage unter aasx/files/
   (221er-Bild, 00001-PDF) werden NICHT uebernommen, stattdessen das
   2273-PDF und ggf. ein eigenes Produktbild (--bild) eingebettet.
6. Pruefbericht: Pflichtfelder ohne Wert im PEP

Lebenszyklusmodule
------------------
Das PEP weist aggregierte Spalten aus (A1-3, A4, A5, B1-7, C1-4, D). Das
Teilmodell modelliert Module einzeln mit Status (User Guide Kap. 9/10):
  A1-3  -> A1-A3
  B1-7  -> B6 (Joule-Verluste); B1-B5, B7 sind laut PEP Kap. 4.4 fuer das
           passive Bauteil nicht anwendbar -> NotRelevant, ohne Wert
  C1-4  -> C2-C4 (aggregiert); C1 liegt laut PEP Kap. 4.5 ausserhalb der
           Systemgrenze -> NotRelevant, ohne Wert
Die Zuordnung steht in STUFEN.

Pruefen mit aas-core3.0:
    from aas_core3 import xmlization, verification
    env = xmlization.environment_from_str(open("aasx/data.xml").read())
    print(list(verification.verify(env)))
Erwartet sind nur AASd-120-Meldungen (idShort an Listenkindern, mit WAGO
abgestimmt; --ohne-idshort-in-listen erzeugt die strikt konforme Form).
"""

import argparse
import re
import zipfile
from pathlib import Path

import pandas as pd
from lxml import etree

import excel_to_aasx as ex
from excel_to_aasx import (Q, finde, idshort_setzen, info_text, kind, leeren,
                           liste_leeren, setze, zahl)

# --------------------------------------------------------------- Einstellungen

TEMPLATE = "epd-type-iii-submodel-template.aasx"
REFERENZ = "wago-00001-v01-01-en-epd-submodel-instance.aasx"

# Status-Vokabular laut User Guide Kap. 10. Das Template selbst benutzt in
# DeclaredStages "Assessed"/"NotAssessed" - hier bei Bedarf umstellen.
STATUS_DEKLARIERT = "Declared"
STATUS_NICHT_RELEVANT = "NotRelevant"

# (stageCode im Teilmodell, Spalte im PEP bzw. None = nicht anwendbar)
STUFEN = [
    ("A1-A3", "A1-3"),
    ("A4", "A4"),
    ("A5", "A5"),
    ("B1", None), ("B2", None), ("B3", None), ("B4", None), ("B5", None),
    ("B6", "B1-7"),
    ("B7", None),
    ("C1", None),
    ("C2-C4", "C1-4"),
    ("D", "D"),
]

# Phasenbezeichnung im Faktorenblatt -> Spalte im PEP
FAKTOR_PHASE = {
    "A1-3_Manufacturing Stages": "A1-3",
    "A4_Distribution phases": "A4",
    "A5_Installation phases": "A5",
    "B6_Use stage": "B1-7",
    "C1-4_End-of-life stages": "C1-4",
    "D_Benefits and loads": "D",
    "Total": "Total",
}

UNIT_ID = "https://wg-epd.example.com/draft/v1/Value/Unit_{}"
INDIKATOR_CD = "https://wg-epd.example.com/draft/v1/Indicator/{}"

# --------------------------------------------------------------- Fachmapping

# PDF-Indikatorname -> (indicatorCode, resultCategory, indicatorName, unit,
#                       baseUnit, Status im PEP)
# Codes und Einheiten-Schreibweise folgen der Referenzinstanz (Registry-
# Notation, z. B. "m3 eq." statt "m3 World eq."), die acht dort fehlenden
# Sachbilanzindikatoren sind nach EN 15804+A2 ergaenzt. baseUnit nur bei
# abgeleiteten Einheiten (User Guide Kap. 8).
INDIKATOREN = {
    # verbindliche Wirkungsindikatoren (PDF S. 9)
    "Climate change - total": ("GWP-total", "ImpactIndicator", "Global Warming Potential (total)", "kg CO2 eq.", "kg", "mandatory"),
    "Climate change - fossil fuels": ("GWP-fossil", "ImpactIndicator", "Global Warming Potential (fossil)", "kg CO2 eq.", "kg", "mandatory"),
    "Climate change - biogenics": ("GWP-biogenic", "ImpactIndicator", "Global Warming Potential (biogenic)", "kg CO2 eq.", "kg", "mandatory"),
    "Climate change - land use and land use transformation":
        ("GWP-luluc", "ImpactIndicator", "Global Warming Potential (land use and land use change)", "kg CO2 eq.", "kg", "mandatory"),
    "Ozone depletion": ("ODP", "ImpactIndicator", "Ozone Depletion Potential", "kg CFC-11 eq.", "kg", "mandatory"),
    "Acidification (AP)": ("AP", "ImpactIndicator", "Acidification Potential", "mol H+ eq.", "mol", "mandatory"),
    "Freshwater eutrophication": ("EP-freshwater", "ImpactIndicator", "Eutrophication Potential (freshwater)", "kg P eq.", "kg", "mandatory"),
    "Marine aquatic eutrophication": ("EP-marine", "ImpactIndicator", "Eutrophication Potential (marine)", "kg N eq.", "kg", "mandatory"),
    "Terrestrial eutrophication": ("EP-terrestrial", "ImpactIndicator", "Eutrophication Potential (terrestrial)", "mol N eq.", "mol", "mandatory"),
    "Photochemical ozone formation": ("POCP", "ImpactIndicator", "Photochemical Ozone Creation Potential", "kg NMVOC eq.", "kg", "mandatory"),
    "Abiotic resource depletion - elements or resource depletion - metals and minerals":
        ("ADPE", "ImpactIndicator", "Abiotic Depletion Potential (elements)", "kg Sb eq.", "kg", "mandatory"),
    "Abiotic resource depletion - fossil fuels or resource depletion - fossils":
        ("ADPF", "ImpactIndicator", "Abiotic Depletion Potential (fossil fuels)", "MJ", "MJ", "mandatory"),
    "Water requirement": ("WDP", "ImpactIndicator", "Water Deprivation Potential", "m3 eq.", "m3", "mandatory"),
    # optionale Wirkungsindikatoren (PDF S. 11)
    "Total use of primary energy during the life cycle": ("Total-PE", "ImpactIndicator", "Total Primary Energy", "MJ", "MJ", "optional"),
    "Emission of fine particles": ("PM", "ImpactIndicator", "Particulate Matter Formation", "disease incidence", None, "optional"),
    "Ionizing radiation, human health": ("IR", "ImpactIndicator", "Ionising Radiation", "kBq U-235 eq.", "kBq", "optional"),
    "Ecotoxicity (fresh water)": ("ETP-fw", "ImpactIndicator", "Ecotoxicity (freshwater)", "CTUe", None, "optional"),
    "Human toxicity, carcinogenic effects": ("HTP-c", "ImpactIndicator", "Human Toxicity Potential (cancer)", "CTUh", None, "optional"),
    "Human toxicity, non-carcinogenic effects": ("HTP-nc", "ImpactIndicator", "Human Toxicity Potential (non-cancer)", "CTUh", None, "optional"),
    "Impacts related to land use/soil quality": ("SQP", "ImpactIndicator", "Soil Quality Potential", "dimensionless", None, "optional"),
    # Ressourceneinsatz (PDF S. 10)
    "Use of renewable primary energy, excluding renewable primary energy resources used as raw materials":
        ("PERE", "ResourceUse", "Renewable Primary Energy (excluding raw materials)", "MJ", "MJ", "mandatory"),
    "Use of renewable primary energy resources used as raw materials":
        ("PERM", "ResourceUse", "Renewable Primary Energy used as raw materials", "MJ", "MJ", "mandatory"),
    "Total use of renewable primary energy resources":
        ("PERT", "ResourceUse", "Total Renewable Primary Energy", "MJ", "MJ", "mandatory"),
    "Use of non-renewable primary energy, excluding non-renewable primary energy resources used as raw materials":
        ("PENRE", "ResourceUse", "Non-renewable Primary Energy (excluding raw materials)", "MJ", "MJ", "mandatory"),
    "Use of non-renewable primary energy resources used as raw materials":
        ("PENRM", "ResourceUse", "Non-renewable Primary Energy used as raw materials", "MJ", "MJ", "mandatory"),
    "Total use of non-renewable primary energy resources":
        ("PENRT", "ResourceUse", "Total Non-renewable Primary Energy", "MJ", "MJ", "mandatory"),
    "Use of secondary materials": ("SM", "ResourceUse", "Use of Secondary Materials", "kg", "kg", "mandatory"),
    "Use of renewable secondary fuels": ("RSF", "ResourceUse", "Use of Renewable Secondary Fuels", "MJ", "MJ", "mandatory"),
    "Use of non-renewable secondary fuels": ("NRSF", "ResourceUse", "Use of Non-renewable Secondary Fuels", "MJ", "MJ", "mandatory"),
    "Net use of fresh water": ("FW", "ResourceUse", "Net Use of Fresh Water", "m3", "m3", "mandatory"),
    # Abfall- und Ausgangsstroeme (PDF S. 10)
    "Hazardous waste disposed of": ("HWD", "Waste", "Hazardous Waste Disposed", "kg", "kg", "mandatory"),
    "Non-hazardous waste disposed of": ("NHWD", "Waste", "Non-Hazardous Waste Disposed", "kg", "kg", "mandatory"),
    "Radioactive waste disposed of": ("RWD", "Waste", "Radioactive Waste Disposed", "kg", "kg", "mandatory"),
    "Components for reuse": ("CRU", "OutputFlow", "Components for Re-use", "kg", "kg", "mandatory"),
    "Materials for recycling": ("MFR", "OutputFlow", "Materials for Recycling", "kg", "kg", "mandatory"),
    "Materials for energy recovery": ("MER", "OutputFlow", "Materials for Energy Recovery", "kg", "kg", "mandatory"),
    "Exported energy": ("EE", "OutputFlow", "Exported Energy", "MJ", "MJ", "mandatory"),
}

# Keine LCIA-Indikatoren, sondern Felder unter ContentDeclaration/biogenicCarbonContent
BIOGEN = {
    "Biogenic carbon content of the product": "biogenicCarbonContentProduct",
    "Biogenic carbon content of the associated packaging": "biogenicCarbonContentPackaging",
}

NORMEN = ["ISO 14025:2006", "ISO 14040", "ISO 14044", "EN 15804+A2",
          "EN 50693:2019", "XP C08-100-1:2016", "NF E38-500:2022"]
REGELN = ["PCR-ed4-EN-2021 09 06", "PSR-0001-ed4-EN-2022 11 16"]


# --------------------------------------------------------------- Werkzeuge

def wert(el):
    return el.findtext(Q("value")) if el is not None else None


def entferne(el):
    if el is not None and el.getparent() is not None:
        el.getparent().remove(el)


def klon(el):
    return etree.fromstring(etree.tostring(el))


# Elemente, die laut XSD hinter aas:qualifiers stehen
NACH_QUALIFIERS = {"embeddedDataSpecifications", "valueType", "value", "valueId", "contentType",
                   "orderRelevant", "semanticIdListElement", "typeValueListElement",
                   "valueTypeListElement"}


def qualifier_anhaengen(el, typ, text):
    """Wie excel_to_aasx.qualifier_anhaengen, legt aas:qualifiers aber an der
    XSD-Position an - auch bei Collections ohne valueType (dort haengte die
    Vorlage-Funktion ans Ende, hinter aas:value)."""
    if el is None:
        return
    if el.find(Q("qualifiers")) is None:
        pos = next((i for i, c in enumerate(el)
                    if etree.QName(c).localname in NACH_QUALIFIERS), len(el))
        el.insert(pos, etree.Element(Q("qualifiers")))
    ex.qualifier_anhaengen(el, typ, text)


def bauplan(quellen, pfad):
    """Ersten Listeneintrag unter pfad aus der ersten Quelle holen, die einen hat."""
    for sm in quellen:
        if sm is None:
            continue
        el = finde(sm, pfad)
        werte = el.find(Q("value")) if el is not None else None
        if werte is not None and len(werte):
            kopie = klon(werte[0])
            leeren(kopie)
            return kopie
    raise SystemExit(f"kein Bauplan fuer {pfad} gefunden (Template/Referenz pruefen)")


def liste_fuellen(liste, plan, eintraege, praefix, fuellen):
    """Liste leeren und je Eintrag eine Kopie des Bauplans anhaengen."""
    werte = liste_leeren(liste)
    for i, e in enumerate(eintraege):
        neu = klon(plan)
        idshort_setzen(neu, f"{praefix}_{i}" if praefix else None)
        fuellen(neu, e)
        werte.append(neu)


def text_liste(liste, plan, texte, praefix):
    liste_fuellen(liste, plan, texte, praefix, lambda el, t: setze(el, t))


def stufen_idshort(code):
    return code.replace("-", "_")


def sci(v):
    """xs:double in wissenschaftlicher Schreibweise, drei signifikante Stellen wie im PEP."""
    return f"{float(v):.2E}"


def masse_kg(text):
    """'0.71E-03 kg' -> 0.00071"""
    m = re.search(r"-?[\d.,]+(?:[Ee][+-]?\d+)?", str(text))
    return float(m.group(0).replace(",", ".")) if m else None


# --------------------------------------------------------------- Excel lesen

def lade(xlsx):
    x = pd.ExcelFile(xlsx)
    info = pd.read_excel(x, "Info").set_index("Feld")["Wert"].to_dict()
    return {
        "info": info,
        "produkte": pd.read_excel(x, "Produktfamilie"),
        "setup": pd.read_excel(x, "LCA_Setup").set_index("Parameter")["Beschreibung"].to_dict(),
        "indikatoren": pd.read_excel(x, "Indikatoren_lang"),
        "faktoren": pd.read_excel(x, "Faktoren_lang"),
    }


def werte_fuer_variante(daten, artikel):
    """{(Indikator, PEP-Spalte): Wert} fuer eine Variante.

    Referenzprodukt: Werte direkt aus dem PEP. Andere Varianten: Referenzwert
    x Multiplikationsfaktor (PEP Kap. 7). Ergebnis auf drei signifikante
    Stellen gerundet, wie die Werte im PEP.
    """
    il = daten["indikatoren"]
    basis = {(r["Indikator"], r["Phase"]): r["Wert"] for _, r in il.iterrows()}
    if artikel == info_text(daten["info"], "Referenzprodukt"):
        return basis

    fl = daten["faktoren"]
    teil = fl[fl["Variante"] == artikel]
    if teil.empty:
        raise SystemExit(f"keine Extrapolationsfaktoren fuer {artikel} im Blatt Faktoren_lang")
    faktoren = {(r["Indikator"], FAKTOR_PHASE[r["Phase"]]): r["Faktor"]
                for _, r in teil.iterrows() if r["Phase"] in FAKTOR_PHASE}
    return {k: float(f"{w * faktoren[k]:.2E}") for k, w in basis.items() if k in faktoren}


# --------------------------------------------------------------- Befuellen

def befuelle(sm, referenz_sm, daten, artikel, pdf=None, bild=None, mit_total=False):
    info, setup = daten["info"], daten["setup"]
    ref = info_text(daten["info"], "Referenzprodukt")
    produkte = daten["produkte"].set_index("Artikelnummer")
    zeile = produkte.loc[artikel]
    masse = float(zeile["Masse deklarierte Einheit (g)"]) / 1000.0
    extrapoliert = artikel != ref

    # -- Bauplaene sichern, bevor die Vorlage geleert wird
    quellen = (sm, referenz_sm)
    plaene = {p: bauplan(quellen, p) for p in (
        "identificationPublication/documentIds",
        "identificationPublication/languages",
        "manufacturerProduct/manufacturingSites",
        "manufacturerProduct/applicableJurisdictions",
        "LCAMethodology/DeclaredStages",
        "lca_StandardsPcr/standardsCompliance",
        "lca_StandardsPcr/productCategoryRules",
        "EPDScope/coveredProductReferences",
    )}
    # Indikator-Bauplan bevorzugt aus der Referenz: nur dort hat er unitId/baseUnit
    plaene["EnvironmentalResults"] = bauplan((referenz_sm, sm), "EnvironmentalResults")

    # -- Vorlage leeren: keine 221-Werte duerfen durchrutschen
    leeren(sm)

    ausgabe = info_text(info, "Ausgabedatum", "02-2026")          # "02-2026"
    monat, jahr = ausgabe.split("-")
    ausgabe_iso = f"{jahr}-{monat}-01"
    ablauf_iso = f"{int(jahr) + 5}-{monat}-01"

    # -- IdentificationPublication
    ip = "identificationPublication"

    def doc_id(el, _):
        setze(kind(el, "documentDomainId"), "PEP ecopassport")
        setze(kind(el, "documentIdentifier"), info_text(info, "PEP-Registriernummer"))
        setze(kind(el, "documentIsPrimary"), "true")
    liste_fuellen(finde(sm, f"{ip}/documentIds"), plaene[f"{ip}/documentIds"],
                  [None], "documentId", doc_id)
    setze(finde(sm, f"{ip}/programOperatorVersion"),
          re.search(r"V\.?(\d+\.\d+)", info_text(info, "PEP-Registriernummer")).group(1))
    setze(finde(sm, f"{ip}/dateOfIssue"), ausgabe_iso)
    setze(finde(sm, f"{ip}/validUntil"), ablauf_iso)
    for feld in ("dateOfIssue", "validUntil"):
        # Das PEP nennt nur Monat und Jahr; der Tag ist ergaenzt.
        qualifier_anhaengen(finde(sm, f"{ip}/{feld}"), "DatePrecision",
                            "month (day of month not stated in the PEP)")
    text_liste(finde(sm, f"{ip}/languages"), plaene[f"{ip}/languages"], ["en"], "language")
    setze(finde(sm, f"{ip}/sourceDataFormat"), "PEP Ecopassport PDF")
    setze(finde(sm, f"{ip}/declarationUrl"), info_text(info, "Deklarations-URL"))

    # -- ProgramOperatorVerification
    pv = "programOperatorVerification"
    setze(finde(sm, f"{pv}/programOperator/Company"), "Association P.E.P. (PEP ecopassport)")
    setze(finde(sm, f"{pv}/programOperator/EmailAddress/EmailAddress"),
          info_text(info, "Programmbetreiber E-Mail"))
    setze(finde(sm, f"{pv}/thirdPartyVerifier/verifierAccreditationId"),
          info_text(info, "Verifizierer-Akkreditierung", "VH44"))
    setze(finde(sm, f"{pv}/epdDeveloper/Company"),
          info_text(info, "EPD-Ersteller", "WAGO Kontakttechnik GmbH & Co. KG"))
    extern = info_text(info, "Verifizierung", "extern").lower().startswith("ext")
    setze(finde(sm, f"{pv}/verificationType"),
          "ThirdPartyVerification" if extern else "InternalVerification")
    # verificationDate: im PEP nicht angegeben -> bleibt leer und entfaellt

    # -- ManufacturerProduct
    mp = "manufacturerProduct"
    hersteller = "WAGO Kontakttechnik GmbH & Co. KG"
    setze(finde(sm, f"{mp}/manufacturerName"), hersteller)

    # Das PEP nennt "one WAGO facility in Germany", aber keinen Ort. Ein
    # Standorteintrag mit Land ist quellengestuetzt; Kennung, Ort und
    # Koordinaten der Vorlage (Minden bzw. illustrativer Lohnfertiger)
    # waeren es nicht und werden entfernt.
    def standort(el, _):
        setze(kind(el, "siteName"), "WAGO production facility, Germany")
        setze(kind(el, "siteType"), "Fixed")
        setze(kind(el, "siteOperator"), hersteller)
        entferne(kind(el, "SiteIdentifiers"))
        loc = kind(el, "Location")
        setze(kind(loc, "countryCode"), "DE")
        for weg in ("region", "locality", "latitude", "longitude", "LocationCodes"):
            entferne(kind(loc, weg))
    liste_fuellen(finde(sm, f"{mp}/manufacturingSites"), plaene[f"{mp}/manufacturingSites"],
                  [None], "ManufacturingSite", standort)

    setze(finde(sm, f"{mp}/productName"), info_text(info, "Produkt"))
    setze(finde(sm, f"{mp}/productArticleNumberOfManufacturer"), artikel)
    klemmstellen = int(artikel[-1])
    setze(finde(sm, f"{mp}/productDescription"),
          f"{zeile['Produktserie']}, {klemmstellen} clamping units. "
          + info_text(info, "Produktbeschreibung"))
    if bild:
        el = finde(sm, f"{mp}/productImage")
        setze(el, f"/aasx/files/{Path(bild).name}")
        ct = el.find(Q("contentType")) if el is not None else None
        if ct is not None:
            ct.text = {"png": "image/png"}.get(Path(bild).suffix.lstrip(".").lower(), "image/jpeg")
    setze(finde(sm, f"{mp}/massPerDeclaredUnit"), zahl(masse, 3))
    setze(finde(sm, f"{mp}/referenceServiceLife"), "30")
    # "001" = Welt (UN M49), wie in der Referenzinstanz: das PEP modelliert die
    # Nutzung weltweit (globaler Strommix in B6).
    text_liste(finde(sm, f"{mp}/applicableJurisdictions"),
               plaene[f"{mp}/applicableJurisdictions"], ["001"], "jurisdiction")
    setze(finde(sm, f"{mp}/manufacturingDescription"),
          "Manufactured at one WAGO production facility in Germany (injection moulding, "
          "metal forming, assembly). Electricity: average WAGO production mix Germany, "
          f"calendar year {info_text(info, 'Referenzjahr Daten', '2024')}.")

    # -- LCAMethodology
    lm = "LCAMethodology"
    setze(finde(sm, f"{lm}/lcaReferenceYear"), info_text(info, "Referenzjahr Daten", "2024"))
    setze(finde(sm, f"{lm}/timeRepresentativenessDescription"),
          f"Primary production data refer to calendar year {info_text(info, 'Referenzjahr Daten', '2024')}; "
          f"declaration valid until {ablauf_iso}")
    setze(finde(sm, f"{lm}/geographicalScope"), "GLO")
    setze(finde(sm, f"{lm}/geographicalDescription"),
          "Primary data from WAGO manufacturing facilities in Germany; background data at "
          "European or global level (ecoinvent); use stage modelled with a global electricity dataset")
    setze(finde(sm, f"{lm}/technologyDescription"),
          "Injection moulding, metal forming and assembly operations; background processes "
          "represent average state-of-the-art technologies as modelled in "
          + str(setup.get("LCA database", "ecoinvent")) + " ("
          + str(setup.get("LCA Software", "openLCA")) + ", "
          + str(setup.get("Impact assessment method", "EN 15804+A2")) + ")")
    setze(finde(sm, f"{lm}/lcaEnergyModel"), "DE")

    def stufe(el, s):
        code, spalte = s
        setze(kind(el, "stageCode"), code)
        setze(kind(el, "stageStatus"), STATUS_DEKLARIERT if spalte else STATUS_NICHT_RELEVANT)
    liste_fuellen(finde(sm, f"{lm}/DeclaredStages"), plaene[f"{lm}/DeclaredStages"],
                  STUFEN, None, stufe)
    for el in finde(sm, f"{lm}/DeclaredStages").find(Q("value")):
        idshort_setzen(el, "stage_" + stufen_idshort(wert(kind(el, "stageCode"))))
    b6 = finde(sm, f"{lm}/DeclaredStages/stage_B6")
    qualifier_anhaengen(b6, "ScopeNote",
                        "PEP column B1-B7 contains B6 only (Joule losses, contact resistance "
                        f"{info_text(info, 'Kontaktwiderstand (Mikroohm)')} micro-ohm, "
                        f"{info_text(info, 'Energie Modul B6 (J)')} J over the reference lifetime; "
                        "scale by the square of the actual current). B1-B5 and B7 not applicable.")
    qualifier_anhaengen(finde(sm, f"{lm}/DeclaredStages/stage_C1"), "ScopeNote",
                        "C1 (de-installation) is excluded from the system boundary: manual process "
                        "requiring no energy (PCR-ed4-EN-2021 09 06).")

    # -- lca_StandardsPcr
    text_liste(finde(sm, "lca_StandardsPcr/standardsCompliance"),
               plaene["lca_StandardsPcr/standardsCompliance"], NORMEN, "standard")
    text_liste(finde(sm, "lca_StandardsPcr/productCategoryRules"),
               plaene["lca_StandardsPcr/productCategoryRules"], REGELN, "rule")

    # -- epdDocument
    if pdf:
        setze(finde(sm, "epdDocument"), f"/aasx/files/{Path(pdf).name}")

    # -- EnvironmentalResults
    werte = werte_fuer_variante(daten, artikel)
    einheiten_pdf = dict(zip(daten["indikatoren"]["Indikator"], daten["indikatoren"]["Einheit"]))
    n = fuelle_ergebnisse(finde(sm, "EnvironmentalResults"), plaene["EnvironmentalResults"],
                          werte, einheiten_pdf, mit_total)
    if extrapoliert:
        qualifier_anhaengen(finde(sm, "EnvironmentalResults"), "PEP/ValueDerivation",
                            f"Extrapolated from reference product {ref} using the multiplication "
                            "factors of PEP section 7 (value of reference product x factor, "
                            "rounded to 3 significant digits)")

    # -- EPDScope
    setze(finde(sm, "EPDScope/scopeType"), "Family")
    alle = list(daten["produkte"]["Artikelnummer"])
    text_liste(finde(sm, "EPDScope/coveredProductReferences"),
               plaene["EPDScope/coveredProductReferences"], alle, "product")
    herkunft = ("values taken directly from the PEP" if not extrapoliert else
                f"values extrapolated from reference product {ref} with the PEP multiplication factors")
    setze(finde(sm, "EPDScope/scopeDescription"),
          f"Product family EPD for the WAGO {info_text(info, 'Produkt')} ({', '.join(alle)}), "
          f"reference product {ref}. This instance describes {artikel}: {herkunft}. "
          "Scope: cradle to grave (A1-A3, A4, A5, B6, C2-C4, D).")

    # -- ContentDeclaration
    verpackung = masse_kg(info_text(info, "Masse Verpackung je Einheit"))
    if not extrapoliert:
        setze(finde(sm, "ContentDeclaration/productMass"),
              zahl(masse_kg(info_text(info, "Masse Referenzprodukt")), 2))
    elif verpackung is not None:
        # Das PEP nennt die Produktmasse nur fuer das Referenzprodukt. Alle
        # Varianten haben laut PEP Kap. 7 dieselbe Verpackung, daher
        # Produktmasse = Masse der deklarierten Einheit - Verpackungsmasse.
        el = finde(sm, "ContentDeclaration/productMass")
        setze(el, zahl(masse - verpackung, 3))
        qualifier_anhaengen(el, "PEP/ValueDerivation",
                            "mass of declared unit minus packaging mass of the reference product "
                            "(identical packaging for all variants, PEP section 7)")
    for name, ziel in BIOGEN.items():
        spalte = (name, "Total")
        if spalte in werte:
            setze(finde(sm, f"ContentDeclaration/biogenicCarbonContent/{ziel}"), zahl(werte[spalte]))
    # recycledContentPercentage: im PEP nicht angegeben -> entfaellt

    # -- FunctionalUnit / DeclaredUnit
    setze(finde(sm, "FunctionalUnit/quantity"), "1")
    setze(finde(sm, "FunctionalUnit/unit"), "A")
    setze(finde(sm, "FunctionalUnit/description"), info_text(info, "Funktionelle Einheit"))
    setze(finde(sm, "DeclaredUnit/quantity"), "1")
    setze(finde(sm, "DeclaredUnit/unit"), "item")
    setze(finde(sm, "DeclaredUnit/description"),
          f"1 item {artikel} (product unit + packaging). Mass = {masse:g} kg per declared unit.")

    # -- Kopf des Teilmodells
    for dn in sm.findall(f"{Q('displayName')}/{Q('langStringNameType')}/{Q('text')}"):
        dn.text = f"WAGO {artikel} EPD Type III Submodel"
    admin = sm.find(Q("administration"))
    sem = sm.find(f"{Q('semanticId')}/{Q('keys')}/{Q('key')}/{Q('value')}")
    if admin is not None and sem is not None and admin.find(Q("templateId")) is not None:
        admin.find(Q("templateId")).text = sem.text
    return n


def fuelle_ergebnisse(liste, plan, werte, einheiten_pdf, mit_total):
    sv_liste = kind(plan, "stageValues")
    sv_plan = klon(sv_liste.find(Q("value"))[0])
    leeren(sv_plan)

    container = liste_leeren(liste)
    n = 0
    for name, (code, kategorie, titel, einheit, basis, status) in INDIKATOREN.items():
        if not any((name, spalte) in werte for _, spalte in STUFEN if spalte):
            continue
        eintrag = klon(plan)
        idshort_setzen(eintrag, code.replace("-", "_"))
        qualifier_anhaengen(eintrag, "PEP/IndicatorStatus", status)
        setze(kind(eintrag, "resultCategory"), kategorie)
        setze(kind(eintrag, "indicatorCode"), code)
        setze(kind(eintrag, "indicatorName"), titel)
        setze(kind(eintrag, "unit"), einheit)
        if einheit != einheiten_pdf.get(name):
            qualifier_anhaengen(kind(eintrag, "unit"), "PEP/UnitAsPrinted", einheiten_pdf.get(name, ""))
        setze(kind(eintrag, "unitId"), UNIT_ID.format(einheit.replace(" ", "_")))
        bu = kind(eintrag, "baseUnit")
        if basis:
            setze(kind(bu, "unit"), basis)
            setze(kind(bu, "unitId"), UNIT_ID.format(basis))
        else:
            entferne(bu)

        stufen = list(STUFEN) + ([("Total", "Total")] if mit_total else [])
        sv = liste_leeren(kind(eintrag, "stageValues"))
        for stufencode, spalte in stufen:
            el = klon(sv_plan)
            idshort_setzen(el, "sv_" + stufen_idshort(stufencode))
            setze(kind(el, "stageCode"), stufencode)
            if spalte and (name, spalte) in werte:
                setze(kind(el, "valueStatus"), STATUS_DEKLARIERT)
                setze(kind(el, "value"), sci(werte[(name, spalte)]))
            elif spalte:
                continue                  # Spalte im PEP leer: Stufe weglassen
            else:
                setze(kind(el, "valueStatus"), STATUS_NICHT_RELEVANT)
                entferne(kind(el, "value"))   # 0.0 != NotRelevant (User Guide Kap. 10)
            sv.append(el)
        container.append(eintrag)
        n += 1
    return n


# --------------------------------------------------------------- Konformitaet

def listen_typen_korrigieren(baum):
    """typeValueListElement/valueTypeListElement an den Inhalt anpassen (AASd-108/109).

    Im Template sind z. B. standardsCompliance als Liste von Collections
    deklariert, enthalten aber Properties; applicableJurisdictions fehlt der
    valueTypeListElement.
    """
    n = 0
    typname = {"property": "Property", "submodelElementCollection": "SubmodelElementCollection",
               "multiLanguageProperty": "MultiLanguageProperty", "file": "File",
               "submodelElementList": "SubmodelElementList"}
    for lst in baum.iter(Q("submodelElementList")):
        werte = lst.find(Q("value"))
        if werte is None or not len(werte):
            continue
        tatsaechlich = {etree.QName(c).localname for c in werte}
        if len(tatsaechlich) != 1:
            continue
        soll = typname.get(tatsaechlich.pop())
        tvle = lst.find(Q("typeValueListElement"))
        if tvle is not None and tvle.text != soll:
            tvle.text = soll
            n += 1
        vtle = lst.find(Q("valueTypeListElement"))
        if soll == "Property":
            vt = werte[0].findtext(Q("valueType")) or "xs:string"
            if vtle is None:
                vtle = etree.Element(Q("valueTypeListElement"))
                # XSD-Reihenfolge: ... typeValueListElement, valueTypeListElement, value
                lst.insert(list(lst).index(tvle) + 1, vtle)
            if vtle.text != vt:
                vtle.text = vt
                n += 1
        elif vtle is not None:
            lst.remove(vtle)
            n += 1
    return n


def cds_deduplizieren(baum):
    """Doppelte ConceptDescription-Ids entfernen, die erste bleibt.

    Das Template enthaelt die 29 Indikator-CDs zweimal: zuerst vollstaendig
    (REAL_MEASURE, Definition mit Einheit), dann als STRING-Rest. Ids muessen
    eindeutig sein; ausserdem repariert konformitaet() je Id nur eine Kopie.
    """
    gesehen, n = set(), 0
    for cd in baum.findall(f"{Q('conceptDescriptions')}/{Q('conceptDescription')}"):
        cid = cd.findtext(Q("id"))
        if cid in gesehen:
            entferne(cd)
            n += 1
        gesehen.add(cid)
    return n


def pruefe(sm):
    """Pflichtfelder (SMT/Cardinality One/OneToMany) ohne Wert sammeln."""
    offene = []

    def lauf(el, pfad):
        inhalt = el.find(Q("value"))
        if inhalt is None:
            inhalt = el.find(Q("submodelElements"))
        if inhalt is None:
            return
        for c in inhalt:
            tag = etree.QName(c).localname
            if tag not in ("submodelElementCollection", "submodelElementList",
                           "property", "multiLanguageProperty", "file"):
                continue
            p = f"{pfad}/{c.findtext(Q('idShort')) or '[]'}"
            if ex.kardinalitaet(c) in ex.PFLICHT and not ex.hat_inhalt(c):
                offene.append(p)
            elif tag in ("submodelElementCollection", "submodelElementList"):
                lauf(c, p)
    lauf(sm, "")
    return offene


# --------------------------------------------------------------- Paket

def schreibe_paket(vorlage, baum, ziel, ids, anhaenge):
    sm = baum.find(f"{Q('submodels')}/{Q('submodel')}")
    sm.find(Q("id")).text = ids["submodel"]
    for shell in baum.iter(Q("assetAdministrationShell")):
        shell.find(Q("id")).text = ids["aas"]
        gai = shell.find(f"{Q('assetInformation')}/{Q('globalAssetId')}")
        if gai is not None:
            gai.text = ids["asset"]
        for key in shell.iter(Q("key")):
            if key.findtext(Q("type")) == "Submodel":
                key.find(Q("value")).text = ids["submodel"]
        break

    ex.aufraeumen(baum)
    bericht = []
    n = listen_typen_korrigieren(baum)
    if n:
        bericht.append(f"AASd-108/109: {n} Listentypen an den Inhalt angepasst")
    n = cds_deduplizieren(baum)
    if n:
        bericht.append(f"{n} doppelte ConceptDescriptions (gleiche Id) entfernt")
    bericht += ex.konformitaet(baum)
    xml = etree.tostring(baum, xml_declaration=True, encoding="utf-8")

    pfade = [f"/aasx/files/{Path(a).name}" for a in anhaenge]
    with zipfile.ZipFile(vorlage) as quelle, \
            zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as z:
        for i in quelle.infolist():
            # Dateien der Vorlage (221er-Bild, 00001-PDF) nicht uebernehmen
            if (i.filename in ("aasx/data.xml", "aasx/_rels/data.xml.rels")
                    or i.filename.startswith(("aasx/files/", "aasx/documents/"))):
                continue
            roh = quelle.read(i.filename)
            if i.filename == "[Content_Types].xml":
                roh = ex.content_types_ergaenzen(
                    roh, {Path(a).suffix.lstrip(".").lower() for a in anhaenge})
            if i.filename == "aasx/_rels/aasx-origin.rels" and anhaenge:
                roh = ex.rels_ergaenzen(roh, pfade)
            z.writestr(i, roh)
        z.writestr("aasx/data.xml", xml)
        if anhaenge:
            # Supplementary-Beziehungen an aasx-origin und am aas-spec-Part,
            # wie in excel_to_aasx.py (beide Lesarten sind verbreitet).
            z.writestr("aasx/_rels/data.xml.rels", ex.rels_ergaenzen(
                b'<?xml version="1.0" encoding="utf-8"?>\n'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/'
                b'package/2006/relationships">\n</Relationships>\n', pfade))
        for a in anhaenge:
            z.writestr(f"aasx/files/{Path(a).name}", Path(a).read_bytes())
    return bericht


# --------------------------------------------------------------- Hauptprogramm

def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel")
    ap.add_argument("--template", default=TEMPLATE)
    ap.add_argument("--referenz", default=REFERENZ,
                    help="befuellte Instanz, liefert Bauplaene fuer im Template leere Listen")
    ap.add_argument("--variante", default=None, help="Standard: Referenzprodukt aus Blatt Info")
    ap.add_argument("--alle-varianten", action="store_true")
    ap.add_argument("--out", default="aasx_2273")
    ap.add_argument("--basis-id", default=ex.BASIS_ID,
                    help="Praefix der WAGO-Identifier (Standard: %(default)s)")
    ap.add_argument("--submodel-kennung", default=ex.SUBMODEL_KENNUNG,
                    help="Kurzname des Teilmodells in der Submodel-Id (Standard: %(default)s)")
    ap.add_argument("--pdf", default=None,
                    help="Quell-PDF fuer epdDocument (Standard: 'Quelldatei' aus Blatt Info, falls vorhanden)")
    ap.add_argument("--bild", default=None, help="Produktbild fuer productImage (jpg/png)")
    ap.add_argument("--mit-total", action="store_true",
                    help="zusaetzlich zu den Modulen den PEP-Summenwert als stageCode 'Total'")
    ap.add_argument("--ohne-idshort-in-listen", action="store_true",
                    help="idShorts an Listenkindern entfernen (strikt AASd-120-konform)")
    ap.add_argument("--behalte-leere-pflicht", action="store_true",
                    help="leere Pflichtfelder (Cardinality One) nicht entfernen")
    args = ap.parse_args()

    ex.IDSHORT_IN_LISTEN = not args.ohne_idshort_in_listen
    ex.LEERE_PFLICHT_ENTFERNEN = not args.behalte_leere_pflicht
    # Indikator-ConceptDescriptions des Templates haben REAL_MEASURE ohne
    # Einheit (AASc-3a-009); die Einheit ist die des Indikators.
    for code, _, _, einheit, _, _ in INDIKATOREN.values():
        ex.CD_EINHEITEN[INDIKATOR_CD.format(code)] = einheit

    daten = lade(args.excel)
    ref = info_text(daten["info"], "Referenzprodukt")
    referenz_sm = ex.lade_submodel(args.referenz)[1] if args.referenz and Path(args.referenz).exists() else None

    pdf = args.pdf
    if pdf is None:
        kandidat = Path(info_text(daten["info"], "Quelldatei"))
        if kandidat.name and kandidat.exists():
            pdf = str(kandidat)
    anhaenge = [a for a in (pdf, args.bild) if a]

    varianten = (list(daten["produkte"]["Artikelnummer"]) if args.alle_varianten
                 else [args.variante or ref])
    ziel_ordner = Path(args.out)
    ziel_ordner.mkdir(parents=True, exist_ok=True)

    for artikel in varianten:
        baum, sm = ex.lade_submodel(args.template)
        n = befuelle(sm, referenz_sm, daten, artikel, pdf, args.bild, args.mit_total)
        offene = pruefe(sm)
        ids = ex.identifier(artikel, args.basis_id, args.submodel_kennung)
        ziel = ziel_ordner / f"wago-{artikel.lower()}-epd-submodel.aasx"
        bericht = schreibe_paket(args.template, baum, ziel, ids, anhaenge)

        quelle = "PEP-Werte" if artikel == ref else f"extrapoliert aus {ref}"
        print(f"{ziel.name:44s} {n:3d} Indikatoren  ({quelle})")
        print(f"   Submodel-Id {ids['submodel']}  |  AAS-Id {ids['aas']}")
        if not pdf:
            print("   Hinweis: kein PDF angegeben - epdDocument (Pflichtfeld) bleibt leer")
        if offene:
            was = "entfernt" if ex.LEERE_PFLICHT_ENTFERNEN else "leer belassen"
            print(f"   Pflichtfelder ohne Wert im PEP ({len(offene)}, {was}):")
            for p in offene:
                print(f"      {p}")
        for zeile_b in bericht:
            print(f"   {zeile_b}")


if __name__ == "__main__":
    main()
