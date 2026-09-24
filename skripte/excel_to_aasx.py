#!/usr/bin/env python3
"""Generischer Generator: Standard-Excel (Vorlage v1) -> AASX mit Teilmodell "EPD Type III".

Aufruf (im Projektordner):
    python skripte/excel_to_aasx.py epd/221/EPD_221.xlsx
    python skripte/excel_to_aasx.py epd/2273/EPD_2273.xlsx --variante 2273-202

    Die AASX landen neben der Excel im Ordner aasx/ (--out fuer einen anderen Ort).
    Optionen: --template (Standard vorlagen/epd-type-iii-submodel-template.aasx),
    --variante (nur diesen Artikel), --nur-pruefen, --mit-total,
    --ohne-idshort-in-listen, --ohne-value-ids, --ohne-einheiten-ids, --behalte-leere

Für jede Zeile im Blatt Produkte (AASX erzeugen <> nein) entsteht eine AASX.
Das Skript enthält KEIN produkt- oder EPD-spezifisches Wissen: alles kommt aus
der Excel (Aufbau siehe epd_excel.py) bzw. aus dem AASX-Template.

Ablauf
------
1. Excel lesen und prüfen (epd_excel.pruefe). Geschlossene Vokabulare werden
   dabei aus den ConceptDescriptions des Templates gelesen. Bei Fehlern wird
   nichts erzeugt (Exit-Code 2).
2. Je Produkt: Template laden, Bauplaene für Listeneintraege sichern, ALLE
   Beispielwerte des Templates leeren, dann aus der Excel befuellen.
   Varianten mit Werteherkunft "Faktor": Wert = Referenzwert x Faktor,
   gerundet auf signifikante_stellen.
3. Leere Elemente entfernen, valueIds für Vokabularwerte setzen (optional),
   Template-Fehler beheben, die aas-core3 meldet (siehe konformitaet()).
4. Mit aas-core3 einlesen, prüfen und in XSD-Reihenfolge schreiben.
   Erwartet sind nur AASd-120-Meldungen (idShort an Listenkindern, WAGO-
   Konvention; --ohne-idshort-in-listen erzeugt die strikte Form). Alles
   andere ist ein Fehler (Exit-Code 1).
5. AASX schreiben: OPC-Teile der Vorlage, neue data.xml, PDF/Bild unter
   /aasx/files/.

Modellierungsregeln (User Guide EPD Type III)
---------------------------------------------
- Kein Wert für ein Modul -> valueStatus NotDeclared ohne value (nie 0).
- Module mit stageStatus NotAssessed erscheinen in stageValues als NotDeclared.
- Summenwerte (stageCode Total) nur mit --mit-total, nie in DeclaredStages.
- Hochgerechnete Werte tragen den Qualifier EPD/ValueDerivation.
"""

import argparse
import copy
import re
import sys
import zipfile
from decimal import Decimal
from pathlib import Path

from lxml import etree

import epd_excel

try:
    from aas_core3 import verification as aas_verification
    from aas_core3 import xmlization as aas_xml
except ImportError:                                   # pragma: no cover
    aas_verification = aas_xml = None

NS = "https://admin-shell.io/aas/3/0"


def Q(tag):
    return f"{{{NS}}}{tag}"


# --------------------------------------------------------------- Einstellungen

TEMPLATE = "epd-type-iii-submodel-template.aasx"
PROJEKT = Path(__file__).resolve().parent.parent          # skripte/.. = Projektordner
TEMPLATE_STANDARD = PROJEKT / "vorlagen" / TEMPLATE
IDSHORT_IN_LISTEN = True          # WAGO-Konvention (AASd-120), siehe --ohne-idshort-in-listen
VALUE_IDS = True                  # valueId fuer Vokabularwerte aus der CD-valueList
EINHEITEN_IDS = True              # unitId + baseUnit je Indikator (User Guide Kap. 8)
LEERE_ENTFERNEN = True

ID_MUSTER = {
    "submodel": "{basis}/submodel/{artikel}/{kennung}",
    "aas": "{basis}/aas/{artikel}",
    "asset": "{basis}/asset/{artikel}",
}

W = "https://wg-epd.example.com/draft/v1/"
# Vokabular-CDs des Templates je Property-idShort. Nicht ueber die semanticId
# der Property aufloesbar: scopeType verweist im Template auf .../ScopeType,
# die valueList steht aber an .../EPDScope/Type.
VOKABULAR_CDS = {
    "scopeType": W + "EPDScope/Type",
    "stageStatus": W + "StageStatus",
    "valueStatus": W + "ValueStatus",
    "siteType": W + "ManufacturingSite/Type",
    "resultCategory": W + "ResultCategory",
    "indicatorCode": W + "IndicatorCode",
    "stageCode": W + "StageCode",
}
GESCHLOSSEN = {"scopeType", "stageStatus", "valueStatus", "siteType", "resultCategory"}

SEM = {
    "documentId": "https://admin-shell.io/vdi/2770/1/0/DocumentId",
    "documentDomainId": "https://admin-shell.io/vdi/2770/1/0/DocumentDomainId",
    "documentIdentifier": "0173-1#02-ABI500#001",
    "documentIsPrimary": "https://admin-shell.io/idta/EPD/DocumentIsPrimary/1/0",
    "unit": W + "Unit",
    "unitId": W + "UnitId",
    "baseUnit": W + "BaseUnit",
}
UNIT_ID = W + "Value/Unit_{}"
INDIKATOR_CD = W + "Indicator/{}"

# Einheiten fuer ConceptDescriptions mit Datentyp *_MEASURE ohne Einheit
# (AASc-3a-009). Indikator-CDs bekommen ihre Einheit aus dem Blatt Indikatoren.
CD_EINHEITEN = {
    "https://admin-shell.io/idta/EPD/MassPerDeclaredUnit/1/0": "kg",
    "https://admin-shell.io/idta/EPD/ReferenceServiceLife/1/0": "a",
    "https://admin-shell.io/idta/EPD/BiogenicCarbon/Product/1/0": "kg C",
    "https://admin-shell.io/idta/EPD/BiogenicCarbon/Packaging/1/0": "kg C",
    W + "ProductMass": "kg",
    W + "RecycledContentPercentage": "%",
    W + "Quantity": "1",
    W + "IndicatorValue": "as 'unit' of the enclosing EnvironmentalResult",
}

# --------------------------------------------------------------- XML-Werkzeug

IDSHORT_MUSTER = re.compile(r"^[a-zA-Z][a-zA-Z0-9_]*$")
SME = ("property", "multiLanguageProperty", "file", "range", "blob", "referenceElement",
       "submodelElementCollection", "submodelElementList", "relationshipElement",
       "annotatedRelationshipElement", "entity", "operation", "capability", "basicEventElement")

warnungen = []
fehlende_pfade = []


def bereinige_idshort(text):
    """idShort auf [a-zA-Z][a-zA-Z0-9_]* bringen ('sv_A1-A3' -> 'sv_A1_A3')."""
    if text is None:
        return None
    neu = re.sub(r"[^a-zA-Z0-9_]", "_", text.strip())
    if not neu or not neu[0].isalpha():
        neu = "x" + neu
    return neu


def lname(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def inhalt(el):
    w = el.find(Q("value"))
    return w if w is not None else el.find(Q("submodelElements"))


def kind(el, idshort):
    """Direktes Kind nach idShort (Gross-/Kleinschreibung egal)."""
    if el is None:
        return None
    werte = inhalt(el)
    if werte is None:
        return None
    gesucht = bereinige_idshort(idshort).lower()
    for c in werte:
        if (bereinige_idshort(c.findtext(Q("idShort"))) or "").lower() == gesucht:
            return c
    return None


def finde(wurzel, pfad):
    el = wurzel
    for teil in pfad.split("/"):
        el = kind(el, teil)
        if el is None:
            return None
    return el


def pfad(sm, p):
    """Wie finde(), merkt sich aber Pfade, die das Template nicht hat."""
    el = finde(sm, p)
    if el is None and p not in fehlende_pfade:
        fehlende_pfade.append(p)
    return el


def setze(el, wert):
    """aas:value einer Property / MultiLanguageProperty / File setzen. Leere Werte: nichts tun."""
    if el is None or wert is None or str(wert).strip() == "":
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
    ids.text = bereinige_idshort(text)


def liste_leeren(el):
    werte = el.find(Q("value"))
    if werte is None:
        werte = etree.SubElement(el, Q("value"))
    for c in list(werte):
        werte.remove(c)
    return werte


def werte_entfernen(el):
    """Alle Nutzwerte unterhalb von el entfernen, verschachtelte Listen leeren."""
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
    if liste is None:
        return None
    w = liste.find(Q("value"))
    if w is None or not len(w):
        return None
    return copy.deepcopy(w[0])


def qualifier_anhaengen(el, typ, text):
    """Zusatzangabe maschinenlesbar anhaengen (AASd-021: je Typ ein Qualifier)."""
    if el is None or not text:
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
    etree.SubElement(q, Q("value")).text = str(text)


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


def neue_smc(idshort, sem, kinder, kardinalitaet=None):
    el = etree.Element(Q("submodelElementCollection"))
    if idshort:
        etree.SubElement(el, Q("idShort")).text = idshort
    if sem:
        el.append(referenz("semanticId", sem))
    if kardinalitaet:
        qualifier_anhaengen(el, "SMT/Cardinality", kardinalitaet)
    w = etree.SubElement(el, Q("value"))
    for k in kinder:
        w.append(k)
    return el


def listen_property(liste, wert, idshort, bauplan=None):
    if bauplan is not None and lname(bauplan) == "property":
        el = copy.deepcopy(bauplan)
        idshort_setzen(el, idshort)
        setze(el, wert)
        return el
    sem = liste.findtext(f"{Q('semanticIdListElement')}/{Q('keys')}/{Q('key')}/{Q('value')}")
    vt = liste.findtext(Q("valueTypeListElement")) or "xs:string"
    return neues_property(idshort, sem, vt, wert)


def dezimal(v, stellen=None):
    """xs:decimal ohne Exponent. stellen: signifikante Stellen (None = unverändert)."""
    if v is None:
        return None
    d = Decimal(f"{float(v):.{stellen - 1}E}") if stellen else Decimal(repr(float(v)))
    s = f"{d.normalize():f}"
    return "0" if s in ("-0", "0") else s


def double(v, stellen=None):
    """xs:double; stellen=None: kuerzeste exakte Darstellung (0.0124, 4.73e-07)."""
    if v is None:
        return None
    s = f"{float(v):.{stellen}g}" if stellen else repr(float(v))
    s = s[:-2] if s.endswith(".0") and "e" not in s else s
    return "0" if s in ("-0", "0", "-0.0") else s


# --------------------------------------------------------------- Template-Vokabular

def lade_template(aasx_pfad):
    with zipfile.ZipFile(aasx_pfad) as z:
        spec = spec_part(z)
        baum = etree.fromstring(z.read(spec))
    sm = baum.find(f"{Q('submodels')}/{Q('submodel')}")
    if sm is None:
        raise SystemExit(f"{aasx_pfad}: kein Submodel in {spec}")
    return baum, sm


def template_vokabular(baum):
    """{vokabular: [werte]} und {vokabular: {wert: valueId}} aus den CD-valueLists."""
    listen, ids = {}, {}
    cds = {cd.findtext(Q("id")): cd for cd in reversed(baum.findall(f".//{Q('conceptDescription')}"))}
    for name, cd_id in VOKABULAR_CDS.items():
        cd = cds.get(cd_id)
        if cd is None:
            continue
        paare = cd.findall(f".//{Q('valueReferencePair')}")
        if not paare:
            continue
        listen[name] = [p.findtext(Q("value")) for p in paare]
        ids[name] = {p.findtext(Q("value")): p.findtext(f"{Q('valueId')}/{Q('keys')}/{Q('key')}/{Q('value')}")
                     for p in paare}
    return listen, ids


def vokabular_abgleichen(tpl):
    """Excel-Auswahllisten (epd_excel.VOKABULAR) gegen das Template prüfen."""
    for name in GESCHLOSSEN:
        if name in tpl and set(tpl[name]) != set(epd_excel.VOKABULAR.get(name, [])):
            warnungen.append(f"Vokabular {name}: Template {tpl[name]} != Excel-Vorlage "
                             f"{epd_excel.VOKABULAR.get(name)} - epd_excel.VOKABULAR anpassen")
        if name not in tpl:
            warnungen.append(f"Vokabular {name}: keine valueList im Template gefunden, "
                             f"verwende Excel-Vorlage {epd_excel.VOKABULAR.get(name)}")


# --------------------------------------------------------------- Werte je Produkt

def werte_fuer(modell, produkt):
    """{(code, stage): (wert, einheit_quelle)} für ein Produkt."""
    ergebnisse, ref = modell["ergebnisse"], modell["referenz"]
    a = produkt["artikel"]
    if produkt["herkunft"] == "Deklariert":
        return {(c, s): v for (x, c, s), v in ergebnisse.items() if x == a}
    stellen = modell["epd"]["signifikante_stellen"] or 3
    out = {}
    for (x, c, s), (wert, einheit) in ergebnisse.items():
        if x != ref:
            continue
        f = modell["faktoren"][(a, c, s)]           # Vollstaendigkeit hat pruefe() sichergestellt
        out[(c, s)] = (float(f"{wert * f:.{stellen}g}"), einheit)
    return out


def kontext(modell, produkt, herkunft_text):
    e = modell["epd"]
    return {
        "artikel": produkt["artikel"],
        "referenzprodukt": modell["referenz"],
        "produktname": produkt.get("produktname") or e["produktname"],
        "serie": produkt.get("serie") or "",
        "registriernummer": e["registriernummer"],
        "masse_kg": dezimal(produkt["masse_kg"]),
        "lca_referenzjahr": e["lca_referenzjahr"],
        "herkunft": herkunft_text,
        "produkte": ", ".join(p["artikel"] for p in modell["produkte"]),
    }


def text(vorlage, ktx):
    return str(vorlage).format_map(ktx) if vorlage else None


# --------------------------------------------------------------- Befuellen

def bauplaene_sichern(sm):
    """Bauplaene für Listeneintraege kopieren, bevor die Beispielwerte geleert werden."""
    er = pfad(sm, "EnvironmentalResults")
    plan_er = erster_eintrag(er)
    if plan_er is None:
        raise SystemExit("Template: 'EnvironmentalResults' fehlt oder ist leer - richtige Vorlage?")
    plan_sv = erster_eintrag(kind(plan_er, "stageValues"))
    if plan_sv is None:
        raise SystemExit("Template: kein Bauplan für 'stageValues' im ersten EnvironmentalResult")
    plaene = {"result": plan_er, "stage": plan_sv,
              "stufe": erster_eintrag(finde(sm, "LCAMethodology/DeclaredStages"))}
    site = erster_eintrag(finde(sm, "manufacturerProduct/manufacturingSites"))
    plaene["site"] = site
    plaene["site_id"] = erster_eintrag(kind(site, "SiteIdentifiers")) if site is not None else None
    for listenpfad in ("lca_StandardsPcr/standardsCompliance", "lca_StandardsPcr/productCategoryRules",
                       "EPDScope/coveredProductReferences", "identificationPublication/languages",
                       "manufacturerProduct/applicableJurisdictions"):
        plaene[listenpfad] = erster_eintrag(finde(sm, listenpfad))
    for p in plaene.values():
        if p is not None:
            werte_entfernen(p)
            for q in p.iter(Q("qualifiers")):          # Beispiel-Qualifier raus, Kardinalitaet bleibt
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


def fuelle_standorte(sm, plaene, standorte):
    liste = pfad(sm, "manufacturerProduct/manufacturingSites")
    if liste is None:
        return
    c = liste_leeren(liste)
    if not standorte:
        return
    if plaene["site"] is None:
        warnungen.append("Template ohne Bauplan für manufacturingSites - Standorte nicht übernommen")
        return
    for i, s in enumerate(standorte, start=1):
        e = copy.deepcopy(plaene["site"])
        idshort_setzen(e, f"ManufacturingSite_{i:02d}")
        setze(kind(e, "siteName"), s.get("name"))
        setze(kind(e, "siteType"), s.get("typ"))
        setze(kind(e, "siteOperator"), s.get("betreiber"))
        loc = kind(e, "Location")
        setze(kind(loc, "countryCode"), s.get("land"))
        setze(kind(loc, "region"), s.get("region"))
        setze(kind(loc, "locality"), s.get("ort"))
        setze(kind(loc, "latitude"), dezimal(s.get("breite")))
        setze(kind(loc, "longitude"), dezimal(s.get("laenge")))
        ids = kind(e, "SiteIdentifiers")
        if ids is not None and s.get("id_wert") and plaene["site_id"] is not None:
            sid = copy.deepcopy(plaene["site_id"])
            idshort_setzen(sid, "SiteIdentifier_01")
            setze(kind(sid, "identifierScheme"), s.get("id_schema") or "Internal")
            setze(kind(sid, "identifierValue"), s.get("id_wert"))
            liste_leeren(ids).append(sid)
        c.append(e)


def befuelle(sm, modell, produkt, dateien, mit_total=False):
    e = modell["epd"]
    ref = modell["referenz"]
    a = produkt["artikel"]
    abgeleitet = produkt["herkunft"] == "Faktor"
    herkunft = (f"values derived from reference product {ref} using the multiplication factors of the EPD"
                if abgeleitet else "values declared directly in the EPD")
    ktx = kontext(modell, produkt, herkunft)
    stellen = e["signifikante_stellen"] or 3

    plaene = bauplaene_sichern(sm)
    werte_entfernen(sm)                                   # Beispielwerte des Templates raus

    # -- Kopf des Teilmodells
    for t in sm.findall(f"{Q('displayName')}/{Q('langStringNameType')}/{Q('text')}"):
        t.text = f"EPD Type III - {a}"
    for t in sm.findall(f"{Q('description')}/{Q('langStringTextType')}/{Q('text')}"):
        t.text = (f"Environmental Product Declaration (EPD Type III) for {a}, based on "
                  f"{e['dokument_domaene']} {e['registriernummer']}")
    admin = sm.find(Q("administration"))
    sem = sm.findtext(f"{Q('semanticId')}/{Q('keys')}/{Q('key')}/{Q('value')}")
    if admin is not None and sem and admin.find(Q("templateId")) is not None:
        admin.find(Q("templateId")).text = sem

    # -- IdentificationPublication
    ip = "identificationPublication"
    docs = pfad(sm, f"{ip}/documentIds")
    if docs is not None:
        liste_leeren(docs).append(neue_smc("documentId_0", SEM["documentId"], [
            neues_property("documentDomainId", SEM["documentDomainId"], "xs:string", e["dokument_domaene"], "One"),
            neues_property("documentIdentifier", SEM["documentIdentifier"], "xs:string", e["registriernummer"], "One"),
            neues_property("documentIsPrimary", SEM["documentIsPrimary"], "xs:boolean", "true", "ZeroToOne"),
        ]))
    setze(pfad(sm, f"{ip}/programOperatorVersion"), e["programm_version"])
    setze(pfad(sm, f"{ip}/dateOfIssue"), e["ausgabedatum"].isoformat())
    setze(pfad(sm, f"{ip}/validUntil"), e["gueltig_bis"].isoformat())
    if e["datumsgenauigkeit"] == "Monat":
        for feld in ("dateOfIssue", "validUntil"):
            qualifier_anhaengen(finde(sm, f"{ip}/{feld}"), "EPD/DatePrecision",
                                "month (day of month not stated in the source EPD)")
    fuelle_liste(sm, f"{ip}/languages", e["sprachen"], plaene, "language")
    setze(pfad(sm, f"{ip}/sourceDataFormat"), e["quelldatenformat"])
    setze(pfad(sm, f"{ip}/declarationUrl"), e["deklarations_url"])

    # -- ProgramOperatorVerification
    pv = "programOperatorVerification"
    setze(pfad(sm, f"{pv}/programOperator/Company"), e["programmbetreiber"])
    setze(pfad(sm, f"{pv}/programOperator/EmailAddress/EmailAddress"), e["programmbetreiber_email"])
    setze(pfad(sm, f"{pv}/thirdPartyVerifier/verifierAccreditationId"), e["verifizierer_akkreditierung"])
    setze(pfad(sm, f"{pv}/epdDeveloper/Company"), e["epd_ersteller"])
    setze(pfad(sm, f"{pv}/verificationType"), e["verifizierungsart"])
    if e["verifizierungsdatum"]:
        setze(pfad(sm, f"{pv}/verificationDate"), e["verifizierungsdatum"].isoformat())

    # -- ManufacturerProduct
    mp = "manufacturerProduct"
    setze(pfad(sm, f"{mp}/manufacturerName"), e["hersteller"])
    fuelle_standorte(sm, plaene, modell["standorte"])
    setze(pfad(sm, f"{mp}/productName"), text(ktx["produktname"], ktx))
    setze(pfad(sm, f"{mp}/productArticleNumberOfManufacturer"), a)
    setze(pfad(sm, f"{mp}/productDescription"), text(produkt.get("beschreibung"), ktx))
    if dateien.get("bild"):
        setze(pfad(sm, f"{mp}/productImage"), dateien["bild"][0])
        ct = finde(sm, f"{mp}/productImage").find(Q("contentType"))
        if ct is not None:
            ct.text = dateien["bild"][1]
    setze(pfad(sm, f"{mp}/massPerDeclaredUnit"), dezimal(produkt["masse_kg"]))
    setze(pfad(sm, f"{mp}/referenceServiceLife"), dezimal(e["referenzlebensdauer_jahre"]))
    fuelle_liste(sm, f"{mp}/applicableJurisdictions", e["geltungsbereich"] or [], plaene, "jurisdiction")
    setze(pfad(sm, f"{mp}/manufacturingDescription"), text(e["herstellungsbeschreibung"], ktx))

    # -- LCAMethodology
    lm = "LCAMethodology"
    setze(pfad(sm, f"{lm}/lcaReferenceYear"), e["lca_referenzjahr"])
    for feld, schl in (("timeRepresentativenessDescription", "zeitliche_repraesentativitaet"),
                       ("geographicalScope", "geografischer_scope"),
                       ("geographicalDescription", "geografische_beschreibung"),
                       ("technologyDescription", "technologie_beschreibung"),
                       ("lcaEnergyModel", "energiemodell")):
        setze(pfad(sm, f"{lm}/{feld}"), text(e[schl], ktx))
    stufen = pfad(sm, f"{lm}/DeclaredStages")
    if stufen is not None and plaene["stufe"] is not None:
        c = liste_leeren(stufen)
        for m in modell["module"]:
            el = copy.deepcopy(plaene["stufe"])
            idshort_setzen(el, f"stage_{m['stageCode']}")
            setze(kind(el, "stageCode"), m["stageCode"])
            setze(kind(el, "stageStatus"), m["stageStatus"])
            qualifier_anhaengen(el, "EPD/StageNote", text(m.get("hinweis"), ktx))
            c.append(el)

    # -- lca_StandardsPcr
    fuelle_liste(sm, "lca_StandardsPcr/standardsCompliance", modell["normen"], plaene, "standard")
    fuelle_liste(sm, "lca_StandardsPcr/productCategoryRules", modell["regeln"], plaene, "rule")

    # -- epdDocument
    if dateien.get("pdf"):
        el = pfad(sm, "epdDocument")
        setze(el, dateien["pdf"][0])
        if el is not None:
            ct = el.find(Q("contentType"))
            if ct is None:
                ct = etree.SubElement(el, Q("contentType"))
            ct.text = dateien["pdf"][1]

    # -- EnvironmentalResults
    werte = werte_fuer(modell, produkt)
    n = fuelle_ergebnisse(sm, plaene, modell, werte, mit_total, stellen if abgeleitet else None)
    if abgeleitet:
        qualifier_anhaengen(finde(sm, "EnvironmentalResults"), "EPD/ValueDerivation",
                            f"Derived: values of reference product {ref} multiplied by the multiplication "
                            f"factors for {a} given in the EPD, rounded to {stellen} significant digits.")

    # -- EPDScope
    setze(pfad(sm, "EPDScope/scopeType"), e["scope_typ"])
    fuelle_liste(sm, "EPDScope/coveredProductReferences", [p["artikel"] for p in modell["produkte"]],
                 plaene, "product")
    setze(pfad(sm, "EPDScope/scopeDescription"), text(e["scope_beschreibung"], ktx))

    # -- ContentDeclaration
    cd = "ContentDeclaration"
    setze(pfad(sm, f"{cd}/productMass"), dezimal(produkt.get("produktmasse_kg")))
    setze(pfad(sm, f"{cd}/recycledContentPercentage"), dezimal(e["recyclinganteil_prozent"]))
    setze(pfad(sm, f"{cd}/biogenicCarbonContent/biogenicCarbonContentProduct"),
          dezimal(produkt.get("biogen_produkt")))
    setze(pfad(sm, f"{cd}/biogenicCarbonContent/biogenicCarbonContentPackaging"),
          dezimal(produkt.get("biogen_verpackung")))

    # -- FunctionalUnit / DeclaredUnit
    for teil, schl in (("FunctionalUnit", "fe"), ("DeclaredUnit", "de")):
        setze(pfad(sm, f"{teil}/quantity"), dezimal(e[f"{schl}_menge"]))
        setze(pfad(sm, f"{teil}/unit"), e[f"{schl}_einheit"])
        setze(pfad(sm, f"{teil}/description"), text(e[f"{schl}_beschreibung"], ktx))
    return n


def einheiten_ergaenzen(eintrag, einheit, basis):
    """unitId und baseUnit (User Guide Kap. 8) hinter 'unit' einfuegen."""
    unit = kind(eintrag, "unit")
    if unit is None:
        return
    neu = [neues_property("unitId", SEM["unitId"], "xs:string", UNIT_ID.format(einheit.replace(" ", "_")),
                          "ZeroToOne")]
    if basis:
        neu.append(neue_smc("baseUnit", SEM["baseUnit"], [
            neues_property("unit", SEM["unit"], "xs:string", basis, "One"),
            neues_property("unitId", SEM["unitId"], "xs:string", UNIT_ID.format(basis.replace(" ", "_")), "One"),
        ], "ZeroToOne"))
    for el in reversed(neu):
        alt = kind(eintrag, el.findtext(Q("idShort")))
        if alt is not None:
            alt.getparent().remove(alt)
        unit.addnext(el)


def fuelle_ergebnisse(sm, plaene, modell, werte, mit_total, stellen):
    container = liste_leeren(finde(sm, "EnvironmentalResults"))
    stufen = [(m["stageCode"], m["stageStatus"], m.get("hinweis")) for m in modell["module"]]
    if mit_total:
        stufen.append((epd_excel.SUMMEN_STUFE, "Assessed", None))
    n = 0
    for code, ind in modell["indikatoren"].items():
        if not any((code, s) in werte for s, _, _ in stufen):
            continue
        e = copy.deepcopy(plaene["result"])
        idshort_setzen(e, code)
        for t in e.findall(f"{Q('description')}/{Q('langStringTextType')}/{Q('text')}"):
            t.text = f"{ind['name']} ({code})"
        for t in e.findall(f"{Q('displayName')}/{Q('langStringNameType')}/{Q('text')}"):
            t.text = code
        qualifier_anhaengen(e, "EPD/IndicatorStatus", ind.get("status"))
        setze(kind(e, "resultCategory"), ind["kategorie"])
        setze(kind(e, "indicatorCode"), code)
        setze(kind(e, "indicatorName"), ind["name"])
        setze(kind(e, "unit"), ind["einheit"])
        gedruckt = {w[1] for (c, _), w in werte.items() if c == code and w[1]}
        if gedruckt - {ind["einheit"]}:
            qualifier_anhaengen(kind(e, "unit"), "EPD/UnitAsPrinted", " | ".join(sorted(gedruckt)))
        if EINHEITEN_IDS:
            einheiten_ergaenzen(e, ind["einheit"], ind.get("basiseinheit"))
        c = liste_leeren(kind(e, "stageValues"))
        for stage, status, hinweis in stufen:
            wert = werte.get((code, stage))
            if stage == epd_excel.SUMMEN_STUFE and wert is None:
                continue
            sv = copy.deepcopy(plaene["stage"])
            idshort_setzen(sv, f"sv_{stage}")
            setze(kind(sv, "stageCode"), stage)
            if wert is not None and status == "Assessed":
                setze(kind(sv, "valueStatus"), "Declared")
                setze(kind(sv, "value"), double(wert[0], stellen))
            else:
                setze(kind(sv, "valueStatus"), "NotDeclared")
                v = kind(sv, "value")                  # Guide Kap. 10: kein Wert statt 0
                if v is not None:
                    v.getparent().remove(v)
            c.append(sv)
        container.append(e)
        n += 1
    return n


# --------------------------------------------------------------- Pflichtfelder / Aufraeumen

PFLICHT = ("One", "OneToMany")


def kardinalitaet(el):
    for q in el.findall(f"{Q('qualifiers')}/{Q('qualifier')}"):
        if q.findtext(Q("type")) == "SMT/Cardinality":
            return q.findtext(Q("value"))
    return None


def hat_inhalt(el):
    tag = lname(el)
    wert = el.find(Q("value"))
    if tag in ("submodelElementCollection", "submodelElementList"):
        return wert is not None and len(wert) > 0
    if tag == "multiLanguageProperty":
        return wert is not None and bool((wert.findtext(f"{Q('langStringTextType')}/{Q('text')}") or "").strip())
    return wert is not None and bool((wert.text or "").strip())


def irgendein_wert(el):
    return any(lname(x) in ("property", "multiLanguageProperty", "file", "range") and hat_inhalt(x)
               for x in el.iter())


def offene_pflichtfelder(sm):
    """Pflichtfelder (Cardinality One/OneToMany) ohne Wert, relativ zum Elternelement."""
    offene = []

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
    return offene


def aufraeumen(baum):
    """Leere Elemente entfernen, mehrere Durchlaeufe (Collections werden erst spaeter leer)."""
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
            if not hat_inhalt(el):
                weg.append((eltern, el))
        if not weg:
            return entfernt
        for eltern, el in weg:
            eltern.remove(el)
            entfernt += 1


def value_ids_setzen(sm, vok_ids):
    """valueId für Properties mit Vokabularwert setzen (wie in der WAGO-Referenzinstanz)."""
    n = 0
    for p in sm.iter(Q("property")):
        tabelle = vok_ids.get(p.findtext(Q("idShort")))
        wert = p.findtext(Q("value"))
        if not tabelle or wert not in tabelle or not tabelle[wert]:
            continue
        alt = p.find(Q("valueId"))
        if alt is not None:
            p.remove(alt)
        p.append(referenz("valueId", tabelle[wert]))
        n += 1
    return n


# --------------------------------------------------------------- Konformitaet

def _lang(el, tag, text_, sprache="en"):
    ls = etree.SubElement(el, Q(tag))
    etree.SubElement(ls, Q("language")).text = sprache
    etree.SubElement(ls, Q("text")).text = text_
    return ls


def _element_beschreibungen(baum):
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


def _kurzname(text_, maxlen=18):
    text_ = (text_ or "").strip()
    return text_ if len(text_) <= maxlen else text_[:maxlen].rstrip("._- ")


def listen_typen(baum):
    """AASd-108/109: typeValueListElement / valueTypeListElement an den Inhalt anpassen."""
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
    n = 0
    for el in baum.iter(Q("idShort")):
        if el.text and not IDSHORT_MUSTER.match(el.text):
            el.text = bereinige_idshort(el.text)
            n += 1
    if n:
        bericht.append(f"{n} ungültige idShorts bereinigt")

    if not IDSHORT_IN_LISTEN:                             # AASd-120
        n = 0
        for lst in baum.iter(Q("submodelElementList")):
            w = lst.find(Q("value"))
            for c in (w if w is not None else []):
                ids = c.find(Q("idShort"))
                if ids is not None:
                    c.remove(ids)
                    n += 1
        if n:
            bericht.append(f"AASd-120: {n} idShorts an Listenkindern entfernt")

    n = listen_typen(baum)
    if n:
        bericht.append(f"AASd-108/109: {n} Typangaben von Listen an den Inhalt angepasst")

    n = 0                                                 # AASd-122/124
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

    cds_el = baum.find(Q("conceptDescriptions"))
    if cds_el is None:
        cds_el = etree.SubElement(baum, Q("conceptDescriptions"))
    gesehen, n = set(), 0
    for cd in list(cds_el):
        i = cd.findtext(Q("id"))
        if i in gesehen:
            cds_el.remove(cd)
            n += 1
        gesehen.add(i)
    if n:
        bericht.append(f"{n} doppelte ConceptDescription-Ids entfernt")

    benutzt = set()
    for el in baum.iter(Q("key")):
        if any(lname(x) == "dataSpecification" for x in el.iterancestors()):
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
                bericht.append(f"AASc-3a-009: keine Einheit für {cd_id} bekannt - CD_EINHEITEN ergänzen")
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
            idshort, text_ = beschr.get(cd_id, (None, None))
            if not text_:
                pn = iec.find(f"{Q('preferredName')}/{Q('langStringPreferredNameTypeIec61360')}")
                text_ = (pn.findtext(Q("text")) if pn is not None else None) or idshort or cd_id
            if defs is None:
                defs = etree.SubElement(iec, Q("definition"))
            _lang(defs, "langStringDefinitionTypeIec61360", text_)
            n_def += 1
    if n_kurz:
        bericht.append(f"IEC 61360: {n_kurz} shortNames auf 18 Zeichen gekürzt")
    if n_unit:
        bericht.append(f"AASc-3a-009: {n_unit} Einheiten in ConceptDescriptions ergänzt")
    if n_def:
        bericht.append(f"AASc-3a-008: {n_def} englische Definitionen ergänzt")

    n = 0
    for sem, (idshort, text_) in beschr.items():
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
        _lang(etree.SubElement(iec, Q("definition")), "langStringDefinitionTypeIec61360", text_ or idshort)
        vorhanden[sem] = cd
        n += 1
    if n:
        bericht.append(f"{n} ConceptDescriptions für semanticIds ohne CD ergänzt")
    return bericht


def normalisieren(baum):
    """Mit aas-core3 einlesen, prüfen, in XSD-Reihenfolge schreiben. -> (xml, befunde|None)"""
    roh = etree.tostring(baum, encoding="unicode")
    if aas_xml is None:
        return etree.tostring(baum, xml_declaration=True, encoding="utf-8", pretty_print=True), None
    env = aas_xml.environment_from_str(roh)
    befunde = [f"{f.path}: {f.cause}" for f in aas_verification.verify(env)]
    neu = etree.fromstring(aas_xml.to_str(env).encode("utf-8"))
    return etree.tostring(neu, xml_declaration=True, encoding="utf-8", pretty_print=True), befunde


# --------------------------------------------------------------- AASX-Paket

REL_NS = "http://admin-shell.io/aasx/relationships/"
CONTENT_TYPES = {"pdf": "application/pdf", "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                 "xml": "application/xml"}


def spec_part(z):
    for name in z.namelist():
        if name.endswith("aasx-origin.rels"):
            roh = z.read(name).decode("utf-8")
            m = (re.search(r'Type="[^"]*aas-spec"[^>]*Target="/?([^"]+)"', roh)
                 or re.search(r'Target="/?([^"]+)"[^>]*Type="[^"]*aas-spec"', roh))
            if m:
                return m.group(1)
    return "aasx/data.xml"


def paketname(datei):
    return re.sub(r"[^A-Za-z0-9._-]", "_", Path(datei).name)


def content_types_ergaenzen(xml, endungen):
    t = xml.decode("utf-8")
    for e in sorted(endungen):
        if re.search(rf'Extension="{e}"', t, re.I):
            continue
        t = t.replace("</Types>", f'<Default Extension="{e}" '
                                  f'ContentType="{CONTENT_TYPES.get(e, "application/octet-stream")}"/></Types>')
    return t.encode("utf-8")


def schreibe_aasx(vorlage, xml, ziel, dateien):
    """AASX schreiben. dateien: {Part-Name ohne fuehrenden Slash: Bytes}.

    Übernommen: OPC-Teile der Vorlage. Nicht übernommen: data.xml, deren Rels,
    Dateien der Vorlage unter aasx/files bzw. aasx/documents. Die Relationships
    werden auf die Form aus IDTA Part 5 (http://admin-shell.io/aasx/...)
    umgestellt, sonst findet Eclipse BaSyx kein AAS im Paket.
    """
    with zipfile.ZipFile(vorlage) as quelle:
        spec = spec_part(quelle)
        spec_rels = str(Path(spec).parent / "_rels" / (Path(spec).name + ".rels")).replace("\\", "/")
        with zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as z:
            for i in quelle.infolist():
                n = i.filename
                if n in (spec, spec_rels) or n.startswith(("aasx/files/", "aasx/documents/")):
                    continue
                roh = quelle.read(n)
                if n == "[Content_Types].xml":
                    roh = content_types_ergaenzen(roh, {Path(d).suffix.lstrip(".").lower() for d in dateien})
                elif n.endswith(".rels"):
                    roh = re.sub(rb"https?://(?:www\.)?admin-shell\.io/aasx/relationships/", REL_NS.encode(), roh)
                z.writestr(i, roh)
            z.writestr(spec, xml)
            if dateien:
                rels = ['<?xml version="1.0" encoding="utf-8"?>'
                        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
                for i, part in enumerate(dateien, start=1):
                    rels.append(f'<Relationship Type="{REL_NS}aas-suppl" Target="/{part}" Id="Rsuppl{i}"/>')
                rels.append("</Relationships>")
                z.writestr(spec_rels, "".join(rels))
            for part, daten in dateien.items():
                z.writestr(part, daten)


def identifier(artikel, basis, kennung):
    felder = {"basis": basis.rstrip("/"), "artikel": artikel, "kennung": kennung}
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


def anhang(basis_ordner, datei):
    """(Part-Name, Pfad in der AASX, Content-Type, Bytes) für eine Datei relativ zur Excel."""
    p = Path(datei)
    p = p if p.is_absolute() else basis_ordner / p
    if not p.is_file():
        return None
    part = f"aasx/files/{paketname(p)}"
    endung = p.suffix.lstrip(".").lower()
    return part, "/" + part, CONTENT_TYPES.get(endung, "application/octet-stream"), p.read_bytes()


# --------------------------------------------------------------- Hauptprogramm

def main(argv=None):
    global IDSHORT_IN_LISTEN, VALUE_IDS, EINHEITEN_IDS, LEERE_ENTFERNEN
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel", help="ausgefüllte Excel nach Vorlage v1 (epd_excel.py)")
    ap.add_argument("--template", default=None,
                    help=f"AASX-Template (Standard: vorlagen/{TEMPLATE})")
    ap.add_argument("--out", default=None, help="Zielordner (Standard: aasx/ neben der Excel)")
    ap.add_argument("--variante", action="append", help="nur diese Artikelnummer(n) erzeugen")
    ap.add_argument("--nur-pruefen", action="store_true", help="Excel prüfen, nichts erzeugen")
    ap.add_argument("--mit-total", action="store_true", help="Summenwerte als stageCode 'Total' ausgeben")
    ap.add_argument("--ohne-idshort-in-listen", action="store_true",
                    help="idShorts an Listenkindern entfernen (strikt AASd-120-konform)")
    ap.add_argument("--ohne-value-ids", action="store_true", help="keine valueIds für Vokabularwerte")
    ap.add_argument("--ohne-einheiten-ids", action="store_true", help="kein unitId/baseUnit je Indikator")
    ap.add_argument("--behalte-leere", action="store_true", help="leere Elemente nicht entfernen")
    args = ap.parse_args(argv)
    IDSHORT_IN_LISTEN = not args.ohne_idshort_in_listen
    VALUE_IDS = not args.ohne_value_ids
    EINHEITEN_IDS = not args.ohne_einheiten_ids
    LEERE_ENTFERNEN = not args.behalte_leere
    warnungen.clear()

    excel = Path(args.excel).resolve()
    if not excel.is_file():
        raise SystemExit(f"{args.excel}: Datei nicht gefunden")
    basis = excel.parent
    template = Path(args.template) if args.template else next(
        (p for p in (basis / TEMPLATE, TEMPLATE_STANDARD, Path.cwd() / TEMPLATE) if p.is_file()), None)
    if template is None or not template.is_file():
        raise SystemExit(f"AASX-Template nicht gefunden ({args.template or TEMPLATE}) - --template angeben")

    # 1. Excel lesen und pruefen, Vokabular aus dem Template
    baum0, _ = lade_template(template)
    tpl_vok, tpl_ids = template_vokabular(baum0)
    vokabular_abgleichen(tpl_vok)
    modell, fehler, wn = epd_excel.pruefe(epd_excel.lese_arbeitsmappe(excel), tpl_vok, basis)
    print(f"{excel.name}: Vorlage v{modell['epd'].get('vorlagen_version')}, {len(modell['produkte'])} Produkte, "
          f"{len(modell['ergebnisse'])} Werte, {len(modell['faktoren'])} Faktoren, Template {template.name}")
    epd_excel.ausgeben(fehler, wn + warnungen)
    if fehler:
        print(f"\nExcel nicht in Ordnung - keine AASX erzeugt.")
        return 2
    if args.nur_pruefen:
        print("\nExcel in Ordnung.")
        return 0
    warnungen.clear()

    produkte = [p for p in modell["produkte"] if p.get("erzeugen") != "nein"]
    if args.variante:
        unbekannt = set(args.variante) - {p["artikel"] for p in modell["produkte"]}
        if unbekannt:
            raise SystemExit(f"--variante {', '.join(sorted(unbekannt))}: nicht im Blatt Produkte")
        produkte = [p for p in modell["produkte"] if p["artikel"] in args.variante]

    e = modell["epd"]
    einheiten_je_code = {c: i["einheit"] for c, i in modell["indikatoren"].items()}
    pdf = anhang(basis, e["epd_dokument"]) if e["epd_dokument"] else None
    ziel_ordner = Path(args.out) if args.out else basis / "aasx"
    ziel_ordner.mkdir(parents=True, exist_ok=True)
    fehler_gesamt = 0
    berichte = []
    print()
    for produkt in produkte:
        fehlende_pfade.clear()
        baum, sm = lade_template(template)
        dateien, verweise = {}, {}
        if pdf:
            dateien[pdf[0]], verweise["pdf"] = pdf[3], (pdf[1], pdf[2])
        if produkt.get("bild"):
            bild = anhang(basis, produkt["bild"])
            if bild:
                dateien[bild[0]], verweise["bild"] = bild[3], (bild[1], bild[2])
        n = befuelle(sm, modell, produkt, verweise, args.mit_total)
        offene = offene_pflichtfelder(sm)
        if LEERE_ENTFERNEN:
            aufraeumen(baum)
        n_vid = value_ids_setzen(sm, tpl_ids) if VALUE_IDS else 0
        ids = identifier(produkt["artikel"], e["id_basis"], e["submodel_kennung"])
        ids_setzen(baum, ids)
        bericht = konformitaet(baum, einheiten_je_code)
        xml, befunde = normalisieren(baum)
        ziel = ziel_ordner / f"{e['dateiname_praefix']}-{produkt['artikel'].lower()}-epd-submodel.aasx"
        schreibe_aasx(template, xml, ziel, dateien)

        quelle = ("deklariert" if produkt["herkunft"] == "Deklariert"
                  else f"Faktoren x {modell['referenz']}")
        print(f"{ziel.name:44s} {n:3d} Indikatoren  ({quelle})  {ids['submodel']}")
        zeilen = [f"Anhänge: {', '.join('/' + d for d in dateien) or '-'}", f"valueIds gesetzt: {n_vid}"]
        if fehlende_pfade:
            zeilen.append(f"Im Template nicht gefunden, nicht befüllt: {', '.join(fehlende_pfade)}")
        if offene:
            zeilen.append(f"Pflichtfelder (laut Template-Kardinalität) ohne Wert, entfernt: {', '.join(offene)}")
        zeilen += [f"Template korrigiert: {b}" for b in bericht]
        if befunde is None:
            zeilen.append("aas-core3 nicht installiert (pip install aas-core3.0): KEINE Prüfung")
            fehler_gesamt += 1
        else:
            erwartet = [b for b in befunde if "AASd-120" in b]
            rest = [b for b in befunde if "AASd-120" not in b]
            zeilen.append(f"aas-core3: {len(erwartet)} x AASd-120 (WAGO-Konvention), {len(rest)} weitere Befunde")
            zeilen += [f"FEHLER {b}" for b in rest]
            fehler_gesamt += len(rest)
        berichte.append(zeilen)
    # Zeilen, die fuer alle Produkte gleich sind, nur einmal ausgeben
    gemeinsam = [z for z in berichte[0] if all(z in b for b in berichte)] if berichte else []
    for i, (produkt, zeilen) in enumerate(zip(produkte, berichte)):
        eigen = [z for z in zeilen if z not in gemeinsam]
        if eigen:
            print(f"\n{produkt['artikel']}:")
            for z in eigen:
                print(f"   {z}")
    if gemeinsam:
        print(f"\nFür {'alle ' + str(len(produkte)) + ' Produkte' if len(produkte) > 1 else 'das Produkt'}:")
        for z in gemeinsam:
            print(f"   {z}")
    if warnungen:
        print("\nWarnungen:")
        for w in dict.fromkeys(warnungen):
            print(f"  - {w}")
    if fehler_gesamt:
        print(f"\n{fehler_gesamt} Befunde - Ergebnis NICHT konform bzw. nicht geprüft.")
        return 1
    print(f"\n{len(produkte)} AASX in {ziel_ordner} - alle konform (aas-core3).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
