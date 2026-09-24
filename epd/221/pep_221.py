#!/usr/bin/env python3
"""PEP WAGO-00001 (221-Serie, Hebelklemmen) -> Standard-Excel EPD_221.xlsx

Aufruf (im Projektordner):  python epd/221/pep_221.py
Danach:                     python skripte/excel_to_aasx.py epd/221/EPD_221.xlsx

Hier steht nur, was diese PEP besonders macht. Tabellen lesen, zuordnen,
pruefen und die Excel schreiben erledigt skripte/pep_bausteine.py.
Neue PEP: Ordner epd/<serie>/ anlegen, PDF hineinlegen, diese Datei dorthin
kopieren, umbenennen und KONFIG anpassen.
"""

import sys
from pathlib import Path

HIER = Path(__file__).resolve().parent                     # epd/<serie>/
sys.path.insert(0, str(HIER.parents[1] / "skripte"))       # gemeinsame Skripte
import pep_bausteine  # noqa: E402

KONFIG = {
    "pdf": "WAGO-00001-V01.01-EN.pdf",
    "ziel": "EPD_221.xlsx",
    # Tabellen werden ueber ihren Inhalt gefunden. Nur falls das scheitert,
    # gedruckte Seitenzahlen vorgeben, z. B. {"faktoren": [14, 15]}.
    "seiten": {},
    # Kopfdaten, die im PDF fehlen oder falsch erkannt werden, hier vorgeben
    # (Schluessel siehe pep_bausteine.KOPFDATEN).
    "kopfdaten": {},
    # PDF S. 4: 221-420/-430 haben zehn, 221-2411/-2421 zwei Klemmstellen,
    # sonst gibt die letzte Ziffer die Anzahl an.
    "klemmstellen": {"221-420": 10, "221-430": 10, "221-2411": 2, "221-2421": 2},
    "produkttext": "Lever-operated splicing connector for solid, stranded and fine-stranded copper conductors "
                   "0.2-6 mm², rated 450 V, up to 32 A.",                                  # PDF S. 4
    "standort": "WAGO production facilities in Germany",                                  # PDF S. 6
    "herstellung": "Production and packaging at several WAGO manufacturing facilities in Germany (injection "
                   "moulding, metal forming and assembly); internal logistics between the facilities modelled "
                   "with primary data. Electricity: average mix of WAGO's production facilities in Germany, "
                   "calendar year {lca_referenzjahr}.",                                   # PDF S. 6
    "normen_zusatz": [],                     # Normen aus dem Deckblatt reichen
    "bilder": {"221-422": "ProductPicture_Wago_221_422.jpg"},   # aus dem AASX-Template
}

if __name__ == "__main__":
    sys.exit(0 if pep_bausteine.pep_zu_vorlage(KONFIG, HIER) else 2)
