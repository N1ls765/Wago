#!/usr/bin/env python3
"""PEP WAGO-00002 (2273-Serie, Steckklemmen) -> Standard-Excel EPD_2273.xlsx

Aufruf (im Projektordner):  python epd/2273/pep_2273.py
Danach:                     python skripte/excel_to_aasx.py epd/2273/EPD_2273.xlsx

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
    "pdf": "WAGO-00002-V01.01-EN.pdf",
    "ziel": "EPD_2273.xlsx",
    "seiten": {},                            # Tabellen werden ueber ihren Inhalt gefunden
    "kopfdaten": {},
    "klemmstellen": {},                      # PDF S. 5: letzte Ziffer = Anzahl Klemmstellen
    "produkttext": "Push-wire splicing connector for solid copper conductors 0.5-2.5 mm², rated 450 V, "
                   "up to 24 A.",                                                         # PDF S. 5
    "standort": "WAGO production facility in Germany",                                    # PDF S. 8
    "herstellung": "Production and packaging at one WAGO manufacturing facility in Germany (injection "
                   "moulding, metal forming and assembly). Electricity: average mix of WAGO's production "
                   "facilities in Germany, calendar year {lca_referenzjahr}.",            # PDF S. 8
    # PDF S. 6 und S. 14: LCA nach ISO 14040/14044 und EN 15804 (+A2, EF 3.1)
    "normen_zusatz": ["ISO 14040", "ISO 14044", "EN 15804+A2"],
    "bilder": {},
}

if __name__ == "__main__":
    sys.exit(0 if pep_bausteine.pep_zu_vorlage(KONFIG, HIER) else 2)
