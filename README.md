# WAGO EPD → AASX → AAS-Server → EPD Explorer

Erzeugt aus Umweltproduktdeklarationen (EPD, bei WAGO: PEP ecopassport) AASX-Dateien mit dem Teilmodell *EPD Type III*, eine AASX je Artikel. Ein eigener AAS-Server (Eclipse BaSyx) stellt die Pakete über die AAS-API bereit, die Oberfläche *WAGO EPD Explorer* liest sie von dort.

## Starten

Voraussetzung ist nur **Docker Desktop**. Im Hauptordner:

```bash
docker compose up -d --build
```

| Was | Adresse |
|---|---|
| EPD Explorer (Oberfläche) | http://localhost:8080 |
| AAS Web UI (Portal) | http://localhost:3000 |
| AAS-API | http://localhost:8081/shells |

Der erste Start dauert einige Minuten (Images laden, Oberfläche bauen), danach etwa 30 Sekunden, bis der AAS-Server bereit ist. Stoppen mit `docker compose down`.

| Nach … | Befehl |
|---|---|
| neu erzeugten AASX-Paketen | `docker compose restart aas-env` |
| Änderungen an der Oberfläche (`frontend/`) | `docker compose up -d --build frontend` |

Zum Entwickeln an der Oberfläche mit automatischem Neuladen: Node.js 22 installieren, nur den Server starten (`docker compose up -d aas-env aas-web-ui`), dann in `frontend/` `npm install` und `npm run dev` (http://localhost:5173).

Für die Skripte (AASX erzeugen) wird zusätzlich Python 3.12+ gebraucht, siehe „Benutzung“ unten.

```
epd/<serie>/<PEP>.pdf ──pep_<serie>.py──▶ epd/<serie>/EPD_<serie>.xlsx ──excel_to_aasx.py──▶ epd/<serie>/aasx/*.aasx
                        je PEP ein schlankes     Standard-Excel (Vorlage v1)    generisch, eines für alle
                        Skript
EPD ohne passendes PDF-Skript: vorlagen/EPD_Vorlage_v1.xlsx von Hand ausfüllen ──▲
```

## Ordnerstruktur

```
├── README.md
├── requirements.txt
├── skripte/                      gemeinsamer Code
│   ├── excel_to_aasx.py          Standard-Excel → AASX (generisch, ohne Produktwissen)
│   ├── epd_excel.py              Aufbau der Standard-Excel: Felder, Vokabular, Lesen, Prüfen, leere Vorlage
│   ├── pep_bausteine.py          gemeinsamer Teil der PEP-Skripte: PDF lesen, zuordnen, prüfen, Excel schreiben
│   └── pdf_grid.py               Tabellenerkennung im PDF
├── vorlagen/
│   ├── EPD_Vorlage_v1.xlsx       leere Standard-Excel
│   └── epd-type-iii-submodel-template.aasx   AASX-Template von WAGO (Struktur + Vokabular)
├── epd/                          je EPD ein Ordner
│   ├── 221/                      PEP WAGO-00001, Serie 221
│   │   ├── WAGO-00001-V01.01-EN.pdf
│   │   ├── pep_221.py            PDF → EPD_221.xlsx
│   │   ├── EPD_221.xlsx
│   │   ├── produktbilder/
│   │   └── aasx/                 erzeugte AASX (9)
│   └── 2273/                     PEP WAGO-00002, Serie 2273 (gleicher Aufbau, 5 AASX)
├── tests/                        pytest (25 Tests)
├── aas-server/                   eigener AAS-Server (Eclipse BaSyx, Docker), lädt epd/*/aasx
├── frontend/                     WAGO EPD Explorer (React + Vite), liest über die AAS-API
└── wago_unterlagen/              von WAGO: User Guide, Referenzinstanz, SmartData-API-Beschreibungen
```

## Benutzung

Alle Befehle im Projektordner:

```bash
pip install -r requirements.txt

python epd/221/pep_221.py                               # PDF -> epd/221/EPD_221.xlsx, danach Excel sichten
python skripte/excel_to_aasx.py epd/221/EPD_221.xlsx    # -> epd/221/aasx/
python skripte/excel_to_aasx.py epd/2273/EPD_2273.xlsx --variante 2273-202

python skripte/epd_excel.py --pruefe epd/221/EPD_221.xlsx          # nur prüfen
python skripte/epd_excel.py --leer vorlagen/EPD_Vorlage_v1.xlsx    # leere Vorlage neu erzeugen
python -m pytest tests -q
```

Exit-Codes von `excel_to_aasx.py`: `0` ok · `1` aas-core3-Befund bzw. aas-core3 fehlt · `2` Fehler in der Excel (es wird nichts erzeugt).

Optionen: `--out` (Standard: `aasx/` neben der Excel), `--mit-total` (Summenwerte als stageCode `Total`), `--ohne-idshort-in-listen` (strikt AASd-120), `--ohne-value-ids`, `--ohne-einheiten-ids`, `--behalte-leere`, `--template`.

## Neue EPD

**WAGO-PEP (gleiche PDF-Vorlage wie 221/2273):**

1. Ordner `epd/<serie>/` anlegen und das PDF hineinlegen.
2. `epd/2273/pep_2273.py` dorthin kopieren, in `pep_<serie>.py` umbenennen und in `KONFIG` eintragen: PDF-Name, Zieldatei, Klemmstellen-Regel, Produkttext, Standort, Herstellungstext, ggf. Normen-Zusätze und Produktbilder.
3. `python epd/<serie>/pep_<serie>.py` ausführen und die Meldungen lesen. Das Skript bricht ab, statt falsche Werte zu schreiben, wenn
   - eine Tabelle nicht gefunden wird: dann Seitenzahlen unter `seiten` vorgeben,
   - Kopfdaten im Text fehlen: dann unter `kopfdaten` vorgeben,
   - ein Indikatorname unbekannt ist: dann den Katalog in `skripte/epd_excel.py` (`STANDARD_INDIKATOREN`) ergänzen,
   - Total ≠ Summe der Module ist (Spalte verrutscht) oder Faktoren fehlen.
4. Die erzeugte Excel sichten, dann `python skripte/excel_to_aasx.py epd/<serie>/EPD_<serie>.xlsx`.

Neue Layout-Eigenheiten werden in `skripte/pep_bausteine.py` gelöst, damit alle PEPs davon profitieren.

**Andere EPD (anderer Programmbetreiber, kein passendes PDF):** `vorlagen/EPD_Vorlage_v1.xlsx` nach `epd/<name>/` kopieren und von Hand ausfüllen (Anleitung im ersten Blatt), dann `excel_to_aasx.py`.

## Modellierungsentscheidungen

- **Vokabular = Template.** Geschlossene Listen werden zur Laufzeit aus den ConceptDescriptions (valueList) des Templates gelesen: `stageStatus` Assessed/NotAssessed, `valueStatus` Declared/NotDeclared, `scopeType` Class/Instance, `siteType` Fixed/Mobile, `resultCategory`. `verificationType` = `ThirdPartyVerification` (Beispielwert des Templates, dort ohne valueList).
- **scopeType = Class** für Familien-EPDs (Referenzprodukt + Varianten über Faktoren), auch in der AASX des Referenzprodukts: Die EPD dahinter beschreibt die Familie. `Instance` = EPD für genau einen Artikel.
- **Module (PEP):** Spalte „B1-7“ → `B6` (B1–B5, B7 `NotAssessed`), Spalte „C1-4“ → `C2-C4` (C1 laut PEP Kap. 4.5 ausgeschlossen → `NotAssessed`). `C2-C4` steht nicht im Vokabular; das Template erlaubt zusätzliche Codes.
- Kein Wert → `NotDeclared` ohne `value`, nie 0 (Leitfaden Kap. 10).
- Varianten: Wert = Referenzwert × Faktor, 3 signifikante Stellen, Qualifier `EPD/ValueDerivation`.
- Nur quellengestützte Werte: `productMass` nur beim Referenzprodukt (das PEP nennt sie nur dort).
- `valueId` für Vokabularwerte, `unitId` und `baseUnit` je Indikator, wie in der WAGO-Referenzinstanz (abschaltbar).

## Prüfstand (24.09.2026)

- 14/14 AASX aas-core3-konform (nur AASd-120, WAGO-Konvention).
- Alle 14 × 222 Indikatorwerte identisch mit den früheren Skripten (`excel_to_aasx_221.py`, `excel_to_aasx_2273.py`, im Git-Verlauf).
- PDF → Standard-Excel direkt liefert dieselben Werte und Faktoren wie der frühere Zwischenschritt über `pep_to_excel_*.py`. Bewusste Unterschiede: Produktname wie im PDF-Titel (221 ohne vorangestelltes „WAGO“), Normen der 2273 in PDF-Reihenfolge.
- Tabellen werden in beiden PDFs ohne Seitenvorgabe gefunden; in beiden PEPs ist Total = Summe der Module für alle 37 Indikatoren.
- 25 Tests: Regression PDF → Excel, Ankerwerte aus dem PDF, Summenprüfung, Konformität, alle Werte nachgerechnet, 10 Fehlerfälle in der Excel.

Hinweise aus den PDFs (Warnungen der PEP-Skripte): Deckblatt nennt „PCR-4-ed4-…“, Fließtext „PCR-ed4-…“; 221-2421 hat Faktoren, aber keine Masse (kein Teilmodell); im Text stehen „221-2431“ (00001) und „2273-0202“ (00002, Kreisdiagramm S. 5), vermutlich Tippfehler.

## Offene Punkte für WAGO (Template / Leitfaden)

1. **AASd-120 und BaSyx:** Mit idShorts an Listenkindern (WAGO-Konvention) lehnt Eclipse BaSyx (Python SDK 1.2) das Teilmodell im strikten Modus ab. Im toleranten Modus fehlt `EnvironmentalResults` stillschweigend. Mit `--ohne-idshort-in-listen` liest BaSyx alles vollständig ein. Listeneinträge sind per Index adressierbar (`EnvironmentalResults[0]`).
2. `scopeType` hat die semanticId `…/ScopeType`, die valueList steht aber an `…/EPDScope/Type`.
3. `Location/region`, `locality`, `latitude`, `longitude` sind im Template Pflicht (One), laut Leitfaden Kap. 13 optional.
4. `baseUnit` ist im Leitfaden eine Collection (unit, unitId), die CD im Template ist STRING.
5. `verificationType`: Leitfaden „Internal/External“, Template „ThirdPartyVerification“ ohne valueList.
6. 8 Indikatoren (PERM, PERT, PENRM, PENRT, SM, RSF, NRSF, FW) und `C2-C4` fehlen im Template-Vokabular.
