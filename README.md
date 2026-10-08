# WAGO EPD → AAS

Hochschulprojekt: EPD-Daten (PEP ecopassport) von WAGO-Verbindungsklemmen als
Verwaltungsschale (AAS) bereitstellen und in einer Oberfläche nutzbar machen.

| Ordner / Datei | Inhalt |
|---|---|
| `pep_to_excel_*.py`, `excel_to_aasx_*.py` | Skripte: PEP-PDF → Excel → AASX (Teilmodell EPD Type III) |
| `aasx/`, `aasx_2273/` | erzeugte AASX-Pakete je Produktvariante |
| `aas-server/` | eigener AAS-Server (Eclipse BaSyx, Docker), lädt die Pakete |
| `frontend/` | WAGO EPD Explorer, liest die Daten über die AAS-API |
| `weitereInformationenWago/` | Leitfaden zum Teilmodell, API-Beschreibung WAGO SmartData |

Schnellstart:

```bash
cd aas-server && docker compose up -d
cd ../frontend && npm install && npm run dev
```
