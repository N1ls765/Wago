# Eigener AAS-Server (Eclipse BaSyx)

Stellt die AASX-Pakete aus `../aasx` (Serie 221) und `../aasx_2273` (Serie 2273)
über die standardisierte AAS-API bereit (IDTA Part 2, API v3.0) und bietet mit der
BaSyx AAS Web UI ein Portal zum Durchklicken.

## Starten

Voraussetzung: Docker Desktop läuft.

```bash
cd aas-server
docker compose up -d
```

| Was | Adresse |
|---|---|
| AAS-API | http://localhost:8081/shells, http://localhost:8081/submodels |
| API-Doku (Swagger) | http://localhost:8081/swagger-ui/index.html |
| AAS Web UI (Portal) | http://localhost:3000 |

Der erste Start dauert etwa 30 Sekunden. Stoppen mit `docker compose down`.

## Neue oder geänderte AASX-Pakete

Der Server hält die Daten nur im Speicher (`basyx.backend=InMemory`) und liest
beim Start alle Pakete aus den eingebundenen Ordnern neu ein. Nach einem
Skript-Lauf reicht also:

```bash
docker compose restart aas-env
```

Ein neuer Ordner (z. B. eine weitere Produktserie) muss in `docker-compose.yml`
als Volume und in `aas-env.properties` unter `basyx.environment` ergänzt werden.

## Beim Start erwartete Warnungen

`Colliding Ids detected for ConceptDescription ...` – jedes AASX-Paket bringt
dieselben ConceptDescriptions mit. BaSyx behält die erste und überspringt die
übrigen. Solange sie identisch sind, ist das unkritisch.

## Nützliche Abfragen

Identifier werden in der URL **base64url-kodiert** (ohne `=`):

```bash
# Alle Verwaltungsschalen
curl "http://localhost:8081/shells"

# EPD-Teilmodell von 221-425, nur Werte (kompakt)
ID=$(printf 'https://wago.com/ids/submodel/221-425/EPD' | base64 -w0 | tr '+/' '-_' | tr -d '=')
curl "http://localhost:8081/submodels/$ID/\$value"

# PDF der EPD
curl -o epd.pdf "http://localhost:8081/submodels/$ID/submodel-elements/epdDocument/attachment"
```

Dieselbe API bietet auch WAGOs SmartData-Dienst an (siehe
`../weitereInformationenWago/SmartData AAS Env.json`). Die Oberfläche kann deshalb
später über `VITE_AAS_URL` auf den echten Dienst umgestellt werden.
