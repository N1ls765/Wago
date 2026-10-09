# WAGO EPD Explorer (Oberfläche)

React + TypeScript + Vite. Die EPD-Daten kommen live vom AAS-Server
(`../aas-server`), nicht mehr aus fest eingetragenen Werten.

## Starten

1. AAS-Server starten (siehe `../aas-server/README.md`)
2. Oberfläche starten:

   ```bash
   cd frontend
   npm install
   npm run dev
   ```

3. http://localhost:5173 öffnen

Ohne lokales Node.js: Im Hauptordner startet `docker compose up -d --build` alles
zusammen, die Oberfläche läuft dann fertig gebaut unter http://localhost:8080.
Dafür sorgen `Dockerfile` (bauen mit Node, ausliefern mit nginx) und `nginx.conf`
(leitet `/aas/...` an den AAS-Server weiter, wie der Vite-Proxy im Dev-Betrieb).
Nach Änderungen an der Oberfläche: `docker compose up -d --build frontend`.

## Wie die Daten geladen werden

Im Dev-Betrieb leitet Vite alle Anfragen an `/aas/...` an den AAS-Server weiter
(`vite.config.ts`). Ablauf in `src/aas/epd.ts`:

1. `GET /shells`: alle Verwaltungsschalen
2. je Teilmodell-Referenz `GET /submodels/{id}/$metadata`: ist die `semanticId`
   `https://admin-shell.io/idta/EPD/EPDTypeIII/1/0`?
3. für das EPD-Teilmodell `GET /submodels/{id}/$value`: Werte ohne Metadaten
4. GWP-total je Lebenszyklusphase summieren (B1–B7 → „B1-B7“, C1 + C2-C4 →
   „C1-C4“); Modul D wird separat ausgewiesen

PDF und Produktbild werden direkt über `.../submodel-elements/<pfad>/attachment` verlinkt.

Verpackungseinheiten (VPE/UVPE), Shop-Links und Produktfotos stehen nicht in der EPD
und deshalb nicht in der AAS. Sie kommen aus `src/shop.ts` und
`public/produktbilder/<Artikel>.jpg` (Hauptbild der Shop-Produktseite, 480 px,
Stand WAGO-Shop 08.10.2026) und müssen bei neuen Artikeln dort ergänzt werden.

Andere Server-Adresse: `.env.example` nach `.env.local` kopieren und anpassen.

## Bekannte Lücken in den Daten

- **Klemmstellen** und **Green Range** gibt es nicht als eigene Felder. Beides wird
  aus `productDescription` gelesen („2 clamping units“, „green range“).
- **Produktbild** liegt in der AAS bisher nur bei 221-422; sonst wird das Shop-Foto gezeigt.

## Prüfen

```bash
npm run build   # Typprüfung + Build
npm run lint
```
