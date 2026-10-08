
import { useEffect, useState } from 'react'
import './App.css'
import wagoLogo from './assets/wago-logo.svg'
import { AAS_URL } from './aas/client'
import {
  ladeProdukte,
  phasen,
  phasenNamen,
  type Produkt
} from './aas/epd'
import { produktbild, shopUrl, verpackung } from './shop'

// =====================================================
// WAGO EPD EXPLORER
// Datenquelle: eigener AAS-Server (../aas-server),
// befuellt mit den AASX-Paketen aus den Skripten
// =====================================================

const format = (wert: number, stellen = 5) =>
    wert.toLocaleString('de-DE', {
      maximumFractionDigits: stellen
    })

const formatDatum = (iso: string) =>
    iso
        ? new Date(iso).toLocaleDateString('de-DE', {
          month: 'long',
          year: 'numeric'
        })
        : '–'

const verifizierungNamen: Record<string, string> = {
  External: 'Extern',
  ThirdPartyVerification: 'Extern (Drittprüfung)',
  Internal: 'Intern'
}

function standardMenge(nummer: string): string {
  return String(verpackung(nummer)?.vpe ?? 100)
}

function berechneGWP(
    produkt: Produkt,
    menge: number
): number {
  return produkt.gwpProStueck * menge
}

function App() {
  const [produkte, setProdukte] = useState<Produkt[] | null>(null)
  const [fehler, setFehler] = useState<string | null>(null)

  useEffect(() => {
    ladeProdukte()
        .then(setProdukte)
        .catch((e: Error) => setFehler(e.message))
  }, [])

  if (fehler) {
    return (
        <div className="app">
          <main>
            <div className="empty-state">
              <h3>AAS-Server nicht erreichbar</h3>
              <p className="error">{fehler}</p>
              <p className="hint">
                Läuft der Server? Im Ordner aas-server:
                {' '}<code>docker compose up -d</code>
              </p>
            </div>
          </main>
        </div>
    )
  }

  if (produkte === null) {
    return (
        <div className="app">
          <main>
            <div className="empty-state">
              EPD-Daten werden vom AAS-Server geladen …
            </div>
          </main>
        </div>
    )
  }

  if (produkte.length === 0) {
    return (
        <div className="app">
          <main>
            <div className="empty-state">
              Der AAS-Server enthält keine EPD-Teilmodelle.
            </div>
          </main>
        </div>
    )
  }

  return <Explorer produkte={produkte} />
}

function Explorer({ produkte }: { produkte: Produkt[] }) {
  const start =
      produkte.find(p => p.nummer === '221-422') ?? produkte[0]

  const [suche, setSuche] = useState('')
  const [serienFilter, setSerienFilter] = useState('Alle')
  const [filter, setFilter] = useState('Alle')
  const [auswahl, setAuswahl] = useState(start.nummer)
  const [menge, setMenge] = useState(standardMenge(start.nummer))
  const [ansicht, setAnsicht] =
      useState<'grafik' | 'tabelle'>('grafik')

  const [vergleichA, setVergleichA] = useState(start.nummer)
  const [vergleichB, setVergleichB] = useState(
      (produkte.find(p => p.nummer !== start.nummer) ?? start).nummer
  )
  const [vergleichMenge, setVergleichMenge] = useState('1000')

  const serien = [...new Set(produkte.map(p => p.serie))]

  function findeProdukt(nummer: string): Produkt {
    return produkte.find(p => p.nummer === nummer) ?? produkte[0]
  }

  const produkt = findeProdukt(auswahl)
  const pack = verpackung(produkt.nummer)

  // Bild aus der AAS bevorzugen, sonst das Shop-Foto
  const dokumentBild =
      produkt.dokument.bildUrl ?? produktbild(produkt.nummer)

  const schnellwahl = pack
      ? [
        { menge: pack.uvpe, label: 'UVPE' },
        { menge: pack.vpe, label: '1 VPE' },
        { menge: pack.vpe * 5, label: '5 VPE' },
        { menge: pack.vpe * 10, label: '10 VPE' }
      ]
      : [1, 10, 100, 1000].map(n => ({ menge: n, label: 'Stück' }))

  const anzahl = Number(menge)
  const gueltigeMenge =
      menge.trim() !== '' &&
      Number.isSafeInteger(anzahl) &&
      anzahl > 0

  const gesamtergebnis = gueltigeMenge
      ? berechneGWP(produkt, anzahl)
      : 0

  const wertProStueck =
      berechneGWP(produkt, 1)

  const sichtbareProdukte = produkte.filter(p => {
    const begriff = suche.trim().toLowerCase()
    const passtSuche =
        p.nummer.includes(begriff) ||
        p.typ.toLowerCase().includes(begriff) ||
        p.name.toLowerCase().includes(begriff)

    const passtSerie =
        serienFilter === 'Alle' || p.serie === serienFilter

    const passtFilter =
        filter === 'Alle' || p.typ === filter

    return passtSuche && passtSerie && passtFilter
  })

  const phasenErgebnisse = phasen.map(phase => {
    const proStueck = produkt.phasen[phase]

    return {
      phase,
      name: phasenNamen[phase],
      proStueck,
      wert: gueltigeMenge && proStueck !== null
          ? proStueck * anzahl
          : 0
    }
  })

  const maxPhase = Math.max(
      ...phasenErgebnisse.map(p => p.wert),
      0.000001
  )

  // Vergleich

  const produktA = findeProdukt(vergleichA)
  const produktB = findeProdukt(vergleichB)

  const vMenge = Number(vergleichMenge)
  const gueltigerVergleich =
      vergleichMenge.trim() !== '' &&
      Number.isSafeInteger(vMenge) &&
      vMenge > 0

  const wertA = gueltigerVergleich
      ? berechneGWP(produktA, vMenge)
      : 0

  const wertB = gueltigerVergleich
      ? berechneGWP(produktB, vMenge)
      : 0

  const maxVergleich = Math.max(
      wertA,
      wertB,
      0.000001
  )

  const klemmenText = (p: Produkt) =>
      p.klemmen === null
          ? ''
          : ` · ${p.klemmen} Klemmstellen`

  const produktOptionen = produkte.map(p => (
      <option
          key={p.nummer}
          value={p.nummer}
      >
        WAGO {p.nummer} – {p.typ}
      </option>
  ))

  // Bei einem Produktwechsel startet die Stückzahl mit
  // einer vollen Verpackungseinheit, wie im Shop
  function wechsleProdukt(nummer: string) {
    setAuswahl(nummer)
    setMenge(standardMenge(nummer))
  }

  function waehleProdukt(nummer: string) {
    wechsleProdukt(nummer)

    document.getElementById('rechner')?.scrollIntoView({
      behavior: 'smooth'
    })
  }

  return (
      <div className="app">

        {/* NAVIGATION */}

        <header className="header">
          <a className="logo" href="#">
            <img src={wagoLogo} alt="WAGO" />
            <span>EPD Explorer</span>
          </a>

          <nav>
            <a href="#produkte">Produkte</a>
            <a href="#rechner">EPD-Rechner</a>
            <a href="#ergebnisse">Ergebnisse</a>
            <a href="#vergleich">Vergleich</a>
            <a href="#dokumente">Dokumentation</a>
          </nav>

          <span className="mock-badge" title={AAS_URL}>
          LIVE · AAS
        </span>
        </header>

        <main>

          {/* STARTSEITE */}

          <section className="hero">
            <div className="hero-content">
            <span className="eyebrow">
              WAGO · SERIE {serien.join(' / ')}
            </span>

              <h1>
                Umweltkennzahlen
                <br />
                einfach berechnen.
              </h1>

              <p>
                Wähle ein WAGO-Produkt,
                gib deine gewünschte Stückzahl ein
                und berechne die Umweltwirkungen
                automatisch.
              </p>

              <a
                  href="#produkte"
                  className="primary-button"
              >
                Produkte entdecken →
              </a>
            </div>

            <div className="hero-stat">
              <span>Ausgewählt: WAGO {produkt.nummer}</span>
              <strong>{format(wertProStueck, 4)}</strong>
              <span>{produkt.einheit} pro Stück</span>

              <div className="hero-stat-line" />

              <small>
                Datengrundlage: {produkt.dokument.kennung}
              </small>
            </div>
          </section>

          {/* KENNZAHLEN */}

          <div className="stats">
            <div className="stat-card">
              <strong>{produkte.length}</strong>
              <span>Produktvarianten</span>
            </div>

            <div className="stat-card">
              <strong>{phasen.length}</strong>
              <span>Lebenszyklusgruppen</span>
            </div>

            <div className="stat-card">
              <strong>{serien.length}</strong>
              <span>WAGO-Produktserien</span>
            </div>
          </div>

          {/* PRODUKTAUSWAHL */}

          <section id="produkte" className="section">
            <div className="section-heading">
              <div>
              <span className="eyebrow">
                PRODUKTÜBERSICHT
              </span>

                <h2>WAGO Verbindungsklemmen</h2>

                <p>
                  Wähle deine Produktvariante.
                  Die EPD-Daten kommen direkt
                  aus der Verwaltungsschale (AAS).
                </p>
              </div>

              <span className="counter">
              {sichtbareProdukte.length} Produkte
            </span>
            </div>

            <div className="toolbar">
              <input
                  aria-label="Produkt suchen"
                  placeholder="Artikelnummer suchen ..."
                  value={suche}
                  onChange={e => setSuche(e.target.value)}
              />

              <select
                  aria-label="Produktserie filtern"
                  value={serienFilter}
                  onChange={e => setSerienFilter(e.target.value)}
              >
                <option value="Alle">Alle Serien</option>
                {serien.map(s => (
                    <option key={s} value={s}>
                      Serie {s}
                    </option>
                ))}
              </select>

              <select
                  aria-label="Produktgruppe filtern"
                  value={filter}
                  onChange={e => setFilter(e.target.value)}
              >
                <option>Alle</option>
                <option>Standard</option>
                <option>Green Range</option>
              </select>
            </div>

            <div className="produkt-grid">
              {sichtbareProdukte.map(p => (
                  <button
                      key={p.nummer}
                      className={
                        auswahl === p.nummer
                            ? 'produkt aktiv'
                            : 'produkt'
                      }
                      onClick={() => waehleProdukt(p.nummer)}
                  >
                    <div className="produkt-top">
                  <span
                      className={
                        p.typ === 'Green Range'
                            ? 'tag green'
                            : 'tag'
                      }
                  >
                    {p.typ}
                  </span>

                      {auswahl === p.nummer && (
                          <span className="selected-icon">
                      ✓
                    </span>
                      )}
                    </div>

                    <div className="produkt-bild">
                      {produktbild(p.nummer) ? (
                          <img
                              src={produktbild(p.nummer)!}
                              alt={`WAGO ${p.nummer}`}
                              loading="lazy"
                          />
                      ) : (
                          <div className="connector">
                            <strong>{p.klemmen ?? '–'}</strong>
                            <span>Klemmstellen</span>
                          </div>
                      )}
                    </div>

                    <h3>WAGO {p.nummer}</h3>

                    <p>
                      Serie {p.serie} · {p.typ}
                      {klemmenText(p)}
                    </p>

                    <span className="produkt-link">
                  {auswahl === p.nummer
                      ? 'Ausgewählt ✓'
                      : 'Produkt auswählen →'}
                </span>
                  </button>
              ))}
            </div>

            {sichtbareProdukte.length === 0 && (
                <div className="empty-state">
                  Keine passenden Produkte gefunden.
                </div>
            )}
          </section>

          {/* EPD-RECHNER */}

          <section id="rechner" className="section">
            <div className="section-heading">
              <div>
              <span className="eyebrow">
                BERECHNUNG
              </span>

                <h2>EPD-Rechner</h2>

                <p>
                  Berechne das Treibhauspotenzial
                  automatisch für deine gewünschte
                  Produktvariante und Stückzahl.
                </p>
              </div>
            </div>

            <div className="rechner-layout">

              <div className="panel">
                <h3>Berechnungsparameter</h3>

                <label htmlFor="variante">
                  Produktvariante
                </label>

                <select
                    id="variante"
                    value={auswahl}
                    onChange={e => wechsleProdukt(e.target.value)}
                >
                  {produktOptionen}
                </select>

                <label htmlFor="menge">
                  Anzahl der Produkte
                </label>

                <input
                    id="menge"
                    type="number"
                    min="1"
                    step={pack ? pack.uvpe : 1}
                    value={menge}
                    onChange={e => setMenge(e.target.value)}
                />

                <div className="quick-buttons">
                  {schnellwahl.map(s => (
                      <button
                          key={s.label}
                          className={
                            menge === String(s.menge)
                                ? 'active'
                                : ''
                          }
                          onClick={() => setMenge(String(s.menge))}
                      >
                        {format(s.menge, 0)}
                        <small>{s.label}</small>
                      </button>
                  ))}
                </div>

                {!gueltigeMenge && (
                    <p className="error">
                      Bitte eine gültige Stückzahl
                      größer 0 eingeben.
                    </p>
                )}

                {pack && (
                    <div className="shop-box">
                      <p>
                        Im WAGO-Shop: VPE {format(pack.vpe, 0)} Stück,
                        {' '}kleinste Packung (UVPE) {format(pack.uvpe, 0)} Stück.
                      </p>

                      {gueltigeMenge && anzahl % pack.uvpe !== 0 && (
                          <p className="shop-hinweis">
                            {format(anzahl, 0)} Stück sind kein Vielfaches
                            {' '}der UVPE. Im Shop wären es
                            {' '}{format(Math.ceil(anzahl / pack.uvpe) * pack.uvpe, 0)} Stück.
                          </p>
                      )}

                      <a
                          className="shop-link"
                          href={shopUrl(produkt.nummer)}
                          target="_blank"
                          rel="noopener noreferrer"
                      >
                        WAGO {produkt.nummer} im Shop ansehen ↗
                      </a>
                    </div>
                )}

                <div className="info-box">
                  <strong>
                    Berechnungsgrundlage
                  </strong>

                  <p>
                    EPD: {produkt.dokument.kennung}
                  </p>

                  <p>
                    Teilmodell: {produkt.submodelId}
                  </p>

                  <p>
                    Indikator: GWP-total, Module
                    {' '}A1–C4 (ohne Modul D)
                  </p>
                </div>
              </div>

              <div className="panel ergebnis-panel">
              <span className="eyebrow">
                DEIN ERGEBNIS
              </span>

                <h3>WAGO {produkt.nummer}</h3>

                <p>
                  Serie {produkt.serie} · {produkt.typ}
                  {klemmenText(produkt)}
                </p>

                <div className="main-result">
                <span>
                  Gesamtes Treibhauspotenzial
                </span>

                  <strong>
                    {gueltigeMenge
                        ? format(gesamtergebnis)
                        : '–'}
                  </strong>

                  <span>
                  {produkt.einheit}
                </span>
                </div>

                <div className="result-row">
                  <span>Wert pro Stück</span>

                  <strong>
                    {format(wertProStueck, 6)}
                    {' '}{produkt.einheit}
                  </strong>
                </div>

                <div className="result-row">
                  <span>Stückzahl</span>

                  <strong>
                    {gueltigeMenge
                        ? format(anzahl, 0)
                        : '–'}
                  </strong>
                </div>

                <div className="formula">
                  {format(wertProStueck, 6)}
                  {' × '}
                  {gueltigeMenge
                      ? format(anzahl, 0)
                      : '–'}
                </div>
              </div>
            </div>
          </section>

          {/* UMWELTKENNZAHLEN */}

          <section id="ergebnisse" className="section">
            <div className="section-heading">
              <div>
              <span className="eyebrow">
                DETAILERGEBNISSE
              </span>

                <h2>Umweltkennzahlen</h2>

                <p>
                  Treibhauspotenzial nach
                  Lebenszyklusphase für WAGO {auswahl}.
                </p>
              </div>
            </div>

            <div className="panel">
              <div className="ansicht-kopf">
                <h3>
                  {ansicht === 'grafik'
                      ? 'Lebenszyklusphasen'
                      : 'Detaillierte Ergebnistabelle'}
                </h3>

                <div
                    className="ansicht-umschalter"
                    role="group"
                    aria-label="Darstellung wählen"
                >
                  <button
                      className={ansicht === 'grafik' ? 'active' : ''}
                      aria-pressed={ansicht === 'grafik'}
                      onClick={() => setAnsicht('grafik')}
                  >
                    Grafik
                  </button>

                  <button
                      className={ansicht === 'tabelle' ? 'active' : ''}
                      aria-pressed={ansicht === 'tabelle'}
                      onClick={() => setAnsicht('tabelle')}
                  >
                    Tabelle
                  </button>
                </div>
              </div>

              {ansicht === 'grafik' && (
              <div className="chart">
                {phasenErgebnisse.map(p => (
                    <div
                        className="chart-row"
                        key={p.phase}
                    >
                  <span>
                    {p.phase}
                    <small>{p.name}</small>
                  </span>

                      <div className="bar-track">
                        <div
                            className="bar"
                            style={{
                              width:
                                  (p.wert / maxPhase * 100)
                                  + '%'
                            }}
                        />
                      </div>

                      <strong>
                        {gueltigeMenge && p.proStueck !== null
                            ? format(p.wert, 6)
                            : '–'}
                      </strong>
                    </div>
                ))}
              </div>
              )}

              {ansicht === 'tabelle' && (
              <div className="table-scroll">
                <table>
                  <thead>
                  <tr>
                    <th>Phase</th>
                    <th>Bezeichnung</th>
                    <th>pro Stück</th>
                    <th>{produkt.einheit}</th>
                  </tr>
                  </thead>

                  <tbody>
                  {phasenErgebnisse.map(p => (
                      <tr key={p.phase}>
                        <td>{p.phase}</td>
                        <td>{p.name}</td>
                        <td>
                          {p.proStueck === null
                              ? 'nicht deklariert'
                              : format(p.proStueck, 6)}
                        </td>
                        <td>
                          {gueltigeMenge && p.proStueck !== null
                              ? format(p.wert, 6)
                              : '–'}
                        </td>
                      </tr>
                  ))}
                  </tbody>
                </table>
              </div>
              )}

              <p className="hint">
                Werte aus dem EPD-Teilmodell der
                Verwaltungsschale. Phasen, die in der EPD
                feiner aufgeteilt sind (z. B. B1–B7),
                werden summiert.
                Modul D ist nicht im Gesamtwert enthalten
                {produkt.modulD !== null && (
                    <> (Modul D: {format(produkt.modulD, 6)}
                      {' '}{produkt.einheit} pro Stück)</>
                )}.
              </p>
            </div>
          </section>

          {/* PRODUKTVERGLEICH */}

          <section id="vergleich" className="section">
            <div className="section-heading">
              <div>
              <span className="eyebrow">
                PRODUKTVERGLEICH
              </span>

                <h2>Varianten vergleichen</h2>

                <p>
                  Vergleiche die Klimawirkung
                  zweier Produkte für dieselbe
                  Stückzahl.
                </p>
              </div>
            </div>

            <div className="panel">

              <div className="vergleich-eingaben">
                <div>
                  <label htmlFor="vergleich-a">
                    Produkt A
                  </label>

                  <select
                      id="vergleich-a"
                      value={vergleichA}
                      onChange={e =>
                          setVergleichA(e.target.value)
                      }
                  >
                    {produktOptionen}
                  </select>
                </div>

                <div>
                  <label htmlFor="vergleich-b">
                    Produkt B
                  </label>

                  <select
                      id="vergleich-b"
                      value={vergleichB}
                      onChange={e =>
                          setVergleichB(e.target.value)
                      }
                  >
                    {produktOptionen}
                  </select>
                </div>

                <div>
                  <label htmlFor="vergleich-menge">
                    Stückzahl
                  </label>

                  <input
                      id="vergleich-menge"
                      type="number"
                      min="1"
                      step="1"
                      value={vergleichMenge}
                      onChange={e =>
                          setVergleichMenge(e.target.value)
                      }
                  />
                </div>
              </div>

              {!gueltigerVergleich && (
                  <p className="error">
                    Bitte eine gültige Stückzahl eingeben.
                  </p>
              )}

              <div className="vergleich-karten">

                <div className="vergleich-karte">
                <span className="eyebrow">
                  PRODUKT A
                </span>

                  <h3>
                    WAGO {produktA.nummer}
                  </h3>

                  <p>
                    {produktA.typ}
                    {klemmenText(produktA)}
                  </p>

                  <div className="vergleich-wert">
                    {gueltigerVergleich
                        ? format(wertA)
                        : '–'}
                  </div>

                  <span>{produktA.einheit}</span>

                  <div className="bar-track">
                    <div
                        className="bar"
                        style={{
                          width: gueltigerVergleich
                              ? (wertA / maxVergleich * 100)
                              + '%'
                              : '0%'
                        }}
                    />
                  </div>
                </div>

                <div className="vergleich-karte">
                <span className="eyebrow">
                  PRODUKT B
                </span>

                  <h3>
                    WAGO {produktB.nummer}
                  </h3>

                  <p>
                    {produktB.typ}
                    {klemmenText(produktB)}
                  </p>

                  <div className="vergleich-wert">
                    {gueltigerVergleich
                        ? format(wertB)
                        : '–'}
                  </div>

                  <span>{produktB.einheit}</span>

                  <div className="bar-track">
                    <div
                        className="bar"
                        style={{
                          width: gueltigerVergleich
                              ? (wertB / maxVergleich * 100)
                              + '%'
                              : '0%'
                        }}
                    />
                  </div>
                </div>
              </div>

              {gueltigerVergleich && (
                  <div className="vergleich-hinweis">
                <span>
                  Absolute Differenz
                </span>

                    <strong>
                      {format(Math.abs(wertA - wertB))}
                      {' '}{produktA.einheit}
                    </strong>
                  </div>
              )}

              <p className="hint">
                Dieser Vergleich zeigt ausschließlich
                die rechnerischen GWP-Werte bei gleichen
                EPD-Nutzungsannahmen. Unterschiedliche
                Funktionen oder Einsatzbedingungen
                werden nicht bewertet.
              </p>
            </div>
          </section>

          {/* DOKUMENTATION */}

          <section id="dokumente" className="section">
            <div className="section-heading">
              <div>
              <span className="eyebrow">
                DOKUMENTATION
              </span>

                <h2>
                  Environmental Product Declaration
                </h2>

                <p>
                  Informationen zur Datengrundlage
                  von WAGO {produkt.nummer}.
                </p>
              </div>
            </div>

            <div className="panel">
            <span className="tag green">
              {produkt.dokument.programmbetreiber || 'EPD'}
            </span>

              <h3 className="dokument-titel">
                {produkt.name || `WAGO – Serie ${produkt.serie}`}
              </h3>

              {dokumentBild && (
                  <img
                      className="produkt-foto"
                      src={dokumentBild}
                      alt={`WAGO ${produkt.nummer}`}
                  />
              )}

              <div className="dokument-grid">
                <div>
                <span>
                  Registrierungsnummer
                </span>

                  <strong>
                    {produkt.dokument.kennung || '–'}
                  </strong>
                </div>

                <div>
                <span>
                  Ausgabedatum
                </span>

                  <strong>
                    {formatDatum(produkt.dokument.ausgabe)}
                  </strong>
                </div>

                <div>
                <span>
                  Gültig bis
                </span>

                  <strong>
                    {formatDatum(produkt.dokument.gueltigBis)}
                  </strong>
                </div>

                <div>
                <span>
                  Referenznutzungsdauer
                </span>

                  <strong>
                    {produkt.dokument.nutzungsdauer
                        ? `${produkt.dokument.nutzungsdauer} Jahre`
                        : '–'}
                  </strong>
                </div>

                <div>
                <span>
                  Verifizierung
                </span>

                  <strong>
                    {verifizierungNamen[produkt.dokument.verifizierung]
                        ?? (produkt.dokument.verifizierung || '–')}
                  </strong>
                </div>

                <div>
                <span>
                  Aktuelle Datenquelle
                </span>

                  <strong>
                    AAS-Server ({AAS_URL})
                  </strong>
                </div>
              </div>

              <a
                  className="dokument-link"
                  href={produkt.dokument.epdUrl}
                  target="_blank"
                  rel="noopener noreferrer"
              >
                EPD als PDF öffnen ↗
              </a>

              <p className="hint">
                Das PDF ist als Datei im EPD-Teilmodell
                der Verwaltungsschale hinterlegt.
                Diese Anwendung stellt keine
                eigenständige EPD-Verifizierung dar.
              </p>
            </div>
          </section>

          {/* FOOTER */}

          <footer>
            <strong>
              WAGO EPD Explorer
            </strong>

            <p>
              Hochschulprojekt · Daten aus der
              Verwaltungsschale (AAS) · {produkte.length} Produkte
            </p>
          </footer>

        </main>
      </div>
  )
}

export default App
