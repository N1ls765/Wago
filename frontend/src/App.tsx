
import { useState } from 'react'
import './App.css'

// =====================================================
// WAGO EPD EXPLORER
// Interaktives Mock-up mit lokalen EPD-Daten
// Spätere Datenquelle: WAGO AAS
// =====================================================

type Phase = 'A1-A3' | 'A4' | 'A5' | 'B1-B7' | 'C1-C4'

type Produkt = {
  nummer: string
  typ: 'Standard' | 'Green Range'
  klemmen: number
  gesamtFaktor: number
  phasenFaktoren: Record<Phase, number>
}

const referenz = {
  nummer: '221-422',
  gwp: 0.0856,
  einheit: 'kg CO₂-Äq.',
  phasen: {
    'A1-A3': 0.0124,
    'A4': 0.00057,
    'A5': 0.00068,
    'B1-B7': 0.0702,
    'C1-C4': 0.00169
  }
}

const phasen: Phase[] = [
  'A1-A3', 'A4', 'A5', 'B1-B7', 'C1-C4'
]

const phasenNamen: Record<Phase, string> = {
  'A1-A3': 'Herstellung',
  'A4': 'Transport',
  'A5': 'Installation',
  'B1-B7': 'Nutzung',
  'C1-C4': 'Entsorgung'
}

// Gerundete Hochrechnungsfaktoren aus der
// bereitgestellten WAGO-221-EPD, Seiten 14–15.
// Die Daten werden später durch AAS-Werte ersetzt.

const produkte: Produkt[] = [
  {
    nummer: '221-412',
    typ: 'Standard',
    klemmen: 2,
    gesamtFaktor: 1.01,
    phasenFaktoren: {
      'A1-A3': 1.07,
      'A4': 0.96,
      'A5': 1.00,
      'B1-B7': 1.00,
      'C1-C4': 0.96
    }
  },
  {
    nummer: '221-413',
    typ: 'Standard',
    klemmen: 3,
    gesamtFaktor: 1.10,
    phasenFaktoren: {
      'A1-A3': 1.58,
      'A4': 1.47,
      'A5': 1.82,
      'B1-B7': 1.00,
      'C1-C4': 1.37
    }
  },
  {
    nummer: '221-415',
    typ: 'Standard',
    klemmen: 5,
    gesamtFaktor: 1.28,
    phasenFaktoren: {
      'A1-A3': 2.57,
      'A4': 2.46,
      'A5': 3.47,
      'B1-B7': 1.00,
      'C1-C4': 2.17
    }
  },
  {
    nummer: '221-420',
    typ: 'Standard',
    klemmen: 10,
    gesamtFaktor: 1.69,
    phasenFaktoren: {
      'A1-A3': 4.88,
      'A4': 4.67,
      'A5': 5.99,
      'B1-B7': 1.00,
      'C1-C4': 3.95
    }
  },
  {
    nummer: '221-422',
    typ: 'Green Range',
    klemmen: 2,
    gesamtFaktor: 1.00,
    phasenFaktoren: {
      'A1-A3': 1.00,
      'A4': 1.00,
      'A5': 1.00,
      'B1-B7': 1.00,
      'C1-C4': 1.00
    }
  },
  {
    nummer: '221-423',
    typ: 'Green Range',
    klemmen: 3,
    gesamtFaktor: 1.09,
    phasenFaktoren: {
      'A1-A3': 1.46,
      'A4': 1.49,
      'A5': 1.82,
      'B1-B7': 1.00,
      'C1-C4': 1.41
    }
  },
  {
    nummer: '221-425',
    typ: 'Green Range',
    klemmen: 5,
    gesamtFaktor: 1.26,
    phasenFaktoren: {
      'A1-A3': 2.39,
      'A4': 2.49,
      'A5': 3.49,
      'B1-B7': 1.00,
      'C1-C4': 2.23
    }
  },
  {
    nummer: '221-430',
    typ: 'Green Range',
    klemmen: 10,
    gesamtFaktor: 1.65,
    phasenFaktoren: {
      'A1-A3': 4.61,
      'A4': 4.79,
      'A5': 6.03,
      'B1-B7': 1.00,
      'C1-C4': 4.12
    }
  },
  {
    nummer: '221-2411',
    typ: 'Standard',
    klemmen: 2,
    gesamtFaktor: 1.14,
    phasenFaktoren: {
      'A1-A3': 1.85,
      'A4': 1.46,
      'A5': 1.60,
      'B1-B7': 1.00,
      'C1-C4': 1.31
    }
  }
]

const format = (wert: number, stellen = 5) =>
    wert.toLocaleString('de-DE', {
      maximumFractionDigits: stellen
    })

function findeProdukt(nummer: string): Produkt {
  return produkte.find(p => p.nummer === nummer)!
}

function berechneGWP(
    produkt: Produkt,
    menge: number
): number {
  return referenz.gwp * produkt.gesamtFaktor * menge
}

function App() {
  const [suche, setSuche] = useState('')
  const [filter, setFilter] = useState('Alle')
  const [auswahl, setAuswahl] = useState('221-422')
  const [menge, setMenge] = useState('100')

  const [vergleichA, setVergleichA] = useState('221-422')
  const [vergleichB, setVergleichB] = useState('221-425')
  const [vergleichMenge, setVergleichMenge] = useState('100')

  const produkt = findeProdukt(auswahl)

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
    const passtSuche =
        p.nummer.includes(suche.trim()) ||
        p.typ.toLowerCase().includes(suche.toLowerCase())

    const passtFilter =
        filter === 'Alle' || p.typ === filter

    return passtSuche && passtFilter
  })

  const phasenErgebnisse = phasen.map(phase => ({
    phase,
    name: phasenNamen[phase],
    faktor: produkt.phasenFaktoren[phase],
    wert: gueltigeMenge
        ? referenz.phasen[phase] *
        produkt.phasenFaktoren[phase] *
        anzahl
        : 0
  }))

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

  function waehleProdukt(nummer: string) {
    setAuswahl(nummer)

    document.getElementById('rechner')?.scrollIntoView({
      behavior: 'smooth'
    })
  }

  return (
      <div className="app">

        {/* NAVIGATION */}

        <header className="header">
          <a className="logo" href="#">
            WAGO <span>EPD Explorer</span>
          </a>

          <nav>
            <a href="#produkte">Produkte</a>
            <a href="#rechner">EPD-Rechner</a>
            <a href="#ergebnisse">Ergebnisse</a>
            <a href="#vergleich">Vergleich</a>
            <a href="#dokumente">Dokumentation</a>
          </nav>

          <span className="mock-badge">
          MOCK-UP
        </span>
        </header>

        <main>

          {/* STARTSEITE */}

          <section className="hero">
            <div className="hero-content">
            <span className="eyebrow">
              WAGO · SERIE 221
            </span>

              <h1>
                Umweltkennzahlen
                <br />
                einfach berechnen.
              </h1>

              <p>
                Wähle ein Produkt der WAGO-Serie 221,
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
              <span>Referenzprodukt 221-422</span>
              <strong>0,0856</strong>
              <span>kg CO₂-Äq. pro Stück</span>

              <div className="hero-stat-line" />

              <small>
                Datengrundlage: WAGO-221-EPD
              </small>
            </div>
          </section>

          {/* KENNZAHLEN */}

          <div className="stats">
            <div className="stat-card">
              <strong>9</strong>
              <span>Produktvarianten</span>
            </div>

            <div className="stat-card">
              <strong>5</strong>
              <span>Lebenszyklusgruppen</span>
            </div>

            <div className="stat-card">
              <strong>221</strong>
              <span>WAGO-Produktserie</span>
            </div>
          </div>

          {/* PRODUKTAUSWAHL */}

          <section id="produkte" className="section">
            <div className="section-heading">
              <div>
              <span className="eyebrow">
                PRODUKTÜBERSICHT
              </span>

                <h2>WAGO Serie 221</h2>

                <p>
                  Wähle deine Produktvariante.
                  Die verfügbaren EPD-Daten werden
                  automatisch übernommen.
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
                      <div className="connector">
                        <strong>{p.klemmen}</strong>
                        <span>Klemmstellen</span>
                      </div>
                    </div>

                    <h3>WAGO {p.nummer}</h3>

                    <p>
                      Serie 221 · {p.typ}
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
                    onChange={e => setAuswahl(e.target.value)}
                >
                  {produkte.map(p => (
                      <option
                          key={p.nummer}
                          value={p.nummer}
                      >
                        WAGO {p.nummer} – {p.typ}
                      </option>
                  ))}
                </select>

                <label htmlFor="menge">
                  Anzahl der Produkte
                </label>

                <input
                    id="menge"
                    type="number"
                    min="1"
                    step="1"
                    value={menge}
                    onChange={e => setMenge(e.target.value)}
                />

                <div className="quick-buttons">
                  {[1, 10, 100, 1000].map(n => (
                      <button
                          key={n}
                          className={
                            menge === String(n)
                                ? 'active'
                                : ''
                          }
                          onClick={() => setMenge(String(n))}
                      >
                        {format(n, 0)}
                      </button>
                  ))}
                </div>

                {!gueltigeMenge && (
                    <p className="error">
                      Bitte eine gültige Stückzahl
                      größer 0 eingeben.
                    </p>
                )}

                <div className="info-box">
                  <strong>
                    Berechnungsgrundlage
                  </strong>

                  <p>
                    Referenzprodukt: WAGO 221-422
                  </p>

                  <p>
                    Variantenfaktor: {format(
                      produkt.gesamtFaktor,
                      2
                  )}
                  </p>

                  <p>
                    Nutzungsannahme: 1 A
                    gemäß bereitgestellter EPD
                  </p>
                </div>
              </div>

              <div className="panel ergebnis-panel">
              <span className="eyebrow">
                DEIN ERGEBNIS
              </span>

                <h3>WAGO {produkt.nummer}</h3>

                <p>
                  {produkt.typ} ·
                  {' '}{produkt.klemmen} Klemmstellen
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
                  kg CO₂-Äq.
                </span>
                </div>

                <div className="result-row">
                  <span>Wert pro Stück</span>

                  <strong>
                    {format(wertProStueck, 6)}
                    {' '}kg CO₂-Äq.
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
                  {format(referenz.gwp, 4)}
                  {' × '}
                  {format(produkt.gesamtFaktor, 2)}
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
              <h3>
                Lebenszyklusphasen
              </h3>

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
                        {gueltigeMenge
                            ? format(p.wert, 6)
                            : '–'}
                      </strong>
                    </div>
                ))}
              </div>

              <h3>
                Detaillierte Ergebnistabelle
              </h3>

              <div className="table-scroll">
                <table>
                  <thead>
                  <tr>
                    <th>Phase</th>
                    <th>Bezeichnung</th>
                    <th>Faktor</th>
                    <th>kg CO₂-Äq.</th>
                  </tr>
                  </thead>

                  <tbody>
                  {phasenErgebnisse.map(p => (
                      <tr key={p.phase}>
                        <td>{p.phase}</td>
                        <td>{p.name}</td>
                        <td>
                          {format(p.faktor, 2)}
                        </td>
                        <td>
                          {gueltigeMenge
                              ? format(p.wert, 6)
                              : '–'}
                        </td>
                      </tr>
                  ))}
                  </tbody>
                </table>
              </div>

              <p className="hint">
                Die Berechnung erfolgt mit gerundeten
                Faktoren aus der bereitgestellten EPD.
                Gesamtwert und Phasenwerte werden
                unabhängig hochgerechnet. Dadurch sind
                Rundungsdifferenzen möglich.
                Modul D ist nicht im Gesamtwert enthalten.
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
                  zweier Produkte der Serie 221
                  für dieselbe Stückzahl.
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
                    {produkte.map(p => (
                        <option
                            key={p.nummer}
                            value={p.nummer}
                        >
                          WAGO {p.nummer} – {p.typ}
                        </option>
                    ))}
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
                    {produkte.map(p => (
                        <option
                            key={p.nummer}
                            value={p.nummer}
                        >
                          WAGO {p.nummer} – {p.typ}
                        </option>
                    ))}
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
                    {produktA.typ} ·
                    {' '}{produktA.klemmen} Klemmstellen
                  </p>

                  <div className="vergleich-wert">
                    {gueltigerVergleich
                        ? format(wertA)
                        : '–'}
                  </div>

                  <span>kg CO₂-Äq.</span>

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
                    {produktB.typ} ·
                    {' '}{produktB.klemmen} Klemmstellen
                  </p>

                  <div className="vergleich-wert">
                    {gueltigerVergleich
                        ? format(wertB)
                        : '–'}
                  </div>

                  <span>kg CO₂-Äq.</span>

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
                      {' '}kg CO₂-Äq.
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
                  der WAGO-Serie 221.
                </p>
              </div>
            </div>

            <div className="panel">
            <span className="tag green">
              PEP ecopassport
            </span>

              <h3 className="dokument-titel">
                WAGO – Serie 221
              </h3>

              <div className="dokument-grid">
                <div>
                <span>
                  Registrierungsnummer
                </span>

                  <strong>
                    WAGO-00001-V01.01-EN
                  </strong>
                </div>

                <div>
                <span>
                  Referenzprodukt
                </span>

                  <strong>
                    221-422 Green Range
                  </strong>
                </div>

                <div>
                <span>
                  Ausgabedatum laut PDF
                </span>

                  <strong>
                    Februar 2026
                  </strong>
                </div>

                <div>
                <span>
                  Referenznutzungsdauer
                </span>

                  <strong>
                    30 Jahre
                  </strong>
                </div>

                <div>
                <span>
                  Programmbetreiber
                </span>

                  <strong>
                    PEP ecopassport
                  </strong>
                </div>

                <div>
                <span>
                  Aktuelle Datenquelle
                </span>

                  <strong>
                    Lokale Mock-Daten
                  </strong>
                </div>
              </div>

              <a
                  className="dokument-link"
                  href="https://www.pep-ecopassport.org/"
                  target="_blank"
                  rel="noopener noreferrer"
              >
                PEP ecopassport öffnen ↗
              </a>

              <p className="hint">
                Die Projektunterlagen befinden sich
                teilweise noch in Entwicklung.
                Dieses Mock-up stellt keine
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
              Hochschulprojekt · Mock-up mit lokalen
              EPD-Daten · Keine Live-AAS-Anbindung
            </p>
          </footer>

        </main>
      </div>
  )
}

export default App