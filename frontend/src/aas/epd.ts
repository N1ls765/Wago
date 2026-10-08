// =====================================================
// EPD-Teilmodell (IDTA "EPD Type III") -> Produktdaten
// fuer die Oberflaeche
// =====================================================

import {
  anhangUrl,
  istEpdSubmodel,
  ladeShells,
  ladeSubmodelWerte,
  submodelIds
} from './client'

export type Phase = 'A1-A3' | 'A4' | 'A5' | 'B1-B7' | 'C1-C4'

export const phasen: Phase[] = [
  'A1-A3', 'A4', 'A5', 'B1-B7', 'C1-C4'
]

export const phasenNamen: Record<Phase, string> = {
  'A1-A3': 'Herstellung',
  'A4': 'Transport',
  'A5': 'Installation',
  'B1-B7': 'Nutzung',
  'C1-C4': 'Entsorgung'
}

export type Produkt = {
  nummer: string
  serie: string
  typ: 'Standard' | 'Green Range'
  klemmen: number | null
  name: string
  submodelId: string

  // Treibhauspotenzial (GWP-total) pro Stueck.
  // null = in der EPD fuer diese Phase nichts deklariert
  einheit: string
  phasen: Record<Phase, number | null>
  modulD: number | null
  gwpProStueck: number

  dokument: {
    kennung: string
    ausgabe: string
    gueltigBis: string
    programmbetreiber: string
    verifizierung: string
    nutzungsdauer: string
    epdUrl: string
    bildUrl: string | null
  }
}

// ---------- ValueOnly-JSON des Teilmodells (nur genutzte Felder)

type Sprachtext = Record<string, string>[]

type StageValue = {
  stageCode: string
  valueStatus: string
  value?: string
}

type EpdWerte = {
  manufacturerProduct: {
    productArticleNumberOfManufacturer: string
    productName?: Sprachtext
    productDescription?: Sprachtext
    referenceServiceLife?: string
    productImage?: { contentType: string, value: string }
  }
  identificationPublication: {
    dateOfIssue?: string
    validUntil?: string
    documentIds?: { documentIdentifier: string }[]
  }
  programOperatorVerification?: {
    programOperator?: { Company?: string }
    verificationType?: string
  }
  EnvironmentalResults: {
    indicatorCode: string
    unit: string
    stageValues: StageValue[]
  }[]
}

function text(wert: Sprachtext | undefined): string {
  if (!wert || wert.length === 0) return ''
  const alle = Object.assign({}, ...wert) as Record<string, string>
  return alle.de ?? alle.en ?? Object.values(alle)[0] ?? ''
}

// stageCode aus dem Teilmodell -> Phasengruppe der Oberflaeche.
// Die Serien liefern unterschiedlich feine Codes
// (z. B. "C1-C4" oder "C1" + "C2-C4").
function phaseZu(stageCode: string): Phase | 'D' | null {
  if (/^A[1-3]/.test(stageCode)) return 'A1-A3'
  if (stageCode === 'A4') return 'A4'
  if (stageCode === 'A5') return 'A5'
  if (stageCode.startsWith('B')) return 'B1-B7'
  if (stageCode.startsWith('C')) return 'C1-C4'
  if (stageCode === 'D') return 'D'
  return null
}

function summiere(werte: StageValue[]) {
  const summen: Record<Phase | 'D', number | null> = {
    'A1-A3': null, 'A4': null, 'A5': null,
    'B1-B7': null, 'C1-C4': null, 'D': null
  }

  for (const sv of werte) {
    const phase = phaseZu(sv.stageCode)
    if (phase === null) continue
    if (sv.valueStatus !== 'Declared' || sv.value === undefined) continue

    summen[phase] = (summen[phase] ?? 0) + Number(sv.value)
  }

  return summen
}

function zuProdukt(submodelId: string, w: EpdWerte): Produkt {
  const mp = w.manufacturerProduct
  const nummer = mp.productArticleNumberOfManufacturer
  const beschreibung = text(mp.productDescription)

  const gwp = w.EnvironmentalResults.find(
      e => e.indicatorCode === 'GWP-total'
  )
  if (!gwp) {
    throw new Error(`Kein GWP-total im EPD von ${nummer}`)
  }

  const summen = summiere(gwp.stageValues)
  const phasenWerte = Object.fromEntries(
      phasen.map(p => [p, summen[p]])
  ) as Record<Phase, number | null>

  // Die Skripte schreiben die Anzahl in die Produktbeschreibung,
  // z. B. "..., green range, 2 clamping units. ..."
  const klemmenText = beschreibung.match(/(\d+)\s+clamping unit/i)

  return {
    nummer,
    serie: nummer.split('-')[0],
    typ: /green range/i.test(beschreibung)
        ? 'Green Range'
        : 'Standard',
    klemmen: klemmenText ? Number(klemmenText[1]) : null,
    name: text(mp.productName),
    submodelId,

    einheit: gwp.unit,
    phasen: phasenWerte,
    modulD: summen.D,
    // Modul D gehoert nicht zum Gesamtwert
    gwpProStueck: phasen.reduce(
        (summe, p) => summe + (phasenWerte[p] ?? 0), 0
    ),

    dokument: {
      kennung: w.identificationPublication.documentIds?.[0]
          ?.documentIdentifier ?? '',
      ausgabe: w.identificationPublication.dateOfIssue ?? '',
      gueltigBis: w.identificationPublication.validUntil ?? '',
      programmbetreiber: w.programOperatorVerification
          ?.programOperator?.Company ?? '',
      verifizierung: w.programOperatorVerification
          ?.verificationType ?? '',
      nutzungsdauer: mp.referenceServiceLife ?? '',
      epdUrl: anhangUrl(submodelId, 'epdDocument'),
      bildUrl: mp.productImage
          ? anhangUrl(submodelId, 'manufacturerProduct.productImage')
          : null
    }
  }
}

function vergleicheNummern(a: Produkt, b: Produkt): number {
  return a.nummer.localeCompare(b.nummer, 'de', { numeric: true })
}

// Alle AAS lesen, jeweils das EPD-Teilmodell suchen und
// in Produkte umwandeln.
export async function ladeProdukte(): Promise<Produkt[]> {
  const shells = await ladeShells()

  const jeShell = await Promise.all(shells.map(async shell => {
    for (const id of submodelIds(shell)) {
      if (await istEpdSubmodel(id)) {
        return zuProdukt(id, await ladeSubmodelWerte<EpdWerte>(id))
      }
    }
    return null
  }))

  return jeShell
      .filter((p): p is Produkt => p !== null)
      .sort(vergleicheNummern)
}
