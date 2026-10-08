// =====================================================
// Verpackungseinheiten laut WAGO-Shop
// Quelle: www.wago.com, Angabe "VPE (UVPE)" auf der
// Produktseite, abgerufen am 08.10.2026.
// Nicht Teil der EPD und daher nicht in der AAS.
// =====================================================

export type Verpackung = {
  vpe: number   // Verpackungseinheit (Karton)
  uvpe: number  // Unterverpackungseinheit, kleinste Packung
}

const VERPACKUNG: Record<string, Verpackung> = {
  '221-412': { vpe: 1000, uvpe: 100 },
  '221-413': { vpe: 500, uvpe: 50 },
  '221-415': { vpe: 250, uvpe: 25 },
  '221-420': { vpe: 150, uvpe: 15 },
  '221-422': { vpe: 1000, uvpe: 100 },
  '221-423': { vpe: 500, uvpe: 50 },
  '221-425': { vpe: 250, uvpe: 25 },
  '221-430': { vpe: 150, uvpe: 15 },
  '221-2411': { vpe: 600, uvpe: 60 },
  '2273-202': { vpe: 1000, uvpe: 100 },
  '2273-203': { vpe: 1000, uvpe: 100 },
  '2273-204': { vpe: 1000, uvpe: 100 },
  '2273-205': { vpe: 1000, uvpe: 100 },
  '2273-208': { vpe: 500, uvpe: 50 }
}

export function verpackung(nummer: string): Verpackung | null {
  return VERPACKUNG[nummer] ?? null
}

// Produktfotos: Hauptbild der Shop-Produktseite (480 px),
// abgelegt unter public/produktbilder/<Artikel>.jpg
export function produktbild(nummer: string): string | null {
  return nummer in VERPACKUNG
      ? `${import.meta.env.BASE_URL}produktbilder/${nummer}.jpg`
      : null
}

// Leitet auf die Produktseite im Shop weiter
export function shopUrl(nummer: string): string {
  return `https://www.wago.com/de/p/${encodeURIComponent(nummer)}`
}
