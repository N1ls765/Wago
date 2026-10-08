// =====================================================
// Zugriff auf den AAS-Server (IDTA Part 2 / API v3.0)
// Getestet mit Eclipse BaSyx AAS Environment, siehe
// ../aas-server/docker-compose.yml
// =====================================================

// Im Dev-Betrieb leitet Vite '/aas' an den AAS-Server weiter
// (vite.config.ts). Fuer einen anderen Server, z. B. die
// WAGO SmartData API, VITE_AAS_URL in .env.local setzen.
export const AAS_URL: string =
    import.meta.env.VITE_AAS_URL ?? '/aas'

export const EPD_SEMANTIC_ID =
    'https://admin-shell.io/idta/EPD/EPDTypeIII/1/0'

type Reference = {
  type: string
  keys: { type: string, value: string }[]
}

export type Shell = {
  id: string
  idShort?: string
  submodels?: Reference[]
}

type SubmodelMetadata = {
  id: string
  idShort?: string
  semanticId?: Reference
}

// Identifier werden in der URL base64url-kodiert (ohne Padding).
export function kodiereId(id: string): string {
  const bytes = new TextEncoder().encode(id)
  let binaer = ''
  bytes.forEach(b => { binaer += String.fromCharCode(b) })

  return btoa(binaer)
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '')
}

async function holeJson<T>(pfad: string): Promise<T> {
  const antwort = await fetch(AAS_URL + pfad)

  if (!antwort.ok) {
    throw new Error(
        `AAS-Server antwortet mit ${antwort.status} auf ${pfad}`
    )
  }

  return antwort.json() as Promise<T>
}

export async function ladeShells(): Promise<Shell[]> {
  const seite = await holeJson<{ result: Shell[] }>(
      '/shells?limit=1000'
  )
  return seite.result
}

export function submodelIds(shell: Shell): string[] {
  return (shell.submodels ?? [])
      .map(ref => ref.keys.at(-1)?.value)
      .filter((id): id is string => id !== undefined)
}

export async function istEpdSubmodel(id: string): Promise<boolean> {
  const meta = await holeJson<SubmodelMetadata>(
      `/submodels/${kodiereId(id)}/$metadata`
  )
  return meta.semanticId?.keys.some(
      k => k.value === EPD_SEMANTIC_ID
  ) ?? false
}

// ValueOnly-Darstellung: nur idShort -> Wert, ohne Metadaten.
export function ladeSubmodelWerte<T>(id: string): Promise<T> {
  return holeJson<T>(`/submodels/${kodiereId(id)}/$value`)
}

// Direkter Link auf eine Datei im Teilmodell (PDF, Bild).
// idShortPath mit Punkten, z. B. 'manufacturerProduct.productImage'
export function anhangUrl(submodelId: string, idShortPath: string): string {
  return `${AAS_URL}/submodels/${kodiereId(submodelId)}` +
      `/submodel-elements/${idShortPath}/attachment`
}
