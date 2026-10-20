import * as XLSX from 'xlsx'
import { cleanName, type Employee } from './parse'
import { REGIONS, type Region } from './exportExcel'

export interface LocationRecord {
  name: string
  zones: Region[]
}

export interface LocationFile {
  fileName: string
  people: LocationRecord[]
}

export interface Match {
  // Region to auto-assign; only set when the person has exactly one zone.
  region?: Region
  zones: Region[]
  matchedName?: string
  method: 'exact' | 'fuzzy' | 'none'
}

const norm = (v: unknown) => String(v ?? '').trim().toLowerCase()

const ZONE_ALIASES: Record<string, Region> = {
  sydney: 'Sydney',
  syd: 'Sydney',
  newcastle: 'Newcastle',
  wollongong: 'Wgong',
  wgong: 'Wgong',
  gong: 'Wgong',
}

export const parseZones = (v: unknown): Region[] => {
  const found = new Set<Region>()
  for (const part of String(v ?? '').split(/[,;/|]/)) {
    const r = ZONE_ALIASES[norm(part).replace(/[^a-z]/g, '')]
    if (r) found.add(r)
  }
  return REGIONS.filter((r) => found.has(r))
}

export async function parseLocationFile(file: File): Promise<LocationFile> {
  const wb = XLSX.read(await file.arrayBuffer())
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' })

  const NAME_COLS = ['name', 'employee', 'employee name']
  const ZONE_COLS = ['zones', 'zone', 'location', 'region', 'branch', 'office']
  const headerIdx = rows.findIndex(
    (r) => r.some((c) => NAME_COLS.includes(norm(c))) && r.some((c) => ZONE_COLS.includes(norm(c))),
  )
  if (headerIdx === -1)
    throw new Error('Could not find a header row with a "Name" column and a "Zones"/"Location" column.')

  const header = rows[headerIdx].map(norm)
  const nameCol = header.findIndex((h) => NAME_COLS.includes(h))
  const zoneCol = header.findIndex((h) => ZONE_COLS.includes(h))

  const people: LocationRecord[] = []
  for (const r of rows.slice(headerIdx + 1)) {
    const name = cleanName(r[nameCol])
    if (name) people.push({ name, zones: parseZones(r[zoneCol]) })
  }
  return { fileName: file.name, people }
}

// --- Name matching -------------------------------------------------------
// Payroll and the staff list don't always spell names the same way
// ("Nicholas" vs "Nick", "Russel" vs "Russell", middle names), so after an
// exact match we fall back to looser rules, accepting only unambiguous hits.

const NICKNAMES: Record<string, string> = {
  nick: 'nicholas', nic: 'nicholas', nicolas: 'nicholas',
  mike: 'michael', mick: 'michael',
  matt: 'matthew', tom: 'thomas', tommy: 'thomas',
  josh: 'joshua', chris: 'christopher', alex: 'alexander',
  sam: 'samuel', dan: 'daniel', danny: 'daniel',
  will: 'william', bill: 'william', ben: 'benjamin',
  jake: 'jacob', joe: 'joseph', jim: 'james', jimmy: 'james',
  rob: 'robert', bob: 'robert', tony: 'anthony', andy: 'andrew',
  steve: 'stephen', dave: 'david', pete: 'peter', jon: 'jonathan',
  zach: 'zachary', zac: 'zachary', nate: 'nathan', jack: 'john',
}

const tokens = (name: string) =>
  name.toLowerCase().replace(/[^a-z\s'-]/g, '').replace(/['-]/g, '').split(/\s+/).filter(Boolean)

const first = (t: string[]) => NICKNAMES[t[0]] ?? t[0]

function editDistance(a: string, b: string) {
  const d = Array.from({ length: b.length + 1 }, (_, i) => i)
  for (let i = 1; i <= a.length; i++) {
    let prev = d[0]
    d[0] = i
    for (let j = 1; j <= b.length; j++) {
      const tmp = d[j]
      d[j] = Math.min(d[j] + 1, d[j - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1))
      prev = tmp
    }
  }
  return d[b.length]
}

const fuzzyRules: ((a: string[], b: string[]) => boolean)[] = [
  // Same first name (nicknames resolved) and same surname, ignoring middle names.
  (a, b) => first(a) === first(b) && a.at(-1) === b.at(-1),
  // Same first name and surname off by one letter.
  (a, b) => {
    const [la, lb] = [a.at(-1)!, b.at(-1)!]
    return first(a) === first(b) && Math.min(la.length, lb.length) >= 4 && editDistance(la, lb) <= 1
  },
]

function findPerson(name: string, people: LocationRecord[]) {
  const key = tokens(name).join(' ')
  const exact = people.filter((p) => tokens(p.name).join(' ') === key)
  if (exact.length === 1) return { person: exact[0], method: 'exact' as const }

  const t = tokens(name)
  if (t.length < 2) return undefined
  for (const rule of fuzzyRules) {
    const hits = people.filter((p) => {
      const pt = tokens(p.name)
      return pt.length >= 2 && rule(t, pt)
    })
    if (hits.length === 1) return { person: hits[0], method: 'fuzzy' as const }
    if (hits.length > 1) return undefined
  }
  return undefined
}

export function matchEmployees(employees: Employee[], people: LocationRecord[]) {
  const result: Record<string, Match> = {}
  for (const e of employees) {
    const hit = findPerson(e.name, people)
    result[e.name] = hit
      ? {
          zones: hit.person.zones,
          region: hit.person.zones.length === 1 ? hit.person.zones[0] : undefined,
          matchedName: hit.person.name,
          method: hit.method,
        }
      : { zones: [], method: 'none' }
  }
  return result
}
