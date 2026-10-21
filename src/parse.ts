import * as XLSX from 'xlsx'

export interface Employee {
  id: string
  name: string
  leaveType: string
  units: number
  amount: number
}

export interface ParsedSheet {
  title: string
  employees: Employee[]
}

export const LEAVE_KINDS = ['Annual', 'Sick', 'RDO'] as const
export type LeaveKind = (typeof LEAVE_KINDS)[number]
export const LEAVE_LABELS: Record<LeaveKind, string> = { Annual: 'Annual leave', Sick: 'Sick leave', RDO: 'RDO' }

// Payroll exports name these differently ("Annual Leave", "Personal/Carer's Leave", "Sick Leave"…).
export const leaveKind = (leaveType: string): LeaveKind | undefined => {
  const t = leaveType.toLowerCase()
  if (/rdo|rostered day/.test(t)) return 'RDO'
  if (t.includes('annual')) return 'Annual'
  if (/sick|personal|carer/.test(t)) return 'Sick'
  return undefined
}

// RDO balances above this many hours are flagged in red.
export const RDO_LIMIT = 38
export const isOverRdoLimit = (e: Employee) => leaveKind(e.leaveType) === 'RDO' && e.units > RDO_LIMIT

const norm = (v: unknown) => String(v ?? '').trim().toLowerCase()

// Collapses double spaces / trailing spaces that payroll exports often contain.
export const cleanName = (v: unknown) => String(v ?? '').replace(/\s+/g, ' ').trim()

const toNumber = (v: unknown) => {
  if (typeof v === 'number') return v
  const n = parseFloat(String(v ?? '').replace(/[$,\s]/g, ''))
  return Number.isFinite(n) ? n : 0
}

export async function parseLeaveFile(file: File): Promise<ParsedSheet> {
  const wb = XLSX.read(await file.arrayBuffer())
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' })

  const headerIdx = rows.findIndex((r) =>
    r.some((c) => ['employee', 'name', 'employee name'].includes(norm(c))),
  )
  if (headerIdx === -1) throw new Error('Could not find a header row with an "Employee" or "Name" column.')

  const header = rows[headerIdx].map(norm)
  const col = (...names: string[]) => header.findIndex((h) => names.includes(h))
  const nameCol = col('employee', 'name', 'employee name')
  const typeCol = col('leave type')
  const unitsCol = col('units', 'hours', 'balance')
  const amountCol = col('amount', 'value', '$')

  const title = rows
    .slice(0, headerIdx)
    .map((r) => cleanName(r[0]))
    .filter(Boolean)
    .join(' · ')

  const employees: Employee[] = []
  for (const r of rows.slice(headerIdx + 1)) {
    const name = cleanName(r[nameCol])
    if (!name || norm(name) === 'total') continue
    employees.push({
      id: `${employees.length}`,
      name,
      leaveType: typeCol >= 0 ? cleanName(r[typeCol]) : '',
      units: unitsCol >= 0 ? toNumber(r[unitsCol]) : 0,
      amount: amountCol >= 0 ? toNumber(r[amountCol]) : 0,
    })
  }
  return { title, employees }
}
