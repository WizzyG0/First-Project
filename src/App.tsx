import { useEffect, useMemo, useState, type ReactNode } from 'react'
import {
  isOverRdoLimit,
  LEAVE_KINDS,
  LEAVE_LABELS,
  leaveKind,
  parseLeaveFile,
  RDO_LIMIT,
  type Employee,
  type LeaveKind,
} from './parse'
import { exportByRegion, REGIONS, type Region } from './exportExcel'
import { matchEmployees, parseLocationFile, type LocationFile } from './locations'
import './App.css'

type Tags = Record<string, Region | undefined>
// Manual corrections on top of the locations file, remembered per employee name.
const OVERRIDES_KEY = 'leave-region-overrides'
// The locations file changes rarely, so it's kept between visits.
const LOCATIONS_KEY = 'leave-locations'
// Employees deleted from the list stay hidden on future uploads until restored.
const REMOVED_KEY = 'leave-removed'

function load<T>(key: string, fallback: T): T {
  try {
    const raw = localStorage.getItem(key)
    return raw ? JSON.parse(raw) : fallback
  } catch {
    return fallback
  }
}

function save(key: string, value: unknown) {
  try {
    if (value == null) localStorage.removeItem(key)
    else localStorage.setItem(key, JSON.stringify(value))
  } catch {
    /* storage unavailable */
  }
}

type SortOrder = 'name' | 'units-desc' | 'units-asc'
const byName = (a: Employee, b: Employee) => a.name.localeCompare(b.name) || a.leaveType.localeCompare(b.leaveType)
const SORTS: Record<SortOrder, { label: string; compare: (a: Employee, b: Employee) => number }> = {
  name: { label: 'Name A–Z', compare: byName },
  'units-desc': { label: 'Units: highest to lowest', compare: (a, b) => b.units - a.units || byName(a, b) },
  'units-asc': { label: 'Units: lowest to highest', compare: (a, b) => a.units - b.units || byName(a, b) },
}

const fmtHours = (n: number) => n.toLocaleString(undefined, { maximumFractionDigits: 2 })
const fmtMoney = (n: number) => n.toLocaleString(undefined, { style: 'currency', currency: 'AUD' })

function DropZone({
  label,
  accept,
  onFile,
  children,
}: {
  label: string
  accept: string
  onFile: (file?: File) => void
  children: ReactNode
}) {
  const [dragging, setDragging] = useState(false)
  return (
    <label
      className={`drop ${dragging ? 'dragging' : ''}`}
      onDragOver={(e) => {
        e.preventDefault()
        setDragging(true)
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault()
        setDragging(false)
        onFile(e.dataTransfer.files[0])
      }}
    >
      <input
        type="file"
        accept={accept}
        onChange={(e) => {
          onFile(e.target.files?.[0])
          e.target.value = ''
        }}
      />
      <span className="drop-label">{label}</span>
      <span>{children}</span>
    </label>
  )
}

export default function App() {
  const [employees, setEmployees] = useState<Employee[]>([])
  const [title, setTitle] = useState('')
  const [fileName, setFileName] = useState('')
  const [locations, setLocations] = useState<LocationFile | null>(() => load(LOCATIONS_KEY, null))
  const [error, setError] = useState('')
  const [overrides, setOverrides] = useState<Tags>(() => load(OVERRIDES_KEY, {}))
  const [filter, setFilter] = useState('')
  const [attentionOnly, setAttentionOnly] = useState(false)
  const [removed, setRemoved] = useState<string[]>(() => load(REMOVED_KEY, []))
  // Which leave balance to show; undefined shows every row in the file.
  const [kind, setKind] = useState<LeaveKind>()
  const [sort, setSort] = useState<SortOrder>('name')

  useEffect(() => save(OVERRIDES_KEY, overrides), [overrides])
  useEffect(() => save(LOCATIONS_KEY, locations), [locations])
  useEffect(() => save(REMOVED_KEY, removed), [removed])

  const removeEmployee = (name: string) => setRemoved((r) => (r.includes(name) ? r : [...r, name]))
  const restoreEmployee = (name: string) => setRemoved((r) => r.filter((n) => n !== name))

  async function handleLeaveFile(file?: File) {
    if (!file) return
    setError('')
    try {
      const parsed = await parseLeaveFile(file)
      setEmployees(parsed.employees)
      setTitle(parsed.title)
      setFileName(file.name)
    } catch (e) {
      setError(`Leave balance file: ${e instanceof Error ? e.message : String(e)}`)
    }
  }

  async function handleLocationFile(file?: File) {
    if (!file) return
    setError('')
    try {
      setLocations(await parseLocationFile(file))
    } catch (e) {
      setError(`Locations file: ${e instanceof Error ? e.message : String(e)}`)
    }
  }

  const active = useMemo(
    () =>
      employees
        .filter((e) => !removed.includes(e.name) && (!kind || leaveKind(e.leaveType) === kind))
        .sort(SORTS[sort].compare),
    [employees, removed, kind, sort],
  )
  const kindLabel = kind ? LEAVE_LABELS[kind] : ''
  const removedHere = employees.filter((e) => removed.includes(e.name))

  const matches = useMemo(
    () => matchEmployees(employees, locations?.people ?? []),
    [employees, locations],
  )

  const tags = useMemo(() => {
    const t: Tags = {}
    for (const e of employees) t[e.name] = overrides[e.name] ?? matches[e.name]?.region
    return t
  }, [employees, overrides, matches])

  // Picking the file's own region (or re-clicking the current one) drops the override.
  const setTag = (name: string, region: Region) =>
    setOverrides((o) => {
      const next = { ...o }
      if (region === tags[name] || region === matches[name]?.region) delete next[name]
      else next[name] = region
      return next
    })
  const resetTag = (name: string) =>
    setOverrides((o) => {
      const next = { ...o }
      delete next[name]
      return next
    })

  const isOverridden = (name: string) =>
    overrides[name] !== undefined && overrides[name] !== matches[name]?.region
  const needsAttention = (e: Employee) =>
    !tags[e.name] || matches[e.name]?.method === 'fuzzy' || isOverridden(e.name)

  const visible = active.filter(
    (e) => e.name.toLowerCase().includes(filter.toLowerCase()) && (!attentionOnly || needsAttention(e)),
  )
  const untagged = active.filter((e) => !tags[e.name]).length
  const autoTagged = active.filter((e) => !isOverridden(e.name) && matches[e.name]?.region).length

  const grouped = useMemo(
    () =>
      REGIONS.map((region) => ({
        region,
        people: active.filter((e) => tags[e.name] === region),
      })),
    [active, tags],
  )

  function note(e: Employee): ReactNode {
    const m = matches[e.name]
    if (isOverridden(e.name))
      return (
        <>
          Set by hand{m?.region ? ` (file says ${m.region})` : ''} ·{' '}
          <button className="link" onClick={() => resetTag(e.name)}>
            reset
          </button>
        </>
      )
    if (!locations) return null
    if (!m || m.method === 'none') return 'Not found in locations file — choose a region'
    const parts: string[] = []
    if (m.method === 'fuzzy') parts.push(`Matched to “${m.matchedName}”`)
    if (!tags[e.name])
      parts.push(m.zones.length ? `Works in ${m.zones.join(', ')} — choose one` : 'No zone listed — choose a region')
    return parts.join(' · ') || null
  }

  return (
    <div className="app">
      <header>
        <h1>Leave Balance by Region</h1>
        {title && <p className="subtitle">{title}</p>}
      </header>

      <div className="uploads">
        <DropZone label="1. Leave balance report" accept=".xlsx,.xls,.csv" onFile={handleLeaveFile}>
          {fileName ? (
            <>
              <strong>{fileName}</strong> — {employees.length} employees. Click or drop to replace.
            </>
          ) : (
            'Click to choose or drop the leave balance Excel file here'
          )}
        </DropZone>
        <DropZone label="2. Employee locations" accept=".csv,.xlsx,.xls" onFile={handleLocationFile}>
          {locations ? (
            <>
              <strong>{locations.fileName}</strong> — {locations.people.length} people. Click or drop to
              replace.
            </>
          ) : (
            'Click to choose or drop the employee list (with a Zones column) here'
          )}
        </DropZone>
      </div>
      {error && <p className="error">{error}</p>}

      {employees.length > 0 && (
        <>
          <section className="panel">
            <div className="toolbar">
              <input
                className="search"
                placeholder="Search names…"
                value={filter}
                onChange={(e) => setFilter(e.target.value)}
              />
              <label className="check">
                <input
                  type="checkbox"
                  checked={attentionOnly}
                  onChange={(e) => setAttentionOnly(e.target.checked)}
                />
                Needs attention only
              </label>
              {locations && <span className="muted">{autoTagged} auto-tagged</span>}
              <span className={untagged ? 'warn' : 'ok'}>
                {untagged ? `${untagged} untagged` : 'All tagged'}
              </span>
              <select
                className="sort"
                aria-label="Sort"
                value={sort}
                onChange={(e) => setSort(e.target.value as SortOrder)}
              >
                {Object.entries(SORTS).map(([value, { label }]) => (
                  <option key={value} value={value}>
                    {label}
                  </option>
                ))}
              </select>
              <div className="kinds" role="group" aria-label="Leave type">
                {LEAVE_KINDS.map((k) => (
                  <button
                    key={k}
                    className={kind === k ? 'active' : ''}
                    aria-pressed={kind === k}
                    onClick={() => setKind(kind === k ? undefined : k)}
                  >
                    {LEAVE_LABELS[k]} balance
                  </button>
                ))}
              </div>
              <button
                className="primary"
                onClick={() => exportByRegion(active, tags, [title, kindLabel].filter(Boolean).join(' · '))}
              >
                Export Excel
              </button>
            </div>

            <table className="list">
              <thead>
                <tr>
                  <th>Employee</th>
                  <th>Leave type</th>
                  <th className="num">Units (hrs)</th>
                  <th className="num">Amount</th>
                  <th>Region</th>
                  <th className="actions" aria-label="Delete" />
                </tr>
              </thead>
              <tbody>
                {visible.map((e) => {
                  const n = note(e)
                  const zones = matches[e.name]?.zones ?? []
                  return (
                    <tr key={e.id} className={tags[e.name] ? '' : 'untagged'}>
                      <td>
                        {e.name}
                        {n && <div className="note">{n}</div>}
                      </td>
                      <td>{e.leaveType}</td>
                      <td className={`num ${isOverRdoLimit(e) ? 'over-limit' : ''}`} title={isOverRdoLimit(e) ? `RDO over ${RDO_LIMIT} hrs` : undefined}>
                        {fmtHours(e.units)}
                      </td>
                      <td className="num">{fmtMoney(e.amount)}</td>
                      <td>
                        <div className="tags">
                          {REGIONS.map((r) => (
                            <button
                              key={r}
                              className={`tag tag-${r.toLowerCase()} ${tags[e.name] === r ? 'active' : ''} ${
                                !tags[e.name] && zones.includes(r) ? 'suggested' : ''
                              }`}
                              onClick={() => setTag(e.name, r)}
                            >
                              {r}
                            </button>
                          ))}
                        </div>
                      </td>
                      <td className="actions">
                        <button
                          className="delete"
                          title={`Remove ${e.name} from the list and export`}
                          onClick={() => removeEmployee(e.name)}
                        >
                          Delete
                        </button>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
            {kind && active.length === 0 && (
              <p className="empty muted">No {kindLabel.toLowerCase()} rows in this file.</p>
            )}

            {removedHere.length > 0 && (
              <div className="removed">
                <span className="muted">Deleted ({removedHere.length}):</span>
                {removedHere.map((e) => (
                  <span key={e.id} className="removed-chip">
                    {e.name}
                    <button className="link" onClick={() => restoreEmployee(e.name)}>
                      restore
                    </button>
                  </span>
                ))}
              </div>
            )}
          </section>

          <h2>Preview{kind && ` — ${kindLabel}`}</h2>
          <div className="regions">
            {grouped.map(({ region, people }) => (
              <table key={region} className={`region region-${region.toLowerCase()}`}>
                <thead>
                  <tr>
                    <th colSpan={4} className="region-title">
                      {region} ({people.length})
                    </th>
                  </tr>
                  <tr>
                    <th>Employee</th>
                    <th>Leave type</th>
                    <th className="num">Units</th>
                    <th className="num">Amount</th>
                  </tr>
                </thead>
                <tbody>
                  {people.map((p) => (
                    <tr key={p.id}>
                      <td>{p.name}</td>
                      <td>{p.leaveType}</td>
                      <td className={`num ${isOverRdoLimit(p) ? 'over-limit' : ''}`}>{fmtHours(p.units)}</td>
                      <td className="num">{fmtMoney(p.amount)}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr>
                    <td colSpan={2}>Total</td>
                    <td className="num">{fmtHours(people.reduce((s, p) => s + p.units, 0))}</td>
                    <td className="num">{fmtMoney(people.reduce((s, p) => s + p.amount, 0))}</td>
                  </tr>
                </tfoot>
              </table>
            ))}
          </div>
        </>
      )}
    </div>
  )
}
