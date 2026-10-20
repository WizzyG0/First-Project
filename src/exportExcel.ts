import ExcelJS from 'exceljs'
import type { Employee } from './parse'

export const REGIONS = ['Sydney', 'Newcastle', 'Wgong'] as const
export type Region = (typeof REGIONS)[number]

const HEADERS = ['Employee', 'Units (hrs)', 'Amount ($)', 'Region']
const TABLE_WIDTH = HEADERS.length
const GAP = 1 // empty columns between region tables

const thin = { style: 'thin' as const, color: { argb: 'FFBFBFBF' } }
const border = { top: thin, left: thin, bottom: thin, right: thin }

export async function exportByRegion(
  employees: Employee[],
  tags: Record<string, Region | undefined>,
  title: string,
) {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Leave by Region')

  ws.getCell(1, 1).value = title || 'Leave Balances by Region'
  ws.getCell(1, 1).font = { bold: true, size: 14 }

  REGIONS.forEach((region, i) => {
    const startCol = 1 + i * (TABLE_WIDTH + GAP)
    const people = employees
      .filter((e) => tags[e.name] === region)
      .sort((a, b) => a.name.localeCompare(b.name))

    // Region title row
    ws.mergeCells(3, startCol, 3, startCol + TABLE_WIDTH - 1)
    const titleCell = ws.getCell(3, startCol)
    titleCell.value = `${region} (${people.length})`
    titleCell.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } }
    titleCell.alignment = { horizontal: 'center' }
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } }

    // Header row
    HEADERS.forEach((h, j) => {
      const c = ws.getCell(4, startCol + j)
      c.value = h
      c.font = { bold: true }
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } }
      c.border = border
    })

    people.forEach((p, r) => {
      const row = 5 + r
      const values = [p.name, p.units, p.amount, region]
      values.forEach((v, j) => {
        const c = ws.getCell(row, startCol + j)
        c.value = v
        c.border = border
      })
      ws.getCell(row, startCol + 1).numFmt = '#,##0.00'
      ws.getCell(row, startCol + 2).numFmt = '$#,##0.00'
    })

    // Totals row
    const totalRow = 5 + people.length
    const colLetter = (offset: number) => ws.getColumn(startCol + offset).letter
    const range = (offset: number) => `${colLetter(offset)}5:${colLetter(offset)}${Math.max(totalRow - 1, 5)}`
    const totals = [
      'Total',
      { formula: `SUM(${range(1)})`, result: people.reduce((s, p) => s + p.units, 0) },
      { formula: `SUM(${range(2)})`, result: people.reduce((s, p) => s + p.amount, 0) },
      '',
    ]
    totals.forEach((v, j) => {
      const c = ws.getCell(totalRow, startCol + j)
      c.value = people.length ? (v as ExcelJS.CellValue) : j === 0 ? 'Total' : 0
      c.font = { bold: true }
      c.border = { ...border, top: { style: 'medium' } }
    })
    ws.getCell(totalRow, startCol + 1).numFmt = '#,##0.00'
    ws.getCell(totalRow, startCol + 2).numFmt = '$#,##0.00'

    ws.getColumn(startCol).width = 28
    ws.getColumn(startCol + 1).width = 13
    ws.getColumn(startCol + 2).width = 14
    ws.getColumn(startCol + 3).width = 12
    if (i < REGIONS.length - 1) ws.getColumn(startCol + TABLE_WIDTH).width = 4
  })

  const buf = await wb.xlsx.writeBuffer()
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const a = document.createElement('a')
  a.href = URL.createObjectURL(blob)
  a.download = `Leave-Balance-by-Region-${new Date().toISOString().slice(0, 10)}.xlsx`
  a.click()
  URL.revokeObjectURL(a.href)
}
