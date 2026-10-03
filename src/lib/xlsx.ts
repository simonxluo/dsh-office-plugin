/** Excel(.xlsx)读写:exceljs 引擎 */
import ExcelJS from 'exceljs'

export interface SheetInput {
  /** 工作表名,缺省 Sheet1/Sheet2… */
  name?: string
  /** 二维数组(首行默认表头) */
  rows?: unknown[][]
  /** 对象数组(键做表头),与 rows 二选一 */
  data?: Record<string, unknown>[]
  /** 首行是否加粗+底纹表头样式,默认 true */
  header?: boolean
  /** 列宽(字符数) */
  widths?: number[]
}

export interface CellStyle {
  bold?: boolean
  /** 填充色,ARGB 十六进制(如 FF DDEBF7 或 DDEBF7) */
  fill?: string
  /** 字体色,同上 */
  color?: string
  align?: 'left' | 'center' | 'right'
  /** 数字格式,如 '#,##0.00'、'yyyy-mm-dd' */
  format?: string
}

export type EditOp =
  | { op: 'set_value'; sheet?: string; cell: string; value?: unknown }
  | { op: 'set_style'; sheet?: string; cell?: string; range?: string; style: CellStyle }
  | { op: 'add_row'; sheet?: string; row?: unknown[] }
  | { op: 'add_sheet'; name: string; rows?: unknown[][] }
  | { op: 'delete_sheet'; name: string }

export interface ReadResult {
  sheets: string[]
  /** TSV 文本:[Sheet: 名] 分段 */
  text: string
  totalRows: number
}

function normalizeColor(c: string): string {
  return c.replace(/^#/, '').length === 6 ? 'FF' + c.replace(/^#/, '') : c.replace(/^#/, '')
}

function applyCellStyle(cell: ExcelJS.Cell, style: CellStyle): void {
  const font = {} as ExcelJS.Font
  if (style.bold) font.bold = true
  if (style.color) font.color = { argb: normalizeColor(style.color) }
  if (Object.keys(font).length > 0) cell.font = font
  if (style.fill) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: normalizeColor(style.fill) } }
  }
  if (style.align) cell.alignment = { horizontal: style.align, vertical: 'middle' }
  if (style.format) cell.numFmt = style.format
}

function objectRows(data: Record<string, unknown>[]): unknown[][] {
  const keys: string[] = []
  for (const item of data) {
    for (const k of Object.keys(item)) if (!keys.includes(k)) keys.push(k)
  }
  return [keys, ...data.map(item => keys.map(k => item[k]))]
}

function rangeCells(range: string): Array<[string, number]> | null {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range.trim().toUpperCase())
  if (!m) return null
  const colOf = (letters: string): number => {
    let n = 0
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
    return n
  }
  const c1 = colOf(m[1]); const r1 = Number(m[2]); const c2 = colOf(m[3]); const r2 = Number(m[4])
  if (c2 < c1 || r2 < r1) return null
  const out: Array<[string, number]> = []
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      let letters = ''
      let n = c
      while (n > 0) {
        const rem = (n - 1) % 26
        letters = String.fromCharCode(65 + rem) + letters
        n = Math.floor((n - 1) / 26)
      }
      out.push([letters, r])
    }
  }
  return out
}

/** 读取工作表(默认第一个),offset 从 0 开始计数据行 */
export async function readWorkbook(
  file: string,
  sheetName: string | undefined,
  offset: number,
  limit: number,
): Promise<ReadResult> {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile(file)
  const sheets = wb.worksheets.map(ws => ws.name)
  const targets = sheetName
    ? wb.worksheets.filter(ws => ws.name === sheetName)
    : wb.worksheets.slice(0, 1)
  if (targets.length === 0) throw new Error(`工作表不存在: ${sheetName}；可用: ${sheets.join(', ') || '(无)'}`)
  const parts: string[] = []
  let totalRows = 0
  for (const ws of targets) {
    parts.push(`[Sheet: ${ws.name}]`)
    const rowCount = ws.rowCount
    totalRows = Math.max(totalRows, rowCount)
    const start = Math.min(offset + 1, rowCount + 1)
    const end = Math.min(rowCount, start + Math.max(limit, 1) - 1)
    const lines: string[] = []
    for (let r = start; r <= end; r++) {
      const cells: string[] = []
      ws.getRow(r).eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cells[colNumber - 1] = cell.text ?? String(cell.value ?? '')
      })
      for (let i = 0; i < cells.length; i++) if (cells[i] === undefined) cells[i] = ''
      lines.push(cells.join('\t'))
    }
    parts.push(lines.length > 0 ? lines.join('\n') : '(该范围无数据)')
  }
  return { sheets, text: parts.join('\n'), totalRows }
}

/** 新建工作簿并写入(覆盖同路径文件) */
export async function writeWorkbook(
  file: string,
  sheets: SheetInput[],
): Promise<{ path: string; written: Array<{ name: string; rows: number }> }> {
  if (!Array.isArray(sheets) || sheets.length === 0) throw new Error('sheets 不能为空(至少一个工作表)')
  const wb = new ExcelJS.Workbook()
  const written: Array<{ name: string; rows: number }> = []
  let auto = 0
  for (const s of sheets) {
    const name = (s.name && String(s.name).trim()) || `Sheet${++auto}`
    const rows = s.rows ?? (s.data ? objectRows(s.data) : [])
    const ws = wb.addWorksheet(name)
    const useHeader = s.header !== false && rows.length > 0
    for (const r of rows) ws.addRow(r as unknown[])
    if (useHeader) {
      const headerRow = ws.getRow(1)
      headerRow.font = { bold: true }
      headerRow.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } }
        cell.alignment = { horizontal: 'center', vertical: 'middle' }
      })
      ws.views = [{ state: 'frozen', ySplit: 1 }]
    }
    if (Array.isArray(s.widths)) {
      s.widths.forEach((w, idx) => {
        ws.getColumn(idx + 1).width = w
      })
    }
    written.push({ name, rows: rows.length })
  }
  await wb.xlsx.writeFile(file)
  return { path: file, written }
}

/** 就地编辑已有工作簿 */
export async function editWorkbook(file: string, ops: EditOp[]): Promise<{ path: string; applied: number }> {
  if (!Array.isArray(ops) || ops.length === 0) throw new Error('ops 不能为空')
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile(file)
  let applied = 0
  for (const op of ops) {
    if (op.op === 'set_value' || op.op === 'set_style' || op.op === 'add_row') {
      const ws = op.sheet ? wb.getWorksheet(op.sheet) : wb.worksheets[0]
      if (!ws) throw new Error(`工作表不存在: ${op.sheet}`)
      if (op.op === 'set_value') {
        if (!op.cell) throw new Error('set_value 需要 cell(如 A1)')
        ws.getCell(op.cell).value = (op.value ?? null) as never
        applied++
      } else if (op.op === 'add_row') {
        ws.addRow((op.row ?? []) as unknown[])
        applied++
      } else {
        if (op.cell) {
          applyCellStyle(ws.getCell(op.cell), op.style ?? {})
          applied++
        } else if (op.range) {
          const cells = rangeCells(op.range)
          if (!cells) throw new Error(`range 格式应为 A1:B2,收到: ${op.range}`)
          for (const [letters, row] of cells) applyCellStyle(ws.getCell(`${letters}${row}`), op.style ?? {})
          applied++
        } else {
          throw new Error('set_style 需要 cell 或 range')
        }
      }
      continue
    }
    if (op.op === 'add_sheet') {
      if (!op.name) throw new Error('add_sheet 需要 name')
      if (wb.getWorksheet(op.name)) throw new Error(`工作表已存在: ${op.name}`)
      const ws = wb.addWorksheet(op.name)
      for (const r of op.rows ?? []) ws.addRow(r as unknown[])
      applied++
      continue
    }
    if (op.op === 'delete_sheet') {
      const ws = wb.getWorksheet(op.name)
      if (!ws) throw new Error(`工作表不存在: ${op.name}`)
      wb.removeWorksheet(ws.id)
      applied++
    }
  }
  await wb.xlsx.writeFile(file)
  return { path: file, applied }
}
