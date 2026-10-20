import { z } from 'zod'
import type { AgentToolCall, AgentToolDef } from '@araoffice/agent-core'
import { workbookOperationSchema, type WorkbookOperation } from '../../domain/workbook-dsl'
import {
  columnLabel,
  parseRange,
  rangeCellCount,
  formatAddress,
  type RangeBounds,
} from '../../domain/cell-address'
import type {
  ApplyOutcome,
  CellFormatState,
  CellScalar,
  ChangePlan,
} from '../../domain/workbook.types'
import { t } from '../i18n/locale'
import { loadGuides } from './guides'

/**
 * The workbook DSL as an AgentSkill tool set: read-only context/reader tools
 * and one propose tool. Mirrors the docx skill's read-before-write discipline
 * (get_document_context / read_blocks / replace_blocks), but the mutating
 * tool never touches the workbook directly — it only computes a ChangePlan
 * and hands it to the SAME plan/apply path the manual flow uses, which now
 * auto-applies immediately (undo via ⌘Z / inline button covers everything;
 * the preview card only remains as a manual fallback when apply fails).
 */

/** raw shape the model sends for one operation; validated against workbookOperationSchema */
export type ProposedOperation = Record<string, unknown>

export interface SheetRef {
  readonly id: string
  readonly name: string
  /** Data extent (from the xlsx dimension or known cells); may drift slightly
   * after structural changes within the session */
  readonly rows?: number
  readonly columns?: number
}

export interface ChartRef {
  readonly path: string
  readonly title: string
  readonly types: string
  readonly sheetId: string
}

export interface ActiveSheetInfo {
  readonly mode: 'demo' | 'lazy' | 'none'
  readonly sheetId: string
  readonly sheetName: string
  /** demo mode only: current revision, needed for the CAS-checked plan() call */
  readonly revision?: number
  /** non-empty cell addresses known to the caller without an extra read */
  readonly knownAddresses: readonly string[]
  /** lazy mode only: the viewport-backed range currently present in Univer */
  readonly loadedRange?: string | undefined
  /** every sheet in the workbook, active one included */
  readonly sheets: readonly SheetRef[]
  /** current selection in A1 notation, when one exists */
  readonly selection?: string | undefined
  /** merged ranges on the active sheet (A1 notation) */
  readonly merges?: readonly string[] | undefined
  /** charts in the workbook (imported files only) */
  readonly charts?: readonly ChartRef[] | undefined
}

function isBlankSheet(info: ActiveSheetInfo, active: SheetRef | undefined): boolean {
  return (
    active?.rows === 1 &&
    active.columns === 1 &&
    info.knownAddresses.length === 0 &&
    (info.merges?.length ?? 0) === 0 &&
    (info.charts?.length ?? 0) === 0
  )
}

export interface SheetsSkillDeps {
  getActiveSheetInfo(): ActiveSheetInfo
  /** Ensure a lazy workbook range is present in Univer before reading it. */
  ensureRangeLoaded?(range: RangeBounds): boolean | Promise<boolean>
  readCells(addresses: readonly string[]): Record<string, { value: CellScalar; formula?: string }>
  /** per-cell explicit formatting of the active sheet; cells with no explicit format are omitted */
  readFormats(addresses: readonly string[]): Record<string, CellFormatState>
  /** formatted report of a sheet's feature state (filters, CF, DV, names, visuals, …) */
  readSheetFeatures(sheetId?: string): string
  /** `applied` resolves with the real apply result (the lazy path applies async);
   * the tool awaits it so the model never hears "applied" for a batch that failed */
  proposeOperations(
    operations: readonly WorkbookOperation[],
    summary: string,
  ): { ok: true; plan: ChangePlan; applied?: Promise<ApplyOutcome> } | { ok: false; error: string }
}

const MAX_READ_ADDRESSES = 100
const MAX_READ_RANGE_CELLS = 2000
/** Read-back after write: max number of formula cells whose results are read back */
const MAX_READBACK_FORMULAS = 10
/** Read-back after write: wait time (ms) for Univer's async formula recalc */
const FORMULA_RECALC_DELAY_MS = 300
const MAX_READ_FORMAT_CELLS = 200

export const SHEETS_TOOLS: AgentToolDef[] = [
  {
    name: 'get_workbook_context',
    description:
      'Get a workbook overview: all sheets (id/name/data-extent rows-columns), active sheet, current selection, known non-empty cell addresses. ' +
      'Note: Initial workbook context is already provided in the prompt context. Only call this tool if workbook structure changed during the conversation. ' +
      'For data-size questions (how many rows / how much data), answer from the data extent here instead of reading block by block; use read_range or read_cells when concrete values are needed.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'read_range',
    description:
      'Read current values/formulas by rectangular range, returning a grid with row numbers and column letters. ' +
      'The requested range is not the worksheet data extent: never infer total row or record count from its ending row; use get_workbook_context. ' +
      'This is the preferred way to read data; max 2000 cells — read larger regions in multiple calls.',
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Range like "A1:D20"; a single cell like "B2" is also accepted',
        },
      },
      required: ['range'],
    },
  },
  {
    name: 'read_formats',
    description:
      'Read explicit cell formats in a range (bold/italic/underline/colors/number format/alignment/borders); only formatted cells are returned. ' +
      'Use when you need to "reuse the format from somewhere" or inspect current formatting; max 200 cells.',
    inputSchema: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'Range like "A1:D20"' },
      },
      required: ['range'],
    },
  },
  {
    name: 'read_sheet_features',
    description:
      "Read a worksheet's feature state: AutoFilter (range and column criteria), conditional formatting rules, data validation rules, defined names, " +
      'freeze panes, hidden/protected status, shapes and images, and page setup pending save this session. ' +
      'Read the current state before modifying or clearing any of these existing settings — never change them blindly.',
    inputSchema: {
      type: 'object',
      properties: {
        sheetId: {
          type: 'string',
          description: 'Target sheet id; reads the active sheet when omitted',
        },
      },
      required: [],
    },
  },
  {
    name: 'read_cells',
    description:
      'Read current values/formulas of specific scattered cells (use read_range for contiguous regions). Always read the affected cells before writing — never assume their contents.',
    inputSchema: {
      type: 'object',
      properties: {
        addresses: {
          type: 'array',
          items: { type: 'string' },
          description: 'List of cell addresses, e.g. ["A1","B2"], max 100',
        },
      },
      required: ['addresses'],
    },
  },
  {
    name: 'propose_operations',
    description:
      'Submit a batch of change operations, applied to the workbook immediately (the user can roll back with the [Undo] button or ⌘Z). Supported operations:\n' +
      '- set_range: {op:"set_range",sheetId,range:"A1:D5",values:[[row1],[row2],...]} (formulas start with =; null clears cell)\n' +
      '- set_cell: {op:"set_cell",sheetId,address:"B2",value:"Text"}\n' +
      '- set_formula: {op:"set_formula",sheetId,address:"E2",formula:"=AVERAGE(C2:D2)"}\n' +
      '- format_range: {op:"format_range",sheetId,range:"A1:D1",format:{bold?:bool,fillColor?:"#4472C4",fontColor?:"#FFFFFF",fontSize?:num,horizontalAlign?:"center",verticalAlign?:"middle",wrapText?:bool,numberFormat?:"0.00%"}}\n' +
      '- add_table: {op:"add_table",sheetId,range:"A1:D10",name?:"TableName",style?:"TableStyleMedium2"}\n' +
      '- set_col_width: {op:"set_col_width",sheetId,cols:[{col:0,width:140}]}\n' +
      '- set_row_height: {op:"set_row_height",sheetId,rows:[{row:1,height:28}]}\n' +
      '- merge_cells: {op:"merge_cells",sheetId,range:"A1:D1"} / unmerge_cells\n' +
      '- sort_range: {op:"sort_range",sheetId,range:"A2:D20",col:2,asc?:bool}\n' +
      '- set_freeze: {op:"set_freeze",sheetId,row?:1,col?:1}\n' +
      '- clear_range: {op:"clear_range",sheetId,range:"A1:D10"} / clear_cell: {op:"clear_cell",sheetId,address:"A1"}\n' +
      '- rename_sheet: {op:"rename_sheet",sheetId,name:"NewName"}\n' +
      '- add_chart: {op:"add_chart",sheetId,type:"column"|"bar"|"line"|"pie"|"area",range:"A1:D10",title?:"Title"}\n' +
      '- insert_rows: {op:"insert_rows",sheetId,row:1,count:5} / delete_rows / insert_cols / delete_cols\n' +
      'Limits: structural ops (insert/delete rows/cols, sheet ops) cannot share a batch with content/format ops; max 2000 cell changes; sheetId must come from prompt context or get_workbook_context.',
    inputSchema: {
      type: 'object',
      properties: {
        operations: {
          type: 'array',
          items: { type: 'object' },
          description: 'Array of operations in the workbook DSL discriminated-union format',
        },
        summary: {
          type: 'string',
          description:
            'One-sentence summary of this batch of changes. Always include a real summary ' +
            '(e.g. "Add seven-column student table"); never an empty string.',
        },
      },
      required: ['operations', 'summary'],
    },
  },
]

export interface ToolExecution {
  output: string
  isError?: boolean
  /** true when propose_operations auto-applied a batch of changes */
  mutated: boolean
  summary: string
}

const fail = (summary: string, output: string): ToolExecution => ({
  output,
  isError: true,
  mutated: false,
  summary,
})

export function buildSheetsContext(deps: SheetsSkillDeps): string {
  const info = deps.getActiveSheetInfo()
  if (info.mode === 'none') return 'No workbook is currently open.'
  const dims = (sheet: SheetRef): string =>
    sheet.rows && sheet.columns
      ? `, data extent about ${sheet.rows} rows × ${sheet.columns} columns`
      : ''
  const active = info.sheets.find((sheet) => sheet.id === info.sheetId)
  const lines = [
    `Active sheet: ${info.sheetName} (id=${info.sheetId}${active ? dims(active) : ''})`,
    info.mode === 'demo'
      ? `Mode: demo workbook, current revision=${info.revision}`
      : 'Mode: imported real xlsx file (some regions may still be streaming in)',
  ]
  if (active?.rows && active.columns) {
    lines.push(
      `Active sheet data area: A1:${columnLabel(active.columns - 1)}${active.rows}` +
        ' (answer data-size questions directly from this — do not tally block by block with read_range)',
    )
  }
  if (info.sheets.length > 1) {
    lines.push(
      `All sheets: ${info.sheets.map((sheet) => `${sheet.name} (id=${sheet.id}${dims(sheet)})`).join(', ')}`,
    )
  }
  if (info.selection) {
    lines.push(`Current selection: ${info.selection}`)
  }
  if (info.loadedRange) {
    lines.push(`Currently loaded viewport: ${info.loadedRange} (not the worksheet data extent)`)
  }
  if (info.merges && info.merges.length > 0) {
    lines.push(`Merged ranges on the active sheet: ${info.merges.slice(0, 50).join(', ')}`)
  }
  if (info.charts && info.charts.length > 0) {
    lines.push(
      'Charts in the workbook (use the path below with edit_chart to edit an existing chart; use add_chart to create one):',
    )
    for (const chart of info.charts) {
      lines.push(
        `- ${chart.path} | title: ${chart.title || '(none)'} | type: ${chart.types} | sheetId: ${chart.sheetId}`,
      )
    }
  }
  if (info.knownAddresses.length > 0) {
    lines.push(
      `Known non-empty cells (may be incomplete): ${info.knownAddresses.slice(0, 200).join(', ')}`,
    )
  } else if (isBlankSheet(info, active)) {
    lines.push(
      'The active sheet is blank. A1:A1 is only the current data extent; template ranges may be written directly without reading beyond A1.',
    )
    lines.push(
      'Do not retry a failed merge blindly: reuse existing merges from this context or unmerge an overlapping range first.',
    )
  } else {
    lines.push(
      'No known non-empty cell information yet; read on demand with read_range/read_cells.',
    )
  }
  return lines.join('\n')
}

function describeFormatState(format: CellFormatState): string {
  const parts: string[] = []
  if (format.bold) parts.push('bold')
  if (format.italic) parts.push('italic')
  if (format.underline) parts.push('underline')
  if (format.strikethrough) parts.push('strikethrough')
  if (format.fontFamily) parts.push(`font ${format.fontFamily}`)
  if (format.fontSize) parts.push(`size ${format.fontSize}`)
  if (format.fontColor) parts.push(`font color ${format.fontColor}`)
  if (format.fillColor) parts.push(`fill ${format.fillColor}`)
  if (format.numberFormat) parts.push(`number format ${format.numberFormat}`)
  if (format.horizontalAlign) parts.push(`align ${format.horizontalAlign}`)
  if (format.verticalAlign) parts.push(`valign ${format.verticalAlign}`)
  if (format.wrapText) parts.push('wrap')
  if (format.border) {
    parts.push(
      `border ${format.border.type}${format.border.color ? ` ${format.border.color}` : ''}`,
    )
  }
  return parts.join(', ') || '(none)'
}

// Cell text is emitted into tab/newline-delimited tool output (read_range grid,
// read_cells lists, plan summaries), where raw control characters would tear the
// line/column structure apart and scramble the model's view of the grid. Escape
// them — and backslash itself, so the encoding stays unambiguous. Univer streams
// in-cell paragraph breaks as \r (the file model uses \n; see edit-journal's
// dataStream conversion), so CR/CRLF are normalized to \n first: the model sees
// a single line-break representation, and echoing the same `\n` escape inside
// JSON string values of write operations round-trips into real line breaks.
function escapeCellText(text: string): string {
  return text
    .replace(/\\/g, '\\\\')
    .replace(/\t/g, '\\t')
    .replace(/\r\n?/g, '\n')
    .replace(/\n/g, '\\n')
}

function formatCellScalar(cell: { value: CellScalar; formula?: string | undefined }): string {
  // Formula cells prefer the value (the AI reasons from computed values, with
  // the formula as provenance); when the value isn't computed yet, give only the
  // formula
  if (cell.formula) {
    return cell.value === null || cell.value === undefined
      ? escapeCellText(cell.formula)
      : `${escapeCellText(String(cell.value))} (${escapeCellText(cell.formula)})`
  }
  if (cell.value === null) return '(empty)'
  return escapeCellText(String(cell.value))
}

function formatPlanSummary(plan: ChangePlan): string {
  const parts: string[] = []
  if (plan.structuralChanges.length > 0) {
    parts.push(plan.structuralChanges.map((change) => change.label).join('; '))
  }
  if (plan.formatChanges.length > 0) {
    parts.push(plan.formatChanges.map((change) => change.label).join('; '))
  }
  if (plan.cellChanges.length > 0) {
    const shown = plan.cellChanges.slice(0, 20)
    const rest = plan.cellChanges.length - shown.length
    parts.push(
      shown
        .map((c) => `${c.address}: ${formatCellScalar(c.before)} → ${formatCellScalar(c.after)}`)
        .join('; ') + (rest > 0 ? `; …${rest} more cells` : ''),
    )
  }
  if (plan.sheetRenames.length > 0) {
    parts.push(plan.sheetRenames.map((r) => `sheet ${r.before} → ${r.after}`).join('; '))
  }
  return parts.join(' | ') || '(no changes)'
}

function describeOperationsFallback(rawOps: unknown): string {
  if (!Array.isArray(rawOps) || rawOps.length === 0) return 'Apply workbook changes'
  const counts = new Map<string, number>()
  for (const op of rawOps) {
    const name =
      op !== null && typeof op === 'object' && 'op' in op && typeof op.op === 'string'
        ? String(op.op).replace(/[_-]+/g, ' ')
        : 'operation'
    counts.set(name, (counts.get(name) ?? 0) + 1)
  }
  const parts = [...counts.entries()].map(([name, count]) =>
    count > 1 ? `${count}× ${name}` : name,
  )
  return `Apply workbook changes (${parts.slice(0, 4).join(', ')})`
}

function normalizeFormatPatch(format: unknown): ProposedOperation {
  if (!format || typeof format !== 'object') return {}
  const f: ProposedOperation = { ...(format as ProposedOperation) }
  if (f.bgColor && !f.fillColor) {
    f.fillColor = f.bgColor
    delete f.bgColor
  }
  if (f.backgroundColor && !f.fillColor) {
    f.fillColor = f.backgroundColor
    delete f.backgroundColor
  }
  if (f.color && !f.fontColor) {
    f.fontColor = f.color
    delete f.color
  }
  if (f.textColor && !f.fontColor) {
    f.fontColor = f.textColor
    delete f.textColor
  }
  if (f.align && !f.horizontalAlign) {
    f.horizontalAlign = f.align
    delete f.align
  }
  if (f.alignment && !f.horizontalAlign) {
    f.horizontalAlign = f.alignment
    delete f.alignment
  }
  if (f.valign && !f.verticalAlign) {
    f.verticalAlign = f.valign
    delete f.valign
  }
  if (f.numFmt && !f.numberFormat) {
    f.numberFormat = f.numFmt
    delete f.numFmt
  }
  if (f.border && typeof f.border !== 'object') {
    f.border = { type: 'all' }
  }
  return f
}

function normalizeRawOperations(rawOps: readonly unknown[]): ProposedOperation[] {
  const result: ProposedOperation[] = []
  for (const item of rawOps) {
    if (!item || typeof item !== 'object') {
      continue
    }
    const obj: ProposedOperation = { ...(item as ProposedOperation) }
    if (!('op' in obj) && 'type' in obj && typeof obj.type === 'string') {
      obj.op = obj.type === 'write_range' ? 'set_range' : obj.type
    }
    if (!('sheetId' in obj) && 'sheet' in obj) {
      obj.sheetId = obj.sheet
    }

    if (obj.op === 'set_col_width') {
      const cols = Array.isArray(obj.cols)
        ? obj.cols
        : Array.isArray(obj.columns)
          ? obj.columns
          : null
      if (cols) {
        for (const c of cols) {
          if (c && typeof c === 'object') {
            const co = c as ProposedOperation
            const colIndex =
              typeof co.col === 'number' ? co.col : typeof co.column === 'number' ? co.column : null
            const colLabel =
              colIndex === null ? String(co.col ?? co.column ?? 'A') : columnLabel(colIndex)
            const rawW = Number(co.width ?? co.widthPx ?? 80)
            const widthPx = rawW <= 50 ? Math.max(10, Math.round(rawW * 7)) : rawW
            result.push({
              op: 'set_col_width',
              sheetId: obj.sheetId,
              column: colLabel,
              count: 1,
              widthPx,
            })
          }
        }
        continue
      }
      if (typeof obj.column === 'number') {
        obj.column = columnLabel(obj.column)
      }
      if (obj.width !== undefined && obj.widthPx === undefined) {
        const rawW = Number(obj.width)
        obj.widthPx = rawW <= 50 ? Math.max(10, Math.round(rawW * 7)) : rawW
        delete obj.width
      }
    } else if (obj.op === 'set_row_height') {
      const rows = Array.isArray(obj.rows) ? obj.rows : null
      if (rows) {
        for (const r of rows) {
          if (r && typeof r === 'object') {
            const ro = r as ProposedOperation
            result.push({
              op: 'set_row_height',
              sheetId: obj.sheetId,
              row: Number(ro.row),
              count: 1,
              heightPoints: Number(ro.height ?? ro.heightPoints ?? 20),
            })
          }
        }
        continue
      }
      if (obj.height !== undefined && obj.heightPoints === undefined) {
        obj.heightPoints = Number(obj.height)
        delete obj.height
      }
    } else if (obj.op === 'format_range' || obj.op === 'set_format') {
      obj.op = 'format_range'
      if (obj.format) obj.format = normalizeFormatPatch(obj.format)
    }
    result.push(obj)
  }
  return result
}

export function executeSheetsTool(
  deps: SheetsSkillDeps,
  call: AgentToolCall,
): ToolExecution | Promise<ToolExecution> {
  switch (call.name) {
    case 'get_workbook_context':
      return {
        output: buildSheetsContext(deps),
        mutated: false,
        summary: t('aiToolWorkbookContext'),
      }

    case 'read_range': {
      const raw = call.input.range
      if (typeof raw !== 'string' || !raw.trim())
        return fail(t('aiToolReadRange'), 'range must be a non-empty string')
      let bounds
      try {
        bounds = parseRange(raw.trim().toUpperCase())
      } catch {
        return fail(t('aiToolReadRange'), `Cannot parse range: ${raw}`)
      }
      if (rangeCellCount(bounds) > MAX_READ_RANGE_CELLS) {
        return fail(
          t('aiToolReadRange'),
          `The range contains more than ${MAX_READ_RANGE_CELLS} cells; read it in multiple calls`,
        )
      }
      const info = deps.getActiveSheetInfo()
      const active = info.sheets.find((sheet) => sheet.id === info.sheetId)
      const blankSheet = isBlankSheet(info, active)
      if (
        !blankSheet &&
        active?.rows !== undefined &&
        active.columns !== undefined &&
        (bounds.endRow >= active.rows || bounds.endColumn >= active.columns)
      ) {
        return fail(
          t('aiToolReadRange'),
          `The requested range is outside the worksheet data extent A1:${columnLabel(active.columns - 1)}${active.rows}.`,
        )
      }
      const outsideDataExtent =
        active?.rows !== undefined &&
        active.columns !== undefined &&
        (bounds.endRow >= active.rows || bounds.endColumn >= active.columns)
      const executeRead = (): ToolExecution => {
        const normalizedRange = `${formatAddress(bounds.startRow, bounds.startColumn)}:${formatAddress(bounds.endRow, bounds.endColumn)}`
        const metadata =
          active?.rows && active.columns
            ? `Read metadata: requested range ${normalizedRange}; authoritative worksheet data extent A1:${columnLabel(active.columns - 1)}${active.rows} (${active.rows} worksheet rows including any header). Do not infer total rows or records from the requested range.` +
              (blankSheet && outsideDataExtent
                ? ' The sheet is blank, so cells outside that extent are treated as empty.'
                : '')
            : `Read metadata: requested range ${normalizedRange}; worksheet data extent is unknown. Do not infer total rows or records from the requested range.`
        const addresses: string[] = []
        for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
          for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
            addresses.push(formatAddress(row, column))
          }
        }
        const cells = deps.readCells(addresses)
        const header = [
          '',
          ...Array.from({ length: bounds.endColumn - bounds.startColumn + 1 }, (_, offset) =>
            columnLabel(bounds.startColumn + offset),
          ),
        ].join('\t')
        const rows: string[] = [metadata, header]
        for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
          const columns: string[] = [String(row + 1)]
          for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
            const cell = cells[formatAddress(row, column)]
            columns.push(
              cell
                ? cell.value === null
                  ? escapeCellText(cell.formula ?? '')
                  : formatCellScalar(cell)
                : '',
            )
          }
          rows.push(columns.join('\t'))
        }
        return {
          output: rows.join('\n'),
          mutated: false,
          summary: t('aiToolReadRangeOf', { range: raw.trim().toUpperCase() }),
        }
      }
      const loading = blankSheet && outsideDataExtent ? true : deps.ensureRangeLoaded?.(bounds)
      if (loading instanceof Promise) {
        return loading.then((loaded) =>
          loaded
            ? executeRead()
            : fail(
                t('aiToolReadRange'),
                'The requested range could not be fully loaded; retry after workbook indexing completes.',
              ),
        )
      }
      if (loading === false) {
        return fail(
          t('aiToolReadRange'),
          'The requested range could not be fully loaded; retry after workbook indexing completes.',
        )
      }
      return executeRead()
    }

    case 'load_guide': {
      const raw = call.input.guides
      if (!Array.isArray(raw) || raw.length === 0)
        return fail(t('aiToolLoadGuide'), 'guides must be a non-empty array')
      const outcome = loadGuides(raw.map(String))
      if (!outcome.ok) return fail(t('aiToolLoadGuide'), outcome.error)
      return {
        output: outcome.content,
        mutated: false,
        summary: t('aiToolLoadGuideOf', { names: raw.join(', ') }),
      }
    }

    case 'read_formats': {
      const raw = call.input.range
      if (typeof raw !== 'string' || !raw.trim())
        return fail(t('aiToolReadFormats'), 'range must be a non-empty string')
      let bounds
      try {
        bounds = parseRange(raw.trim().toUpperCase())
      } catch {
        return fail(t('aiToolReadFormats'), `Cannot parse range: ${raw}`)
      }
      if (rangeCellCount(bounds) > MAX_READ_FORMAT_CELLS) {
        return fail(
          t('aiToolReadFormats'),
          `The range contains more than ${MAX_READ_FORMAT_CELLS} cells; read it in multiple calls`,
        )
      }
      const addresses: string[] = []
      for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
        for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
          addresses.push(formatAddress(row, column))
        }
      }
      const formats = deps.readFormats(addresses)
      const lines = Object.entries(formats).map(
        ([address, format]) => `${address}: ${describeFormatState(format)}`,
      )
      return {
        output: lines.length > 0 ? lines.join('\n') : 'No explicit formats in this range.',
        mutated: false,
        summary: t('aiToolReadFormatsOf', { range: raw.trim().toUpperCase() }),
      }
    }

    case 'read_sheet_features': {
      const raw = call.input.sheetId
      const sheetId = typeof raw === 'string' && raw.trim() ? raw.trim() : undefined
      return {
        output: deps.readSheetFeatures(sheetId),
        mutated: false,
        summary: t('aiToolSheetFeatures'),
      }
    }

    case 'read_cells': {
      const raw = call.input.addresses
      if (!Array.isArray(raw) || raw.length === 0)
        return fail(t('aiToolReadCells'), 'addresses must be a non-empty array')
      const addresses = raw.slice(0, MAX_READ_ADDRESSES).map(String)
      const cells = deps.readCells(addresses)
      const lines = addresses.map((addr) => {
        const cell = cells[addr]
        return `${addr}: ${cell ? formatCellScalar(cell) : '(unknown)'}`
      })
      return {
        output: lines.join('\n'),
        mutated: false,
        summary: t('aiToolReadCellsCount', { count: addresses.length }),
      }
    }

    case 'propose_operations': {
      let rawOps = call.input.operations
      const summaryInput = call.input.summary
      if (typeof rawOps === 'string') {
        try {
          rawOps = JSON.parse(rawOps)
        } catch {
          // retain raw value for schema validation error
        }
      }
      if (!Array.isArray(rawOps) || rawOps.length === 0) {
        return fail(t('aiToolPropose'), 'operations must be a non-empty array')
      }
      rawOps = normalizeRawOperations(rawOps)
      // The model sometimes emits summary:"" or Chinese text on non-Chinese prompts:
      // derive a fallback from the operations instead of showing raw Chinese.
      const rawSummary = typeof summaryInput === 'string' ? summaryInput.trim() : ''
      const hasChinese = /[\u4e00-\u9fa5]/.test(rawSummary)
      const summaryText =
        rawSummary && !hasChinese ? rawSummary : describeOperationsFallback(rawOps)
      let operations: WorkbookOperation[]
      try {
        operations = z.array(workbookOperationSchema).parse(rawOps)
      } catch (e) {
        return fail(
          summaryText || t('aiToolPropose'),
          e instanceof Error ? e.message : 'Invalid operation format',
        )
      }
      const outcome = deps.proposeOperations(operations, summaryText)
      if (!outcome.ok) return fail(summaryText || t('aiToolPropose'), outcome.error)
      const summary = summaryText
      const finish = (): ToolExecution | Promise<ToolExecution> => {
        const warnings =
          outcome.plan.warnings.length > 0 ? `\nNote: ${outcome.plan.warnings.join('; ')}` : ''
        const opCount =
          outcome.plan.cellChanges.length +
          outcome.plan.formatChanges.length +
          outcome.plan.sheetRenames.length +
          outcome.plan.structuralChanges.length
        const base = `Auto-applied ${opCount} change(s) (undo via the side panel [Undo] button or ⌘Z): ${formatPlanSummary(outcome.plan)}${warnings}`
        // Read-back after write (write → verify): formula cells fetch their
        // computed values after the async recalc, so the AI sees real results and
        // errors like #REF!/#DIV/0! instead of just what it wrote.
        const formulaAddrs = outcome.plan.cellChanges.flatMap((c) =>
          c.after.formula ? [c.address] : [],
        )
        if (formulaAddrs.length === 0) {
          return { output: base, mutated: true, summary }
        }
        return (async (): Promise<ToolExecution> => {
          await new Promise((resolve) => setTimeout(resolve, FORMULA_RECALC_DELAY_MS))
          const shown = formulaAddrs.slice(0, MAX_READBACK_FORMULAS)
          const cells = deps.readCells(shown)
          const lines = shown.map((addr) => {
            const v = cells[addr]?.value
            return `${addr} = ${v === null || v === undefined ? '(still computing; verify with read_cells)' : String(v)}`
          })
          const rest = formulaAddrs.length - shown.length
          const hasError = lines.some((l) =>
            /#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!)/.test(l),
          )
          return {
            output:
              `${base}\nFormula results: ${lines.join('; ')}${rest > 0 ? `; …${rest} more formula cells` : ''}` +
              (hasError
                ? '\n⚠️ Formula error values present — check references/divisors and fix them.'
                : ''),
            mutated: true,
            summary,
          }
        })()
      }
      if (!outcome.applied) return finish()
      return outcome.applied.then((applied) =>
        applied.ok
          ? finish()
          : fail(
              t('aiToolPropose'),
              `Apply failed — the workbook is UNCHANGED: ${applied.reason ?? 'unknown reason'}. ` +
                'Do not tell the user the changes were made; adjust the operations and retry, or explain the failure.',
            ),
      )
    }

    default:
      return fail(call.name, `Unknown tool: ${call.name}`)
  }
}
/** Legacy names kept for tests: prefer SHEETS_TOOLS / buildSheetsContext / executeSheetsTool(deps, call). */
export const WORKBOOK_TOOLS = SHEETS_TOOLS
export const buildWorkbookContext = buildSheetsContext
export function executeWorkbookTool(
  call: AgentToolCall,
  deps: SheetsSkillDeps,
): ToolExecution | Promise<ToolExecution> {
  return executeSheetsTool(deps, call)
}
