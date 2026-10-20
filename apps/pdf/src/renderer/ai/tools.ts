import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { AgentToolCall, AgentToolDef, ToolExecution } from '@araoffice/agent-core'
import type { OutlineNode } from '../OutlinePanel'
import type { SearchIndex } from '../search'
import { searchInIndex } from '../search'
import type { FormValueInput, MarkupType } from '../../shared/ipc'
import { t } from '../i18n/locale'

/** Text cap per read_pages fed back to the model (the payload is resent in full each turn, so volume must be limited) */
const READ_CHUNK_CHARS = 24_000
const MAX_MATCHES = 40

/** Highlighting a whole document can bury the model in annotations; match count is the guard */
const MAX_MARKUPS = 50

/** pdfjs AnnotationType.WIDGET — form fields render as widget annotations */
const WIDGET = 20

interface WidgetAnnotation {
  fieldName?: string
  fieldValue?: unknown
  checkBox?: boolean
  radioButton?: boolean
  options?: { exportValue?: unknown }[]
  readOnly?: boolean
}

/** Capability surface App provides to AI tools; all getters, since the loop outlives render closures */
export interface PdfSkillDeps {
  doc(): PDFDocumentProxy | null
  fileName(): string
  pageCount(): number
  /** Original page number of the currently visible page (1-based) */
  currentPage(): number
  readOnly(): boolean
  outline(): OutlineNode[] | null
  searchIndex(): Promise<SearchIndex> | null
  isDeleted(origIdx: number): boolean
  /** Original page number → scroll to that page; returns false if the page was deleted */
  gotoPage(origPage: number): boolean
  addMarkup(type: MarkupType, origIdx: number, rects: [number, number, number, number][]): void
  formEdits(): ReadonlyMap<string, FormValueInput>
  applyFormEdit(v: FormValueInput): void
  rotatePage(origIdx: number, dir: 90 | -90): void
  deletePage(origIdx: number): boolean
}

export const PDF_TOOLS: AgentToolDef[] = [
  {
    name: 'read_pages',
    description:
      'Read the text content of a page range (with [Page N] markers). Read the relevant pages before answering questions; at most 10 pages per call, over-long output is truncated.',
    inputSchema: {
      type: 'object',
      properties: {
        start: { type: 'integer', description: 'Start page number (1-based)' },
        end: {
          type: 'integer',
          description: 'End page number (inclusive); if omitted, only the start page is read',
        },
      },
      required: ['start'],
    },
  },
  {
    name: 'search_text',
    description:
      'Search the full text for a string; returns the page number and a context excerpt for each hit. Prefer this when locating which page something is on.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Text to search for (case-insensitive)' },
      },
      required: ['query'],
    },
  },
  {
    name: 'goto_page',
    description: 'Scroll the reading view to the given page so the user can see it.',
    inputSchema: {
      type: 'object',
      properties: { page: { type: 'integer', description: 'Page number (1-based)' } },
      required: ['page'],
    },
  },
  {
    name: 'get_outline',
    description:
      'Read the document outline (bookmarks) tree, including entry titles. Returns empty if the document has no outline.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'list_form_fields',
    description:
      'List the form fields with their kind and current value. Call this before fill_form_field to discover names and kinds.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'fill_form_field',
    description:
      'Fill one form field by name. For checkboxes pass "true" or "false"; for radio/choice fields the option value must match. Field lookups come from list_form_fields.',
    inputSchema: {
      type: 'object',
      properties: {
        name: { type: 'string', description: 'Field name from list_form_fields' },
        value: { type: 'string', description: 'New value; checkbox accepts "true"/"false"' },
      },
      required: ['name', 'value'],
    },
  },
  {
    name: 'rotate_page',
    description:
      'Rotate one page a quarter turn. The change shows in the view immediately and is undoable by the user.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        direction: {
          type: 'string',
          enum: ['cw', 'ccw'],
          description: 'Direction; clockwise by default',
        },
      },
      required: ['page'],
    },
  },
  {
    name: 'delete_page',
    description:
      'Remove one page from the document. Destructive but undoable; confirm with the user when the request is ambiguous.',
    inputSchema: {
      type: 'object',
      properties: { page: { type: 'integer', description: 'Page number (1-based)' } },
      required: ['page'],
    },
  },
  {
    name: 'markup_text',
    description:
      'Highlight, underline or strike out every occurrence of a query, appended as real annotations (undoable). Prefer searching first; cap of 50 matches per call.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Text to mark' },
        type: {
          type: 'string',
          enum: ['highlight', 'underline', 'strikeout', 'squiggly'],
          description: 'Markup kind; highlight by default',
        },
        page: {
          type: 'integer',
          description: 'Restrict to one page (1-based); omitted marks every page',
        },
      },
      required: ['query'],
    },
  },
]

function err(output: string, summary: string): ToolExecution {
  return { output, isError: true, summary }
}

/** Widget annotations from every page, for the form tools */
interface FieldEntry {
  name: string
  kind: FormValueInput['kind']
  widget: WidgetAnnotation
  /** 1-based orig page number the field's widget sits on */
  origPage: number
}

const kindOf = (w: WidgetAnnotation): FormValueInput['kind'] =>
  w.radioButton ? 'radio' : w.checkBox ? 'checkbox' : Array.isArray(w.options) ? 'choice' : 'text'

async function collectFields(deps: PdfSkillDeps): Promise<FieldEntry[]> {
  const d = deps.doc()
  if (!d) return []
  const out: FieldEntry[] = []
  for (let n = 1; n <= d.numPages; n++) {
    const page = await d.getPage(n)
    const annotations = await page.getAnnotations()
    // SAFETY: pdfjs ships widget fields (fieldName/fieldValue/checkBox/radioButton/options) untyped
    // in its annotation data; the doc is ours to walk, so the fields convey invariant meaning
    const annotations_ = annotations as unknown as (WidgetAnnotation & {
      annotationType?: number
    })[]
    for (const a of annotations_) {
      if (a.annotationType !== WIDGET || typeof a.fieldName !== 'string' || !a.fieldName) continue
      out.push({ name: a.fieldName, kind: kindOf(a), widget: a, origPage: n })
    }
    page.cleanup()
  }
  return out
}

async function listFormFields(deps: PdfSkillDeps): Promise<ToolExecution> {
  const fields = await collectFields(deps)
  const lines = fields.map((f) => {
    const value = f.widget.fieldValue
    const shown = f.kind === 'checkbox' ? String(value === true) : String(value ?? '')
    return `- ${f.name} (${f.kind}) = "${shown}"${f.widget.readOnly ? ' [read-only]' : ''}`
  })
  return {
    output: lines.join('\n') || 'The document has no form fields',
    summary: t('aiToolFields', { count: fields.length }),
  }
}

const CHECKBOX_ALIASES = new Set(['true', 'yes', 'checked', 'an', '1', '✓'])

async function fillFormField(
  deps: PdfSkillDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const name = String(input.name ?? '').trim()
  const raw = input.value === undefined ? '' : String(input.value)
  const summary = t('aiToolFill', { name })
  if (!name) return err('name must not be empty', summary)

  const fields = await collectFields(deps)
  const field = fields.find((f) => f.name === name)
  if (!field) {
    const known = fields.map((f) => f.name).join(', ')
    return err(
      known
        ? `Unknown field "${name}" (fields: ${known})`
        : `Unknown field "${name}" (the document has no form fields)`,
      summary,
    )
  }
  if (field.widget.readOnly) return err(`Field "${name}" is read-only`, summary)

  const kind = field.kind
  if (kind === 'checkbox') {
    const checked = CHECKBOX_ALIASES.has(raw.toLowerCase())
    deps.applyFormEdit({ name, kind, checked })
    return { output: `Checked "${name}"`, summary }
  }
  if (raw === '') return err(`Empty value for "${name}"`, summary)
  if (kind === 'radio' || kind === 'choice') {
    const options = (field.widget.options ?? []).map((o) => String(o.exportValue))
    if (options.length > 0 && !options.includes(raw)) {
      return err(
        `"${raw}" is not one of the options for "${name}" (options: ${options.join(', ')})`,
        summary,
      )
    }
    deps.applyFormEdit({ name, kind, value: raw })
    return { output: `Set "${name}" to "${raw}"`, summary }
  }
  deps.applyFormEdit({ name, kind: 'text', value: raw })
  return { output: `Set "${name}" to "${raw}"`, summary }
}

/** Validate a 1-based page number; returns the original page index or an error */
function resolvePage(deps: PdfSkillDeps, raw: unknown): { origIdx: number } | { bad: string } {
  const page = Number(raw)
  if (!Number.isInteger(page) || page < 1 || page > deps.pageCount()) {
    return {
      bad: `Page number ${String(raw)} is out of range (document has ${deps.pageCount()} pages)`,
    }
  }
  if (deps.isDeleted(page - 1)) return { bad: `Page ${page} has been deleted (unsaved)` }
  return { origIdx: page - 1 }
}

async function readPages(
  deps: PdfSkillDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const doc = deps.doc()
  if (!doc) return err('Document not ready', t('aiToolReadPages', { start: '?', end: '?' }))
  const start = Number(input.start)
  const end = Math.min(Number(input.end ?? start), start + 9)
  const summary = t('aiToolReadPages', { start, end })
  if (!Number.isInteger(start) || start < 1 || end < start || start > doc.numPages) {
    return err(`Invalid page range (document has ${doc.numPages} pages)`, summary)
  }
  let out = ''
  for (let n = start; n <= Math.min(end, doc.numPages); n++) {
    const page = await doc.getPage(n)
    const content = await page.getTextContent()
    let text = ''
    for (const item of content.items) {
      if ('str' in item) {
        text += item.str
        if (item.hasEOL) text += '\n'
      }
    }
    page.cleanup()
    out += `[Page ${n}]\n${text.trim()}\n\n`
    if (out.length > READ_CHUNK_CHARS) {
      out = `${out.slice(0, READ_CHUNK_CHARS)}\n… (truncated; read the rest in further calls)`
      break
    }
  }
  return {
    output: out.trim() || '(No extractable text in this range; the pages may be scanned images)',
    summary,
  }
}

async function searchText(
  deps: PdfSkillDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const query = String(input.query ?? '').trim()
  if (!query) return err('query must not be empty', t('aiToolSearch', { query: '', count: 0 }))
  const indexPromise = deps.searchIndex()
  if (!indexPromise) return err('Document not ready', t('aiToolSearch', { query, count: 0 }))
  const index = await indexPromise
  const matches = searchInIndex(index, query)
  const lines: string[] = []
  for (const m of matches.slice(0, MAX_MATCHES)) {
    const entry = index[m.pageIndex]!
    const pos = entry.lower.indexOf(query.toLowerCase())
    const from = Math.max(0, pos - 40)
    const snippet = entry.text.slice(from, pos + query.length + 40).replace(/\s+/g, ' ')
    lines.push(`Page ${m.pageIndex + 1}: …${snippet}…`)
  }
  if (matches.length > MAX_MATCHES)
    lines.push(`(${matches.length} matches total; only the first ${MAX_MATCHES} listed)`)
  return {
    output: lines.join('\n') || 'No matches found',
    summary: t('aiToolSearch', { query, count: matches.length }),
  }
}

/** Password-protected (or otherwise read-only) documents refuse every write */
function guardWrite(deps: PdfSkillDeps): ToolExecution | null {
  if (deps.readOnly()) {
    return err(
      'The document is password-protected; changes are disabled in this session',
      'Read-only document',
    )
  }
  return null
}

async function rotatePage(
  deps: PdfSkillDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const page = Number(input.page)
  const summary = t('aiToolRotate', { page })
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const dir = input.direction === 'ccw' ? -90 : 90
  deps.rotatePage(r.origIdx, dir)
  return {
    output: `Rotated page ${page} ${dir === 90 ? 'clockwise' : 'counter-clockwise'}`,
    summary,
  }
}

async function deletePageTool(
  deps: PdfSkillDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const page = Number(input.page)
  const summary = t('aiToolDelete', { page })
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  if (!deps.deletePage(r.origIdx)) {
    return err('Deleting every page is not allowed — one page has to remain', summary)
  }
  return { output: `Removed page ${page}`, summary }
}

async function markupText(
  deps: PdfSkillDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const query = String(input.query ?? '').trim()
  const pageFilter = Number(input.page)
  // Absent type means the schema default; only an explicit unknown kind is rejected
  const type = (typeof input.type === 'string' ? input.type : 'highlight') as MarkupType
  const valid = ['highlight', 'underline', 'strikeout', 'squiggly'].includes(type)
  if (!query) return err('query must not be empty', t('aiToolMarkup', { page: '?' }))

  const indexPromise = deps.searchIndex()
  if (!indexPromise) return err('Document not ready', t('aiToolMarkup', { page: '?' }))
  if (!valid) {
    return err(
      `type must be one of highlight, underline, strikeout, squiggly (got "${String(input.type)}")`,
      t('aiToolMarkup', { page: pageFilter || '?' }),
    )
  }
  const index = await indexPromise

  let matches = searchInIndex(index, query)
  if (Number.isInteger(pageFilter)) matches = matches.filter((m) => m.pageIndex === pageFilter - 1)
  if (matches.length === 0) {
    return err(
      t('aiToolSearch', { query, count: 0 }),
      t('aiToolMarkup', { page: pageFilter || '?' }),
    )
  }

  const capped = matches.slice(0, MAX_MARKUPS)
  const byPage = new Map<number, [number, number, number, number][]>()
  for (const m of capped) {
    for (const rect of m.rects) {
      const list = byPage.get(m.pageIndex)
      if (list) list.push(rect)
      else byPage.set(m.pageIndex, [rect])
    }
  }
  for (const [pageIndex, rects] of byPage) deps.addMarkup(type, pageIndex, rects)

  const pages = [...byPage.keys()].map((p) => p + 1)
  const marks = capped.reduce((sum, m) => sum + m.rects.length, 0)
  const lines = pages.map((p) => {
    const rects = byPage.get(p - 1)!.length
    return `Page ${p}: ${rects} marks (${type})`
  })
  if (matches.length > capped.length) {
    lines.push(`(${matches.length - capped.length} more matches were left unmarked)`)
  }
  lines.push(`(${marks} marks added across ${pages.length} page${pages.length === 1 ? '' : 's'})`)
  return { output: lines.join('\n'), summary: t('aiToolMarkup', { page: pages[0]! }) }
}

export async function executePdfTool(
  deps: PdfSkillDeps,
  call: AgentToolCall,
): Promise<ToolExecution> {
  const input = call.input
  switch (call.name) {
    case 'read_pages':
      return readPages(deps, input)
    case 'search_text':
      return searchText(deps, input)
    case 'goto_page': {
      const summary = t('aiToolGoto', { page: Number(input.page) })
      const r = resolvePage(deps, input.page)
      if ('bad' in r) return err(r.bad, summary)
      deps.gotoPage(r.origIdx + 1)
      return { output: `Jumped to page ${r.origIdx + 1}`, summary }
    }
    case 'list_form_fields':
      return listFormFields(deps)
    case 'fill_form_field': {
      const denied = guardWrite(deps)
      if (denied) return denied
      return fillFormField(deps, input)
    }
    case 'rotate_page': {
      const denied = guardWrite(deps)
      if (denied) return denied
      return rotatePage(deps, input)
    }
    case 'delete_page': {
      const denied = guardWrite(deps)
      if (denied) return denied
      return deletePageTool(deps, input)
    }
    case 'markup_text': {
      const denied = guardWrite(deps)
      if (denied) return denied
      return markupText(deps, input)
    }
    case 'get_outline': {
      const outline = deps.outline()
      const lines: string[] = []
      const walk = (nodes: OutlineNode[], depth: number) => {
        for (const n of nodes) {
          lines.push(`${'  '.repeat(depth)}${n.title}`)
          if (n.items) walk(n.items, depth + 1)
        }
      }
      if (outline) walk(outline, 0)
      return {
        output: lines.join('\n') || 'The document has no outline',
        summary: t('aiToolOutline'),
      }
    }
    default:
      return err(`Unknown tool: ${call.name}`, call.name)
  }
}
