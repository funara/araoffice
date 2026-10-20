import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import {
  TABLE_HEADER_FILL,
  type Block,
  type TableCell,
  type TableModel,
} from '@araoffice/docx-engine'
import { tableModelToPmNode, type PmMark, type PmNode } from '../editor/convert'
import { DEFAULT_TABLE_BORDERS } from '../editor/extensions'
import { equationBlockJson, inlineEquationNodeJson } from '../editor/equation'
import { TRACK_IGNORE } from '../editor/revisions'
import { countWords } from '../word-count'

/**
 * Agent protocol: the model runs a multi-turn tool-use loop — it reads the document skeleton
 * (block list with numbered addressing), reads full block content on demand,
 * and mutates the document exclusively through tools (tools.ts). Questions
 * are answered directly in chat without touching the document. Context size
 * stays bounded: previews are clipped and full content is pulled lazily.
 */

// ---- context budgets (characters, ≈4 chars/token) ----

const SELECTION_MAX_CHARS = 24_000
const DOC_CONTEXT_MAX_CHARS = 8_000
const PREVIEW_MAX_CHARS = 60
const PREVIEW_TIGHT_CHARS = 20
export { countWords }
export { buildDocsContext as buildDocContext }

export interface SelectionScope {
  /** top-level child indexes covered by the selection (inclusive) */
  startIndex: number
  endIndex: number
  /** true when the user has an actual range selected (not just a caret) */
  isRange: boolean
}

export function getSelectionScope(editor: Editor): SelectionScope {
  const { from, to, empty } = editor.state.selection
  const doc = editor.state.doc
  let startIndex = -1
  let endIndex = -1
  let index = 0
  doc.forEach((node, offset) => {
    const nodeFrom = offset
    const nodeTo = offset + node.nodeSize
    if (nodeTo > from && nodeFrom < to) {
      if (startIndex === -1) startIndex = index
      endIndex = index
    } else if (empty && from >= nodeFrom && from <= nodeTo && startIndex === -1) {
      startIndex = index
      endIndex = index
    }
    index++
  })
  if (startIndex === -1) {
    startIndex = doc.childCount - 1
    endIndex = doc.childCount - 1
  }
  return { startIndex, endIndex, isRange: !empty }
}

/** ProseMirror positions of a top-level child index range */
export function blockRangePositions(
  editor: Editor,
  startIndex: number,
  endIndex: number,
): { from: number; to: number } {
  const doc = editor.state.doc
  let from = 0
  let to = 0
  let index = 0
  doc.forEach((node, offset) => {
    if (index === startIndex) from = offset
    if (index === endIndex) to = offset + node.nodeSize
    index++
  })
  return { from, to }
}

// ---- tracked deletions (pending revisions are not current content) ----

const hasDelMark = (node: ProseMirrorNode) => node.marks.some((m) => m.type.name === 'del')

/** block text as it reads once pending tracked deletions are applied (textContent minus del runs) */
export function liveText(node: ProseMirrorNode): string {
  if ((node.attrs?.blockRevision as { kind?: string } | null)?.kind === 'del') return ''
  let out = ''
  const walk = (child: ProseMirrorNode): void => {
    if (hasDelMark(child)) return
    if (child.isText) out += child.text ?? ''
    else if (child.isLeaf) out += child.type.spec.leafText?.(child) ?? ''
    else child.forEach(walk)
  }
  node.forEach(walk)
  return out
}

/**
 * The whole block is a pending deletion revision (struck through in the editor).
 * Checked structurally, not by text length: atom leaves (inline formulas, ruby)
 * carry no textContent, so a live formula must still count as live content.
 */
export function isTrackedDeleted(node: ProseMirrorNode): boolean {
  if ((node.attrs?.blockRevision as { kind?: string } | null)?.kind === 'del') return true
  let hasContent = false
  let hasLive = false
  node.descendants((child) => {
    if (child.isText || (child.isInline && child.isLeaf)) {
      hasContent = true
      if (!hasDelMark(child)) hasLive = true
    }
  })
  return hasContent && !hasLive
}

// ---- blocks -> restricted HTML ----

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function inlineToHtml(content: PmNode[] | undefined): string {
  if (!content) return ''
  let html = ''
  for (const node of content) {
    // del-marked runs are pending deletions, not current content
    if ((node.marks ?? []).some((m: PmMark) => m.type === 'del')) continue
    if (node.type === 'hardBreak') {
      html += '<br>'
      continue
    }
    if (node.type === 'docInlineMath') {
      // editor-created formulas round-trip as LaTeX; parsed-from-docx ones
      // degrade to the flat token strip (structure lost if the model rewrites)
      const source = String(node.attrs?.latex ?? '') || String(node.attrs?.text ?? '')
      if (source) html += `<formula>${escapeHtml(source)}</formula>`
      continue
    }
    if (node.type === 'docRuby') {
      html += escapeHtml(String(node.attrs?.base ?? ''))
      continue
    }
    if (node.type !== 'text' || !node.text) continue
    let text = escapeHtml(node.text)
    const marks = node.marks ?? []
    const has = (type: string) => marks.some((m: PmMark) => m.type === type)
    if (has('bold')) text = `<strong>${text}</strong>`
    if (has('italic')) text = `<em>${text}</em>`
    if (has('underline')) text = `<u>${text}</u>`
    if (has('strike')) text = `<s>${text}</s>`
    const link = marks.find((m: PmMark) => m.type === 'link')
    if (link?.attrs?.href) text = `<a href="${escapeHtml(String(link.attrs.href))}">${text}</a>`
    html += text
  }
  return html
}

/** plain text of a paragraph's inline content, hard breaks as newlines */
function inlineToPlainText(content: PmNode[] | undefined): string {
  if (!content) return ''
  return content
    .filter((n) => !(n.marks ?? []).some((m: PmMark) => m.type === 'del'))
    .map((n) => (n.type === 'hardBreak' ? '\n' : n.type === 'text' ? (n.text ?? '') : ''))
    .join('')
}

/** paragraph carries the <pre> preset (see CODE_BLOCK_PRESET) */
function isCodeParagraph(node: PmNode): boolean {
  return (
    node.attrs?.shadingFill === CODE_BLOCK_PRESET.shadingFill &&
    node.attrs?.borders === CODE_BLOCK_PRESET.borders
  )
}

/** paragraph carries the <blockquote> preset (see QUOTE_BLOCK_PRESET) */
function isQuoteParagraph(node: PmNode): boolean {
  return node.attrs?.borders === QUOTE_BLOCK_PRESET.borders && Number(node.attrs?.indentLeft) > 0
}

/** docTable PmNode -> restricted <table> HTML preserving rich inline runs, formulas, and headers */
function serializeTableNodeToHtml(node: PmNode): string {
  const rows = (node.content ?? []).map((rowNode) => {
    const cells = (rowNode.content ?? []).map((cellNode) => {
      const isHeader = cellNode.type === 'docTableHeader'
      const tag = isHeader ? 'th' : 'td'
      const paras = (cellNode.content ?? []).filter(
        (p) => p.type === 'docParagraph' || p.type === 'docListItem',
      )
      const inner = paras
        .map((p) => {
          const h = inlineToHtml(p.content)
          return isHeader ? h.replace(/^<strong>([\s\S]*?)<\/strong>$/, '$1') : h
        })
        .join('<br>')
      return `<${tag}>${inner}</${tag}>`
    })
    return `<tr>${cells.join('')}</tr>`
  })
  return `<table>${rows.join('')}</table>`
}

/**
 * Serialize top-level nodes [startIndex, endIndex] into restricted HTML.
 * Consecutive list items of the same kind are grouped into ul/ol.
 */
export function serializeRangeToHtml(editor: Editor, startIndex: number, endIndex: number): string {
  const json = editor.getJSON() as PmNode
  const children = (json.content ?? []).slice(startIndex, endIndex + 1)
  const parts: string[] = []
  let listBuffer: { kind: string; items: string[] } | null = null

  const flushList = () => {
    if (!listBuffer) return
    const tag = listBuffer.kind === 'ordered' ? 'ol' : 'ul'
    parts.push(`<${tag}>${listBuffer.items.join('')}</${tag}>`)
    listBuffer = null
  }

  for (let i = 0; i < children.length; i++) {
    const node = children[i]
    if (isTrackedDeleted(editor.state.doc.child(startIndex + i))) {
      // the deleted block still separates the surrounding lists in the document
      flushList()
      continue
    }
    if (node.type === 'docHeading') {
      flushList()
      const level = Math.min(Math.max(Number(node.attrs?.level) || 1, 1), 6)
      parts.push(`<h${level}>${inlineToHtml(node.content)}</h${level}>`)
    } else if (node.type === 'docListItem') {
      const kind = (node.attrs?.kind as string) ?? 'bullet'
      if (!listBuffer || listBuffer.kind !== kind) {
        flushList()
        listBuffer = { kind, items: [] }
      }
      listBuffer.items.push(`<li>${inlineToHtml(node.content)}</li>`)
    } else if (node.type === 'docParagraph') {
      flushList()
      if (isCodeParagraph(node)) {
        parts.push(`<pre>${escapeHtml(inlineToPlainText(node.content))}</pre>`)
      } else if (isQuoteParagraph(node)) {
        parts.push(`<blockquote>${inlineToHtml(node.content)}</blockquote>`)
      } else {
        parts.push(`<p>${inlineToHtml(node.content)}</p>`)
      }
    } else if (node.type === 'docTable') {
      flushList()
      parts.push(serializeTableNodeToHtml(node))
    } else {
      flushList()
      const formula = node.attrs?.formulaDisplay as { tokens?: string[]; latex?: string } | null
      if (formula?.latex) {
        // full LaTeX recovered: the model can rewrite the formula losslessly
        parts.push(`<formula>${escapeHtml(formula.latex)}</formula>`)
      } else if (formula?.tokens?.length) {
        parts.push(
          `<p>[Protected formula: ${escapeHtml(formula.tokens.join(' '))}, structure is read-only, kept as is]</p>`,
        )
      } else {
        const label = String(node.attrs?.label ?? node.attrs?.blockType ?? 'content')
        parts.push(`<p>[Protected content: ${escapeHtml(label)}, kept as is]</p>`)
      }
    }
  }
  flushList()
  return parts.join('\n')
}

// ---- document context + layered request builders ----

interface ContextEntry {
  index: number
  type: string
  preview: string
  isHeading: boolean
  deleted: boolean
}

/**
 * Numbered structure snapshot for command-mode addressing: one line per
 * top-level block. Block numbers MUST equal the live PM doc order at execution
 * time, so this is rebuilt per request and never cached.
 */
export function buildDocumentContext(editor: Editor): string {
  const entries: ContextEntry[] = []
  let index = 0
  let fullText = ''
  let hasPendingDeletions = false
  editor.state.doc.forEach((node) => {
    let type: string
    let preview: string
    let isHeading = false
    let deleted: boolean
    if (node.type.name === 'docProtected') {
      // protected blocks can only be block-level deletions (blockRevision)
      deleted = isTrackedDeleted(node)
      type = String(node.attrs.label || node.attrs.blockType || 'protected')
      preview = String(node.attrs.previewText ?? '')
      if (deleted) hasPendingDeletions = true
      else fullText += node.textContent
    } else {
      deleted = isTrackedDeleted(node)
      const live = liveText(node)
      // deleted blocks show their struck text so the model can talk about the
      // revision, but that text never counts as current content
      preview = deleted ? node.textContent : live
      if (!deleted) fullText += live
      if (deleted || live !== node.textContent) hasPendingDeletions = true
      if (node.type.name === 'docHeading') {
        type = `h${Math.min(Math.max(Number(node.attrs.level) || 1, 1), 6)}`
        isHeading = true
      } else if (node.type.name === 'docListItem') {
        type = 'li'
      } else if (node.type.name === 'docTable') {
        type = 'table'
      } else {
        type = 'p'
      }
    }
    entries.push({ index, type, preview: preview.replace(/\s+/g, ' ').trim(), isHeading, deleted })
    index++
  })

  const render = (bodyMax: number) =>
    entries.map((e) => {
      const body = clip(e.preview, e.isHeading ? PREVIEW_MAX_CHARS : bodyMax)
      return `${e.index}|${e.type}|${e.deleted ? '[tracked deletion] ' : ''}${body}`
    })
  let lines = render(PREVIEW_MAX_CHARS)
  if (lines.join('\n').length > DOC_CONTEXT_MAX_CHARS) lines = render(PREVIEW_TIGHT_CHARS)
  if (lines.join('\n').length > DOC_CONTEXT_MAX_CHARS) {
    // keep both ends (numbering must stay verifiable), elide the middle
    let dropStart = Math.floor(lines.length / 3)
    let dropEnd = lines.length - Math.floor(lines.length / 3)
    while (
      dropStart > 1 &&
      dropEnd < lines.length - 1 &&
      [...lines.slice(0, dropStart), ...lines.slice(dropEnd)].join('\n').length >
        DOC_CONTEXT_MAX_CHARS
    ) {
      dropStart = Math.max(1, dropStart - 10)
      dropEnd = Math.min(lines.length - 1, dropEnd + 10)
    }
    lines = [
      ...lines.slice(0, dropStart),
      `…(${dropEnd - dropStart} blocks elided here; numbering is continuous, extrapolate indexes if needed)…`,
      ...lines.slice(dropEnd),
    ]
  }

  const scope = getSelectionScope(editor)
  const selLine = scope.isRange
    ? scope.startIndex === scope.endIndex
      ? `Current selection: block ${scope.startIndex}`
      : `Current selection: blocks ${scope.startIndex}-${scope.endIndex}`
    : `No selection; cursor is in block ${scope.startIndex}`

  // authoritative stats for answer mode (previews above may be clipped)
  const statsLine = `Full-text stats: words ${countWords(fullText)}, characters (no spaces) ${
    fullText.replace(/\s/g, '').length
  }, characters (with spaces) ${fullText.length}`

  return [
    `The document has ${entries.length} blocks, listed below (index|type|content preview):`,
    ...lines,
    statsLine,
    ...(hasPendingDeletions
      ? [
          'Tracked changes: blocks tagged [tracked deletion] and struck-through text inside other blocks are pending deletion revisions — that text is already deleted, is excluded from stats/read_blocks, and must never be deleted or rewritten again; the user accepts/rejects revisions in the Review tab.',
        ]
      : []),
    selLine,
  ].join('\n')
}

function clip(text: string, max: number): string {
  return text.length > max ? text.slice(0, max) + '…' : text
}

/** Blank = exactly one empty paragraph; a textContent check would misread image/chart-only documents as blank. */
export function isBlankDocument(editor: Editor): boolean {
  const doc = editor.state.doc
  if (doc.childCount !== 1) return false
  const first = doc.child(0)
  return first.type.name === 'docParagraph' && first.content.size === 0
}

/**
 * Per-turn context: fresh document skeleton + selection fragment injected
 * as a selection chip.
 */
export function buildDocsContext(editor: Editor): string {
  const isEmptyDoc = isBlankDocument(editor)
  const scope = getSelectionScope(editor)
  const selectionHtml = scope.isRange
    ? clip(serializeRangeToHtml(editor, scope.startIndex, scope.endIndex), SELECTION_MAX_CHARS)
    : ''
  return [
    isEmptyDoc
      ? 'The document is currently blank.'
      : `Document block list:\n${buildDocumentContext(editor)}`,
    selectionHtml
      ? `Content selected by the user (blocks ${scope.startIndex}-${scope.endIndex}):\n${selectionHtml}`
      : '',
  ]
    .filter(Boolean)
    .join('\n\n')
}

// ---- restricted HTML fragment -> PmNode[] ----

export interface NumIds {
  bullet: string | null
  ordered: string | null
}

/**
 * Preset paragraph formats for <pre> / <blockquote>: combinations of existing
 * Block-model properties, so these blocks stay fully editable and round-trip
 * byte-stable through parse -> generate -> parse (no schema/signature change).
 */
export const CODE_BLOCK_PRESET = {
  font: 'Consolas',
  shadingFill: 'F2F2F2',
  borders: 'tblr',
} as const
export const QUOTE_BLOCK_PRESET = {
  color: '666666',
  indentLeft: 720,
  borders: 'l',
} as const

/** pick an existing numbering id of the right kind so new list items join real docx numbering */
export function findNumId(blocks: Block[], kind: 'bullet' | 'ordered'): string | null {
  for (const block of blocks) {
    if (block.type === 'listItem' && block.list?.kind === kind) return block.list.numId
  }
  return null
}

const INLINE_MARK_TAGS: Record<string, string> = {
  strong: 'bold',
  b: 'bold',
  em: 'italic',
  i: 'italic',
  u: 'underline',
  s: 'strike',
  del: 'strike',
  strike: 'strike',
}

function parseInline(element: Node, marks: PmMark[]): PmNode[] {
  const nodes: PmNode[] = []
  element.childNodes.forEach((child) => {
    if (child.nodeType === Node.TEXT_NODE) {
      const text = (child.textContent ?? '').replace(/\s+/g, ' ')
      if (text)
        nodes.push({ type: 'text', text, ...(marks.length > 0 ? { marks: [...marks] } : {}) })
      return
    }
    if (child.nodeType !== Node.ELEMENT_NODE) return
    const el = child as Element
    const tag = el.tagName.toLowerCase()
    if (tag === 'br') {
      nodes.push({ type: 'hardBreak' })
      return
    }
    if (tag === 'a') {
      const href = el.getAttribute('href') ?? ''
      nodes.push(...parseInline(el, [...marks, { type: 'link', attrs: { href, rId: null } }]))
      return
    }
    if (tag === 'formula') {
      const latex = (el.textContent ?? '').trim()
      if (!latex) return
      try {
        nodes.push(inlineEquationNodeJson(latex))
      } catch (e) {
        const reason = e instanceof Error ? e.message : String(e)
        throw new Error(`Cannot parse the LaTeX in <formula> (${reason}): ${latex}`, { cause: e })
      }
      return
    }
    const markType = INLINE_MARK_TAGS[tag]
    if (markType) {
      const next = marks.some((m) => m.type === markType) ? marks : [...marks, { type: markType }]
      nodes.push(...parseInline(el, next))
      return
    }
    // unknown inline tag (span etc.): keep its text content
    nodes.push(...parseInline(el, marks))
  })
  return nodes
}

function blockNode(type: string, attrs: Record<string, unknown>, content: PmNode[]): PmNode {
  const node: PmNode = {
    type,
    attrs: { docxIndex: null, styleId: null, aiChanged: true, ...attrs },
  }
  if (content.length > 0) node.content = content
  return node
}

function parseList(
  el: Element,
  kind: 'bullet' | 'ordered',
  ilvl: number,
  numIds: NumIds,
  out: PmNode[],
): void {
  el.childNodes.forEach((child) => {
    if (child.nodeType !== Node.ELEMENT_NODE) return
    const item = child as Element
    const tag = item.tagName.toLowerCase()
    if (tag === 'ul' || tag === 'ol') {
      parseList(item, tag === 'ol' ? 'ordered' : 'bullet', ilvl + 1, numIds, out)
      return
    }
    if (tag !== 'li') return
    // extract nested lists first so their items follow this one
    const nested: Element[] = []
    item.querySelectorAll(':scope > ul, :scope > ol').forEach((n) => {
      nested.push(n)
      n.remove()
    })
    out.push(
      blockNode(
        'docListItem',
        { kind, numId: numIds[kind], ilvl: Math.min(ilvl, 4) },
        parseInline(item, []),
      ),
    )
    for (const n of nested) {
      parseList(n, n.tagName.toLowerCase() === 'ol' ? 'ordered' : 'bullet', ilvl + 1, numIds, out)
    }
  })
}

/** paragraph texts of one table cell: <br> and nested <p>/<div> split paragraphs */
function cellParas(cell: Element): string[] {
  const paras: string[] = []
  let current = ''
  const push = () => {
    paras.push(current.replace(/\s+/g, ' ').trim())
    current = ''
  }
  const walk = (node: Node) => {
    node.childNodes.forEach((child) => {
      if (child.nodeType === Node.TEXT_NODE) {
        current += child.textContent ?? ''
        return
      }
      if (child.nodeType !== Node.ELEMENT_NODE) return
      const el = child as Element
      const tag = el.tagName.toLowerCase()
      if (tag === 'br') push()
      else if (tag === 'p' || tag === 'div') {
        if (current.trim()) push()
        walk(el)
        push()
      } else walk(el)
    })
  }
  walk(cell)
  if (current.trim()) push()
  const out = paras.filter((p, i) => p !== '' || (i > 0 && i < paras.length - 1))
  return out.length > 0 ? out : ['']
}

/**
 * <table> -> protected table block: display TableModel + generated w:tbl
 * fragment. Cell texts are patched into the fragment at save time via the
 * existing genXml branch of pmDocToSavePlan (patchTableCellTexts), so the
 * model row/col counts must mirror the generated grid.
 */
function parseTable(el: Element): PmNode | null {
  const trs = Array.from(el.querySelectorAll('tr'))
  if (trs.length === 0) return null
  const rawRows = trs.map((tr) => Array.from(tr.children).filter((c) => /^t[hd]$/i.test(c.tagName)))
  const cols = Math.max(...rawRows.map((r) => r.length))
  if (cols === 0) return null
  const headerRow = rawRows[0].some((c) => c.tagName.toLowerCase() === 'th')
  const rows: TableCell[][] = rawRows.map((cells, r) => {
    const isHeader = headerRow && r === 0
    const row: TableCell[] = cells.map((cell) => ({
      paras: cellParas(cell),
      ...(isHeader ? { bold: true, fill: TABLE_HEADER_FILL } : {}),
    }))
    while (row.length < cols)
      row.push({ paras: [''], ...(isHeader ? { bold: true, fill: TABLE_HEADER_FILL } : {}) })
    return row
  })
  const table: TableModel = {
    rows,
    borders: DEFAULT_TABLE_BORDERS,
    colWidthsPct: Array.from({ length: cols }, () => 100 / cols),
  }
  return tableModelToPmNode(table)
}

/** <pre> -> mono/shaded paragraph; newlines preserved as hard breaks */
function parseCodeBlock(el: Element): PmNode | null {
  const text = (el.textContent ?? '').replace(/^\n/, '').replace(/\s+$/, '')
  if (!text) return null
  // Latin slot only: monospace applies to code text, an inherited CJK font stays intact
  const mark: PmMark = { type: 'docTextStyle', attrs: { fontAscii: CODE_BLOCK_PRESET.font } }
  const content: PmNode[] = []
  text.split('\n').forEach((line, i) => {
    if (i > 0) content.push({ type: 'hardBreak' })
    if (line !== '') content.push({ type: 'text', text: line, marks: [mark] })
  })
  return blockNode(
    'docParagraph',
    { shadingFill: CODE_BLOCK_PRESET.shadingFill, borders: CODE_BLOCK_PRESET.borders },
    content,
  )
}

/** <blockquote> -> indented gray paragraph with a left border */
function parseBlockquote(el: Element): PmNode | null {
  const mark: PmMark = { type: 'docTextStyle', attrs: { color: QUOTE_BLOCK_PRESET.color } }
  const inline = parseInline(el, [mark])
  if (inline.length === 0) return null
  return blockNode(
    'docParagraph',
    { indentLeft: QUOTE_BLOCK_PRESET.indentLeft, borders: QUOTE_BLOCK_PRESET.borders },
    inline,
  )
}

/**
 * Parse a restricted HTML fragment returned by the model into top-level PmNodes.
 * Tolerates markdown code fences and plain-text responses.
 */
export function parseHtmlFragment(raw: string, numIds: NumIds): PmNode[] {
  let text = raw.trim()
  const fence = /```(?:html)?\s*([\s\S]*?)```/.exec(text)
  if (fence) text = fence[1].trim()
  if (!text) return []

  // plain text response (no tags): one paragraph per blank-line-separated chunk
  if (!/<[a-z][\s\S]*>/i.test(text)) {
    return text
      .split(/\n{2,}/)
      .map((para) =>
        blockNode('docParagraph', {}, [{ type: 'text', text: para.replace(/\s+/g, ' ').trim() }]),
      )
      .filter((n) => n.content?.[0]?.text)
  }

  const parsed = new DOMParser().parseFromString(text, 'text/html')
  const out: PmNode[] = []
  const pushFormula = (el: Element) => {
    const latex = (el.textContent ?? '').trim()
    if (!latex) return
    try {
      out.push(equationBlockJson(latex))
    } catch (e) {
      const reason = e instanceof Error ? e.message : String(e)
      throw new Error(`Cannot parse the LaTeX in <formula> (${reason}): ${latex}`, { cause: e })
    }
  }
  parsed.body.childNodes.forEach((child) => {
    if (child.nodeType === Node.TEXT_NODE) {
      const t = (child.textContent ?? '').trim()
      if (t) out.push(blockNode('docParagraph', {}, [{ type: 'text', text: t }]))
      return
    }
    if (child.nodeType !== Node.ELEMENT_NODE) return
    const el = child as Element
    const tag = el.tagName.toLowerCase()
    const headingMatch = /^h([1-6])$/.exec(tag)
    if (headingMatch) {
      out.push(blockNode('docHeading', { level: Number(headingMatch[1]) }, parseInline(el, [])))
    } else if (tag === 'ul' || tag === 'ol') {
      parseList(el, tag === 'ol' ? 'ordered' : 'bullet', 0, numIds, out)
    } else if (tag === 'table') {
      const node = parseTable(el)
      if (node) out.push(node)
    } else if (tag === 'pre') {
      const node = parseCodeBlock(el)
      if (node) out.push(node)
    } else if (tag === 'blockquote') {
      const node = parseBlockquote(el)
      if (node) out.push(node)
    } else if (tag === 'formula') {
      pushFormula(el)
    } else if (tag === 'p' || tag === 'div') {
      const inline = parseInline(el, [])
      if (inline.length > 0) out.push(blockNode('docParagraph', {}, inline))
    } else {
      // unknown block-ish tag: salvage the text
      const t = (el.textContent ?? '').trim()
      if (t) out.push(blockNode('docParagraph', {}, [{ type: 'text', text: t }]))
    }
  })
  return out
}

// ---- applying parsed fragments ----

/** record the AI's content change as tracked revisions under this author */
export interface AiTrack {
  author: string
}

function revisionDate(): string {
  return new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
}

/** blocks whose content ins/del marks can fully represent (tracked replace) */
const TRACKABLE_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

/**
 * Replace the top-level block range with the parsed nodes (marked aiChanged).
 * With `track`, and when both sides are plain text blocks, this becomes a
 * tracked rewrite instead: the old blocks stay struck through
 * (del) and the new blocks follow with ins marks — accept/reject via Review.
 */
/**
 * When replacing blocks (e.g. translation, rewrite, edit), preserve formatting
 * from the original blocks:
 * 1. For tables: inherit borders, colWidthsPct, colwidth, tblStyleId, cellMar,
 *    align, and docxIndex so tables never lose borders, column widths, or byte-level patching.
 * 2. For paragraphs/headings: inherit font, size, line spacing, margins, indents,
 *    alignment, and docxIndex so the document's original formatting is preserved.
 */
function preserveBlockFormatting(newNodes: PmNode[], oldBlocks: ProseMirrorNode[]): PmNode[] {
  if (oldBlocks.length === 0) return newNodes

  const oldTables = oldBlocks.filter((b) => b.type.name === 'docTable')
  let oldTableIdx = 0

  return newNodes.map((newNode, idx) => {
    // 1. Table preservation:
    if (newNode.type === 'docTable') {
      const oldTable = oldTables[oldTableIdx++] ?? oldTables[0]
      if (oldTable) {
        const oldAttrs = (oldTable.attrs ?? {}) as Record<string, unknown>
        const mergedAttrs = {
          ...newNode.attrs,
          docxIndex: oldAttrs.docxIndex ?? newNode.attrs?.docxIndex ?? null,
          borders: oldAttrs.borders ?? newNode.attrs?.borders ?? DEFAULT_TABLE_BORDERS,
          colWidthsPct: oldAttrs.colWidthsPct ?? newNode.attrs?.colWidthsPct ?? null,
          widthPx: oldAttrs.widthPx ?? newNode.attrs?.widthPx ?? null,
          widthPct: oldAttrs.widthPct ?? newNode.attrs?.widthPct ?? null,
          cellMar: oldAttrs.cellMar ?? newNode.attrs?.cellMar ?? null,
          tblAlign: oldAttrs.tblAlign ?? newNode.attrs?.tblAlign ?? null,
          indentTwips: oldAttrs.indentTwips ?? newNode.attrs?.indentTwips ?? null,
          tblStyleId: oldAttrs.tblStyleId ?? newNode.attrs?.tblStyleId ?? null,
          bidiVisual: oldAttrs.bidiVisual ?? newNode.attrs?.bidiVisual ?? false,
        }

        const rows = (newNode.content ?? []).map((rowNode, rIdx) => {
          const oldRow =
            oldTable.maybeChild(rIdx) ??
            oldTable.maybeChild(Math.min(rIdx, oldTable.childCount - 1))
          const cells = (rowNode.content ?? []).map((cellNode, cIdx) => {
            const oldCell =
              oldRow?.maybeChild(cIdx) ??
              oldRow?.maybeChild(Math.min(cIdx, (oldRow?.childCount ?? 1) - 1))
            if (!oldCell) return cellNode
            const oldCellAttrs = (oldCell.attrs ?? {}) as Record<string, unknown>
            return {
              ...cellNode,
              attrs: {
                ...cellNode.attrs,
                colwidth: oldCellAttrs.colwidth ?? cellNode.attrs?.colwidth ?? null,
                cellMar: oldCellAttrs.cellMar ?? cellNode.attrs?.cellMar ?? null,
                fill:
                  oldCellAttrs.fill === undefined
                    ? (cellNode.attrs?.fill ?? null)
                    : oldCellAttrs.fill,
                color: oldCellAttrs.color ?? cellNode.attrs?.color ?? null,
                align: oldCellAttrs.align ?? cellNode.attrs?.align ?? null,
                vAlign: oldCellAttrs.vAlign ?? cellNode.attrs?.vAlign ?? null,
                borders: oldCellAttrs.borders ?? cellNode.attrs?.borders ?? null,
              },
            }
          })
          return {
            ...rowNode,
            attrs: {
              ...rowNode.attrs,
              heightTwips: oldRow?.attrs?.heightTwips ?? rowNode.attrs?.heightTwips ?? null,
              rawTrPr: oldRow?.attrs?.rawTrPr ?? rowNode.attrs?.rawTrPr ?? null,
            },
            content: cells,
          }
        })

        return {
          ...newNode,
          attrs: mergedAttrs,
          content: rows,
        }
      }
      return {
        ...newNode,
        attrs: {
          ...newNode.attrs,
          borders: newNode.attrs?.borders ?? DEFAULT_TABLE_BORDERS,
        },
      }
    }

    // 2. Paragraph and heading preservation:
    const correspondingOld =
      idx < oldBlocks.length && oldBlocks[idx].type.name === newNode.type
        ? oldBlocks[idx]
        : oldBlocks.find((b) => b.type.name === newNode.type)

    if (correspondingOld) {
      const oldAttrs = (correspondingOld.attrs ?? {}) as Record<string, unknown>
      const preservedKeys = [
        'align',
        'lineSpacing',
        'lineRule',
        'lineRawTwips',
        'indentLeft',
        'indentRight',
        'indentFirstLine',
        'spaceBefore',
        'spaceAfter',
        'pageBreakBefore',
        'shadingFill',
        'borders',
        'styleId',
      ] as const
      const preserved: Record<string, unknown> = {}
      for (const k of preservedKeys) {
        if (oldAttrs[k] !== undefined && oldAttrs[k] !== null) {
          preserved[k] = oldAttrs[k]
        }
      }
      return {
        ...newNode,
        attrs: {
          ...preserved,
          ...newNode.attrs,
          docxIndex: idx < oldBlocks.length ? (oldBlocks[idx].attrs?.docxIndex ?? null) : null,
          aiChanged: true,
        },
      }
    }

    return newNode
  })
}

export function replaceBlockRange(
  editor: Editor,
  startIndex: number,
  endIndex: number,
  nodes: PmNode[],
  track?: AiTrack,
): boolean {
  if (nodes.length === 0) return false
  const { from, to } = blockRangePositions(editor, startIndex, endIndex)
  const oldBlocks: ProseMirrorNode[] = []
  for (let i = startIndex; i <= endIndex && i < editor.state.doc.childCount; i++) {
    oldBlocks.push(editor.state.doc.child(i))
  }
  const preservedNodes = preserveBlockFormatting(nodes, oldBlocks)
  const pmNodes = preservedNodes.map((n) => editor.schema.nodeFromJSON(n))
  let oldTrackable = true
  editor.state.doc.nodesBetween(from, to, (node, _pos, parent) => {
    if (parent === editor.state.doc && !TRACKABLE_TYPES.has(node.type.name)) oldTrackable = false
    return false
  })
  const newTrackable = nodes.every((n) => TRACKABLE_TYPES.has(n.type))
  if (track && oldTrackable && newTrackable) {
    const { ins, del } = editor.schema.marks
    const date = revisionDate()
    // revision marks are the change indicator; no yellow aiChanged on top
    const tracked = preservedNodes.map((n) =>
      editor.schema.nodeFromJSON({ ...n, attrs: { ...n.attrs, aiChanged: false } }),
    )
    const inserted = tracked.reduce((size, n) => size + n.nodeSize, 0)
    const tr = editor.state.tr
    tr.setMeta(TRACK_IGNORE, true)
    if (editor.state.doc.textBetween(from, to, '\n').trim() === '') {
      // nothing to strike through (blank paragraphs): replace outright
      tr.replaceWith(from, to, tracked)
      tr.addMark(from, from + inserted, ins.create({ author: track.author, date }))
    } else {
      tr.insert(to, tracked)
      tr.addMark(to, to + inserted, ins.create({ author: track.author, date }))
      tr.addMark(from, to, del.create({ author: track.author, date }))
    }
    editor.view.dispatch(tr)
    return true
  }
  if (track) {
    const date = revisionDate()
    const oldNodes: ProseMirrorNode[] = []
    editor.state.doc.nodesBetween(from, to, (node, _pos, parent) => {
      if (parent === editor.state.doc) {
        oldNodes.push(
          node.type.create(
            {
              ...node.attrs,
              aiChanged: false,
              blockRevision: { kind: 'del', author: track.author, date },
            },
            node.content,
            node.marks,
          ),
        )
      }
      return false
    })
    const inserted = pmNodes.map((node) =>
      node.type.create(
        {
          ...node.attrs,
          aiChanged: false,
          blockRevision: { kind: 'ins', author: track.author, date },
        },
        node.content,
        node.marks,
      ),
    )
    const tr = editor.state.tr.replaceWith(from, to, [...oldNodes, ...inserted])
    tr.setMeta(TRACK_IGNORE, true)
    editor.view.dispatch(tr)
    return true
  }
  editor.view.dispatch(editor.state.tr.replaceWith(from, to, pmNodes))
  return true
}

/** insert the parsed nodes after the given top-level block index */
export function insertBlocksAfter(
  editor: Editor,
  index: number,
  nodes: PmNode[],
  track?: AiTrack,
): boolean {
  if (nodes.length === 0) return false
  const { to } = blockRangePositions(editor, index, index)
  const pmNodes = nodes.map((n) =>
    editor.schema.nodeFromJSON(track ? { ...n, attrs: { ...n.attrs, aiChanged: false } } : n),
  )
  const tr = editor.state.tr.insert(to, pmNodes)
  if (track) {
    const date = revisionDate()
    let offset = to
    for (const node of pmNodes) {
      if (TRACKABLE_TYPES.has(node.type.name)) {
        tr.addMark(
          offset,
          offset + node.nodeSize,
          editor.schema.marks.ins.create({ author: track.author, date }),
        )
      } else {
        tr.setNodeMarkup(offset, undefined, {
          ...node.attrs,
          blockRevision: { kind: 'ins', author: track.author, date },
        })
      }
      offset += node.nodeSize
    }
    tr.setMeta(TRACK_IGNORE, true)
  }
  editor.view.dispatch(tr)
  return true
}
