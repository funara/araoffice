/**
 * Word-parity text statistics.
 *
 * Word's rule, verified against Word itself (\`scripts/word-stats.mjs\`):
 *   Words = Asian characters, counted one by one (punctuation included)
 *         + non-Asian words, where a word is any run of non-space characters
 * The dialog also reports the two addends separately — Chinese users care about the
 * Asian-character figure.
 *
 * "Any run of non-space characters" is the part that is easy to get wrong: Word counts
 * \`1.0\`, \`state-of-the-art\` and \`cld_aiops_source_completeness\` as ONE word each, not one
 * per punctuation-delimited fragment. Splitting on punctuation instead inflates the
 * figure badly on technical documents — measured on a 232-page SOP full of dotted
 * version numbers and snake_case table identifiers, it reported 30,306 words against
 * Word's 27,460 (+10.4%). Clauses where Word's own text and our model differ are
 * covered in \`tests/word-count.test.ts\`.
 */

// Han (incl. radicals/compat/ext-B+), kana, hangul, bopomofo, CJK symbols
// and punctuation (U+3001 up: the ideographic space stays whitespace),
// fullwidth forms
import type { Block, ParsedDoc, TableModel } from '@araoffice/docx-engine'
import { computeListMarkers, type ListItemRef } from '@araoffice/docx-engine'

const ASIAN_SRC =
  '[\\u1100-\\u11FF\\u2E80-\\u2FDF\\u3001-\\u303F\\u3040-\\u30FF\\u3100-\\u312F' +
  '\\u3130-\\u318F\\u3190-\\u31BF\\u3400-\\u4DBF\\u4E00-\\u9FFF\\uAC00-\\uD7AF' +
  '\\uF900-\\uFAFF\\uFF01-\\uFF60\\uFFE0-\\uFFE6]|[\\uD840-\\uD87F][\\uDC00-\\uDFFF]'

const ASIAN_RE = new RegExp(ASIAN_SRC, 'g')

/** Asian characters separate words, exactly as whitespace does. */
const NON_WORD_RUN_RE = new RegExp('(?:' + ASIAN_SRC + '|\\s)+', 'g')

export function asianCharCount(text: string): number {
  return (text.match(ASIAN_RE) ?? []).length
}

export function nonAsianWordCount(text: string): number {
  let n = 0
  for (const part of text.split(NON_WORD_RUN_RE)) if (part.length > 0) n++
  return n
}

/** Word's Words figure: asian chars + non-asian words */
export function countWords(text: string): number {
  return asianCharCount(text) + nonAsianWordCount(text)
}

/** minimal structural view of a ProseMirror node (keeps this module dependency-free) */
export interface FlowNode {
  isText?: boolean
  text?: string | null
  type: { name: string }
  attrs: Record<string, unknown>
  descendants: (fn: (node: FlowNode) => boolean | void) => void
}

/** protected display content that lives in another part; Word's count does not include it */
const UNCOUNTED_PREVIEW_LABELS = new Set(['Chart', 'SmartArt', 'Embedded object'])

function tableWords(table: TableModel): number {
  let words = 0
  for (const row of table.rows) {
    for (const cell of row) {
      for (const para of cell.paras) words += countWords(para)
      for (const nested of cell.nestedTables ?? []) words += tableWords(nested)
    }
  }
  return words
}

/** words the editor renders from the parsed model instead of editable flow text */
function displayWords(block: Block): number {
  // only protected blocks are missing from the flow text; everything else (paragraphs,
  // headings, lists, tables, images) is already in the editor document
  if (block.type !== 'passthrough') return 0
  const fd = block.fieldDisplay
  if (fd) {
    if (fd.kind === 'tocLine') {
      // entries carry their number either as literal text (inside left) or as a
      // numbering marker — never count both
      const num = fd.num ?? ''
      const left = fd.left ?? ''
      return (
        countWords(num && !left.startsWith(num) ? `${num} ${left}` : left) +
        countWords(fd.right ?? '')
      )
    }
    return fd.kind === 'text' ? countWords(fd.left ?? '') : 0
  }
  if (block.table) return tableWords(block.table)
  if (block.textboxes) {
    let words = 0
    for (const tb of block.textboxes) {
      for (const para of tb.paras) for (const run of para.runs) words += countWords(run.text ?? '')
    }
    return words
  }
  if (block.previewText && !UNCOUNTED_PREVIEW_LABELS.has(block.label ?? '')) {
    return countWords(block.previewText)
  }
  return 0
}

/**
 * Words for the whole document, matching Word's Review ▸ Word Count figure.
 *
 * Flow text (paragraphs, headings, lists, tables) comes from the editor so edits are
 * reflected; everything the editor keeps as protected display models — TOC entries,
 * field results, text boxes, equations, protected tables, section-break paragraphs —
 * plus footnotes and endnotes comes from the parsed document. Headers/footers are
 * excluded, exactly like Word.
 */
export function countDocumentWords(parsed: ParsedDoc, doc: FlowNode): number {
  let words = 0
  const listRefs: ListItemRef[] = []
  doc.descendants((node) => {
    if (node.isText) {
      words += countWords(node.text ?? '')
      return false
    }
    if (node.type.name === 'docListItem') {
      // Word counts ordered list/heading numbers as words (bullets are not counted);
      // keep descending so the item's own text is counted too
      listRefs.push({
        numId: (node.attrs.numId as string | null) ?? null,
        ilvl: Number(node.attrs.ilvl) || 0,
      })
      return true
    }
    if (node.type.name === 'docInlineMath') {
      words += countWords(String(node.attrs.text ?? ''))
      return false
    }
    if (node.type.name === 'docRuby') {
      words += countWords(`${node.attrs.base ?? ''} ${node.attrs.rt ?? ''}`)
      return false
    }
    if (node.type.name === 'docNestedTable') {
      const model = node.attrs.model as TableModel | null
      if (model) words += tableWords(model)
      return false
    }
    return true
  })
  if (listRefs.length > 0 && parsed.numbering.size > 0) {
    const markers = computeListMarkers(listRefs, parsed.numbering)
    for (let i = 0; i < markers.length; i++) {
      const marker = markers[i]
      if (!marker) continue
      const ref = listRefs[i]
      const level = (ref.numId ? parsed.numbering.get(ref.numId) : undefined)?.levels[ref.ilvl]
      if (level?.numFmt === 'bullet') continue
      words += countWords(marker)
    }
  }
  for (const block of parsed.blocks) words += displayWords(block)
  for (const note of parsed.footnotes ?? []) words += countWords(note.text ?? '')
  for (const note of parsed.endnotes ?? []) words += countWords(note.text ?? '')
  return words
}
