import { PDFArray, PDFBool, PDFDict, PDFDocument, PDFHexString, PDFName, PDFString } from 'pdf-lib'
import type { PDFObject } from 'pdf-lib'
import type { A11yReport, A11yRuleResult } from '../shared/ipc'

/**
 * The accessibility report: what the file itself can answer, plus honest "manual" verdicts where
 * the answer is not in the file (colour contrast, reading order, screen flicker) — the same split
 * Acrobat's checker draws between automated rules and ones needing a person.
 *
 * Structural verdicts stay report-only. The one fix that lives here is metadata-level (title and
 * language below): plain key writes pdf-lib can do without touching content, unlike structure
 * trees, which hand-building would violate the byte-preservation invariant with.
 *
 * The structure-tree walk follows the never-crash rules: depth cap and an identity-based seen-set,
 * because a corrupt or cyclic tree is untrusted input like any other. pdf-lib resolves each ref to
 * a stable object, so object identity doubles as the "visited" marker.
 */

/** Past this many pages, navigation is expected to exist as bookmarks */
const BOOKMARK_PAGES = 20
const MAX_DEPTH = 64

interface StructureScan {
  figures: number
  figuresWithoutAlt: number
  headings: number
  tables: number
}

const emptyScan = (): StructureScan => ({
  figures: 0,
  figuresWithoutAlt: 0,
  headings: 0,
  tables: 0,
})

/** /Title-style values arrive as PDFString, PDFHexString, or — when stored indirect — a PDFRef */
function decodeStr(ctx: PDFDocument['context'], obj: PDFObject | undefined): string | null {
  // A ref only resolves through the context: dict methods read their own keys, not arbitrary
  // objects — the trap behind the outline read-back failure earlier in this session
  const resolved = ctx.lookup(obj as never)
  if (resolved instanceof PDFString || resolved instanceof PDFHexString) {
    const text = resolved.decodeText()
    return text.trim().length > 0 ? text : null
  }
  return null
}

/** Alt for text is textual; a placed string means the rule's target field can hold content */
const FIGURE = 'Fig'
const TABLE = 'Table'
const HEADING = /^H[1-6]$/

function scanNode(
  dict: PDFDict,
  ctx: PDFDocument['context'],
  depth: number,
  seen: WeakSet<object>,
  acc: StructureScan,
): void {
  if (depth > MAX_DEPTH) return
  if (seen.has(dict)) return
  seen.add(dict)
  const s = ctx.lookupMaybe(dict.get(PDFName.of('S')), PDFName)?.decodeText()
  if (s === FIGURE) {
    acc.figures += 1
    // /Alt is optional text; a figure that cannot say what it shows is exactly this rule's finding
    if (!dict.has(PDFName.of('Alt'))) acc.figuresWithoutAlt += 1
  } else if (s === TABLE) {
    acc.tables += 1
  } else if (s && HEADING.test(s)) {
    acc.headings += 1
  }
  const kids = ctx.lookupMaybe(dict.get(PDFName.of('K')), PDFArray)
  if (!kids) return
  for (let i = 0; i < kids.size(); i++) {
    // Kids are element dicts or marked-content references ([pageRef, mcId]); only dicts recurse
    const kid = ctx.lookupMaybe(kids.get(i), PDFDict)
    if (kid) scanNode(kid, ctx, depth + 1, seen, acc)
  }
}

function scanTree(pdfDoc: PDFDocument, catalog: PDFDict): StructureScan {
  const ctx = pdfDoc.context
  const root = ctx.lookupMaybe(catalog.get(PDFName.of('StructTreeRoot')), PDFDict)
  if (!root) return emptyScan()
  const acc = emptyScan()
  scanNode(root, ctx, 0, new WeakSet(), acc)
  return acc
}

function scanFiguresVerdict(scan: StructureScan, hasTree: boolean): A11yRuleResult['verdict'] {
  if (!hasTree) return 'manual'
  if (scan.figures === 0) return 'pass'
  return scan.figuresWithoutAlt === 0 ? 'pass' : 'fail'
}

function headingsVerdict(scan: StructureScan, hasTree: boolean): A11yRuleResult['verdict'] {
  if (!hasTree) return 'manual'
  // Headings present is the automatable half; a good order needs a reader
  return scan.headings > 0 ? 'pass' : 'manual'
}

function bookmarksVerdict(pageCount: number, hasOutline: boolean): A11yRuleResult['verdict'] {
  if (pageCount < BOOKMARK_PAGES) return 'manual'
  return hasOutline ? 'pass' : 'fail'
}

/**
 * Run the accessibility report over a document. Read-only: load, walk, verdict, return.
 */
export async function collectA11yReport(bytes: Uint8Array): Promise<A11yReport> {
  const pdfDoc = await PDFDocument.load(bytes, { updateMetadata: false })
  const ctx = pdfDoc.context
  const catalog = pdfDoc.catalog
  const pageCount = pdfDoc.getPageCount()

  const markInfo = ctx.lookupMaybe(catalog.get(PDFName.of('MarkInfo')), PDFDict)
  const markedRef = markInfo?.get(PDFName.of('Marked'))
  const markedValue = markedRef ? ctx.lookup(markedRef) : undefined
  const marked = markedValue instanceof PDFBool ? markedValue.asBoolean() === true : false
  const treeRef = catalog.get(PDFName.of('StructTreeRoot'))
  const hasTree = Boolean(treeRef)
  const scan = marked && hasTree ? scanTree(pdfDoc, catalog) : emptyScan()

  const info = ctx.lookupMaybe(ctx.trailerInfo.Info, PDFDict)
  const lang = decodeStr(ctx, catalog.get(PDFName.of('Lang')))
  const title = decodeStr(ctx, info?.get(PDFName.of('Title')))
  const hasBookmarks = catalog.has(PDFName.of('Outlines'))
  const hasForms = catalog.has(PDFName.of('AcroForm'))

  return {
    pageCount,
    rules: [
      { rule: 'title', verdict: title ? 'pass' : 'fail', count: pageCount },
      { rule: 'language', verdict: lang ? 'pass' : 'fail' },
      { rule: 'tagged', verdict: marked && hasTree ? 'pass' : 'fail' },
      {
        rule: 'altFigures',
        verdict: scanFiguresVerdict(scan, marked && hasTree),
        count: scan.figures,
      },
      {
        rule: 'bookmarksLong',
        verdict: bookmarksVerdict(pageCount, hasBookmarks),
        count: pageCount,
      },
      { rule: 'headings', verdict: headingsVerdict(scan, marked && hasTree), count: scan.headings },
      { rule: 'tables', verdict: 'manual', count: scan.tables },
      { rule: 'forms', verdict: 'manual', count: hasForms ? 1 : 0 },
    ],
  }
}

/**
 * Fill the two metadata gaps the report calls out: a missing catalogue /Lang and a missing or
 * blank /Info /Title. Only writes what is absent; returns the new bytes and how many fields
 * went in, so the caller can skip the save entirely (byte-identical round trip returns nothing).
 * The title comes from the caller (the file name is the honest default); the lang tag likewise.
 */
export async function fixA11yMetadataBytes(
  bytes: Uint8Array,
  title: string,
  lang: string,
): Promise<{ bytes: Uint8Array; fixed: number }> {
  const pdfDoc = await PDFDocument.load(bytes, { updateMetadata: false })
  const ctx = pdfDoc.context
  const catalog = pdfDoc.catalog
  let fixed = 0

  if (!decodeStr(ctx, catalog.get(PDFName.of('Lang')))) {
    catalog.set(PDFName.of('Lang'), PDFHexString.fromText(lang))
    fixed += 1
  }

  let info = ctx.lookupMaybe(ctx.trailerInfo.Info, PDFDict)
  if (!info) {
    // The trailer must carry the dict itself — pdf-lib serializes from trailerInfo.Info
    info = ctx.obj({})
    ctx.trailerInfo.Info = ctx.register(info)
  }
  if (!decodeStr(ctx, info.get(PDFName.of('Title')))) {
    // fromText round-trips arbitrary text (CJK filenames included) where PDFString mangles it
    info.set(PDFName.of('Title'), PDFHexString.fromText(title))
    fixed += 1
  }

  return { bytes: await pdfDoc.save({ useObjectStreams: false }), fixed }
}

/** PDFRawStream is referenced by the scan tree fixtures only through re-exports; keep the import list honest */
