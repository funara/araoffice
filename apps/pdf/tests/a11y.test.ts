import { describe, expect, it } from 'vitest'
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRef } from 'pdf-lib'
import { collectA11yReport, fixA11yMetadataBytes } from '../src/main/a11y-check'
import { applyOutline } from '../src/main/outline-writer'

/**
 * Fixtures are hand-written low-level structures, because pdf-lib has no tagging helpers: the
 * structure tree, /Lang and /MarkInfo are exactly the values this report reads, so the tests
 * construct them directly — the same reasoning the outline and sanitize suites follow.
 */

async function doc(pages = 1): Promise<PDFDocument> {
  const d = await PDFDocument.create()
  for (let i = 0; i < pages; i++) d.addPage([300, 300])
  return d
}

/** Attached: a /MarkInfo with Marked=true, a /StructTreeRoot, plus an optional /Lang */
async function tagDocument(d: PDFDocument, lang?: string): Promise<void> {
  const markInfo = d.context.register(d.context.obj({ Marked: true }))
  d.catalog.set(PDFName.of('MarkInfo'), markInfo)
  const treeRef = d.context.register(
    d.context.obj({ Type: 'StructTreeRoot', K: d.context.obj([]) }),
  )
  d.catalog.set(PDFName.of('StructTreeRoot'), treeRef)
  // fromText, not of: of() takes raw bytes and decodeText mangles text without the BOM
  if (lang) d.catalog.set(PDFName.of('Lang'), d.context.register(PDFHexString.fromText(lang)))
}

/**
 * Structure-tree fixture builder. /S classifies the node; /K holds children.
 */
type StructInput = { s: string; alt?: string; kids?: StructInput[] }

function buildTree(d: PDFDocument, input: StructInput): void {
  const build = (node: StructInput): PDFRef => {
    const obj = d.context.obj({ S: node.s, K: d.context.obj([]) })
    if (node.alt !== undefined)
      obj.set(PDFName.of('Alt'), d.context.register(d.context.obj(node.alt)))
    node.kids?.forEach((kid) => {
      const kidObj = build(kid)
      obj.lookupMaybe(PDFName.of('K'), PDFArray)?.push(kidObj)
    })
    return d.context.register(obj)
  }
  // Attach at the root's /K
  // The dict's lookupMaybe is keyed by /Name; catalog.get returns a PDFRef, so use the name
  const root = d.catalog.lookupMaybe(PDFName.of('StructTreeRoot'), PDFDict)
  if (!root) throw new Error('fixture: no struct tree root')
  const top = build(input)
  root.lookupMaybe(PDFName.of('K'), PDFArray)?.push(top)
}

const report = async (bytes: Uint8Array) =>
  (await collectA11yReport(bytes)).rules.map((r) => `${r.rule}:${r.verdict}`)

describe('fixA11yMetadataBytes', () => {
  it('fills the two gaps the report calls out, and the report then runs them all-pass', async () => {
    const bytes = await (await doc()).save({ useObjectStreams: false })
    expect(await report(bytes)).toEqual(expect.arrayContaining(['title:fail', 'language:fail']))

    const fixed = await fixA11yMetadataBytes(bytes, 'sample-report', 'en')
    expect(fixed.fixed).toBe(2)
    const verdicts = await report(fixed.bytes)
    expect(verdicts).toContain('title:pass')
    expect(verdicts).toContain('language:pass')

    // Title round-trips the caller's text; /Lang lands as written
    const reload = await PDFDocument.load(fixed.bytes, { updateMetadata: false })
    expect(reload.getTitle()).toBe('sample-report')
    const lang = reload.catalog.lookupMaybe(PDFName.of('Lang'), PDFHexString)
    expect(lang?.decodeText()).toBe('en')
  })

  it('is idempotent and keeps an existing non-blank title', async () => {
    const d = await doc()
    d.setTitle('Kept')
    const bytes = await d.save({ useObjectStreams: false })

    const once = await fixA11yMetadataBytes(bytes, 'ignored', 'de')
    expect(once.fixed).toBe(1) // title already present, only /Lang goes in
    const reloaded = await PDFDocument.load(once.bytes, { updateMetadata: false })
    expect(reloaded.getTitle()).toBe('Kept')

    const twice = await fixA11yMetadataBytes(once.bytes, 'ignored', 'de')
    expect(twice.fixed).toBe(0)
  })
})

describe('collectA11yReport', () => {
  it('fails the three base rules on a bare document', async () => {
    const bytes = await (await doc()).save({ useObjectStreams: false })
    expect(await report(bytes)).toEqual([
      'title:fail',
      'language:fail',
      'tagged:fail',
      'altFigures:manual',
      'bookmarksLong:manual',
      'headings:manual',
      'tables:manual',
      'forms:manual',
    ])
  })

  it('passes title and language when the file says so', async () => {
    const d = await doc()
    d.setTitle('Annual Report')
    // Same /Lang encoding the writer and sanitize use; ctx.obj('en') would become a PDFName
    d.catalog.set(PDFName.of('Lang'), d.context.register(PDFHexString.fromText('en')))
    const bytes = await d.save({ useObjectStreams: false })
    expect(await report(bytes)).toEqual([
      'title:pass',
      'language:pass',
      'tagged:fail',
      'altFigures:manual',
      'bookmarksLong:manual',
      'headings:manual',
      'tables:manual',
      'forms:manual',
    ])
  })

  it('walks the tag tree: figures with and without Alt, headings, tables', async () => {
    const d = await doc()
    d.setTitle('Tagged')
    await tagDocument(d, 'en')
    buildTree(d, {
      s: 'Document',
      kids: [
        { s: 'Fig' }, // lacks /Alt
        { s: 'Fig', alt: 'A chart of quarterly revenue' },
        { s: 'H1' },
        { s: 'Table' },
      ],
    })
    const bytes = await d.save({ useObjectStreams: false })
    const r = await collectA11yReport(bytes)
    expect(r.rules.map((x) => x.verdict)).toEqual([
      'pass',
      'pass',
      'pass',
      'fail',
      'manual',
      'pass',
      'manual',
      'manual',
    ])
    // Counts, for the renderer's report line
    const figures = r.rules.find((x) => x.rule === 'altFigures')
    expect(figures?.count).toBe(2)
    const headings = r.rules.find((x) => x.rule === 'headings')
    expect(headings?.count).toBe(1)
    const tables = r.rules.find((x) => x.rule === 'tables')
    expect(tables?.count).toBe(1)
  })

  it('treats a tagged document without figures as an Alt pass, not a fail', async () => {
    const d = await doc()
    d.setTitle('Tagged')
    await tagDocument(d, 'en')
    buildTree(d, { s: 'Document', kids: [{ s: 'P' }] })
    const bytes = await d.save({ useObjectStreams: false })
    const r = await collectA11yReport(bytes)
    expect(r.rules.find((x) => x.rule === 'altFigures')?.verdict).toBe('pass')
    expect(r.rules.find((x) => x.rule === 'altFigures')?.count).toBe(0)
  })

  it('demands bookmarks on long documents and accepts them on short ones', async () => {
    const short = await (await doc()).save({ useObjectStreams: false })
    const shortReport = await collectA11yReport(short)
    expect(shortReport.rules.find((x) => x.rule === 'bookmarksLong')?.verdict).toBe('manual')

    const long = await doc(25)
    const longBytes = await long.save({ useObjectStreams: false })
    expect(
      (await collectA11yReport(longBytes)).rules.find((x) => x.rule === 'bookmarksLong')?.verdict,
    ).toBe('fail')

    applyOutline(long, [{ title: 'Overview', pageIndex: 0 }])
    const longTagged = await long.save({ useObjectStreams: false })
    expect(
      (await collectA11yReport(longTagged)).rules.find((x) => x.rule === 'bookmarksLong')?.verdict,
    ).toBe('pass')
  })

  it('counts AcroForm presence without inventing an automated form verdict', async () => {
    const d = await doc()
    const page = d.getPage(0)
    d.getForm().createTextField('name').addToPage(page, { x: 20, y: 200, width: 160, height: 20 })
    const bytes = await d.save({ useObjectStreams: false })
    const formRule = (await collectA11yReport(bytes)).rules.find((x) => x.rule === 'forms')
    expect(formRule?.verdict).toBe('manual')
    expect(formRule?.count).toBe(1)
  })

  it('survives a cyclic structure tree without hanging', async () => {
    const d = await doc()
    d.setTitle('Cyclic')
    await tagDocument(d, 'en')
    // Build with a cycle: the root children map contains a child whose /K points at the top node
    const ctx = d.context
    const root = d.catalog.lookupMaybe(PDFName.of('StructTreeRoot'), PDFDict)
    if (!root) throw new Error('fixture: no struct tree root')
    const top = ctx.register(ctx.obj({ S: ctx.obj('Document'), K: ctx.obj([]) }))
    const kid = ctx.register(ctx.obj({ S: ctx.obj('Fig'), K: ctx.obj([]) }))
    ctx.lookup(top, PDFDict).lookupMaybe(PDFName.of('K'), PDFArray)?.push(kid)
    ctx.lookup(kid, PDFDict).lookupMaybe(PDFName.of('K'), PDFArray)?.push(top)
    root.lookupMaybe(PDFName.of('K'), PDFArray)?.push(top)
    const bytes = await d.save({ useObjectStreams: false })
    const r = await collectA11yReport(bytes)
    expect(r.rules.find((x) => x.rule === 'altFigures')?.verdict).toBe('fail')
    expect(r.rules.find((x) => x.rule === 'headings')?.count).toBe(0)
  })
})
