import { describe, expect, it } from 'vitest'
import { PDFArray, PDFDict, PDFDocument, PDFName } from 'pdf-lib'
import { sanitizePdfBytes, scanSanitizeCounts } from '../src/main/save-pdf'
import type { SanitizeRequest } from '../src/shared/ipc'

/** Every category on: the case that removes the most and is the easiest to over-assert */
const ALL: SanitizeRequest = {
  metadata: true,
  javascript: true,
  attachments: true,
  comments: true,
  thumbnails: true,
  privateData: true,
}

const NONE: SanitizeRequest = {
  metadata: false,
  javascript: false,
  attachments: false,
  comments: false,
  thumbnails: false,
  privateData: false,
}

/**
 * A document with something removable in every category, built against the low-level API because
 * pdf-lib only offers high-level accessors for a subset (title, author, subject, keywords — but no
 * XMP, embedded files or JavaScript).
 *
 * /Thumb is written as a plain dict: sanitize only has to find and drop the entry, and a viewer is
 * never asked to interpret it here.
 */
async function makeSanitizable(pageCount = 2): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  doc.setTitle('Secret Title')
  doc.setAuthor('Secret Author')
  doc.setSubject('Secret Subject')
  for (let i = 0; i < pageCount; i++) doc.addPage([300, 300])

  const ctx = doc.context
  const xmp = ctx.stream('<x:xmpmeta>SECRET</x:xmpmeta>', { Type: 'Metadata', Subtype: 'XML' })
  doc.catalog.set(PDFName.of('Metadata'), ctx.register(xmp))
  doc.catalog.set(
    PDFName.of('Names'),
    ctx.obj({
      JavaScript: ctx.obj({ Names: ctx.obj([]) }),
      EmbeddedFiles: ctx.obj({ Names: ctx.obj([]) }),
      // A branch sanitize does not target, to pin that the /Names dict is edited in place
      Dests: ctx.obj({ Names: ctx.obj([]) }),
    }),
  )
  doc.catalog.set(
    PDFName.of('OpenAction'),
    ctx.obj({ Type: ctx.obj('Action'), S: ctx.obj('JavaScript') }),
  )
  doc.catalog.set(PDFName.of('AA'), ctx.obj({}))
  doc.catalog.set(PDFName.of('Collection'), ctx.obj({}))
  doc.catalog.set(PDFName.of('PieceInfo'), ctx.obj({ ARA: ctx.obj({}) }))

  for (const page of doc.getPages()) {
    page.node.set(
      PDFName.of('Annots'),
      ctx.obj([ctx.obj({ Type: ctx.obj('Annot'), Subtype: ctx.obj('Text') })]),
    )
    page.node.set(PDFName.of('Thumb'), ctx.obj({ Type: ctx.obj('XObject') }))
    page.node.set(PDFName.of('PieceInfo'), ctx.obj({ ARA: ctx.obj({}) }))
  }
  return doc.save({ useObjectStreams: false })
}

/** The catalog dict of a saved-and-reloaded document, for has() probes */
async function catalogOf(bytes: Uint8Array) {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  return doc.catalog
}

describe('scanSanitizeCounts', () => {
  it('matches the removal report of the same run — the census and removal share one walk', async () => {
    const bytes = await makeSanitizable()
    const counts = await scanSanitizeCounts(bytes)
    const run = await sanitizePdfBytes(bytes, ALL)
    expect(counts).toEqual(run!.report)
  })

  it('never mutates the source bytes', async () => {
    const bytes = await makeSanitizable()
    const snapshot = bytes.slice()
    await scanSanitizeCounts(bytes)
    expect(Buffer.from(bytes).equals(Buffer.from(snapshot))).toBe(true)
  })
})

describe('sanitizePdfBytes', () => {
  it('removes everything requested and leaves the pages alone', async () => {
    const source = await makeSanitizable()
    const { bytes, report } = (await sanitizePdfBytes(source, ALL))!

    // Exact where the fixture controls every occurrence
    expect(report.javascript).toBe(3) // /Names JavaScript + /OpenAction + /AA
    expect(report.attachments).toBe(2) // EmbeddedFiles + /Collection
    expect(report.comments).toBe(2) // one markup annot per page
    expect(report.thumbnails).toBe(2)
    expect(report.privateData).toBe(3) // catalog + two pages
    // pdf-lib's create() writes its own Info fields, so the count is not a closed set here
    expect(report.metadata).toBeGreaterThanOrEqual(4) // 3 fixture fields + XMP

    const out = await PDFDocument.load(bytes, { updateMetadata: false })
    expect(out.getPageCount()).toBe(2)
    const catalog = out.catalog
    expect(catalog.has(PDFName.of('Metadata'))).toBe(false)
    expect(catalog.has(PDFName.of('OpenAction'))).toBe(false)
    expect(catalog.has(PDFName.of('AA'))).toBe(false)
    expect(catalog.has(PDFName.of('Collection'))).toBe(false)
    expect(catalog.has(PDFName.of('PieceInfo'))).toBe(false)
    const names = catalog.lookupMaybe(PDFName.of('Names'), PDFDict)!
    expect(names.has(PDFName.of('JavaScript'))).toBe(false)
    expect(names.has(PDFName.of('EmbeddedFiles'))).toBe(false)
    const info = out.context.lookupMaybe(out.context.trailerInfo.Info, PDFDict)!
    for (const key of ['Title', 'Author', 'Subject', 'CreationDate', 'ModDate']) {
      expect(info.has(PDFName.of(key)), `/Info ${key}`).toBe(false)
    }
    // The names tree survives with its other branch: a wholesale /Names delete would have taken
    // named destinations down along with the scripts and files
    expect(names.has(PDFName.of('Dests'))).toBe(true)
  })

  it('removes only the requested category and reports nothing else', async () => {
    const { bytes, report } = (await sanitizePdfBytes(await makeSanitizable(), {
      ...NONE,
      comments: true,
    }))!
    expect(report).toEqual({
      metadata: 0,
      javascript: 0,
      attachments: 0,
      comments: 2,
      thumbnails: 0,
      privateData: 0,
    })

    const out = await PDFDocument.load(bytes, { updateMetadata: false })
    // Untouched categories are untouched: privacy work must not eat the document as a side effect
    expect(out.getTitle()).toBe('Secret Title')
    expect(out.getAuthor()).toBe('Secret Author')
    expect((await catalogOf(bytes)).has(PDFName.of('Metadata'))).toBe(true)
    const page = out.getPage(0)
    expect(page.node.has(PDFName.of('Thumb'))).toBe(true)
    expect(page.node.has(PDFName.of('PieceInfo'))).toBe(true)
  })

  it('keeps form field widgets when removing comments', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const ctx = doc.context
    doc
      .getPage(0)
      .node.set(
        PDFName.of('Annots'),
        ctx.obj([
          ctx.obj({ Type: ctx.obj('Annot'), Subtype: ctx.obj('Text') }),
          ctx.obj({ Type: ctx.obj('Annot'), Subtype: ctx.obj('Widget') }),
        ]),
      )
    const bytes = await doc.save({ useObjectStreams: false })

    const { bytes: saved, report } = (await sanitizePdfBytes(bytes, { ...NONE, comments: true }))!
    expect(report.comments).toBe(1)

    const out = await PDFDocument.load(saved, { updateMetadata: false })
    const annots = out.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)!
    expect(annots.size()).toBe(1)
    const survivor = out.context.lookupMaybe(annots.get(0), PDFDict)
    expect(survivor?.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText()).toBe('Widget')
  })

  it('returns an all-zero report for a document with nothing to remove', async () => {
    const { report, bytes } = (await sanitizePdfBytes(await makeSanitizable(1), {
      ...NONE,
      thumbnails: true,
    }))!
    expect(report.thumbnails).toBe(1)
    expect((await PDFDocument.load(bytes)).getPageCount()).toBe(1)
  })

  it('refuses an empty request instead of rewriting the file for nothing', async () => {
    expect(await sanitizePdfBytes(await makeSanitizable(), NONE)).toBeNull()
  })
})
