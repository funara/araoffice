import { describe, expect, it } from 'vitest'
import { PDFArray, PDFDict, PDFDocument, PDFName, PDFStream, PDFString } from 'pdf-lib'
import { deleteAnnotsBytes } from '../src/main/delete-annots'
import { annotationsToRows } from '../src/renderer/comments'
import type { AnnotMatch } from '../src/shared/ipc'
import type { StringKey } from '../src/renderer/i18n/locale'

/**
 * The matcher is judged by consequences, not internals: exactly the targeted annotation goes,
 * everything else — text, other annots, widgets — survives; and every refusal leaves the file
 * unchanged.
 */

type AnnotSpec = {
  subtype: string
  rect: number[]
  contents: string
  nm?: string
  widget?: boolean
}

function annotSpec(ctx: PDFDocument['context'], a: AnnotSpec) {
  return ctx.obj({
    Type: ctx.obj('Annot'),
    Subtype: ctx.obj(a.subtype),
    Rect: ctx.obj(a.rect),
    // ctx.obj('string') builds a NAME here, not text — annotations carry PDFString values
    Contents: PDFString.of(a.contents),
    ...(a.nm ? { NM: PDFString.of(a.nm) } : {}),
  })
}

async function docWith(pagesPerAnnots: AnnotSpec[][]): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  pagesPerAnnots.forEach(() => doc.addPage([300, 300]))
  for (let i = 0; i < pagesPerAnnots.length; i++) {
    // Real content on every page: the untouched-stream assertion needs an actual content stream
    doc.getPage(i).drawText('anchor', { x: 20, y: 20, size: 12 })
  }
  for (let i = 0; i < pagesPerAnnots.length; i++) {
    const annots = pagesPerAnnots[i]!
    if (annots.length > 0)
      doc
        .getPage(i)
        .node.set(
          PDFName.of('Annots'),
          doc.context.obj(annots.map((a) => annotSpec(doc.context, a))),
        )
  }
  return doc.save({ useObjectStreams: false })
}

const TWO = [
  [
    { subtype: 'Text', rect: [10, 100, 90, 140], contents: 'first note', nm: 'n-1' },
    { subtype: 'Highlight', rect: [0, 200, 150, 216], contents: 'x' },
  ],
]
const TEXT_FIRST: AnnotMatch = {
  pageIndex: 0,
  subtype: 'Text',
  rect: [10, 100, 90, 140],
  contents: 'first note',
  name: 'n-1',
}

const annotCount = async (bytes: Uint8Array, page = 0): Promise<number> => {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  return doc.getPage(page).node.lookupMaybe(PDFName.of('Annots'), PDFArray)?.size() ?? 0
}

/** Encoded page contents, for the untouched-stream assertion */
const streamBytes = async (bytes: Uint8Array, page = 0): Promise<Uint8Array> => {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  const c = doc.getPage(page).node.get(PDFName.of('Contents'))
  const resolved = doc.context.lookup(c as never)
  // /Contents may be one stream or an array of them — pdf-lib appends every draw as a member
  const parts =
    resolved instanceof PDFArray
      ? Array.from({ length: resolved.size() }, (_, i) =>
          doc.context.lookup(resolved.get(i) as never),
        )
      : [resolved]
  const streams = parts.filter((p): p is PDFStream => p instanceof PDFStream)
  if (streams.length === 0)
    throw new Error(`fixture: unexpected contents shape (${resolved?.constructor?.name})`)
  const total = streams.reduce((n, s) => n + s.getContents().length, 0)
  const out = new Uint8Array(total)
  let at = 0
  for (const s of streams) {
    out.set(s.getContents(), at)
    at += s.getContents().length
  }
  return out
}

const label = (_subtype: string, key: StringKey): string => key

describe('annotationsToRows', () => {
  it('skips Widget, Link and Popup, and keeps addressable rows with an exact identity', () => {
    const rows = annotationsToRows(
      [
        { subtype: 'Widget', rect: [1, 1, 2, 2] },
        { subtype: 'Link', rect: [1, 1, 2, 2] },
        { subtype: 'Popup', rect: [1, 1, 2, 2] },
        {
          subtype: 'Text',
          rect: [10, 100, 90, 140],
          contents: '  hello\nworld  ',
          title: 'Ara ',
          id: 'p1_7', // pdf.js ids never enter the identity
        },
      ],
      1,
      label,
    )
    expect(rows).toHaveLength(1)
    expect(rows[0]!.match).toEqual({
      pageIndex: 1,
      subtype: 'Text',
      rect: [10, 100, 90, 140],
      contents: 'hello\nworld',
      name: '',
    })
    expect(rows[0]!.excerpt).toBe('Ara: hello world')
    expect(rows[0]!.label).toBe('drawNote') // Text annots read as notes, the same key the tools use
    expect(rows[0]!.saved).toBe(true)
  })

  it('labels known subtypes through the strings map and trims long excerpts', () => {
    const rows = annotationsToRows(
      [
        { subtype: 'Highlight', rect: [0, 0, 10, 10] },
        { subtype: 'Stamp', rect: [0, 0, 10, 10], contents: 'x'.repeat(200) },
      ],
      0,
      label,
    )
    expect(rows[0]!.label).toBe('highlight')
    // Unknown subtypes keep their raw PDF subtype name, as viewers do
    expect(rows[1]!.label).toBe('Stamp')
    expect(rows[1]!.match.subtype).toBe('Stamp')
    // excerpt of 200 x's trims to 80 chars plus the ellipsis
    expect(rows[1]!.excerpt).toHaveLength(81)
  })
})

describe('deleteAnnotsBytes', () => {
  it('removes exactly the target annotation and keeps the rest whole', async () => {
    const bytes = await docWith(TWO)
    const out = await deleteAnnotsBytes(bytes, [TEXT_FIRST])

    const doc = await PDFDocument.load(out, { updateMetadata: false })
    const annots = doc.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    expect(annots?.size()).toBe(1)
    const survivor = annots?.lookup(0, PDFDict)
    expect(survivor?.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText()).toBe('Highlight')
    expect(survivor?.lookupMaybe(PDFName.of('Rect'), PDFArray)?.size()).toBe(4)
  })

  it('leaves the page content stream identical', async () => {
    const bytes = await docWith(TWO)
    const out = await deleteAnnotsBytes(bytes, [TEXT_FIRST])
    expect(Buffer.from(await streamBytes(out)).equals(Buffer.from(await streamBytes(bytes)))).toBe(
      true,
    )
  })

  it('refuses an unknown match and writes nothing', async () => {
    const bytes = await docWith(TWO)
    await expect(
      deleteAnnotsBytes(bytes, [{ ...TEXT_FIRST, contents: 'different' }]),
    ).rejects.toThrow(/not found/)
    expect(await annotCount(bytes)).toBe(2)
  })

  it('refuses an ambiguous match: identical annots share the identity', async () => {
    const bytes = await docWith([
      [
        { subtype: 'Text', rect: [10, 100, 90, 140], contents: 'dupe' },
        { subtype: 'Text', rect: [10, 100, 90, 140], contents: 'dupe' },
      ],
    ])
    const target: AnnotMatch = {
      pageIndex: 0,
      subtype: 'Text',
      rect: [10, 100, 90, 140],
      contents: 'dupe',
      name: '',
    }
    await expect(deleteAnnotsBytes(bytes, [target])).rejects.toThrow(/ambiguous/)
    expect(await annotCount(bytes)).toBe(2)
  })

  it('keeps widgets: they are fields, not the targeted comment', async () => {
    const bytes = await docWith([
      [
        { subtype: 'Text', rect: [10, 100, 90, 140], contents: 'note' },
        { subtype: 'Widget', rect: [10, 100, 90, 140], contents: '' },
      ],
    ])
    const out = await deleteAnnotsBytes(bytes, [
      { pageIndex: 0, subtype: 'Text', rect: [10, 100, 90, 140], contents: 'note', name: '' },
    ])
    expect(await annotCount(out)).toBe(1)
    const doc = await PDFDocument.load(out, { updateMetadata: false })
    const annots = doc.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    expect(
      annots?.lookup(0, PDFDict)?.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText(),
    ).toBe('Widget')
  })
})
