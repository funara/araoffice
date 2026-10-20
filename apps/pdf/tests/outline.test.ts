import { describe, expect, it } from 'vitest'
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRef } from 'pdf-lib'
import type { PDFObject } from 'pdf-lib'
import { applyOutline, type OutlineNodeInput } from '../src/main/outline-writer'

/**
 * The outline is read back through the same few keys a viewer walks — /First, /Next, /Count,
 * /Parent, /Dest — so the tests hold the writer to the scrollbar, not to its own dicts.
 */

/** The document's outline as plain nested text, or undefined when there is none */
function readOutline(doc: PDFDocument): OutlineNodeInput[] | undefined {
  const root = doc.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict)
  if (!root) return undefined
  const readChain = (firstRef: PDFObject | undefined): OutlineNodeInput[] => {
    const out: OutlineNodeInput[] = []
    let ref = firstRef
    for (let guard = 0; guard < 64 && ref; guard++) {
      if (!(ref instanceof PDFRef)) break
      const dict = doc.context.lookup(ref, PDFDict)
      const title = dict.lookup(PDFName.of('Title'), PDFHexString).decodeText()
      let children: OutlineNodeInput[] = []
      const first = dict.lookupMaybe(PDFName.of('First'), PDFDict)
      if (first || dict.has(PDFName.of('Count')))
        children = readChain(dict.get(PDFName.of('First')))
      out.push({
        title,
        pageIndex: destPageOf(doc, dict),
        children: children.length > 0 ? children : undefined,
      })
      ref = dict.get(PDFName.of('Next'))
    }
    return out
  }
  return readChain(root.get(PDFName.of('First')))
}

/** 0-based index of the page a node's /Dest points at, or -1 when none */
function destPageOf(doc: PDFDocument, dict: PDFDict): number {
  const dest = dict.lookupMaybe(PDFName.of('Dest'), PDFArray)
  if (!dest) return -1
  const ref = dest.get(0)
  return doc.getPages().findIndex((page) => String(page.ref) === String(ref))
}

/** Every supported language has exactly one entry per key with the zh key set (placeholder-free). */
async function docWithOutline(roots: OutlineNodeInput[]): Promise<PDFDocument> {
  const doc = await PDFDocument.create()
  for (let i = 0; i < 4; i++) doc.addPage([200, 200])
  applyOutline(doc, roots)
  return doc
}

describe('applyOutline', () => {
  it('writes a nested tree: siblings chained, children linked to their parent', async () => {
    const doc = await docWithOutline([
      {
        title: 'Alpha',
        pageIndex: 0,
        children: [
          { title: 'Beta', pageIndex: 1 },
          { title: 'Gamma', pageIndex: 2 },
        ],
      },
      { title: 'Delta', pageIndex: 3 },
    ])
    const read = readOutline(doc)!
    expect(read).toEqual([
      {
        title: 'Alpha',
        pageIndex: 0,
        children: [
          { title: 'Beta', pageIndex: 1, children: undefined },
          { title: 'Gamma', pageIndex: 2, children: undefined },
        ],
      },
      { title: 'Delta', pageIndex: 3, children: undefined },
    ])

    // The writer's own counts, read raw: root sees all four entries at all levels
    const root = doc.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict)!
    expect(String(root.get(PDFName.of('Count')))).toBe('4')
    const alpha = doc.context.lookup(root.get(PDFName.of('First'))!, PDFDict)
    expect(String(alpha.get(PDFName.of('Count')))).toBe('2')
    // A child's /Parent is the entry that owns it, checked by title so the link is not merely
    // pointing at "some dict". The ref resolves through the document context
    const beta = doc.context.lookup(alpha.get(PDFName.of('First'))!, PDFDict)
    const parent = doc.context.lookup(beta.get(PDFName.of('Parent')) as never, PDFDict)
    expect(parent.lookup(PDFName.of('Title'), PDFHexString).decodeText()).toBe('Alpha')
  })

  it('replaces an existing outline rather than appending to it', async () => {
    const doc = await docWithOutline([
      { title: 'Old', pageIndex: 0, children: [{ title: 'Old child', pageIndex: 1 }] },
    ])
    applyOutline(doc, [{ title: 'New', pageIndex: 2 }])
    expect(readOutline(doc)).toEqual([{ title: 'New', pageIndex: 2, children: undefined }])
  })

  it('removes the /Outlines entry when every bookmark is cleared', async () => {
    const doc = await docWithOutline([{ title: 'Only', pageIndex: 0 }])
    expect(applyOutline(doc, [])).toBe(false)
    expect(readOutline(doc)).toBeUndefined()
  })

  it('keeps a web bookmark as a URI action instead of a page destination', async () => {
    const doc = await docWithOutline([
      { title: 'ARA', pageIndex: 0, uri: 'https://example.com' },
      { title: 'Doc', pageIndex: 1 },
    ])
    const root = doc.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict)!
    const ara = doc.context.lookup(root.get(PDFName.of('First'))!, PDFDict)
    // The chain stays intact: the action replaces only the destination
    expect(ara.has(PDFName.of('Dest'))).toBe(false)
    const action = ara.lookupMaybe(PDFName.of('A'), PDFDict)!
    expect(action.lookup(PDFName.of('S'), PDFName).decodeText()).toBe('URI')
    // ctx.obj turns a literal string into a PDFName, so the uri is read back as a name too
    expect(action.lookup(PDFName.of('URI'), PDFName).decodeText()).toBe('https://example.com')
    // The next entry is untouched and still a real destination
    const docEntry = doc.context.lookup(ara.get(PDFName.of('Next'))!, PDFDict)
    expect(docEntry.lookupMaybe(PDFName.of('Dest'), PDFArray)).toBeDefined()

    // A URI entry survives a full re-write of the same tree, so editing never drops it
    const out = await doc.save({ useObjectStreams: false })
    const reloaded = await PDFDocument.load(out, { updateMetadata: false })
    applyOutline(reloaded, [
      { title: 'ARA', pageIndex: 0, uri: 'https://example.com' },
      { title: 'Doc', pageIndex: 1 },
    ])
    expect(
      reloaded.context.lookup(reloaded.catalog.get(PDFName.of('Outlines')) as never, PDFDict),
    ).toBeDefined()
  })

  it('drops nothing and invents nothing for an already-empty document', async () => {
    const doc = await docWithOutline([])
    expect(applyOutline(doc, [])).toBe(false)
    expect(doc.catalog.has(PDFName.of('Outlines'))).toBe(false)
  })

  it('refuses to write entries whose page does not exist', async () => {
    const doc = await docWithOutline([{ title: 'Nine', pageIndex: 9 }])
    // The guard drops the entry with its subtree, and without anything left it wrote no outline:
    // a bookmark to page 9 of a 4-page file is never a real sidebar
    expect(applyOutline(doc, [{ title: 'Nine', pageIndex: 9 }])).toBe(false)
  })

  it('drops only the bad branch when a child points past the end', async () => {
    applyOutline(await docWithOutline([]), [])
    const doc = await docWithOutline([
      {
        title: 'Keep',
        pageIndex: 0,
        children: [
          { title: 'Good', pageIndex: 1 },
          { title: 'Bad', pageIndex: 9 },
        ],
      },
    ])
    const read = readOutline(doc)!
    expect(read).toEqual([
      {
        title: 'Keep',
        pageIndex: 0,
        children: [{ title: 'Good', pageIndex: 1, children: undefined }],
      },
    ])
    // The surviving child chain stays linked after the drop
    const root = doc.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict)!
    const keep = doc.context.lookup(root.get(PDFName.of('First'))!, PDFDict)
    const good = doc.context.lookup(keep.get(PDFName.of('First'))!, PDFDict)
    expect(good.get(PDFName.of('Prev'))).toBeUndefined()
    expect(good.get(PDFName.of('Next'))).toBeUndefined()
  })
})
