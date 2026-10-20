import { createHash } from 'node:crypto'
import { mkdtempSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRawStream } from 'pdf-lib'
import {
  applySaveRequest,
  combinePdfsBytes,
  extractPagesBytes,
  insertPdfBytes,
  savePdfToPath,
  splitPdfBytes,
} from '../src/main/save-pdf'
import type { SavePdfRequest } from '../src/shared/ipc'

/** 1x1 red pixel PNG */
const TINY_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

async function makePdf(sizes: [number, number][]): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  for (const size of sizes) doc.addPage(size)
  return doc.save({ useObjectStreams: false })
}

/**
 * A minimal, valid PDF whose page tree is empty. pdf-lib cannot produce one (create() seeds a
 * page and save() always writes at least one), but a damaged or hostile file can look like this,
 * which is the case the combine guard has to survive.
 */
function emptyPdfBytes(): Uint8Array {
  const objs = ['<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Pages /Kids [] /Count 0 >>']
  let body = '%PDF-1.7\n'
  const offsets: number[] = []
  objs.forEach((o, i) => {
    offsets.push(body.length)
    body += `${i + 1} 0 obj\n${o}\nendobj\n`
  })
  const xrefAt = body.length
  body += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n`
  for (const off of offsets) body += `${String(off).padStart(10, '0')} 00000 n \n`
  body += `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xrefAt}\n%%EOF\n`
  // Pure ASCII, so string length is byte length and the xref offsets hold
  return new Uint8Array([...body].map((c) => c.charCodeAt(0)))
}

const request = (over: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path: '/tmp/test.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  ...over,
})

function pageAnnots(doc: PDFDocument, pageIndex: number): PDFDict[] {
  const annots = doc.getPage(pageIndex).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (!annots) return []
  return Array.from({ length: annots.size() }, (_, i) => annots.lookup(i, PDFDict))
}

const subtypeOf = (annot: PDFDict) => annot.lookup(PDFName.of('Subtype'), PDFName).decodeText()

/** The operator text of an annotation's /AP /N appearance stream (uncompressed: pdf-lib writes
    these with no filter), so a test can assert what the appearance actually draws. */
function apOps(doc: PDFDocument, annot: PDFDict): string {
  const ref = annot.lookup(PDFName.of('AP'), PDFDict).get(PDFName.of('N'))
  const stream = doc.context.lookup(ref!) as PDFRawStream
  return Buffer.from(stream.contents).toString('latin1')
}

describe('extractPagesBytes', () => {
  it('extracts the requested pages in the given order', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const out = await PDFDocument.load(await extractPagesBytes(bytes, [2, 0]))
    expect(out.getPageCount()).toBe(2)
    expect(out.getPage(0).getWidth()).toBe(300)
    expect(out.getPage(1).getWidth()).toBe(100)
  })

  it('silently drops out-of-range page indices', async () => {
    const bytes = await makePdf([[100, 100]])
    const out = await PDFDocument.load(await extractPagesBytes(bytes, [-1, 0, 5]))
    expect(out.getPageCount()).toBe(1)
  })
})

describe('insertPdfBytes', () => {
  it('inserts all pages at the front when afterPageIndex is -1', async () => {
    const dst = await makePdf([[100, 100]])
    const src = await makePdf([
      [200, 200],
      [300, 300],
    ])
    const { merged, count } = await insertPdfBytes(dst, src, -1)
    expect(count).toBe(2)
    const out = await PDFDocument.load(merged)
    expect(out.getPageCount()).toBe(3)
    expect(out.getPage(0).getWidth()).toBe(200)
    expect(out.getPage(2).getWidth()).toBe(100)
  })

  it('clamps an out-of-range afterPageIndex to the document end', async () => {
    const dst = await makePdf([[100, 100]])
    const src = await makePdf([[200, 200]])
    const { merged } = await insertPdfBytes(dst, src, 99)
    const out = await PDFDocument.load(merged)
    expect(out.getPage(1).getWidth()).toBe(200)
  })
})

describe('savePdfToPath', () => {
  const sha256 = (path: string) => createHash('sha256').update(readFileSync(path)).digest('hex')
  const highlight = {
    pageIndex: 0,
    type: 'highlight' as const,
    color: [1, 0.87, 0.35] as [number, number, number],
    quads: [[10, 100, 60, 100, 10, 88, 60, 88]],
  }

  it('Save As writes the edits to the target only and never mutates the source', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'gen-pdf-'))
    const src = join(dir, 'original.pdf')
    const dst = join(dir, 'copy.pdf')
    writeFileSync(src, await makePdf([[612, 792]]))
    const srcHash = sha256(src)
    const srcInode = statSync(src).ino

    await savePdfToPath(src, dst, request({ path: src, targetPath: dst, markups: [highlight] }))

    // Source: same inode, same bytes
    expect(sha256(src)).toBe(srcHash)
    expect(statSync(src).ino).toBe(srcInode)
    // Target: valid PDF containing the new annotation
    const out = await PDFDocument.load(new Uint8Array(readFileSync(dst)))
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Highlight'])
    // No temp files left behind
    expect(readdirSync(dir).sort()).toEqual(['copy.pdf', 'original.pdf'])
  })

  it('in-place save (target === source) replaces the file atomically', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'gen-pdf-'))
    const src = join(dir, 'doc.pdf')
    writeFileSync(src, await makePdf([[612, 792]]))

    await savePdfToPath(src, src, request({ path: src, markups: [highlight] }))

    const out = await PDFDocument.load(new Uint8Array(readFileSync(src)))
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Highlight'])
    expect(readdirSync(dir)).toEqual(['doc.pdf'])
  })

  it('a failed save leaves the source and target untouched and cleans up temp files', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'gen-pdf-'))
    const src = join(dir, 'original.pdf')
    writeFileSync(src, await makePdf([[612, 792]]))
    const srcHash = sha256(src)

    // Apply failure (unknown form field): nothing may be written anywhere
    await expect(
      savePdfToPath(
        src,
        join(dir, 'copy.pdf'),
        request({ path: src, formValues: [{ name: 'missing', kind: 'text', value: 'x' }] }),
      ),
    ).rejects.toThrow()
    expect(sha256(src)).toBe(srcHash)
    expect(readdirSync(dir)).toEqual(['original.pdf'])

    // Write failure (target directory does not exist): source intact, temp cleaned up
    await expect(
      savePdfToPath(src, join(dir, 'no-such-dir', 'copy.pdf'), request({ path: src })),
    ).rejects.toThrow()
    expect(sha256(src)).toBe(srcHash)
    expect(readdirSync(dir)).toEqual(['original.pdf'])
  })
})

describe('applySaveRequest', () => {
  it('applies page rotation deltas on top of the existing rotation', async () => {
    const bytes = await makePdf([[100, 100]])
    const saved = await applySaveRequest(
      bytes,
      request({ rotations: [{ pageIndex: 0, delta: 90 }] }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getPage(0).getRotation().angle).toBe(90)
  })

  it('writes markup annotations with an appearance stream', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        markups: [
          {
            pageIndex: 0,
            type: 'highlight',
            color: [1, 0.87, 0.35],
            quads: [[10, 100, 60, 100, 10, 88, 60, 88]],
          },
          {
            pageIndex: 0,
            type: 'underline',
            color: [0.17, 0.4, 1],
            quads: [[10, 80, 60, 80, 10, 68, 60, 68]],
          },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Highlight', 'Underline'])
    // Every markup gets a hand-written /AP /N appearance
    for (const a of annots) {
      expect(a.lookup(PDFName.of('AP'), PDFDict).has(PDFName.of('N'))).toBe(true)
    }
  })

  it('writes note and shape drawing annotations', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        drawings: [
          { kind: 'note', pageIndex: 0, color: [1, 0, 0], at: [50, 700], contents: 'hello note' },
          {
            kind: 'ink',
            pageIndex: 0,
            color: [0, 0, 1],
            width: 2,
            paths: [[10, 10, 20, 20, 30, 15]],
          },
          { kind: 'rect', pageIndex: 0, color: [0, 1, 0], width: 1, rect: [40, 40, 90, 80] },
          {
            kind: 'arrow',
            pageIndex: 0,
            color: [0, 0, 0],
            width: 2,
            from: [100, 100],
            to: [200, 150],
          },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Text', 'Ink', 'Square', 'Line'])
  })

  it('writes a typewriter box as a FreeText annotation carrying its text', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        drawings: [
          {
            kind: 'freetext',
            pageIndex: 0,
            image: TINY_PNG,
            rect: [80, 560, 220, 580],
            contents: 'typed 中文 note',
          },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['FreeText'])
    // The rasterized glyphs are the appearance; the text stays a real comment viewers can list
    expect(apOps(out, annots[0]!)).toContain('/Im0 Do')
    expect(annots[0]!.lookup(PDFName.of('Contents'), PDFHexString).decodeText()).toBe(
      'typed 中文 note',
    )
  })

  it('distinguishes a stamp from typewriter text by subtype', async () => {
    // Both are image-backed, so the subtype is the only thing telling a viewer whether to treat
    // the annotation as a signature/stamp or as editable text
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        drawings: [
          { kind: 'image', pageIndex: 0, image: TINY_PNG, rect: [10, 10, 40, 40] },
          {
            kind: 'freetext',
            pageIndex: 0,
            image: TINY_PNG,
            rect: [60, 60, 120, 80],
            contents: 'note',
          },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Stamp', 'FreeText'])
    // A stamp carries no /Contents; the typewriter box does
    expect(annots[0]!.has(PDFName.of('Contents'))).toBe(false)
    expect(annots[1]!.has(PDFName.of('Contents'))).toBe(true)
  })

  it('draws squiggly as a wave and underline as a straight segment', async () => {
    const quad = [[10, 100, 60, 100, 10, 88, 60, 88]]
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        markups: [
          { pageIndex: 0, type: 'squiggly', color: [0.86, 0.22, 0.18], quads: quad },
          { pageIndex: 0, type: 'underline', color: [0.17, 0.4, 1], quads: quad },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Squiggly', 'Underline'])
    // The squiggle is the only difference between the two, so this pins the wave itself
    expect(apOps(out, annots[0]!)).toContain(' c')
    expect(apOps(out, annots[1]!)).not.toContain(' c')
  })

  it('writes a polyline and a polygon with their vertices', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        drawings: [
          {
            kind: 'polyline',
            pageIndex: 0,
            color: [1, 0, 0],
            width: 1,
            points: [10, 10, 50, 40, 90, 20],
          },
          {
            kind: 'polygon',
            pageIndex: 0,
            color: [0, 0, 1],
            width: 2,
            points: [100, 100, 160, 100, 130, 140],
          },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['PolyLine', 'Polygon'])
    expect(String(annots[0]!.lookup(PDFName.of('Vertices'), PDFArray))).toBe(
      '[ 10 10 50 40 90 20 ]',
    )
    // The closed shape joins its last point back to the first; the open one only strokes
    const polylineOps = apOps(out, annots[0]!).split('\n')
    const polygonOps = apOps(out, annots[1]!).split('\n')
    expect(polylineOps).toContain('S')
    expect(polygonOps).toContain('s')
    expect(polylineOps).not.toContain('s')
  })

  it('ignores malformed polyline points instead of writing a broken annotation', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        drawings: [
          { kind: 'polyline', pageIndex: 0, color: [1, 0, 0], width: 1, points: [10, 10] },
          { kind: 'polygon', pageIndex: 0, color: [1, 0, 0], width: 1, points: [10, 10, 50, 40] },
          { kind: 'polyline', pageIndex: 0, color: [1, 0, 0], width: 1, points: [1, 2, 3] },
        ],
      }),
    )
    expect(pageAnnots(await PDFDocument.load(saved), 0)).toHaveLength(0)
  })

  it('writes a caret annotation carrying its comment', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        drawings: [
          {
            kind: 'caret',
            pageIndex: 0,
            color: [1, 0, 0],
            at: [80, 600],
            size: 12,
            contents: 'insert here',
          },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Caret'])
    expect(apOps(out, annots[0]!)).toContain(' l S')
    expect(annots[0]!.lookup(PDFName.of('Contents'), PDFHexString).decodeText()).toBe('insert here')
  })

  it('ignores markups and drawings addressing missing pages', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        markups: [
          { pageIndex: 9, type: 'highlight', color: [1, 1, 0], quads: [[0, 1, 1, 1, 0, 0, 1, 0]] },
        ],
        drawings: [{ kind: 'note', pageIndex: 9, color: [1, 0, 0], at: [0, 0], contents: 'x' }],
      }),
    )
    expect(pageAnnots(await PDFDocument.load(saved), 0)).toHaveLength(0)
  })

  it('writes an image drawing as a Stamp annotation with an image appearance', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        drawings: [{ kind: 'image', pageIndex: 0, image: TINY_PNG, rect: [100, 500, 300, 600] }],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Stamp'])
    const rect = annots[0]!.lookup(PDFName.of('Rect'), PDFArray)
    expect(String(rect)).toBe('[ 100 500 300 600 ]')
    expect(annots[0]!.lookup(PDFName.of('AP'), PDFDict).has(PDFName.of('N'))).toBe(true)
  })

  it('counter-rotates the image appearance on rotated pages', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        rotations: [{ pageIndex: 0, delta: 90 }],
        drawings: [{ kind: 'image', pageIndex: 0, image: TINY_PNG, rect: [100, 500, 300, 600] }],
      }),
    )
    const out = await PDFDocument.load(saved)
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Stamp'])
  })

  it('embeds PNG stamps without failing', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await applySaveRequest(
      bytes,
      request({
        stamps: [{ pageIndex: 0, image: TINY_PNG, rect: [0, 0, 612, 792], opacity: 0.2 }],
      }),
    )
    expect((await PDFDocument.load(saved)).getPageCount()).toBe(1)
  })

  it('applies metadata and splits keywords on mixed separators', async () => {
    const bytes = await makePdf([[100, 100]])
    const saved = await applySaveRequest(
      bytes,
      request({ metadata: { title: 'My Title', author: 'Me', keywords: 'a, b；c，d' } }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getTitle()).toBe('My Title')
    expect(out.getAuthor()).toBe('Me')
    expect(out.getKeywords()).toContain('a')
    expect(out.getKeywords()).toContain('d')
  })

  it('deletes pages by original index but never removes the last page', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const saved = await applySaveRequest(bytes, request({ deletedPages: [0, 2] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPageCount()).toBe(1)
    expect(out.getPage(0).getWidth()).toBe(200)

    const savedAll = await applySaveRequest(bytes, request({ deletedPages: [0, 1, 2] }))
    expect((await PDFDocument.load(savedAll)).getPageCount()).toBe(1)
  })

  it('reorders pages by original index', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const saved = await applySaveRequest(bytes, request({ pageOrder: [2, 0, 1] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPages().map((p) => p.getWidth())).toEqual([300, 100, 200])
  })

  it('applies the reorder when combined with deletions', async () => {
    // Regression: pdf-lib's removePage does not invalidate its page cache,
    // so a getPages() call after the deletion loop returns the stale
    // pre-deletion list. applySaveRequest must derive the remaining pages
    // from the pre-deletion snapshot instead of re-reading them.
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const saved = await applySaveRequest(bytes, request({ deletedPages: [1], pageOrder: [2, 0] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPageCount()).toBe(2)
    expect(out.getPages().map((p) => p.getWidth())).toEqual([300, 100])
  })

  it('skips the reorder when deleting every page leaves the guarded last page', async () => {
    // Deleting all pages keeps one via the last-page guard; the order list
    // then matches nothing alive, so the reorder must be skipped safely.
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
    ])
    const saved = await applySaveRequest(
      bytes,
      request({ deletedPages: [0, 1], pageOrder: [1, 0] }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getPageCount()).toBe(1)
  })

  it('fills text fields and checkboxes', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 300])
    const form = doc.getForm()
    form.createTextField('user.name').addToPage(page, { x: 20, y: 200, width: 200, height: 20 })
    form.createCheckBox('user.agree').addToPage(page, { x: 20, y: 150, width: 16, height: 16 })
    const bytes = await doc.save({ useObjectStreams: false })

    const saved = await applySaveRequest(
      bytes,
      request({
        formValues: [
          { name: 'user.name', kind: 'text', value: 'Alice' },
          { name: 'user.agree', kind: 'checkbox', checked: true },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getForm().getTextField('user.name').getText()).toBe('Alice')
    expect(out.getForm().getCheckBox('user.agree').isChecked()).toBe(true)
  })

  it('falls back to NeedAppearances when form values cannot be WinAnsi-encoded', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 300])
    doc.getForm().createTextField('cjk').addToPage(page, { x: 20, y: 200, width: 200, height: 20 })
    const bytes = await doc.save({ useObjectStreams: false })

    const saved = await applySaveRequest(
      bytes,
      request({ formValues: [{ name: 'cjk', kind: 'text', value: '中文测试' }] }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getForm().getTextField('cjk').getText()).toBe('中文测试')
    const needAppearances = out.getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))
    expect(String(needAppearances)).toBe('true')
  })
})

describe('splitPdfBytes', () => {
  it('produces one document per range, in order', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
      [400, 400],
    ])
    const parts = await splitPdfBytes(bytes, [[0, 1], [2], [3]])
    expect(parts).toHaveLength(3)
    const widths = await Promise.all(
      parts.map(async (p) => (await PDFDocument.load(p)).getPages().map((x) => x.getWidth())),
    )
    expect(widths).toEqual([[100, 200], [300], [400]])
  })

  it('returns an empty list for no ranges', async () => {
    expect(await splitPdfBytes(await makePdf([[100, 100]]), [])).toEqual([])
  })
})

describe('combinePdfsBytes', () => {
  const outlineOf = (doc: PDFDocument): PDFDict | undefined =>
    doc.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict)

  /** Walk the /Next chain of the top-level outline items */
  const outlineTitles = (doc: PDFDocument): string[] => {
    const root = outlineOf(doc)
    if (!root) return []
    const titles: string[] = []
    let ref = root.get(PDFName.of('First'))
    while (ref) {
      const item = doc.context.lookup(ref, PDFDict)
      titles.push(item.lookup(PDFName.of('Title'), PDFHexString).decodeText())
      ref = item.get(PDFName.of('Next'))
    }
    return titles
  }

  it('concatenates every source in the given order', async () => {
    const merged = await combinePdfsBytes(
      [
        await makePdf([[100, 100]]),
        await makePdf([
          [200, 200],
          [250, 250],
        ]),
      ],
      null,
    )
    const out = await PDFDocument.load(merged)
    expect(out.getPages().map((p) => p.getWidth())).toEqual([100, 200, 250])
    // No titles asked for, so no outline is invented
    expect(outlineOf(out)).toBeUndefined()
  })

  it('links one bookmark per source to its first page', async () => {
    const merged = await combinePdfsBytes(
      [
        await makePdf([[100, 100]]),
        await makePdf([
          [200, 200],
          [250, 250],
        ]),
      ],
      ['First', 'Second'],
    )
    const out = await PDFDocument.load(merged)
    expect(outlineTitles(out)).toEqual(['First', 'Second'])

    const root = outlineOf(out)!
    expect(String(root.get(PDFName.of('Count')))).toBe('2')
    const first = out.context.lookup(root.get(PDFName.of('First'))!, PDFDict)
    const second = out.context.lookup(root.get(PDFName.of('Last'))!, PDFDict)
    // Each bookmark points at the page its source starts on, and the chain is doubly linked
    expect(first.lookup(PDFName.of('Title'), PDFHexString).decodeText()).toBe('First')
    expect(second.lookup(PDFName.of('Title'), PDFHexString).decodeText()).toBe('Second')
    const destPage = (item: PDFDict) => {
      const dest = item.lookup(PDFName.of('Dest'), PDFArray)
      return out.getPages().findIndex((p) => p.ref === dest.get(0))
    }
    expect(destPage(first)).toBe(0)
    expect(destPage(second)).toBe(1)
    expect(first.get(PDFName.of('Prev'))).toBeUndefined()
    expect(second.get(PDFName.of('Next'))).toBeUndefined()
  })

  it('skips a bookmark whose source contributes no pages', async () => {
    // A page-less source cannot be built with pdf-lib, so the fixture is written by hand.
    // The premise is asserted first: if pdf-lib ever stops loading such a file, fail here
    // rather than let the rest of the test pass for the wrong reason.
    const emptyBytes = emptyPdfBytes()
    expect((await PDFDocument.load(emptyBytes)).getPageCount()).toBe(0)

    const merged = await combinePdfsBytes(
      [emptyBytes, await makePdf([[300, 300]])],
      ['Empty', 'Real'],
    )
    const out = await PDFDocument.load(merged)
    // A page-less source has nothing to point at, so it must not become a bookmark — that would
    // either dangle or land on the next source's first page
    expect(outlineTitles(out)).toEqual(['Real'])
    const root = outlineOf(out)!
    const only = out.context.lookup(root.get(PDFName.of('First'))!, PDFDict)
    const dest = only.lookup(PDFName.of('Dest'), PDFArray)
    expect(out.getPages().findIndex((p) => p.ref === dest.get(0))).toBe(0)
  })
})
