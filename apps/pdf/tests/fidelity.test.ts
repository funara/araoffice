import { describe, expect, it } from 'vitest'
import { PDFDocument, PDFName, PDFNumber, PDFString } from 'pdf-lib'
import { applySaveRequest } from '../src/main/save-pdf'
import type { SavePdfRequest } from '../src/shared/ipc'

/**
 * The PDF save fidelity contract (AGENTS.md: "opening/editing/saving ... PDF must preserve
 * untouched content, styling, metadata, and raw bytes").
 *
 * pdf-lib re-serializes the whole object graph, so the "raw bytes" half of that rule does not
 * hold for PDF today. This suite pins down exactly which half does, so the gap is a known,
 * tested quantity rather than a surprise — and so a future incremental-save path has a target
 * to flip. Do not "fix" a failing assertion here by loosening it; the expectations below are
 * the contract.
 */

const request = (over: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path: '/tmp/test.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  ...over,
})

/** What ARA does on every save: parse, then write the whole file back out.
   Bytes stay plain Uint8Array: pdf-lib's own validation rejects a foreign-realm Buffer here. */
const roundTrip = async (bytes: Uint8Array): Promise<Uint8Array> => {
  const doc = await PDFDocument.load(new Uint8Array(bytes), { updateMetadata: false })
  return doc.save({ useObjectStreams: false })
}

const latin = (b: Uint8Array) => Buffer.from(b).toString('latin1')

const sameBytes = (a: Uint8Array, b: Uint8Array) =>
  a.length === b.length && a.every((v, i) => v === b[i])

const concat = (a: Uint8Array, b: Uint8Array): Uint8Array => {
  const out = new Uint8Array(a.length + b.length)
  out.set(a, 0)
  out.set(b, a.length)
  return out
}

/**
 * Append a hand-written incremental update, the way a signing or annotation tool would: the new
 * object plus an xref section and a trailer whose /Prev points at the original xref. This is the
 * shape of every revision history and every signed PDF.
 */
async function appendRevision(orig: Uint8Array, objNum: number): Promise<Uint8Array> {
  const doc = await PDFDocument.load(new Uint8Array(orig), { updateMetadata: false })
  const prev = Number(latin(orig).match(/startxref\s+(\d+)\s+%%EOF\s*$/)?.[1])
  if (!prev) throw new Error('fixture: no startxref in the base file')

  const body = `${objNum} 0 obj\n<< /Rev2Marker (REVISION_TWO) >>\nendobj\n`
  const at = orig.length
  const xrefAt = at + body.length
  const xref =
    `xref\n0 1\n0000000000 65535 f \n${objNum} 1\n${String(at).padStart(10, '0')} 00000 n \n` +
    `trailer\n<< /Size ${objNum + 1} /Root ${doc.context.trailerInfo.Root} ` +
    `/Info ${doc.context.trailerInfo.Info} /Prev ${prev} >>\nstartxref\n${xrefAt}\n%%EOF\n`
  // ASCII only, so UTF-8 encoding gives the exact bytes the PDF syntax needs
  return concat(orig, new TextEncoder().encode(body + xref))
}

/** A base document with a free object number to hang the revision's object on. */
async function seedDoc(): Promise<{ bytes: Uint8Array; freeObjNum: number }> {
  const doc = await PDFDocument.create()
  doc.addPage([200, 200])
  const bytes = await doc.save({ useObjectStreams: false })
  const used = doc.context.enumerateIndirectObjects().map(([ref]) => ref.objectNumber)
  return { bytes, freeObjNum: Math.max(...used) + 1 }
}

describe('PDF fidelity: what a save preserves', () => {
  it('keeps a foreign object graph reachable from the catalog', async () => {
    // The case that matters: a producer-specific structure ARA knows nothing about. If this
    // were dropped, opening a document in ARA and saving it would silently strip another
    // application's data — the exact failure the fidelity rule exists to prevent.
    const doc = await PDFDocument.create()
    doc.addPage([200, 200])
    const nested = doc.context.register(doc.context.obj({ Marker: PDFString.of('ORPHAN_OBJ') }))
    doc.catalog.set(PDFName.of('ARAForeign'), doc.context.obj({ Nested: { Deep: nested } }))

    const out = latin(await roundTrip(await doc.save({ useObjectStreams: false })))
    expect(out).toContain('ARAForeign')
    expect(out).toContain('ORPHAN_OBJ')
  })

  it('keeps foreign keys on a page dictionary', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([200, 200])
    page.node.set(PDFName.of('ARAExtraPageKey'), PDFNumber.of(42))

    const out = latin(await roundTrip(await doc.save({ useObjectStreams: false })))
    expect(out).toContain('ARAExtraPageKey')
    expect(out).toContain('42')
  })

  it('keeps /PieceInfo private application data', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([200, 200])
    doc.catalog.set(
      PDFName.of('PieceInfo'),
      doc.context.obj({ ARA: { Private: PDFString.of('PRIVATE_APP_DATA') } }),
    )

    expect(latin(await roundTrip(await doc.save({ useObjectStreams: false })))).toContain(
      'PRIVATE_APP_DATA',
    )
  })

  it('keeps an unreferenced object', async () => {
    // Garbage-collecting this would be defensible, but Acrobat does not, and "only ever add"
    // is the safer default while ARA cannot tell a stale object from one reached through a
    // structure it does not parse.
    const doc = await PDFDocument.create()
    doc.addPage([100, 100])
    doc.context.register(doc.context.obj({ Marker: PDFString.of('UNREFERENCED') }))

    expect(latin(await roundTrip(await doc.save({ useObjectStreams: false })))).toContain(
      'UNREFERENCED',
    )
  })

  it('keeps an incremental revision alive through an unrelated edit', async () => {
    const { bytes, freeObjNum } = await seedDoc()
    const rev2 = await appendRevision(bytes, freeObjNum)

    const out = latin(
      await applySaveRequest(
        new Uint8Array(rev2),
        request({
          markups: [
            {
              pageIndex: 0,
              type: 'highlight',
              color: [1, 1, 0],
              quads: [[10, 10, 60, 10, 10, 0, 60, 0]],
            },
          ],
        }),
      ),
    )
    // The revision's content survives an edit, so data is never lost by rewriting.
    expect(out).toContain('REVISION_TWO')
  })
})

describe('PDF fidelity: what a save does not preserve', () => {
  // The known gap. Each of these is a user-visible consequence, not an implementation detail.

  it('is byte-stable on a document ARA itself wrote without object streams', async () => {
    // The narrow good news, pinned so a future change cannot silently regress it: re-saving a
    // plain-serialized file is idempotent. This is what the existing save tests rely on, and it
    // is why the gap below is about foreign documents rather than about saving in general.
    const { bytes } = await seedDoc()
    expect(sameBytes(await roundTrip(bytes), bytes)).toBe(true)
  })

  it('rewrites a document written with object streams, changing its bytes', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([200, 200])
    const compact = await doc.save({ useObjectStreams: true })
    expect(latin(compact)).toContain('/ObjStm')

    const out = await roundTrip(compact)
    // Consequence: most real-world PDFs are written with object streams, so the common case is
    // a size increase on every save that is never described as byte-preserving. Measured ~50%
    // on small files; the ratio falls as the document grows.
    expect(latin(out)).not.toContain('/ObjStm')
    expect(out.length).toBeGreaterThan(compact.length)
    expect(sameBytes(out, compact)).toBe(false)
  })

  it('collapses the revision chain, dropping /Prev', async () => {
    const { bytes, freeObjNum } = await seedDoc()
    const rev2 = await appendRevision(bytes, freeObjNum)
    expect(latin(rev2)).toMatch(/\/Prev\b/)

    const out = await roundTrip(rev2)
    // Consequence: revision history is gone. A signed document also loses its ByteRange
    // coverage, which is why incremental save is the prerequisite for digital signatures and
    // is deliberately deferred rather than approximated.
    expect(latin(out)).not.toMatch(/\/Prev\b/)
    expect(latin(out)).toContain('REVISION_TWO')
    expect(sameBytes(out, rev2)).toBe(false)
  })
})
