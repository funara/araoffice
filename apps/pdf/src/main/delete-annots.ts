import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFString,
  PDFHexString,
} from 'pdf-lib'
import type { PDFObject } from 'pdf-lib'
import type { AnnotMatch } from '../shared/ipc'

/**
 * Annotation deletion for the comments panel: a main-side matcher that removes exactly the
 * annotation the renderer pointed at, refusing instead of guessing.
 *
 * The renderer's pdf.js rows carry { subtype, rect, contents, name } — pdf.js synthesizes
 * unstable ids, so the file-side match uses decoded /Subtype plus the raw /Rect (with a
 * small tolerance for float re-encoding), and only when the /Contents and /NM fields agree.
 * Zero matches is an error; several identical candidates is an ambiguity error — the caller
 * must send something more distinguishing rather than this module guessing.
 */

const RECT_TOL = 0.01

/** Untrusted file data: no class-asserting lookups (pdf-lib's assert throws); resolve + check */
function decodeMaybe(ctx: PDFDocument['context'], value: PDFObject | undefined): string {
  if (!value) return ''
  // A ref only resolves through the context: the same call shape as the a11y module's decodeStr
  const resolved = ctx.lookup(value as never)
  if (resolved instanceof PDFString || resolved instanceof PDFHexString)
    return resolved.decodeText()
  return ''
}

/** Same annotation when every identity field agrees (rect within float tolerance) */
function matches(ctx: PDFDocument['context'], dict: PDFDict, m: AnnotMatch): boolean {
  const subtype = ctx.lookup(dict.get(PDFName.of('Subtype')))
  if (!(subtype instanceof PDFName) || subtype.decodeText() !== m.subtype) return false
  const rect = dict.lookupMaybe(PDFName.of('Rect'), PDFArray)
  if (!rect || rect.size() !== m.rect.length) return false
  for (let i = 0; i < rect.size(); i++) {
    // /Rect entries are numbers (possibly indirect refs) — anything unreadable is not the row
    const num = ctx.lookup(rect.get(i))
    const expected = m.rect[i]
    if (!(num instanceof PDFNumber) || expected === undefined) return false
    if (Math.abs(num.asNumber() - expected) > RECT_TOL) return false
  }
  const contents = decodeMaybe(ctx, dict.get(PDFName.of('Contents')))
  if (contents !== (m.contents || '')) return false
  return decodeMaybe(ctx, dict.get(PDFName.of('NM'))) === (m.name || '')
}

/**
 * Delete the one annotation matching each target. Refuses the whole batch on any unresolved
 * or ambiguous target — a partial application would leave the file inconsistent with what
 * the panel showed the user. Returns nothing on success; the caller saves.
 */
export async function deleteAnnotsBytes(
  bytes: Uint8Array,
  targets: AnnotMatch[],
): Promise<Uint8Array> {
  const pdfDoc = await PDFDocument.load(bytes, { updateMetadata: false })
  const ctx = pdfDoc.context

  for (const m of targets) {
    const page = pdfDoc.getPages()[m.pageIndex]
    if (!page) throw new Error(`pdf: page ${m.pageIndex + 1} does not exist`)
    const annots = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!annots) throw new Error(`pdf: page ${m.pageIndex + 1} has no annotations left`)

    const hits: number[] = []
    for (let i = 0; i < annots.size(); i++) {
      const dict = ctx.lookupMaybe(annots.get(i), PDFDict)
      if (dict && matches(ctx, dict, m)) hits.push(i)
    }
    if (hits.length === 0) throw new Error('pdf: the annotation to remove was not found')
    if (hits.length > 1) {
      throw new Error(
        'pdf: the annotation is ambiguous — identical annotations share this page position',
      )
    }
    for (const hit of hits) annots.remove(hit)
  }

  return pdfDoc.save({ useObjectStreams: false })
}
