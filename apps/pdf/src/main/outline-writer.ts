import { PDFDict, PDFDocument, PDFHexString, PDFName } from 'pdf-lib'
import type { PDFRef } from 'pdf-lib'

/**
 * The document outline as a tree, the shape the PDF spec means and what a viewer's sidebar
 * expects, so an edit never flattens a document's existing nesting.
 *
 * Fidelity notes:
 * - An entry whose page is out of range is dropped, and with it its whole subtree. The caller
 *   resolves destinations to a page it knows first, so this only ever fires on a document whose
 *   structure changed underneath — it is a guard, not a silent pruning pass.
 * - Removing every bookmark removes the /Outlines entry entirely rather than writing an empty one,
 *   which is what a viewer treats as "no sidebar".
 * - An edit that writes new entries leaves the old entry objects orphaned in the file (the save
 *   path preserves unreferenced objects), so a replaced outline costs space, not behaviour.
 */

export interface OutlineNodeInput {
  title: string
  /** 0-based page index the bookmark jumps to; ignored when uri is set */
  pageIndex: number
  /** External target: written as a /URI action instead of a page destination, so web and
      mailto bookmarks the document already had survive an edit untouched */
  uri?: string
  children?: OutlineNodeInput[]
}

/** Visible descendants of a node, all of them: the writer treats every branch as expanded */
function countOpen(children: OutlineNodeInput[] | undefined): number {
  if (!children?.length) return 0
  return children.reduce((sum, child) => sum + 1 + countOpen(child.children), 0)
}

/**
 * Write the outline tree. Returns the number of top-level entries written — zero means nothing
 * valid was given, and the caller should remove the /Outlines entry instead of pointing at junk.
 */
function writeOutlineTree(pdfDoc: PDFDocument, roots: OutlineNodeInput[]): number {
  const usable = roots.filter(
    (root) => Boolean(root.uri) || (root.pageIndex >= 0 && root.pageIndex < pdfDoc.getPageCount()),
  )
  if (usable.length === 0) return 0
  const ctx = pdfDoc.context
  const outlinesRef = ctx.register(
    ctx.obj({
      Type: 'Outlines',
      Count: usable.reduce((sum, r) => sum + 1 + countOpen(r.children), 0),
    }),
  )

  const writeNode = (node: OutlineNodeInput, parentRef: PDFRef): PDFRef => {
    const ref = ctx.register(
      // A URI bookmark carries an action rather than a destination; the sidebar looks the same
      ctx.obj({
        Title: PDFHexString.fromText(node.title),
        Parent: parentRef,
        ...(node.uri
          ? { A: { S: 'URI', URI: node.uri, Type: 'Action' } }
          : { Dest: [pdfDoc.getPage(node.pageIndex).ref, PDFName.of('Fit')] }),
      }),
    )
    const dict = ctx.lookup(ref, PDFDict)
    const kids = (node.children ?? []).filter(
      (child) =>
        Boolean(child.uri) || (child.pageIndex >= 0 && child.pageIndex < pdfDoc.getPageCount()),
    )
    if (kids.length > 0) {
      const refs = kids.map((kid) => writeNode(kid, ref))
      dict.set(PDFName.of('First'), refs[0]!)
      dict.set(PDFName.of('Last'), refs[refs.length - 1]!)
      dict.set(PDFName.of('Count'), ctx.obj(countOpen(kids)))
      refs.forEach((kidRef, i) => {
        if (i > 0) ctx.lookup(kidRef, PDFDict).set(PDFName.of('Prev'), refs[i - 1]!)
        if (i < refs.length - 1) ctx.lookup(kidRef, PDFDict).set(PDFName.of('Next'), refs[i + 1]!)
      })
    }
    return ref
  }

  const rootRefs = usable.map((root) => writeNode(root, outlinesRef))
  rootRefs.forEach((ref, i) => {
    if (i > 0) ctx.lookup(ref, PDFDict).set(PDFName.of('Prev'), rootRefs[i - 1]!)
    if (i < rootRefs.length - 1) ctx.lookup(ref, PDFDict).set(PDFName.of('Next'), rootRefs[i + 1]!)
  })
  ctx.lookup(outlinesRef, PDFDict).set(PDFName.of('First'), rootRefs[0]!)
  ctx.lookup(outlinesRef, PDFDict).set(PDFName.of('Last'), rootRefs[rootRefs.length - 1]!)
  pdfDoc.catalog.set(PDFName.of('Outlines'), outlinesRef)
  return usable.length
}

/**
 * Apply an outline tree to a document: writes it when entries exist, removes the /Outlines entry
 * when the caller has cleared every bookmark, and leaves an out-of-range structure in the file
 * alone rather than guessing what it pointed at.
 *
 * `true` when the document now has an outline (so the caller reloads and re-reads it).
 */
export function applyOutline(pdfDoc: PDFDocument, roots: OutlineNodeInput[]): boolean {
  if (roots.length === 0) {
    pdfDoc.catalog.delete?.(PDFName.of('Outlines'))
    return false
  }
  return writeOutlineTree(pdfDoc, roots) > 0
}
