import { readFile } from 'node:fs/promises'
import { atomicWriteFile } from '@araoffice/electron-utils'
import {
  PDFArray,
  PDFBool,
  PDFDict,
  PDFDocument,
  PDFDropdown,
  PDFHexString,
  PDFName,
  PDFOptionList,
  PDFObject,
  degrees,
} from 'pdf-lib'
import type { PDFPage, PDFRef } from 'pdf-lib'
import type {
  DrawingInput,
  FormValueInput,
  MarkupInput,
  MarkupType,
  MetadataInput,
  SanitizeReport,
  SanitizeRequest,
  SavePdfRequest,
} from '../shared/ipc'
import { MIN_POLY_POINTS } from '../shared/ipc'
import { applyOutline } from './outline-writer'

const num = (v: number) => Math.round(v * 100) / 100

const quadBounds = (q: number[]) => {
  const xs = [q[0]!, q[2]!, q[4]!, q[6]!]
  const ys = [q[1]!, q[3]!, q[5]!, q[7]!]
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)] as const
}

/**
 * Hand-written appearance stream (/AP /N), so viewers don't self-draw from QuadPoints —
 * Acrobat/Preview/pdfjs all render from AP for consistent results.
 * Highlight uses Multiply blending to mimic a highlighter; underline/strikeout are stroked
 * segments drawn along the "visual bottom edge" (pageRot is the page's final /Rotate;
 * at 90/270 line height runs along the x axis).
 */
/**
 * One markup stroke from (ax,ay) to (bx,by): a straight segment for underline/strikeout, or a
 * wave for squiggly. `h` is the line height, which sets the wave amplitude so the squiggle stays
 * proportional to the text at every page rotation. Works along the line direction, so the four
 * page rotations need no special cases.
 */
function markupLine(
  type: MarkupType,
  ax: number,
  ay: number,
  bx: number,
  by: number,
  h: number,
): string[] {
  if (type !== 'squiggly') return [`${num(ax)} ${num(ay)} m ${num(bx)} ${num(by)} l S`]
  const dx = bx - ax
  const dy = by - ay
  const len = Math.hypot(dx, dy)
  if (len < 0.5) return []
  const amp = Math.max(0.6, h * 0.11)
  // Even half-periods so the wave starts and ends on the baseline
  const half = Math.max(2, 2 * Math.round(len / (amp * 6)))
  const sx = dx / half
  const sy = dy / half
  // Perpendicular unit vector: the wave swings across the line, not along it
  const px = -dy / len
  const py = dx / len
  const ops = [`${num(ax)} ${num(ay)} m`]
  for (let i = 0; i < half; i++) {
    const x0 = ax + sx * i
    const y0 = ay + sy * i
    const x1 = ax + sx * (i + 1)
    const y1 = ay + sy * (i + 1)
    // Control points sit at the crest, 4/3 over so the cubic peaks at the amplitude
    const k = ((i % 2 === 0 ? 1 : -1) * amp * 4) / 3
    const cx1 = x0 + sx / 3 + px * k
    const cy1 = y0 + sy / 3 + py * k
    const cx2 = x1 - sx / 3 + px * k
    const cy2 = y1 - sy / 3 + py * k
    ops.push(`${num(cx1)} ${num(cy1)} ${num(cx2)} ${num(cy2)} ${num(x1)} ${num(y1)} c`)
  }
  ops.push('S')
  return ops
}

function markupAppearance(
  pdfDoc: PDFDocument,
  m: MarkupInput,
  rect: number[],
  pageRot: number,
): ReturnType<typeof pdfDoc.context.stream> {
  const [r, g, b] = m.color
  const ops: string[] = []
  if (m.type === 'highlight') {
    ops.push('/GsM gs', `${r} ${g} ${b} rg`)
    for (const q of m.quads) {
      const [x1, y1, x2, y2] = quadBounds(q)
      ops.push(`${num(x1)} ${num(y1)} ${num(x2 - x1)} ${num(y2 - y1)} re f`)
    }
  } else {
    // squiggly rides the underline position, strikeout runs through the glyphs
    const t = m.type === 'strikeout' ? 0.46 : 0.08
    ops.push(`${r} ${g} ${b} RG`)
    for (const q of m.quads) {
      const [x1, y1, x2, y2] = quadBounds(q)
      const h = pageRot % 180 === 0 ? y2 - y1 : x2 - x1
      ops.push(`${Math.max(0.8, num(h * 0.06))} w`)
      if (pageRot === 90) {
        const x = x2 - h * t
        ops.push(...markupLine(m.type, x, y1, x, y2, h))
      } else if (pageRot === 270) {
        const x = x1 + h * t
        ops.push(...markupLine(m.type, x, y1, x, y2, h))
      } else {
        const y = pageRot === 180 ? y2 - h * t : y1 + h * t
        ops.push(...markupLine(m.type, x1, y, x2, y, h))
      }
    }
  }
  return pdfDoc.context.stream(ops.join('\n'), {
    Type: 'XObject',
    Subtype: 'Form',
    BBox: rect,
    Resources:
      m.type === 'highlight'
        ? { ExtGState: { GsM: { Type: 'ExtGState', BM: 'Multiply', ca: 1 } } }
        : {},
  })
}

const SUBTYPE: Record<MarkupInput['type'], string> = {
  highlight: 'Highlight',
  underline: 'Underline',
  strikeout: 'StrikeOut',
  squiggly: 'Squiggly',
}

function addMarkup(pdfDoc: PDFDocument, page: PDFPage, m: MarkupInput): void {
  const xs = m.quads.flatMap((q) => [q[0]!, q[2]!, q[4]!, q[6]!])
  const ys = m.quads.flatMap((q) => [q[1]!, q[3]!, q[5]!, q[7]!])
  const rect = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
  const pageRot = ((page.getRotation().angle % 360) + 360) % 360
  const apRef = pdfDoc.context.register(markupAppearance(pdfDoc, m, rect, pageRot))
  const annot = pdfDoc.context.obj({
    Type: 'Annot',
    Subtype: SUBTYPE[m.type],
    Rect: rect,
    QuadPoints: m.quads.flat(),
    C: m.color,
    F: 4, // print
    T: 'ARAOffice',
    P: page.ref,
    AP: { N: apRef },
  })
  appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
}

function appendAnnot(pdfDoc: PDFDocument, page: PDFPage, annotRef: PDFRef): void {
  const existing = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (existing) {
    existing.push(annotRef)
  } else {
    page.node.set(PDFName.of('Annots'), pdfDoc.context.obj([annotRef]))
  }
}

/** 4-segment Bezier approximation of an ellipse */
function ellipseOps(x1: number, y1: number, x2: number, y2: number): string[] {
  const k = 0.5522847
  const cx = (x1 + x2) / 2
  const cy = (y1 + y2) / 2
  const rx = Math.abs(x2 - x1) / 2
  const ry = Math.abs(y2 - y1) / 2
  const n = num
  return [
    `${n(cx + rx)} ${n(cy)} m`,
    `${n(cx + rx)} ${n(cy + k * ry)} ${n(cx + k * rx)} ${n(cy + ry)} ${n(cx)} ${n(cy + ry)} c`,
    `${n(cx - k * rx)} ${n(cy + ry)} ${n(cx - rx)} ${n(cy + k * ry)} ${n(cx - rx)} ${n(cy)} c`,
    `${n(cx - rx)} ${n(cy - k * ry)} ${n(cx - k * rx)} ${n(cy - ry)} ${n(cx)} ${n(cy - ry)} c`,
    `${n(cx + k * rx)} ${n(cy - ry)} ${n(cx + rx)} ${n(cy - k * ry)} ${n(cx + rx)} ${n(cy)} c`,
    'S',
  ]
}

/**
 * Image signature/stamp: a Stamp annotation whose appearance stream draws the embedded PNG.
 * The image is counter-rotated against the page's final /Rotate (viewers rotate annotation
 * appearances with the page), so it displays upright — matching the renderer preview.
 *
 * Subtype is Stamp for signatures and stamps, FreeText for typewriter text: the bitmap is what
 * keeps CJK working, while FreeText keeps the annotation a text comment that viewers list,
 * sort and search.
 */
async function addImageAnnot(
  pdfDoc: PDFDocument,
  page: PDFPage,
  d: { image: string; rect: [number, number, number, number]; contents?: string },
  subtype: 'Stamp' | 'FreeText',
): Promise<void> {
  const png = await pdfDoc.embedPng(d.image)
  const [x1, y1, x2, y2] = d.rect
  const rw = x2 - x1
  const rh = y2 - y1
  const rot = ((page.getRotation().angle % 360) + 360) % 360
  // cm matrix mapping the image unit square into the BBox, pre-counter-rotated for the page
  const cm =
    rot === 90
      ? `0 ${num(rh)} ${num(-rw)} 0 ${num(rw)} 0`
      : rot === 180
        ? `${num(-rw)} 0 0 ${num(-rh)} ${num(rw)} ${num(rh)}`
        : rot === 270
          ? `0 ${num(-rh)} ${num(rw)} 0 0 ${num(rh)}`
          : `${num(rw)} 0 0 ${num(rh)} 0 0`
  const ap = pdfDoc.context.stream(`q ${cm} cm /Im0 Do Q`, {
    Type: 'XObject',
    Subtype: 'Form',
    BBox: [0, 0, num(rw), num(rh)],
    Resources: { XObject: { Im0: png.ref } },
  })
  const annot = pdfDoc.context.obj({
    Type: 'Annot',
    Subtype: subtype,
    Rect: [num(x1), num(y1), num(x2), num(y2)],
    F: 4,
    P: page.ref,
    AP: { N: pdfDoc.context.register(ap) },
  })
  if (d.contents !== undefined) {
    annot.set(PDFName.of('Contents'), PDFHexString.fromText(d.contents))
  }
  annot.set(PDFName.of('T'), PDFHexString.fromText('ARAOffice'))
  appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
}

/** Drawing annots: hand-written AP for Ink/Square/Circle/Line; notes are standard Text annots (viewer draws the icon) */
function addDrawing(pdfDoc: PDFDocument, page: PDFPage, d: DrawingInput): void {
  if (d.kind === 'image' || d.kind === 'freetext') return // handled by addImageAnnot (needs async embed)
  const [r, g, b] = d.color

  if (d.kind === 'note') {
    const [x, y] = d.at
    const annot = pdfDoc.context.obj({
      Type: 'Annot',
      Subtype: 'Text',
      Rect: [num(x), num(y - 18), num(x + 20), num(y)],
      Name: 'Comment',
      C: d.color,
      F: 4,
      P: page.ref,
    })
    annot.set(PDFName.of('Contents'), PDFHexString.fromText(d.contents))
    annot.set(PDFName.of('T'), PDFHexString.fromText('ARAOffice'))
    appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
    return
  }

  if (d.kind === 'caret') {
    // A thin stroked wedge sitting on the insertion point, the way a caret reads on screen
    const [x, y] = d.at
    const w = Math.max(3, d.size * 0.4)
    const h = Math.max(2, d.size * 0.5)
    const lw = Math.max(0.6, num(d.size * 0.08))
    const pad = lw / 2
    const ap = pdfDoc.context.stream(
      `${lw} w 1 J 1 j ${r} ${g} ${b} RG ${num(x)} ${num(y)} m ` +
        `${num(x + w / 2)} ${num(y + h)} l ${num(x + w)} ${num(y)} l S`,
      {
        Type: 'XObject',
        Subtype: 'Form',
        BBox: [num(x - pad), num(y - pad), num(x + w + pad), num(y + h + pad)],
      },
    )
    const annot = pdfDoc.context.obj({
      Type: 'Annot',
      Subtype: 'Caret',
      Rect: [num(x), num(y), num(x + w), num(y + h)],
      C: d.color,
      F: 4,
      P: page.ref,
      AP: { N: pdfDoc.context.register(ap) },
    })
    annot.set(PDFName.of('Contents'), PDFHexString.fromText(d.contents))
    annot.set(PDFName.of('T'), PDFHexString.fromText('ARAOffice'))
    appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
    return
  }

  const ops: string[] = [`${num(d.width)} w 1 J 1 j ${r} ${g} ${b} RG`]
  let xs: number[] = []
  let ys: number[] = []
  let subtype: string

  if (d.kind === 'ink') {
    subtype = 'Ink'
    for (const path of d.paths) {
      if (path.length < 4) continue
      ops.push(`${num(path[0]!)} ${num(path[1]!)} m`)
      for (let i = 2; i < path.length; i += 2) ops.push(`${num(path[i]!)} ${num(path[i + 1]!)} l`)
      ops.push('S')
      for (let i = 0; i < path.length; i += 2) {
        xs.push(path[i]!)
        ys.push(path[i + 1]!)
      }
    }
  } else if (d.kind === 'rect' || d.kind === 'ellipse') {
    const [x1, y1, x2, y2] = d.rect
    subtype = d.kind === 'rect' ? 'Square' : 'Circle'
    if (d.kind === 'rect') ops.push(`${num(x1)} ${num(y1)} ${num(x2 - x1)} ${num(y2 - y1)} re S`)
    else ops.push(...ellipseOps(x1, y1, x2, y2))
    xs = [x1, x2]
    ys = [y1, y2]
  } else if (d.kind === 'polygon' || d.kind === 'polyline') {
    const closed = d.kind === 'polygon'
    // An odd coordinate list is malformed; too few points cannot form the shape
    if (d.points.length < MIN_POLY_POINTS[d.kind] || d.points.length % 2 !== 0) return
    subtype = closed ? 'Polygon' : 'PolyLine'
    ops.push(`${num(d.points[0]!)} ${num(d.points[1]!)} m`)
    for (let i = 2; i < d.points.length; i += 2) {
      ops.push(`${num(d.points[i]!)} ${num(d.points[i + 1]!)} l`)
    }
    ops.push(closed ? 's' : 'S')
    for (let i = 0; i < d.points.length; i += 2) {
      xs.push(d.points[i]!)
      ys.push(d.points[i + 1]!)
    }
  } else {
    const [fx, fy] = d.from
    const [tx, ty] = d.to
    subtype = 'Line'
    ops.push(`${num(fx)} ${num(fy)} m ${num(tx)} ${num(ty)} l S`)
    xs = [fx, tx]
    ys = [fy, ty]
    if (d.kind === 'arrow') {
      const ang = Math.atan2(ty - fy, tx - fx)
      const len = Math.max(9, d.width * 4.5)
      for (const off of [-0.45, 0.45]) {
        const hx = tx - len * Math.cos(ang + off)
        const hy = ty - len * Math.sin(ang + off)
        ops.push(`${num(tx)} ${num(ty)} m ${num(hx)} ${num(hy)} l S`)
        xs.push(hx)
        ys.push(hy)
      }
    }
  }

  const pad = d.width + 2
  const rect = [
    Math.min(...xs) - pad,
    Math.min(...ys) - pad,
    Math.max(...xs) + pad,
    Math.max(...ys) + pad,
  ]
  const ap = pdfDoc.context.stream(ops.join('\n'), { Type: 'XObject', Subtype: 'Form', BBox: rect })
  const annot = pdfDoc.context.obj({
    Type: 'Annot',
    Subtype: subtype,
    Rect: rect,
    C: d.color,
    F: 4,
    P: page.ref,
    BS: { W: d.width },
    AP: { N: pdfDoc.context.register(ap) },
  })
  if (d.kind === 'ink') annot.set(PDFName.of('InkList'), pdfDoc.context.obj(d.paths))
  if (d.kind === 'line' || d.kind === 'arrow') {
    annot.set(PDFName.of('L'), pdfDoc.context.obj([...d.from, ...d.to]))
  }
  if (d.kind === 'polygon' || d.kind === 'polyline') {
    annot.set(PDFName.of('Vertices'), pdfDoc.context.obj(d.points))
  }
  annot.set(PDFName.of('T'), PDFHexString.fromText('ARAOffice'))
  appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
}

function applyFormValues(pdfDoc: PDFDocument, values: FormValueInput[]): void {
  const form = pdfDoc.getForm()
  for (const v of values) {
    if (v.kind === 'text') {
      form.getTextField(v.name).setText(v.value ?? '')
    } else if (v.kind === 'radio') {
      const rg = form.getRadioGroup(v.name)
      if (v.value) rg.select(v.value)
      else rg.clear()
    } else if (v.kind === 'choice') {
      const f = form.getField(v.name)
      if (f instanceof PDFDropdown || f instanceof PDFOptionList) {
        if (v.value) f.select(v.value)
        else f.clear()
      }
    } else {
      const cb = form.getCheckBox(v.name)
      if (v.checked) cb.check()
      else cb.uncheck()
    }
  }
}

/** Split into one document per range of original page indices, in the given order */
export async function splitPdfBytes(bytes: Uint8Array, ranges: number[][]): Promise<Uint8Array[]> {
  const out: Uint8Array[] = []
  // Sequential: extractPagesBytes loads the source each time, and the ranges share it
  for (const range of ranges) out.push(await extractPagesBytes(bytes, range))
  return out
}

/**
 * Merge several documents into a new one, in the given order. With `bookmarkTitles`, each source
 * gets one top-level bookmark at its first page.
 *
 * Not carried over: a source's own nested bookmarks (only the per-file entry is written), and
 * colliding form field names across sources are left as-is rather than renamed or merged.
 */
export async function combinePdfsBytes(
  sources: Uint8Array[],
  bookmarkTitles: string[] | null,
): Promise<Uint8Array> {
  const out = await PDFDocument.create()
  const parts: { pageIndex: number; pages: number; title: string }[] = []
  for (const [i, bytes] of sources.entries()) {
    const src = await PDFDocument.load(bytes, { updateMetadata: false })
    const indices = src.getPageIndices()
    parts.push({
      pageIndex: out.getPageCount(),
      pages: indices.length,
      title: bookmarkTitles?.[i] ?? '',
    })
    const copied = await out.copyPages(src, indices)
    for (const page of copied) out.addPage(page)
  }
  if (bookmarkTitles) {
    applyOutline(
      out,
      // A source with no pages has nothing to point at: its bookmark would land on the next
      // source's first page, so it is dropped rather than written wrong
      parts.flatMap((p) => (p.pages > 0 ? [{ title: p.title, pageIndex: p.pageIndex }] : [])),
    )
  }
  return out.save({ useObjectStreams: false })
}

const SANITIZE_KEYS = [
  'metadata',
  'javascript',
  'attachments',
  'comments',
  'thumbnails',
  'privateData',
] as const

/**
 * Remove the chosen categories and report how much went, for a result toast.
 *
 * The walk is parameterized by `apply` so the preview census and the removal share one code path
 * and can never drift. Absent structures contribute nothing. /Info is blanked rather than deleted:
 * pdf-lib serializes from trailerInfo and does not model a missing Info dict, so dropping fields
 * instead of the object is the safe form; the values are still gone by the time the file is written.
 * Form-field and annotation action scripts are removed only as far as deleting the objects they
 * live in goes, and form fields themselves are kept.
 */
function walkSanitize(
  pdfDoc: PDFDocument,
  request: SanitizeRequest,
  apply: boolean,
): SanitizeReport {
  const ctx = pdfDoc.context
  const catalog = pdfDoc.catalog
  const report = Object.fromEntries(SANITIZE_KEYS.map((key) => [key, 0])) as SanitizeReport
  /** Counting pairs with a single `apply=false` branch keeps census and removal byte-equivalent */
  const drop = (dict: PDFDict, key: PDFName) => {
    if (apply) dict.delete(key)
  }

  if (request.metadata) {
    const info = ctx.lookupMaybe(ctx.trailerInfo.Info, PDFDict)
    if (info) {
      for (const key of [
        'Title',
        'Author',
        'Subject',
        'Keywords',
        'Creator',
        'Producer',
        'CreationDate',
        'ModDate',
      ]) {
        if (info.has(PDFName.of(key))) {
          drop(info, PDFName.of(key))
          report.metadata += 1
        }
      }
    }
    if (catalog.has(PDFName.of('Metadata'))) {
      drop(catalog, PDFName.of('Metadata'))
      report.metadata += 1
    }
  }

  if (request.javascript || request.attachments) {
    const names = catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
    if (names) {
      if (request.javascript && names.has(PDFName.of('JavaScript'))) {
        drop(names, PDFName.of('JavaScript'))
        report.javascript += 1
      }
      if (request.attachments && names.has(PDFName.of('EmbeddedFiles'))) {
        drop(names, PDFName.of('EmbeddedFiles'))
        report.attachments += 1
      }
    }
    if (request.attachments && catalog.has(PDFName.of('Collection'))) {
      drop(catalog, PDFName.of('Collection'))
      report.attachments += 1
    }
  }

  if (request.javascript) {
    for (const key of ['OpenAction', 'AA']) {
      if (catalog.has(PDFName.of(key))) {
        drop(catalog, PDFName.of(key))
        report.javascript += 1
      }
    }
  }

  for (const page of pdfDoc.getPages()) {
    if (request.comments) {
      const annots = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
      if (annots && annots.size() > 0) {
        // Widget annotations carry form fields, so they stay: deleting them would orphan the
        // AcroForm entries behind fields the user still has. Only markup annots are removed.
        const keep: PDFObject[] = []
        for (let i = 0; i < annots.size(); i++) {
          const entry = annots.get(i)
          const subtype = ctx
            .lookupMaybe(entry, PDFDict)
            ?.lookupMaybe(PDFName.of('Subtype'), PDFName)
          if (subtype?.decodeText() === 'Widget') {
            if (apply) keep.push(entry)
          } else report.comments += 1
        }
        // Only the annot rewrite is apply-only; the page-tail census below must still run
        if (apply && keep.length < annots.size()) {
          if (keep.length > 0) page.node.set(PDFName.of('Annots'), ctx.obj(keep))
          else page.node.delete(PDFName.of('Annots'))
        }
      }
    }
    if (request.thumbnails && page.node.has(PDFName.of('Thumb'))) {
      drop(page.node, PDFName.of('Thumb'))
      report.thumbnails += 1
    }
    if (request.privateData && page.node.has(PDFName.of('PieceInfo'))) {
      drop(page.node, PDFName.of('PieceInfo'))
      report.privateData += 1
    }
  }

  if (request.privateData && catalog.has(PDFName.of('PieceInfo'))) {
    drop(catalog, PDFName.of('PieceInfo'))
    report.privateData += 1
  }

  return report
}

/**
 * Read-only census for the sanitize dialog: which categories hold anything in this file.
 * Runs the exact removal walk without applying it, never touches the bytes.
 */
export async function scanSanitizeCounts(bytes: Uint8Array): Promise<SanitizeReport> {
  const pdfDoc = await PDFDocument.load(bytes, { updateMetadata: false })
  const all: SanitizeRequest = {
    metadata: true,
    javascript: true,
    attachments: true,
    comments: true,
    thumbnails: true,
    privateData: true,
  }
  return walkSanitize(pdfDoc, all, false)
}

export async function sanitizePdfBytes(
  bytes: Uint8Array,
  request: SanitizeRequest,
): Promise<{ bytes: Uint8Array; report: SanitizeReport } | null> {
  if (!SANITIZE_KEYS.some((key) => request[key])) return null

  const pdfDoc = await PDFDocument.load(bytes, { updateMetadata: false })
  const report = walkSanitize(pdfDoc, request, true)
  return { bytes: await pdfDoc.save({ useObjectStreams: false }), report }
}

/** Extract the given pages (original indices) into bytes of a new PDF */
export async function extractPagesBytes(bytes: Uint8Array, pages: number[]): Promise<Uint8Array> {
  const src = await PDFDocument.load(bytes, { updateMetadata: false })
  const out = await PDFDocument.create()
  const valid = pages.filter((p) => p >= 0 && p < src.getPageCount())
  const copied = await out.copyPages(src, valid)
  for (const p of copied) out.addPage(p)
  return out.save({ useObjectStreams: false })
}

/** Insert all pages of another PDF after afterPageIndex (-1 = front); returns merged bytes and inserted page count */
export async function insertPdfBytes(
  bytes: Uint8Array,
  otherBytes: Uint8Array,
  afterPageIndex: number,
): Promise<{ merged: Uint8Array; count: number }> {
  const dst = await PDFDocument.load(bytes, { updateMetadata: false })
  const src = await PDFDocument.load(otherBytes, { updateMetadata: false })
  const copied = await dst.copyPages(src, src.getPageIndices())
  let at = Math.min(Math.max(afterPageIndex + 1, 0), dst.getPageCount())
  for (const p of copied) dst.insertPage(at++, p)
  return { merged: await dst.save({ useObjectStreams: false }), count: copied.length }
}

function applyMetadata(pdfDoc: PDFDocument, meta: MetadataInput): void {
  if (meta.title !== undefined) pdfDoc.setTitle(meta.title)
  if (meta.author !== undefined) pdfDoc.setAuthor(meta.author)
  if (meta.subject !== undefined) pdfDoc.setSubject(meta.subject)
  if (meta.keywords !== undefined) {
    pdfDoc.setKeywords(
      meta.keywords
        .split(/[,，;；]/)
        .map((k) => k.trim())
        .filter(Boolean),
    )
  }
  pdfDoc.setModificationDate(new Date())
}

/**
 * Apply the request to the PDF at sourcePath and atomically write the result to targetPath
 * (temp file next to the target + rename, so a mid-write crash can't corrupt it).
 * The source file is only ever read: Save As (targetPath !== sourcePath) must never mutate
 * the original document, and a failed or cancelled save leaves both paths untouched.
 * In-place Save passes targetPath === sourcePath.
 */
export async function savePdfToPath(
  sourcePath: string,
  targetPath: string,
  request: SavePdfRequest,
): Promise<void> {
  const bytes = await applySaveRequest(new Uint8Array(await readFile(sourcePath)), request)
  await atomicWriteFile(targetPath, bytes)
}

/** Apply markups + form values + page ops, returning new bytes. Original objects are not reordered (pdf-lib keeps untouched objects). */
export async function applySaveRequest(
  bytes: Uint8Array,
  request: SavePdfRequest,
): Promise<Uint8Array> {
  const pdfDoc = await PDFDocument.load(bytes, { updateMetadata: false })
  if (request.formValues.length > 0) applyFormValues(pdfDoc, request.formValues)
  const pages = pdfDoc.getPages()
  // Apply rotations first so markup appearances draw lines for the page's final orientation
  for (const r of request.rotations ?? []) {
    const page = pages[r.pageIndex]
    if (page) page.setRotation(degrees((page.getRotation().angle + r.delta) % 360))
  }
  for (const m of request.markups) {
    const page = pages[m.pageIndex]
    if (page) addMarkup(pdfDoc, page, m)
  }
  for (const d of request.drawings ?? []) {
    const page = pages[d.pageIndex]
    if (!page) continue
    if (d.kind === 'image') await addImageAnnot(pdfDoc, page, d, 'Stamp')
    else if (d.kind === 'freetext') await addImageAnnot(pdfDoc, page, d, 'FreeText')
    else addDrawing(pdfDoc, page, d)
  }
  for (const s of request.stamps ?? []) {
    const page = pages[s.pageIndex]
    if (!page) continue
    const png = await pdfDoc.embedPng(s.image)
    const [x1, y1, x2, y2] = s.rect
    page.drawImage(png, {
      x: x1,
      y: y1,
      width: x2 - x1,
      height: y2 - y1,
      opacity: s.opacity ?? 1,
    })
  }
  if (request.metadata) applyMetadata(pdfDoc, request.metadata)
  // Deletions go last, in descending order; earlier ops all address original page indices
  for (const idx of [...(request.deletedPages ?? [])].sort((a, b) => b - a)) {
    if (idx >= 0 && idx < pdfDoc.getPageCount() && pdfDoc.getPageCount() > 1) pdfDoc.removePage(idx)
  }
  // Reorder last: pageOrder gives the new order of remaining-after-delete pages by original index.
  // pdf-lib's removePage never invalidates its page cache, so getPages() here would return the
  // stale pre-deletion list — derive the surviving pages from the pre-deletion snapshot instead.
  const order = request.pageOrder
  if (order && order.length > 0) {
    const deletedSet = new Set(request.deletedPages ?? [])
    const target = order.flatMap((o) => (deletedSet.has(o) ? [] : pages[o] ? [pages[o]!] : []))
    if (target.length === pdfDoc.getPageCount()) {
      while (pdfDoc.getPageCount() > 0) pdfDoc.removePage(0)
      for (const p of target) pdfDoc.addPage(p)
    }
  }
  try {
    return await pdfDoc.save({ useObjectStreams: false })
  } catch (err) {
    // Form values beyond WinAnsi (e.g. CJK) make pdf-lib's appearance generation fail:
    // skip it and set NeedAppearances so viewers rebuild them (Acrobat/pdfjs both support this)
    if (request.formValues.length === 0) throw err
    pdfDoc.getForm().acroForm.dict.set(PDFName.of('NeedAppearances'), PDFBool.True)
    return await pdfDoc.save({ useObjectStreams: false, updateFieldAppearances: false })
  }
}
