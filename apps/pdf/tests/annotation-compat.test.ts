import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { applySaveRequest } from '../src/main/save-pdf'
import type { SavePdfRequest } from '../src/shared/ipc'

/**
 * Compatibility of what we write with the renderer ARA actually displays documents with.
 *
 * pdf-lib will happily write an annotation dict that no viewer can render — a subtype it accepts
 * but a renderer rejects, a missing appearance, geometry in the wrong space. `save-pdf.test.ts`
 * checks the structure we intend; this file checks that pdfjs accepts it, which is the closest
 * thing to "the user can see it" that a test can assert without a display.
 */

const request = (over: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path: '/tmp/test.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  ...over,
})

async function blank(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  doc.addPage([612, 792])
  return doc.save({ useObjectStreams: false })
}

/** 1x1 red pixel PNG */
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

type PdfjsAnnotation = { subtype: string; rect: number[]; hasAppearance: boolean }

/** The legacy build: the modern entry point expects a browser canvas. Options stay minimal, with
    `useSystemFonts` off so the result does not depend on the fonts installed on the machine. */
async function loadDoc(bytes: Uint8Array) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  return pdfjs.getDocument({ data: bytes, useSystemFonts: false }).promise
}

async function pdfjsAnnotations(bytes: Uint8Array): Promise<PdfjsAnnotation[]> {
  const page = await (await loadDoc(bytes)).getPage(1)
  return (await page.getAnnotations()) as PdfjsAnnotation[]
}

describe('annotations are accepted by pdfjs', () => {
  it('recognises every annotation kind we write, with an appearance', async () => {
    const quad = [[20, 700, 200, 700, 20, 688, 200, 688]]
    const saved = await applySaveRequest(
      await blank(),
      request({
        markups: [
          { pageIndex: 0, type: 'highlight', color: [1, 0.87, 0.35], quads: quad },
          { pageIndex: 0, type: 'underline', color: [0.17, 0.4, 1], quads: quad },
          { pageIndex: 0, type: 'strikeout', color: [0.86, 0.22, 0.18], quads: quad },
          { pageIndex: 0, type: 'squiggly', color: [0.55, 0.25, 0.85], quads: quad },
        ],
        drawings: [
          { kind: 'ink', pageIndex: 0, color: [0, 0, 1], width: 2, paths: [[40, 600, 80, 620]] },
          { kind: 'rect', pageIndex: 0, color: [0, 0, 1], width: 2, rect: [40, 560, 160, 600] },
          { kind: 'ellipse', pageIndex: 0, color: [0, 0, 1], width: 2, rect: [40, 500, 160, 540] },
          {
            kind: 'arrow',
            pageIndex: 0,
            color: [0, 0, 1],
            width: 2,
            from: [40, 460],
            to: [160, 480],
          },
          {
            kind: 'polyline',
            pageIndex: 0,
            color: [0, 0, 1],
            width: 2,
            points: [40, 400, 120, 440, 200, 410],
          },
          {
            kind: 'polygon',
            pageIndex: 0,
            color: [0, 0.5, 0],
            width: 2,
            points: [40, 300, 160, 300, 100, 360],
          },
          {
            kind: 'caret',
            pageIndex: 0,
            color: [1, 0, 0],
            at: [250, 200],
            size: 14,
            contents: 'insert',
          },
          { kind: 'note', pageIndex: 0, color: [1, 0.8, 0], at: [300, 200], contents: 'note' },
          {
            kind: 'freetext',
            pageIndex: 0,
            image: PNG,
            rect: [300, 560, 420, 580],
            contents: 'typed note',
          },
        ],
      }),
    )

    const annots = await pdfjsAnnotations(saved)
    expect(annots.map((a) => a.subtype)).toEqual([
      'Highlight',
      'Underline',
      'StrikeOut',
      'Squiggly',
      'Ink',
      'Square',
      'Circle',
      'Line',
      'PolyLine',
      'Polygon',
      'Caret',
      'Text',
      'FreeText',
    ])
    // A viewer draws from the appearance stream when one is present, so every shape and markup
    // annotation must carry one or it stays invisible until the viewer synthesises the geometry.
    // Text (sticky note) is the one deliberate exception: it has no /AP and the viewer draws the
    // icon from /Name ("Comment"), which is what the PDF spec expects. Adding an /AP there would
    // make viewers that draw both show the note twice.
    const withAppearance = annots.filter((a) => a.subtype !== 'Text')
    for (const a of withAppearance) {
      expect(a.hasAppearance, `${a.subtype} has no appearance`).toBe(true)
    }
    expect(annots.find((a) => a.subtype === 'Text')!.hasAppearance).toBe(false)
  })

  it('keeps annotation geometry inside the page it belongs to', async () => {
    const saved = await applySaveRequest(
      await blank(),
      request({
        drawings: [
          {
            kind: 'polygon',
            pageIndex: 0,
            color: [0, 0.5, 0],
            width: 2,
            points: [40, 300, 160, 300, 100, 360],
          },
          {
            kind: 'caret',
            pageIndex: 0,
            color: [1, 0, 0],
            at: [250, 200],
            size: 14,
            contents: 'insert',
          },
        ],
      }),
    )
    const [polygon, caret] = await pdfjsAnnotations(saved)
    // A rect outside the media box means the coordinates went in the wrong space
    for (const a of [polygon!, caret!]) {
      expect(a.rect[0]).toBeGreaterThanOrEqual(0)
      expect(a.rect[1]).toBeGreaterThanOrEqual(0)
      expect(a.rect[2]).toBeLessThanOrEqual(612)
      expect(a.rect[3]).toBeLessThanOrEqual(792)
    }
    expect(polygon!.rect[0]).toBeLessThan(polygon!.rect[2])
    expect(polygon!.rect[1]).toBeLessThan(polygon!.rect[3])
  })

  it('renders the page without throwing', async () => {
    const saved = await applySaveRequest(
      await blank(),
      request({
        markups: [
          {
            pageIndex: 0,
            type: 'squiggly',
            color: [0.55, 0.25, 0.85],
            quads: [[20, 700, 200, 700, 20, 688, 200, 688]],
          },
        ],
      }),
    )
    const doc = await loadDoc(saved)
    const page = await doc.getPage(1)
    // Building the operator list walks the content stream and the annotation appearances,
    // so a malformed appearance stream fails here rather than silently at render time
    const op = await page.getOperatorList()
    expect(op.fnArray.length).toBeGreaterThan(0)
  })
})
