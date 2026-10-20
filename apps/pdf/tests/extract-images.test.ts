import { deflateSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { PDFDocument, PDFName, PDFRawStream } from 'pdf-lib'
import { collectEmbeddedImages } from '../src/main/extract-images'

/**
 * Fixtures are written by hand against the low-level API, because the encoder only has to trust
 * the geometry a dict states: embedding via pdf-lib's own embedPng/addPage helpers would leave
 * the tests unable to construct the shapes being classified (predictor images, the wrong Mystream
 * filter, a stream whose inflated length disagrees with its dimensions).
 */

/** Register a raw stream on the document and hang it off page one of the /XObject dict */
function attach(
  doc: PDFDocument,
  name: string,
  // LiteralObject is not exported from pdf-lib; index it off the stream method's own signature
  dict: Parameters<PDFDocument['context']['stream']>[1],
  contents: Uint8Array,
): void {
  // context.stream takes contents first, unlike PDFRawStream.of(dict, contents)
  const stream = doc.context.stream(contents, dict) as PDFRawStream
  doc.getPage(0).node.setXObject(PDFName.of(name), doc.context.register(stream))
}

const RGB = { ColorSpace: 'DeviceRGB', BitsPerComponent: 8 }

/** One row, two RGB pixels; returns the uncompressed sample bytes as well for the BMP assertions */
function rgbPixels(): { raw: Uint8Array; packed: Uint8Array } {
  const raw = new Uint8Array([10, 20, 30, 200, 150, 100])
  // Wrapped so the test realm owns the bytes: zlib's Buffer fails pdf-lib's Uint8Array check
  return { raw, packed: new Uint8Array(deflateSync(raw)) }
}

describe('collectEmbeddedImages', () => {
  /** BMP pixel-data offset, read as a 4-byte little-endian field — verifies the header, too */
  const le32 = (bmp: Uint8Array, pos: number) =>
    bmp[pos]! | (bmp[pos + 1]! << 8) | (bmp[pos + 2]! << 16) | (bmp[pos + 3]! << 24)
  it('passes a DCTDecode stream through byte for byte as .jpg', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3, 0xd9])
    attach(doc, 'Im0', { Subtype: 'Image', Filter: 'DCTDecode', Width: 2, Height: 1, ...RGB }, jpeg)

    const { images, skipped } = await collectEmbeddedImages(
      await doc.save({ useObjectStreams: false }),
    )
    expect(skipped).toBe(0)
    expect(images).toHaveLength(1)
    expect(images[0]!.ext).toBe('jpg')
    expect(images[0]!.name).toBe('Im0')
    // Passing through means the export is the file's own bytes, not a re-encoding of them
    expect(Array.from(images[0]!.data)).toEqual(Array.from(jpeg))
    expect(images[0]!.pages).toEqual([1])
  })

  it('decodes a plain Flate RGB image into a bottom-up BGR bitmap', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const { raw, packed } = rgbPixels()
    attach(
      doc,
      'Im1',
      { Subtype: 'Image', Filter: 'FlateDecode', Width: 2, Height: 1, ...RGB },
      packed,
    )
    const { images } = await collectEmbeddedImages(await doc.save({ useObjectStreams: false }))
    expect(images).toHaveLength(1)
    expect(images[0]!.ext).toBe('bmp')

    const bmp = images[0]!.data
    expect(String.fromCharCode(bmp[0]!, bmp[1]!)).toBe('BM')
    expect(bmp[28]).toBe(24) // bit depth
    // Two pixels at 3 bytes, padded to a 4-byte row
    expect(bmp[34]).toBe(8)
    // Bottom-up: at 1-row height the memory row is the pixel row
    // Samples are BGR on disk, and 2px×3B pads to a 4-byte row. bmp[10] is the full
    // pixel-data offset (header + palette already included)
    expect(Array.from(bmp.slice(le32(bmp, 10)))).toEqual([30, 20, 10, 100, 150, 200, 0, 0])
    expect(Array.from(raw)).toEqual([10, 20, 30, 200, 150, 100])
  })

  it('writes a grayscale image with its identity palette', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const raw = new Uint8Array([0, 255])
    attach(
      doc,
      'G',
      {
        Subtype: 'Image',
        Filter: 'FlateDecode',
        Width: 2,
        Height: 1,
        ColorSpace: 'DeviceGray',
        BitsPerComponent: 8,
      },
      new Uint8Array(deflateSync(raw)),
    )
    const { images } = await collectEmbeddedImages(await doc.save({ useObjectStreams: false }))
    expect(images).toHaveLength(1)
    const bmp = images[0]!.data
    expect(bmp[28]).toBe(8)
    // 1x1 gray pixels land through the identity palette: byte 0 maps to (0,0,0), byte 255 to white
    const dataStart = le32(bmp, 10)
    expect(Array.from(bmp.slice(dataStart, dataStart + 4))).toEqual([0, 255, 0, 0])
  })

  it('exports an image shared by two pages once, with both pages listed', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    doc.addPage([300, 300])
    const { packed } = rgbPixels()
    const stream = doc.context.stream(packed, {
      Subtype: 'Image',
      Filter: 'FlateDecode',
      Width: 2,
      Height: 1,
      ...RGB,
    }) as PDFRawStream
    const ref = doc.context.register(stream)
    doc.getPage(0).node.setXObject(PDFName.of('Shared'), ref)
    doc.getPage(1).node.setXObject(PDFName.of('Again'), ref)

    const { images } = await collectEmbeddedImages(await doc.save({ useObjectStreams: false }))
    expect(images).toHaveLength(1)
    expect(images[0]!.pages).toEqual([1, 2])
  })

  it('reaches images inside a Form XObject', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const { packed } = rgbPixels()
    const inner = doc.context.stream(packed, {
      Subtype: 'Image',
      Filter: 'FlateDecode',
      Width: 2,
      Height: 1,
      ...RGB,
    })
    const form = doc.context.stream(new Uint8Array(), {
      Subtype: 'Form',
      BBox: doc.context.obj([]),
      Resources: doc.context.obj({
        XObject: doc.context.obj({ Im1: doc.context.register(inner) }),
      }),
    })
    doc.getPage(0).node.setXObject(PDFName.of('Form0'), doc.context.register(form))

    const { images, skipped } = await collectEmbeddedImages(
      await doc.save({ useObjectStreams: false }),
    )
    expect(images).toHaveLength(1)
    expect(images[0]!.pages).toEqual([1])
    expect(skipped).toBe(0)
  })

  it('skips what it cannot decode faithfully instead of exporting it wrong', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const { raw, packed } = rgbPixels()
    // Predictor > 1: pre-filtered rows, the shape of an embedded PNG
    attach(
      doc,
      'Pred',
      {
        Subtype: 'Image',
        Filter: 'FlateDecode',
        DecodeParms: doc.context.obj({ Predictor: 12 }),
        Width: 2,
        Height: 1,
        ...RGB,
      },
      new Uint8Array(deflateSync(raw)),
    )
    attach(
      doc,
      'Jpx',
      { Subtype: 'Image', Filter: 'JPXDecode', Width: 2, Height: 1, ...RGB },
      packed,
    )
    attach(
      doc,
      'Masked',
      {
        Subtype: 'Image',
        Filter: 'FlateDecode',
        SMask: doc.context.obj({}),
        Width: 2,
        Height: 1,
        ...RGB,
      },
      packed,
    )
    attach(
      doc,
      'ICC',
      {
        Subtype: 'Image',
        Filter: 'FlateDecode',
        ColorSpace: 'ICCBased',
        Width: 2,
        Height: 1,
        BitsPerComponent: 8,
      },
      packed,
    )
    attach(
      doc,
      'Short',
      { Subtype: 'Image', Filter: 'FlateDecode', Width: 4, Height: 2, ...RGB },
      packed, // inflates to 6 bytes, the dict says 24
    )
    attach(doc, 'NotImage', { Subtype: 'Form' }, new Uint8Array()) // recursed into, not counted

    const { images, skipped } = await collectEmbeddedImages(
      await doc.save({ useObjectStreams: false }),
    )
    expect(images).toHaveLength(0)
    // The five unexportable images. The Form container is recursed into rather than counted: it
    // holds no image of its own, so it contributes neither an export nor a skip
    expect(skipped).toBe(5)
  })

  it('reports an empty scan for a document with no images', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const { images, skipped } = await collectEmbeddedImages(
      await doc.save({ useObjectStreams: false }),
    )
    expect(images).toEqual([])
    expect(skipped).toBe(0)
  })
})
