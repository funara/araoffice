import { inflateSync } from 'node:zlib'
import {
  PDFArray,
  PDFBool,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFRawStream,
  PDFRef,
} from 'pdf-lib'
import type { PDFObject } from 'pdf-lib'

/**
 * Embedded image extraction. Fidelity first: an image is written out either exactly as the file
 * carries it, or not at all — the same bar the save path applies to untouched objects.
 *
 * - `/Filter DCTDecode` (a JPEG) passes through unchanged as .jpg: the bytes the file carries are
 *   the bytes a viewer gets.
 * - A single plain `FlateDecode` pass at 8-bit DeviceRGB or DeviceGray, with no mask, no decode
 *   array and no predictor, is decoded and written as BMP — no image library needed, because the
 *   data is already exactly the geometry BMP stores.
 * - Everything else is skipped and counted rather than written: JPX, CCITT, JBIG2, LZW, PNG
 *   predictors (the shape of almost every embedded PNG and most scans), ICC or indexed colour,
 *   image masks, decode arrays and other bit depths. A plausible-looking wrong export is worse
 *   than a missing one, and the page still shows the content in the viewer where the renderer
 *   handles it.
 *
 * Images shared across pages are exported once, with every page that draws them reported. Form
 * XObjects are recursed into, since pages usually draw images through them.
 */

export interface EmbeddedImage {
  ext: 'jpg' | 'bmp'
  data: Uint8Array
  /** 1-based page numbers that draw this image, ascending */
  pages: number[]
  /** Its name in the /Resources of the first page that drew it */
  name: string
}

export interface EmbeddedImageScan {
  images: EmbeddedImage[]
  skipped: number
}

/** A 64k-side 24-bit image is far past any screenshot; larger streams are treated as bombs */
const MAX_PIXELS = 64_000_000
const MAX_DEPTH = 4

/** /Filter as a list of names, accepting both the single-name and array forms */
function filtersOf(dict: PDFDict, ctx: PDFDocument['context']): string[] {
  const single = ctx.lookupMaybe(dict.get(PDFName.of('Filter')), PDFName)
  if (single) return [single.decodeText()!]
  const list = ctx.lookupMaybe(dict.get(PDFName.of('Filter')), PDFArray)
  if (!list) return []
  const out: string[] = []
  for (let i = 0; i < list.size(); i++) {
    const name = ctx.lookupMaybe(list.get(i), PDFName)
    // A filter list is names; anything else is not a stream this module wrote bytes for
    if (!name) return []
    out.push(name.decodeText()!)
  }
  return out
}

/**
 * BMP encoder for the plain path: bottom-up rows, 4-byte row padding, BGR sample order, and an
 * identity grayscale palette for the one-channel case. Pure bytes, because the input is already
 * exactly the geometry the format stores.
 */
function bmpEncode(
  pixels: Uint8Array,
  width: number,
  height: number,
  gray: boolean,
): Uint8Array | null {
  const channels = gray ? 1 : 3
  const rowBytes = Math.ceil((width * channels) / 4) * 4
  const palette = gray ? 256 : 0
  const offset = 14 + 40 + palette * 4
  const dataSize = rowBytes * height
  if (!Number.isSafeInteger(dataSize) || dataSize < 0 || dataSize > 1 << 28) return null

  const out = new Uint8Array(offset + dataSize)
  const le16 = (pos: number, value: number) => {
    out[pos]! = value & 0xff
    out[pos + 1]! = (value >>> 8) & 0xff
  }
  const le32 = (pos: number, value: number) => {
    for (let i = 0; i < 4; i++) out[pos + i]! = (value >>> (8 * i)) & 0xff
  }

  // BITMAPFILEHEADER, then BITMAPINFOHEADER
  out[0]! = 'B'.charCodeAt(0)
  out[1]! = 'M'.charCodeAt(0)
  le32(2, out.length)
  le32(10, offset)
  le32(14, 40)
  le32(18, width)
  le32(22, height) // positive height: rows stored bottom-up
  le16(26, 1)
  le16(28, gray ? 8 : 24)
  le32(30, 0) // BI_RGB
  le32(34, dataSize)
  le32(38, 2835) // 72 dpi
  le32(42, 2835)
  le32(46, palette)
  le32(50, palette)

  for (let i = 0; i < palette; i++) {
    const at = 54 + i * 4
    out[at] = i
    out[at + 1] = i
    out[at + 2] = i
    out[at + 3] = 0
  }

  for (let y = 0; y < height; y++) {
    const dst = (height - 1 - y) * rowBytes + offset
    for (let x = 0; x < width; x++) {
      if (gray) {
        out[dst + x] = pixels[y * width + x]!
      } else {
        const p = (y * width + x) * 3
        out[dst + x * 3] = pixels[p + 2]!
        out[dst + x * 3 + 1] = pixels[p + 1]!
        out[dst + x * 3 + 2] = pixels[p]!
      }
    }
  }
  return out
}

/** The image as BMP bytes, or null when it cannot be decoded faithfully here */
function flateToBitmap(stream: PDFRawStream, ctx: PDFDocument['context']): Uint8Array | null {
  const dict: PDFDict = stream.dict
  const filters = filtersOf(dict, ctx)
  // Anything beyond a single Flate pass is not this decoder's business
  if (filters.length !== 1 || filters[0] !== 'FlateDecode') return null

  const parms = ctx.lookupMaybe(dict.get(PDFName.of('DecodeParms')), PDFDict)
  // Predictor > 1 means rows arrive pre-filtered, the shape of an embedded PNG; inflating that
  // without implementing the filters would give plausible-looking garbage
  const predictor = ctx.lookupMaybe(parms?.get(PDFName.of('Predictor')), PDFNumber)?.asNumber() ?? 1
  if (predictor > 1) return null

  const cs = ctx.lookupMaybe(dict.get(PDFName.of('ColorSpace')), PDFName)?.decodeText()
  if (cs !== 'DeviceRGB' && cs !== 'DeviceGray') return null
  if (dict.has(PDFName.of('SMask')) || dict.has(PDFName.of('Mask'))) return null
  if (dict.has(PDFName.of('Decode'))) return null
  if (ctx.lookupMaybe(dict.get(PDFName.of('ImageMask')), PDFBool)?.asBoolean()) return null

  const bpc = ctx.lookupMaybe(dict.get(PDFName.of('BitsPerComponent')), PDFNumber)?.asNumber()
  const width = ctx.lookupMaybe(dict.get(PDFName.of('Width')), PDFNumber)?.asNumber()
  const height = ctx.lookupMaybe(dict.get(PDFName.of('Height')), PDFNumber)?.asNumber()
  if (bpc !== 8 || !width || !height || width <= 0 || height <= 0) return null
  if (width * height > MAX_PIXELS) return null

  let decoded: Uint8Array
  try {
    decoded = inflateSync(stream.contents)
  } catch {
    return null
  }
  const channels = cs === 'DeviceGray' ? 1 : 3
  // Exact: surplus bytes mean the stream is not the colour data this encoder assumes
  if (decoded.length !== width * height * channels) return null

  return bmpEncode(decoded, width, height, cs === 'DeviceGray')
}

/** Dedupe key for an /XObject entry: refs compare by tag, direct objects by serialization */
function xobjectKey(value: PDFObject): string {
  return value instanceof PDFRef ? value.tag : `direct:${value.toString()}`
}

function walkXObjects(
  ctx: PDFDocument['context'],
  xobjects: PDFDict | undefined,
  pageNo: number,
  depth: number,
  found: Map<string, { stream: PDFRawStream; name: string; pages: number[] }>,
  skipped: { count: number },
): void {
  if (!xobjects || depth > MAX_DEPTH) return
  for (const [name, value] of xobjects.entries()) {
    // lookupMaybe has no PDFRawStream overload, so resolve untyped and narrow: every parsed
    // stream is one, and anything else in the /XObject list is not extractable bytes
    const maybe = ctx.lookup(value)
    if (!(maybe instanceof PDFRawStream)) continue
    const stream: PDFRawStream = maybe
    const dict: PDFDict = stream.dict
    const subtype = ctx.lookupMaybe(dict.get(PDFName.of('Subtype')), PDFName)?.decodeText()
    if (subtype === 'Image') {
      const seen = found.get(xobjectKey(value))
      if (seen) {
        if (!seen.pages.includes(pageNo)) seen.pages.push(pageNo)
      } else {
        found.set(xobjectKey(value), { stream, name: name.decodeText(), pages: [pageNo] })
      }
    } else if (subtype === 'Form') {
      const nested = ctx.lookupMaybe(dict.get(PDFName.of('Resources')), PDFDict)
      walkXObjects(
        ctx,
        nested?.lookupMaybe(PDFName.of('XObject'), PDFDict),
        pageNo,
        depth + 1,
        found,
        skipped,
      )
    } else {
      // Not drawable image data, and not a container either
      skipped.count++
    }
  }
}

/**
 * Single ordered pass over the pages. Each image keeps the pages that draw it; the skip count is
 * the caller's report line, so a file that is mostly predictors says so instead of looking empty.
 */
export async function collectEmbeddedImages(bytes: Uint8Array): Promise<EmbeddedImageScan> {
  const pdfDoc = await PDFDocument.load(bytes, { updateMetadata: false })
  const ctx = pdfDoc.context
  const found = new Map<string, { stream: PDFRawStream; name: string; pages: number[] }>()
  const skipped = { count: 0 }

  pdfDoc.getPages().forEach((page, i) => {
    const resources = ctx.lookupMaybe(page.node.get(PDFName.of('Resources')), PDFDict)
    walkXObjects(
      ctx,
      resources?.lookupMaybe(PDFName.of('XObject'), PDFDict),
      i + 1,
      0,
      found,
      skipped,
    )
  })

  const images: EmbeddedImage[] = []
  for (const { stream, name, pages } of found.values()) {
    const filters = filtersOf(stream.dict as PDFDict, ctx)
    if (filters.length === 1 && filters[0] === 'DCTDecode') {
      images.push({ ext: 'jpg', data: stream.contents, pages, name })
      continue
    }
    const bmp = flateToBitmap(stream, ctx)
    if (bmp) images.push({ ext: 'bmp', data: bmp, pages, name })
    else skipped.count++
  }

  return { images, skipped: skipped.count }
}
