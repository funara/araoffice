/**
 * pptx package management — open the zip, archive the original by SHA-256, and read
 * parts and .rels.
 *
 * Byte fidelity: PackageArchive holds the original bytes of every entry; on save,
 * unmodified entries are written back byte-for-byte (handled by the patch layer).
 * This module only handles reading and metadata.
 */
import JSZip from 'jszip'
import { createHash } from 'node:crypto'
import { XMLParser } from 'fast-xml-parser'
import type { SlideSize } from './types'
import { asXmlNode, xmlArray } from './xml-utils'

const relsParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  isArray: (name) => name === 'Relationship' || name === 'sldId' || name === 'Override',
})
const MAX_ZIP_PARTS = 10_000
const MAX_PART_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
const MAX_TOTAL_UNCOMPRESSED_BYTES = 1.5 * 1024 * 1024 * 1024

function getUncompressedSize(file: JSZip.JSZipObject): number {
  if (typeof file === 'object' && file !== null && '_data' in file) {
    const data = file._data
    if (typeof data === 'object' && data !== null && 'uncompressedSize' in data) {
      const size = data.uncompressedSize
      if (typeof size === 'number') return size
    }
  }
  return 0
}

export function assertZipWithinLimits(zip: JSZip): void {
  const files = Object.values(zip.files).filter((f) => !f.dir)
  if (files.length > MAX_ZIP_PARTS) {
    throw new Error(`pptx rejected: ${files.length} parts exceeds the ${MAX_ZIP_PARTS} limit`)
  }
  let total = 0
  for (const file of files) {
    const size = getUncompressedSize(file)
    if (size > MAX_PART_UNCOMPRESSED_BYTES) {
      throw new Error(
        `pptx rejected: part ${file.name} declares ${size} uncompressed bytes (limit ${MAX_PART_UNCOMPRESSED_BYTES})`,
      )
    }
    if (size > 0) total += size
  }
  if (total > MAX_TOTAL_UNCOMPRESSED_BYTES) {
    throw new Error(
      `pptx rejected: total uncompressed size ${total} exceeds the ${MAX_TOTAL_UNCOMPRESSED_BYTES} limit`,
    )
  }
}

export interface Relationship {
  id: string
  type: string
  target: string
  targetMode?: string
}

const EMPTY_RELS: ReadonlyMap<string, Relationship> = new Map()

function parseSlideDimension(value: unknown): number {
  const dimension = Number(value)
  if (!Number.isInteger(dimension) || dimension <= 0) {
    throw new Error('pptx: invalid slide size')
  }
  return dimension
}

function parseRelsXml(xml: string): ReadonlyMap<string, Relationship> {
  const rels = new Map<string, Relationship>()
  const doc = asXmlNode(relsParser.parse(xml))
  const list = asXmlNode(doc.Relationships).Relationship
  for (const r of xmlArray(list)) {
    const id = String(r['@_Id'] ?? '')
    rels.set(id, {
      id,
      type: String(r['@_Type'] ?? ''),
      target: String(r['@_Target'] ?? ''),
      ...(r['@_TargetMode'] != null ? { targetMode: String(r['@_TargetMode']) } : {}),
    })
  }
  return rels
}

export class PackageArchive {
  private constructor(
    private readonly zip: JSZip,
    /** Original bytes of every entry, keyed by path inside the zip */
    readonly entries: Map<string, Uint8Array>,
    readonly originalHash: string,
    private readonly relsCache = new Map<string, ReadonlyMap<string, Relationship>>(),
  ) {}

  static async open(bytes: Uint8Array): Promise<PackageArchive> {
    const originalHash = createHash('sha256').update(bytes).digest('hex')
    const zip = await JSZip.loadAsync(bytes)
    assertZipWithinLimits(zip)
    // ponytail: lazy inflate skipped — entries is a public Map<string, Uint8Array>
    // consumed as raw bytes by the save path (buildZip), slide-transfer, theme-apply
    // and copy-across-decks; a lazy wrapper would ripple through all of them.
    const entries = new Map<string, Uint8Array>()
    const names = Object.keys(zip.files)
    for (const name of names) {
      const file = zip.files[name]
      if (file.dir) continue
      entries.set(name, await file.async('uint8array'))
    }
    return new PackageArchive(zip, entries, originalHash)
  }

  has(path: string): boolean {
    return this.entries.has(path)
  }

  /** Read a part as a UTF-8 string (for XML parts). */
  readText(path: string): string | null {
    const bytes = this.entries.get(path)
    if (!bytes) return null
    return Buffer.from(bytes).toString('utf8')
  }

  readBytes(path: string): Uint8Array | null {
    return this.entries.get(path) ?? null
  }

  /**
   * Read a part's relationships file. partPath e.g. 'ppt/slides/slide1.xml' →
   * 'ppt/slides/_rels/slide1.xml.rels'.
   *
   * Memoized by rels XML content (not path — entries are rewritten in place by edit
   * operations like hyperlink insertion, then re-read). Identical rels across many
   * slides share one parse. Returned view is read-only; do not mutate.
   */
  readRels(partPath: string): ReadonlyMap<string, Relationship> {
    const xml = this.readText(relsPathFor(partPath))
    if (!xml) return EMPTY_RELS
    let cached = this.relsCache.get(xml)
    if (!cached) {
      cached = parseRelsXml(xml)
      this.relsCache.set(xml, cached)
    }
    return cached
  }

  /**
   * Read the presentation's slide size and the slide part paths in order.
   */
  readPresentation(): { size: SlideSize; slidePaths: string[] } {
    const presXml = this.readText('ppt/presentation.xml')
    if (!presXml) throw new Error('pptx: missing ppt/presentation.xml')

    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      isArray: (name) => name === 'p:sldId',
    })
    const pres = asXmlNode(parser.parse(presXml))
    const rootRaw = pres['p:presentation'] ?? pres.presentation
    if (!rootRaw) throw new Error('pptx: malformed presentation.xml')
    const root = asXmlNode(rootRaw)

    // Slide size
    const szRaw = root['p:sldSz'] ?? root.sldSz
    const sz = szRaw ? asXmlNode(szRaw) : null
    const size: SlideSize = sz
      ? {
          cx: parseSlideDimension(sz['@_cx']),
          cy: parseSlideDimension(sz['@_cy']),
        }
      : { cx: 9144000, cy: 6858000 }

    // Slide order: presentation.xml.rels maps r:id to slide parts
    const rels = this.readRels('ppt/presentation.xml')
    const sldIdLst = asXmlNode(root['p:sldIdLst'] ?? root.sldIdLst)
    const slidePaths: string[] = []
    for (const id of xmlArray(sldIdLst['p:sldId'])) {
      const rId = id['@_r:id'] ?? id['@_id']
      if (!rId) continue
      const rel = rels.get(String(rId))
      if (!rel || rel.targetMode === 'External') continue
      slidePaths.push(resolveTarget('ppt/presentation.xml', rel.target))
    }
    return { size, slidePaths }
  }

  /** Resolve a slide's layout / master part paths (via the rels chain). */
  resolveSlideChain(slidePath: string): {
    layoutPath?: string
    masterPath?: string
    themePath?: string
  } {
    const slideRels = this.readRels(slidePath)
    let layoutPath: string | undefined
    for (const rel of slideRels.values()) {
      if (rel.type.endsWith('/slideLayout') && rel.targetMode !== 'External') {
        layoutPath = resolveTarget(slidePath, rel.target)
        break
      }
    }
    let masterPath: string | undefined
    let themePath: string | undefined
    if (layoutPath) {
      const layoutRels = this.readRels(layoutPath)
      for (const rel of layoutRels.values()) {
        if (rel.type.endsWith('/slideMaster') && rel.targetMode !== 'External') {
          masterPath = resolveTarget(layoutPath, rel.target)
          break
        }
      }
    }
    if (masterPath) {
      const masterRels = this.readRels(masterPath)
      for (const rel of masterRels.values()) {
        if (rel.type.endsWith('/theme') && rel.targetMode !== 'External') {
          themePath = resolveTarget(masterPath, rel.target)
          break
        }
      }
    }
    return { layoutPath, masterPath, themePath }
  }
}

/** 'ppt/slides/slide1.xml' → 'ppt/slides/_rels/slide1.xml.rels' */
export function relsPathFor(partPath: string): string {
  const idx = partPath.lastIndexOf('/')
  const dir = idx >= 0 ? partPath.slice(0, idx) : ''
  const file = idx >= 0 ? partPath.slice(idx + 1) : partPath
  return `${dir ? dir + '/' : ''}_rels/${file}.rels`
}

/**
 * Resolve a relative target into an absolute path inside the zip.
 * basePart is the referencing part's path (its directory is the base); target may be
 * something like '../slideLayouts/slideLayout1.xml'.
 */
export function resolveTarget(basePart: string, target: string): string {
  const baseDir = basePart.slice(0, basePart.lastIndexOf('/'))
  const parts = target.startsWith('/') ? [] : baseDir.split('/').filter(Boolean)
  for (const seg of target.split('/')) {
    if (seg === '.' || seg === '') continue
    if (seg === '..') {
      if (!parts.length) throw new Error('pptx: relationship target escapes archive root')
      parts.pop()
    } else {
      parts.push(seg)
    }
  }
  return parts.join('/')
}
