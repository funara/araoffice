/**
 * Real font metrics from the metric-compatible fonts bundled with the app
 * (`apps/docs/src/renderer/fonts`), for headless consumers: the pagination parity
 * harness and any Node-side layout check.
 *
 * The baseline PDFs are laid out by LibreOffice, which ships the same metric-compatible
 * substitute families (Calibri→Carlito, Cambria→Caladea, Times/Arial/Courier→Liberation).
 * Reading our own copies of those files gives the same advances and line heights as the
 * baseline engine, on every platform — no LibreOffice or Word installation required.
 *
 * Coverage is Latin; CJK families have no file here and fall through to
 * {@link HeuristicMetrics}, whose 1.0em CJK advance is exact (CJK advances are the em
 * square). Unmatched families and unparseable files fall back the same way.
 */
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import * as opentype from 'opentype.js'
import { HeuristicMetrics, type FontMetricsProvider } from '../../src/renderer/line-metrics'
import {
  BundledFontMetrics,
  variantSuffix,
  type OpentypeFontLike,
} from '../../src/renderer/font-metrics'

const __dirname = dirname(fileURLToPath(import.meta.url))

/** The app's bundled metric-compatible fonts. */
export const BUNDLED_FONT_DIR = join(__dirname, '../../src/renderer/fonts')

const cache = new Map<string, OpentypeFontLike | null>()

/** Parse one bundled font file; null when absent or unreadable. */
export function loadBundledFontFile(path: string): OpentypeFontLike | null {
  const hit = cache.get(path)
  if (hit !== undefined) return hit
  let wrapped: OpentypeFontLike | null = null
  if (existsSync(path)) {
    try {
      const buf = readFileSync(path)
      const font = opentype.parse(
        buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer,
      )
      wrapped = {
        unitsPerEm: font.unitsPerEm,
        ascender: font.ascender,
        descender: font.descender,
        // Single line spacing = hhea ascent + descent + lineGap (Word's rule).
        lineGap: (font.tables.hhea as { lineGap?: number } | undefined)?.lineGap ?? 0,
        getAdvanceWidth: (t: string, s: number) => font.getAdvanceWidth(t, s),
        charToGlyphIndex: (ch: string) => font.charToGlyphIndex(ch),
      }
    } catch {
      wrapped = null
    }
  }
  cache.set(path, wrapped)
  return wrapped
}

/**
 * Real metrics from the bundled fonts; the heuristic table when a family has no file.
 *
 * `DOCS_METRICS=heuristic` forces the old name-table/fraction metrics so the harness can
 * measure the two providers against the same baseline.
 */
export function bundledFontMetrics(): FontMetricsProvider {
  if (process.env.DOCS_METRICS === 'heuristic') return new HeuristicMetrics()
  if (!existsSync(BUNDLED_FONT_DIR)) return new HeuristicMetrics()
  return new BundledFontMetrics(
    (family, bold, italic) =>
      loadBundledFontFile(join(BUNDLED_FONT_DIR, `${family}-${variantSuffix(bold, italic)}.ttf`)) ??
      undefined,
  )
}

/** True when the bundled font directory is available (i.e. real metrics are usable). */
export function bundledFontsAvailable(): boolean {
  return existsSync(BUNDLED_FONT_DIR)
}
