/**
 * Real font metrics for line breaking.
 *
 * Replaces the hardcoded per-character advance fractions (`charAdvanceEm`) and the
 * family-name line-height table (`lineHeightFactor`) with metrics read from actual
 * font data, so a paragraph's line count and line heights come from the same numbers
 * the renderer will use.
 *
 * Two real backends, one interface:
 *  - {@link BundledFontMetrics}: opentype.js over the metric-compatible TTFs bundled in
 *    `fonts/` (Carlito/Caladea/Liberation). Deterministic, DOM-free, works headlessly —
 *    the parity harness and any Node consumer use this.
 *  - {@link CanvasFontMetrics}: measured through the browser's own text engine with the
 *    same CSS family chain the canvas renders with. Exact for what the DOM will lay out,
 *    so the pagination model and the on-screen line boxes cannot drift apart.
 *
 * Both fall back to {@link HeuristicMetrics} for families with no metric-compatible
 * equivalent (CJK, and any font we have no file for), where the generic 1.0em CJK
 * advance and the name-table factor remain the best available estimate.
 */
import {
  HeuristicMetrics,
  cssFontFamily,
  lineHeightFactor,
  setMetricsResolver,
  type FontMetrics,
  type FontMetricsProvider,
  type RunStyle,
} from './line-metrics'

/** The subset of opentype.js's Font this module needs (keeps the runtime dependency optional). */
export interface OpentypeFontLike {
  unitsPerEm: number
  ascender: number
  descender: number
  /** hhea line gap; folded into the ascent for single-line spacing */
  lineGap?: number
  getAdvanceWidth(text: string, fontSize: number): number
  charToGlyphIndex?(char: string): number
}

/** Word font family → bundled metric-compatible font family shipped in `fonts/`. */
export const BUNDLED_FAMILY_MAP: ReadonlyArray<readonly [RegExp, string]> = [
  [/calibri|carlito/i, 'Carlito'],
  [/cambria|caladea/i, 'Caladea'],
  [/times|liberation ?serif/i, 'LiberationSerif'],
  [/^arial|helvetica|liberation ?sans/i, 'LiberationSans'],
  [/courier|liberation ?mono/i, 'LiberationMono'],
]

/** Variant suffix used by the bundled font file names. */
export function variantSuffix(bold: boolean, italic: boolean): string {
  return bold && italic ? 'BoldItalic' : bold ? 'Bold' : italic ? 'Italic' : 'Regular'
}

/** The bundled family that stands in for a Word family, or undefined when there is none. */
export function bundledFamilyFor(fontFamily: string): string | undefined {
  return BUNDLED_FAMILY_MAP.find(([re]) => re.test(fontFamily))?.[1]
}

/**
 * Metrics from real font files. `fileFor` resolves a bundled family + variant to font
 * data; the caller supplies it so this module needs no filesystem or network access.
 */
export class BundledFontMetrics implements FontMetricsProvider {
  private readonly cache = new Map<string, OpentypeFontLike | null>()

  constructor(
    private readonly fileFor: (
      family: string,
      bold: boolean,
      italic: boolean,
    ) => OpentypeFontLike | undefined,
    private readonly fallback: FontMetricsProvider = new HeuristicMetrics(),
  ) {}

  private font(style: RunStyle): OpentypeFontLike | undefined {
    const family = bundledFamilyFor(style.fontFamily)
    if (!family) return undefined
    const key = family + '/' + variantSuffix(style.bold, style.italic)
    const hit = this.cache.get(key)
    if (hit !== undefined) return hit ?? undefined
    let font: OpentypeFontLike | undefined
    try {
      font = this.fileFor(family, style.bold, style.italic)
    } catch {
      font = undefined
    }
    this.cache.set(key, font ?? null)
    return font
  }

  metrics(style: RunStyle): FontMetrics {
    const font = this.font(style)
    if (!font) return this.fallback.metrics(style)
    const scale = style.fontSizePx / font.unitsPerEm
    // Word's single line spacing is the font's hhea ascent + descent + line gap.
    const ascent = (font.ascender + (font.lineGap ?? 0)) * scale
    const descent = Math.abs(font.descender) * scale
    return { ascent, descent, lineHeight: ascent + descent }
  }

  measure(text: string, style: RunStyle): number {
    const font = this.font(style)
    if (!font) return this.fallback.measure(text, style)
    // A glyph the font lacks (CJK inside a Latin face) would measure as .notdef:
    // hand the whole run to the fallback, whose 1.0em CJK advance is correct.
    if (font.charToGlyphIndex) {
      for (const ch of text) {
        if (ch !== '\n' && font.charToGlyphIndex(ch) === 0)
          return this.fallback.measure(text, style)
      }
    }
    try {
      return font.getAdvanceWidth(text, style.fontSizePx)
    } catch {
      return this.fallback.measure(text, style)
    }
  }
}

/**
 * Metrics measured through the browser's own text engine, using the same CSS family
 * chain the canvas renders with. Because it is the same shaper that lays out the DOM,
 * the pagination model cannot disagree with the on-screen line boxes about widths.
 */
export class CanvasFontMetrics implements FontMetricsProvider {
  private ctx: CanvasRenderingContext2D | null | undefined
  private readonly cache = new Map<string, FontMetrics>()

  constructor(private readonly fallback: FontMetricsProvider = new HeuristicMetrics()) {}

  private context(): CanvasRenderingContext2D | null {
    if (this.ctx !== undefined) return this.ctx
    try {
      this.ctx =
        typeof document === 'undefined'
          ? null
          : (document.createElement('canvas').getContext('2d') ?? null)
    } catch {
      this.ctx = null
    }
    return this.ctx
  }

  private apply(style: RunStyle): CanvasRenderingContext2D | null {
    const ctx = this.context()
    if (!ctx) return null
    const weight = style.bold ? 'bold ' : ''
    const slant = style.italic ? 'italic ' : ''
    ctx.font = `${slant}${weight}${style.fontSizePx}px ${cssFontFamily(style.fontFamily)}`
    return ctx
  }

  metrics(style: RunStyle): FontMetrics {
    const key = `${style.fontFamily}|${style.fontSizePx}|${style.bold ? 1 : 0}|${style.italic ? 1 : 0}`
    const hit = this.cache.get(key)
    if (hit) return hit
    const ctx = this.apply(style)
    if (!ctx) return this.fallback.metrics(style)
    const m = ctx.measureText('Hxy')
    const ascent = m.fontBoundingBoxAscent
    const descent = m.fontBoundingBoxDescent
    let out: FontMetrics
    if (typeof ascent === 'number' && typeof descent === 'number' && ascent + descent > 0) {
      out = { ascent, descent, lineHeight: ascent + descent }
    } else {
      // Older engines: no font-bounding-box metrics; the name table is the best estimate.
      out = this.fallback.metrics(style)
    }
    this.cache.set(key, out)
    return out
  }

  measure(text: string, style: RunStyle): number {
    const ctx = this.apply(style)
    if (!ctx) return this.fallback.measure(text, style)
    try {
      return ctx.measureText(text).width
    } catch {
      return this.fallback.measure(text, style)
    }
  }
}

// ─── Active provider ────────────────────────────────────────────────────────

/**
 * The provider `computeLineMetrics` uses when the caller passes none.
 *
 * Defaults to browser-measured metrics in a renderer and to the heuristic table in a
 * headless environment; Node consumers (the parity harness) install
 * {@link BundledFontMetrics} explicitly through {@link setDefaultMetrics}.
 */
let activeMetrics: FontMetricsProvider | undefined

/** The metrics provider in force, choosing the environment default on first use. */
export function defaultFontMetrics(): FontMetricsProvider {
  if (!activeMetrics) {
    activeMetrics =
      typeof document === 'undefined' ? new HeuristicMetrics() : new CanvasFontMetrics()
  }
  return activeMetrics
}

/** Install the process-wide metrics provider (tests, harnesses, and explicit wiring). */
export function setDefaultMetrics(provider: FontMetricsProvider | undefined): void {
  activeMetrics = provider
}

/** Reset to the environment default (renderer: browser-measured; headless: heuristic). */
export function resetDefaultMetrics(): void {
  activeMetrics = undefined
}

/**
 * Single-line height of a family as a multiple of its font size, from real metrics.
 *
 * This is the value the canvas writes into `--doc-line-factor`, and the same one the
 * pagination model multiplies a font size by — so a paragraph's on-screen line box and
 * its modelled line box are the same number by construction. Falls back to the family
 * name table when the environment has neither real font files nor a text engine.
 */
export function realLineFactor(fontFamily: string, bold = false, italic = false): number {
  const probe = 100
  const m = defaultFontMetrics().metrics({ fontFamily, fontSizePx: probe, bold, italic })
  const factor = m.lineHeight / probe
  return Number.isFinite(factor) && factor > 0.5 && factor < 3
    ? factor
    : lineHeightFactor(fontFamily)
}

// Installing the resolver here (rather than in line-metrics) keeps the dependency
// one-way: line-metrics stays provider-agnostic, this module decides what "real" means.
setMetricsResolver(defaultFontMetrics)
