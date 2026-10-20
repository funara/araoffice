import { describe, expect, it } from 'vitest'
import { HeuristicMetrics, lineHeightFactor } from '../src/renderer/line-metrics'
import {
  BundledFontMetrics,
  bundledFamilyFor,
  realLineFactor,
  resetDefaultMetrics,
  setDefaultMetrics,
  variantSuffix,
} from '../src/renderer/font-metrics'
import { activeMetrics } from '../src/renderer/line-metrics'
import { bundledFontMetrics, loadBundledFontFile } from './helpers/bundled-fonts'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const FONT_DIR = join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/fonts')
const real = bundledFontMetrics()
const heuristic = new HeuristicMetrics()

const style = (fontFamily: string, pt: number, bold = false) => ({
  fontFamily,
  fontSizePx: (pt * 96) / 72,
  bold,
  italic: false,
})

describe('bundled font metrics', () => {
  it('resolves the Word families that have a metric-compatible bundled face', () => {
    expect(bundledFamilyFor('Calibri')).toBe('Carlito')
    expect(bundledFamilyFor('Cambria')).toBe('Caladea')
    expect(bundledFamilyFor('Times New Roman')).toBe('LiberationSerif')
    expect(bundledFamilyFor('Arial')).toBe('LiberationSans')
    expect(bundledFamilyFor('Courier New')).toBe('LiberationMono')
    expect(bundledFamilyFor('Yu Gothic')).toBeUndefined()
  })

  it('names the variant files the way the font directory does', () => {
    expect(variantSuffix(false, false)).toBe('Regular')
    expect(variantSuffix(true, false)).toBe('Bold')
    expect(variantSuffix(false, true)).toBe('Italic')
    expect(variantSuffix(true, true)).toBe('BoldItalic')
  })

  it('reads a real single-line height instead of the name table', () => {
    for (const family of ['Calibri', 'Cambria', 'Times New Roman', 'Arial']) {
      const m = real.metrics(style(family, 11))
      const factor = m.lineHeight / style(family, 11).fontSizePx
      expect(factor).toBeGreaterThan(1.0)
      expect(factor).toBeLessThan(1.4)
      // the name table is a per-family constant; the real file must agree closely
      expect(Math.abs(factor - lineHeightFactor(family))).toBeLessThan(0.02)
    }
  })

  it('measures advances that differ from the heuristic fractions', () => {
    // Calibri's real advance is materially narrower than the 0.52em default the
    // heuristic used, which is what pushed paragraphs onto extra lines.
    const sample = 'The quick brown fox jumps over the lazy dog'
    const realW = real.measure(sample, style('Calibri', 10))
    const heurW = heuristic.measure(sample, style('Calibri', 10))
    expect(realW).toBeLessThan(heurW)
    expect(realW / heurW).toBeGreaterThan(0.85)
  })

  it('measures a monospace face wider than the heuristic default', () => {
    const sample = 'The quick brown fox'
    expect(real.measure(sample, style('Courier New', 10))).toBeGreaterThan(
      heuristic.measure(sample, style('Courier New', 10)),
    )
  })

  it('falls back for a family with no bundled file', () => {
    const s = style('Some Unknown Face', 11)
    expect(real.metrics(s)).toEqual(heuristic.metrics(s))
    expect(real.measure('hello', s)).toBe(heuristic.measure('hello', s))
  })

  it('hands CJK text to the fallback when the Latin file lacks the glyphs', () => {
    const s = style('Calibri', 11)
    expect(real.measure('中文', s)).toBe(heuristic.measure('中文', s))
  })

  it('caches font files: the loader returns the same object for the same path', () => {
    const p = join(FONT_DIR, 'Carlito-Regular.ttf')
    expect(loadBundledFontFile(p)).toBe(loadBundledFontFile(p))
  })

  it('gives the canvas line factor and the pagination model the same number', () => {
    // The invariant that keeps the on-screen line boxes and the modelled ones together:
    // doc-style-css writes realLineFactor(family) into --doc-line-factor, and the model
    // multiplies a font size by the same provider's lineHeight/fontSizePx.
    setDefaultMetrics(bundledFontMetrics())
    try {
      for (const family of ['Calibri', 'Cambria', 'Times New Roman', 'Arial']) {
        const s = style(family, 11)
        const modelled = activeMetrics().metrics(s).lineHeight / s.fontSizePx
        expect(realLineFactor(family)).toBeCloseTo(modelled, 10)
      }
    } finally {
      resetDefaultMetrics()
    }
  })

  it('falls back to the heuristic table when a file is missing', () => {
    const missing = new BundledFontMetrics(() => undefined)
    const s = style('Calibri', 11)
    expect(missing.metrics(s)).toEqual(heuristic.metrics(s))
  })
})
