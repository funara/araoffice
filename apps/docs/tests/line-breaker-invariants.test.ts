/**
 * Line-breaker invariants and metric wiring.
 *
 * The 27-document parity corpus scores page *starts* with a fuzzy text match, which is
 * too coarse to notice a change in how a paragraph breaks: an 8.6% width error in the
 * metrics still landed on the same page starts. These tests pin the breaker itself.
 */
import { describe, expect, it } from 'vitest'
import {
  HeuristicMetrics,
  computeLineMetrics,
  simulateLines,
  type FontMetricsProvider,
} from '../src/renderer/line-metrics'
import { bundledFontMetrics } from './helpers/bundled-fonts'

const metrics = bundledFontMetrics()
const heuristic = new HeuristicMetrics()

const runs = (text: string, fontFamily = 'Calibri', pt = 11) => [
  { text, fontFamily, sizeHalfPoints: pt * 2 },
]
const lines = (text: string, width: number, m: FontMetricsProvider = metrics) =>
  simulateLines(runs(text), width, m, 11, 'Calibri')

/** Every character the paragraph contributed, ignoring the whitespace a break drops. */
const content = (s: string) => s.replace(/\s+/g, '')

const LATIN =
  'The quick brown fox jumps over the lazy dog and then keeps running well past the edge of the column.'
const CJK = '本文件說明雲端遷移的標準作業程序，包含前置檢查、執行步驟與回復方式。'
const MIXED = 'Platform 平台 DAG Name 名稱與 Script 指令的對應關係如下所述。'

describe('line breaker invariants', () => {
  for (const [name, text] of [
    ['Latin', LATIN],
    ['CJK', CJK],
    ['mixed script', MIXED],
  ] as const) {
    it(`loses no content when breaking ${name} text`, () => {
      for (const width of [120, 260, 420, 624]) {
        expect(
          content(
            lines(text, width)
              .map((l) => l.text)
              .join(''),
          ),
        ).toBe(content(text))
      }
    })

    it(`never starts a line with a space ${name}`, () => {
      for (const width of [120, 260, 420]) {
        const ls = lines(text, width)
        for (const l of ls.slice(1)) expect(l.text.startsWith(' ')).toBe(false)
      }
    })
  }

  it('never starts a line with a character kinsoku forbids there', () => {
    const text = '這是、測試。內容，包含；各種！標點？以及（括號）。'
    for (const width of [80, 140, 220, 360]) {
      for (const l of lines(text, width)) {
        expect('、。，；！？）】」'.includes(l.text[0] ?? '')).toBe(false)
      }
    }
  })

  it('never ends a line with an opening bracket', () => {
    const text = '說明（括號內的文字很長很長很長）後繼續。'
    for (const width of [80, 140, 220]) {
      for (const l of lines(text, width)) {
        const last = l.text[l.text.length - 1] ?? ''
        expect('（【「'.includes(last)).toBe(false)
      }
    }
  })

  it('never needs more lines when the column gets wider', () => {
    for (const text of [LATIN, CJK, MIXED]) {
      let prev = Number.POSITIVE_INFINITY
      for (let width = 100; width <= 700; width += 25) {
        const n = lines(text, width).length
        expect(n).toBeLessThanOrEqual(prev)
        prev = n
      }
    }
  })

  it('gives every line a positive height', () => {
    for (const text of [LATIN, CJK, MIXED]) {
      for (const l of lines(text, 300)) expect(l.naturalLineH).toBeGreaterThan(0)
    }
  })

  it('breaks a hard line break into its own line', () => {
    expect(lines('first\nsecond', 400).length).toBe(2)
    expect(lines('first\nsecond', 400)[0].text).toBe('first')
  })

  it('keeps a hyphenated word breakable and a numeric fraction intact', () => {
    expect(lines('well-known', 30).length).toBeGreaterThan(1)
    expect(lines('3/4', 30)[0].text).toContain('3/4')
  })
})

describe('metric wiring', () => {
  it('produces different line counts from the heuristic fractions somewhere', () => {
    // A single width can coincide; over a sweep the two providers must diverge, or the
    // real provider is no longer wired in.
    let differing = 0
    for (let width = 120; width <= 700; width += 10) {
      if (lines(LATIN, width).length !== lines(LATIN, width, heuristic).length) differing++
    }
    expect(differing).toBeGreaterThan(3)
  })

  it('measures the bundled Carlito advance the same way every run', () => {
    // A frozen fact about the bundled font file, not about our code: 11pt Carlito over
    // this sample. Guards against the provider silently falling back to heuristics.
    const w = metrics.measure('The quick brown fox jumps over the lazy dog', {
      fontFamily: 'Calibri',
      fontSizePx: (11 * 96) / 72,
      bold: false,
      italic: false,
    })
    expect(w).toBeGreaterThan(240)
    expect(w).toBeLessThan(320)
    expect(w).not.toBe(
      heuristic.measure('The quick brown fox jumps over the lazy dog', {
        fontFamily: 'Calibri',
        fontSizePx: (11 * 96) / 72,
        bold: false,
        italic: false,
      }),
    )
  })

  it('reports a line count consistent with the measured advance', () => {
    const width = 300
    const ls = lines(LATIN, width)
    // greedy packing can waste space, but it can never beat the ideal fit
    const total = metrics.measure(LATIN, {
      fontFamily: 'Calibri',
      fontSizePx: (11 * 96) / 72,
      bold: false,
      italic: false,
    })
    expect(ls.length).toBeGreaterThanOrEqual(Math.ceil(total / width) - 1)
    expect(ls.length).toBeLessThanOrEqual(Math.ceil(total / width) + 3)
    expect(computeLineMetrics({ runs: runs(LATIN), availWidthPx: width, metrics }).lineCount).toBe(
      ls.length,
    )
  })
})
