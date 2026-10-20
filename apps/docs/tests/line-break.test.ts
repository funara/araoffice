import { describe, expect, it } from 'vitest'
import {
  applyKinsoku,
  breakOpportunities,
  lineBreakOpportunities,
} from '../src/renderer/line-break'

/** Break positions as the list of left-hand pieces, which reads better in failures. */
const pieces = (text: string, wordWrap = true) => {
  const out: string[] = []
  let prev = 0
  for (const at of breakOpportunities(text, wordWrap)) {
    out.push(text.slice(prev, at))
    prev = at
  }
  out.push(text.slice(prev))
  return out
}

describe('line-break opportunities (UAX #14 subset + kinsoku)', () => {
  it('breaks after spaces, never before them', () => {
    expect(pieces('Hello world again')).toEqual(['Hello ', 'world ', 'again'])
  })

  it('does not break inside a Latin word', () => {
    expect(pieces('internationalization')).toEqual(['internationalization'])
  })

  it('breaks after a hyphen or slash inside a word', () => {
    expect(pieces('well-known')).toEqual(['well-', 'known'])
    expect(pieces('a/b/c')).toEqual(['a/', 'b/', 'c'])
  })

  it('does not break inside a numeric fraction or a hyphenated number', () => {
    expect(pieces('3/4')).toEqual(['3/4'])
    expect(pieces('SOP-260904')).toEqual(['SOP-260904'])
  })

  it('breaks between East Asian characters', () => {
    expect(pieces('中文測試')).toEqual(['中', '文', '測', '試'])
  })

  it('does not break inside a Latin run embedded in CJK', () => {
    expect(pieces('中文English混合')).toEqual(['中', '文', 'English', '混', '合'])
  })

  it('never starts a line with a character kinsoku forbids there', () => {
    // 、 and 。 may not begin a line, so no opportunity may sit directly before them
    const breaks = breakOpportunities('這是、測試。')
    for (const at of breaks) {
      expect('、。'.includes('這是、測試。'[at])).toBe(false)
    }
  })

  it('never ends a line with an opening bracket', () => {
    const text = '（括號）內容'
    const breaks = breakOpportunities(text)
    for (const at of breaks) {
      expect('（'.includes(text[at - 1])).toBe(false)
    }
  })

  it('allows a break between two East Asian characters when word wrap is off', () => {
    const text = '中文測試'
    expect(applyKinsoku(text, [], false).length).toBe(3)
    expect(applyKinsoku(text, [], true).length).toBe(0)
  })

  it('reports ascending, unique, in-range offsets', () => {
    const text = 'The 中文 test, with punctuation; and more.'
    const breaks = lineBreakOpportunities(text)
    const sorted = [...breaks].sort((a, b) => a - b)
    expect(breaks).toEqual(sorted)
    expect(new Set(breaks).size).toBe(breaks.length)
    for (const at of breaks) {
      expect(at).toBeGreaterThan(0)
      expect(at).toBeLessThan(text.length)
    }
  })

  it('never returns an offset inside a surrogate pair', () => {
    const text = 'a \u{1F600} b'
    for (const at of breakOpportunities(text)) {
      const prev = text.charCodeAt(at - 1)
      expect(prev >= 0xd800 && prev <= 0xdbff).toBe(false)
    }
  })
})
