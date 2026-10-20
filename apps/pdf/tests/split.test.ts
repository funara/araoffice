import { describe, expect, it } from 'vitest'
import { splitRanges } from '../src/renderer/split'

/** Flatten to 0-based page indices: a correct split must reproduce 0..total-1 exactly once */
const flatten = (parts: number[][]) => parts.flat()

describe('splitRanges', () => {
  it('splits into equal parts with a short remainder last', () => {
    expect(splitRanges(7, { kind: 'every', size: 3 })).toEqual([[0, 1, 2], [3, 4, 5], [6]])
  })

  it('yields one part when the step exceeds the document', () => {
    expect(splitRanges(3, { kind: 'every', size: 10 })).toEqual([[0, 1, 2]])
  })

  it('treats a zero or negative step as one page rather than looping forever', () => {
    expect(splitRanges(3, { kind: 'every', size: 0 })).toEqual([[0], [1], [2]])
    expect(splitRanges(3, { kind: 'every', size: -5 })).toEqual([[0], [1], [2]])
    expect(splitRanges(3, { kind: 'every', size: 2.7 })).toEqual([[0, 1], [2]])
  })

  it('splits at the listed pages, always starting a part at page 1', () => {
    // 1-based 4 and 6 begin new files: pages 1-3, 4-5, 6-8
    expect(splitRanges(8, { kind: 'at', pages: [4, 6] })).toEqual([
      [0, 1, 2],
      [3, 4],
      [5, 6, 7],
    ])
  })

  it('ignores a listed page 1 instead of emitting an empty part', () => {
    expect(splitRanges(4, { kind: 'at', pages: [1, 3] })).toEqual([
      [0, 1],
      [2, 3],
    ])
  })

  it('sorts, de-duplicates and drops out-of-range boundaries', () => {
    // Unsorted, with a duplicate 4 and two impossible pages (0 is below, 99 above)
    expect(splitRanges(5, { kind: 'at', pages: [4, 3, 4, 0, 99] })).toEqual([[0, 1], [2], [3, 4]])
  })

  it('keeps every page exactly once, whatever the mode', () => {
    for (const total of [1, 2, 5, 12, 13]) {
      for (const parts of [
        splitRanges(total, { kind: 'every', size: 4 }),
        splitRanges(total, { kind: 'at', pages: [2, 5, 9] }),
      ]) {
        expect(flatten(parts), `total=${total}`).toEqual(Array.from({ length: total }, (_, i) => i))
      }
    }
  })

  it('returns no parts for an empty or invalid document', () => {
    expect(splitRanges(0, { kind: 'every', size: 2 })).toEqual([])
    expect(splitRanges(-3, { kind: 'at', pages: [1] })).toEqual([])
    expect(splitRanges(Number.NaN, { kind: 'every', size: 2 })).toEqual([])
  })

  it('handles a single listed page in a one-page document', () => {
    expect(splitRanges(1, { kind: 'at', pages: [1] })).toEqual([[0]])
  })
})
