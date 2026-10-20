/** How to divide a document into separate files */
export type SplitMode =
  /** Every `size` pages, the last part holding whatever remains */
  | { kind: 'every'; size: number }
  /** At the given 1-based page numbers: each listed page begins a new file */
  | { kind: 'at'; pages: number[] }

const span = (from: number, to: number) =>
  Array.from({ length: Math.max(0, to - from) }, (_, i) => from + i)

/**
 * Consecutive page ranges (0-based indices) to write as separate documents, in order.
 *
 * Page 1 is always a boundary even when the caller does not list it, so the first part never
 * starts partway into the document. Ranges are kept in ascending order and never overlap, so a
 * split can neither drop nor duplicate a page.
 */
export function splitRanges(total: number, mode: SplitMode): number[][] {
  if (!Number.isFinite(total) || total <= 0) return []
  if (mode.kind === 'every') {
    // A zero or negative step would loop forever or emit nothing; one page is the useful floor
    const size = Math.max(1, Math.floor(mode.size) || 1)
    const parts: number[][] = []
    for (let start = 0; start < total; start += size) {
      parts.push(span(start, Math.min(start + size, total)))
    }
    return parts
  }
  const inRange = mode.pages.filter((p) => Number.isFinite(p) && p >= 1 && p <= total)
  const boundaries = [...new Set([1, ...inRange])].sort((a, b) => a - b)
  return boundaries.map((start, i) => span(start - 1, (boundaries[i + 1] ?? total + 1) - 1))
}
