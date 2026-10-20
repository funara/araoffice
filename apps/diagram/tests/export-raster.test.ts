import { describe, expect, it } from 'vitest'
import { svgIntrinsicSize } from '../src/renderer/maxgraph/exportDiagram'

describe('svgIntrinsicSize', () => {
  it('prefers the viewBox', () => {
    const svg =
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 180" width="640px" height="360px"></svg>'
    expect(svgIntrinsicSize(svg)).toEqual({ width: 320, height: 180 })
  })

  it('falls back to width/height attributes', () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="800" height="600"></svg>'
    expect(svgIntrinsicSize(svg)).toEqual({ width: 800, height: 600 })
  })

  it('accepts px suffixes and floats', () => {
    const svg = '<svg width="102.5px" height="68.25px"></svg>'
    expect(svgIntrinsicSize(svg)).toEqual({ width: 102.5, height: 68.25 })
  })

  it('rejects roots without usable dimensions', () => {
    expect(svgIntrinsicSize('<svg xmlns="http://www.w3.org/2000/svg"></svg>')).toBeNull()
    expect(svgIntrinsicSize('<svg width="100%" height="100%"></svg>')).toBeNull()
    expect(svgIntrinsicSize('not svg')).toBeNull()
  })

  it('falls back to attributes when the viewBox is degenerate', () => {
    expect(svgIntrinsicSize('<svg viewBox="0 0 0 9" width="1" height="9">')).toEqual({
      width: 1,
      height: 9,
    })
  })
})
