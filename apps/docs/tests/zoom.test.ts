import { describe, expect, it } from 'vitest'
import {
  anchorContentPoint,
  canvasHPad,
  clampZoom,
  MAX_ZOOM,
  MIN_ZOOM,
  scrollDeltaForAnchor,
} from '../src/renderer/zoom'

describe('clampZoom', () => {
  it('clamps to the zoom bounds', () => {
    expect(clampZoom(0)).toBe(MIN_ZOOM)
    expect(clampZoom(1000)).toBe(MAX_ZOOM)
    expect(clampZoom(100)).toBe(100)
    expect(clampZoom(106)).toBe(106)
  })
})

describe('canvasHPad', () => {
  it('is the normal gutter while the page fits', () => {
    expect(canvasHPad(800, 1000)).toBe(24)
    expect(canvasHPad(1000, 1000)).toBe(24)
  })

  it('grows symmetrically once the page overflows, leaving an equal gutter at the extremes', () => {
    // 1200 page in a 1000 pane: overflow 200, so 100 of scroll room each side (+ the 24 gutter)
    expect(canvasHPad(1200, 1000)).toBe(124)
  })
})

describe('anchorContentPoint', () => {
  it('recovers the unscaled content point under the anchor', () => {
    expect(
      anchorContentPoint({ anchorX: 110, anchorY: 60, rect: { left: 10, top: 20 }, factor: 2 }),
    ).toEqual({ x: 50, y: 20 })
  })

  it('is the raw offset at 100%', () => {
    expect(
      anchorContentPoint({ anchorX: 110, anchorY: 60, rect: { left: 10, top: 20 }, factor: 1 }),
    ).toEqual({ x: 100, y: 40 })
  })
})

describe('scrollDeltaForAnchor', () => {
  it('keeps the anchored point under the cursor when zooming in', () => {
    const point = anchorContentPoint({
      anchorX: 400,
      anchorY: 300,
      rect: { left: 0, top: 0 },
      factor: 1,
    })
    const { dx, dy } = scrollDeltaForAnchor({
      anchorX: 400,
      anchorY: 300,
      rect: { left: 0, top: 0 },
      factor: 2,
      point,
    })
    expect(dx).toBe(400)
    expect(dy).toBe(300)
    // the on-screen position after scrolling by the delta is rect.left + point.x * factor - dx
    expect(0 + point.x * 2 - dx).toBe(400)
    expect(0 + point.y * 2 - dy).toBe(300)
  })

  it('is a no-op when the scale is unchanged', () => {
    expect(
      scrollDeltaForAnchor({
        anchorX: 500,
        anchorY: 200,
        rect: { left: 500 - 123, top: 200 - 45 },
        factor: 1,
        point: { x: 123, y: 45 },
      }),
    ).toEqual({ dx: 0, dy: 0 })
  })

  it('accounts for a canvas rect that shifted during re-layout', () => {
    expect(
      scrollDeltaForAnchor({
        anchorX: 200,
        anchorY: 100,
        rect: { left: -30, top: 0 },
        factor: 2,
        point: { x: 100, y: 50 },
      }),
    ).toEqual({ dx: -30, dy: 0 })
  })
})
