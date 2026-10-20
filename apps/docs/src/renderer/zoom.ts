/** Document zoom bounds. Zoom is a pure view magnification: pagination never changes with it. */
export const MIN_ZOOM = 50
export const MAX_ZOOM = 200

export function clampZoom(next: number): number {
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, next))
}

/**
 * Symmetric horizontal scroll room for the canvas. Below the fit ratio this is
 * the normal 24px gutter; once the page is wider than the pane it grows so the
 * page can stay centred and both paper edges remain reachable by scrolling
 * (auto margins alone collapse to 0 on overflow).
 */
export function canvasHPad(pageW: number, paneW: number, min = 24): number {
  return Math.max(min, (pageW - paneW) / 2 + min)
}

export interface RectLike {
  left: number
  top: number
}

export interface AnchorInput {
  /** Anchor point in client coordinates (pointer for Ctrl/Cmd+wheel, else viewport centre) */
  anchorX: number
  anchorY: number
  /** Bounding rect of the zoomed canvas at the current scale */
  rect: RectLike
  /** Current scale factor (zoom / 100) */
  factor: number
}

/** Unscaled content-space point sitting under a client-coordinate anchor */
export function anchorContentPoint(input: AnchorInput): { x: number; y: number } {
  const { anchorX, anchorY, rect, factor } = input
  return { x: (anchorX - rect.left) / factor, y: (anchorY - rect.top) / factor }
}

export interface ScrollDeltaInput {
  anchorX: number
  anchorY: number
  /** Canvas rect at the new scale */
  rect: RectLike
  /** New scale factor (zoom / 100) */
  factor: number
  /** Content-space point to keep under the anchor */
  point: { x: number; y: number }
}

/** scrollLeft/scrollTop delta that keeps `point` under the anchor after the scale change */
export function scrollDeltaForAnchor(input: ScrollDeltaInput): { dx: number; dy: number } {
  const { anchorX, anchorY, rect, factor, point } = input
  // scrolling by +delta moves the content by -delta, so the on-screen position of
  // `point` becomes rect.left + point.x * factor - dx; solving = anchorX gives:
  return {
    dx: rect.left + point.x * factor - anchorX,
    dy: rect.top + point.y * factor - anchorY,
  }
}
