import type { Graph } from '@maxgraph/core'

export type RasterMime = 'image/png' | 'image/jpeg'

/**
 * Export visual diagram canvas to standalone SVG string, cropped tightly
 * to diagram bounds with comfortable margin.
 */
export function exportGraphToSvg(graph: Graph, padding = 20): string {
  const container = graph.container
  const svgEl = container?.querySelector('svg')
  if (!svgEl) {
    throw new Error('No SVG element found in diagram canvas')
  }

  const bounds = graph.getGraphBounds()
  const x = Math.floor(bounds.x) - padding
  const y = Math.floor(bounds.y) - padding
  const width = Math.max(100, Math.ceil(bounds.width) + padding * 2)
  const height = Math.max(60, Math.ceil(bounds.height) + padding * 2)

  const clone = svgEl.cloneNode(true) as SVGSVGElement
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg')
  clone.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink')
  clone.setAttribute('version', '1.1')
  clone.setAttribute('viewBox', `${x} ${y} ${width} ${height}`)
  clone.setAttribute('width', String(width))
  clone.setAttribute('height', String(height))
  clone.style.width = `${width}px`
  clone.style.height = `${height}px`
  clone.style.display = 'block'
  clone.style.backgroundColor = '#ffffff'

  // Embed default font CSS if not already present
  const styleEl = document.createElementNS('http://www.w3.org/2000/svg', 'style')
  styleEl.textContent = `
    text { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", sans-serif; }
  `
  clone.insertBefore(styleEl, clone.firstChild)

  return new XMLSerializer().serializeToString(clone)
}

/** Parse a static width/height for the diagram out of an SVG string (viewBox preferred). */
export function svgIntrinsicSize(svgString: string): { width: number; height: number } | null {
  const root = svgString.match(/<svg[^>]*>/)?.[0]
  if (!root) return null

  const consume = (value: string | undefined): number | null => {
    if (!value) return null
    const match = value.trim().match(/^(\d+(?:\.\d+)?|\.\d+)(?:px)?$/i)
    if (!match) return null
    const parsed = Number.parseFloat(match[1])
    return Number.isFinite(parsed) && parsed > 0 ? parsed : null
  }

  const viewBox = root
    .match(/\bviewBox\s*=\s*"([^"]*)"/)?.[1]
    ?.trim()
    .split(/\s+/)
    .map(Number)
  if (
    viewBox &&
    viewBox.length === 4 &&
    viewBox.every(Number.isFinite) &&
    viewBox[2] > 0 &&
    viewBox[3] > 0
  ) {
    return { width: viewBox[2], height: viewBox[3] }
  }

  const width = consume(root.match(/\bwidth\s*=\s*"([^"]*)"/)?.[1])
  const height = consume(root.match(/\bheight\s*=\s*"([^"]*)"/)?.[1])
  return width && height ? { width, height } : null
}

export interface RasterSvgOptions {
  mime: RasterMime
  /** Source pixel dimensions; measured via svgIntrinsicSize when omitted */
  width?: number
  height?: number
  /** Supersampling factor for crisp output (default 2) */
  scale?: number
  /** Canvas background; JPEG has no alpha so a background is always filled (default #ffffff) */
  background?: string
  /** encoder quality for JPEG (default 0.92) */
  quality?: number
}

/** Rasterize an SVG string to a base64 PNG or JPEG data URL, rendered at 2× scale for high-density crispness. */
export function svgStringToRasterDataUrl(
  svgString: string,
  options: RasterSvgOptions,
): Promise<string> {
  const { mime, scale = 2, background = '#ffffff', quality = 0.92 } = options
  const intrinsic = svgIntrinsicSize(svgString)
  const width = options.width ?? intrinsic?.width
  const height = options.height ?? intrinsic?.height
  if (!width || !height) {
    return Promise.reject(new Error('Could not determine SVG dimensions for rasterization'))
  }

  const { promise, resolve, reject } = Promise.withResolvers<string>()
  const img = new Image()
  const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' })
  const url = URL.createObjectURL(blob)

  img.onload = () => {
    URL.revokeObjectURL(url)
    const canvas = document.createElement('canvas')
    canvas.width = width * scale
    canvas.height = height * scale

    const ctx = canvas.getContext('2d')
    if (!ctx) {
      reject(new Error('Canvas 2D context unavailable'))
      return
    }

    ctx.fillStyle = background
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    ctx.scale(scale, scale)
    ctx.drawImage(img, 0, 0, width, height)

    try {
      resolve(canvas.toDataURL(mime, quality))
    } catch (err) {
      reject(err instanceof Error ? err : new Error(String(err)))
    }
  }

  img.onerror = (err) => {
    URL.revokeObjectURL(url)
    reject(new Error(`Failed to render SVG to image: ${err}`))
  }

  img.src = url
  return promise
}

/** Rasterize the visual diagram canvas at 2× scale as a base64 PNG or JPEG data URL. */
export function exportGraphToRaster(graph: Graph, mime: RasterMime, padding = 20): Promise<string> {
  const svgString = exportGraphToSvg(graph, padding)
  const bounds = graph.getGraphBounds()
  return svgStringToRasterDataUrl(svgString, {
    mime,
    width: Math.max(100, Math.ceil(bounds.width) + padding * 2),
    height: Math.max(60, Math.ceil(bounds.height) + padding * 2),
  })
}
