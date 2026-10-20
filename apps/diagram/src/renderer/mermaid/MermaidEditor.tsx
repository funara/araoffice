import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent, WheelEvent as ReactWheelEvent } from 'react'
import { IconCopy } from '../components/icons'
import { renderMermaidDiagram } from './renderMermaid'

export const MERMAID_ZOOM_MIN = 25
export const MERMAID_ZOOM_MAX = 200
export const MERMAID_ZOOM_STEP = 15

export function clampMermaidZoom(value: number): number {
  return Math.min(MERMAID_ZOOM_MAX, Math.max(MERMAID_ZOOM_MIN, value))
}

export interface MermaidEditorProps {
  code: string
  onChange: (code: string) => void
  zoom: number
  onZoomChange: (zoom: number) => void
}

interface MermaidPreviewProps {
  svg: string
  error: string | null
  isRendering: boolean
  zoom: number
  onZoomChange: (zoom: number) => void
}

interface MermaidPreviewToolbarProps {
  svg: string
}

function MermaidPreviewToolbar({ svg }: MermaidPreviewToolbarProps) {
  const handleCopySvg = async () => {
    if (!svg) return
    try {
      await navigator.clipboard.writeText(svg)
    } catch {
      // ignore clipboard error
    }
  }

  return (
    <div className="mermaid-preview-actions">
      <span className="mermaid-pan-hint">Drag to pan · Ctrl+scroll to zoom</span>
      <button
        type="button"
        className="mermaid-action-btn"
        onClick={handleCopySvg}
        disabled={!svg}
        title="Copy rendered SVG"
      >
        <IconCopy size={14} />
        <span>Copy</span>
      </button>
    </div>
  )
}

interface MermaidPreviewCanvasProps {
  svg: string
  error: string | null
  zoom: number
  onZoomChange: (zoom: number) => void
}

interface SvgBounds {
  x: number
  y: number
  width: number
  height: number
}

function readSvgViewBox(svgEl: SVGSVGElement): SvgBounds | null {
  const values = svgEl.getAttribute('viewBox')?.trim().split(/\s+/).map(Number)
  if (!values || values.length !== 4 || values.some((value) => !Number.isFinite(value))) return null
  const [x, y, width, height] = values
  return width && height ? { x, y, width, height } : null
}

export function mermaidViewport(
  bounds: SvgBounds,
  pan: { x: number; y: number },
  zoom: number,
): string {
  const scale = 100 / zoom
  const width = bounds.width * scale
  const height = bounds.height * scale
  return `${bounds.x + (bounds.width - width) / 2 - pan.x} ${
    bounds.y + (bounds.height - height) / 2 - pan.y
  } ${width} ${height}`
}

function MermaidPreviewCanvas({ svg, error, zoom, onZoomChange }: MermaidPreviewCanvasProps) {
  const [pan, setPan] = useState({ x: 0, y: 0 })
  const [isPanning, setIsPanning] = useState(false)
  const mountRef = useRef<HTMLDivElement>(null)
  const svgRef = useRef<SVGSVGElement | null>(null)
  const boundsRef = useRef<SvgBounds | null>(null)
  const dragRef = useRef<{
    pointerId: number
    startX: number
    startY: number
    startPanX: number
    startPanY: number
    unitsPerPixelX: number
    unitsPerPixelY: number
  } | null>(null)

  useLayoutEffect(() => {
    const mount = mountRef.current
    if (!mount) return
    mount.replaceChildren()
    svgRef.current = null
    boundsRef.current = null
    if (!svg) return

    const doc = new DOMParser().parseFromString(svg, 'image/svg+xml')
    const svgEl = doc.querySelector('svg')
    if (!(svgEl instanceof SVGSVGElement)) return

    svgEl.style.width = '100%'
    svgEl.style.height = '100%'
    mount.append(svgEl)
    svgRef.current = svgEl

    const nativeBounds = readSvgViewBox(svgEl)
    if (!nativeBounds) return
    boundsRef.current = nativeBounds
    setPan({ x: 0, y: 0 })
    svgEl.setAttribute('viewBox', mermaidViewport(nativeBounds, { x: 0, y: 0 }, zoom))

    const frame = requestAnimationFrame(() => {
      if (!svgEl.isConnected) return
      const content = svgEl.getBBox()
      if (!content.width || !content.height) return
      const padding = 16
      const minX = Math.min(nativeBounds.x, content.x) - padding
      const minY = Math.min(nativeBounds.y, content.y) - padding
      const bounds = {
        x: minX,
        y: minY,
        width:
          Math.max(nativeBounds.x + nativeBounds.width, content.x + content.width) - minX + padding,
        height:
          Math.max(nativeBounds.y + nativeBounds.height, content.y + content.height) -
          minY +
          padding,
      }
      boundsRef.current = bounds
      svgEl.setAttribute('viewBox', mermaidViewport(bounds, { x: 0, y: 0 }, zoom))
    })
    return () => cancelAnimationFrame(frame)
  }, [svg])

  useEffect(() => {
    const svgEl = svgRef.current
    const bounds = boundsRef.current
    if (svgEl && bounds) svgEl.setAttribute('viewBox', mermaidViewport(bounds, pan, zoom))
  }, [pan, zoom])

  const handlePointerDown = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (event.button !== 0) return
    const svgEl = svgRef.current
    if (!svgEl) return
    const rect = svgEl.getBoundingClientRect()
    const viewBox = readSvgViewBox(svgEl)
    if (!rect.width || !rect.height || !viewBox) return

    event.currentTarget.setPointerCapture(event.pointerId)
    dragRef.current = {
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      startPanX: pan.x,
      startPanY: pan.y,
      unitsPerPixelX: viewBox.width / rect.width,
      unitsPerPixelY: viewBox.height / rect.height,
    }
    setIsPanning(true)
  }

  const handlePointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    const drag = dragRef.current
    if (!drag || drag.pointerId !== event.pointerId) return
    setPan({
      x: drag.startPanX + (event.clientX - drag.startX) * drag.unitsPerPixelX,
      y: drag.startPanY + (event.clientY - drag.startY) * drag.unitsPerPixelY,
    })
  }

  const handlePointerUp = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (dragRef.current?.pointerId !== event.pointerId) return
    dragRef.current = null
    setIsPanning(false)
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId)
    }
  }

  const handleWheel = (event: ReactWheelEvent<HTMLDivElement>) => {
    if (!event.ctrlKey && !event.metaKey) return
    const delta = event.deltaY || event.deltaX
    if (delta === 0) return
    event.preventDefault()
    onZoomChange(clampMermaidZoom(zoom + (delta < 0 ? MERMAID_ZOOM_STEP : -MERMAID_ZOOM_STEP)))
  }

  if (error) {
    return (
      <div className="mermaid-error-banner">
        <div className="mermaid-error-title">Syntax Error</div>
        <pre className="mermaid-error-body">{error}</pre>
      </div>
    )
  }

  return (
    <div
      className={`mermaid-svg-container is-draggable${isPanning ? ' is-panning' : ''}`}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={handlePointerUp}
      onPointerCancel={handlePointerUp}
      onWheel={handleWheel}
      title="Drag to pan the preview"
    >
      {svg && <div ref={mountRef} className="mermaid-svg-wrapper" />}
      {!svg && <div className="mermaid-empty-hint">Enter Mermaid syntax to see preview</div>}
    </div>
  )
}

function MermaidPreview({ svg, error, isRendering, zoom, onZoomChange }: MermaidPreviewProps) {
  return (
    <div className="mermaid-preview-pane">
      <div className="mermaid-pane-header">
        <div className="mermaid-preview-header-left">
          <span className="mermaid-pane-title">Live Preview</span>
          {isRendering && <span className="mermaid-status-badge">Rendering...</span>}
        </div>
        <MermaidPreviewToolbar svg={svg} />
      </div>
      <MermaidPreviewCanvas svg={svg} error={error} zoom={zoom} onZoomChange={onZoomChange} />
    </div>
  )
}

export function MermaidEditor({ code, onChange, zoom, onZoomChange }: MermaidEditorProps) {
  const [svg, setSvg] = useState('')
  const [error, setError] = useState<string | null>(null)
  const [isRendering, setIsRendering] = useState(false)
  const rootRef = useRef<HTMLDivElement>(null)
  // ponytail: track render sequence to discard out-of-date renders during rapid typing
  const renderSeqRef = useRef(0)

  useEffect(() => {
    const seq = ++renderSeqRef.current
    setIsRendering(true)

    let cancelled = false
    const timer = setTimeout(async () => {
      const result = await renderMermaidDiagram(code)
      if (cancelled || seq !== renderSeqRef.current) return
      setIsRendering(false)
      if (result.error) {
        setError(result.error)
      } else {
        setError(null)
        setSvg(result.svg || '')
      }
    }, 250)

    return () => {
      cancelled = true
      clearTimeout(timer)
    }
  }, [code])

  return (
    <div ref={rootRef} className="mermaid-editor-root">
      <div className="mermaid-code-pane">
        <div className="mermaid-pane-header">
          <span className="mermaid-pane-title">Mermaid Source</span>
        </div>
        <textarea
          className="mermaid-textarea"
          value={code}
          onChange={(e) => onChange(e.target.value)}
          placeholder="Enter Mermaid diagram code (e.g. flowchart TD\n  A --> B)..."
          spellCheck={false}
        />
      </div>
      <MermaidPreview
        svg={svg}
        error={error}
        isRendering={isRendering}
        zoom={zoom}
        onZoomChange={onZoomChange}
      />
    </div>
  )
}
