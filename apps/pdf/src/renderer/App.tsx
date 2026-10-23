import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { CSSProperties, ReactElement, ReactNode, RefObject } from 'react'
// legacy build: the modern build relies on new APIs like Math.sumPrecise that the current
// Electron V8 lacks, making embedded font parsing fail and whole pages render as garbled raw char codes
import { GlobalWorkerOptions, TextLayer, getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import { AraSpark } from '@araoffice/ui'
import { AiPanel } from './ai/AiPanel'
import type { PdfSkillDeps } from './ai/tools'
import {
  MARKUP_COLORS,
  geomDispSize,
  pdfRectToCss,
  quadToRect,
  selectionQuadsByPage,
  viewToPdf,
} from './annotations'
import type { LocalMarkup, PageGeom } from './annotations'
import { DRAW_COLORS, DrawLayer, cssRgb } from './DrawLayer'
import type { DrawTool, LocalDrawing } from './DrawLayer'
import { FormLayer } from './FormLayer'
import { navAction } from './keyNav'
import { rowOfVisIdx, rowTopOffsets, spreadRows, stepPage } from './spread'
import { hasUnsavedEdits } from './dirty'
import { LinkLayer } from './LinkLayer'
import { OutlinePanel } from './OutlinePanel'
import { CommentsPanel } from './CommentsPanel'
import { annotationsToRows, type CommentRow } from './comments'
import type { OutlineEditOp, OutlineNode } from './OutlinePanel'
import { PdfPrintPreview } from './PdfPrintPreview'
import { PropertiesDialog } from './PropertiesDialog'
import { SignatureDialog } from './SignatureDialog'
import type { SignatureData } from './SignatureDialog'
import { StampDialog } from './StampDialog'
import { buildStamps, renderTextBox } from './stamps'
import type { HeaderFooterConfig, WatermarkConfig } from './stamps'
import { buildSearchIndex, indexToText, searchInIndex } from './search'
import { splitRanges } from './split'
import type { SplitMode } from './split'
import type { SearchIndex, SearchMatch } from './search'
import { useI18n, t as tModule } from './i18n/locale'
import { htmlLang } from '@araoffice/i18n'
import { useAutosave } from './useAutosave'
import type {
  DrawingInput,
  FormValueInput,
  MarkupType,
  MetadataInput,
  SanitizeReport,
  SanitizeRequest,
  OutlineEditNode,
  A11yReport,
  A11yFixRequest,
  A11yFixResult,
  DeleteAnnotationsResult,
  SanitizeResult,
  SanitizePreviewResult,
  StampInput,
} from '../shared/ipc'

GlobalWorkerOptions.workerSrc = workerUrl

// cmaps/standard fonts/wasm are statically copied by the build into pdfjs/ of the renderer output (same path on the dev server)
const ASSET_BASE = new URL('pdfjs/', document.baseURI).href

const ZOOM_STEPS = [0.5, 0.67, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4]
const MIN_SCALE = ZOOM_STEPS[0]
const MAX_SCALE = ZOOM_STEPS[ZOOM_STEPS.length - 1]
const PAGE_GAP = 16
const SCROLL_PAD = 24
// ── Sidebar (thumbnails / outline) width: drag the divider to resize; persisted ──
const SIDEBAR_W_KEY = 'araoffice-pdf-sidebar-width'
const SIDEBAR_W_DEFAULT = 150
const SIDEBAR_W_MIN = 120
/** pane padding (10px × 2) + thumb box borders (2px × 2) */
const SIDEBAR_CHROME = 24

const clampSidebarW = (w: number): number =>
  Math.min(Math.max(w, SIDEBAR_W_MIN), Math.min(320, Math.round(window.innerWidth * 0.4)))

const loadSidebarW = (): number => {
  const saved = Number(localStorage.getItem(SIDEBAR_W_KEY))
  return Number.isFinite(saved) && saved > 0 ? clampSidebarW(saved) : SIDEBAR_W_DEFAULT
}

interface PageSize {
  width: number
  height: number
}

type FitMode = 'width' | 'page' | null

const DOC_OPTS = {
  cMapUrl: `${ASSET_BASE}cmaps/`,
  cMapPacked: true,
  standardFontDataUrl: `${ASSET_BASE}standard_fonts/`,
  wasmUrl: `${ASSET_BASE}wasm/`,
}

/** Which items in the container are within the (expanded) viewport — shared lazy-render basis
    for pages/thumbnails. Rebuild the observer when enabled flips (sidebar toggles unmount/remount the root) */
function useVisibleSet(
  rootRef: RefObject<HTMLElement | null>,
  count: number,
  rootMargin: string,
  enabled = true,
): { visible: Set<number>; setItemRef: (idx: number) => (el: HTMLElement | null) => void } {
  const [visible, setVisible] = useState<Set<number>>(new Set())
  const itemRefs = useRef<(HTMLElement | null)[]>([])
  useEffect(() => {
    const root = rootRef.current
    if (!enabled || !root || count === 0) return
    const io = new IntersectionObserver(
      (entries) => {
        setVisible((prev) => {
          const next = new Set(prev)
          for (const e of entries) {
            const idx = Number((e.target as HTMLElement).dataset.idx)
            if (e.isIntersecting) next.add(idx)
            else next.delete(idx)
          }
          return next
        })
      },
      { root, rootMargin },
    )
    for (const el of itemRefs.current) if (el) io.observe(el)
    return () => io.disconnect()
  }, [rootRef, count, rootMargin, enabled])
  return {
    visible,
    setItemRef: (idx) => (el) => {
      itemRefs.current[idx] = el
    },
  }
}

/** Single page: renders canvas + text layer (select/copy) when visible, released once off-viewport */
function PdfPage({
  doc,
  pageNo,
  scale,
  rotationDelta,
  visible,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  scale: number
  /** Unsaved rotation delta (clockwise degrees) */
  rotationDelta: number
  visible: boolean
}) {
  const holderRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const holder = holderRef.current
    if (!visible || !holder) return
    let cancelled = false
    let renderTask: RenderTask | null = null
    let allocatedCanvas: HTMLCanvasElement | null = null
    void (async () => {
      const page = await doc.getPage(pageNo)
      if (cancelled) return
      const viewport = page.getViewport({ scale, rotation: (page.rotate + rotationDelta) % 360 })
      // Cap at 2x: on hi-dpi screens a 3x-dpr full-page bitmap doubles memory with no visible gain
      const dpr = Math.min(window.devicePixelRatio || 1, 2)
      const canvas = document.createElement('canvas')
      allocatedCanvas = canvas
      canvas.width = Math.floor(viewport.width * dpr)
      canvas.height = Math.floor(viewport.height * dpr)
      canvas.style.width = `${Math.floor(viewport.width)}px`
      canvas.style.height = `${Math.floor(viewport.height)}px`
      renderTask = page.render({
        canvas,
        viewport,
        transform: dpr === 1 ? undefined : [dpr, 0, 0, dpr, 0, 0],
      })
      try {
        await renderTask.promise
      } catch {
        return // cancelled
      }
      if (cancelled) return
      const textDiv = document.createElement('div')
      textDiv.className = 'textLayer'
      holder.replaceChildren(canvas, textDiv)
      const textLayer = new TextLayer({
        textContentSource: page.streamTextContent(),
        container: textDiv,
        viewport,
      })
      try {
        await textLayer.render()
      } catch {
        /* cancelled */
      }
    })()
    return () => {
      cancelled = true
      renderTask?.cancel()
      if (allocatedCanvas) {
        allocatedCanvas.width = 0
        allocatedCanvas.height = 0
        allocatedCanvas = null
      }
      holder.replaceChildren()
    }
  }, [doc, pageNo, scale, rotationDelta, visible])
  return <div ref={holderRef} className="pdf-page-content" />
}

/** A wavy underline as a repeating SVG tile, matching the wave the main process writes into the
    squiggly appearance stream */
const squigglyTile = (stroke: string) =>
  `url("data:image/svg+xml,${encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="6" height="4" viewBox="0 0 6 4"><path d="M0 2.5 Q1.5 0.5 3 2.5 Q4.5 4.5 6 2.5" fill="none" stroke="${stroke}" stroke-width="1.2"/></svg>`,
  )}")`

/** Overlay for unsaved markups; click to select (deletion is explicit via the delete popup or Delete key) */
function MarkupOverlay({
  markups,
  geom,
  scale,
  selectedId,
  selectTitle,
  onSelect,
}: {
  markups: LocalMarkup[]
  geom: PageGeom
  scale: number
  selectedId: string | null
  selectTitle: string
  onSelect: (id: string, x: number, y: number) => void
}) {
  return (
    <>
      {markups.flatMap((m) =>
        m.quads.map((q, i) => {
          const [r, g, b] = m.color
          const style: CSSProperties = pdfRectToCss(geom, quadToRect(q), scale)
          if (m.type === 'highlight') {
            style.background = `rgba(${r * 255}, ${g * 255}, ${b * 255}, 0.4)`
          } else {
            const bar = `rgb(${r * 255}, ${g * 255}, ${b * 255})`
            if (m.type === 'squiggly') {
              style.backgroundImage = squigglyTile(bar)
              style.backgroundRepeat = 'repeat-x'
              style.backgroundPosition = 'left bottom'
              style.backgroundSize = '6px 4px'
            } else if (m.type === 'underline') {
              style.borderBottom = `2px solid ${bar}`
            } else {
              style.backgroundImage = `linear-gradient(${bar}, ${bar})`
            }
          }
          return (
            <div
              key={`${m.id}-${i}`}
              className={`pdf-markup pdf-markup-${m.type}${m.id === selectedId ? ' pdf-markup-selected' : ''}`}
              style={style}
              title={selectTitle}
              onClick={(e) => onSelect(m.id, e.clientX, e.clientY)}
            />
          )
        }),
      )}
    </>
  )
}

/** Thumbnail: rendered once per (doc, rotation, raster width) when visible and cached.
 *  rasterW only changes when a sidebar drag ends, so a resize re-rasters each
 *  visible thumb once; while dragging the canvas just CSS-stretches. */
function PdfThumb({
  doc,
  pageNo,
  rotationDelta,
  visible,
  rasterW,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  rotationDelta: number
  visible: boolean
  rasterW: number
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const renderedKeyRef = useRef<string | null>(null)
  useEffect(() => {
    const canvas = canvasRef.current
    const key = `${rotationDelta}:${rasterW}`
    if (!visible || !canvas || renderedKeyRef.current === key) return
    renderedKeyRef.current = key
    let cancelled = false
    void (async () => {
      const page = await doc.getPage(pageNo)
      if (cancelled) return
      const rotation = (page.rotate + rotationDelta) % 360
      const scale = rasterW / page.getViewport({ scale: 1, rotation }).width
      const viewport = page.getViewport({ scale, rotation })
      const dpr = Math.min(window.devicePixelRatio || 1, 2)
      canvas.width = Math.floor(viewport.width * dpr)
      canvas.height = Math.floor(viewport.height * dpr)
      try {
        await page.render({
          canvas,
          viewport,
          transform: dpr === 1 ? undefined : [dpr, 0, 0, dpr, 0, 0],
        }).promise
      } catch {
        renderedKeyRef.current = null
      }
    })()
    return () => {
      cancelled = true
    }
  }, [doc, pageNo, rotationDelta, visible, rasterW])
  // Reset the cache key when the doc changes (save reload); re-render next time it's visible
  useEffect(() => {
    renderedKeyRef.current = null
  }, [doc])
  return <canvas ref={canvasRef} style={{ width: '100%' }} />
}

// ── ribbon icons (aligned with slides' rb-big visual language) ──

/** Constant painted stroke instead of proportional scaling — same rule as the
 *  slides icons: ~1.5px lines at 20px+, ~1.25px on 13-19px glyphs, ~1.1px below.
 *  stroke-width is in 24-canvas units: units = painted-px × 24 / rendered-px. */
function pinnedStroke(size: number): number {
  const painted = size >= 20 ? 1.5 : size >= 13 ? 1.25 : 1.1
  return (painted * 24) / size
}

function Icon({ size = 28, children }: { size?: number; children: ReactNode }): ReactElement {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={pinnedStroke(size)}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

const IconThumbs = () => (
  <Icon>
    <rect x="4.5" y="5" width="6" height="6.5" rx="1" />
    <rect x="4.5" y="14" width="6" height="5" rx="1" />
    <path d="M14 6 L19.5 6 M14 10 L19.5 10 M14 15 L19.5 15 M14 18 L19.5 18" />
  </Icon>
)
const IconHighlight = () => (
  <Icon>
    <path d="M6.04 15.09 L13.54 7.59 L16.63 10.68 L9.13 18.18 L6.04 18.18 L6.04 15.09 Z" />
    <path d="M13.54 6.71 L15.75 4.5 L18.84 7.59 L16.63 9.79" />
    <path d="M5.16 19.5 L17.51 19.5" strokeWidth={2.2} />
  </Icon>
)
const IconUnderline = () => (
  <Icon>
    <path d="M7.56 4.5 L7.56 11.17 A 4.44 4.44 0 0 0 16.44 11.17 L16.44 4.5" />
    <path d="M5.89 19.5 L18.11 19.5" strokeWidth={1.85} />
  </Icon>
)
const IconStrike = () => (
  <Icon>
    <path d="M16.29 7.21 C15.72 5.52 14.03 4.5 12 4.5 C9.52 4.5 7.71 5.85 7.71 7.88 C7.71 9.46 8.73 10.36 10.65 10.93" />
    <path d="M8.62 16.91 C9.18 18.48 10.87 19.5 13.02 19.5 C15.5 19.5 17.41 18.26 17.41 16.12 C17.41 15.44 17.3 14.88 16.96 14.42" />
    <path d="M4.67 12.28 L19.33 12.28" strokeWidth={1.65} />
  </Icon>
)
const IconSquiggly = () => (
  <Icon>
    <path d="M4.5 14 Q7 9.5 9.5 14 Q12 18.5 14.5 14 Q17 9.5 19.5 14" />
  </Icon>
)
const IconInk = () => (
  <Icon>
    <path d="M16.15 4.85 L19.15 7.85 L8.9 18.1 L4.9 19.1 L5.9 15.1 Z" />
    <path d="M14.15 6.85 L17.15 9.85" />
  </Icon>
)
const IconRect = () => (
  <Icon>
    <rect x="4.5" y="6.64" width="15" height="10.71" rx="1.07" />
  </Icon>
)
const IconEllipse = () => (
  <Icon>
    <ellipse cx="12" cy="12" rx="7.5" ry="5.57" />
  </Icon>
)
const IconArrow = () => (
  <Icon>
    <path d="M4.5 19.2 L19.5 4.8" />
    <path d="M12.9 4.8 L19.5 4.8 L19.5 11.4" />
  </Icon>
)
const IconNote = () => (
  <Icon>
    <path d="M4.5 5.57 L19.5 5.57 L19.5 15.21 L10.93 15.21 L6.64 18.43 L6.64 15.21 L4.5 15.21 Z" />
    <path d="M8.25 9.32 L15.75 9.32 M8.25 12 L13.07 12" />
  </Icon>
)
const IconPolyline = () => (
  <Icon>
    <path d="M4.5 17.5 L9.5 9.5 L14 14 L19.5 5.5" />
  </Icon>
)
const IconPolygon = () => (
  <Icon>
    <path d="M12 4.5 L19.5 10 L16.6 18.8 L7.4 18.8 L4.5 10 Z" />
  </Icon>
)
const IconCaret = () => (
  <Icon>
    <path d="M8 18.5 L12 6.5 L16 18.5" />
  </Icon>
)
const IconFreeText = () => (
  <Icon>
    <path d="M6 8.5 V6.5 H18 V8.5" />
    <path d="M12 6.5 V18" />
    <path d="M9.5 18 H14.5" />
  </Icon>
)
const IconSign = () => (
  <Icon>
    <path d="M5.5 15.1 C7.8 12.3 9.5 9 9.2 7 C9 5.7 7.9 5.9 7.6 7.4 C7.2 9.6 8.6 13.4 10.5 14.9 C12 16.1 13.9 15.3 14.7 13.8 C15.1 13 15.9 13 16.3 13.8 C16.7 14.7 17.7 15 18.5 14.4" />
    <path d="M4.75 18.6 L19.25 18.6" />
  </Icon>
)
const IconExportImg = () => (
  <Icon>
    <rect x="4.5" y="6.75" width="15" height="10.5" rx="1" />
    <circle cx="9" cy="10.55" r="1.2" />
    <path d="M4.8 14.95 L9 11.75 L12.4 14.35 L15 12.35 L19.2 15.75" />
  </Icon>
)
const IconExportText = () => (
  <Icon>
    <path d="M6 5.5 H14.5 L18 9 V18.5 H6 Z" />
    <path d="M14 5.5 V9.5 H18" />
    <path d="M8.75 12.5 H15.25 M8.75 15.25 H13" />
  </Icon>
)
const IconSpread = () => (
  <Icon>
    <rect x="4.5" y="6" width="6.5" height="12" rx="1" />
    <rect x="13" y="6" width="6.5" height="12" rx="1" />
  </Icon>
)
const IconSinglePage = () => (
  <Icon>
    <rect x="6.81" y="4.5" width="10.38" height="15" rx="1.15" />
  </Icon>
)
const IconWatermark = () => (
  <Icon>
    <rect x="4.5" y="5.04" width="15" height="13.93" rx="1.07" />
    <path d="M7.71 15.75 L15.75 7.71" />
    <path d="M7.71 11.46 L11.46 7.71 M12.54 15.75 L16.29 12" />
  </Icon>
)
const IconProps = () => (
  <Icon>
    <circle cx="12" cy="12" r="7.5" />
    <path d="M12 10.96 L12 15.65" />
    <circle cx="12" cy="8.46" r="0.94" fill="currentColor" stroke="none" />
  </Icon>
)
const IconA11y = () => (
  <Icon>
    <circle cx="12" cy="5.6" r="2" />
    <path d="M5.5 9.8 C9.8 11 14.2 11 18.5 9.8" />
    <path d="M12 11 L12 14.4" />
    <path d="M12 14.4 L9.4 19" />
    <path d="M12 14.4 L14.6 19" />
  </Icon>
)
const IconRotateL = () => (
  <Icon>
    <path d="M8.28 10.3 L4.53 10.3 L4.53 6.55" />
    <path d="M4.75 9.98 A 7.5 7.5 0 1 1 4.53 12.98" />
  </Icon>
)
const IconRotateR = () => (
  <Icon>
    <path d="M15.72 10.3 L19.47 10.3 L19.47 6.55" />
    <path d="M19.25 9.98 A 7.5 7.5 0 1 0 19.47 12.98" />
  </Icon>
)
const IconDeletePage = () => (
  <Icon>
    <path d="M7.7 4.5 H13.7 L17.2 8 V18.5 A1 1 0 0 1 16.2 19.5 H7.7 A1 1 0 0 1 6.7 18.5 V5.5 A1 1 0 0 1 7.7 4.5 Z" />
    <path d="M13.7 4.5 V8 H17.2" />
    <path d="M9.7 11.75 L14.2 16.25 M14.2 11.75 L9.7 16.25" />
  </Icon>
)
const IconExtract = () => (
  <Icon>
    <path d="M7.2 4.5 H13.2 L16.7 8 V11.5" />
    <path d="M6.2 5.5 V18.5 A1 1 0 0 0 7.2 19.5 H11.2" />
    <path d="M14.95 13.5 V19 M12.2 16.5 L14.95 19.25 L17.7 16.5" />
  </Icon>
)
const IconSplit = () => (
  <Icon>
    <path d="M7.5 4.5 H13.5 L17 8 V18.5 A1 1 0 0 1 16 19.5 H7.5 A1 1 0 0 1 6.5 18.5 V5.5 A1 1 0 0 1 7.5 4.5 Z" />
    <path d="M13.5 4.5 V8 H17" />
    <path d="M6.5 12.25 H17" strokeDasharray="2.4 2" />
  </Icon>
)
const IconCombine = () => (
  <Icon>
    <path d="M8 4.5 H13.5 L16.5 7.5 V13" />
    <path d="M11 9.5 H16.5 L19.5 12.5 V18 A1 1 0 0 1 18.5 19 H11.5 A1 1 0 0 1 10.5 18 V10.5 A1 1 0 0 1 11 9.5 Z" />
    <path d="M16.5 9.5 V12.5 H19.5" />
  </Icon>
)
const IconSanitize = () => (
  <Icon>
    <path d="M12 3.5 L18.5 6 V12 C18.5 15.5 16 18 12 19.5 C8 18 5.5 15.5 5.5 12 V6 Z" />
    <path d="M9.5 11.5 L11.5 13.5 L14.75 10" />
  </Icon>
)
const IconImgStack = () => (
  <Icon>
    <path d="M7.5 8 H16.5 M10 4.5 H19" />
    <rect x="4.5" y="11" width="13" height="8.5" rx="1" />
  </Icon>
)
const IconInsertPdf = () => (
  <Icon>
    <path d="M7.7 4.5 H13.7 L17.2 8 V18.5 A1 1 0 0 1 16.2 19.5 H7.7 A1 1 0 0 1 6.7 18.5 V5.5 A1 1 0 0 1 7.7 4.5 Z" />
    <path d="M13.7 4.5 V8 H17.2" />
    <path d="M11.95 11 V17 M8.95 14 H14.95" />
  </Icon>
)
const IconFitWidth = () => (
  <Icon>
    <path d="M4.5 5.57 L4.5 18.43 M19.5 5.57 L19.5 18.43" />
    <path d="M7.71 12 L16.29 12 M10.07 9.64 L7.71 12 L10.07 14.36 M13.93 9.64 L16.29 12 L13.93 14.36" />
  </Icon>
)
const IconFitPage = () => (
  <Icon>
    <rect x="7" y="4.5" width="10" height="15" rx="1" />
    <path d="M12 8 L12 16 M9.8 10.2 L12 8 L14.2 10.2 M9.8 13.8 L12 16 L14.2 13.8" />
  </Icon>
)
const IconOutline = () => (
  <Icon>
    <path d="M4.84 4.78 L19.5 4.78 M8.22 9.29 L19.5 9.29 M8.22 13.8 L19.5 13.8 M11.61 18.32 L19.5 18.32" />
    <circle cx="5.4" cy="9.29" r="0.9" fill="currentColor" stroke="none" />
    <circle cx="5.4" cy="13.8" r="0.9" fill="currentColor" stroke="none" />
    <circle cx="8.79" cy="18.32" r="0.9" fill="currentColor" stroke="none" />
  </Icon>
)

const IconComments = () => (
  <Icon>
    <path d="M20.6 11.6 a 7.4 7.4 0 0 1 -7.4 7.2 8 8 0 0 1 -3.5 -0.7 L5.2 19.3 l 1.1 -3.8 A 7.2 7.2 0 0 1 4.6 11.6 A 7.4 7.4 0 0 1 12 4.4 a 7.4 7.4 0 0 1 8.6 7.2 Z" />
    <path d="M8.6 11.6 L15.4 11.6" />
  </Icon>
)

const IconDrawColor = () => (
  <Icon>
    <path d="M12 4.5 C14.2 7.3 17.25 9.2 17.25 12.4 C17.25 15.4 14.9 17.5 12 17.5 C9.1 17.5 6.75 15.4 6.75 12.4 C6.75 9.2 9.8 7.3 12 4.5 Z" />
  </Icon>
)
/* dropdown chevron, same glyph as slides' RbCaret */
const RbCaret = () => (
  <svg className="rb-caret" width="10" height="10" viewBox="0 0 24 24" fill="none" aria-hidden>
    <path
      d="M5.5 9.25 12 15.75l6.5-6.5"
      stroke="currentColor"
      strokeWidth="2.6"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
)
const IconSearch = () => (
  <Icon>
    <circle cx="10.61" cy="10.61" r="6.11" />
    <path d="M15.28 15.28 L19.5 19.5" />
  </Icon>
)
const IconPrint = () => (
  <Icon>
    <path d="M7.71 8.79 L7.71 4.5 L16.29 4.5 L16.29 8.79" />
    <rect x="5.04" y="8.79" width="13.93" height="6.96" rx="1.07" />
    <path d="M7.71 13.07 L16.29 13.07 L16.29 19.5 L7.71 19.5 Z" />
  </Icon>
)
const IconZoomIn = () => (
  <Icon>
    <circle cx="11" cy="11" r="6" />
    <path d="M11 8 V14 M8 11 H14 M15.5 15.5 L19.5 19.5" />
  </Icon>
)
const IconZoomOut = () => (
  <Icon>
    <circle cx="11" cy="11" r="6" />
    <path d="M8 11 H14 M15.5 15.5 L19.5 19.5" />
  </Icon>
)
const IconActualSize = () => (
  <Icon>
    <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
    <path d="M9 10 L10.5 8.5 V15.5 M13 10.5 A1.5 1.5 0 0 1 14.5 9 A1.5 1.5 0 0 1 16 10.5 V14 A1.5 1.5 0 0 1 14.5 15.5 A1.5 1.5 0 0 1 13 14 Z" />
  </Icon>
)
const IconFullscreen = () => (
  <Icon>
    <path d="M4.5 9 V4.5 H9 M15 4.5 H19.5 V9 M19.5 15 V19.5 H15 M9 19.5 H4.5 V15" />
  </Icon>
)
const AssistantMark = ({ size = 26 }: { size?: number }) => (
  <AraSpark size={size} className="ara-assistant-badge" />
)

const IconSummarize = () => (
  <Icon>
    <path d="M4 6 H20 M4 11 H15 M4 16 H11 M4 21 H9" />
    <path d="M16 15 L18 17 L22 13" />
  </Icon>
)

const IconKeyPoints = () => (
  <Icon>
    <circle cx="6" cy="7" r="1.5" fill="currentColor" />
    <circle cx="6" cy="12" r="1.5" fill="currentColor" />
    <circle cx="6" cy="17" r="1.5" fill="currentColor" />
    <path d="M11 7 H20 M11 12 H20 M11 17 H20" />
  </Icon>
)

const IconExplain = () => (
  <Icon>
    <circle cx="12" cy="12" r="9" />
    <path d="M12 17 V16 M12 13 C12 11 14 11 14 9.5 A2 2 0 1 0 10.5 8" />
  </Icon>
)

const IconTranslate = () => (
  <Icon>
    <path d="M5 8 H13 M9 5 V8 M6 8 C6.5 11 8.5 13.5 12 15 M8 11 C9 12 10.5 13 12 13.5" />
    <path d="M14 19 L17.5 10 L21 19 M15.2 16.5 H19.8" />
  </Icon>
)

/** Design-supplied glyphs on the 1:16 stroke:canvas ratio (24-canvas / 1.5 stroke),
 *  geometry shared with docs/slides/sheets. Size 16 → pinnedStroke(16) = 1.5,
 *  so the shared Icon wrapper renders identical output. */
const IconUndo = () => (
  <Icon size={16}>
    <path d="M5.91026 4L2.5 7.14791L5.91026 10.8205" />
    <path d="M3.96154 7.41028H15.1636C18.5169 7.41028 21.3646 10.1484 21.4953 13.5C21.6334 17.0416 18.707 20.0769 15.1636 20.0769H6.88384" />
  </Icon>
)
const IconRedo = () => (
  <Icon size={16}>
    <path d="M18.0897 4L21.5 7.14791L18.0897 10.8205" />
    <path d="M20.0385 7.41028H8.83636C5.4831 7.41028 2.63537 10.1484 2.5047 13.5C2.36657 17.0416 5.29296 20.0769 8.83636 20.0769H17.1162" />
  </Icon>
)
const IconSave = () => (
  <Icon size={16}>
    <path d="M3 4.5C3 3.67158 3.67158 3 4.5 3H17.1407L21 6.60325V19.5C21 20.3285 20.3285 21 19.5 21H4.5C3.67158 21 3 20.3285 3 19.5V4.5Z" />
    <path d="M12.0042 3L12 6.6923C12 6.86225 11.7761 7 11.5 7H7.5C7.22385 7 7 6.86225 7 6.6923V3H12.0042Z" />
    <path d="M7 13H17" />
    <path d="M7 17H12.0042" />
  </Icon>
)

const rgbToHex = (c: readonly [number, number, number]): string =>
  `#${c
    .map((v) =>
      Math.round(v * 255)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`
const hexToRgb = (hex: string): [number, number, number] => [
  parseInt(hex.slice(1, 3), 16) / 255,
  parseInt(hex.slice(3, 5), 16) / 255,
  parseInt(hex.slice(5, 7), 16) / 255,
]

interface ThumbMenu {
  x: number
  y: number
  origIdx: number
}

const DRAW_TOOLS = [
  { tool: 'ink' as const, icon: IconInk, key: 'drawInk' as const },
  { tool: 'rect' as const, icon: IconRect, key: 'drawRect' as const },
  { tool: 'ellipse' as const, icon: IconEllipse, key: 'drawEllipse' as const },
  { tool: 'arrow' as const, icon: IconArrow, key: 'drawArrow' as const },
  { tool: 'polyline' as const, icon: IconPolyline, key: 'drawPolyline' as const },
  { tool: 'polygon' as const, icon: IconPolygon, key: 'drawPolygon' as const },
  { tool: 'note' as const, icon: IconNote, key: 'drawNote' as const },
  { tool: 'caret' as const, icon: IconCaret, key: 'drawCaret' as const },
  { tool: 'freetext' as const, icon: IconFreeText, key: 'drawFreeText' as const },
]

/** Line height a placed caret spans, and the point size a typewriter box is typed at (PDF pt) */
const CARET_SIZE = 12
const FREETEXT_SIZE = 11

/** Sanitize categories in their display order: a key into SanitizeRequest and its label together,
    so neither can drift from the other */
const SANITIZE_ROWS = [
  { key: 'metadata', label: 'sanitizeMeta' },
  { key: 'javascript', label: 'sanitizeJs' },
  { key: 'attachments', label: 'sanitizeAttachments' },
  { key: 'comments', label: 'sanitizeComments' },
  { key: 'thumbnails', label: 'sanitizeThumbs' },
  { key: 'privateData', label: 'sanitizePrivate' },
] as const

/** The report's rule set in display order; verdicts get glyphs, labels get the strings */
const A11Y_ROWS = [
  { rule: 'title', label: 'a11yRuleTitle' },
  { rule: 'language', label: 'a11yRuleLanguage' },
  { rule: 'tagged', label: 'a11yRuleTagged' },
  { rule: 'altFigures', label: 'a11yRuleAlt' },
  { rule: 'bookmarksLong', label: 'a11yRuleBookmarks' },
  { rule: 'headings', label: 'a11yRuleHeadings' },
  { rule: 'tables', label: 'a11yRuleTables' },
  { rule: 'forms', label: 'a11yRuleForms' },
] as const

/** The two rules the one-click fix covers; the dialog offers it only while one still fails */
const a11yMetaFail = (rules: A11yReport['rules']): boolean =>
  rules.some((r) => (r.rule === 'title' || r.rule === 'language') && r.verdict === 'fail')

/** Drawing stroke width (PDF pt); thin lines stay crisp under zoom */
const STROKE_WIDTH = 2

/** Full snapshot of unsaved edits (for undo/redo; data is small, whole-copy replace is safest) */
/** Watermark/header-footer are kept as config and rendered in final page order only at save time, so page numbers survive reorders/deletions */
interface StampConfig {
  wm: WatermarkConfig | null
  hf: HeaderFooterConfig | null
}

interface EditSnapshot {
  markups: LocalMarkup[]
  drawings: LocalDrawing[]
  stampCfg: StampConfig | null
  formEdits: Map<string, FormValueInput>
  rotations: Map<number, number>
  deleted: Set<number>
  order: number[] | null
  metadata: MetadataInput | null
}

/** Selected annotation with the anchor of its floating delete popup; a stamp click selects the whole watermark/header-footer set */
type AnnotSelection =
  | { kind: 'markup' | 'drawing'; id: string; x: number; y: number }
  | { kind: 'stamp'; x: number; y: number }

/** Page ranges like "1-3,5" → list of 1-based page numbers; null if invalid */
function parsePageRanges(input: string, max: number): number[] | null {
  const out = new Set<number>()
  for (const part of input.split(/[,，]/)) {
    const s = part.trim()
    if (!s) continue
    const m = /^(\d+)\s*[-–]\s*(\d+)$|^(\d+)$/.exec(s)
    if (!m) return null
    const a = Number(m[1] ?? m[3])
    const b = Number(m[2] ?? m[3])
    if (a < 1 || b > max || a > b) return null
    for (let i = a; i <= b; i++) out.add(i)
  }
  return out.size > 0 ? [...out].sort((x, y) => x - y) : null
}

/**
 * Bookmark-editor tree helpers. All are pure: the App holds the draft as state and applies them
 * through setOutlineDraft, so every op is one immutable replacement.
 */

/** Replace the sibling list that owns `path`'s parent with fn's result */
const mapOutlinePath = (
  nodes: OutlineEditNode[],
  path: number[],
  fn: (siblings: OutlineEditNode[]) => OutlineEditNode[],
): OutlineEditNode[] => {
  if (path.length === 0) return fn(nodes)
  const [head, ...rest] = path
  return nodes.map((node, i) =>
    i === head ? { ...node, children: mapOutlinePath(node.children ?? [], rest, fn) } : node,
  )
}

/** The node at `path`, or null when the path does not exist in the tree */
const outlineNodeAt = (nodes: OutlineEditNode[], path: number[]): OutlineEditNode | null => {
  let level = nodes
  let node: OutlineEditNode | null = null
  for (const i of path) {
    node = level[i] ?? null
    if (!node) return null
    level = node.children ?? []
  }
  return node
}

const removeOutlineNode = (nodes: OutlineEditNode[], path: number[]): OutlineEditNode[] =>
  mapOutlinePath(nodes, path.slice(0, -1), (siblings) =>
    siblings.filter((_, i) => i !== path[path.length - 1]),
  )

/** Swap a node with the neighbour at delta −1 or +1; no-op at an edge */
const moveOutlineNode = (
  nodes: OutlineEditNode[],
  path: number[],
  delta: -1 | 1,
): OutlineEditNode[] => {
  const i = path[path.length - 1]!
  return mapOutlinePath(nodes, path.slice(0, -1), (siblings) => {
    const j = i + delta
    if (j < 0 || j >= siblings.length) return siblings
    const copy = [...siblings]
    ;[copy[i]!, copy[j]!] = [copy[j]!, copy[i]!]
    return copy
  })
}

const renameOutlineNode = (
  nodes: OutlineEditNode[],
  path: number[],
  title: string,
): OutlineEditNode[] =>
  mapOutlinePath(nodes, path.slice(0, -1), (siblings) =>
    siblings.map((node, i) => (i === path[path.length - 1] ? { ...node, title } : node)),
  )

/**
 * Rebase resolved pages onto the display order the file is being applied with.
 *
 * The draft's pages are 0-based positions in the display order at resolve time; pending page
 * deletes or reorders shift them, so the rebase walks through the original index they resolved to
 * and finds its current position. A node whose page is gone is dropped with its subtree, the same
 * guard the writer applies. URI bookmarks skip the lookup and pass through untouched.
 */
const rebaseOutline = (nodes: OutlineEditNode[], from: number[], to: number[]): OutlineEditNode[] =>
  nodes
    .map((node): OutlineEditNode | null => {
      if (node.uri) return { ...node, children: rebaseOutline(node.children ?? [], from, to) }
      const currentIndex = from[node.pageIndex]
      const idx = currentIndex === undefined ? -1 : to.indexOf(currentIndex)
      if (idx < 0) return null
      return { ...node, pageIndex: idx, children: rebaseOutline(node.children ?? [], from, to) }
    })
    .filter((n): n is OutlineEditNode => n !== null)

/** Scale factor applied when placing a signature: 1/3 of the displayed page width,
    capped at 1/6 of its height so tall images stay signature-sized */
const signPlaceK = (sig: SignatureData, dispW: number, dispH: number): number =>
  Math.min(dispW / 3 / sig.width, dispH / 6 / sig.height)

/** Click-to-place overlay: a translucent ghost of the pending signature follows the
    cursor at its actual landing size, and clicking drops it centered on that point */
function SignDropOverlay({
  sig,
  dispW,
  dispH,
  scale,
  color,
  title,
  onPlace,
}: {
  sig: SignatureData
  /** Displayed page size at scale=1 (view coords) */
  dispW: number
  dispH: number
  scale: number
  color: [number, number, number]
  title: string
  onPlace: (vx: number, vy: number) => void
}): ReactElement {
  const [pt, setPt] = useState<[number, number] | null>(null)
  const k = signPlaceK(sig, dispW, dispH) * scale
  const w = sig.width * k
  const h = sig.height * k
  return (
    <div
      className="pdf-sign-drop"
      title={title}
      onPointerMove={(e) => {
        const box = e.currentTarget.getBoundingClientRect()
        setPt([e.clientX - box.left, e.clientY - box.top])
      }}
      onPointerLeave={() => setPt(null)}
      onClick={(e) => {
        const box = e.currentTarget.getBoundingClientRect()
        onPlace((e.clientX - box.left) / scale, (e.clientY - box.top) / scale)
      }}
    >
      {pt && (
        <div
          className="pdf-sign-ghost"
          style={{
            left: Math.min(Math.max(pt[0] - w / 2, 0), Math.max(dispW * scale - w, 0)),
            top: Math.min(Math.max(pt[1] - h / 2, 0), Math.max(dispH * scale - h, 0)),
            width: w,
            height: h,
          }}
        >
          {sig.kind === 'image' ? (
            <img src={`data:image/png;base64,${sig.image}`} alt="" draggable={false} />
          ) : (
            <svg viewBox={`0 0 ${sig.width} ${sig.height}`} preserveAspectRatio="none">
              {sig.paths.map((p, i) => {
                const pts: string[] = []
                for (let j = 0; j < p.length; j += 2) pts.push(`${p[j]},${p[j + 1]}`)
                return (
                  <polyline
                    key={i}
                    points={pts.join(' ')}
                    fill="none"
                    stroke={cssRgb(color)}
                    strokeWidth={2.4}
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                )
              })}
            </svg>
          )}
        </div>
      )}
    </div>
  )
}

export default function App() {
  const { t, lang: uiLang } = useI18n()
  const [doc, setDoc] = useState<PDFDocumentProxy | null>(null)
  const [filePath, setFilePath] = useState('')
  const [aiPanelKey, setAiPanelKey] = useState(0)
  const [status, setStatus] = useState<'loading' | 'error' | 'empty' | 'password' | 'ready'>(
    'loading',
  )
  const [sizes, setSizes] = useState<PageSize[]>([])
  const [baseRots, setBaseRots] = useState<number[]>([])
  const [scale, setScale] = useState(1)
  const [currentPage, setCurrentPage] = useState(1)
  const [sidebar, setSidebar] = useState<'thumbs' | 'outline' | 'comments' | null>('thumbs')
  const [sidebarW, setSidebarW] = useState(loadSidebarW)
  /** raster width for thumbnails — only updated when a drag ends (re-rastering every frame would jank) */
  const [thumbRasterW, setThumbRasterW] = useState(() => loadSidebarW() - SIDEBAR_CHROME)
  // Re-clamp when the window shrinks (max is 40% of the window), same as slides
  useEffect(() => {
    const onResize = () => setSidebarW((w) => clampSidebarW(w))
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])
  /** Drag to resize: width follows the pointer (rAF-throttled); persisted on release */
  const startSidebarResize = (e: React.PointerEvent<HTMLDivElement>) => {
    e.preventDefault()
    const startX = e.clientX
    const startW = sidebarW
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    let w = startW
    let raf = 0
    const onMove = (ev: PointerEvent) => {
      w = clampSidebarW(startW + ev.clientX - startX)
      if (!raf)
        raf = requestAnimationFrame(() => {
          raf = 0
          setSidebarW(w)
        })
    }
    const onUp = () => {
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', onUp)
      if (raf) cancelAnimationFrame(raf)
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
      setSidebarW(w)
      setThumbRasterW(w - SIDEBAR_CHROME)
      localStorage.setItem(SIDEBAR_W_KEY, String(Math.round(w)))
    }
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', onUp)
  }
  const [aiCollapsed, setAiCollapsed] = useState(() => localStorage.getItem('ai-pdf-open') === '0')
  const [aiPreset, setAiPreset] = useState<{ text: string; nonce: number } | null>(null)
  const triggerAiPreset = (promptText: string) => {
    setAiCollapsed(false)
    setAiPreset({ text: promptText, nonce: Date.now() })
  }

  useEffect(() => {
    localStorage.setItem('ai-pdf-open', aiCollapsed ? '0' : '1')
  }, [aiCollapsed])
  const [spread, setSpread] = useState<1 | 2>(1)
  const [outline, setOutline] = useState<OutlineNode[] | null>(null)
  const [markups, setMarkups] = useState<LocalMarkup[]>([])
  const [drawings, setDrawings] = useState<LocalDrawing[]>([])
  const [drawTool, setDrawTool] = useState<DrawTool | null>(null)
  const [drawColor, setDrawColor] = useState<[number, number, number]>(DRAW_COLORS[0]!.rgb)
  const [colorOpen, setColorOpen] = useState(false)
  const [notePrompt, setNotePrompt] = useState<{
    kind: 'note' | 'caret' | 'freetext'
    origIdx: number
    at: [number, number]
  } | null>(null)
  const [noteText, setNoteText] = useState('')
  const [stampCfg, setStampCfg] = useState<StampConfig | null>(null)
  /** User-defined page order (original page indices); null means unreordered */
  const [order, setOrder] = useState<number[] | null>(null)
  const [metadata, setMetadata] = useState<MetadataInput | null>(null)
  const [stampDlg, setStampDlg] = useState(false)
  const [propsDlg, setPropsDlg] = useState(false)
  const [fileSize, setFileSize] = useState(0)
  const [dragFrom, setDragFrom] = useState<number | null>(null)
  const [dragOver, setDragOver] = useState<number | null>(null)
  const [signDlg, setSignDlg] = useState(false)
  /** Confirmed signature awaiting placement; when non-null the page enters click-to-place mode */
  const [pendingSign, setPendingSign] = useState<SignatureData | null>(null)
  const [exporting, setExporting] = useState(false)
  const [showPrintPreview, setShowPrintPreview] = useState(false)
  const [formEdits, setFormEdits] = useState<Map<string, FormValueInput>>(new Map())
  const [rotations, setRotations] = useState<Map<number, number>>(new Map())
  const [deleted, setDeleted] = useState<Set<number>>(new Set())
  const [selPopup, setSelPopup] = useState<{ x: number; y: number } | null>(null)
  const [selected, setSelected] = useState<AnnotSelection | null>(null)
  const [deleteToast, setDeleteToast] = useState(false)
  const toastTimerRef = useRef<number | null>(null)
  const [thumbMenu, setThumbMenu] = useState<ThumbMenu | null>(null)
  const [saveState, setSaveState] = useState<'idle' | 'saving' | 'saved' | 'error'>('idle')
  const [saveError, setSaveError] = useState('')
  /** Autosave gate: this file was saved explicitly at least once */
  /** Path waiting behind the unsaved-changes prompt (Save/Discard/Cancel) */
  const [pendingOpen, setPendingOpen] = useState<string | null>(null)
  /** dirty in ref form: openPath runs outside the render cycle (IPC mount, menu re-open) */
  const dirtyRef = useRef(false)
  const savedOnceRef = useRef(false)
  const [searchOpen, setSearchOpen] = useState(false)
  const [searchQuery, setSearchQuery] = useState('')
  const [searchMatches, setSearchMatches] = useState<SearchMatch[]>([])
  const [searchCur, setSearchCur] = useState(0)
  const [undoStack, setUndoStack] = useState<EditSnapshot[]>([])
  const [redoStack, setRedoStack] = useState<EditSnapshot[]>([])
  const [pwInput, setPwInput] = useState('')
  const [pwWrong, setPwWrong] = useState(false)
  type PdfRibbonTab = 'home' | 'comment' | 'organize' | 'view' | 'ai'
  const [ribbonTab, setRibbonTab] = useState<PdfRibbonTab>('home')
  const [isFullscreen, setIsFullscreen] = useState(() => Boolean(document.fullscreenElement))
  useEffect(() => {
    const onFsChange = () => setIsFullscreen(Boolean(document.fullscreenElement))
    document.addEventListener('fullscreenchange', onFsChange)
    return () => document.removeEventListener('fullscreenchange', onFsChange)
  }, [])
  const toggleFullscreen = () => {
    if (document.fullscreenElement) void document.exitFullscreen()
    else void document.documentElement.requestFullscreen()
  }
  const [extractDlg, setExtractDlg] = useState(false)
  const [extractInput, setExtractInput] = useState('')
  const [extractInvalid, setExtractInvalid] = useState(false)
  const [splitDlg, setSplitDlg] = useState(false)
  const [splitMode, setSplitMode] = useState<'every' | 'at' | 'bookmark'>('every')
  const [splitEvery, setSplitEvery] = useState('2')
  const [splitPages, setSplitPages] = useState('')
  const [splitInvalid, setSplitInvalid] = useState(false)
  const [sanitizeDlg, setSanitizeDlg] = useState(false)
  const [sanitizeCfg, setSanitizeCfg] = useState<SanitizeRequest>({
    metadata: true,
    javascript: true,
    attachments: true,
    comments: true,
    thumbnails: true,
    privateData: true,
  })
  /** Non-null once a sanitize run has applied: the dialog then shows its report instead of the checkboxes */
  const [sanitizeReport, setSanitizeReport] = useState<SanitizeReport | null>(null)
  /** Per-category census of the current file, fetched when the dialog opens, for the row badges */
  const [sanitizePreview, setSanitizePreview] = useState<SanitizeReport | null>(null)
  const [a11yDlg, setA11yDlg] = useState(false)
  const [a11yReport, setA11yReport] = useState<A11yReport | null>(null)
  /**
   * Display page numbers (1-based) of the top-level bookmarks, used as split boundaries.
   *
   * Destinations reference original page indices, which a deletion shifts, so each is mapped
   * through visList to a position in the document as it stands — the same numbering the split
   * ranges use, because they index the saved file.
   */
  const outlineBoundaryPages = async (): Promise<number[]> => {
    if (!doc || !outline) return []
    const pages: number[] = []
    for (const node of outline) {
      if (node.dest == null) continue
      try {
        const arr = typeof node.dest === 'string' ? await doc.getDestination(node.dest) : node.dest
        if (!Array.isArray(arr)) continue
        const ref = arr[0]
        const origIdx =
          typeof ref === 'number'
            ? ref
            : await doc.getPageIndex(ref as Parameters<PDFDocumentProxy['getPageIndex']>[0])
        const visIdx = visList.indexOf(origIdx)
        if (visIdx >= 0) pages.push(visIdx + 1)
      } catch {
        /* Ignore corrupted destinations, the same way goToDest does */
      }
    }
    return [...new Set(pages)].sort((a, b) => a - b)
  }

  /**
   * Split the saved file into separate PDFs. Ranges index the file as saved, which is the current
   * display order (flushThen saves pending deletions and reordering first), so they are built from
   * pageCount rather than original indices.
   */
  const doSplit = () =>
    flushThen(async () => {
      if (!doc || exporting) return
      let mode: SplitMode
      if (splitMode === 'every') {
        const n = Number(splitEvery)
        if (!Number.isInteger(n) || n < 1) {
          setSplitInvalid(true)
          return
        }
        mode = { kind: 'every', size: n }
      } else if (splitMode === 'at') {
        const pages = parsePageRanges(splitPages, pageCount)
        if (!pages) {
          setSplitInvalid(true)
          return
        }
        mode = { kind: 'at', pages }
      } else {
        const pages = await outlineBoundaryPages()
        if (pages.length === 0) {
          setSplitInvalid(true)
          return
        }
        mode = { kind: 'at', pages }
      }
      const ranges = splitRanges(pageCount, mode)
      if (ranges.length === 0) return
      setSplitDlg(false)
      setSplitInvalid(false)
      setExporting(true)
      try {
        const result = await window.pdfApi.splitPdf({
          path: filePath,
          ranges,
          baseName: fileName.replace(/\.pdf$/i, ''),
        })
        if (!result.ok) opFailed(result.error)
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    })

  /** Merge picked files; the main process shows both the file picker and the save dialog */
  const doCombine = () =>
    flushThen(async () => {
      if (exporting) return
      setExporting(true)
      try {
        const result = await window.pdfApi.combinePdfs({ bookmarkPerFile: true })
        if (!result.ok) opFailed(result.error)
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    })

  /** Sanitize the saved file in place, then reload the open document from it */
  const previewSanitize = async () => {
    if (!doc) return
    try {
      const result: SanitizePreviewResult = await window.pdfApi.sanitizePreview({ path: filePath })
      setSanitizePreview(result.ok ? result.counts : null)
    } catch {
      setSanitizePreview(null)
    }
  }

  const doSanitize = () =>
    flushThen(async () => {
      if (!doc || exporting) return
      setExporting(true)
      try {
        const result: SanitizeResult = await window.pdfApi.sanitizePdf({
          path: filePath,
          ...sanitizeCfg,
        })
        if (!result.ok) {
          opFailed(result.error)
          return
        }
        // The file on disk no longer matches the open document, so reload it — a stale doc would
        // otherwise write the removed content straight back on the next save
        setSanitizeReport(result.report)
        await loadDoc(filePath, doc)
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    })

  const coalesceKeyRef = useRef<string | null>(null)
  const passwordRef = useRef<string | undefined>(undefined)
  const fitModeRef = useRef<FitMode>('width')
  const scrollRef = useRef<HTMLDivElement>(null)
  const thumbsRef = useRef<HTMLDivElement>(null)
  const searchInputRef = useRef<HTMLInputElement>(null)
  const searchIndexRef = useRef<{ doc: PDFDocumentProxy; promise: Promise<SearchIndex> } | null>(
    null,
  )
  const searchJumpRef = useRef<{ matches: SearchMatch[]; cur: number } | null>(null)

  /** Visible pages (with unsaved reorder, deleted pages hidden): position → original page index */
  const visList = useMemo(() => {
    const base = order ?? sizes.map((_, i) => i)
    return base.filter((i) => !deleted.has(i))
  }, [sizes, deleted, order])
  const pageCount = visList.length

  const rows = useMemo(() => spreadRows(visList, spread), [visList, spread])

  /** Visible position → row index */
  const rowOfVis = useCallback((visIdx: number) => rowOfVisIdx(visIdx, spread), [spread])
  const fileName = filePath.split(/[\\/]/).pop() ?? filePath

  const rotDelta = useCallback((origIdx: number) => rotations.get(origIdx) ?? 0, [rotations])
  /** Page geometry: unrotated size + total display rotation; the single entry point for overlay coord conversion */
  const pageGeom = useCallback(
    (origIdx: number): PageGeom => {
      const s = sizes[origIdx]!
      return { pw: s.width, ph: s.height, rot: (baseRots[origIdx] ?? 0) + rotDelta(origIdx) }
    },
    [sizes, baseRots, rotDelta],
  )
  /** Page display size (width/height swapped under rotation) */
  const dispSize = useCallback(
    (origIdx: number): PageSize => geomDispSize(pageGeom(origIdx)),
    [pageGeom],
  )

  const { visible: visibleRows, setItemRef: setRowRef } = useVisibleSet(
    scrollRef,
    rows.length,
    '800px 0px',
  )
  const { visible: visibleThumbs, setItemRef: setThumbRef } = useVisibleSet(
    thumbsRef,
    pageCount,
    '400px 0px',
    sidebar === 'thumbs',
  )

  const loadDoc = useCallback(async (path: string, previous: PDFDocumentProxy | null) => {
    const data = await window.pdfApi.readFile(path)
    const loaded = await getDocument({
      data: new Uint8Array(data),
      password: passwordRef.current,
      ...DOC_OPTS,
    }).promise
    const all: PageSize[] = []
    const rots: number[] = []
    for (let i = 1; i <= loaded.numPages; i++) {
      const page = await loaded.getPage(i)
      // Unrotated size; display size is derived by geom from the total rotation
      const vp = page.getViewport({ scale: 1, rotation: 0 })
      all.push({ width: vp.width, height: vp.height })
      rots.push(page.rotate ?? 0)
    }
    setSizes(all)
    setBaseRots(rots)
    setDoc(loaded)
    setMarkups([])
    setDrawings([])
    setStampCfg(null)
    setFormEdits(new Map())
    setRotations(new Map())
    setDeleted(new Set())
    setOrder(null)
    setMetadata(null)
    setFileSize(data.byteLength)
    setSelected(null)
    setDeleteToast(false)
    setUndoStack([])
    setRedoStack([])
    void loaded.getOutline().then(
      (o) => setOutline(o && o.length > 0 ? (o as OutlineNode[]) : null),
      () => setOutline(null),
    )
    // pdfjs-dist 6.x removed PDFDocumentProxy.destroy(); go through the loading task
    if (previous) void previous.loadingTask.destroy()
  }, [])

  const doOpenPath = useCallback(
    async (path: string) => {
      try {
        setAiPanelKey((key) => key + 1)
        setFilePath(path)
        // A newly opened file starts outside the autosave gate
        savedOnceRef.current = false
        await loadDoc(path, null)
        setStatus('ready')
      } catch (err) {
        if ((err as Error | null)?.name === 'PasswordException') {
          setPwWrong(passwordRef.current !== undefined)
          setStatus('password')
          return
        }
        console.error('[pdf] open failed:', err)
        setStatus('error')
      }
    },
    [loadDoc],
  )

  /** Replacing the document loses pre-first-save edits with no recovery copy — prompt first */
  const openPath = useCallback(
    async (path: string) => {
      if (dirtyRef.current) {
        setPendingOpen(path)
        return
      }
      await doOpenPath(path)
    },
    [doOpenPath],
  )

  useEffect(() => {
    void (async () => {
      const path = await window.pdfApi.consumePending()
      if (!path) {
        setStatus('empty')
        return
      }
      await openPath(path)
    })()
  }, [openPath])

  // Unsaved-changes prompt before openPath replaces the document. Save persists first and
  // falls through to the open on success (failed save keeps the doc + prompt). Discard drops
  // the pending edits; Cancel keeps the current document. saveState 'saving' → doOpenPath
  // (any further save call resolves immediately against the fresh doc) — autosave already
  // can't overlap it (blur/interval are paused while the modal blocks the window).
  const confirmOpen = async (choice: 'save' | 'discard'): Promise<void> => {
    const path = pendingOpen
    setPendingOpen(null)
    if (!path) return
    if (choice === 'save' && !(await save())) return
    await doOpenPath(path)
  }

  /** Documents opened with a password are treated as read-only: pdf-lib can't write back encrypted files */
  const readOnly = status === 'ready' && passwordRef.current !== undefined

  const clampScale = (s: number) => Math.min(MAX_SCALE, Math.max(MIN_SCALE, s))

  /** Overall size of a row (in spread mode widths add up, including the page gap) */
  const rowSize = useCallback(
    (row: number[]): PageSize => {
      const dims = row.map((i) => dispSize(i))
      return {
        width: dims.reduce((w, d) => w + d.width, 0) + (dims.length - 1) * PAGE_GAP,
        height: Math.max(...dims.map((d) => d.height)),
      }
    },
    [dispSize],
  )

  const recomputeFit = useCallback(() => {
    const mode = fitModeRef.current
    const el = scrollRef.current
    if (!mode || !el || rows.length === 0) return
    const dims = rows.map((r) => rowSize(r))
    const maxW = Math.max(...dims.map((s) => s.width))
    const availW = el.clientWidth - SCROLL_PAD * 2
    if (mode === 'width') {
      setScale(clampScale(availW / maxW))
    } else {
      const maxH = Math.max(...dims.map((s) => s.height))
      setScale(clampScale(Math.min(availW / maxW, (el.clientHeight - PAGE_GAP * 2) / maxH)))
    }
  }, [rows, rowSize])

  useEffect(() => {
    if (status !== 'ready') return
    recomputeFit()
    const el = scrollRef.current
    if (!el) return
    const ro = new ResizeObserver(recomputeFit)
    ro.observe(el)
    return () => ro.disconnect()
  }, [status, recomputeFit])

  const rowTops = useMemo(
    () =>
      rowTopOffsets(
        rows.map((row) => rowSize(row).height),
        scale,
        PAGE_GAP,
      ),
    [rows, rowSize, scale],
  )
  const rowTop = useCallback((rowIdx: number) => rowTops[rowIdx] ?? PAGE_GAP, [rowTops])

  /** Page-top offset of a visible position (used for search positioning) */
  const pageTop = useCallback((visIdx: number) => rowTop(rowOfVis(visIdx)), [rowTop, rowOfVis])

  const handleScroll = useCallback(() => {
    const el = scrollRef.current
    if (!el || rows.length === 0) return
    const anchor = el.scrollTop + el.clientHeight * 0.4
    let rowIdx = 0
    for (let i = 0; i < rows.length; i++) {
      if (rowTop(i) <= anchor) rowIdx = i
      else break
    }
    const page = visList.indexOf(rows[rowIdx]![0]!) + 1
    setCurrentPage(page)
  }, [rows, rowTop, visList])

  const scrollToPage = (n: number) => {
    const el = scrollRef.current
    if (!el) return
    const target = Math.min(Math.max(1, n), pageCount)
    el.scrollTop = rowTop(rowOfVis(target - 1)) - PAGE_GAP / 2
  }

  /** Scale scroll position proportionally when zooming so the visual anchor doesn't jump */
  const applyScale = (next: number, mode: FitMode) => {
    fitModeRef.current = mode
    const el = scrollRef.current
    const clamped = clampScale(next)
    if (el && scale > 0) {
      const ratio = clamped / scale
      requestAnimationFrame(() => {
        el.scrollTop *= ratio
      })
    }
    setScale(clamped)
  }

  const zoomIn = () => applyScale(ZOOM_STEPS.find((s) => s > scale + 0.001) ?? MAX_SCALE, null)
  const zoomOut = () =>
    applyScale([...ZOOM_STEPS].reverse().find((s) => s < scale - 0.001) ?? MIN_SCALE, null)

  const dirty = hasUnsavedEdits({
    markups,
    drawings,
    stampCfg,
    formEdits,
    rotations,
    deleted,
    order,
    metadata,
  })
  dirtyRef.current = dirty

  // Mirror dirty state to the main process (close-tab/close-window guard)
  useEffect(() => {
    window.pdfApi.setDirty(dirty)
  }, [dirty])

  // ── Undo/redo: push a full snapshot before each change; consecutive input on the same form field coalesces into one step ──

  const snapshot = (): EditSnapshot => ({
    markups,
    drawings,
    stampCfg,
    formEdits,
    rotations,
    deleted,
    order,
    metadata,
  })

  const pushUndo = (coalesceKey?: string) => {
    if (coalesceKey && coalesceKeyRef.current === coalesceKey) return
    coalesceKeyRef.current = coalesceKey ?? null
    setUndoStack((prev) => [...prev.slice(-49), snapshot()])
    setRedoStack([])
  }

  const applySnapshot = (s: EditSnapshot) => {
    setMarkups(s.markups)
    setDrawings(s.drawings)
    setStampCfg(s.stampCfg)
    setFormEdits(s.formEdits)
    setRotations(s.rotations)
    setDeleted(s.deleted)
    setOrder(s.order)
    setMetadata(s.metadata)
    // The selected annotation may no longer exist in the restored snapshot
    setSelected(null)
  }

  const undo = () => {
    const top = undoStack[undoStack.length - 1]
    if (!top) return
    setRedoStack((r) => [...r, snapshot()])
    setUndoStack((u) => u.slice(0, -1))
    applySnapshot(top)
    coalesceKeyRef.current = null
  }

  const redo = () => {
    const top = redoStack[redoStack.length - 1]
    if (!top) return
    setUndoStack((u) => [...u, snapshot()])
    setRedoStack((r) => r.slice(0, -1))
    applySnapshot(top)
    coalesceKeyRef.current = null
  }

  // ── Full-text search ──

  /** Text index cached per doc; invalidated and rebuilt after a save reload */
  const getSearchIndex = useCallback((): Promise<SearchIndex> | null => {
    if (!doc) return null
    if (searchIndexRef.current?.doc !== doc) {
      searchIndexRef.current = { doc, promise: buildSearchIndex(doc) }
    }
    return searchIndexRef.current.promise
  }, [doc])

  useEffect(() => {
    if (!searchOpen || !searchQuery.trim()) {
      setSearchMatches([])
      setSearchCur(0)
      return
    }
    let cancelled = false
    const timer = setTimeout(() => {
      void getSearchIndex()?.then((idx) => {
        if (cancelled) return
        setSearchMatches(searchInIndex(idx, searchQuery.trim()))
        setSearchCur(0)
      })
    }, 200)
    return () => {
      cancelled = true
      clearTimeout(timer)
    }
  }, [searchOpen, searchQuery, getSearchIndex])

  /** Pages with unsaved deletion are excluded from match navigation */
  const activeMatches = useMemo(
    () => searchMatches.filter((m) => !deleted.has(m.pageIndex)),
    [searchMatches, deleted],
  )
  const searchCurClamped = Math.min(searchCur, Math.max(0, activeMatches.length - 1))

  const gotoMatch = useCallback(
    (idx: number) => {
      const m = activeMatches[idx]
      const el = scrollRef.current
      if (!m || !el) return
      const visIdx = visList.indexOf(m.pageIndex)
      if (visIdx < 0) return
      const box = pdfRectToCss(pageGeom(m.pageIndex), m.rects[0] ?? [0, 0, 0, 0], scale)
      el.scrollTop = Math.max(0, pageTop(visIdx) + box.top - el.clientHeight * 0.35)
    },
    [activeMatches, visList, pageGeom, scale, pageTop],
  )

  // Scroll to the current match on new results or position changes (unrelated changes like zoom don't re-scroll)
  useEffect(() => {
    if (!searchOpen || activeMatches.length === 0) return
    const last = searchJumpRef.current
    if (last && last.matches === activeMatches && last.cur === searchCurClamped) return
    searchJumpRef.current = { matches: activeMatches, cur: searchCurClamped }
    gotoMatch(searchCurClamped)
  }, [searchOpen, activeMatches, searchCurClamped, gotoMatch])

  const searchStep = (dir: 1 | -1) => {
    const n = activeMatches.length
    if (n === 0) return
    setSearchCur((searchCurClamped + dir + n) % n)
  }

  const openSearch = () => {
    setSearchOpen(true)
    requestAnimationFrame(() => {
      searchInputRef.current?.focus()
      searchInputRef.current?.select()
    })
  }

  const closeSearch = () => setSearchOpen(false)

  /** Mouse released over selected text → show the markup bar centered above the selection box (below if it doesn't fit) */
  const handleMouseUp = () => {
    setTimeout(() => {
      const el = scrollRef.current
      const sel = window.getSelection()
      if (!el || !sel || sel.isCollapsed || sel.rangeCount === 0) {
        setSelPopup(null)
        return
      }
      if (!el.contains(sel.getRangeAt(0).commonAncestorContainer)) return
      if (readOnly) return
      const box = sel.getRangeAt(0).getBoundingClientRect()
      if (box.width < 1 && box.height < 1) return
      setSelPopup({
        x: Math.min(Math.max(box.left + box.width / 2, 70), window.innerWidth - 70),
        y: box.top >= 52 ? box.top - 44 : Math.min(box.bottom + 8, window.innerHeight - 44),
      })
    }, 0)
  }

  const applyMarkup = (type: MarkupType) => {
    const el = scrollRef.current
    if (!el || readOnly) return
    const byVisPage = selectionQuadsByPage(
      el,
      visList.map((i) => pageGeom(i)),
      scale,
    )
    setSelPopup(null)
    if (!byVisPage) return
    const added: LocalMarkup[] = []
    for (const [visIdx, quads] of byVisPage) {
      const origIdx = visList[visIdx]
      if (origIdx === undefined) continue
      added.push({
        id: `m${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
        pageIndex: origIdx,
        type,
        color: MARKUP_COLORS[type],
        quads,
      })
    }
    if (added.length === 0) return
    pushUndo()
    setMarkups((prev) => [...prev, ...added])
    window.getSelection()?.removeAllRanges()
  }

  // ── Annotation selection: click selects, deletion is explicit (delete popup / Delete key) ──

  /** Clamp the delete popup anchor into the window, preferring a spot above the click (same rules as the markup bar) */
  const popupPos = (x: number, y: number) => ({
    x: Math.min(Math.max(x, 70), window.innerWidth - 70),
    y: y >= 96 ? y - 48 : Math.min(y + 12, window.innerHeight - 44),
  })

  /** Shift a drawing by a PDF-space delta (drag-to-move on the page) */
  const moveDrawing = (id: string, dx: number, dy: number) => {
    pushUndo()
    setSelected(null)
    setDrawings((prev) =>
      prev.map((d) => {
        if (d.id !== id) return d
        const input = d.input
        switch (input.kind) {
          case 'ink':
            return {
              ...d,
              input: {
                ...input,
                paths: input.paths.map((p) => p.map((v, i) => (i % 2 === 0 ? v + dx : v + dy))),
              },
            }
          case 'rect':
          case 'ellipse':
          case 'image':
          case 'freetext':
            return {
              ...d,
              input: {
                ...input,
                rect: [
                  input.rect[0] + dx,
                  input.rect[1] + dy,
                  input.rect[2] + dx,
                  input.rect[3] + dy,
                ] as [number, number, number, number],
              },
            }
          case 'line':
          case 'arrow':
            return {
              ...d,
              input: {
                ...input,
                from: [input.from[0] + dx, input.from[1] + dy] as [number, number],
                to: [input.to[0] + dx, input.to[1] + dy] as [number, number],
              },
            }
          case 'polyline':
          case 'polygon':
            return {
              ...d,
              input: {
                ...input,
                points: input.points.map((v, i) => (i % 2 === 0 ? v + dx : v + dy)),
              },
            }
          case 'caret':
            return {
              ...d,
              input: {
                ...input,
                at: [input.at[0] + dx, input.at[1] + dy] as [number, number],
              },
            }
          default:
            return d
        }
      }),
    )
  }

  /** Replace an image drawing's rect (corner-handle resize) */
  const resizeDrawing = (id: string, rect: [number, number, number, number]) => {
    pushUndo()
    setDrawings((prev) =>
      prev.map((d) =>
        d.id === id && (d.input.kind === 'image' || d.input.kind === 'freetext')
          ? { ...d, input: { ...d.input, rect } }
          : d,
      ),
    )
  }

  const deleteSelected = () => {
    const sel = selected
    if (!sel) return
    pushUndo()
    if (sel.kind === 'markup') setMarkups((prev) => prev.filter((m) => m.id !== sel.id))
    else if (sel.kind === 'drawing') setDrawings((prev) => prev.filter((d) => d.id !== sel.id))
    else setStampCfg(null)
    setSelected(null)
    // Transient "deleted · undo" toast so the removal is visible and reversible in place
    if (toastTimerRef.current !== null) window.clearTimeout(toastTimerRef.current)
    setDeleteToast(true)
    toastTimerRef.current = window.setTimeout(() => setDeleteToast(false), 5000)
  }

  const opFailed = (error: string) => {
    setSaveError(error)
    setSaveState('error')
  }

  /** Pending edits in SavePdfRequest form; shared by in-place Save and Save As */
  const editsPayload = () => ({
    markups: markups.map(({ id: _id, ...rest }) => rest),
    drawings: drawings.map((d) => d.input),
    stamps: stampCfg ? renderStamps(stampCfg, visList) : [],
    formValues: [...formEdits.values()],
    rotations: [...rotations].map(([pageIndex, delta]) => ({ pageIndex, delta })),
    deletedPages: [...deleted],
    ...(order ? { pageOrder: visList } : {}),
    ...(metadata ? { metadata } : {}),
  })

  /** Resolved when the running save() lands; Save As serializes behind it */
  const saveInFlightRef = useRef<Promise<boolean> | null>(null)

  const save = (autosave = false): Promise<boolean> => {
    if (!dirty || saveState === 'saving' || !filePath) return Promise.resolve(!dirty)
    // An explicit save opts this file into autosave
    if (!autosave) savedOnceRef.current = true
    const run = (async (): Promise<boolean> => {
      setSaveState('saving')
      const result = await window.pdfApi.save({ path: filePath, ...editsPayload() })
      if (!result.ok) {
        opFailed(result.error)
        return false
      }
      // Reload: changes are in the file now, canvas renders directly, overlays/pending ops are cleared
      try {
        const el = scrollRef.current
        const scrollTop = el?.scrollTop ?? 0
        await loadDoc(filePath, doc)
        requestAnimationFrame(() => {
          if (scrollRef.current) scrollRef.current.scrollTop = scrollTop
        })
      } catch {
        /* Save already succeeded; a reload failure doesn't block (takes effect on next open) */
      }
      setSaveState('saved')
      setTimeout(() => setSaveState((s) => (s === 'saved' ? 'idle' : s)), 2000)
      return true
    })()
    const tracked = run.finally(() => {
      if (saveInFlightRef.current === tracked) saveInFlightRef.current = null
    })
    saveInFlightRef.current = tracked
    return tracked
  }

  // Autosave (same strategy as Docs): every 30s and on window blur, silently persist pending
  // edits via the regular save() path; skipped while a save is in flight or without a file path.
  // Gated on one explicit save first: a PDF opened only to read must never be
  // overwritten because a thumbnail got dragged or a markup tool tapped — Save (⌘S / the
  // toolbar button / File ▸ Save) is what opts this file into unattended writes.
  useAutosave(
    () => savedOnceRef.current && dirty && saveState !== 'saving' && filePath !== '' && !readOnly,
    () => void save(true),
  )

  // ── Page operations ──

  const rotatePage = (origIdx: number, dir: 90 | -90) => {
    if (readOnly) return
    pushUndo()
    setRotations((prev) => {
      const next = new Map(prev)
      const nv = ((next.get(origIdx) ?? 0) + dir + 360) % 360
      if (nv === 0) next.delete(origIdx)
      else next.set(origIdx, nv)
      return next
    })
    // Image stamps are always drawn upright (both in the overlay and in the saved
    // appearance), so a 90° page turn swaps their displayed width/height. Swap the
    // user-space rect around its center to keep the bitmap's aspect ratio intact.
    setDrawings((prev) =>
      prev.map((d) => {
        if (d.input.kind !== 'image' || d.input.pageIndex !== origIdx) return d
        const [x1, y1, x2, y2] = d.input.rect
        const cx = (x1 + x2) / 2
        const cy = (y1 + y2) / 2
        const hw = (x2 - x1) / 2
        const hh = (y2 - y1) / 2
        return { ...d, input: { ...d.input, rect: [cx - hh, cy - hw, cx + hh, cy + hw] } }
      }),
    )
  }

  const deletePage = (origIdx: number) => {
    if (pageCount <= 1 || readOnly) return
    pushUndo()
    setDeleted((prev) => new Set(prev).add(origIdx))
    setMarkups((prev) => prev.filter((m) => m.pageIndex !== origIdx))
    setDrawings((prev) => prev.filter((d) => d.input.pageIndex !== origIdx))
  }

  // ── Drawing annotations ──

  const newId = () => `d${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`

  const commitDrawing = (origIdx: number, input: DrawingInput) => {
    pushUndo()
    setDrawings((prev) => [...prev, { id: newId(), input: { ...input, pageIndex: origIdx } }])
  }

  /** Render stamps in current page order; page numbers depend on visList, so both preview and save compute fresh */
  const renderStamps = useCallback(
    (cfg: StampConfig, pages: number[]): StampInput[] =>
      buildStamps(
        pages.map((origIdx, i) => ({
          origIdx,
          pw: sizes[origIdx]!.width,
          ph: sizes[origIdx]!.height,
          displayNo: i + 1,
        })),
        cfg.wm,
        cfg.hf,
      ),
    [sizes],
  )

  const applyStamps = (wm: WatermarkConfig | null, hf: HeaderFooterConfig | null) => {
    setStampDlg(false)
    if (!wm && !hf) return
    pushUndo()
    setStampCfg({ wm, hf })
  }

  /** Stamp preview for visible pages (only pages in rendered rows, so large docs don't render every canvas) */
  const stampPreview = useMemo(() => {
    if (!stampCfg) return new Map<number, StampInput[]>()
    const shown = new Set([...visibleRows].flatMap((r) => rows[r] ?? []))
    const byPage = new Map<number, StampInput[]>()
    for (const s of renderStamps(stampCfg, visList)) {
      if (!shown.has(s.pageIndex)) continue
      const list = byPage.get(s.pageIndex)
      if (list) list.push(s)
      else byPage.set(s.pageIndex, [s])
    }
    return byPage
  }, [stampCfg, visList, rows, visibleRows, renderStamps])

  /** Thumbnail drag-and-drop reorder: move the page at position from to position to */
  const movePage = (from: number, to: number) => {
    if (from === to || readOnly) return
    pushUndo()
    const next = [...visList]
    const [moved] = next.splice(from, 1)
    next.splice(to, 0, moved!)
    // order must cover all original pages (deleted ones included, kept at the tail so they don't affect the result)
    const rest = sizes.map((_, i) => i).filter((i) => !next.includes(i))
    setOrder([...next, ...rest])
  }

  /** Place signature centered on the click point (view coords, scale=1); sized via signPlaceK
      to match the ghost preview. Points map view→PDF individually so it stays upright on
      rotated pages. */
  const placeSignature = (origIdx: number, vx: number, vy: number) => {
    const sig = pendingSign
    if (!sig) return
    const geom = pageGeom(origIdx)
    const disp = geomDispSize(geom)
    const k = signPlaceK(sig, disp.width, disp.height)
    const targetW = sig.width * k
    const targetH = sig.height * k
    const left = Math.min(Math.max(vx - targetW / 2, 0), Math.max(disp.width - targetW, 0))
    const top = Math.min(Math.max(vy - targetH / 2, 0), Math.max(disp.height - targetH, 0))
    pushUndo()
    if (sig.kind === 'image') {
      const [ax, ay] = viewToPdf(geom, left, top)
      const [bx, by] = viewToPdf(geom, left + targetW, top + targetH)
      const rect: [number, number, number, number] = [
        Math.min(ax, bx),
        Math.min(ay, by),
        Math.max(ax, bx),
        Math.max(ay, by),
      ]
      setDrawings((prev) => [
        ...prev,
        { id: newId(), input: { kind: 'image', pageIndex: origIdx, image: sig.image, rect } },
      ])
    } else {
      const paths = sig.paths.map((p) => {
        const out: number[] = []
        for (let i = 0; i < p.length; i += 2) {
          out.push(...viewToPdf(geom, left + p[i]! * k, top + p[i + 1]! * k))
        }
        return out
      })
      setDrawings((prev) => [
        ...prev,
        {
          id: newId(),
          input: { kind: 'ink', pageIndex: origIdx, color: drawColor, width: 1.6, paths },
        },
      ])
    }
    setPendingSign(null)
  }

  /** Export PNG: current page or all visible pages, 150dpi equivalent */
  const exportImages = (allPages: boolean) =>
    flushThen(async () => {
      if (!doc || exporting) return
      setExporting(true)
      try {
        const targets = allPages ? visList : [curOrigIdx].filter((i) => i >= 0)
        const images: string[] = []
        const pageNumbers: number[] = []
        const canvas = document.createElement('canvas')
        for (const origIdx of targets) {
          const page = await doc.getPage(origIdx + 1)
          const viewport = page.getViewport({
            scale: 150 / 72,
            rotation: (page.rotate + rotDelta(origIdx)) % 360,
          })
          canvas.width = Math.floor(viewport.width)
          canvas.height = Math.floor(viewport.height)
          await page.render({ canvas, viewport }).promise
          images.push(canvas.toDataURL('image/png').split(',')[1] ?? '')
          pageNumbers.push(visList.indexOf(origIdx) + 1)
        }
        canvas.width = 0
        canvas.height = 0
        const result = await window.pdfApi.exportImages({
          images,
          pageNumbers,
          baseName: fileName.replace(/\.pdf$/i, ''),
        })
        if (!result.ok) opFailed(result.error)
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    })

  /** Export the document's text. Reuses the cached search index, so nothing is re-extracted */
  const exportText = () =>
    flushThen(async () => {
      if (!doc || exporting) return
      setExporting(true)
      try {
        const index = await getSearchIndex()
        const text = index ? indexToText(index) : ''
        if (!text) {
          opFailed(t('exportTextEmpty'))
          return
        }
        const result = await window.pdfApi.exportText({
          text,
          suggestedName: `${fileName.replace(/\.pdf$/i, '')}.txt`,
        })
        if (!result.ok) opFailed(result.error)
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    })

  /** Export the file's embedded bitmaps. Reads the saved file, so unsaved edits flush first */
  const doExportEmbeddedImages = () =>
    flushThen(async () => {
      if (!doc || exporting) return
      setExporting(true)
      try {
        const result = await window.pdfApi.exportEmbeddedImages({ path: filePath })
        if (!result.ok) opFailed(result.error)
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    })

  const confirmNote = () => {
    const target = notePrompt
    const text = noteText.trim()
    setNotePrompt(null)
    setNoteText('')
    if (!target || !text) return
    // Typewriter text: rasterize the glyphs (CJK-safe, same reason as the watermark), then hang
    // the annotation box down-right of the click, which is where the text visually starts
    if (target.kind === 'freetext') {
      const box = renderTextBox(text, FREETEXT_SIZE, cssRgb(drawColor))
      if (!box) return
      const [x, y] = target.at
      pushUndo()
      setDrawings((prev) => [
        ...prev,
        {
          id: newId(),
          input: {
            kind: 'freetext',
            pageIndex: target.origIdx,
            image: box.image,
            rect: [x, y - box.height, x + box.width, y],
            contents: text,
          },
        },
      ])
      return
    }
    pushUndo()
    setDrawings((prev) => [
      ...prev,
      {
        id: newId(),
        input:
          target.kind === 'caret'
            ? {
                kind: 'caret',
                pageIndex: target.origIdx,
                color: drawColor,
                at: target.at,
                size: CARET_SIZE,
                contents: text,
              }
            : {
                kind: 'note',
                pageIndex: target.origIdx,
                color: drawColor,
                at: target.at,
                contents: text,
              },
      },
    ])
  }

  /** Extract/insert work on the file on disk — flush unsaved changes first */
  const flushThen = async (fn: () => Promise<void>) => {
    if (dirty && !(await save())) return
    await fn()
  }

  const extractPage = (origIdx: number) =>
    flushThen(async () => {
      const base = fileName.replace(/\.pdf$/i, '')
      const result = await window.pdfApi.extractPages({
        path: filePath,
        pages: [origIdx],
        suggestedName: `${base}-p${origIdx + 1}.pdf`,
      })
      if (!result.ok) opFailed(result.error)
    })

  const openExtractDlg = () => {
    setExtractInput(String(currentPage))
    setExtractInvalid(false)
    setExtractDlg(true)
  }

  /** Extract dialog confirm: visible page-number ranges → original page indices */
  const confirmExtract = () => {
    const pages = parsePageRanges(extractInput, pageCount)
    if (!pages) {
      setExtractInvalid(true)
      return
    }
    setExtractDlg(false)
    void flushThen(async () => {
      const base = fileName.replace(/\.pdf$/i, '')
      const label = pages.length === 1 ? `p${pages[0]}` : `p${pages[0]}-${pages[pages.length - 1]}`
      const result = await window.pdfApi.extractPages({
        path: filePath,
        pages: pages.map((n) => visList[n - 1]!),
        suggestedName: `${base}-${label}.pdf`,
      })
      if (!result.ok) opFailed(result.error)
    })
  }

  const insertPdf = (afterOrigIdx: number) =>
    flushThen(async () => {
      const result = await window.pdfApi.insertPdf({ path: filePath, afterPageIndex: afterOrigIdx })
      if (!result.ok) {
        opFailed(result.error)
        return
      }
      if (!('canceled' in result)) await loadDoc(filePath, doc)
    })

  /** Print: save first (markups/forms/page ops all into the file), then reload from the file to render, avoiding a destroyed old doc */
  const printDoc = () => {
    if (!doc) return
    setShowPrintPreview(true)
  }

  /** Capability surface for AI tools; rebuilt each render (AiPanel mirrors it via refs to get the latest) */
  const aiApi: PdfSkillDeps = {
    doc: () => doc,
    fileName: () => fileName,
    pageCount: () => sizes.length,
    currentPage: () => (visList[currentPage - 1] ?? 0) + 1,
    readOnly: () => readOnly,
    outline: () => outline,
    searchIndex: getSearchIndex,
    isDeleted: (i) => deleted.has(i),
    gotoPage: (p) => {
      const visIdx = visList.indexOf(p - 1)
      if (visIdx < 0) return false
      scrollToPage(visIdx + 1)
      return true
    },
    addMarkup: (type, origIdx, rects) => {
      pushUndo()
      const quads = rects.map((r) => [r[0], r[3], r[2], r[3], r[0], r[1], r[2], r[1]])
      setMarkups((prev) => [
        ...prev,
        {
          id: `m${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
          pageIndex: origIdx,
          type,
          color: MARKUP_COLORS[type],
          quads,
        },
      ])
    },
    formEdits: () => formEdits,
    applyFormEdit: (v) => {
      pushUndo()
      setFormEdits((prev) => new Map(prev).set(v.name, v))
    },
    rotatePage,
    deletePage: (origIdx) => {
      if (pageCount <= 1 || readOnly) return false
      deletePage(origIdx)
      return true
    },
  }

  /** Internal destination of a Link annotation → jump to that page */
  const goToDest = async (dest: unknown) => {
    if (!doc) return
    try {
      const arr = typeof dest === 'string' ? await doc.getDestination(dest) : dest
      if (!Array.isArray(arr)) return
      const ref = arr[0]
      const origIdx =
        typeof ref === 'number'
          ? ref
          : await doc.getPageIndex(ref as Parameters<PDFDocumentProxy['getPageIndex']>[0])
      const visIdx = visList.indexOf(origIdx)
      if (visIdx >= 0) scrollToPage(visIdx + 1)
    } catch {
      /* Ignore corrupted destinations */
    }
  }

  // ── Bookmark editing ──
  const [outlineEditMode, setOutlineEditMode] = useState(false)
  /** The tree being edited plus the display order it resolved against, so Apply can rebase */
  const [outlineDraft, setOutlineDraft] = useState<{
    nodes: OutlineEditNode[]
    visList: number[]
  } | null>(null)
  /** Text prompts for rename/add/addChild; path addresses the node the op targets */
  const [outlinePrompt, setOutlinePrompt] = useState<{
    mode: 'rename' | 'add' | 'addChild'
    path: number[]
    text: string
  } | null>(null)

  /**
   * Resolve the pdfjs outline into an editable tree.
   *
   * Pages become 0-based positions in the current display order (visList), which is what Apply
   * rebases against. A bookmark whose destination is gone keeps nothing: its page is not in the
   * document any more. Web link bookmarks keep their url untouched.
   */
  const enterOutlineEdit = async () => {
    if (!doc || !outline || readOnly) return
    // Captured: the closures below outlive this guard, and TS does not narrow through them
    const loaded = doc
    async function walk(n: OutlineNode): Promise<OutlineEditNode | null> {
      if (n.url) {
        return { title: n.title, pageIndex: -1, uri: n.url, children: await walkAll(n) }
      }
      if (n.dest == null) return null
      try {
        const arr = typeof n.dest === 'string' ? await loaded.getDestination(n.dest) : n.dest
        if (!Array.isArray(arr)) return null
        const ref = arr[0]
        const origIdx =
          typeof ref === 'number'
            ? ref
            : await loaded.getPageIndex(ref as Parameters<PDFDocumentProxy['getPageIndex']>[0])
        const displayIdx = visList.indexOf(origIdx)
        if (displayIdx < 0) return null
        return { title: n.title, pageIndex: displayIdx, children: await walkAll(n) }
      } catch {
        return null
      }
    }
    async function walkAll(n: OutlineNode): Promise<OutlineEditNode[]> {
      return (await Promise.all((n.items ?? []).map(walk))).filter(
        (child): child is OutlineEditNode => child !== null,
      )
    }
    setOutlineEditMode(true)
    setOutlineDraft({
      nodes: (await Promise.all(outline.map(walk))).filter((n): n is OutlineEditNode => n !== null),
      visList,
    })
  }

  /** Escape closes the a11y dialog too: rows of text, no text field to take the key */
  useEffect(() => {
    if (!a11yDlg) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setA11yDlg(false)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [a11yDlg])

  /** Collect the accessibility report through the IPC grant, then show it */
  const openA11yReport = () =>
    flushThen(async () => {
      if (!doc || exporting) return
      setExporting(true)
      try {
        const result = await window.pdfApi.a11yCheck({ path: filePath })
        if (!result.ok) {
          opFailed(result.error)
          return
        }
        setA11yReport(result.report)
        setA11yDlg(true)
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    })

  /** One-click fill of the report's two metadata gaps: title from the file name, lang from the UI */
  const doA11yFix = () =>
    flushThen(async () => {
      if (!doc || exporting) return
      setExporting(true)
      try {
        const request: A11yFixRequest = { path: filePath, lang: htmlLang(uiLang) }
        const result: A11yFixResult = await window.pdfApi.a11yFix(request)
        if (!result.ok) {
          opFailed(result.error)
          return
        }
        // The file changed on disk: reload like sanitize does, then re-run the report in place so
        // the dialog's own rows flip to pass
        await loadDoc(filePath, doc)
        const fresh = await window.pdfApi.a11yCheck({ path: filePath })
        if (fresh.ok) setA11yReport(fresh.report)
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    })

  const endOutlineEdit = () => {
    setOutlineEditMode(false)
    setOutlineDraft(null)
  }

  /** Comments panel: a flat row per annotation; the file's annots arrive over IPC-backed walk */
  const [commentRows, setCommentRows] = useState<CommentRow[]>([])
  const [commentsKey, setCommentsKey] = useState(0)

  useEffect(() => {
    if (sidebar !== 'comments') return
    let cancelled = false
    const walk = async () => {
      const current = doc
      if (!current) {
        setCommentRows([])
        return
      }
      try {
        const out: CommentRow[] = []
        for (let i = 0; i < current.numPages; i++) {
          const page = await current.getPage(i + 1)
          if (cancelled) return
          const annots = await page.getAnnotations()
          if (cancelled) return
          out.push(...annotationsToRows(annots, i, (_subtype, key) => tModule(key)))
        }
        setCommentRows(out)
      } catch {
        setCommentRows([])
      }
    }
    void walk()
    return () => {
      cancelled = true
    }
    // The walk reads the doc and the open state; reloads change `doc`, deletions bump commentsKey
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sidebar === 'comments', doc, commentsKey])

  const jumpComment = (row: CommentRow) => {
    const visIdx = visList.indexOf(row.pageIndex)
    if (visIdx < 0) return
    scrollToPage(visIdx + 1)
  }

  const removeComment = (row: CommentRow) => {
    if (!row.saved) {
      // Pending markup: the same undo-backed state removal the on-canvas delete popup does
      pushUndo()
      setMarkups((prev) => prev.filter((m) => m.id !== row.id))
      return
    }
    void flushThen(async () => {
      if (!doc || exporting) return
      setExporting(true)
      try {
        const result: DeleteAnnotationsResult = await window.pdfApi.deleteAnnotations({
          path: filePath,
          targets: [row.match],
        })
        if (!result.ok) {
          opFailed(result.error)
          // The file may have moved on since the row list was built: resync it
          setCommentsKey((k) => k + 1)
          return
        }
        await loadDoc(filePath, doc)
        setCommentsKey((k) => k + 1)
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    })
  }

  /** Escape closes the sanitize dialog: no text field inside it would otherwise take the key */
  useEffect(() => {
    if (!sanitizeDlg) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setSanitizeDlg(false)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [sanitizeDlg])

  const handleOutlineEdit = (path: number[], op: OutlineEditOp) => {
    if (op === 'rename') {
      const node = outlineDraft ? outlineNodeAt(outlineDraft.nodes, path) : null
      setOutlinePrompt({ mode: 'rename', path, text: node?.title ?? '' })
      return
    }
    if (op === 'addChild') {
      // Lands on the page in view, to be renamed by whoever just wrote it
      setOutlinePrompt({ mode: 'addChild', path, text: '' })
      return
    }
    if (op === 'delete') {
      setOutlineDraft((d) => (d ? { ...d, nodes: removeOutlineNode(d.nodes, path) } : d))
      return
    }
    setOutlineDraft((d) =>
      d ? { ...d, nodes: moveOutlineNode(d.nodes, path, op === 'up' ? -1 : 1) } : d,
    )
  }

  const openOutlinePrompt = (mode: 'add' | 'addChild', path: number[]): void =>
    setOutlinePrompt({ mode, path, text: '' })

  const confirmOutlinePrompt = () => {
    const prompt = outlinePrompt
    const title = (prompt?.text ?? '').trim()
    setOutlinePrompt(null)
    if (!prompt || !title) return
    const newPage = currentPage - 1
    setOutlineDraft((d) => {
      if (!d) return d
      if (prompt.mode === 'rename') {
        return { ...d, nodes: renameOutlineNode(d.nodes, prompt.path, title) }
      }
      const entry: OutlineEditNode = { title, pageIndex: newPage }
      if (prompt.mode === 'add') {
        return { ...d, nodes: mapOutlinePath(d.nodes, [], (list) => [...list, entry]) }
      }
      return {
        ...d,
        nodes: mapOutlinePath(d.nodes, prompt.path, (list) =>
          list.map((node, i) =>
            i === prompt.path[prompt.path.length - 1]
              ? { ...node, children: [...(node.children ?? []), entry] }
              : node,
          ),
        ),
      }
    })
  }

  /**
   * Write the draft into the file. flushThen first keeps the file true to the display order the
   * draft was rebased against, and the reload afterwards replaces the sidebar with what landed.
   */
  const doApplyOutline = () =>
    flushThen(async () => {
      const draft = outlineDraft
      if (!doc || !draft || exporting) return
      setExporting(true)
      try {
        const nodes = rebaseOutline(draft.nodes, draft.visList, visList)
        const result = await window.pdfApi.setOutline({ path: filePath, nodes })
        if (!result.ok) {
          opFailed(result.error)
          return
        }
        endOutlineEdit()
        await loadDoc(filePath, doc)
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    })

  const curOrigIdx = visList[currentPage - 1] ?? -1

  // Clicking elsewhere closes the thumbnail context menu
  useEffect(() => {
    if (!thumbMenu) return
    const close = (e: PointerEvent) => {
      if (!(e.target as Element | null)?.closest?.('.thumb-menu')) setThumbMenu(null)
    }
    window.addEventListener('pointerdown', close)
    return () => window.removeEventListener('pointerdown', close)
  }, [thumbMenu])

  // Clicking elsewhere closes the draw-color palette
  useEffect(() => {
    if (!colorOpen) return
    const close = (e: PointerEvent) => {
      if (!(e.target as Element | null)?.closest?.('.rb-drop-wrap')) setColorOpen(false)
    }
    window.addEventListener('pointerdown', close)
    return () => window.removeEventListener('pointerdown', close)
  }, [colorOpen])

  // Main process picked "Save" in the close prompt → save and report the result
  useEffect(() => {
    return window.pdfApi.onCloseSaveRequest(() => {
      void save().then((ok) => window.pdfApi.sendCloseSaveResult(ok))
    })
  })

  // Shortcuts: ⌘S/⌘F/⌘P/⌘±/⌘0 + page navigation (only ⌘ combos kept while an input control is focused)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null
      const inEditable =
        !!target &&
        (target.tagName === 'INPUT' ||
          target.tagName === 'TEXTAREA' ||
          target.tagName === 'SELECT' ||
          target.isContentEditable)
      if (e.metaKey || e.ctrlKey) {
        const k = e.key.toLowerCase()
        if (k === 's') {
          e.preventDefault()
          void save()
        } else if (k === 'z') {
          e.preventDefault()
          if (e.shiftKey) redo()
          else undo()
        } else if (k === 'f') {
          e.preventDefault()
          openSearch()
        } else if (k === 'p' && !e.shiftKey) {
          e.preventDefault()
          void printDoc()
        } else if (e.key === '=' || e.key === '+') {
          e.preventDefault()
          zoomIn()
        } else if (e.key === '-') {
          e.preventDefault()
          zoomOut()
        } else if (e.key === '0') {
          e.preventDefault()
          fitModeRef.current = 'width'
          recomputeFit()
        }
        return
      }
      if (e.key === 'Escape') {
        if (pendingSign) setPendingSign(null)
        else if (drawTool) setDrawTool(null)
        else if (selected) setSelected(null)
        else if (searchOpen) closeSearch()
        return
      }
      if (inEditable) return
      if ((e.key === 'Delete' || e.key === 'Backspace') && selected) {
        e.preventDefault()
        deleteSelected()
        return
      }
      const el = scrollRef.current
      if (!el) return
      const inThumbs = !!thumbsRef.current?.contains(document.activeElement)
      const action = navAction(e.key, inThumbs)
      if (!action) return
      e.preventDefault()
      switch (action.type) {
        case 'scrollViewport':
          el.scrollTop += action.dir * (el.clientHeight - 40)
          break
        case 'scrollEdge':
          el.scrollTop = action.edge === 'top' ? 0 : el.scrollHeight
          break
        case 'scrollBy':
          el.scrollTop += action.delta
          break
        case 'stepPage': {
          const target = stepPage(visList, spread, currentPage, action.dir)
          scrollToPage(target)
          if (inThumbs) {
            const thumbEl = thumbsRef.current?.querySelector<HTMLElement>(
              `[data-idx="${target - 1}"]`,
            )
            thumbEl?.focus({ preventScroll: true })
            thumbEl?.scrollIntoView({ block: 'nearest' })
          }
          break
        }
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  })

  // Ctrl/⌘ + wheel zoom (native listener: React's wheel is passive and can't preventDefault)
  useEffect(() => {
    const el = scrollRef.current
    if (!el) return
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return
      e.preventDefault()
      if (e.deltaY < 0) zoomIn()
      else zoomOut()
    }
    el.addEventListener('wheel', onWheel, { passive: false })
    return () => el.removeEventListener('wheel', onWheel)
  })

  if (status === 'password') {
    return (
      <div className="app">
        <div className="pdf-placeholder">
          <form
            className="pdf-password"
            onSubmit={(e) => {
              e.preventDefault()
              passwordRef.current = pwInput
              setStatus('loading')
              void openPath(filePath)
            }}
          >
            <div className="pdf-password-title">{t('pwTitle')}</div>
            <input
              type="password"
              className="pdf-password-input"
              placeholder={t('pwPlaceholder')}
              value={pwInput}
              autoFocus
              onChange={(e) => setPwInput(e.target.value)}
            />
            {pwWrong && <div className="pdf-password-error">{t('pwWrong')}</div>}
            <button type="submit" className="pdf-password-btn" disabled={!pwInput}>
              {t('pwOpen')}
            </button>
          </form>
        </div>
      </div>
    )
  }

  if (status !== 'ready' || !doc) {
    return (
      <div className="app">
        <div className="pdf-placeholder">
          {status === 'loading' ? t('loading') : status === 'error' ? t('loadError') : t('noFile')}
        </div>
      </div>
    )
  }

  const menuOrig = thumbMenu?.origIdx ?? -1

  return (
    <div className="app">
      <div className="ribbon">
        <div className="ribbon-tabs">
          <button
            className="qa-btn"
            title={`${t('save')} (⌘S)`}
            aria-label={t('save')}
            disabled={!dirty || saveState === 'saving'}
            onClick={() => void save()}
          >
            <IconSave />
          </button>
          <button
            className="qa-btn"
            title={`${t('undo')} (⌘Z)`}
            disabled={undoStack.length === 0}
            onClick={undo}
          >
            <IconUndo />
          </button>
          <button
            className="qa-btn"
            title={`${t('redo')} (⇧⌘Z)`}
            disabled={redoStack.length === 0}
            onClick={redo}
          >
            <IconRedo />
          </button>
          <span className="ribbon-tabs-divider" />
          <button
            className={`ribbon-tab${ribbonTab === 'home' ? ' active' : ''}`}
            onClick={() => setRibbonTab('home')}
          >
            {t('tabHome')}
          </button>
          <button
            className={`ribbon-tab${ribbonTab === 'ai' ? ' active' : ''}`}
            onClick={() => setRibbonTab('ai')}
          >
            {t('tabAi')}
          </button>
          <button
            className={`ribbon-tab${ribbonTab === 'comment' ? ' active' : ''}`}
            onClick={() => setRibbonTab('comment')}
          >
            {t('tabComment')}
          </button>
          <button
            className={`ribbon-tab${ribbonTab === 'organize' ? ' active' : ''}`}
            onClick={() => setRibbonTab('organize')}
          >
            {t('tabOrganize')}
          </button>
          <button
            className={`ribbon-tab${ribbonTab === 'view' ? ' active' : ''}`}
            onClick={() => setRibbonTab('view')}
          >
            {t('tabView')}
          </button>
          <span className="ribbon-tabs-spacer" />
          <span className="ribbon-file" title={filePath}>
            {fileName}
          </span>
          {readOnly && <span className="tb-readonly">{t('roEncrypted')}</span>}
          {/* Unsaved-changes indicator next to the file name: the file on disk
              is only touched by an explicit save until then */}
          {saveState === 'saving' ? (
            <span className="tb-save-pending">{t('saving')}</span>
          ) : (
            dirty &&
            saveState !== 'error' && <span className="tb-save-pending">{t('unsaved')}</span>
          )}
          {saveState === 'error' && (
            <span className="tb-save-error" title={saveError}>
              {t('saveFailed')}
            </span>
          )}
          {saveState === 'saved' && <span className="tb-save-ok">{t('savedOk')}</span>}
        </div>
        <div className="ribbon-body">
          {ribbonTab === 'home' && (
            <>
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className={`rb-big ai-entry${aiCollapsed ? '' : ' active'}`}
                    title={t('aiOpenAssistant')}
                    onClick={() => setAiCollapsed((v) => !v)}
                  >
                    <span className="rb-big-icon">
                      <AssistantMark size={26} />
                    </span>
                    <span>ARA Assistant</span>
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className={`rb-big${sidebar === 'thumbs' ? ' active' : ''}`}
                    title={t('thumbs')}
                    onClick={() => setSidebar((v) => (v === 'thumbs' ? null : 'thumbs'))}
                  >
                    <span className="rb-big-icon">
                      <IconThumbs />
                    </span>
                    {t('thumbs')}
                  </button>
                  <button
                    className={`rb-big${sidebar === 'outline' ? ' active' : ''}`}
                    disabled={!outline}
                    title={t('outline')}
                    onClick={() => setSidebar((v) => (v === 'outline' ? null : 'outline'))}
                  >
                    <span className="rb-big-icon">
                      <IconOutline />
                    </span>
                    {t('outline')}
                  </button>
                  <button
                    className={`rb-big${sidebar === 'comments' ? ' active' : ''}`}
                    disabled={!doc}
                    title={t('comments')}
                    onClick={() => setSidebar((v) => (v === 'comments' ? null : 'comments'))}
                  >
                    <span className="rb-big-icon">
                      <IconComments />
                    </span>
                    {t('comments')}
                  </button>
                  <button
                    className={`rb-big${searchOpen ? ' active' : ''}`}
                    title={`${t('search')} (⌘F)`}
                    onClick={() => (searchOpen ? closeSearch() : openSearch())}
                  >
                    <span className="rb-big-icon">
                      <IconSearch />
                    </span>
                    {t('search')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    title={t('fitWidth')}
                    onClick={() => {
                      fitModeRef.current = 'width'
                      recomputeFit()
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconFitWidth />
                    </span>
                    {t('fitWidth')}
                  </button>
                  <button
                    className="rb-big"
                    title={t('fitPage')}
                    onClick={() => {
                      fitModeRef.current = 'page'
                      recomputeFit()
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconFitPage />
                    </span>
                    {t('fitPage')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={curOrigIdx < 0 || readOnly}
                    onClick={() => rotatePage(curOrigIdx, -90)}
                    title={t('rotateLeft')}
                  >
                    <span className="rb-big-icon">
                      <IconRotateL />
                    </span>
                    {t('rotateLeft')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={curOrigIdx < 0 || readOnly}
                    onClick={() => rotatePage(curOrigIdx, 90)}
                    title={t('rotateRight')}
                  >
                    <span className="rb-big-icon">
                      <IconRotateR />
                    </span>
                    {t('rotateRight')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group" onMouseDown={(e) => e.preventDefault()}>
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    title={t('highlight')}
                    onClick={() => applyMarkup('highlight')}
                  >
                    <span className="rb-big-icon">
                      <IconHighlight />
                    </span>
                    {t('highlight')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    title={t('underline')}
                    onClick={() => applyMarkup('underline')}
                  >
                    <span className="rb-big-icon">
                      <IconUnderline />
                    </span>
                    {t('underline')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    title={t('strikeout')}
                    onClick={() => applyMarkup('strikeout')}
                  >
                    <span className="rb-big-icon">
                      <IconStrike />
                    </span>
                    {t('strikeout')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    title={t('squiggly')}
                    onClick={() => applyMarkup('squiggly')}
                  >
                    <span className="rb-big-icon">
                      <IconSquiggly />
                    </span>
                    {t('squiggly')}
                  </button>
                  <button
                    className={`rb-big${drawTool === 'ink' ? ' active' : ''}`}
                    disabled={readOnly}
                    title={t('drawInk')}
                    onClick={() => setDrawTool((v) => (v === 'ink' ? null : 'ink'))}
                  >
                    <span className="rb-big-icon">
                      <IconInk />
                    </span>
                    {t('drawInk')}
                  </button>
                  <div className="rb-drop-wrap">
                    <button
                      className={`rb-big${colorOpen ? ' active' : ''}`}
                      disabled={readOnly}
                      title={t('drawColor')}
                      onClick={() => setColorOpen((v) => !v)}
                    >
                      <span className="rb-big-icon">
                        <span className="rb-big-icon-colored">
                          <IconDrawColor />
                          <span
                            className="rb-color-bar"
                            style={{ background: cssRgb(drawColor) }}
                          />
                        </span>
                        <RbCaret />
                      </span>
                      {t('drawColor')}
                    </button>
                    {colorOpen && (
                      <div className="rb-drop rb-color-grid">
                        {DRAW_COLORS.map((c) => (
                          <button
                            key={c.name}
                            className={`rb-swatch${cssRgb(drawColor) === cssRgb(c.rgb) ? ' active' : ''}`}
                            style={{ background: cssRgb(c.rgb) }}
                            title={c.name}
                            onClick={() => {
                              setDrawColor(c.rgb)
                              setColorOpen(false)
                            }}
                          />
                        ))}
                        <label className="rb-color-more" title={t('drawColor')}>
                          <input
                            type="color"
                            value={rgbToHex(drawColor)}
                            onChange={(e) => setDrawColor(hexToRgb(e.target.value))}
                          />
                          {t('moreColors')}
                        </label>
                      </div>
                    )}
                  </div>
                  <button
                    className={`rb-big${pendingSign ? ' active' : ''}`}
                    disabled={readOnly}
                    title={t('signTitle')}
                    onClick={() => (pendingSign ? setPendingSign(null) : setSignDlg(true))}
                  >
                    <span className="rb-big-icon">
                      <IconSign />
                    </span>
                    {t('sign')}
                  </button>
                </div>
              </div>
              <div className="ribbon-spacer" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className={`rb-big${isFullscreen ? ' active' : ''}`}
                    title={t('fullscreen')}
                    onClick={toggleFullscreen}
                  >
                    <span className="rb-big-icon">
                      <IconFullscreen />
                    </span>
                    {t('fullscreen')}
                  </button>
                  <button className="rb-big" title={`${t('print')} (⌘P)`} onClick={printDoc}>
                    <span className="rb-big-icon">
                      <IconPrint />
                    </span>
                    {t('print')}
                  </button>
                  <button
                    className="rb-big"
                    title={t('propsTitle')}
                    onClick={() => setPropsDlg(true)}
                  >
                    <span className="rb-big-icon">
                      <IconProps />
                    </span>
                    {t('props')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={exporting}
                    onClick={() => void openA11yReport()}
                    title={t('a11yTitle')}
                  >
                    <span className="rb-big-icon">
                      <IconA11y />
                    </span>
                    {t('a11yTitle')}
                  </button>
                </div>
              </div>
            </>
          )}
          {ribbonTab === 'comment' && (
            <>
              <div className="ribbon-group" onMouseDown={(e) => e.preventDefault()}>
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    title={t('highlight')}
                    onClick={() => applyMarkup('highlight')}
                  >
                    <span className="rb-big-icon">
                      <IconHighlight />
                    </span>
                    {t('highlight')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    title={t('underline')}
                    onClick={() => applyMarkup('underline')}
                  >
                    <span className="rb-big-icon">
                      <IconUnderline />
                    </span>
                    {t('underline')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    title={t('strikeout')}
                    onClick={() => applyMarkup('strikeout')}
                  >
                    <span className="rb-big-icon">
                      <IconStrike />
                    </span>
                    {t('strikeout')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    title={t('squiggly')}
                    onClick={() => applyMarkup('squiggly')}
                  >
                    <span className="rb-big-icon">
                      <IconSquiggly />
                    </span>
                    {t('squiggly')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  {DRAW_TOOLS.map(({ tool, icon: DrawIcon, key }) => (
                    <button
                      key={tool}
                      className={`rb-big${drawTool === tool ? ' active' : ''}`}
                      disabled={readOnly}
                      title={t(key)}
                      onClick={() => setDrawTool((v) => (v === tool ? null : tool))}
                    >
                      <span className="rb-big-icon">
                        <DrawIcon />
                      </span>
                      {t(key)}
                    </button>
                  ))}
                  <div className="rb-drop-wrap">
                    <button
                      className={`rb-big${colorOpen ? ' active' : ''}`}
                      disabled={readOnly}
                      title={t('drawColor')}
                      onClick={() => setColorOpen((v) => !v)}
                    >
                      <span className="rb-big-icon">
                        <span className="rb-big-icon-colored">
                          <IconDrawColor />
                          <span
                            className="rb-color-bar"
                            style={{ background: cssRgb(drawColor) }}
                          />
                        </span>
                        <RbCaret />
                      </span>
                      {t('drawColor')}
                    </button>
                    {colorOpen && (
                      <div className="rb-drop rb-color-grid">
                        {DRAW_COLORS.map((c) => (
                          <button
                            key={c.name}
                            className={`rb-swatch${cssRgb(drawColor) === cssRgb(c.rgb) ? ' active' : ''}`}
                            style={{ background: cssRgb(c.rgb) }}
                            title={c.name}
                            onClick={() => {
                              setDrawColor(c.rgb)
                              setColorOpen(false)
                            }}
                          />
                        ))}
                        <label className="rb-color-more" title={t('drawColor')}>
                          <input
                            type="color"
                            value={rgbToHex(drawColor)}
                            onChange={(e) => setDrawColor(hexToRgb(e.target.value))}
                          />
                          {t('moreColors')}
                        </label>
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className={`rb-big${pendingSign ? ' active' : ''}`}
                    disabled={readOnly}
                    title={t('signTitle')}
                    onClick={() => (pendingSign ? setPendingSign(null) : setSignDlg(true))}
                  >
                    <span className="rb-big-icon">
                      <IconSign />
                    </span>
                    {t('sign')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    title={t('stampTitle')}
                    onClick={() => setStampDlg(true)}
                  >
                    <span className="rb-big-icon">
                      <IconWatermark />
                    </span>
                    {t('watermark')}
                  </button>
                </div>
              </div>
            </>
          )}
          {ribbonTab === 'organize' && (
            <>
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    disabled={curOrigIdx < 0 || readOnly}
                    onClick={() => rotatePage(curOrigIdx, -90)}
                    title={t('rotateLeft')}
                  >
                    <span className="rb-big-icon">
                      <IconRotateL />
                    </span>
                    {t('rotateLeft')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={curOrigIdx < 0 || readOnly}
                    onClick={() => rotatePage(curOrigIdx, 90)}
                    title={t('rotateRight')}
                  >
                    <span className="rb-big-icon">
                      <IconRotateR />
                    </span>
                    {t('rotateRight')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    onClick={() => void insertPdf(curOrigIdx)}
                    title={t('insertPdf')}
                  >
                    <span className="rb-big-icon">
                      <IconInsertPdf />
                    </span>
                    {t('insertPdf')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={curOrigIdx < 0 || readOnly}
                    onClick={openExtractDlg}
                    title={t('extractPage')}
                  >
                    <span className="rb-big-icon">
                      <IconExtract />
                    </span>
                    {t('extractPage')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={exporting || pageCount < 2}
                    onClick={() => {
                      setSplitInvalid(false)
                      setSplitDlg(true)
                    }}
                    title={t('splitPdf')}
                  >
                    <span className="rb-big-icon">
                      <IconSplit />
                    </span>
                    {t('splitPdf')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={exporting}
                    onClick={() => void doCombine()}
                    title={t('combineFiles')}
                  >
                    <span className="rb-big-icon">
                      <IconCombine />
                    </span>
                    {t('combineFiles')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={curOrigIdx < 0 || pageCount <= 1 || readOnly}
                    onClick={() => deletePage(curOrigIdx)}
                    title={t('deletePage')}
                  >
                    <span className="rb-big-icon">
                      <IconDeletePage />
                    </span>
                    {t('deletePage')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    title={t('exportImagesAll')}
                    disabled={exporting}
                    onClick={() => void exportImages(true)}
                  >
                    <span className="rb-big-icon">
                      <IconExportImg />
                    </span>
                    {exporting ? t('exporting') : t('exportImages')}
                  </button>
                  <button
                    className="rb-big"
                    title={t('exportText')}
                    disabled={exporting}
                    onClick={() => void exportText()}
                  >
                    <span className="rb-big-icon">
                      <IconExportText />
                    </span>
                    {t('exportText')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={exporting || readOnly}
                    onClick={() => {
                      setSanitizeReport(null)
                      setSanitizePreview(null)
                      setSanitizeDlg(true)
                      void previewSanitize()
                    }}
                    title={t('sanitizeTitle')}
                  >
                    <span className="rb-big-icon">
                      <IconSanitize />
                    </span>
                    {t('sanitizeTitle')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={exporting}
                    onClick={() => void doExportEmbeddedImages()}
                    title={t('exportEmbeddedImages')}
                  >
                    <span className="rb-big-icon">
                      <IconImgStack />
                    </span>
                    {t('exportEmbeddedImages')}
                  </button>
                </div>
              </div>
            </>
          )}
          {ribbonTab === 'view' && (
            <>
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className={`rb-big${spread === 1 ? ' active' : ''}`}
                    title={t('singlePage')}
                    onClick={() => setSpread(1)}
                  >
                    <span className="rb-big-icon">
                      <IconSinglePage />
                    </span>
                    {t('singlePage')}
                  </button>
                  <button
                    className={`rb-big${spread === 2 ? ' active' : ''}`}
                    title={t('twoPage')}
                    onClick={() => setSpread(2)}
                  >
                    <span className="rb-big-icon">
                      <IconSpread />
                    </span>
                    {t('twoPage')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    title={t('actualSize')}
                    onClick={() => {
                      fitModeRef.current = null
                      setScale(1)
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconActualSize />
                    </span>
                    {t('actualSize')}
                  </button>
                  <button
                    className="rb-big"
                    title={t('fitWidth')}
                    onClick={() => {
                      fitModeRef.current = 'width'
                      recomputeFit()
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconFitWidth />
                    </span>
                    {t('fitWidth')}
                  </button>
                  <button
                    className="rb-big"
                    title={t('fitPage')}
                    onClick={() => {
                      fitModeRef.current = 'page'
                      recomputeFit()
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconFitPage />
                    </span>
                    {t('fitPage')}
                  </button>
                  <button className="rb-big" title={t('zoomIn')} onClick={zoomIn}>
                    <span className="rb-big-icon">
                      <IconZoomIn />
                    </span>
                    {t('zoomIn')}
                  </button>
                  <button className="rb-big" title={t('zoomOut')} onClick={zoomOut}>
                    <span className="rb-big-icon">
                      <IconZoomOut />
                    </span>
                    {t('zoomOut')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className={`rb-big${sidebar === 'thumbs' ? ' active' : ''}`}
                    onClick={() => setSidebar((v) => (v === 'thumbs' ? null : 'thumbs'))}
                    title={t('thumbs')}
                  >
                    <span className="rb-big-icon">
                      <IconThumbs />
                    </span>
                    {t('thumbs')}
                  </button>
                  <button
                    className={`rb-big${sidebar === 'outline' ? ' active' : ''}`}
                    disabled={!outline}
                    onClick={() => setSidebar((v) => (v === 'outline' ? null : 'outline'))}
                    title={t('outline')}
                  >
                    <span className="rb-big-icon">
                      <IconOutline />
                    </span>
                    {t('outline')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className={`rb-big${isFullscreen ? ' active' : ''}`}
                    title={t('fullscreen')}
                    onClick={toggleFullscreen}
                  >
                    <span className="rb-big-icon">
                      <IconFullscreen />
                    </span>
                    {t('fullscreen')}
                  </button>
                </div>
              </div>
            </>
          )}
          {ribbonTab === 'ai' && (
            <>
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className={`rb-big ai-entry${aiCollapsed ? '' : ' active'}`}
                    title={t('aiOpenAssistant')}
                    onClick={() => setAiCollapsed((v) => !v)}
                  >
                    <span className="rb-big-icon">
                      <AssistantMark size={26} />
                    </span>
                    <span>ARA Assistant</span>
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    title={t('aiQuickSummaryPrompt')}
                    onClick={() => triggerAiPreset(t('aiQuickSummaryPrompt'))}
                  >
                    <span className="rb-big-icon">
                      <IconSummarize />
                    </span>
                    {t('aiQuickSummary')}
                  </button>
                  <button
                    className="rb-big"
                    title={t('aiQuickKeyPointsPrompt')}
                    onClick={() => triggerAiPreset(t('aiQuickKeyPointsPrompt'))}
                  >
                    <span className="rb-big-icon">
                      <IconKeyPoints />
                    </span>
                    {t('aiQuickKeyPoints')}
                  </button>
                  <button
                    className="rb-big"
                    title={t('aiExplainPrompt')}
                    onClick={() => triggerAiPreset(t('aiExplainPrompt'))}
                  >
                    <span className="rb-big-icon">
                      <IconExplain />
                    </span>
                    {t('aiExplain')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    title={t('aiTranslatePrompt')}
                    onClick={() => triggerAiPreset(t('aiTranslatePrompt'))}
                  >
                    <span className="rb-big-icon">
                      <IconTranslate />
                    </span>
                    {t('aiTranslate')}
                  </button>
                </div>
              </div>
            </>
          )}
        </div>
      </div>
      <div className="app-main">
        {/* dock wrapper animates the width between panel and rail (docs-style 180ms ease);
            the panel stays mounted while collapsed so the chat history survives */}
        <div className={`ai-dock${aiCollapsed ? ' collapsed' : ''}`}>
          {aiCollapsed && (
            <button
              className="ai-rail"
              title={t('aiOpenAssistant')}
              onClick={() => setAiCollapsed(false)}
            >
              <span className="ai-rail-icon" aria-hidden="true">
                <AraSpark size={21} />
              </span>
            </button>
          )}
          <AiPanel
            key={aiPanelKey}
            api={aiApi}
            onCollapse={() => setAiCollapsed(true)}
            preset={aiPreset}
            onClearPreset={() => setAiPreset(null)}
          />
        </div>
        <div className="app-content">
          <div className="pdf-body">
            {sidebar === 'outline' && outline && (
              <div className="pdf-thumbs pdf-outline-pane" style={{ width: sidebarW }}>
                <div className="pdf-outline-toolbar">
                  {!outlineEditMode && (
                    <button className="pdf-outline-tool" onClick={() => void enterOutlineEdit()}>
                      {t('outlineEdit')}
                    </button>
                  )}
                  {outlineEditMode && (
                    <>
                      <button
                        className="pdf-outline-tool"
                        onClick={() => openOutlinePrompt('add', [])}
                      >
                        {t('outlineAdd')}
                      </button>
                      <button className="pdf-outline-tool" onClick={() => void doApplyOutline()}>
                        {t('outlineApply')}
                      </button>
                      <button className="pdf-outline-tool" onClick={() => setOutlineDraft(null)}>
                        {t('outlineDiscard')}
                      </button>
                    </>
                  )}
                </div>
                <OutlinePanel
                  outline={outline}
                  onGoToDest={(dest) => void goToDest(dest)}
                  draft={outlineEditMode ? (outlineDraft?.nodes ?? []) : undefined}
                  onEdit={handleOutlineEdit}
                />
              </div>
            )}
            {sidebar === 'comments' && (
              <div className="pdf-thumbs pdf-outline-pane" style={{ width: sidebarW }}>
                <CommentsPanel
                  rows={[
                    ...markups.map((m): CommentRow => ({
                      id: m.id,
                      pageIndex: m.pageIndex,
                      match: {
                        pageIndex: m.pageIndex,
                        subtype: m.type,
                        rect: [],
                        contents: '',
                        name: '',
                      },
                      excerpt: t(m.type),
                      label: t(m.type),
                      saved: false,
                    })),
                    ...commentRows,
                  ]}
                  onJump={jumpComment}
                  onRemove={removeComment}
                />
              </div>
            )}
            {sidebar === 'thumbs' && (
              <div ref={thumbsRef} className="pdf-thumbs" style={{ width: sidebarW }}>
                {visList.map((origIdx, v) => {
                  const size = dispSize(origIdx)
                  return (
                    <div
                      key={origIdx}
                      ref={setThumbRef(v)}
                      data-idx={v}
                      tabIndex={-1}
                      className={`pdf-thumb${currentPage === v + 1 ? ' pdf-thumb-active' : ''}${
                        dragOver === v && dragFrom !== null && dragFrom !== v
                          ? ' pdf-thumb-dropbefore'
                          : ''
                      }`}
                      draggable={!readOnly}
                      onDragStart={(e) => {
                        e.dataTransfer.effectAllowed = 'move'
                        setDragFrom(v)
                      }}
                      onDragOver={(e) => {
                        if (dragFrom === null) return
                        e.preventDefault()
                        e.dataTransfer.dropEffect = 'move'
                        setDragOver(v)
                      }}
                      onDragLeave={() => setDragOver((o) => (o === v ? null : o))}
                      onDrop={(e) => {
                        e.preventDefault()
                        if (dragFrom !== null) movePage(dragFrom, v)
                        setDragFrom(null)
                        setDragOver(null)
                      }}
                      onDragEnd={() => {
                        setDragFrom(null)
                        setDragOver(null)
                      }}
                      onClick={() => scrollToPage(v + 1)}
                      onContextMenu={(e) => {
                        e.preventDefault()
                        setThumbMenu({
                          x: Math.min(e.clientX, window.innerWidth - 190),
                          y: Math.min(e.clientY, window.innerHeight - 190),
                          origIdx,
                        })
                      }}
                    >
                      <div
                        className="pdf-thumb-box"
                        style={{ aspectRatio: `${size.width} / ${size.height}` }}
                      >
                        <PdfThumb
                          doc={doc}
                          pageNo={origIdx + 1}
                          rotationDelta={rotDelta(origIdx)}
                          visible={visibleThumbs.has(v)}
                          rasterW={thumbRasterW}
                        />
                      </div>
                      <span className="pdf-thumb-no">{v + 1}</span>
                    </div>
                  )
                })}
              </div>
            )}
            {(sidebar === 'thumbs' || (sidebar === 'outline' && !!outline)) && (
              <div className="pdf-side-resizer" onPointerDown={startSidebarResize} />
            )}
            <div
              ref={scrollRef}
              className={`pdf-scroll${drawTool ? ' pdf-drawing' : ''}`}
              onScroll={() => {
                handleScroll()
                setSelPopup(null)
                setSelected(null)
              }}
              onMouseUp={drawTool ? undefined : handleMouseUp}
              onClick={(e) => {
                // Clicking anywhere that isn't an annotation clears the selection
                if (
                  !(e.target as Element).closest?.(
                    '.pdf-markup, .pdf-draw-shape, .pdf-note-pin, .pdf-stamp-preview',
                  )
                )
                  setSelected(null)
              }}
            >
              {rows.map((row, r) => (
                <div key={r} ref={setRowRef(r)} data-idx={r} className="pdf-row">
                  {row.map((origIdx) => {
                    const rowVisible = visibleRows.has(r)
                    const size = dispSize(origIdx)
                    const geom = pageGeom(origIdx)
                    return (
                      <div
                        key={origIdx}
                        className="pdf-page"
                        style={
                          {
                            width: Math.floor(size.width * scale),
                            height: Math.floor(size.height * scale),
                            '--scale-factor': scale,
                          } as CSSProperties
                        }
                      >
                        <PdfPage
                          doc={doc}
                          pageNo={origIdx + 1}
                          scale={scale}
                          rotationDelta={rotDelta(origIdx)}
                          visible={rowVisible}
                        />
                        {pendingSign && (
                          <SignDropOverlay
                            sig={pendingSign}
                            dispW={geomDispSize(geom).width}
                            dispH={geomDispSize(geom).height}
                            scale={scale}
                            color={drawColor}
                            title={t('signHint')}
                            onPlace={(vx, vy) => placeSignature(origIdx, vx, vy)}
                          />
                        )}
                        {rowVisible && (
                          <>
                            {/* Preview of unsaved stamps; clicking selects the whole watermark/header-footer set */}
                            {(stampPreview.get(origIdx) ?? []).map((s, si) => (
                              <img
                                key={si}
                                className={`pdf-stamp-preview${selected?.kind === 'stamp' ? ' pdf-stamp-selected' : ''}`}
                                src={`data:image/png;base64,${s.image}`}
                                alt=""
                                title={t('removeStamp')}
                                style={{
                                  ...pdfRectToCss(geom, s.rect, scale),
                                  opacity: s.opacity ?? 1,
                                }}
                                onClick={(e) =>
                                  setSelected({ kind: 'stamp', ...popupPos(e.clientX, e.clientY) })
                                }
                              />
                            ))}
                            {searchOpen && (
                              <div className="pdf-search-layer">
                                {activeMatches.flatMap((m, mi) =>
                                  m.pageIndex === origIdx
                                    ? m.rects.map((r, ri) => (
                                        <div
                                          key={`${mi}-${ri}`}
                                          className={`pdf-search-hit${mi === searchCurClamped ? ' pdf-search-hit-cur' : ''}`}
                                          style={pdfRectToCss(geom, r, scale)}
                                        />
                                      ))
                                    : [],
                                )}
                              </div>
                            )}
                            <MarkupOverlay
                              markups={markups.filter((m) => m.pageIndex === origIdx)}
                              geom={geom}
                              scale={scale}
                              selectedId={selected?.kind === 'markup' ? selected.id : null}
                              selectTitle={t('removeMarkup')}
                              onSelect={(id, x, y) =>
                                setSelected({ kind: 'markup', id, ...popupPos(x, y) })
                              }
                            />
                            <DrawLayer
                              geom={geom}
                              scale={scale}
                              pageWidth={size.width}
                              pageHeight={size.height}
                              drawings={drawings.filter((d) => d.input.pageIndex === origIdx)}
                              tool={readOnly ? null : drawTool}
                              color={drawColor}
                              strokeWidth={STROKE_WIDTH}
                              selectedId={selected?.kind === 'drawing' ? selected.id : null}
                              selectTitle={t('removeMarkup')}
                              onCommit={(input) => commitDrawing(origIdx, input)}
                              onPlaceAt={(kind, at) => {
                                setNoteText('')
                                setNotePrompt({ kind, origIdx, at })
                              }}
                              onSelect={(id, x, y) =>
                                setSelected({ kind: 'drawing', id, ...popupPos(x, y) })
                              }
                              onMove={readOnly ? undefined : moveDrawing}
                              onResize={readOnly ? undefined : resizeDrawing}
                            />
                            <LinkLayer
                              doc={doc}
                              pageNo={origIdx + 1}
                              geom={geom}
                              scale={scale}
                              onGoToDest={(dest) => void goToDest(dest)}
                            />
                            <FormLayer
                              doc={doc}
                              pageNo={origIdx + 1}
                              geom={geom}
                              scale={scale}
                              readOnly={readOnly}
                              edits={formEdits}
                              onEdit={(v2) => {
                                pushUndo(`form:${v2.name}`)
                                setFormEdits((prev) => new Map(prev).set(v2.name, v2))
                              }}
                            />
                          </>
                        )}
                      </div>
                    )
                  })}
                </div>
              ))}
            </div>
            {searchOpen && (
              <div className="pdf-search-bar">
                <input
                  ref={searchInputRef}
                  className="pdf-search-input"
                  aria-label={t('search')}
                  placeholder={t('search')}
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') searchStep(e.shiftKey ? -1 : 1)
                    else if (e.key === 'Escape') closeSearch()
                  }}
                />
                <span className="pdf-search-count">
                  {searchQuery.trim()
                    ? activeMatches.length > 0
                      ? t('searchCount', {
                          current: searchCurClamped + 1,
                          total: activeMatches.length,
                        })
                      : t('searchNoResults')
                    : ''}
                </span>
                <button
                  className="rb-icon"
                  title={t('searchPrev')}
                  disabled={activeMatches.length === 0}
                  onClick={() => searchStep(-1)}
                >
                  ‹
                </button>
                <button
                  className="rb-icon"
                  title={t('searchNext')}
                  disabled={activeMatches.length === 0}
                  onClick={() => searchStep(1)}
                >
                  ›
                </button>
                <button className="rb-icon" onClick={closeSearch}>
                  ×
                </button>
              </div>
            )}
            {selPopup && (
              <div
                className="pdf-sel-popup"
                style={{ left: selPopup.x, top: selPopup.y }}
                onMouseDown={(e) => e.preventDefault()}
              >
                <button
                  type="button"
                  title={t('highlight')}
                  onClick={() => applyMarkup('highlight')}
                >
                  <span className="sel-swatch sel-swatch-hl" />
                </button>
                <button
                  type="button"
                  title={t('underline')}
                  onClick={() => applyMarkup('underline')}
                >
                  <span className="sel-swatch sel-swatch-ul">U</span>
                </button>
                <button
                  type="button"
                  title={t('strikeout')}
                  onClick={() => applyMarkup('strikeout')}
                >
                  <span className="sel-swatch sel-swatch-st">S</span>
                </button>
                <button type="button" title={t('squiggly')} onClick={() => applyMarkup('squiggly')}>
                  <span className="sel-swatch sel-swatch-sq">~</span>
                </button>
              </div>
            )}
            {selected && (
              <div
                className="pdf-del-popup"
                style={{ left: selected.x, top: selected.y }}
                onMouseDown={(e) => e.preventDefault()}
              >
                <button type="button" onClick={deleteSelected}>
                  {t('deleteAnnotation')}
                </button>
              </div>
            )}
            {deleteToast && (
              <div className="pdf-toast">
                <span>{t('annotationDeleted')}</span>
                <button
                  type="button"
                  onClick={() => {
                    setDeleteToast(false)
                    undo()
                  }}
                >
                  {t('undo')}
                </button>
              </div>
            )}
            {thumbMenu && (
              <div className="thumb-menu file-menu" style={{ left: thumbMenu.x, top: thumbMenu.y }}>
                <button
                  onClick={() => {
                    rotatePage(menuOrig, -90)
                    setThumbMenu(null)
                  }}
                >
                  {t('rotateLeft')}
                </button>
                <button
                  onClick={() => {
                    rotatePage(menuOrig, 90)
                    setThumbMenu(null)
                  }}
                >
                  {t('rotateRight')}
                </button>
                <button
                  disabled={pageCount <= 1}
                  onClick={() => {
                    deletePage(menuOrig)
                    setThumbMenu(null)
                  }}
                >
                  {t('deletePage')}
                </button>
                <button
                  onClick={() => {
                    setThumbMenu(null)
                    void extractPage(menuOrig)
                  }}
                >
                  {t('extractPage')}
                </button>
                <button
                  onClick={() => {
                    setThumbMenu(null)
                    void insertPdf(menuOrig)
                  }}
                >
                  {t('insertPdf')}
                </button>
              </div>
            )}
            {stampDlg && (
              <StampDialog t={t} onCancel={() => setStampDlg(false)} onApply={applyStamps} />
            )}
            {propsDlg && (
              <PropertiesDialog
                doc={doc}
                fileName={fileName}
                fileSize={fileSize}
                pageCount={pageCount}
                pending={metadata}
                readOnly={readOnly}
                t={t}
                onCancel={() => setPropsDlg(false)}
                onApply={(meta) => {
                  setPropsDlg(false)
                  pushUndo()
                  setMetadata(meta)
                }}
              />
            )}
            {signDlg && (
              <SignatureDialog
                color={drawColor}
                t={t}
                onCancel={() => setSignDlg(false)}
                onConfirm={(sig) => {
                  setSignDlg(false)
                  setPendingSign(sig)
                }}
              />
            )}
            {notePrompt && (
              <div className="pdf-modal-mask" onClick={() => setNotePrompt(null)}>
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">
                    {notePrompt.kind === 'freetext' ? t('drawFreeText') : t('noteTitle')}
                  </div>
                  <textarea
                    className="pdf-modal-textarea"
                    value={noteText}
                    placeholder={t('notePlaceholder')}
                    autoFocus
                    onChange={(e) => setNoteText(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) confirmNote()
                      else if (e.key === 'Escape') setNotePrompt(null)
                    }}
                  />
                  <div className="pdf-modal-actions">
                    <button className="pdf-modal-btn" onClick={() => setNotePrompt(null)}>
                      {t('cancel')}
                    </button>
                    <button
                      className="pdf-modal-btn primary"
                      disabled={!noteText.trim()}
                      onClick={confirmNote}
                    >
                      {t('ok')}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {outlinePrompt && (
              <div className="pdf-modal-mask" onClick={() => setOutlinePrompt(null)}>
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">
                    {t(
                      outlinePrompt.mode === 'rename'
                        ? 'outlineRename'
                        : outlinePrompt.mode === 'add'
                          ? 'outlineAdd'
                          : 'outlineAddChild',
                    )}
                  </div>
                  <input
                    className="pdf-modal-input"
                    value={outlinePrompt.text}
                    autoFocus
                    onChange={(e) => setOutlinePrompt({ ...outlinePrompt, text: e.target.value })}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') confirmOutlinePrompt()
                      else if (e.key === 'Escape') setOutlinePrompt(null)
                    }}
                  />
                  <div className="pdf-modal-actions">
                    <button className="pdf-modal-btn" onClick={() => setOutlinePrompt(null)}>
                      {t('cancel')}
                    </button>
                    <button
                      className="pdf-modal-btn primary"
                      disabled={!outlinePrompt.text.trim()}
                      onClick={confirmOutlinePrompt}
                    >
                      {t('ok')}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {a11yDlg && a11yReport && (
              <div className="pdf-modal-mask" onClick={() => setA11yDlg(false)}>
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">{t('a11yTitle')}</div>
                  <div className="pdf-modal-hint">
                    {t('totalPagesN', { n: a11yReport.pageCount })}
                  </div>
                  <div className="pdf-a11y-rows">
                    {a11yReport.rules.map((r) => {
                      const row = A11Y_ROWS.find((x) => x.rule === r.rule)
                      const glyph = r.verdict === 'pass' ? '✓' : r.verdict === 'fail' ? '✕' : '~'
                      return (
                        <div
                          key={r.rule}
                          className={`pdf-a11y-row pdf-a11y-${r.verdict}`}
                          title={r.verdict === 'manual' ? t('a11yManual') : undefined}
                        >
                          <span className="pdf-a11y-verdict">{glyph}</span>
                          <span className="pdf-a11y-label">{t(row!.label)}</span>
                          {r.count !== undefined && r.count > 0 && (
                            <span className="pdf-a11y-count">{r.count}</span>
                          )}
                        </div>
                      )
                    })}
                  </div>
                  {a11yMetaFail(a11yReport.rules) && (
                    <div className="pdf-modal-hint">
                      {t('a11yFixHint', { lang: htmlLang(uiLang) })}
                    </div>
                  )}
                  <div className="pdf-modal-actions">
                    {a11yMetaFail(a11yReport.rules) && (
                      <button className="pdf-modal-btn" onClick={() => void doA11yFix()}>
                        {t('a11yFix')}
                      </button>
                    )}
                    <button className="pdf-modal-btn primary" onClick={() => setA11yDlg(false)}>
                      {t('ok')}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {splitDlg && (
              <div className="pdf-modal-mask" onClick={() => setSplitDlg(false)}>
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">{t('splitPdf')}</div>
                  <div className="pdf-modal-choices">
                    {(['every', 'at', 'bookmark'] as const).map((m) => (
                      <label key={m} className="pdf-modal-choice">
                        <input
                          type="radio"
                          name="split-mode"
                          checked={splitMode === m}
                          onChange={() => {
                            setSplitMode(m)
                            setSplitInvalid(false)
                          }}
                        />
                        {t(
                          m === 'every'
                            ? 'splitEveryN'
                            : m === 'at'
                              ? 'splitAtPages'
                              : 'splitAtBookmarks',
                        )}
                      </label>
                    ))}
                  </div>
                  {splitMode === 'every' && (
                    <input
                      className={`pdf-modal-input${splitInvalid ? ' invalid' : ''}`}
                      value={splitEvery}
                      inputMode="numeric"
                      autoFocus
                      onChange={(e) => {
                        setSplitEvery(e.target.value)
                        setSplitInvalid(false)
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') void doSplit()
                        else if (e.key === 'Escape') setSplitDlg(false)
                      }}
                    />
                  )}
                  {splitMode === 'at' && (
                    <input
                      className={`pdf-modal-input${splitInvalid ? ' invalid' : ''}`}
                      value={splitPages}
                      placeholder={t('extractRangeHint', { total: pageCount })}
                      autoFocus
                      onChange={(e) => {
                        setSplitPages(e.target.value)
                        setSplitInvalid(false)
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') void doSplit()
                        else if (e.key === 'Escape') setSplitDlg(false)
                      }}
                    />
                  )}
                  {splitMode === 'bookmark' && (
                    <div className="pdf-modal-hint">{t('splitAtBookmarks')}</div>
                  )}
                  <div className="pdf-modal-actions">
                    <button className="pdf-modal-btn" onClick={() => setSplitDlg(false)}>
                      {t('cancel')}
                    </button>
                    <button className="pdf-modal-btn primary" onClick={() => void doSplit()}>
                      {t('ok')}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {sanitizeDlg && (
              <div className="pdf-modal-mask" onClick={() => setSanitizeDlg(false)}>
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">{t('sanitizeTitle')}</div>
                  {sanitizeReport ? (
                    <>
                      {/* The report phase: numbers are language-neutral, so the counts reuse the
                          same category labels and need no extra strings */}
                      {SANITIZE_ROWS.map((row) =>
                        sanitizeReport[row.key] > 0 ? (
                          <div key={row.key} className="pdf-modal-choice">
                            {t(row.label)} · {sanitizeReport[row.key]}
                          </div>
                        ) : null,
                      )}
                      {SANITIZE_ROWS.every((row) => sanitizeReport[row.key] === 0) && (
                        <div className="pdf-modal-hint">0</div>
                      )}
                      <div className="pdf-modal-actions">
                        <button
                          className="pdf-modal-btn primary"
                          onClick={() => setSanitizeDlg(false)}
                        >
                          {t('ok')}
                        </button>
                      </div>
                    </>
                  ) : (
                    <>
                      <div className="pdf-modal-choices">
                        {SANITIZE_ROWS.map((row) => (
                          <label key={row.key} className="pdf-modal-choice">
                            <input
                              type="checkbox"
                              checked={sanitizeCfg[row.key]}
                              onChange={(e) => {
                                const next = { ...sanitizeCfg }
                                next[row.key] = e.target.checked
                                setSanitizeCfg(next)
                              }}
                            />
                            {t(row.label)}
                            <span className="pdf-sanitize-count" title={t('sanitizeTitle')}>
                              {sanitizePreview?.[row.key] || ''}
                            </span>
                          </label>
                        ))}
                      </div>
                      <div className="pdf-modal-actions">
                        <button className="pdf-modal-btn" onClick={() => setSanitizeDlg(false)}>
                          {t('cancel')}
                        </button>
                        <button className="pdf-modal-btn primary" onClick={() => void doSanitize()}>
                          {t('ok')}
                        </button>
                      </div>
                    </>
                  )}
                </div>
              </div>
            )}
            {extractDlg && (
              <div className="pdf-modal-mask" onClick={() => setExtractDlg(false)}>
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">{t('extractRangeTitle')}</div>
                  <input
                    className={`pdf-modal-input${extractInvalid ? ' invalid' : ''}`}
                    value={extractInput}
                    placeholder={t('extractRangeHint', { total: pageCount })}
                    autoFocus
                    onChange={(e) => {
                      setExtractInput(e.target.value)
                      setExtractInvalid(false)
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') confirmExtract()
                      else if (e.key === 'Escape') setExtractDlg(false)
                    }}
                  />
                  <div className="pdf-modal-actions">
                    <button className="pdf-modal-btn" onClick={() => setExtractDlg(false)}>
                      {t('cancel')}
                    </button>
                    <button className="pdf-modal-btn primary" onClick={confirmExtract}>
                      {t('ok')}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {pendingOpen && (
              <div
                className="pdf-modal-mask"
                role="dialog"
                aria-modal="true"
                aria-label={t('openDirtyTitle')}
              >
                <div className="pdf-modal">
                  <div className="pdf-modal-title">{t('openDirtyTitle')}</div>
                  <div className="pdf-modal-hint">{t('openDirtyDetail')}</div>
                  <div className="pdf-modal-actions">
                    <button className="pdf-modal-btn" onClick={() => void confirmOpen('discard')}>
                      {t('discard')}
                    </button>
                    <button className="pdf-modal-btn" onClick={() => setPendingOpen(null)}>
                      {t('cancel')}
                    </button>
                    <button
                      className="pdf-modal-btn primary"
                      disabled={saveState === 'saving'}
                      onClick={() => void confirmOpen('save')}
                    >
                      {t('save')}
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>

          <footer className="status-bar">
            <div className="status-left">
              <span className="status-item">
                {t('appPageOf', { current: currentPage, total: pageCount })}
              </span>
            </div>
            <div className="status-right">
              <button className="zoom-btn" title={t('zoomOut')} onClick={zoomOut}>
                −
              </button>
              <input
                className="zoom-slider"
                type="range"
                min={MIN_SCALE * 100}
                max={MAX_SCALE * 100}
                step={5}
                value={Math.round(scale * 100)}
                onChange={(e) => applyScale(Number(e.target.value) / 100, null)}
              />
              <button className="zoom-btn" title={t('zoomIn')} onClick={zoomIn}>
                +
              </button>
              <span className="zoom-value">{Math.round(scale * 100)}%</span>
            </div>
          </footer>
        </div>
      </div>
      {showPrintPreview && doc && (
        <PdfPrintPreview
          doc={doc}
          visList={visList}
          rotDelta={rotDelta}
          dispSize={dispSize}
          onClose={() => setShowPrintPreview(false)}
        />
      )}
    </div>
  )
}
