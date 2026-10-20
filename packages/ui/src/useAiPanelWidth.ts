import { useCallback, useEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent } from 'react'

const PANEL_WIDTH_MAX = 720

/** Clamp the resizable AI panel width so neither pane collapses (max is 60% of the window) */
export function clampAiPanelWidth(w: number, min = 280): number {
  return Math.min(Math.max(w, min), Math.min(PANEL_WIDTH_MAX, Math.round(window.innerWidth * 0.6)))
}

export function loadAiPanelWidth(key: string, defaultValue = 360, min = 280): number {
  const saved = Number(localStorage.getItem(key))
  return Number.isFinite(saved) && saved > 0 ? clampAiPanelWidth(saved, min) : defaultValue
}

export interface AiPanelWidth {
  readonly panelWidth: number
  readonly setPanelWidth: React.Dispatch<React.SetStateAction<number>>
  readonly asideRef: React.RefObject<HTMLElement | null>
  readonly resizing: boolean
  readonly startResize: (e: ReactPointerEvent<HTMLDivElement>) => void
}

/**
 * Resizable AI panel width: persisted, clamped, drives the dock/grid column via
 * `--ai-panel-width`. Canonical copy; was duplicated ~60 lines x5 across
 * docs/slides/pdf/markdown/diagram AiPanel.tsx files.
 * Dock sync stays with the caller (it owns the `open` dep).
 */
export function useAiPanelWidth(key: string, min = 280, defaultValue = 360): AiPanelWidth {
  const [panelWidth, setPanelWidth] = useState(() => loadAiPanelWidth(key, defaultValue, min))
  const asideRef = useRef<HTMLElement | null>(null)
  const [resizing, setResizing] = useState(false)
  const cleanupRef = useRef<(() => void) | null>(null)
  useEffect(() => () => cleanupRef.current?.(), [])

  useEffect(() => {
    const onResize = (): void => setPanelWidth((w) => clampAiPanelWidth(w, min))
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [min])

  const startResize = useCallback(
    (e: ReactPointerEvent<HTMLDivElement>) => {
      e.preventDefault()
      const resizer = e.currentTarget
      setResizing(true)
      document.body.style.cursor = 'col-resize'
      document.body.style.userSelect = 'none'
      const onMove = (ev: PointerEvent): void => {
        setPanelWidth(clampAiPanelWidth(window.innerWidth - ev.clientX, min))
      }
      let done = false
      const cleanup = (): void => {
        if (done) return
        done = true
        cleanupRef.current = null
        window.removeEventListener('pointermove', onMove)
        window.removeEventListener('pointerup', cleanup)
        window.removeEventListener('pointercancel', cleanup)
        resizer.removeEventListener('lostpointercapture', cleanup)
        document.body.style.cursor = ''
        document.body.style.userSelect = ''
        setResizing(false)
        setPanelWidth((w) => {
          localStorage.setItem(key, String(Math.round(w)))
          return w
        })
      }
      cleanupRef.current = cleanup
      window.addEventListener('pointermove', onMove)
      window.addEventListener('pointerup', cleanup)
      window.addEventListener('pointercancel', cleanup)
      // lostpointercapture also fires if the resizer is unmounted mid-drag (panel collapse)
      resizer.addEventListener('lostpointercapture', cleanup)
      resizer.setPointerCapture(e.pointerId)
    },
    [key, min],
  )

  return { panelWidth, setPanelWidth, asideRef, resizing, startResize }
}
