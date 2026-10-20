import { useEffect, useRef } from 'react'
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist'
import { useI18n } from './i18n/locale'

interface Props {
  doc: PDFDocumentProxy
  visList: number[]
  rotDelta: (origIdx: number) => number
  dispSize: (origIdx: number) => { width: number; height: number }
  onClose: () => void
}

function PdfPrintPage({
  doc,
  origIdx,
  rotationDelta,
  width,
  height,
}: {
  doc: PDFDocumentProxy
  origIdx: number
  rotationDelta: number
  width: number
  height: number
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null)

  useEffect(() => {
    let cancelled = false
    let renderTask: RenderTask | null = null
    void (async () => {
      const page = await doc.getPage(origIdx + 1)
      if (cancelled) return
      const rotation = (page.rotate + rotationDelta) % 360
      const viewport = page.getViewport({ scale: 96 / 72, rotation })
      const canvas = canvasRef.current
      if (!canvas) return
      const dpr = Math.min(window.devicePixelRatio || 1, 2)
      canvas.width = Math.floor(viewport.width * dpr)
      canvas.height = Math.floor(viewport.height * dpr)
      canvas.style.width = `${Math.floor(viewport.width)}px`
      canvas.style.height = `${Math.floor(viewport.height)}px`
      renderTask = page.render({
        canvas,
        viewport,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
      })
      try {
        await renderTask.promise
      } catch {
        /* cancelled */
      }
    })()
    return () => {
      cancelled = true
      renderTask?.cancel()
    }
  }, [doc, origIdx, rotationDelta, width, height])

  return (
    <div style={{ width: `${width}px`, height: `${height}px`, overflow: 'hidden' }}>
      <canvas ref={canvasRef} style={{ display: 'block' }} />
    </div>
  )
}

export function PdfPrintPreview({ doc, visList, rotDelta, dispSize, onClose }: Props) {
  const { t } = useI18n()

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [onClose])

  const handlePrint = () => {
    window.print()
  }

  return (
    <div className="pagination-preview pdf-print-preview">
      <div className="pv-toolbar">
        <span className="pv-title">{t('printPreview')}</span>
        <span className="pv-count">{t('totalPagesN', { n: visList.length })}</span>
        <span className="pv-hint">{t('printPreviewHint')}</span>
        <button type="button" className="pv-close pv-print" onClick={handlePrint}>
          {t('print')}
        </button>
        <button type="button" className="pv-close" onClick={onClose}>
          {t('cancel')}
        </button>
      </div>
      <div className="pv-scroll">
        {visList.map((origIdx) => {
          const size = dispSize(origIdx)
          const pageW = Math.floor(size.width * (96 / 72))
          const pageH = Math.floor(size.height * (96 / 72))
          return (
            <div
              key={origIdx}
              className="pv-page"
              style={
                {
                  width: `${pageW}px`,
                  height: `${pageH}px`,
                  '--pv-page-h': `${pageH}px`,
                } as React.CSSProperties
              }
            >
              <PdfPrintPage
                doc={doc}
                origIdx={origIdx}
                rotationDelta={rotDelta(origIdx)}
                width={pageW}
                height={pageH}
              />
            </div>
          )
        })}
      </div>
    </div>
  )
}
