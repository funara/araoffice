import { useEffect, useRef } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from './i18n/locale'
import { buildPrintNode } from './export/printHtml'

interface Props {
  editor?: Editor | null
  onClose: () => void
}

export function MarkdownPrintPreview({ editor, onClose }: Props) {
  const { t } = useI18n()
  const pageRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', handleKey)
    return () => window.removeEventListener('keydown', handleKey)
  }, [onClose])

  useEffect(() => {
    const target =
      editor?.view?.dom ??
      (document.querySelector('.doc-page .doc-editor') as HTMLElement | null) ??
      (document.querySelector('.doc-editor') as HTMLElement | null)
    if (!pageRef.current || !target) return
    const clone = buildPrintNode(target)
    clone.classList.add('doc-editor')
    pageRef.current.replaceChildren(clone)
  }, [editor])

  const handlePrint = () => {
    window.print()
  }

  return (
    <div className="pagination-preview markdown-print-preview">
      <div className="pv-toolbar">
        <span className="pv-title">{t('printPreview')}</span>
        <span className="pv-hint">{t('printPreviewHint')}</span>
        <button type="button" className="pv-close pv-print" onClick={handlePrint}>
          {t('print')}
        </button>
        <button type="button" className="pv-close" onClick={onClose}>
          {t('close')}
        </button>
      </div>
      <div className="pv-scroll">
        <div className="pv-page md-print-sheet" ref={pageRef} />
      </div>
    </div>
  )
}
