import type { ReactElement } from 'react'
import type { CommentRow } from './comments'
import { t } from './i18n/locale'

/**
 * Flat comment list: pending unsaved markups first, then the file's own annotations, each a
 * jump button beside a delete action. Pure presentation — fetching, undo and deletion flows
 * live in App, the same split the outline panel keeps.
 */
export function CommentsPanel({
  rows,
  onJump,
  onRemove,
}: {
  rows: CommentRow[]
  onJump: (row: CommentRow) => void
  /** The row to delete; App routes pending rows state-side and saved rows through the file */
  onRemove: (row: CommentRow) => void
}): ReactElement {
  if (rows.length === 0) return <div className="pdf-outline pdf-comments">{t('commentsZero')}</div>
  return (
    <div className="pdf-outline pdf-comments">
      {rows.map((r, i) => (
        <div key={i} className="pdf-comments-row">
          <button
            className={`pdf-outline-item${r.saved ? '' : ' pdf-comments-pending'}`}
            title={r.excerpt || r.label}
            onClick={() => onJump(r)}
          >
            <span className="pdf-comments-page">{r.pageIndex + 1}</span>
            <span className="pdf-comments-label">{r.label}</span>
            {r.excerpt && <span className="pdf-comments-excerpt">{r.excerpt}</span>}
          </button>
          <button
            className="pdf-outline-tool"
            title={t(r.saved ? 'deleteAnnotation' : 'removeMarkup')}
            onClick={() => onRemove(r)}
          >
            ✕
          </button>
        </div>
      ))}
    </div>
  )
}
