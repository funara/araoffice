import { useState, useEffect } from 'react'
import { useI18n } from './locale'

interface FeedbackModalProps {
  onClose: () => void
}

/**
 * Feedback goes straight to the public issue tracker: there is no feedback
 * service, so nothing is collected, stored, or sent to a server.
 */
export function FeedbackModal({ onClose }: FeedbackModalProps) {
  const { t } = useI18n()
  const [issuesUrl, setIssuesUrl] = useState<string | null>(null)

  useEffect(() => {
    void window.aiOffice
      .getSupportLinks()
      .then((links) => setIssuesUrl(links.issuesUrl))
      .catch(() => setIssuesUrl(null))
  }, [])

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onClose])

  return (
    <div className="modal-overlay fb-overlay" onClick={onClose}>
      <div
        className="modal fb-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('feedbackTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Ambient Glow */}
        <div className="fb-ambient-glow" aria-hidden="true" />

        {/* Header */}
        <header className="fb-header">
          <div className="fb-header-left">
            <span className="fb-eyebrow">
              <span className="fb-spark" aria-hidden="true">
                ✦
              </span>
              ARAOFFICE
            </span>
            <h2 className="fb-title">{t('feedbackTitle')}</h2>
          </div>
          <button
            type="button"
            className="fb-close-btn"
            onClick={onClose}
            aria-label={t('cancel')}
            title={t('cancel')}
          >
            <svg
              width="18"
              height="18"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.2"
            >
              <line x1="18" y1="6" x2="6" y2="18" />
              <line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        </header>

        <div className="fb-form">
          <p className="fb-issues-body">{t('feedbackIssuesBody')}</p>
          <div className="fb-footer">
            <button type="button" className="btn btn-secondary" onClick={onClose}>
              {t('cancel')}
            </button>
            {issuesUrl && (
              <button
                type="button"
                className="btn btn-primary fb-submit-btn"
                onClick={() => void window.aiOfficeWindow.openExternal(issuesUrl)}
              >
                {t('feedbackOpenIssues')}
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
