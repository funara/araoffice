import { useEffect } from 'react'
import type { AppUpdateState } from '../../shared/update-api'
import { useI18n } from './locale'

interface UpdateRestartModalProps {
  onClose: () => void
  updateState: AppUpdateState
  onRestart: () => void
  onOpenChangelog?: () => void
}

export function UpdateRestartModal({
  onClose,
  updateState,
  onRestart,
  onOpenChangelog,
}: UpdateRestartModalProps) {
  const { t } = useI18n()
  const version = updateState.availableVersion || updateState.currentVersion || ''

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onClose])

  return (
    <div className="modal-overlay update-modal-overlay" onClick={onClose}>
      <div
        className="modal update-restart-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('updateReadyTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Subtle Ambient Top Glow */}
        <div className="update-ambient-glow" aria-hidden="true" />

        {/* Top Header Controls */}
        <div className="update-modal-top">
          <div
            className={`update-eyebrow-pill ${onOpenChangelog ? 'clickable' : ''}`}
            onClick={
              onOpenChangelog
                ? () => {
                    onClose()
                    onOpenChangelog()
                  }
                : undefined
            }
            role={onOpenChangelog ? 'button' : undefined}
            tabIndex={onOpenChangelog ? 0 : undefined}
          >
            <span className="update-sparkle" aria-hidden="true">
              ✦
            </span>
            <span>Update Ready · v{version}</span>
          </div>
          <button
            type="button"
            className="update-close-btn"
            onClick={onClose}
            aria-label={t('cancel')}
            title={t('cancel')}
          >
            <svg
              width="15"
              height="15"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <line x1="18" y1="6" x2="6" y2="18" />
              <line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        </div>

        {/* Hero Icon */}
        <div className="update-modal-badge">
          <div className="update-badge-icon">
            <svg
              width="26"
              height="26"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.2"
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden="true"
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
              <polyline points="7 10 12 15 17 10" />
              <line x1="12" y1="15" x2="12" y2="3" />
            </svg>
          </div>
        </div>

        {/* Title & Description */}
        <h3 className="update-modal-title">{t('updateReadyTitle')}</h3>
        <p className="update-modal-desc">{t('updateReadyDesc', { v: version })}</p>
        <div className="update-notice-box">
          <div className="update-notice-item">
            <svg
              width="15"
              height="15"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.2"
              strokeLinecap="round"
              strokeLinejoin="round"
              className="notice-icon-check"
              aria-hidden="true"
            >
              <polyline points="20 6 9 17 4 12" />
            </svg>
            <span>The app will close and apply the update silently.</span>
          </div>
          <div className="update-notice-item">
            <svg
              width="15"
              height="15"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.2"
              strokeLinecap="round"
              strokeLinejoin="round"
              className="notice-icon-spin"
              aria-hidden="true"
            >
              <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67" />
            </svg>
            <span>ARAOffice will reopen automatically in a few seconds.</span>
          </div>
          <div className="update-notice-item update-notice-warn">
            <svg
              width="15"
              height="15"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.2"
              strokeLinecap="round"
              strokeLinejoin="round"
              className="notice-icon-warn"
              aria-hidden="true"
            >
              <circle cx="12" cy="12" r="10" />
              <line x1="12" y1="8" x2="12" y2="12" />
              <line x1="12" y1="16" x2="12.01" y2="16" />
            </svg>
            <span>Please do not shut down Windows during this process.</span>
          </div>
        </div>

        {/* Actions */}
        <div className="update-modal-actions">
          <button type="button" className="btn btn-secondary update-later-btn" onClick={onClose}>
            {t('restartLater')}
          </button>
          <button
            type="button"
            className="btn btn-primary update-restart-btn"
            onClick={onRestart}
            autoFocus
          >
            <svg
              width="15"
              height="15"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.2"
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden="true"
            >
              <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67" />
            </svg>
            <span>{t('restartNow')}</span>
          </button>
        </div>
      </div>
    </div>
  )
}
