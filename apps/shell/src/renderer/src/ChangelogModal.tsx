import { useEffect, useState } from 'react'
import type { AppUpdateState, ChangelogEntry } from '../../shared/update-api'
import { useI18n } from './locale'

interface ChangelogModalProps {
  onClose: () => void
  updateState: AppUpdateState
  onRestart: () => void
}

const BADGE_MAP: Record<string, { label: string; className: string }> = {
  new: { label: 'NEW', className: 'chip-new' },
  feature: { label: 'NEW', className: 'chip-new' },
  feat: { label: 'NEW', className: 'chip-new' },
  add: { label: 'NEW', className: 'chip-new' },
  fix: { label: 'FIX', className: 'chip-fix' },
  fixes: { label: 'FIX', className: 'chip-fix' },
  bug: { label: 'FIX', className: 'chip-fix' },
}

function parseFormattedLine(line: string) {
  const tagMatch = line.match(/^(\*\*([a-zA-Z0-9_\s-]+)\*\*|\[([a-zA-Z0-9_\s-]+)\]):\s*(.*)$/)
  if (tagMatch) {
    const rawTag = (tagMatch[2] || tagMatch[3] || '').trim().toLowerCase()
    const rest = tagMatch[4]
    const badgeInfo = BADGE_MAP[rawTag]
    if (badgeInfo) {
      return { badge: badgeInfo, content: rest }
    }
    return { badge: null, content: `${tagMatch[2] || tagMatch[3]}: ${rest}` }
  }
  return { badge: null, content: line }
}

function formatInline(str: string) {
  const parts = str.split(/(\*\*.*?\*\*|`.*?`)/g)
  return parts.map((part, i) => {
    if (part.startsWith('**') && part.endsWith('**')) {
      return <strong key={i}>{part.slice(2, -2)}</strong>
    }
    if (part.startsWith('`') && part.endsWith('`')) {
      return (
        <code key={i} className="cl-code">
          {part.slice(1, -1)}
        </code>
      )
    }
    return part
  })
}

function renderMarkdownBody(text: string) {
  if (!text) return null
  const lines = text.split('\n')
  const elements: React.ReactNode[] = []

  let listItems: React.ReactNode[] = []
  let listKey = 0

  const flushList = () => {
    if (listItems.length > 0) {
      elements.push(
        <ul key={`list-${listKey++}`} className="cl-list">
          {listItems}
        </ul>,
      )
      listItems = []
    }
  }

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i].trim()
    if (!rawLine) {
      flushList()
      continue
    }

    if (rawLine.startsWith('### ') || rawLine.startsWith('## ') || rawLine.startsWith('# ')) {
      flushList()
      const headingText = rawLine.replace(/^#+\s*/, '')
      elements.push(
        <h4 key={`h-${i}`} className="cl-heading">
          {headingText}
        </h4>,
      )
    } else if (rawLine.startsWith('- ') || rawLine.startsWith('* ')) {
      const itemText = rawLine.slice(2).trim()
      const { badge, content } = parseFormattedLine(itemText)
      listItems.push(
        <li key={`li-${i}`} className={`cl-list-item ${badge ? 'has-badge' : ''}`}>
          {badge && <span className={`cl-tag ${badge.className}`}>{badge.label}</span>}
          <span className="cl-item-text">{formatInline(content)}</span>
        </li>,
      )
    } else {
      flushList()
      const { badge, content } = parseFormattedLine(rawLine)
      elements.push(
        <p key={`p-${i}`} className="cl-paragraph">
          {badge && <span className={`cl-tag ${badge.className}`}>{badge.label}</span>}
          <span>{formatInline(content)}</span>
        </p>,
      )
    }
  }
  flushList()

  return elements
}

export function ChangelogModal({ onClose, updateState, onRestart }: ChangelogModalProps) {
  const { t, dateLocale } = useI18n()
  const [entries, setEntries] = useState<ChangelogEntry[]>([])
  const [loading, setLoading] = useState(true)
  const [checkingManual, setCheckingManual] = useState(false)

  const loadData = () => {
    setLoading(true)
    void window.aiOffice.getChangelog().then((res) => {
      setEntries(res)
      setLoading(false)
    })
  }

  useEffect(() => {
    loadData()
  }, [])

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onClose])

  const handleManualCheck = () => {
    setCheckingManual(true)
    void window.aiOffice.checkForUpdates()
    setTimeout(() => {
      loadData()
      setCheckingManual(false)
    }, 1200)
  }

  const isDownloaded = updateState.status === 'downloaded'
  const isDownloading = updateState.status === 'downloading'

  return (
    <div className="modal-overlay cl-overlay" onClick={onClose}>
      <div
        className="cl-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('changelogTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Ambient Glow */}
        <div className="cl-ambient-glow" aria-hidden="true" />

        {/* Modal Header */}
        <header className="cl-header">
          <div className="cl-header-left">
            <span className="cl-eyebrow">
              <span className="cl-spark" aria-hidden="true">
                ✦
              </span>
              ARAOFFICE
            </span>
            <h2 className="cl-title">{t('changelogTitle')}</h2>
          </div>
          <button
            type="button"
            className="cl-close-btn"
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

        {/* Update Banner (when ready) */}
        {isDownloaded ? (
          <div className="cl-update-banner ready">
            <div className="cl-banner-content">
              <div className="cl-pulse-ring ready" />
              <div>
                <strong>{t('updateReadyTitle')}</strong>
                <p>{t('updateReadyDesc', { v: updateState.availableVersion ?? '' })}</p>
              </div>
            </div>
            <button type="button" className="btn btn-primary cl-banner-action" onClick={onRestart}>
              {t('restartNow')}
            </button>
          </div>
        ) : isDownloading ? (
          <div className="cl-update-banner downloading">
            <div className="cl-banner-content">
              <div className="cl-pulse-ring downloading" />
              <div>
                <strong>{t('updateAvailable')}</strong>
                <p>{t('updateDownloading', { p: updateState.percent ?? 0 })}</p>
              </div>
            </div>
            <div className="cl-progress-bar">
              <div className="cl-progress-fill" style={{ width: `${updateState.percent ?? 0}%` }} />
            </div>
          </div>
        ) : null}

        {/* Body */}
        <div className="cl-body">
          {loading ? (
            <div className="cl-empty-state">
              <div className="cl-spinner" />
              <span>{t('checkUpdates')}…</span>
            </div>
          ) : entries.length === 0 ? (
            <div className="cl-empty-state">
              <p>{t('changelogEmpty')}</p>
            </div>
          ) : (
            <div className="cl-timeline">
              {entries.map((entry, index) => {
                let formattedDate = ''
                try {
                  if (entry.publishedAt) {
                    formattedDate = new Date(entry.publishedAt).toLocaleDateString(dateLocale, {
                      year: 'numeric',
                      month: 'short',
                      day: 'numeric',
                    })
                  }
                } catch {
                  formattedDate = entry.publishedAt
                }

                return (
                  <article key={entry.version || index} className="cl-card">
                    {/* Clean Card Header: Version + Date */}
                    <div className="cl-card-header">
                      <span className="cl-version-badge">v{entry.version}</span>
                      {formattedDate && <time className="cl-card-date">{formattedDate}</time>}
                    </div>

                    {/* Clean Body */}
                    <div className="cl-card-body">
                      {renderMarkdownBody(entry.body || t('changelogEmpty'))}
                    </div>

                    {/* External Link */}
                    {entry.htmlUrl && (
                      <div className="cl-card-footer">
                        <button
                          type="button"
                          className="cl-link-btn"
                          onClick={() => {
                            if (entry.htmlUrl)
                              void window.aiOfficeWindow.openExternal(entry.htmlUrl)
                          }}
                        >
                          GitHub Release
                          <svg
                            width="11"
                            height="11"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2.2"
                          >
                            <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6" />
                            <polyline points="15 3 21 3 21 9" />
                            <line x1="10" y1="14" x2="21" y2="3" />
                          </svg>
                        </button>
                      </div>
                    )}
                  </article>
                )
              })}
            </div>
          )}
        </div>

        {/* Footer */}
        {(() => {
          const isDownloading = updateState.status === 'downloading'
          const isDownloaded = updateState.status === 'downloaded'
          const isChecking = checkingManual || updateState.status === 'checking'
          const isError = updateState.status === 'error'
          // main flags the no-auto-update platform through the error text (the status union is fixed)
          const manualUrl =
            isError && updateState.error?.includes('Auto-update is not available on this platform')
              ? (updateState.error?.match(/https:\/\/\S+/)?.[0] ?? null)
              : null

          let statusDotClass: string
          let statusText: string

          if (manualUrl) {
            statusDotClass = 'error'
            statusText = 'Manual download required'
          } else if (isChecking) {
            statusDotClass = 'checking'
            statusText = 'Checking for updates…'
          } else if (isDownloading) {
            statusDotClass = 'downloading'
            statusText = `Downloading update (${updateState.percent ?? 0}%)…`
          } else if (isDownloaded) {
            statusDotClass = 'ready'
            statusText = `v${updateState.availableVersion ?? ''} · Ready to install`
          } else if (isError) {
            statusDotClass = 'error'
            statusText = 'Update check failed'
          } else {
            statusDotClass = ''
            statusText = `${updateState.currentVersion ? `v${updateState.currentVersion}` : ''} · ${t('upToDate')}`
          }

          return (
            <footer className="cl-footer">
              <div className="cl-footer-status">
                <span className={`cl-status-dot ${statusDotClass}`} />
                <span>{statusText}</span>
              </div>

              <div className="cl-footer-actions">
                {isDownloaded ? (
                  <button type="button" className="btn btn-primary cl-btn-done" onClick={onRestart}>
                    {t('restartNow')}
                  </button>
                ) : manualUrl ? (
                  <button
                    type="button"
                    className="btn btn-secondary cl-btn-check"
                    onClick={() => void window.aiOfficeWindow.openExternal(manualUrl)}
                  >
                    Download Update
                  </button>
                ) : (
                  <button
                    type="button"
                    className={`btn btn-secondary cl-btn-check ${isChecking ? 'loading' : ''}`}
                    onClick={handleManualCheck}
                    disabled={isChecking || isDownloading || loading}
                  >
                    <svg
                      width="13"
                      height="13"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2.2"
                      className={`cl-refresh-icon ${isChecking ? 'spin' : ''}`}
                    >
                      <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67" />
                    </svg>
                    <span>{isChecking ? 'Checking…' : t('checkUpdates')}</span>
                  </button>
                )}
                <button
                  type="button"
                  className={`btn ${isDownloaded ? 'btn-secondary' : 'btn-primary'} cl-btn-done`}
                  onClick={onClose}
                >
                  {t('cancel') === 'Cancel' ? 'Done' : t('cancel')}
                </button>
              </div>
            </footer>
          )
        })()}
      </div>
    </div>
  )
}
