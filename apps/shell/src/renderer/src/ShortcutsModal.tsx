import { useEffect } from 'react'
import { useI18n } from './locale'

interface ShortcutsModalProps {
  onClose: () => void
}

interface ShortcutItem {
  keys: string[]
  description: string
}

interface ShortcutGroup {
  category: string
  items: ShortcutItem[]
}

export function ShortcutsModal({ onClose }: ShortcutsModalProps) {
  const { t } = useI18n()
  const isMac = navigator.platform.toUpperCase().includes('MAC')
  const mod = isMac ? '⌘' : 'Ctrl'

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onClose])

  const groups: ShortcutGroup[] = [
    {
      category: 'Tabs & Navigation',
      items: [
        { keys: [mod, 'Tab'], description: 'Switch to next open tab' },
        { keys: [mod, 'Shift', 'Tab'], description: 'Switch to previous tab' },
        { keys: [mod, 'W'], description: 'Close active tab' },
        { keys: [mod, 'O'], description: 'Open & browse files' },
        { keys: [mod, '/'], description: 'Toggle shortcuts cheat sheet' },
      ],
    },
    {
      category: 'File & Editing',
      items: [
        { keys: [mod, 'S'], description: 'Save current document' },
        { keys: [mod, 'Shift', 'S'], description: 'Save as new file' },
        { keys: [mod, 'Z'], description: 'Undo last action' },
        { keys: [mod, 'Y'], description: 'Redo last action' },
        { keys: [mod, 'F'], description: 'Find & search in document' },
      ],
    },
    {
      category: 'AI Assistant',
      items: [
        { keys: [mod, 'L'], description: 'Toggle / focus AI Assistant panel' },
        { keys: ['Esc'], description: 'Dismiss modal / cancel generation' },
      ],
    },
  ]

  return (
    <div className="modal-overlay shortcuts-modal-overlay" onClick={onClose}>
      <div
        className="modal shortcuts-modal"
        role="dialog"
        aria-modal="true"
        aria-label="Keyboard Shortcuts"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="shortcuts-ambient-glow" aria-hidden="true" />

        {/* Modal Header */}
        <div className="shortcuts-modal-header">
          <div className="shortcuts-title-wrap">
            <div className="shortcuts-icon-badge">
              <svg
                width="18"
                height="18"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.2"
              >
                <rect x="2" y="4" width="20" height="16" rx="2" />
                <path d="M6 8h.01M10 8h.01M14 8h.01M18 8h.01M6 12h.01M10 12h.01M14 12h.01M18 12h.01M8 16h8" />
              </svg>
            </div>
            <div>
              <h3 className="shortcuts-modal-title">Keyboard Shortcuts</h3>
              <p className="shortcuts-modal-subtitle">Essential hotkeys for navigating ARAOffice</p>
            </div>
          </div>
          <button
            type="button"
            className="shortcuts-close-btn"
            onClick={onClose}
            aria-label={t('cancel')}
          >
            <svg
              width="15"
              height="15"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.2"
            >
              <line x1="18" y1="6" x2="6" y2="18" />
              <line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        </div>

        {/* Shortcut Groups Grid */}
        <div className="shortcuts-grid">
          {groups.map((grp) => (
            <div key={grp.category} className="shortcuts-group-card">
              <h4 className="shortcuts-group-title">{grp.category}</h4>
              <div className="shortcuts-list">
                {grp.items.map((it, idx) => (
                  <div key={idx} className="shortcut-row">
                    <span className="shortcut-desc">{it.description}</span>
                    <div className="shortcut-keys">
                      {it.keys.map((k, kIdx) => (
                        <kbd key={kIdx} className="shortcut-kbd">
                          {k}
                        </kbd>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>

        {/* Modal Footer */}
        <div className="shortcuts-modal-footer">
          <span className="shortcuts-footer-tip">
            Tip: Press <kbd className="shortcut-kbd-inline">Ctrl + /</kbd> anytime on Home to toggle
            this cheatsheet
          </span>
          <button
            type="button"
            className="btn btn-secondary shortcuts-done-btn"
            onClick={onClose}
            autoFocus
          >
            Close
          </button>
        </div>
      </div>
    </div>
  )
}
