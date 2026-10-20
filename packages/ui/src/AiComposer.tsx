import type React from 'react'
import { useEffect, useRef } from 'react'
import { IconArrowUp, IconSend, IconStop } from './icons'

// Keep in sync with the CSS `max-height` on `.ai-input-box textarea` (7 lines à 20px)
const MAX_TEXTAREA_HEIGHT = 140

/**
 * The AI panel input box shared by docs and sheets: auto-growing textarea
 * (Enter sends, Shift+Enter newline, Esc stops) plus a footer with optional
 * app-specific controls, a shortcut hint, and the send/stop button.
 * Renders the `.ai-input-box` class family; each app themes it in its own CSS.
 */
export function AiComposer({
  value,
  busy,
  disabled = false,
  placeholder,
  hintIdle,
  hintBusy,
  hintIdleTitle,
  sendLabel,
  stopLabel,
  ariaLabel,
  header,
  footerStart,
  iconOnly = false,
  stopIcon,
  textareaRef,
  sendIconEnabled,
  sendIconDisabled,
  authGate,
  onChange,
  onSend,
  onStop,
  onPasteFiles,
}: {
  readonly value: string
  readonly busy: boolean
  readonly disabled?: boolean
  readonly placeholder: string
  readonly hintIdle: string
  readonly hintBusy: string
  readonly hintIdleTitle?: string | undefined
  readonly sendLabel: string
  readonly stopLabel: string
  readonly ariaLabel?: string | undefined
  /** content inside the box above the textarea (attachment chips, …) — composer style */
  readonly header?: React.ReactNode
  /** extra controls at the left of the footer (attach buttons, toggles, …) */
  readonly footerStart?: React.ReactNode
  /** compact variant: no hint text, icon-only enter/stop button */
  readonly iconOnly?: boolean | undefined
  /** custom art for the icon-only stop button while busy; falls back to IconStop */
  readonly stopIcon?: React.ReactNode
  /** custom art for the send button states */
  readonly sendIconEnabled?: React.ReactNode
  readonly sendIconDisabled?: React.ReactNode
  /** pass a ref to focus the textarea from outside */
  readonly textareaRef?: React.RefObject<HTMLTextAreaElement | null> | undefined
  readonly authGate?: React.ReactNode
  readonly onChange: (next: string) => void
  readonly onSend: () => void
  readonly onStop: () => void
  /** clipboard files pasted into the textarea (screenshots, copied files); text paste stays native */
  readonly onPasteFiles?: ((files: File[]) => void) | undefined
}): React.JSX.Element {
  const innerRef = useRef<HTMLTextAreaElement | null>(null)
  const ref = textareaRef ?? innerRef
  const canSend = value.trim().length > 0 && !busy && !disabled

  // auto-grow up to ~6 lines; empty clears the inline height outright so the
  // CSS min-height governs (a hidden-at-measure pass can leave a stale value).
  useEffect(() => {
    const ta = ref.current
    if (!ta) return
    if (value === '') {
      ta.style.height = ''
      return
    }
    ta.style.height = 'auto'
    ta.style.height = `${Math.min(ta.scrollHeight, MAX_TEXTAREA_HEIGHT)}px`
  }, [value, ref])

  if (authGate) return <>{authGate}</>

  return (
    <div className="ai-input-box">
      {header}
      <textarea
        ref={ref}
        disabled={disabled}
        value={value}
        placeholder={placeholder}
        aria-label={ariaLabel}
        rows={1}
        onChange={(e) => onChange(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
            e.preventDefault()
            if (canSend) onSend()
          } else if (e.key === 'Escape' && busy) {
            e.preventDefault()
            onStop()
          }
        }}
        onPaste={(e) => {
          if (!onPasteFiles) return
          const files = Array.from(e.clipboardData.files)
          if (files.length === 0) return
          e.preventDefault()
          onPasteFiles(files)
        }}
      />
      <div className="ai-input-footer">
        {footerStart}
        {!iconOnly && (
          <span className="ai-input-hint" title={busy ? undefined : hintIdleTitle}>
            {busy ? hintBusy : hintIdle}
          </span>
        )}
        {/*
          Holds the send button against the right edge. The hint alone cannot do
          it: it renders only when !iconOnly (every app passes iconOnly), and each
          app restates the footer/hint rules in its own stylesheet - so the
          guarantee lives here, where no stylesheet can drop it.
        */}
        <span aria-hidden="true" style={{ flex: '1 1 0%', minWidth: 0 }} />
        {busy ? (
          <button
            className="ai-send-btn ai-stop-btn"
            onClick={onStop}
            title={stopLabel}
            aria-label={stopLabel}
          >
            {iconOnly ? (stopIcon ?? <IconStop size={20} />) : <IconStop size={20} />}
            {!iconOnly && stopLabel}
          </button>
        ) : (
          <button
            className="ai-send-btn"
            onClick={onSend}
            disabled={!canSend}
            title={sendLabel}
            aria-label={sendLabel}
          >
            {iconOnly ? (
              ((canSend ? sendIconEnabled : sendIconDisabled) ?? <IconArrowUp size={18} />)
            ) : (
              <IconSend size={16} />
            )}
            {!iconOnly && sendLabel}
          </button>
        )}
      </div>
    </div>
  )
}
