import type React from 'react'
import { useEffect, useRef } from 'react'
import { IconClock } from './icons'

export interface AiSnapshotItem {
  readonly id?: string | number | undefined
  readonly label: string
  readonly time: string
}

export interface AiSnapshotsPopoverProps<T extends AiSnapshotItem> {
  readonly isOpen: boolean
  readonly onClose: () => void
  readonly snapshots: readonly T[]
  readonly onRollback: (snapshot: T) => void
  readonly title?: string | undefined
  readonly rollbackLabel?: string | undefined
  readonly emptyLabel?: string | undefined
  readonly busy?: boolean | undefined
}

export function AiSnapshotsPopover<T extends AiSnapshotItem>({
  isOpen,
  onClose,
  snapshots,
  onRollback,
  title = 'Version Snapshots',
  rollbackLabel = 'Rollback',
  emptyLabel = 'No snapshots available',
  busy = false,
}: AiSnapshotsPopoverProps<T>): React.JSX.Element | null {
  const popoverRef = useRef<HTMLDivElement | null>(null)

  useEffect(() => {
    if (!isOpen) return
    const onPointerDown = (e: PointerEvent) => {
      if (popoverRef.current && !popoverRef.current.contains(e.target as Node)) {
        onClose()
      }
    }
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('pointerdown', onPointerDown)
    window.addEventListener('keydown', onKeyDown)
    return () => {
      window.removeEventListener('pointerdown', onPointerDown)
      window.removeEventListener('keydown', onKeyDown)
    }
  }, [isOpen, onClose])

  if (!isOpen) return null

  return (
    <div className="ai-snapshots-popover" ref={popoverRef} role="dialog" aria-label={title}>
      <div className="ai-snapshots-popover-header">
        <span className="ai-snapshots-popover-title">
          <IconClock size={14} />
          <span>{title}</span>
          <span className="ai-snapshots-count">{snapshots.length}</span>
        </span>
        <button
          type="button"
          className="ai-snapshots-close-btn"
          onClick={onClose}
          title="Close snapshots"
          aria-label="Close snapshots"
        >
          ✕
        </button>
      </div>

      <div className="ai-snapshots-list">
        {snapshots.length === 0 ? (
          <div className="ai-snapshots-empty">{emptyLabel}</div>
        ) : (
          snapshots.map((s, idx) => (
            <div key={s.id ?? `${s.time}-${idx}`} className="ai-snapshot-row">
              <div className="ai-snapshot-info" title={s.label}>
                <span className="ai-snapshot-time">{s.time}</span>
                <span className="ai-snapshot-label">{s.label}</span>
              </div>
              <button
                type="button"
                className="ai-snapshot-rollback-btn"
                disabled={busy}
                onClick={() => {
                  onRollback(s)
                  onClose()
                }}
              >
                {rollbackLabel}
              </button>
            </div>
          ))
        )}
      </div>
    </div>
  )
}
