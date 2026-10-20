import type React from 'react'
import { useEffect, useMemo, useRef, useState } from 'react'
import {
  AiComposer,
  AraSpark,
  AiPanelHeader,
  AiPanelShell,
  AiProcessCard,
  AiSnapshotsPopover,
  AiTypingIndicator,
  AiUnconfiguredCard,
  clampAiPanelWidth,
  extractThinking,
  IconAttach,
  IconCheck,
  IconCopy,
  IconNewChat,
  IconSidebarCollapse,
  IconUndo,
  Markdown,
  useAiReady,
} from '@araoffice/ui'
import type { ChangePlan } from '../../domain/workbook.types'
import { ATTACHMENT_IMAGE_EXTS, type AttachmentMeta } from '../../shared/desktop-api'
import { getLang, useI18n, type TFunc } from '../i18n/locale'
import fileImageIcon from '../assets/file-image.png'

/** Clipboard bitmap MIME → attachment extension (matches the main process's
 * ATTACHMENT_IMAGE_EXTS) */
const PASTE_MIME_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
}

const AI_CONTENT_STARTERS_ID = [
  {
    label: 'Analisis Data & Tren',
    prompt:
      'Analisis data pada lembar ini, temukan tren utama dan anomali, serta hitung total dan metrik ringkasannya.',
  },
  {
    label: 'Neraca Keuangan',
    prompt:
      'Buat tabel neraca keuangan (balance sheet) lengkap dengan aset lancar, aset tetap, kewajiban, ekuitas, dan formula total yang seimbang.',
  },
  {
    label: 'Anggaran Bulanan',
    prompt:
      'Buat tabel anggaran bulanan dengan kategori pengeluaran, estimasi biaya, pengeluaran aktual, selisih, dan formula SUM.',
  },
  {
    label: 'Format Tabel Profesional',
    prompt:
      'Rapikan format tabel ini: gunakan header yang jelas, format angka/mata uang yang sesuai, dan styling profesional.',
  },
] as const

const AI_CONTENT_STARTERS_EN = [
  {
    label: 'Analyze Data & Trends',
    prompt:
      'Analyze the data on this sheet, identify key trends and anomalies, and calculate summary metrics.',
  },
  {
    label: 'Financial Balance Sheet',
    prompt:
      'Create a complete financial balance sheet table with assets, liabilities, equity, and balanced total formulas.',
  },
  {
    label: 'Monthly Budget Tracker',
    prompt:
      'Create a monthly budget tracker with categories, budgeted amounts, actual expenses, variance, and SUM formulas.',
  },
  {
    label: 'Format Table Professionally',
    prompt:
      'Format and style this table professionally with clear headers, proper number formatting, and clean layout.',
  },
] as const

const AI_BUILD_STARTERS_ID = [
  {
    label: 'Neraca Keuangan',
    prompt:
      'Buat tabel neraca keuangan (balance sheet) lengkap dengan aset lancar, aset tetap, kewajiban, ekuitas, dan formula total yang seimbang.',
  },
  {
    label: 'Anggaran Bulanan',
    prompt:
      'Buat tabel anggaran bulanan dengan kategori, estimasi biaya, pengeluaran aktual, selisih, dan formula SUM.',
  },
  {
    label: 'Laporan Penjualan',
    prompt:
      'Buat tabel laporan performa penjualan bulanan dengan nama produk, target, realisasi, persentase pencapaian, dan formula otomatis.',
  },
  {
    label: 'Pelacak Proyek',
    prompt:
      'Buat tabel pelacak tugas proyek dengan nama tugas, penanggung jawab, deadline, status, dan prioritas.',
  },
] as const

const AI_BUILD_STARTERS_EN = [
  {
    label: 'Financial Balance Sheet',
    prompt:
      'Create a complete financial balance sheet table with assets, liabilities, equity, and balanced total formulas.',
  },
  {
    label: 'Monthly Budget Tracker',
    prompt:
      'Create a monthly budget tracker with categories, budgeted amount, actual expenses, variance, and SUM formulas.',
  },
  {
    label: 'Sales Performance Report',
    prompt:
      'Create a monthly sales performance report table with products, targets, actuals, achievement percentage, and formulas.',
  },
  {
    label: 'Project Task Tracker',
    prompt:
      'Create a project task tracker with task name, assignee, due date, status, and priority.',
  },
] as const

export interface SheetSnapshot {
  readonly id: string
  readonly label: string
  readonly time: string
  readonly snapshot?: unknown
}

/** Card name slot width: 190 card - 2 border - 8/14 padding - 40 icon - 10 gap */
const CARD_NAME_MAX_WIDTH = 116
let cardNameCtx: CanvasRenderingContext2D | null = null

/** Ellipsize like the design: cut at the limit, strip trailing -_./spaces so
 *  punctuation never sits against the …; CSS text-overflow stays as fallback */
function truncateCardName(name: string): string {
  cardNameCtx ??= document.createElement('canvas').getContext('2d')
  if (!cardNameCtx) return name
  // must match the stack the card name actually renders with (body font in styles.css)
  cardNameCtx.font =
    "500 13px 'Segoe UI', -apple-system, BlinkMacSystemFont, 'PingFang SC', 'Microsoft YaHei', sans-serif"
  if (cardNameCtx.measureText(name).width <= CARD_NAME_MAX_WIDTH) return name
  let lo = 1
  let hi = name.length
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1
    if (cardNameCtx.measureText(`${name.slice(0, mid)}…`).width <= CARD_NAME_MAX_WIDTH) lo = mid
    else hi = mid - 1
  }
  return `${name.slice(0, lo).replace(/[-_.\s]+$/, '')}…`
}

function formatAttachmentSize(bytes: number): string {
  return bytes >= 1024 * 1024
    ? `${(bytes / (1024 * 1024)).toFixed(2)} MB`
    : `${(bytes / 1024).toFixed(2)} KB`
}

/** Resizable panel width: persisted, clamped to min/max, drives the .sheet-body grid column via --ai-panel-width */
const PANEL_WIDTH_KEY = 'sheets-ai-panel-width'
const PANEL_WIDTH_MIN = 360

function loadPanelWidth(): number | null {
  const saved = Number(localStorage.getItem(PANEL_WIDTH_KEY))
  return Number.isFinite(saved) && saved > 0 ? clampAiPanelWidth(saved, PANEL_WIDTH_MIN) : null
}

export interface AiToolChip {
  readonly summary: string
  readonly isError: boolean
  /** still executing: rendered as a spinner chip, replaced in place when the tool finishes */
  readonly running?: boolean
  /** Tool name (title tooltip) */
  readonly name?: string
  /** Tool output (truncated UI-side); when present the row expands to details */
  readonly output?: string
}

export interface AiChatMessage {
  readonly role: 'user' | 'assistant'
  readonly text: string
  readonly thinking?: string | undefined
  readonly tools: readonly AiToolChip[]
  readonly streaming?: boolean | undefined
  readonly isError?: boolean | undefined
  readonly turnLimit?: boolean | undefined
  readonly stuck?: boolean | undefined
  readonly undelivered?: boolean | undefined
  /** Set when this message reflects an auto-applied plan; renders an inline [Undo] button. */
  readonly autoApplied?: { readonly opCount: number } | undefined
  readonly attachments?: readonly AttachmentMeta[] | undefined
}

export function AiChatPanel({
  isOpen,
  hasContent,
  chat,
  historicChat = [],
  attachments,
  attachNotice,
  onPickAttachments,
  onAddAttachmentPaths,
  onAddPastedImage,
  onRemoveAttachment,
  prompt,
  preview,
  aiBusy,
  onPromptChange,
  onSend,
  onStop,
  onNewChat,
  onContinue,
  onUndo,
  onExpand,
  onCollapse,
  snapshots = [],
  onRollback,
  onJumpTarget,
}: {
  readonly isOpen: boolean
  /** the workbook has cells with content — empty workbooks get "build me a sheet" copy instead */
  readonly hasContent: boolean
  readonly chat: readonly AiChatMessage[]
  readonly historicChat?: readonly AiChatMessage[]
  /// Chat attachments (chips + 📎 button + drag onto the panel), same structure
  /// as the docs/slides AI panels.
  readonly attachments: readonly AttachmentMeta[]
  readonly attachNotice: string | null
  readonly onPickAttachments: () => void
  readonly onAddAttachmentPaths: (paths: readonly string[]) => void
  /// Clipboard-pasted bitmaps (screenshots etc. without a local path): bytes +
  /// extension
  readonly onAddPastedImage: (data: ArrayBuffer, ext: string) => void
  readonly onRemoveAttachment: (path: string) => void
  readonly prompt: string
  readonly preview: ChangePlan | null
  readonly aiBusy: boolean
  readonly onPromptChange: (prompt: string) => void
  /** Send the composer text, or the given instruction when provided (used by the failed-run Retry) */
  readonly onSend: (instruction?: string) => void
  readonly onStop: () => void
  readonly onContinue: () => void
  readonly onNewChat: () => void
  readonly onUndo: () => void
  readonly onExpand: () => void
  readonly onCollapse: () => void
  readonly snapshots?: readonly SheetSnapshot[]
  readonly onRollback?: (snapshot: SheetSnapshot) => void
  readonly onJumpTarget?: (target: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const { isReady } = useAiReady({
    getAiSettings: window.desktopApi?.getAiSettings
      ? () => window.desktopApi.getAiSettings()
      : undefined,
    onAiSettingsChanged: window.desktopApi?.onAiSettingsChanged
      ? (h) => window.desktopApi.onAiSettingsChanged!(h)
      : undefined,
  })
  const [snapshotsOpen, setSnapshotsOpen] = useState(false)
  const chatRef = useRef<HTMLDivElement | null>(null)
  const inputRef = useRef<HTMLTextAreaElement | null>(null)
  const stickToBottomRef = useRef(true)
  const [dragOver, setDragOver] = useState(false)
  const asideRef = useRef<HTMLElement | null>(null)
  const [resizing, setResizing] = useState(false)
  const displaySnapshots = useMemo(() => {
    if (!snapshots || snapshots.length === 0) return []
    const seenIds = new Set<string>()
    const result: SheetSnapshot[] = []
    for (const s of snapshots) {
      if (seenIds.has(s.id)) continue
      seenIds.add(s.id)
      result.push(s)
    }
    return result
  }, [snapshots])
  /** data-URL previews for image attachments, keyed by path (ARA Assistant composer thumbnails) */
  const [attachmentPreviews, setAttachmentPreviews] = useState<Record<string, string>>({})
  /** image paths with a read already issued — one readAttachmentImage per attach, even while pending */
  const previewRequestedRef = useRef(new Set<string>())
  useEffect(() => {
    const alive = new Set(attachments.map((a) => a.path))
    // drop previews (and request markers) of removed attachments, so memory is reclaimed and a re-attach re-reads
    setAttachmentPreviews((prev) => {
      const stale = Object.keys(prev).filter((p) => !alive.has(p))
      if (stale.length === 0) return prev
      const next = { ...prev }
      for (const p of stale) delete next[p]
      return next
    })
    for (const p of previewRequestedRef.current) {
      if (!alive.has(p)) previewRequestedRef.current.delete(p)
    }
    for (const a of attachments) {
      if (!ATTACHMENT_IMAGE_EXTS.has(a.ext) || previewRequestedRef.current.has(a.path)) continue
      previewRequestedRef.current.add(a.path)
      void window.desktopApi.readAttachmentImage(a.path).then((r) => {
        if (!previewRequestedRef.current.has(a.path)) return // removed while the read was in flight
        if (r.ok && r.base64 && r.mime) {
          setAttachmentPreviews((prev) => ({
            ...prev,
            [a.path]: `data:${r.mime};base64,${r.base64}`,
          }))
        }
      })
    }
  }, [attachments])
  /** paints the strip's scrollbar thumb while the user scrolls it (cleared 800ms after the last event) */
  const attachScrollFadeRef = useRef(0)
  const onAttachmentsScroll = (e: React.UIEvent<HTMLDivElement>): void => {
    const el = e.currentTarget
    el.classList.add('is-scrolling')
    window.clearTimeout(attachScrollFadeRef.current)
    attachScrollFadeRef.current = window.setTimeout(() => el.classList.remove('is-scrolling'), 800)
  }
  /** Wall-clock start of the current run (aiBusy false→true), drives the elapsed badge */
  const busyStartRef = useRef(0)
  useEffect(() => {
    if (aiBusy) busyStartRef.current = Date.now()
  }, [aiBusy])

  // Restore the persisted panel width (the grid column tracks --ai-panel-width on .sheet-body)
  useEffect(() => {
    if (!isOpen) return
    const saved = loadPanelWidth()
    if (saved === null) return
    const area = asideRef.current?.closest('.sheet-body') as HTMLElement | null
    area?.style.setProperty('--ai-panel-width', `${saved}px`)
  }, [isOpen])

  // Re-clamp the panel width when the window shrinks (max is 60% of the window)
  useEffect(() => {
    if (!isOpen) return
    const onResize = (): void => {
      const area = asideRef.current?.closest('.sheet-body') as HTMLElement | null
      const current = parseFloat(area?.style.getPropertyValue('--ai-panel-width') ?? '')
      if (!area || !Number.isFinite(current)) return
      area.style.setProperty('--ai-panel-width', `${clampAiPanelWidth(current, PANEL_WIDTH_MIN)}px`)
    }
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [isOpen])

  // follow the stream, but stop yanking once the user scrolls up to read
  useEffect(() => {
    if (stickToBottomRef.current) {
      chatRef.current?.scrollTo({ top: chatRef.current.scrollHeight })
    }
  }, [chat, preview])

  const resizeCleanupRef = useRef<(() => void) | null>(null)
  useEffect(() => () => resizeCleanupRef.current?.(), [])

  /** Drag the left edge to resize: the panel is flush with the window's right edge, so width = window.innerWidth - clientX. */
  const startResize = (e: React.PointerEvent<HTMLDivElement>): void => {
    e.preventDefault()
    const area = asideRef.current?.closest('.sheet-body') as HTMLElement | null
    if (!area) return
    const resizer = e.currentTarget
    setResizing(true)
    area.style.transition = 'none'
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    let width = 0
    const onMove = (ev: PointerEvent): void => {
      width = clampAiPanelWidth(window.innerWidth - ev.clientX, PANEL_WIDTH_MIN)
      area.style.setProperty('--ai-panel-width', `${width}px`)
    }
    let done = false
    const cleanup = (): void => {
      if (done) return
      done = true
      resizeCleanupRef.current = null
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', cleanup)
      window.removeEventListener('pointercancel', cleanup)
      resizer.removeEventListener('lostpointercapture', cleanup)
      area.style.transition = ''
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
      setResizing(false)
      if (width > 0) localStorage.setItem(PANEL_WIDTH_KEY, String(Math.round(width)))
    }
    resizeCleanupRef.current = cleanup
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', cleanup)
    window.addEventListener('pointercancel', cleanup)
    // lostpointercapture also fires if the resizer is unmounted mid-drag (panel collapse)
    resizer.addEventListener('lostpointercapture', cleanup)
    resizer.setPointerCapture(e.pointerId)
  }

  const onChatScroll = (): void => {
    const el = chatRef.current
    if (!el) return
    stickToBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 48
  }

  if (!isOpen) {
    return (
      <aside className="copilot collapsed">
        <button className="expand-copilot ai-rail" onClick={onExpand} title={t('aiOpenAssistant')}>
          <span className="ai-rail-icon" aria-hidden="true">
            <AraSpark size={21} />
          </span>
        </button>
      </aside>
    )
  }

  const canSend = prompt.trim().length > 0 && !aiBusy && isReady

  const send = (): void => {
    if (!canSend) return
    stickToBottomRef.current = true
    onSend()
  }

  const onDrop = (e: React.DragEvent): void => {
    e.preventDefault()
    e.stopPropagation()
    setDragOver(false)
    const paths = Array.from(e.dataTransfer.files)
      .map((f) => window.desktopApi.getPathForFile(f))
      .filter(Boolean)
    if (paths.length > 0) onAddAttachmentPaths(paths)
  }

  /** Files pasted into the input: ones with a local path go the regular
   * attachment route; pure bitmaps like screenshots are persisted by the host */
  const onPasteFiles = (files: File[]): void => {
    const paths: string[] = []
    for (const f of files) {
      const p = window.desktopApi.getPathForFile(f)
      if (p) {
        paths.push(p)
        continue
      }
      const ext = PASTE_MIME_EXT[f.type] ?? f.name.split('.').pop()?.toLowerCase() ?? 'bin'
      void f.arrayBuffer().then((buf) => onAddPastedImage(buf, ext))
    }
    if (paths.length > 0) onAddAttachmentPaths(paths)
  }

  return (
    <AiPanelShell
      ref={asideRef}
      className={`copilot${dragOver ? ' ai-panel-dragover' : ''}${resizing ? ' ai-panel-resizing' : ''}`}
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes('Files')) {
          e.preventDefault()
          e.stopPropagation()
          setDragOver(true)
        }
      }}
      onDragLeave={(e) => {
        if (!e.currentTarget.contains(e.relatedTarget as Node | null)) setDragOver(false)
      }}
      onDrop={onDrop}
    >
      <div
        className="ai-panel-resizer"
        onPointerDown={startResize}
        role="separator"
        aria-orientation="vertical"
        aria-label="ARA Assistant"
      />
      <AiPanelHeader
        title="ARA Assistant"
        showNewChat={chat.length > 0 || historicChat.length > 0}
        onNewChat={onNewChat}
        onCollapse={onCollapse}
        newChatLabel={t('aiNewChat')}
        collapseLabel={t('aiCollapsePanel')}
        newChatIcon={<IconNewChat size={15} />}
        collapseIcon={<IconSidebarCollapse size={15} />}
        snapshotsCount={displaySnapshots.length}
        isSnapshotsOpen={snapshotsOpen}
        onToggleSnapshots={() => setSnapshotsOpen((v) => !v)}
        live={aiBusy}
      >
        <AiSnapshotsPopover
          isOpen={snapshotsOpen}
          onClose={() => setSnapshotsOpen(false)}
          snapshots={displaySnapshots}
          onRollback={(s) => onRollback?.(s)}
          title={t('aiSnapshotsTitle')}
          rollbackLabel={t('aiRollback')}
        />
      </AiPanelHeader>

      <div className="ai-chat" ref={chatRef} onScroll={onChatScroll}>
        {/* Past conversation (read-only transcript), shown continuously with the current turn */}
        {historicChat.length > 0 && (
          <>
            {historicChat.map((entry, i) => (
              <div key={`h${i}`} className={`ai-msg ai-msg-${entry.role} ai-msg-historic`}>
                {(entry.thinking || entry.tools.length > 0) && (
                  <AiProcessCard
                    thinking={entry.thinking}
                    tools={entry.tools}
                    isStreaming={false}
                  />
                )}
                {entry.text && <Markdown text={entry.text} />}
                {entry.role === 'user' &&
                  entry.attachments?.map((attachment) => (
                    <span className="ai-msg-attachment" key={attachment.path}>
                      <span className="ai-attachment-card-meta">
                        <span className="ai-attachment-card-name">
                          {truncateCardName(attachment.name)}
                        </span>
                        <span className="ai-attachment-card-size">
                          {formatAttachmentSize(attachment.sizeBytes)}
                        </span>
                      </span>
                    </span>
                  ))}
              </div>
            ))}
            <div className="ai-history-sep">{t('aiHistorySep')}</div>
          </>
        )}
        {chat.length === 0 && historicChat.length === 0 && (
          <div className="ai-chat-empty">
            {isReady ? (
              <>
                <div className="ai-chat-empty-title">
                  {t(hasContent ? 'aiEmptyTitle' : 'aiEmptyBuildTitle')}
                </div>
                <div className="ai-chat-empty-body">
                  {t(hasContent ? 'aiEmptyBodyLine1' : 'aiEmptyBuildBody')}
                </div>
                <div className="ai-starter-list">
                  {(hasContent
                    ? getLang() === 'id'
                      ? AI_CONTENT_STARTERS_ID
                      : AI_CONTENT_STARTERS_EN
                    : getLang() === 'id'
                      ? AI_BUILD_STARTERS_ID
                      : AI_BUILD_STARTERS_EN
                  ).map(({ label, prompt: pText }) => (
                    <button
                      key={label}
                      className="ai-starter"
                      type="button"
                      onClick={() => {
                        onPromptChange(pText)
                        inputRef.current?.focus()
                      }}
                    >
                      {label}
                    </button>
                  ))}
                </div>
              </>
            ) : (
              <AiUnconfiguredCard onConfigure={() => void window.desktopApi?.configureAi?.()} />
            )}
          </div>
        )}
        {chat.map((entry, index) => (
          <div
            key={index}
            className={`ai-msg ai-msg-${entry.role}${entry.isError ? ' ai-msg-error' : ''}${entry.role === 'assistant' && entry.streaming ? ' ai-msg-streaming' : ''}`}
          >
            {entry.role === 'user' ? (
              <>
                {entry.text}
                {entry.undelivered && (
                  <div className="ai-msg-undelivered">
                    {t('aiUndelivered')}
                    {!aiBusy && (
                      <button className="ai-retry-btn" onClick={() => onSend(entry.text)}>
                        {t('aiRetry')}
                      </button>
                    )}
                  </div>
                )}
              </>
            ) : (
              <>
                {(() => {
                  const extracted = extractThinking(entry.text)
                  const displayThinking = entry.thinking || extracted.thinking
                  const displayText = extracted.thinking ? extracted.text : entry.text
                  const hasProcess = Boolean(displayThinking || entry.tools.length > 0)
                  return (
                    <>
                      {hasProcess && (
                        <AiProcessCard
                          thinking={displayThinking}
                          tools={entry.tools}
                          isStreaming={entry.streaming}
                        />
                      )}
                      {displayText ? (
                        <Markdown text={displayText} onJumpTarget={onJumpTarget} />
                      ) : (
                        entry.streaming &&
                        !hasProcess && (
                          <span className="ai-typing-row">
                            <AiTypingIndicator label={t('aiWorking')} />
                          </span>
                        )
                      )}
                    </>
                  )
                })()}
                {!entry.streaming && !aiBusy && (entry.autoApplied || entry.text) && (
                  <div className="ai-msg-footer">
                    {entry.autoApplied && entry.autoApplied.opCount > 0 ? (
                      <div className="ai-applied-badge">
                        <IconCheck size={12} />
                        <span className="ai-applied-badge-text">
                          {t('aiAutoApplied', { count: entry.autoApplied.opCount })}
                        </span>
                        <button className="ai-undo-btn" onClick={onUndo} title={t('aiUndoTitle')}>
                          <IconUndo size={12} />
                          <span>{t('aiUndo')}</span>
                        </button>
                      </div>
                    ) : (
                      <span />
                    )}
                    {entry.text && (
                      <div className="ai-msg-toolbar">
                        <button
                          className="ai-msg-tool-btn"
                          onClick={() => void navigator.clipboard.writeText(entry.text)}
                          aria-label="Copy reply"
                          title="Copy reply"
                        >
                          <IconCopy size={13} />
                        </button>
                      </div>
                    )}
                  </div>
                )}
                {entry.turnLimit && !entry.stuck && index === chat.length - 1 && !aiBusy && (
                  <button className="ai-retry-btn" onClick={onContinue}>
                    {t('aiContinue')}
                  </button>
                )}
              </>
            )}
          </div>
        ))}

        {preview && (
          <section className="preview ai-preview-card" aria-label={t('aiPreviewAria')}>
            <h3>{t('aiProposedChanges')}</h3>
            {preview.structuralChanges.map((change, index) => (
              <div className="change" key={`structural-${index}`}>
                <strong>{t('aiChangeStructure')}</strong>
                <span>{change.label}</span>
              </div>
            ))}
            {preview.formatChanges.map((change, index) => (
              <div className="change" key={`format-${index}`}>
                <strong>{t('aiChangeFormat')}</strong>
                <span>{change.label}</span>
              </div>
            ))}
            {preview.cellChanges.slice(0, MAX_PREVIEW_CELL_ROWS).map((change) => (
              <div className="change" key={`${change.sheetId}-${change.address}`}>
                <strong>{change.address}</strong>
                <span>
                  {formatCell(change.before, t)} → {formatCell(change.after, t)}
                </span>
              </div>
            ))}
            {preview.cellChanges.length > MAX_PREVIEW_CELL_ROWS && (
              <div className="change">
                <strong>…</strong>
                <span>
                  {t('aiMoreCells', { count: preview.cellChanges.length - MAX_PREVIEW_CELL_ROWS })}
                </span>
              </div>
            )}
            {preview.sheetRenames.map((rename) => (
              <div className="change" key={rename.sheetId}>
                <strong>{t('aiChangeSheet')}</strong>
                <span>
                  {rename.before} → {rename.after}
                </span>
              </div>
            ))}
            {preview.warnings.map((warning) => (
              <div className="change" key={warning}>
                <strong>⚠</strong>
                <span>{warning}</span>
              </div>
            ))}
          </section>
        )}
      </div>
      {/* Reclaimed bottom space: version snapshots now live in AiPanelHeader popover */}

      <div className="ai-composer">
        {attachNotice && <div className="ai-attach-notice">{attachNotice}</div>}
        <AiComposer
          header={
            attachments.length > 0 && (
              <div className="ai-attachments" onScroll={onAttachmentsScroll}>
                {attachments.map((attachment) =>
                  ATTACHMENT_IMAGE_EXTS.has(attachment.ext) ? (
                    <span
                      key={attachment.path}
                      className="ai-attachment-thumb"
                      title={attachment.path}
                    >
                      {attachmentPreviews[attachment.path] ? (
                        <img src={attachmentPreviews[attachment.path]} alt={attachment.name} />
                      ) : (
                        <span className="ai-attachment-thumb-pending" aria-hidden>
                          <img src={fileImageIcon} alt="" />
                        </span>
                      )}
                      <button
                        className="ai-attachment-thumb-remove"
                        onClick={() => onRemoveAttachment(attachment.path)}
                        title={t('aiRemoveAttachment')}
                        aria-label={t('aiRemoveAttachment')}
                      >
                        <svg width="16" height="16" viewBox="0 0 32 32" aria-hidden>
                          <path
                            d="M24 9.4L22.6 8L16 14.6L9.4 8L8 9.4l6.6 6.6L8 22.6L9.4 24l6.6-6.6l6.6 6.6l1.4-1.4l-6.6-6.6L24 9.4z"
                            fill="currentColor"
                            stroke="currentColor"
                            strokeWidth="0.25"
                          />
                        </svg>
                      </button>
                    </span>
                  ) : (
                    <span
                      key={attachment.path}
                      className="ai-attachment-card"
                      title={attachment.path}
                    >
                      <span className="ai-attachment-card-meta">
                        <span className="ai-attachment-card-name">
                          {truncateCardName(attachment.name)}
                        </span>
                        <span className="ai-attachment-card-size">
                          {formatAttachmentSize(attachment.sizeBytes)}
                        </span>
                      </span>
                      <button
                        className="ai-attachment-thumb-remove"
                        onClick={() => onRemoveAttachment(attachment.path)}
                        title={t('aiRemoveAttachment')}
                        aria-label={t('aiRemoveAttachment')}
                      >
                        <svg width="16" height="16" viewBox="0 0 32 32" aria-hidden>
                          <path
                            d="M24 9.4L22.6 8L16 14.6L9.4 8L8 9.4l6.6 6.6L8 22.6L9.4 24l6.6-6.6l6.6 6.6l1.4-1.4l-6.6-6.6L24 9.4z"
                            fill="currentColor"
                            stroke="currentColor"
                            strokeWidth="0.25"
                          />
                        </svg>
                      </button>
                    </span>
                  ),
                )}
              </div>
            )
          }
          disabled={!isReady}
          value={prompt}
          busy={aiBusy}
          placeholder={t(hasContent ? 'aiComposerPlaceholder' : 'aiComposerPlaceholderBuild')}
          hintIdle={t('aiHintIdle')}
          hintBusy={t('aiHintBusy')}
          hintIdleTitle={t('aiHintIdleTitle')}
          sendLabel={t('aiSend')}
          stopLabel={t('aiStop')}
          ariaLabel={t('aiInstructionAria')}
          iconOnly
          footerStart={
            <>
              <button
                className="ai-attach-btn"
                onClick={onPickAttachments}
                title={t('aiAttachTitle')}
              >
                <IconAttach size={18} />
              </button>
            </>
          }
          textareaRef={inputRef}
          onChange={onPromptChange}
          onSend={send}
          onStop={onStop}
          onPasteFiles={onPasteFiles}
        />
      </div>
    </AiPanelShell>
  )
}

const MAX_PREVIEW_CELL_ROWS = 50

function formatCell(
  cell: { readonly value: unknown; readonly formula?: string | undefined },
  t: TFunc,
): string {
  if (cell.formula) return cell.formula
  if (cell.value === null) return t('aiCellEmpty')
  return String(cell.value)
}
