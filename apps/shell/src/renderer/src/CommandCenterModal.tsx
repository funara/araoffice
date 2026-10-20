import { useEffect, useMemo, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { IconArrowUp, IconStop, Markdown } from '@araoffice/ui'
import type { RecentEntry } from '../../shared/home-api'
import type { ActiveTabContext, TabSummary } from '../../shared/tabs-api'
import { LANGUAGE_OPTIONS, useI18n } from './locale'
import { appendFollowUp, cancelOmnibarStream, executeOmnibarQuery } from './ai/omnibar-service'
import {
  HANDOVER_TRIGGER_REGEX,
  actionableHandoverGoal,
  fallbackHandover,
  parseExplicitHandoverApp,
  synthesizeHandover,
  type TargetAppKind,
} from './ai/handover'
import './command-center.css'
import { FileBadge } from './Home'

type HandoverState =
  | { status: 'idle' }
  | { status: 'busy' }
  | {
      status: 'ready'
      prompt: string
      editors: TargetAppKind[]
      outline: Array<{ label: string; depth: number }>
    }

interface CommandCenterModalProps {
  isOpen: boolean
  isDarkTheme: boolean
  onClose: () => void
  onToggleTheme: () => void
  onOpenFeedback: () => void
  onOpenFeatures?: () => void
  onOpenShortcuts?: () => void
  onOpenDonation?: () => void
}

interface CommandItem {
  id: string
  title: string
  category: string
  shortcut?: string
  icon: ReactElement
  action: () => void
}

interface SelectableItem {
  type: 'ai' | 'tab' | 'command' | 'file'
  aiPrompt?: string
  tab?: TabSummary
  command?: CommandItem
  file?: RecentEntry
}

interface OmnibarChatMessage {
  id: string
  role: 'user' | 'assistant'
  text: string
  isStreaming?: boolean
  error?: string
  followUps?: string[]
  timestamp: number
}

// Morphing ARA spark: single inline SVG so the 4-point star can scale/rotate via CSS.
// Idle: slow subtle pulse. Live (streaming/synthesizing): faster pronounced morph.
function AraSpark({ live }: { live?: boolean }) {
  return (
    <svg
      className={`cmd-spark${live ? ' cmd-spark-live' : ''}`}
      viewBox="0 0 16 16"
      aria-hidden="true"
    >
      <path
        d="M8 0c.7 4.5 3.5 7.3 8 8-4.5.7-7.3 3.5-8 8-.7-4.5-3.5-7.3-8-8 4.5-.7 7.3-3.5 8-8z"
        fill="currentColor"
      />
    </svg>
  )
}

function AraDynamicIcon({ live }: { live?: boolean }) {
  const [phase, setPhase] = useState<'star' | 'mascot'>('star')

  useEffect(() => {
    if (live) return
    const timer = setInterval(() => {
      setPhase((p) => (p === 'star' ? 'mascot' : 'star'))
    }, 3800)
    return () => clearInterval(timer)
  }, [live])

  return (
    <span
      className={`cmd-morph-wrap ${phase === 'mascot' ? 'show-mascot' : 'show-star'}${live ? ' is-live' : ''}`}
      aria-hidden="true"
    >
      <span className="cmd-morph-star">
        <AraSpark live={live} />
      </span>
      <img className="cmd-morph-mascot" src="./ARA_Assistant.png" alt="" width="28" height="28" />
    </span>
  )
}

// Streaming indicator: three dots while the answer streams in with no text yet.
// Model <think> blocks are never surfaced in the omnibar — the user sees the answer only.
function OmnibarStreamingDots({
  isStreaming,
  hasText,
}: {
  isStreaming?: boolean
  hasText?: boolean
}) {
  if (isStreaming && !hasText) {
    return (
      <div className="cmd-thinking cmd-thinking-live" role="status" aria-label="Thinking">
        <span className="cmd-thinking-dots" aria-hidden="true">
          <span />
          <span />
          <span />
        </span>
      </div>
    )
  }
  return null
}
// TabKind -> file ext for FileBadge (keyed by extension, not kind).
const TAB_KIND_EXT: Record<TabSummary['kind'], string> = {
  home: '',
  docs: 'docx',
  sheets: 'xlsx',
  slides: 'pptx',
  pdf: 'pdf',
  markdown: 'md',
  diagram: 'drawio',
}
// ponytail: short EN labels + tile colors, same language as Home badges;
// HANDOVER_LABELS precedent (EN-only). Localize with the dock when shipping.
const OPEN_TILE: Record<TargetAppKind, { letter: string; bg: string; short: string }> = {
  docs: { letter: 'W', bg: '#3276CD', short: 'Docs' },
  sheets: { letter: 'X', bg: '#4FA16B', short: 'Sheets' },
  slides: { letter: 'S', bg: '#D33922', short: 'Slides' },
  markdown: { letter: 'N', bg: '#7C3AED', short: 'Notes' },
  diagram: { letter: 'D', bg: '#d97706', short: 'Diagram' },
}

// Open split: top-1 pick is the button, the rest live under ▾. Choosing from
// the menu promotes that editor to the button (rerank) and opens it.
// Remount per prompt (key={prompt}) so the rank resets with each Generate.
function OpenSplit({
  editors,
  onOpen,
  isPrimaryCard = false,
}: {
  editors: TargetAppKind[]
  onOpen: (kind: TargetAppKind) => void
  isPrimaryCard?: boolean
}) {
  const [top, setTop] = useState<TargetAppKind>(editors[0] ?? 'docs')
  const [menuOpen, setMenuOpen] = useState(false)
  const wrapRef = useRef<HTMLSpanElement>(null)
  const chevRef = useRef<HTMLButtonElement>(null)
  const ordered = useMemo(() => {
    const allKnown: TargetAppKind[] = ['docs', 'sheets', 'slides', 'markdown', 'diagram']
    const combined = [...new Set([...editors, ...allKnown])]
    if (!combined.includes(top)) return combined
    return [top, ...combined.filter((e) => e !== top)]
  }, [editors, top])

  useEffect(() => {
    if (!menuOpen) return
    const onDocClick = (e: MouseEvent) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) setMenuOpen(false)
    }
    document.addEventListener('click', onDocClick)
    return () => document.removeEventListener('click', onDocClick)
  }, [menuOpen])

  const closeMenu = () => {
    setMenuOpen(false)
    chevRef.current?.focus()
  }

  const primary = ordered[0] ?? 'docs'
  return (
    <span className={`cmd-openwrap${isPrimaryCard ? ' cmd-openwrap-card' : ''}`} ref={wrapRef}>
      <span className="cmd-split">
        <button type="button" className="cmd-split-primary" onClick={() => onOpen(primary)}>
          <span
            className="cmd-tile cmd-open-tile"
            style={{ backgroundColor: OPEN_TILE[primary].bg }}
            aria-hidden="true"
          >
            {OPEN_TILE[primary].letter}
          </span>
          <span>Open in {OPEN_TILE[primary].short}</span>
          {isPrimaryCard && <span className="cmd-split-badge">Handover Draft →</span>}
        </button>
        <button
          ref={chevRef}
          type="button"
          className="cmd-split-chev"
          aria-haspopup="menu"
          aria-expanded={menuOpen}
          aria-label="Choose another editor"
          onClick={(e) => {
            e.stopPropagation()
            if (menuOpen) {
              setMenuOpen(false)
            } else {
              setMenuOpen(true)
              setTimeout(
                () => wrapRef.current?.querySelector<HTMLButtonElement>('.cmd-open-item')?.focus(),
                25,
              )
            }
          }}
        >
          ▾
        </button>
      </span>
      {menuOpen && (
        <span className="cmd-open-menu" role="menu">
          <span className="cmd-open-cap">Open with — best first</span>
          {ordered.map((kind, i) => (
            <button
              key={kind}
              type="button"
              role="menuitem"
              className="cmd-open-item"
              onClick={() => {
                setTop(kind)
                setMenuOpen(false)
                onOpen(kind)
              }}
              onKeyDown={(e) => {
                const items = Array.from(
                  wrapRef.current?.querySelectorAll<HTMLButtonElement>('.cmd-open-item') ?? [],
                )
                const at = items.indexOf(e.currentTarget)
                if (e.key === 'ArrowDown') {
                  e.preventDefault()
                  ;(items[at + 1] ?? items[0])?.focus()
                } else if (e.key === 'ArrowUp') {
                  e.preventDefault()
                  ;(items[at - 1] ?? items[items.length - 1])?.focus()
                } else if (e.key === 'Escape') {
                  e.preventDefault()
                  closeMenu()
                }
              }}
            >
              <span
                className="cmd-tile cmd-open-tile"
                style={{ backgroundColor: OPEN_TILE[kind].bg }}
                aria-hidden="true"
              >
                {OPEN_TILE[kind].letter}
              </span>
              <span>{i === 0 ? `Open in ${OPEN_TILE[kind].short}` : OPEN_TILE[kind].short}</span>
              {i === 0 && <small>Best match</small>}
            </button>
          ))}
        </span>
      )}
    </span>
  )
}

function HandoverActionCard({
  status,
  editors,
  onOpen,
  activeContext,
  lang,
}: {
  status: 'busy' | 'ready'
  editors: TargetAppKind[]
  onOpen: (kind: TargetAppKind) => void
  activeContext: ActiveTabContext
  lang: string
}) {
  const primary = editors[0] ?? 'docs'
  const isIndo = lang.startsWith('id')
  const isHome = activeContext.kind === 'home'
  const tile = OPEN_TILE[primary]

  const title = isIndo ? `Siap disusun di ARA ${tile.short}` : `Ready to build in ARA ${tile.short}`

  const subtitle =
    !isHome && activeContext.kind === primary
      ? isIndo
        ? `Menulis draf langsung ke dalam ${activeContext.title}`
        : `Writes the draft directly into ${activeContext.title}`
      : isIndo
        ? `Membuka tab ${tile.short} baru dan langsung menyusun dokumen berdasarkan brief ini.`
        : `Opens a new ${tile.short} tab and automatically drafts the structure from this brief.`

  if (status === 'busy') {
    return (
      <div className="cmd-handover-card cmd-handover-card-busy">
        <div className="cmd-handover-badge-row">
          <span className="cmd-synthesizing-spinner" aria-hidden="true" />
          <span className="cmd-handover-title">
            {isIndo ? 'Menyiapkan draft editor…' : 'Preparing editor draft…'}
          </span>
        </div>
        <div className="cmd-handover-desc">
          {isIndo
            ? 'Menyintesis brief dan memilih editor terbaik…'
            : 'Synthesizing brief and selecting the best editor…'}
        </div>
      </div>
    )
  }

  return (
    <div className="cmd-handover-card">
      <div className="cmd-handover-badge-row">
        <span
          className="cmd-tile cmd-open-tile"
          style={{ backgroundColor: tile.bg }}
          aria-hidden="true"
        >
          {tile.letter}
        </span>
        <span className="cmd-handover-title">{title}</span>
      </div>
      <div className="cmd-handover-action-row">
        <OpenSplit editors={editors} onOpen={onOpen} isPrimaryCard />
      </div>
      <div className="cmd-handover-desc">{subtitle}</div>
    </div>
  )
}

function parentDirName(path: string): string {
  const parts = path.split(/[\\/]/).filter(Boolean)
  return parts.length > 1 ? (parts[parts.length - 2] ?? '') : ''
}

function getDisplayableAiText(text: string): string {
  if (!text) return ''
  return (
    text
      .replace(
        /(?:>\s*)?(?:\*\*Target Editor Prompt:\*\*|Target Editor Prompt:)[^\n\r]*(?:\r?\n[^\n\r]*)*/gi,
        '',
      )
      .replace(/:::followups([\s\S]*?)(?:\s:::[ \t]*(?:\r?\n|$)|$)/gi, '')
      // Drop orphan fence remnants the model sometimes leaves behind: a lone
      // language tag line ("code") or an empty block leaves distracting chrome
      // (header + copy button) with nothing inside.
      .replace(/^[ \t]*```[a-zA-Z0-9_-]*[ \t]*\r?\n(?=\s*(?:```|$))/gm, '')
      .replace(/^[ \t]*```[ \t]*\r?$/gm, '')
      .replace(/([.!?])([A-Z\u00C0-\u024F\u4E00-\u9FFF])/g, '$1 $2')
      .trim()
  )
}

export function CommandCenterModal({
  isOpen,
  onClose,
  onToggleTheme,
  onOpenFeedback,
  onOpenFeatures,
  onOpenShortcuts,
  onOpenDonation,
}: CommandCenterModalProps) {
  const { lang, setLang, t } = useI18n()
  const [query, setQuery] = useState('')
  const [preAiQuery, setPreAiQuery] = useState('')
  const [selectedIndex, setSelectedIndex] = useState(0)
  const [openTabs, setOpenTabs] = useState<TabSummary[]>([])
  const [recentMatches, setRecentMatches] = useState<RecentEntry[]>([])
  const [activeContext, setActiveContext] = useState<ActiveTabContext>({
    id: 'home',
    kind: 'home',
    title: 'Home',
  })
  const [chatMessages, setChatMessages] = useState<OmnibarChatMessage[]>([])
  const [isAiMode, setIsAiMode] = useState(false)
  const [toast, setToast] = useState<string | null>(null)
  const [handover, setHandover] = useState<HandoverState>({ status: 'idle' })
  const handoverSeq = useRef(0)
  const [visible, setVisible] = useState(isOpen)
  const [closing, setClosing] = useState(false)

  // Open: show instantly. Close: hold 130ms for the exit fade, then unmount.
  useEffect(() => {
    if (isOpen) {
      setVisible(true)
      setClosing(false)
      return
    }
    if (!visible) return
    setClosing(true)
    const timer = setTimeout(() => {
      setVisible(false)
      setClosing(false)
    }, 130)
    return () => clearTimeout(timer)
  }, [isOpen, visible])

  const isAiModeRef = useRef(isAiMode)
  isAiModeRef.current = isAiMode
  const chatMessagesRef = useRef(chatMessages)
  chatMessagesRef.current = chatMessages
  const handoverRef = useRef(handover)
  handoverRef.current = handover

  const inputRef = useRef<HTMLInputElement>(null)
  const chatInputRef = useRef<HTMLTextAreaElement>(null)
  const listRef = useRef<HTMLDivElement>(null)
  const chatScrollRef = useRef<HTMLDivElement>(null)

  const showToast = (msg: string) => {
    setToast(msg)
    setTimeout(() => setToast(null), 2500)
  }

  // Cancelled streams never resolve (preload drops the listener on cancel):
  // freeze bubbles so the streaming dots unmount instead of sticking.
  const cancelAndFreezeStreaming = () => {
    cancelOmnibarStream()
    setChatMessages((prev) => prev.map((m) => (m.isStreaming ? { ...m, isStreaming: false } : m)))
  }

  // Load active context & tabs on modal open
  useEffect(() => {
    if (isOpen) {
      setSelectedIndex(0)
      setToast(null)

      if (window.aiOfficeTabs) {
        void window.aiOfficeTabs.getActiveContext().then(setActiveContext)
        void window.aiOfficeTabs.list().then((tabs) => {
          setOpenTabs(tabs.filter((t) => t.kind !== 'home'))
        })
      }

      if (isAiModeRef.current && chatMessagesRef.current.length > 0) {
        setQuery('')
        setTimeout(() => chatInputRef.current?.focus(), 25)
      } else {
        setIsAiMode(false)
        setQuery('')
        setPreAiQuery('')
      }
    } else {
      cancelAndFreezeStreaming()
    }
  }, [isOpen])

  // Focus appropriate input on mode switch
  useEffect(() => {
    if (!isOpen) return
    if (isAiMode) {
      setTimeout(() => chatInputRef.current?.focus(), 30)
    } else {
      setTimeout(() => inputRef.current?.focus(), 30)
    }
  }, [isAiMode, isOpen])

  // Auto-scroll chat view to bottom as tokens/messages arrive
  useEffect(() => {
    if (isAiMode && chatScrollRef.current) {
      chatScrollRef.current.scrollTop = chatScrollRef.current.scrollHeight
    }
  }, [chatMessages, isAiMode])

  // Search recent files
  useEffect(() => {
    const trimmed = query.trim().toLowerCase()
    let cancelled = false
    if (!window.aiOffice?.recents) return

    void window.aiOffice.recents({ limit: 40 }).then((page) => {
      if (cancelled) return
      if (trimmed) {
        const matches = page.entries.filter(
          (e) => e.name.toLowerCase().includes(trimmed) || e.path.toLowerCase().includes(trimmed),
        )
        setRecentMatches(matches.slice(0, 8))
      } else {
        setRecentMatches(page.entries.slice(0, 4))
      }
    })
    return () => {
      cancelled = true
    }
  }, [query])

  const canGenerateHandover = chatMessages.some(
    (m) => m.role === 'assistant' && !m.isStreaming && m.text?.trim() && !m.error,
  )

  // Handover open: prompt already synthesized, content always empty.
  const handleHandoverOpen = async (
    kind: TargetAppKind,
    override?: { prompt: string; outline?: Array<{ label: string; depth: number }> },
  ) => {
    const currentHandover = handoverRef.current
    if (currentHandover.status !== 'ready' && !override) return
    const targetPrompt =
      override?.prompt ?? (currentHandover.status === 'ready' ? currentHandover.prompt : '')
    const targetOutline =
      override?.outline ??
      (currentHandover.status === 'ready' ? currentHandover.outline : undefined)
    onClose()
    try {
      await window.aiOfficeTabs.dispatchAction({
        type: 'create-doc',
        kind,
        content: '',
        prompt: targetPrompt,
        summary: targetPrompt.slice(0, 200),
        ...(kind === 'diagram' && targetOutline
          ? { diagramMode: 'visual' as const, outline: targetOutline }
          : {}),
      })
    } catch (err) {
      console.error('[CommandCenter] Failed to open handover tab:', err)
    }
  }

  // Thread-level handover: one Generate call synthesizes prompt + editors.
  const generateHandover = (messages?: OmnibarChatMessage[]) => {
    const msgs = Array.isArray(messages) ? messages : chatMessagesRef.current
    if (handoverRef.current.status === 'busy') return
    const seq = handoverSeq.current + 1
    handoverSeq.current = seq
    setHandover({ status: 'busy' })
    void synthesizeHandover(msgs, lang).then((res) => {
      if (handoverSeq.current !== seq) return
      if (res && res.prompt && res.editors.length > 0) {
        setHandover({
          status: 'ready',
          prompt: res.prompt,
          editors: res.editors,
          outline: res.outline,
        })
      } else {
        setHandover({ status: 'idle' })
      }
    })
  }

  // Trigger Omnibar AI query (supports continuing conversation thread)
  const triggerAiQuery = (userPrompt: string, continueThread = false) => {
    const trimmed = userPrompt.trim()
    if (!trimmed) return
    if (chatInputRef.current) chatInputRef.current.style.height = ''

    // Fast Conversational Handover Trigger:
    if (continueThread && HANDOVER_TRIGGER_REGEX.test(trimmed)) {
      const explicitApp = parseExplicitHandoverApp(trimmed)
      const currentHandover = handoverRef.current
      if (currentHandover.status === 'ready') {
        const targetKind =
          explicitApp && currentHandover.editors.includes(explicitApp)
            ? explicitApp
            : (currentHandover.editors[0] ?? 'docs')
        setQuery('')
        void handleHandoverOpen(targetKind)
        return
      }
      if (currentHandover.status === 'busy' || canGenerateHandover) {
        const fb = fallbackHandover(chatMessagesRef.current, lang)
        if (fb) {
          const targetKind =
            explicitApp && fb.editors.includes(explicitApp)
              ? explicitApp
              : (fb.editors[0] ?? 'docs')
          setHandover({
            status: 'ready',
            prompt: fb.prompt,
            editors: fb.editors,
            outline: [],
          })
          setQuery('')
          void handleHandoverOpen(targetKind, { prompt: fb.prompt })
          return
        }
      }
    }

    // Fresh entry from command search: stash the query so Back/Esc restores it.
    if (!continueThread && !isAiModeRef.current) setPreAiQuery(trimmed)

    // Freeze an in-flight answer instead of dropping it: the old filter
    // removed every streaming bubble, so the previous AI response vanished
    // when the user sent a follow-up mid-stream.
    const updatedThread = (
      continueThread ? appendFollowUp(chatMessages, trimmed) : []
    ) as OmnibarChatMessage[]
    const userMsgId = `user-${Date.now()}`
    if (continueThread) {
      const userMsg = updatedThread[updatedThread.length - 1] as OmnibarChatMessage
      userMsg.id = userMsgId
      userMsg.timestamp = Date.now()
    } else {
      updatedThread.push({ id: userMsgId, role: 'user', text: trimmed, timestamp: Date.now() })
    }

    const assistantMsgId = `assistant-${Date.now() + 1}`
    const assistantMsg: OmnibarChatMessage = {
      id: assistantMsgId,
      role: 'assistant',
      text: '',
      isStreaming: true,
      timestamp: Date.now() + 1,
    }
    setChatMessages([...updatedThread, assistantMsg])
    handoverSeq.current += 1
    setHandover({ status: 'idle' })
    setIsAiMode(true)
    setQuery('') // Clear input so user is immediately ready for follow-up

    const messagesPayload = updatedThread.map((m) => ({
      role: m.role,
      text: m.text,
    }))

    void executeOmnibarQuery(messagesPayload, activeContext, lang, {
      onUpdate(res) {
        setChatMessages((prev) =>
          prev.map((msg) =>
            msg.id === assistantMsgId
              ? {
                  ...msg,
                  text: res.text,
                  isStreaming: res.isStreaming,
                  error: res.error,
                  followUps: res.followUps,
                }
              : msg,
          ),
        )
      },
      onError(err) {
        setChatMessages((prev) =>
          prev.map((msg) =>
            msg.id === assistantMsgId
              ? {
                  ...msg,
                  text: '',
                  isStreaming: false,
                  error: err,
                }
              : msg,
          ),
        )
      },
      onDone(res) {
        const updatedWithAssistant: OmnibarChatMessage[] = [
          ...updatedThread,
          { id: assistantMsgId, role: 'assistant', text: res.text, timestamp: Date.now() },
        ]
        setChatMessages((prev) =>
          prev.map((msg) =>
            msg.id === assistantMsgId
              ? {
                  ...msg,
                  text: res.text,
                  isStreaming: false,
                  followUps: res.followUps,
                }
              : msg,
          ),
        )
        if (actionableHandoverGoal(updatedWithAssistant)) {
          generateHandover(updatedWithAssistant)
        }
      },
    })
  }

  // Stop covers both streaming answers and handover synthesis: the synth
  // promise has no cancel, so bump the seq to drop its late resolve.
  const stopAll = () => {
    cancelAndFreezeStreaming()
    handoverSeq.current += 1
    if (handoverRef.current.status === 'busy') setHandover({ status: 'idle' })
  }

  const backToCommands = () => {
    cancelAndFreezeStreaming()
    setIsAiMode(false)
    setQuery(preAiQuery)
    setTimeout(() => inputRef.current?.focus(), 25)
  }

  const newConversation = () => {
    cancelOmnibarStream()
    handoverSeq.current += 1
    setHandover({ status: 'idle' })
    setChatMessages([])
    setIsAiMode(false)
    setQuery('')
    setPreAiQuery('')
  }

  // Base shell commands
  // icon is data: single-letter tile in the format color (same badge
  // language as Home). System commands share one neutral tile so the
  // list speaks a single icon language — never mixed badges + outlines.
  const baseCommands: CommandItem[] = useMemo(() => {
    const formatTile = (label: string, bg: string) => (
      <span className="cmd-tile" style={{ backgroundColor: bg }} aria-hidden="true">
        {label}
      </span>
    )
    const sysTile = (glyph: string) => (
      <span className="cmd-tile cmd-tile-sys" aria-hidden="true">
        {glyph}
      </span>
    )
    const commands: CommandItem[] = [
      {
        id: 'new-docx',
        title: t('cmdNewDocx'),
        category: 'Create Document',
        shortcut: 'Alt+W',
        icon: formatTile('W', '#3276CD'),
        action: () => {
          onClose()
          void window.aiOffice.newDoc()
        },
      },
      {
        id: 'new-xlsx',
        title: t('cmdNewXlsx'),
        category: 'Create Document',
        shortcut: 'Alt+X',
        icon: formatTile('X', '#4FA16B'),
        action: () => {
          onClose()
          void window.aiOffice.newSheet()
        },
      },
      {
        id: 'new-pptx',
        title: t('cmdNewPptx'),
        category: 'Create Document',
        shortcut: 'Alt+P',
        icon: formatTile('S', '#D33922'),
        action: () => {
          onClose()
          void window.aiOffice.newSlide()
        },
      },
      {
        id: 'new-md',
        title: t('cmdNewMd'),
        category: 'Create Document',
        shortcut: 'Alt+M',
        icon: formatTile('N', '#7C3AED'),
        action: () => {
          onClose()
          void window.aiOffice.newMarkdown()
        },
      },
      {
        id: 'new-diagram',
        title: t('cmdNewDiagram'),
        category: 'Create Document',
        shortcut: 'Alt+D',
        icon: formatTile('D', '#d97706'),
        action: () => {
          onClose()
          void window.aiOffice.newDiagram()
        },
      },
      {
        id: 'open-file',
        title: t('cmdOpenFile'),
        category: 'File Operations',
        shortcut: 'Ctrl+O',
        icon: sysTile('⤴'),
        action: () => {
          onClose()
          void window.aiOffice.browse()
        },
      },
      {
        id: 'open-pdf',
        title: t('cmdOpenPdf'),
        category: 'File Operations',
        icon: formatTile('P', '#dc2626'),
        action: () => {
          onClose()
          void window.aiOffice.openPdf()
        },
      },
      {
        id: 'switch-theme',
        title: t('cmdSwitchTheme'),
        category: t('cmdSectionSystem'),
        shortcut: t('theme'),
        icon: sysTile('◐'),
        action: () => {
          onToggleTheme()
          showToast(t('cmdThemeSwitched'))
        },
      },
      {
        id: 'switch-language',
        title: t('cmdSwitchLanguage'),
        category: t('cmdSectionSystem'),
        shortcut: t('language'),
        icon: sysTile('文'),
        action: () => {
          const next =
            LANGUAGE_OPTIONS[
              (LANGUAGE_OPTIONS.findIndex((o) => o.value === lang) + 1) % LANGUAGE_OPTIONS.length
            ]
          if (next) {
            setLang(next.value)
            showToast(next.label)
          }
        },
      },
      {
        id: 'check-updates',
        title: t('checkUpdates'),
        category: t('cmdSectionSystem'),
        shortcut: t('cmdShortcutUpdate'),
        icon: sysTile('↓'),
        action: () => {
          onClose()
          void window.aiOffice.checkForUpdates().then(() => showToast(t('cmdUpdateChecking')))
        },
      },
      {
        id: 'open-shortcuts',
        title: t('cmdShortcuts'),
        category: t('cmdSectionSystem'),
        shortcut: 'F1',
        icon: sysTile('⌨'),
        action: () => {
          onClose()
          onOpenShortcuts?.()
        },
      },
      {
        id: 'open-support',
        title: t('supportTitle'),
        category: t('cmdSectionSystem'),
        shortcut: t('cmdShortcutSupport'),
        icon: sysTile('♥'),
        action: () => {
          onClose()
          onOpenDonation?.()
        },
      },
      {
        id: 'open-feedback',
        title: t('cmdFeedback'),
        category: t('cmdSectionSystem'),
        shortcut: 'Feedback',
        icon: sysTile('✎'),
        action: () => {
          onClose()
          onOpenFeedback()
        },
      },
      ...(onOpenFeatures
        ? [
            {
              id: 'open-features',
              title: t('changelogTitle'),
              category: t('cmdSectionSystem'),
              shortcut: t('cmdShortcutNew'),
              icon: sysTile('✦'),
              action: () => {
                onClose()
                onOpenFeatures()
              },
            },
          ]
        : []),
    ]
    return commands
  }, [
    onClose,
    onToggleTheme,
    onOpenFeedback,
    onOpenFeatures,
    onOpenShortcuts,
    onOpenDonation,
    lang,
    setLang,
    t,
  ])

  // Filter commands
  const filteredCommands = useMemo(() => {
    const trimmed = query
      .trim()
      .toLowerCase()
      .replace(/^[>/]\s*/, '')
    if (!trimmed) return baseCommands
    return baseCommands.filter(
      (c) =>
        c.title.toLowerCase().includes(trimmed) ||
        c.category.toLowerCase().includes(trimmed) ||
        c.shortcut?.toLowerCase().includes(trimmed),
    )
  }, [query, baseCommands])

  // Filter matching open tabs
  const matchingTabs = useMemo(() => {
    const trimmed = query.trim().toLowerCase()
    if (!trimmed) return openTabs
    return openTabs.filter((t) => t.title.toLowerCase().includes(trimmed))
  }, [query, openTabs])

  const CATEGORIES = useMemo(
    () => [
      { key: 'Create Document', label: t('cmdSectionCreate') },
      { key: 'File Operations', label: t('cmdSectionFiles') },
      { key: t('cmdSectionSystem'), label: t('cmdSectionSystem') },
    ],
    [t],
  )

  // A leading > or / narrows the list to commands only, mirroring
  // classic palettes. Tab is intentionally NOT a mode switcher.
  const effectiveMode = /^[/>]/.test(query.trim()) ? 'cmd' : 'search'

  // Build selectable items list
  const allItems: SelectableItem[] = useMemo(() => {
    const items: SelectableItem[] = []
    const trimmed = query.trim()
    const commandsOnly = effectiveMode === 'cmd'

    // Top item: ask the assistant if the user typed something
    if (trimmed && !trimmed.startsWith('>') && !trimmed.startsWith('/') && !commandsOnly) {
      items.push({ type: 'ai', aiPrompt: trimmed })
    }

    if (!commandsOnly) {
      for (const tab of matchingTabs) {
        items.push({ type: 'tab', tab })
      }
    }
    for (const cat of CATEGORIES) {
      const catCmds = filteredCommands.filter((c) => c.category === cat.key)
      for (const cmd of catCmds) {
        items.push({ type: 'command', command: cmd })
      }
    }
    if (!commandsOnly) {
      for (const file of recentMatches) {
        items.push({ type: 'file', file })
      }
    }
    return items
  }, [query, effectiveMode, matchingTabs, filteredCommands, recentMatches, CATEGORIES])

  const totalItems = allItems.length

  useEffect(() => {
    if (selectedIndex >= totalItems) {
      setSelectedIndex(0)
    }
  }, [selectedIndex, totalItems])

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault()
      if (isAiMode) {
        const isAnyBusy =
          chatMessages.some((m) => m.isStreaming) || handoverRef.current.status === 'busy'
        if (isAnyBusy) {
          stopAll()
        } else {
          setIsAiMode(false)
          setChatMessages([])
          handoverSeq.current += 1
          setHandover({ status: 'idle' })
          setQuery(preAiQuery)
        }
      } else {
        onClose()
      }
    } else if (e.key === 'ArrowDown') {
      if (isAiMode) return
      e.preventDefault()
      setSelectedIndex((prev) => (totalItems > 0 ? (prev + 1) % totalItems : 0))
    } else if (e.key === 'ArrowUp') {
      if (isAiMode) return
      e.preventDefault()
      setSelectedIndex((prev) => (totalItems > 0 ? (prev - 1 + totalItems) % totalItems : 0))
    } else if (e.altKey && (e.key === 'a' || e.key === 'A')) {
      if (isAiMode) return
      e.preventDefault()
      if (query.trim()) triggerAiQuery(query.trim())
      else {
        setIsAiMode(true)
        setQuery('')
      }
      return
    } else if (e.key === 'Enter') {
      // Chat textarea: Shift+Enter is a newline, plain Enter sends.
      if (isAiMode && e.shiftKey) return
      e.preventDefault()
      if (isAiMode) {
        if (query.trim()) {
          triggerAiQuery(query.trim(), true)
        }
        return
      }

      const item = allItems[selectedIndex]
      if (!item) return
      if (item.type === 'ai' && item.aiPrompt) {
        triggerAiQuery(item.aiPrompt)
      } else if (item.type === 'tab' && item.tab) {
        onClose()
        void window.aiOfficeTabs.activate(item.tab.id)
      } else if (item.type === 'command' && item.command) {
        item.command.action()
      } else if (item.type === 'file' && item.file) {
        onClose()
        void window.aiOffice.openPath(item.file.path)
      }
    }
  }

  if (!visible) return null

  const tabBadgeExt = (tab: TabSummary) => {
    if (tab.kind === 'diagram') return tab.diagramMode === 'mermaid' ? 'mmd' : 'drawio'
    return TAB_KIND_EXT[tab.kind] ?? tab.kind
  }
  const renderFileBadge = (ext: string) => <FileBadge ext={ext.toLowerCase()} size={24} />

  const isAiStreaming = chatMessages.some((m) => m.isStreaming)
  // ponytail: composer hints hardcoded EN (this file already hardcodes Send /
  // back labels); locale tables when shipping beyond EN.
  const composerHint =
    handover.status === 'busy'
      ? 'Generating…'
      : isAiStreaming
        ? 'Thinking…'
        : handover.status === 'ready'
          ? 'Ready to open'
          : null

  return (
    <div
      className={`cmd-backdrop${closing ? ' cmd-closing' : ''}`}
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose()
      }}
    >
      <div
        className={`cmd-dialog${isAiMode ? ' cmd-dialog-chat' : ''}`}
        role="dialog"
        aria-modal="true"
        aria-label="Command Center"
      >
        {isAiMode ? (
          <>
            <div className="cmd-chat-header">
              <button
                className="cmd-chat-back"
                type="button"
                onClick={backToCommands}
                title="Back to command search"
              >
                ← Commands
              </button>
              <span className="cmd-chat-sep" aria-hidden="true" />
              <span className="cmd-chat-brand">
                <span className="cmd-tile cmd-tile-ai" aria-hidden="true">
                  <AraSpark live={isAiStreaming || handover.status === 'busy'} />
                </span>
                ARA Assistant
              </span>
              {activeContext.id !== 'home' && (
                <span className="cmd-crumb" title={`Active: ${activeContext.title}`}>
                  {activeContext.title}
                </span>
              )}
              <span className="cmd-chat-header-right">
                <button
                  className="cmd-chat-new"
                  type="button"
                  onClick={newConversation}
                  title="Start a new conversation"
                >
                  + New
                </button>
                <span className="cmd-chat-sep" aria-hidden="true" />
                <button
                  className="cmd-chat-close"
                  type="button"
                  onClick={onClose}
                  title="Close (Esc)"
                  aria-label="Close"
                >
                  ✕
                </button>
              </span>
            </div>
            {/* AI Streamed Conversation View (Scrollable Middle) */}
            <div className="cmd-ai-view" ref={chatScrollRef}>
              {chatMessages.map((msg, index) => {
                if (msg.role === 'user') {
                  return (
                    <div key={msg.id || index} className="cmd-ai-bubble-user-wrap">
                      <div className="cmd-ai-bubble-user">{msg.text}</div>
                    </div>
                  )
                }

                return (
                  <div key={msg.id || index} className="cmd-ai-card">
                    <div className="cmd-ai-card-header">
                      <span className="cmd-ai-author-badge">
                        <span className="cmd-ai-author-icon" aria-hidden="true">
                          <AraSpark live={msg.isStreaming} />
                        </span>
                        <span className="cmd-ai-author-name">ARA Assistant</span>
                      </span>
                      {msg.isStreaming && <span className="cmd-ai-status-pulse">Thinking</span>}
                    </div>
                    <OmnibarStreamingDots
                      isStreaming={msg.isStreaming}
                      hasText={Boolean(msg.text?.trim())}
                    />
                    {msg.error ? (
                      <div className="cmd-msg-error-container">
                        <div style={{ color: '#ef4444', padding: '6px 0' }}>⚠️ {msg.error}</div>
                        {!msg.isStreaming && (
                          <button
                            type="button"
                            className="cmd-mini-btn"
                            onClick={() => {
                              const prevUser = [...chatMessages]
                                .reverse()
                                .find((m) => m.role === 'user' && m.text?.trim())
                              if (prevUser?.text) triggerAiQuery(prevUser.text, false)
                            }}
                          >
                            ↻ Retry
                          </button>
                        )}
                      </div>
                    ) : msg.text ? (
                      <div className="cmd-ai-content">
                        <Markdown text={getDisplayableAiText(msg.text)} />
                      </div>
                    ) : null}
                    {!msg.isStreaming &&
                      !msg.error &&
                      index === chatMessages.length - 1 &&
                      (handover.status === 'ready' || handover.status === 'busy') && (
                        <HandoverActionCard
                          status={handover.status}
                          editors={handover.status === 'ready' ? handover.editors : []}
                          onOpen={(kind) => void handleHandoverOpen(kind)}
                          activeContext={activeContext}
                          lang={lang}
                        />
                      )}
                  </div>
                )
              })}
            </div>

            <div className="cmd-chat-input-bar">
              <div className="cmd-chat-input-box">
                <textarea
                  ref={chatInputRef}
                  rows={1}
                  className="cmd-chat-input-field"
                  placeholder={
                    activeContext.id === 'home'
                      ? t('cmdChatPlaceholderGeneral')
                      : t('cmdChatPlaceholderContext', { title: activeContext.title })
                  }
                  value={query}
                  onChange={(e) => {
                    setQuery(e.target.value)
                    e.target.style.height = 'auto'
                    e.target.style.height = `${Math.min(e.target.scrollHeight, 140)}px`
                  }}
                  onKeyDown={handleKeyDown}
                />
                <div className="cmd-chat-actions">
                  {composerHint && <span className="cmd-hint">{composerHint}</span>}
                  {isAiStreaming || handover.status === 'busy' ? (
                    <button
                      className="cmd-send-btn cmd-send-stop"
                      type="button"
                      onClick={stopAll}
                      title="Stop generating (Esc)"
                      aria-label="Stop generating"
                    >
                      <IconStop size={16} />
                    </button>
                  ) : (
                    <button
                      className="cmd-send-btn"
                      type="button"
                      disabled={!query.trim()}
                      onClick={() => {
                        if (query.trim()) {
                          triggerAiQuery(query.trim(), true)
                        }
                      }}
                      title="Send message (Enter)"
                      aria-label="Send message"
                    >
                      <IconArrowUp size={18} />
                    </button>
                  )}
                </div>
              </div>
            </div>
          </>
        ) : (
          <>
            <div className="cmd-input-bar">
              <span
                className={`cmd-search-icon cmd-prompt-icon${effectiveMode === 'cmd' ? ' is-cmd-mode' : ''}${query.trim() ? ' has-query' : ''}`}
                aria-hidden="true"
              >
                &gt;_
              </span>
              <input
                ref={inputRef}
                type="text"
                className="cmd-input"
                placeholder={
                  // ponytail: hardcoded EN like composerHint; locale tables when shipping beyond EN.
                  effectiveMode === 'cmd' ? 'Type a command…' : 'Ask ARA, search, or run a command…'
                }
                value={query}
                onChange={(e) => {
                  setQuery(e.target.value)
                  setSelectedIndex(0)
                }}
                onKeyDown={handleKeyDown}
              />
              <button
                type="button"
                className={`cmd-ask-btn${query.trim() ? ' has-query' : ''}`}
                title={`${t('cmdAskAi')} (Alt+A)`}
                aria-label={t('cmdAskAi')}
                onClick={() => {
                  if (query.trim()) triggerAiQuery(query.trim())
                  else {
                    setIsAiMode(true)
                    setQuery('')
                  }
                }}
              >
                <AraDynamicIcon live={isAiStreaming} />
              </button>
            </div>
            <div className="cmd-list" ref={listRef} role="listbox">
              {/* Top Prompt: Ask AI */}
              {allItems.length > 0 && allItems[0].type === 'ai' && (
                <div className="cmd-group">
                  <div
                    className={`cmd-item${selectedIndex === 0 ? ' active' : ''}`}
                    role="option"
                    aria-selected={selectedIndex === 0}
                    onClick={() => triggerAiQuery(allItems[0].aiPrompt!)}
                    onMouseEnter={() => setSelectedIndex(0)}
                  >
                    <span className="cmd-tile cmd-tile-ai" aria-hidden="true">
                      <AraSpark />
                    </span>
                    <div className="cmd-file-info">
                      <span className="cmd-item-title">
                        Ask ARA Assistant: &ldquo;{allItems[0].aiPrompt}&rdquo;
                      </span>
                      <span className="cmd-file-path">
                        {t('cmdAskAboutContext', { title: activeContext.title })}
                      </span>
                    </div>
                    <span className="cmd-item-kbd">↵ {t('cmdAskAi')}</span>
                  </div>
                </div>
              )}

              {/* 1. Matching Open Tabs */}
              {matchingTabs.length > 0 && effectiveMode !== 'cmd' && (
                <div className="cmd-group">
                  <div className="cmd-section-title">{t('cmdOpenTabs')}</div>
                  {matchingTabs.map((tab) => {
                    const isSelected =
                      allItems.findIndex((i) => i.type === 'tab' && i.tab?.id === tab.id) ===
                      selectedIndex
                    return (
                      <div
                        key={`tab-${tab.id}`}
                        className={`cmd-item${isSelected ? ' active' : ''}`}
                        role="option"
                        aria-selected={isSelected}
                        onClick={() => {
                          onClose()
                          void window.aiOfficeTabs.activate(tab.id)
                        }}
                        onMouseEnter={() =>
                          setSelectedIndex(
                            allItems.findIndex((i) => i.type === 'tab' && i.tab?.id === tab.id),
                          )
                        }
                      >
                        {renderFileBadge(tabBadgeExt(tab))}
                        <span className="cmd-item-title">{tab.title}</span>
                        <span className="cmd-item-kbd">↵ switch</span>
                      </div>
                    )
                  })}
                </div>
              )}

              {/* 2. Commands Grouped by Function */}
              {CATEGORIES.map((cat) => {
                const catCmds = filteredCommands.filter((c) => c.category === cat.key)
                if (catCmds.length === 0) return null
                return (
                  <div key={cat.key} className="cmd-group">
                    <div className="cmd-section-title">{cat.label}</div>
                    {catCmds.map((cmd) => {
                      const isSelected =
                        allItems.findIndex(
                          (i) => i.type === 'command' && i.command?.id === cmd.id,
                        ) === selectedIndex
                      return (
                        <div
                          key={cmd.id}
                          className={`cmd-item${isSelected ? ' active' : ''}`}
                          role="option"
                          aria-selected={isSelected}
                          onClick={cmd.action}
                          onMouseEnter={() =>
                            setSelectedIndex(
                              allItems.findIndex(
                                (i) => i.type === 'command' && i.command?.id === cmd.id,
                              ),
                            )
                          }
                        >
                          <span className="cmd-item-icon">{cmd.icon}</span>
                          <span className="cmd-item-title">{cmd.title}</span>
                          {cmd.shortcut && <span className="cmd-item-kbd">{cmd.shortcut}</span>}
                        </div>
                      )
                    })}
                  </div>
                )
              })}

              {/* 3. Matching Recent Files */}
              {recentMatches.length > 0 && effectiveMode !== 'cmd' && (
                <div className="cmd-group">
                  <div className="cmd-section-title">
                    {query.trim() ? t('cmdMatchingFiles') : t('secRecent')}
                  </div>
                  {recentMatches.map((file) => {
                    const isSelected =
                      allItems.findIndex((i) => i.type === 'file' && i.file?.path === file.path) ===
                      selectedIndex
                    return (
                      <div
                        key={file.path}
                        className={`cmd-item${isSelected ? ' active' : ''}`}
                        role="option"
                        aria-selected={isSelected}
                        onClick={() => {
                          onClose()
                          void window.aiOffice.openPath(file.path)
                        }}
                        onMouseEnter={() =>
                          setSelectedIndex(
                            allItems.findIndex(
                              (i) => i.type === 'file' && i.file?.path === file.path,
                            ),
                          )
                        }
                      >
                        <span className="cmd-item-icon">{renderFileBadge(file.ext)}</span>
                        <div className="cmd-file-info">
                          <span className="cmd-item-title">{file.name}</span>
                          <span className="cmd-file-path">{parentDirName(file.path)}</span>
                        </div>
                        {file.starred ? (
                          <span className="cmd-starred-badge">★</span>
                        ) : (
                          <span className="cmd-item-kbd">file</span>
                        )}
                      </div>
                    )
                  })}
                </div>
              )}

              {totalItems === 0 && (
                <div className="cmd-empty">
                  <span className="cmd-empty-icon">∅</span>
                  <span>{t('cmdEmptySearch', { query })}</span>
                </div>
              )}
            </div>
          </>
        )}

        {/* Footer */}
        <div className="cmd-footer">
          <div className="cmd-footer-hints">
            {isAiMode ? (
              <>
                <span>
                  <kbd>↵</kbd> {t('cmdSendFollowUp')}
                </span>
                <span>
                  <kbd>esc</kbd> {t('cmdBackClose')}
                </span>
              </>
            ) : (
              <>
                <span>
                  <kbd>↑↓</kbd> {t('cmdNavigate')}
                </span>
                <span>
                  <kbd>↵</kbd> {t('cmdSelectAsk')}
                </span>
                <span>
                  <kbd>esc</kbd> {t('cmdBackClose')}
                </span>
              </>
            )}
          </div>
          <span className="cmd-footer-count">
            {isAiMode ? (
              'thread · ' + chatMessages.length + ' msgs'
            ) : (
              <>
                {totalItems} {totalItems === 1 ? 'result' : 'results'}
              </>
            )}
          </span>
        </div>
      </div>

      {/* Action Toast Feedback */}
      {toast && (
        <div className="cmd-toast">
          <span>✨</span>
          <span>{toast}</span>
        </div>
      )}
    </div>
  )
}
