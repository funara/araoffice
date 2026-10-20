import type React from 'react'
import { useEffect, useRef, useState } from 'react'
import {
  AgentLoop,
  composeSkills,
  CONTINUE_INSTRUCTION,
  type AgentImage,
  type ToolDisplay,
} from '@araoffice/agent-core'
import type { RenderSlide } from '@araoffice/pptx-render'
import type { AiSettings, AttachmentAddResult, AttachmentMeta } from '../../shared/ipc'
import { ATTACHMENT_IMAGE_EXTS } from '../../shared/ipc'
import {
  createSlidesSkill,
  type SlidesSkillDeps,
  type ClarifyQuestion,
  type DeckProgressEvent,
  type PageProgressItem,
} from './slides-skill'
import { createFilesSkill } from './files-skill'
import { createElectronTransport } from './transport'
import { renderSlidesToPngBase64 } from '../export-render'
import fileImageIcon from '../assets/file-image.png'
import { isQcEnabled, qcSlidePage, QC_MAX_PAGES } from './slide-qc'
import { useI18n, t as tGlobal, aiLangDirective, type TFunc } from '../i18n/locale'
import {
  AiPanelHeader,
  AiPanelShell,
  AraSpark,
  AiProcessCard,
  AiSnapshotsPopover,
  AiTypingIndicator,
  AiUnconfiguredCard,
  extractThinking,
  IconArrowUp,
  IconAttach,
  IconNewChat,
  IconRefresh,
  IconSidebarCollapseLeft,
  IconStop,
  Markdown,
  openExternalUrl,
  useAiPanelWidth,
  useAiReady,
} from '@araoffice/ui'

interface ToolActivity {
  name: string
  summary: string
  /** still executing: rendered as a spinner chip, replaced in place when the tool finishes */
  running?: boolean
  isError?: boolean
  /** Full tool output (truncated to 2000 chars) */
  output?: string
  /**
   * Side-channel display data: structured UI data filled by tools, not in LLM context.
   * Takes priority over name-based inference.
   */
  display?: ToolDisplay
}

/** Max stored chars of tool output (UI-side truncation, doesn't affect what the LLM receives) */
const TOOL_OUTPUT_MAX_CHARS = 2000
/** tools that build actual deck content — a run ending after only cosmetic
 * mutations (background, notes, styling) still counts as "deck not built" */
const DECK_CONTENT_TOOLS = new Set([
  'add_slide',
  'add_text_box',
  'add_shape',
  'add_chart',
  'add_smartart',
  'add_table',
  'edit_chart',
  'edit_table_cell',
  'edit_table_structure',
  'edit_table_style',
  'set_element_text',
  'execute_slide_script',
])
/** content-building continuation rounds fired after plan/clarify — bounded so a
 * weak model that stalls can't burn tokens forever */
const MAX_DECK_CONTINUATIONS = 8

/** Rollback point registered by the main process at the end of an AI run that edited the deck */
interface DeckSnapshot {
  id: number
  label: string
  time: string
}

/** Clipboard bitmap MIME → attachment extension (matches ATTACHMENT_IMAGE_EXTS) */
const PASTE_MIME_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
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

/** Cap on tool args/output persisted to the transcript (the store layer has another 16k truncation fallback) */
const PERSIST_TOOL_FIELD_MAX = 16_000

/** Tool args → JSON string (truncated; returns undefined on serialization failure, doesn't block persistence) */
function safeJsonInput(input: unknown): string | undefined {
  try {
    const s = JSON.stringify(input)
    return s && s !== '{}' ? s.slice(0, PERSIST_TOOL_FIELD_MAX) : undefined
  } catch {
    return undefined
  }
}

/** Generation progress snapshot in the chat stream (same card updated in real time) */
interface DeckProgressSnapshot {
  style?: { label: string; status: 'running' | 'done' | 'error'; summary: string }
  plan?: {
    label: string
    done: number
    total: number
    status: 'running' | 'done' | 'error'
    summary: string
  }
  images?: {
    label: string
    done: number
    total: number
    status: 'running' | 'done' | 'error'
    summary: string
  }
  pages?: {
    label: string
    done: number
    total: number
    status: 'running' | 'done' | 'error'
    summary: string
    items: PageProgressItem[]
  }
  finalTotal?: number // Total page count from the done event
  isDone?: boolean
}

interface ChatEntry {
  role: 'user' | 'assistant'
  text: string
  thinking?: string
  error?: string
  streaming?: boolean
  turnLimit?: boolean
  stuck?: boolean
  /** the run failed and this user message was rolled back out of the model context */
  undelivered?: boolean
  tools?: ToolActivity[]
  attachments?: AttachmentMeta[]
  /** Generation progress card (only one per turn, replaced in real time) */
  deckProgress?: DeckProgressSnapshot
}

/** Empty deck → generation starters; deck with content → polish starters */
const starterPrompts = (t: TFunc, deckEmpty: boolean): string[] =>
  deckEmpty
    ? [t('aiStarterGenReport'), t('aiStarterGenLaunch'), t('aiStarterGenTraining')]
    : [t('aiStarterPolishTitle'), t('aiStarterTighten'), t('aiStarterProofread')]

interface AiPanelProps {
  slides: RenderSlide[]
  current: number
  selectedIds: string[]
  /** no slide carries real content yet — the empty-state copy offers generation instead of polish */
  deckEmpty?: boolean
  /** Decoded image cache from App (SlideThumb needs it) — powers AI-vision slide screenshots */
  images: Map<string, HTMLImageElement>
  applySlide: (slideIndex: number, updated: RenderSlide) => void
  applyDeck: (slides: RenderSlide[], goTo?: number) => void
  fitWidthPx: number
  settings: AiSettings
  /** Preset instruction pushed from the ribbon/start screen; sent immediately when autoRun. When displayText exists the chat bubble shows only it while the full text still goes to the model.
      attachments are local files added in the start-screen input, taking effect with the first message.
      slideShot attaches a rendering of the current slide so the model sees what it's editing (AI Beautify) */
  preset?: {
    text: string
    nonce: number
    autoRun?: boolean
    displayText?: string
    attachments?: AttachmentMeta[]
    slideShot?: boolean
  } | null
  /** false shows only the collapsed rail; the component stays mounted so panel state survives */
  open?: boolean
  /** expand from the collapsed rail */
  onExpand?: () => void
  onCollapse?: () => void
  /** Visible rollback action; uses the same main-process history as Cmd/Ctrl+Z. */
  onUndo?: () => void
  /** Callback to update the path after AI generation lands on disk (title bar sync) */
  onPathChange?: (path: string) => void
  /** Generation progress callback (for the canvas top progress bar) */
  onDeckProgress?: (event: DeckProgressEvent | null) => void
  /** Absolute path of the currently open file (for chat history persistence) */
  currentFilePath?: string | null
}

/** Resizable panel width: persisted; min/max match the Excel panel */
const PANEL_WIDTH_KEY = 'slides-ai-panel-width'

export function AiPanel({
  slides,
  current,
  selectedIds,
  deckEmpty,
  images,
  applySlide,
  applyDeck,
  fitWidthPx,
  settings,
  preset,
  open = true,
  onExpand,
  onCollapse,
  onPathChange,
  onDeckProgress,
  currentFilePath,
}: AiPanelProps) {
  const { t } = useI18n()
  const { isReady } = useAiReady({
    getAiSettings: window.slidesApi?.getAiSettings
      ? () => window.slidesApi.getAiSettings()
      : undefined,
    onAiSettingsChanged: window.slidesApi?.onAiSettingsChanged
      ? (h) => window.slidesApi.onAiSettingsChanged!(h)
      : undefined,
    initialSettings: settings,
  })
  const [input, setInput] = useState('')
  const [busy, setBusy] = useState(false)
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [snapshots, setSnapshots] = useState<DeckSnapshot[]>([])
  const [snapshotsOpen, setSnapshotsOpen] = useState(false)
  const [copiedIdx, setCopiedIdx] = useState<number | null>(null)
  const [attachments, setAttachments] = useState<AttachmentMeta[]>([])
  const [attachNotice, setAttachNotice] = useState<string | null>(null)
  /** data-URL previews for image attachments, keyed by path */
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
      void window.desktop.readAttachmentImage(a.path).then((r) => {
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
  const [dragOver, setDragOver] = useState(false)
  const { panelWidth, asideRef, resizing, startResize } = useAiPanelWidth(PANEL_WIDTH_KEY, 280)

  // The .ai-dock wrapper owns the animated width (Excel-parity 180ms slide);
  // it tracks the resizable panel width through this variable
  // `open` dep: the aside ref only exists while expanded
  useEffect(() => {
    const dock = asideRef.current?.closest('.ai-dock') as HTMLElement | null
    dock?.style.setProperty('--ai-panel-width', `${panelWidth}px`)
  }, [panelWidth, open])
  /** Past conversation restored from JSONL (read-only transcript, not fed to the model) */
  const [historicChat, setHistoricChat] = useState<ChatEntry[]>([])
  /* Answered clarify receipts (answered-state card): view-only record, not chat data */
  const [clarifyAnswers, setClarifyAnswers] = useState<
    Array<{ afterIdx: number; qa: Array<{ q: string; a: string }> }>
  >([])
  const logRef = useRef<HTMLDivElement>(null)
  const clarifyWrapRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  /** Stop following once the user scrolls up; re-attach when near the bottom */
  const stickToBottomRef = useRef(true)
  /** Current chat's projectId/chatId, set after resolve succeeds */
  const chatRefIds = useRef<{ projectId: string; chatId: string } | null>(null)

  // The loop instance survives across renders; closures read refs for the latest state
  const slidesRef = useRef(slides)
  slidesRef.current = slides
  const localSlidesRef = useRef<RenderSlide[]>(slides)
  localSlidesRef.current = slides
  const currentRef = useRef(current)
  currentRef.current = current
  const selectedRef = useRef(selectedIds)
  selectedRef.current = selectedIds
  const applySlideRef = useRef(applySlide)
  applySlideRef.current = applySlide
  const applyDeckRef = useRef(applyDeck)
  applyDeckRef.current = applyDeck
  const onPathChangeRef = useRef(onPathChange)
  onPathChangeRef.current = onPathChange
  const onDeckProgressRef = useRef(onDeckProgress)
  onDeckProgressRef.current = onDeckProgress
  const settingsRef = useRef(settings)
  settingsRef.current = settings
  const imagesRef = useRef(images)
  imagesRef.current = images
  const attachmentsRef = useRef(attachments)
  attachmentsRef.current = attachments
  const runAttachmentsRef = useRef<readonly AttachmentMeta[]>([])
  // Paths of text attachments already read via read_attachment
  const readAttachmentPathsRef = useRef<Set<string>>(new Set())

  const instructionRef = useRef('')
  const lastInstructionRef = useRef('')
  /** Paired with lastInstructionRef: keeps the bubble showing only the user's request on retries */
  const lastDisplayTextRef = useRef<string | undefined>(undefined)
  /** Mirror of the last turn's (the final reply's turn) tool activity — used when persisting the assistant message,
      avoiding side effects inside the setState updater (StrictMode double-invokes updaters, duplicating history writes) */
  const lastTurnToolsRef = useRef<ToolActivity[]>([])
  /** Tool activity of the whole run (with args/output, accumulated across turns) — for full transcript persistence */
  const runToolsRef = useRef<
    Array<{ name: string; summary: string; isError?: boolean; input?: string; output?: string }>
  >([])
  /** A completed plan keeps firing content-building continuations while each
   * round makes progress (capped) — weak models stop after one page otherwise. */
  const nativeDeckPlanPendingRef = useRef(false)
  /** continuation rounds fired since plan/clarify */
  const nativeDeckContinuationsRef = useRef(0)
  /** content mutations since the last continuation fired — the chain stops when
   * a round adds nothing (deck complete or the model stalled) */
  const nativeDeckProgressRef = useRef(0)
  const internalRunRef = useRef(false)

  // ── Chat history persistence ──────────────────────────────────────────────
  /** Resolve chatId and load history on first mount (AiPanel resets by key; no need to watch currentFilePath changes) */
  useEffect(() => {
    const api = (window as Window & { projectApi?: typeof window.projectApi }).projectApi
    if (!api) return
    const tempChatId = `unsaved-${Date.now()}`
    void api
      .resolveChat({ filePath: currentFilePath ?? null, tempChatId })
      .then((ids) => {
        chatRefIds.current = ids
        return api.loadChat({ projectId: ids.projectId, chatId: ids.chatId, limit: 200 })
      })
      .then((msgs) => {
        if (msgs.length === 0) return
        setHistoricChat(
          msgs.map((m) => ({
            role: m.role,
            text: m.text,
            attachments: m.attachments?.flatMap((attachment) =>
              attachment.path && attachment.ext
                ? [
                    {
                      name: attachment.name,
                      path: attachment.path,
                      ext: attachment.ext,
                      sizeBytes: attachment.sizeBytes ?? 0,
                    },
                  ]
                : [],
            ),
            tools: m.tools?.map((t) => ({
              name: t.name,
              summary: t.summary,
              isError: t.isError,
              output: t.output ? t.output.slice(0, TOOL_OUTPUT_MAX_CHARS) : undefined,
            })),
          })),
        )
        // Restore model context: follow-ups after reopening a file continue the earlier conversation (only when the loop is idle with no history)
        loopRef.current?.restore(msgs.map((m) => ({ role: m.role, text: m.text })))
      })
      .catch(() => {
        /* History load failures are silent */
      })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  /** After an unsaved draft lands and gets a real path, bind the unsaved-* history to that file (recoverable by path on reopen) */
  useEffect(() => {
    const ids = chatRefIds.current
    const api = (window as Window & { projectApi?: typeof window.projectApi }).projectApi
    if (!api || !ids || !currentFilePath || !ids.chatId.startsWith('unsaved-')) return
    void api
      .rebindChat({
        projectId: ids.projectId,
        tempChatId: ids.chatId,
        newFilePath: currentFilePath,
      })
      .then((r) => {
        if (r?.chatId) chatRefIds.current = r
      })
      .catch(() => {
        /* Silent */
      })
  }, [currentFilePath])

  /** Persist one message (fails silently). tools include args/output (truncated by the store layer); attachments store metadata only. */
  const persistMessage = (
    role: 'user' | 'assistant',
    text: string,
    tools?: Array<{
      name: string
      summary: string
      isError?: boolean
      input?: string
      output?: string
    }>,
    attachments?: AttachmentMeta[],
  ) => {
    const ids = chatRefIds.current
    const api = (window as Window & { projectApi?: typeof window.projectApi }).projectApi
    if (!ids || !api) return
    void api
      .appendChat({
        projectId: ids.projectId,
        chatId: ids.chatId,
        role,
        text,
        ...(tools && tools.length > 0 ? { tools } : {}),
        ...(attachments && attachments.length > 0
          ? {
              attachments: attachments.map((a) => ({
                name: a.name,
                path: a.path,
                ext: a.ext,
                sizeBytes: a.sizeBytes,
              })),
            }
          : {}),
      })
      .catch(() => {
        /* Silent */
      })
  }

  // Resolved when the user submits/cancels; the AI takes the answer and continues.
  const [activeClarify, setActiveClarify] = useState<ClarifyQuestion[] | null>(null)
  const clarifyResolverRef = useRef<((r: { answers: string; cancelled?: boolean }) => void) | null>(
    null,
  )

  /** Synchronous re-entry guard between runWith trigger and loop.run (see the comment inside runWith) */
  const runStartingRef = useRef(false)
  /** Pages landed by this run's generation calls, pending the post-generation layout QC pass */
  const qcPagesRef = useRef<number[]>([])
  const qcAbortRef = useRef<AbortController | null>(null)
  const qcRunningRef = useRef(false)
  /** Latest runQcPass closure; the loop's onDone (built once) calls through this ref */
  const runQcPassRef = useRef<() => Promise<void>>(() => Promise.resolve())
  /** DeckAccess reused by the QC pass (same executors as the main loop's slides skill) */
  const accessRef = useRef<SlidesSkillDeps | null>(null)
  /** Wall-clock start of the current run, drives the elapsed badge */
  const runStartedAtRef = useRef(0)
  const historyBatchActiveRef = useRef(false)
  const inputEditedSinceRunRef = useRef(false)

  const finishHistoryBatch = async () => {
    if (!historyBatchActiveRef.current) return
    historyBatchActiveRef.current = false
    const id = await window.slidesApi.endHistoryBatch()
    if (typeof id !== 'number') return
    const label = (lastDisplayTextRef.current ?? instructionRef.current).slice(0, 40)
    setSnapshots((prev) =>
      [{ id, label, time: new Date().toLocaleTimeString() }, ...prev].slice(0, 20),
    )
  }

  const rollback = async (snapshot: DeckSnapshot) => {
    const restored = await window.slidesApi.aiSnapshotRestore(snapshot.id)
    if (!restored) return
    applyDeckRef.current(restored, Math.min(currentRef.current, restored.length - 1))
    setSnapshots((prev) => prev.filter((s) => s.id !== snapshot.id))
  }

  const patchLastAssistant = (
    patch: Partial<ChatEntry> | ((last: ChatEntry) => Partial<ChatEntry>),
  ) => {
    setChat((prev) => {
      const next = [...prev]
      const last = next[next.length - 1]
      if (!last || last.role !== 'assistant') return prev
      next[next.length - 1] = { ...last, ...(typeof patch === 'function' ? patch(last) : patch) }
      return next
    })
  }

  const runWithRef =
    useRef<
      (
        instruction: string,
        displayText?: string,
        opts?: { slideShot?: boolean; internal?: boolean },
      ) => void
    >(undefined)
  const loopRef = useRef<AgentLoop | null>(null)
  useEffect(() => {
    const access: SlidesSkillDeps = {
      getSlides: () => localSlidesRef.current ?? slidesRef.current,
      getCurrent: () => currentRef.current,
      getSelectedIds: () => selectedRef.current,
      applySlide: (i, updated) => {
        if (localSlidesRef.current) {
          localSlidesRef.current = localSlidesRef.current.map((s, idx) => (idx === i ? updated : s))
        }
        applySlideRef.current(i, updated)
      },
      applyDeck: (all, goTo) => {
        localSlidesRef.current = all
        applyDeckRef.current(all, goTo)
      },
      askClarification: (questions: ClarifyQuestion[]) => {
        return new Promise<{ answers: string; cancelled?: boolean }>((resolve) => {
          clarifyResolverRef.current = resolve
          setActiveClarify(questions)
        })
      },
      fitWidthPx,
      saveStyleTemplate: async (name, data) => {
        try {
          return await window.slidesApi.saveStyleTemplate(name, data)
        } catch {
          return { ok: false, error: String('') }
        }
      },
      listStyleTemplates: async () => {
        try {
          return await window.slidesApi.listStyleTemplates()
        } catch {
          return []
        }
      },
      loadStyleTemplate: async (name) => {
        try {
          return await window.slidesApi.loadStyleTemplate(name)
        } catch {
          return { ok: false, error: String('') }
        }
      },
    }
    accessRef.current = access
    loopRef.current = new AgentLoop({
      transport: createElectronTransport(() => settingsRef.current),
      systemSuffix: aiLangDirective,
      skill: composeSkills('slides+files', '', [
        createSlidesSkill(access),
        createFilesSkill(
          () => [...runAttachmentsRef.current],
          (path) => readAttachmentPathsRef.current.add(path),
        ),
      ]),
      maxTurns: null,
      maxToolCalls: null,
      events: {
        onThinking: (chunk) => {
          patchLastAssistant((last) => ({
            thinking: (last.thinking ?? '') + chunk,
          }))
        },
        onText: (raw) => {
          const { thinking, text } = extractThinking(raw)
          patchLastAssistant((last) => ({
            text,
            ...(thinking ? { thinking: last.thinking ? last.thinking : thinking } : {}),
          }))
        },
        onToolStart: (call) => {
          // Live "running" chip: replaced in place by onToolExecuted
          const activity: ToolActivity = {
            name: call.name,
            summary: call.name.replace(/[_-]+/g, ' '),
            running: true,
          }
          patchLastAssistant((last) => ({ tools: [...(last.tools ?? []), activity] }))
        },
        onToolExecuted: ({ call, execution }) => {
          const activity: ToolActivity = {
            name: call.name,
            summary: execution.summary,
            isError: execution.isError,
            output: execution.output ? execution.output.slice(0, TOOL_OUTPUT_MAX_CHARS) : undefined,
            // Side channel: display comes from tools, not into LLM context, UI only
            display: execution.display,
          }
          lastTurnToolsRef.current.push(activity)
          if (call.name === 'plan_deck') {
            // plan_deck scaffolds and populates all planned slides atomically; no continuation needed
            nativeDeckPlanPendingRef.current = false
            nativeDeckContinuationsRef.current = 0
            nativeDeckProgressRef.current = 0
          } else if (call.name === 'ask_clarification') {
            nativeDeckPlanPendingRef.current = false
            nativeDeckContinuationsRef.current = 0
            nativeDeckProgressRef.current = 0
          } else if (execution.mutated && DECK_CONTENT_TOOLS.has(call.name)) {
            nativeDeckProgressRef.current += 1
          }
          if (!execution.display) {
            runToolsRef.current.push({
              name: call.name,
              summary: execution.summary,
              isError: execution.isError,
              input: safeJsonInput(call.input),
              output: execution.output
                ? execution.output.slice(0, PERSIST_TOOL_FIELD_MAX)
                : undefined,
            })
          }
          patchLastAssistant((last) => {
            // Swap out the running placeholder pushed by onToolStart (parse-fail calls have none)
            const tools = [...(last.tools ?? [])]
            if (tools.at(-1)?.running) tools.pop()
            return { tools: [...tools, activity] }
          })
        },
        onTurnEnd: () => {
          lastTurnToolsRef.current = []
          // ponytail: keep tools and thinking in one unified assistant response rather than fragmenting into multiple empty bubbles
        },
        onDone: ({ text, cancelled, turnLimit, stuck }) => {
          const finalText = stuck
            ? [text, tGlobal('aiStoppedNote')].filter(Boolean).join('\n\n')
            : turnLimit
              ? [text, tGlobal('aiTurnLimit')].filter(Boolean).join('\n\n')
              : text || (cancelled ? tGlobal('aiStoppedNote') : '')
          const ranTools = runToolsRef.current.length > 0
          const historyText =
            finalText ||
            (ranTools ? '(completed tool actions; no text reply)' : tGlobal('aiNoResponse'))
          const deckRunFailed = cancelled || turnLimit || stuck
          const continueNativeDeck =
            !deckRunFailed &&
            nativeDeckPlanPendingRef.current &&
            nativeDeckContinuationsRef.current < MAX_DECK_CONTINUATIONS &&
            // first continuation always fires (plan-then-stop is the failure being
            // rescued); later ones only while the previous round built content
            (nativeDeckContinuationsRef.current === 0 || nativeDeckProgressRef.current > 0)
          if (continueNativeDeck) {
            nativeDeckContinuationsRef.current += 1
            nativeDeckProgressRef.current = 0
          } else {
            // chain finished: deck complete, budget exhausted, no progress, or failed
            nativeDeckPlanPendingRef.current = false
            nativeDeckContinuationsRef.current = 0
            nativeDeckProgressRef.current = 0
          }
          setChat((prev) => {
            const next = [...prev]
            const last = next.at(-1)
            if (!last || last.role !== 'assistant') return prev
            // Tool-heavy runs often end with an empty closing turn. Earlier
            // bubbles already show the executed work — drop the empty trailing
            // bubble instead of mislabeling the whole run as "no content".
            if (!finalText && !last.text && !last.tools?.length && ranTools) {
              next.pop()
              return next
            }
            next[next.length - 1] = {
              ...last,
              streaming: false,
              turnLimit,
              stuck,
              text: finalText || (last.tools?.length ? last.text : tGlobal('aiNoResponse')),
              // A stop mid-tool can leave a running placeholder behind — drop it
              tools: last.tools?.filter((tl) => !tl.running),
            }
            return next
          })
          if (!cancelled && !continueNativeDeck) {
            persistMessage('assistant', historyText, runToolsRef.current)
          }
          void finishHistoryBatch().finally(() => {
            if (continueNativeDeck) {
              setBusy(false)
              window.setTimeout(() => {
                if (nativeDeckContinuationsRef.current === 0) return
                runWithRef.current?.(
                  'Continue building the planned deck now: add the NEXT planned page with real content (never bracketed placeholder text). Do not explain or search. If every planned page already exists, reply exactly: Deck complete.',
                  undefined,
                  { internal: true },
                )
              }, 0)
              return
            }
            setBusy(false)
            // Post-generation layout QC: only after a completed run that landed generated pages
            if (cancelled) qcPagesRef.current = []
            else if (qcPagesRef.current.length > 0) void runQcPassRef.current()
          })
        },
        onError: (error) => {
          qcPagesRef.current = []
          setChat((prev) => {
            const next = [...prev]
            if (!internalRunRef.current) {
              // the loop rolled this run's user message out of the model context — surface that
              for (let i = next.length - 1; i >= 0; i--) {
                const entry = next[i]!
                if (entry.role === 'user') {
                  next[i] = { ...entry, undelivered: true }
                  break
                }
              }
            }
            const last = next.at(-1)
            if (last?.role === 'assistant') {
              next[next.length - 1] = {
                ...last,
                streaming: false,
                error,
                tools: last.tools?.filter((tl) => !tl.running),
              }
            }
            return next
          })
        },
      },
    })
  }, [fitWidthPx])
  const pendingAutoRunRef = useRef<{
    text: string
    displayText?: string
    slideShot?: boolean
  } | null>(null)

  useEffect(() => {
    if (!preset) return
    // Attachments from the start screen: merge into attachments first (ref updated synchronously so this runWith can read them),
    // then trigger the run. Deduplicated by path, safe under StrictMode double runs.
    if (preset.attachments && preset.attachments.length > 0) {
      const unique = preset.attachments.filter(
        (a) => !attachmentsRef.current.some((item) => item.path === a.path),
      )
      if (attachmentsRef.current.length === 0 && unique.length > 0) {
        attachmentsRef.current = [unique[0]]
        setAttachments([unique[0]])
      }
    }
    if (preset.autoRun) {
      if (isReady) {
        runWithRef.current?.(preset.text, preset.displayText, {
          slideShot: preset.slideShot ?? false,
        })
      } else {
        pendingAutoRunRef.current = {
          text: preset.text,
          displayText: preset.displayText,
          slideShot: preset.slideShot ?? false,
        }
      }
    } else {
      setInput(preset.text)
      inputRef.current?.focus()
    }
  }, [preset, isReady])

  useEffect(() => {
    if (isReady && pendingAutoRunRef.current) {
      const pending = pendingAutoRunRef.current
      pendingAutoRunRef.current = null
      runWithRef.current?.(pending.text, pending.displayText, {
        slideShot: pending.slideShot,
      })
    }
  }, [isReady])

  // `open` dep: re-expanding lands on messages streamed while collapsed
  useEffect(() => {
    if (stickToBottomRef.current) {
      logRef.current?.scrollTo({ top: logRef.current.scrollHeight })
    }
  }, [chat, open])

  useEffect(() => {
    if (activeClarify) {
      window.requestAnimationFrame(() => {
        clarifyWrapRef.current?.scrollIntoView({ block: 'start', behavior: 'smooth' })
      })
    }
  }, [activeClarify])

  const onLogScroll = () => {
    const el = logRef.current
    if (!el) return
    stickToBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 48
  }

  // Input box auto-sizes, up to seven lines (keep in sync with the CSS max-height);
  // empty clears the inline height outright so the CSS min-height governs
  // (a hidden-at-measure pass can leave a stale value), same as the shared AiComposer.
  useEffect(() => {
    const ta = inputRef.current
    if (!ta) return
    if (input === '') {
      ta.style.height = ''
      return
    }
    ta.style.height = 'auto'
    ta.style.height = `${Math.min(ta.scrollHeight, 168)}px`
    // `open` dep: re-measure after expand restores a draft
  }, [input, open])

  const run = () => runWith(input.trim())

  /** Image attachments read as base64, sent multimodally with this user message (≤5MB per image, max 20; isomorphic to docs) */
  const MAX_IMAGES_PER_MESSAGE = 20
  const collectImageAttachments = async (
    turnAttachments: readonly AttachmentMeta[] = attachmentsRef.current,
  ): Promise<AgentImage[]> => {
    const imageAtts = turnAttachments.filter((a) => ATTACHMENT_IMAGE_EXTS.has(a.ext))
    const images: AgentImage[] = []
    const failures: string[] = []
    for (const att of imageAtts.slice(0, MAX_IMAGES_PER_MESSAGE)) {
      const result = await window.desktop.readAttachmentImage(att.path)
      if (result.ok && result.base64 && result.mime) {
        images.push({ base64: result.base64, mime: result.mime })
      } else {
        failures.push(result.error ?? t('aiReadFailed', { name: att.name }))
      }
    }
    if (imageAtts.length > MAX_IMAGES_PER_MESSAGE) {
      failures.push(t('aiTooManyImages', { max: MAX_IMAGES_PER_MESSAGE }))
    }
    if (failures.length > 0) {
      setAttachNotice(failures.join(';'))
      window.setTimeout(() => setAttachNotice(null), 5000)
    }
    return images
  }

  /** Current slide rendered at pixelRatio 1 (vision-friendly size); null when rendering fails */
  const captureSlideShot = async (pageIndex: number): Promise<AgentImage | null> => {
    const slide = slidesRef.current[pageIndex]
    if (!slide) return null
    try {
      const [png] = await renderSlidesToPngBase64([slide], imagesRef.current, 1)
      return png ? { base64: png, mime: 'image/png' } : null
    } catch {
      return null
    }
  }

  const runWith = (
    instruction: string,
    displayText?: string,
    opts?: { slideShot?: boolean; internal?: boolean },
  ) => {
    runWithRef.current = runWith
    if (!isReady) return
    const loop = loopRef.current
    // runStartingRef: loop.run is called only after attachments are read asynchronously, during which loop.busy is still false,
    // so duplicate triggers must be blocked synchronously (e.g. StrictMode double-running the preset autoRun effect),
    // otherwise two sets of bubbles get pushed and the earlier assistant placeholder stays at "thinking" forever.
    // qcRunningRef: the post-generation QC pass edits the deck outside the main loop — no concurrent runs
    if (!instruction || !loop || loop.busy || runStartingRef.current || qcRunningRef.current) return
    const internal = opts?.internal === true
    const turnAttachments = internal ? [] : attachmentsRef.current
    if (!internal) runAttachmentsRef.current = turnAttachments
    internalRunRef.current = internal
    if (!internal) {
      nativeDeckPlanPendingRef.current = false
      nativeDeckContinuationsRef.current = 0
      nativeDeckProgressRef.current = 0
      instructionRef.current = instruction
      lastInstructionRef.current = instruction
      lastDisplayTextRef.current = displayText
    }
    runStartingRef.current = true
    setInput('')
    inputEditedSinceRunRef.current = false
    lastTurnToolsRef.current = []
    runToolsRef.current = []
    stickToBottomRef.current = true
    const shown = displayText ?? instruction
    setChat((prev) => {
      const next: ChatEntry[] = [
        ...prev.map((entry) =>
          entry.role === 'assistant' && entry.streaming ? { ...entry, streaming: false } : entry,
        ),
      ]
      if (!internal) next.push({ role: 'user', text: shown, attachments: turnAttachments })
      next.push({ role: 'assistant', text: '', streaming: true })
      return next
    })
    if (!internal) {
      attachmentsRef.current = []
      setAttachments([])
    }
    runStartedAtRef.current = Date.now()
    setBusy(true)
    // Persist user-visible requests only; continuation instructions are model-only orchestration.
    if (!internal) persistMessage('user', shown, undefined, turnAttachments)
    void collectImageAttachments(turnAttachments)
      .then(async (images) => {
        // AI Beautify sends the current slide's rendering along, so the model sees what it edits;
        // the note rides on the model instruction only — the chat bubble stays the localized preset text
        let modelInstruction = instruction
        if (opts?.slideShot) {
          const shot = await captureSlideShot(currentRef.current)
          if (shot) {
            images.push(shot)
            modelInstruction += `\n\n(Attached image: the current rendering of this slide, slideIndex ${currentRef.current}. Use it to spot visual issues the element inventory can't show.)`
          }
        }
        // Clear the flag before run: loop.run sets running synchronously, leaving no re-entry window
        runStartingRef.current = false
        if (await window.slidesApi.beginHistoryBatch()) historyBatchActiveRef.current = true
        loop.run(modelInstruction, images)
      })
      .catch(() => {
        runStartingRef.current = false
        void finishHistoryBatch().finally(() => setBusy(false))
      })
  }
  runWithRef.current = runWith

  /**
   * Post-generation layout QC: each page landed by this run gets one focused vision pass in a
   * fresh AgentLoop (screenshot + inventory → constrained fixes via execute_slide_script).
   * Each page's edits sit in their own history batch; if the deterministic audit says the page
   * got worse, that batch is rolled back. Progress streams into one assistant chat entry.
   */
  const runQcPass = async () => {
    const pages = qcPagesRef.current
    qcPagesRef.current = []
    const access = accessRef.current
    if (pages.length === 0 || !access || qcRunningRef.current || !isQcEnabled()) return
    qcRunningRef.current = true
    const controller = new AbortController()
    qcAbortRef.current = controller
    const capped = pages.slice(0, QC_MAX_PAGES)
    const transport = createElectronTransport(() => settingsRef.current)
    const header = tGlobal('aiQcStart', { count: capped.length })
    const lines: string[] = []
    const renderEntry = () => [header, ...lines].join('\n')
    setBusy(true)
    stickToBottomRef.current = true
    setChat((prev) => [...prev, { role: 'assistant', text: header, streaming: true }])
    try {
      for (const page of capped) {
        if (controller.signal.aborted) break
        const shot = await captureSlideShot(page)
        if (!shot) {
          if (slidesRef.current[page]) lines.push(tGlobal('aiQcPageSkipped', { n: page + 1 }))
          continue
        }
        const batchOpened = await window.slidesApi.beginHistoryBatch()
        const result = await qcSlidePage({
          access,
          transport,
          pageIndex: page,
          screenshot: shot,
          systemSuffix: aiLangDirective,
          signal: controller.signal,
        })
        const batchId = batchOpened ? await window.slidesApi.endHistoryBatch() : null
        if (controller.signal.aborted) break
        if (result.error) {
          lines.push(tGlobal('aiQcPageFailed', { n: page + 1, error: result.error }))
        } else if (result.edited && result.postIssues > result.preIssues) {
          // The fix made the deterministic audit worse — undo this page's batch
          if (typeof batchId === 'number') {
            const restored = await window.slidesApi.aiSnapshotRestore(batchId)
            if (restored)
              applyDeckRef.current(restored, Math.min(currentRef.current, restored.length - 1))
          }
          lines.push(tGlobal('aiQcPageReverted', { n: page + 1 }))
        } else if (result.edited) {
          const summary =
            result.reply && result.reply.toUpperCase() !== 'OK'
              ? result.reply
              : tGlobal('aiQcPageFixedDefault')
          lines.push(tGlobal('aiQcPageFixed', { n: page + 1, summary }))
          // Kept edits become a visible rollback point, same as main-run snapshots
          if (typeof batchId === 'number') {
            const label = tGlobal('aiQcPageFixed', { n: page + 1, summary }).slice(0, 40)
            setSnapshots((prev) =>
              [{ id: batchId, label, time: new Date().toLocaleTimeString() }, ...prev].slice(0, 20),
            )
          }
        } else {
          lines.push(tGlobal('aiQcPageOk', { n: page + 1 }))
        }
        patchLastAssistant({ text: renderEntry() })
      }
      if (pages.length > capped.length) {
        lines.push(tGlobal('aiQcCapped', { count: pages.length - capped.length }))
      }
      if (controller.signal.aborted) lines.push(tGlobal('aiQcStopped'))
    } finally {
      qcRunningRef.current = false
      qcAbortRef.current = null
      const finalText = renderEntry()
      patchLastAssistant({ streaming: false, text: finalText })
      persistMessage('assistant', finalText)
      setBusy(false)
    }
  }
  runQcPassRef.current = runQcPass

  /** Finish pending survey cards as "skipped" (called on stop/new conversation, avoiding leftover cards and orphan promises) */
  const dismissClarify = () => {
    clarifyResolverRef.current?.({ answers: '', cancelled: true })
    clarifyResolverRef.current = null
    setActiveClarify(null)
  }

  const cancel = () => {
    nativeDeckPlanPendingRef.current = false
    nativeDeckContinuationsRef.current = 0
    nativeDeckProgressRef.current = 0
    dismissClarify()
    qcAbortRef.current?.abort()
    loopRef.current?.cancel()
  }

  // Abort a QC pass still running when the panel unmounts (new file / panel remount by key)
  useEffect(() => () => qcAbortRef.current?.abort(), [])

  const retry = () => runWith(lastInstructionRef.current, lastDisplayTextRef.current)
  const continueRun = () => runWith(CONTINUE_INSTRUCTION, t('aiContinue'))

  const newChat = () => {
    nativeDeckPlanPendingRef.current = false
    nativeDeckContinuationsRef.current = 0
    nativeDeckProgressRef.current = 0
    dismissClarify()
    qcAbortRef.current?.abort()
    loopRef.current?.reset()
    setBusy(false)
    setChat([])
    attachmentsRef.current = []
    runAttachmentsRef.current = []
    setAttachments([])
    inputRef.current?.focus()
  }

  const copyMessage = (text: string, idx: number) => {
    void navigator.clipboard.writeText(text)
    setCopiedIdx(idx)
    window.setTimeout(() => setCopiedIdx((cur) => (cur === idx ? null : cur)), 1200)
  }

  const mergeAttachments = (result: AttachmentAddResult | null) => {
    if (!result) return
    const current = attachmentsRef.current
    const accepted = result.accepted.filter((a) => !current.some((item) => item.path === a.path))
    if (accepted.length > 0) {
      if (current.length > 0) {
        setAttachNotice('Only one attachment is allowed per chat. Remove the current file first.')
        window.setTimeout(() => setAttachNotice(null), 5000)
      } else {
        attachmentsRef.current = [accepted[0]]
        setAttachments([accepted[0]])
        if (accepted.length > 1) {
          setAttachNotice('Only one attachment is allowed per chat; additional files were ignored.')
          window.setTimeout(() => setAttachNotice(null), 5000)
        }
      }
    }
    if (result.rejected.length > 0) {
      setAttachNotice(result.rejected.join(';'))
      window.setTimeout(() => setAttachNotice(null), 5000)
    }
  }

  const pickAttachments = async () => mergeAttachments(await window.desktop.pickAttachments())

  const onDrop = async (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
    setDragOver(false)
    const paths = Array.from(e.dataTransfer.files)
      .map((f) => window.desktop.getPathForFile(f))
      .filter(Boolean)
    if (paths.length > 0) mergeAttachments(await window.desktop.addAttachmentPaths(paths))
  }

  /** Files pasted into the input box: those with local paths go the regular attachment route; pure bitmaps like screenshots land in a temp file first */
  const onPasteFiles = async (files: File[]) => {
    const paths: string[] = []
    for (const f of files) {
      const p = window.desktop.getPathForFile(f)
      if (p) {
        paths.push(p)
        continue
      }
      const ext = PASTE_MIME_EXT[f.type] ?? f.name.split('.').pop()?.toLowerCase() ?? 'bin'
      mergeAttachments(await window.desktop.addPastedImage(await f.arrayBuffer(), ext))
    }
    if (paths.length > 0) mergeAttachments(await window.desktop.addAttachmentPaths(paths))
  }

  const removeAttachment = (path: string) => {
    const next = attachmentsRef.current.filter((a) => a.path !== path)
    attachmentsRef.current = next
    setAttachments(next)
  }

  // collapsed: rail only — after all hooks, so the instance and its state survive
  if (!open) {
    return (
      <button className="ai-rail" title={t('appAiRailExpand')} onClick={onExpand}>
        <span className="ai-rail-icon" aria-hidden="true">
          <AraSpark size={21} />
        </span>
      </button>
    )
  }

  return (
    <AiPanelShell
      ref={asideRef}
      style={{ width: '100%' }}
      className={`ai-panel${activeClarify ? ' ai-panel-clarifying' : ''}${dragOver ? ' ai-panel-dragover' : ''}${resizing ? ' ai-panel-resizing' : ''}`}
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
        onNewChat={newChat}
        onCollapse={onCollapse}
        newChatLabel={t('aiNewChat')}
        collapseLabel={t('aiCollapsePanel')}
        newChatIcon={<IconNewChat size={15} />}
        collapseIcon={<IconSidebarCollapseLeft size={15} />}
        snapshotsCount={snapshots.length}
        isSnapshotsOpen={snapshotsOpen}
        onToggleSnapshots={() => setSnapshotsOpen((v) => !v)}
        live={busy}
      >
        <AiSnapshotsPopover
          isOpen={snapshotsOpen}
          onClose={() => setSnapshotsOpen(false)}
          snapshots={snapshots}
          onRollback={(s) => void rollback(s)}
          title={t('aiSnapshotsTitle')}
          rollbackLabel={t('aiRollback')}
          busy={busy}
        />
      </AiPanelHeader>

      <div ref={logRef} className="ai-chat" onScroll={onLogScroll}>
        {/* Past conversation (read-only transcript, not fed to the model), displayed continuously with the current turn */}
        {historicChat.length > 0 && (
          <>
            {historicChat.map((entry, i) => (
              <div key={`h${i}`} className={`ai-msg ai-msg-${entry.role} ai-msg-historic`}>
                {(entry.thinking || (entry.tools && entry.tools.length > 0)) && (
                  <AiProcessCard
                    thinking={entry.thinking}
                    tools={entry.tools}
                    isStreaming={false}
                    onOpenExternalUrl={openExternalUrl}
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
                  {t(deckEmpty ? 'aiEmptyGenTitle' : 'aiEmptyTitle')}
                </div>
                <div className="ai-chat-empty-body">
                  {t(deckEmpty ? 'aiEmptyGenBody1' : 'aiEmptyBody1')}
                  <br />
                  {t(deckEmpty ? 'aiEmptyGenBody2' : 'aiEmptyBody2')}
                </div>
                <div className="ai-starter-list">
                  {starterPrompts(t, deckEmpty ?? false).map((p) => (
                    <button
                      key={p}
                      className="ai-starter"
                      onClick={() => {
                        setInput(p)
                        inputRef.current?.focus()
                      }}
                    >
                      {p}
                    </button>
                  ))}
                </div>
              </>
            ) : (
              <AiUnconfiguredCard onConfigure={() => void window.slidesApi?.configureAi?.()} />
            )}
          </div>
        )}
        {chat.map((entry, i) => {
          if (
            entry.role === 'assistant' &&
            !entry.text &&
            !entry.streaming &&
            !entry.error &&
            !entry.tools?.length
          ) {
            return null
          }
          const isLast = i === chat.length - 1
          // Action row appears once per completed reply: on the turn's final segment only
          // (mid-turn segments have a following assistant entry; the live turn ends when !busy)
          const nextEntry = chat[i + 1]
          const turnEnded = nextEntry ? nextEntry.role === 'user' : !busy
          const showToolbar =
            entry.role === 'assistant' &&
            !entry.streaming &&
            turnEnded &&
            !!(entry.text || entry.error)
          return (
            <div
              key={i}
              className={`ai-msg ai-msg-${entry.role}${entry.role === 'assistant' && entry.streaming ? ' ai-msg-streaming' : ''}`}
            >
              {entry.role === 'assistant' ? (
                <>
                  {(() => {
                    const extracted = extractThinking(entry.text)
                    const displayThinking = entry.thinking || extracted.thinking
                    const displayText = extracted.thinking ? extracted.text : entry.text
                    const hasProcess = Boolean(
                      displayThinking || (entry.tools && entry.tools.length > 0),
                    )
                    return (
                      <>
                        {hasProcess && (
                          <AiProcessCard
                            thinking={displayThinking}
                            tools={entry.tools}
                            isStreaming={entry.streaming}
                            onOpenExternalUrl={openExternalUrl}
                          />
                        )}
                        {!displayText && entry.streaming && !hasProcess ? (
                          <span className="ai-typing-row">
                            <AiTypingIndicator label={t('aiContinuing')} />
                          </span>
                        ) : (
                          displayText && <Markdown text={displayText} />
                        )}
                      </>
                    )
                  })()}
                </>
              ) : (
                <>
                  {entry.text}
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
                </>
              )}
              {entry.role === 'user' && entry.undelivered && (
                <div className="ai-msg-undelivered">{t('aiUndelivered')}</div>
              )}
              {entry.error && (
                <div className="ai-msg-error">{t('aiMsgError', { error: entry.error })}</div>
              )}
              {entry.deckProgress && <DeckProgressCard progress={entry.deckProgress} />}
              {showToolbar && (
                <div className="ai-msg-toolbar">
                  {entry.text && (
                    <button
                      className="ai-msg-tool-btn"
                      onClick={() => copyMessage(entry.text, i)}
                      aria-label={t('aiCopyReply')}
                      data-tip={t('aiCopyReply')}
                    >
                      {copiedIdx === i ? (
                        <svg
                          width="14"
                          height="14"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="1.8"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <polyline points="20 6 9 17 4 12" />
                        </svg>
                      ) : (
                        <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
                          <path
                            d="M14.6113 5.34253C16.0608 5.3428 17.2363 6.518 17.2363 7.96753V15.5066C17.2361 16.956 16.0607 18.1313 14.6113 18.1316H7.07227C5.62267 18.1316 4.44751 16.9561 4.44727 15.5066V7.96753C4.44732 6.51783 5.62255 5.34253 7.07227 5.34253H14.6113ZM7.07227 6.59253C6.31291 6.59253 5.69732 7.20819 5.69727 7.96753V15.5066C5.69751 16.2658 6.31302 16.8816 7.07227 16.8816H14.6113C15.3703 16.8813 15.9861 16.2656 15.9863 15.5066V7.96753C15.9863 7.20835 15.3705 6.5928 14.6113 6.59253H7.07227ZM10.0176 2.8689C10.3626 2.86905 10.6426 3.14882 10.6426 3.4939C10.6425 3.83888 10.3626 4.11874 10.0176 4.1189H4.59961C3.84022 4.1189 3.22461 4.73451 3.22461 5.4939V11.324C3.22433 11.6689 2.94461 11.949 2.59961 11.949C2.25461 11.949 1.97489 11.6689 1.97461 11.324V5.4939C1.97461 4.04415 3.14987 2.8689 4.59961 2.8689H10.0176Z"
                            fill="currentColor"
                          />
                        </svg>
                      )}
                    </button>
                  )}
                  {isLast && !busy && lastInstructionRef.current && (
                    <button
                      className="ai-msg-tool-btn"
                      onClick={retry}
                      aria-label={t('aiRegenerate')}
                      data-tip={t('aiRegenerate')}
                    >
                      <IconRefresh size={12} />
                    </button>
                  )}
                </div>
              )}
              {entry.turnLimit && !entry.stuck && isLast && !busy && (
                <div className="ai-turn-limit-banner">
                  <div className="ai-turn-limit-body">
                    <span className="ai-turn-limit-icon">⏳</span>
                    <span className="ai-turn-limit-text">
                      {aiLangDirective() === 'id'
                        ? 'Batas langkah tercapai. Lanjutkan pembuatan untuk menyelesaikan slide yang tersisa.'
                        : 'Tool-call step limit reached. Continue generation to finish remaining slides.'}
                    </span>
                  </div>
                  <button type="button" className="ai-continue-action-btn" onClick={continueRun}>
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor">
                      <polygon points="5 3 19 12 5 21 5 3" />
                    </svg>
                    {t('aiContinue')}
                  </button>
                </div>
              )}
              {clarifyAnswers
                .filter((c) => c.afterIdx === i)
                .map((c, k) => (
                  <div key={`ca${k}`} className="ai-clarify-answered">
                    <div className="ai-clarify-answered-head">
                      <span className="ai-clarify-answered-badge">📋 {t('aiClarifyTitle')}</span>
                    </div>
                    {c.qa.map((pair, m) => (
                      <div key={m} className="ai-clarify-answered-row">
                        <div className="ai-clarify-answered-q">{pair.q}</div>
                        <div className="ai-clarify-answered-a">{pair.a}</div>
                      </div>
                    ))}
                  </div>
                ))}
            </div>
          )
        })}
        {activeClarify && (
          <div className="ai-clarify-active-wrap" ref={clarifyWrapRef}>
            <ClarifyCard
              questions={activeClarify}
              onSubmit={(answers, qa) => {
                clarifyResolverRef.current?.({ answers })
                clarifyResolverRef.current = null
                setActiveClarify(null)
                setClarifyAnswers((prev) => [...prev, { afterIdx: chat.length - 1, qa }])
              }}
              onSkip={() => {
                clarifyResolverRef.current?.({ answers: '', cancelled: true })
                clarifyResolverRef.current = null
                setActiveClarify(null)
              }}
            />
          </div>
        )}
      </div>

      {/* Reclaimed bottom space: version snapshots now live in AiPanelHeader popover */}

      {isReady ? (
        <div className="ai-composer">
          {attachNotice && <div className="ai-attach-notice">{attachNotice}</div>}
          <div className="ai-input-box">
            {attachments.length > 0 && (
              <div className="ai-attachments" onScroll={onAttachmentsScroll}>
                {attachments.map((a) =>
                  ATTACHMENT_IMAGE_EXTS.has(a.ext) ? (
                    <span key={a.path} className="ai-attachment-thumb" title={a.path}>
                      {attachmentPreviews[a.path] ? (
                        <img src={attachmentPreviews[a.path]} alt={a.name} />
                      ) : (
                        <span className="ai-attachment-thumb-pending" aria-hidden>
                          <img src={fileImageIcon} alt="" />
                        </span>
                      )}
                      <button
                        className="ai-attachment-thumb-remove"
                        onClick={() => removeAttachment(a.path)}
                        title={t('aiRemoveAttachment')}
                        aria-label={t('aiRemoveAttachment')}
                      >
                        <svg width="16" height="16" viewBox="0 0 32 32" aria-hidden>
                          <path
                            d="M24 9.4L22.6 8L16 14.6L9.4 8L8 9.4l6.6 6.6L8 22.6l1.4 1.4 6.6-6.6 6.6 6.6 1.4-1.4-6.6-6.6L24 9.4z"
                            fill="currentColor"
                            stroke="currentColor"
                            strokeWidth="0.25"
                          />
                        </svg>
                      </button>
                    </span>
                  ) : (
                    <span key={a.path} className="ai-attachment-card" title={a.path}>
                      <span className="ai-attachment-card-meta">
                        <span className="ai-attachment-card-name">{truncateCardName(a.name)}</span>
                        <span className="ai-attachment-card-size">
                          {formatAttachmentSize(a.sizeBytes)}
                        </span>
                      </span>
                      <button
                        className="ai-attachment-thumb-remove"
                        onClick={() => removeAttachment(a.path)}
                        title={t('aiRemoveAttachment')}
                        aria-label={t('aiRemoveAttachment')}
                      >
                        <svg width="16" height="16" viewBox="0 0 32 32" aria-hidden>
                          <path
                            d="M24 9.4L22.6 8L16 14.6 9.4 8 8 9.4l6.6 6.6L8 22.6l1.4 1.4 6.6-6.6 6.6 6.6 1.4-1.4-6.6-6.6L24 9.4z"
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
            )}
            <textarea
              ref={inputRef}
              disabled={!isReady || Boolean(activeClarify)}
              value={input}
              data-slides-ai-input="true"
              data-deck-undo-ready={!busy && !inputEditedSinceRunRef.current ? 'true' : 'false'}
              placeholder={
                activeClarify
                  ? 'Please answer the questions above…'
                  : t(deckEmpty ? 'aiInputPlaceholderGen' : 'aiInputPlaceholder')
              }
              aria-label={t(deckEmpty ? 'aiInputPlaceholderGen' : 'aiInputPlaceholder')}
              onChange={(e) => {
                inputEditedSinceRunRef.current = true
                setInput(e.target.value)
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
                  e.preventDefault()
                  run()
                } else if (e.key === 'Escape' && busy) {
                  e.preventDefault()
                  cancel()
                }
              }}
              onPaste={(e) => {
                const files = Array.from(e.clipboardData.files)
                if (files.length === 0) return
                e.preventDefault()
                void onPasteFiles(files)
              }}
              rows={1}
            />
            <div className="ai-input-footer">
              <button
                className="ai-attach-btn"
                onClick={pickAttachments}
                title={t('aiAttachTitle')}
                disabled={!isReady}
              >
                <IconAttach size={18} />
              </button>
              {busy ? (
                <button
                  className="ai-send-btn ai-stop-btn"
                  onClick={cancel}
                  title={t('aiStopGeneration')}
                  aria-label={t('aiStop')}
                >
                  <IconStop size={20} />
                </button>
              ) : (
                <button
                  className="ai-send-btn"
                  onClick={run}
                  disabled={!isReady || !input.trim()}
                  title={t('aiSend')}
                  aria-label={t('aiSend')}
                >
                  <IconArrowUp size={18} />
                </button>
              )}
            </div>
          </div>
        </div>
      ) : null}
    </AiPanelShell>
  )
}

/** Generation progress card: structured step archive (style → plan → images → pages → done) */
function DeckProgressCard({ progress }: { progress: DeckProgressSnapshot }) {
  const { t } = useI18n()
  // Collapsed by default: while generating only the one-line head shows (fewer concurrent loaders);
  // expanding is a view-only toggle
  const [open, setOpen] = useState(false)
  const { style, plan, images, pages, isDone, finalTotal } = progress

  type StepStatus = 'done' | 'error' | 'running'

  // Fix the step display order (only steps that have appeared are shown).
  // Labels come from skill-layer events (backend-defined steps pattern);
  // in progress shows a live summary (e.g. "planned 5/10 page outlines…"), frozen to the label when done.
  const steps: Array<{ key: string; label: string; stepStatus: StepStatus }> = []

  const stepView = (
    status: 'running' | 'done' | 'error',
    label: string,
    summary: string,
  ): { label: string; stepStatus: StepStatus } => ({
    // In progress/failed read summary; success freezes to the label
    label: status === 'done' ? label : summary,
    stepStatus: status === 'done' ? 'done' : status === 'error' ? 'error' : 'running',
  })

  if (style) {
    steps.push({ key: 'style', ...stepView(style.status, style.label, style.summary) })
  }
  if (plan) {
    steps.push({ key: 'plan', ...stepView(plan.status, plan.label, plan.summary) })
  }
  if (images) {
    steps.push({ key: 'images', ...stepView(images.status, images.label, images.summary) })
  }
  if (pages) {
    const allDone = isDone || pages.status === 'done'
    const hasError = pages.items.some((p) => p.status === 'error')
    steps.push({
      key: 'pages',
      label: allDone
        ? `${pages.label}${pages.total > 0 ? t('aiPagesSuffix', { n: pages.total }) : ''}`
        : pages.summary || pages.label,
      stepStatus: allDone ? (hasError ? 'error' : 'done') : 'running',
    })
  }

  // Show the summary when done; if any step errored, change the title to failed (avoiding perpetual "generating…" + spinner)
  const doneSummary = isDone && finalTotal != null ? t('aiProgressDone', { n: finalTotal }) : null
  const hasStepError = steps.some((s) => s.stepStatus === 'error')

  if (steps.length === 0) return null

  return (
    <div className="deck-progress-card">
      <button
        type="button"
        className="deck-progress-head"
        aria-expanded={open}
        onClick={() => setOpen(!open)}
      >
        {!doneSummary && !hasStepError && <span className="deck-progress-spinner" aria-hidden />}
        {doneSummary ? (
          <span className="deck-progress-done">{doneSummary}</span>
        ) : (
          <span className={`deck-progress-title${hasStepError ? ' is-error' : ''}`}>
            {hasStepError ? t('aiProgressFailed') : t('aiProgressTitle')}
          </span>
        )}
        <span className={`ai-tool-chip-caret${open ? ' open' : ''}`} aria-hidden>
          ›
        </span>
      </button>
      <div className={`deck-progress-body${open ? ' open' : ''}`}>
        <div className="deck-progress-body-inner">
          <div className="deck-progress-steps">
            {steps.map((step) => (
              <div key={step.key} className="deck-progress-step">
                <span className={`deck-progress-icon ${step.stepStatus}`}>
                  {step.stepStatus === 'running' ? (
                    <span className="deck-progress-spinner" />
                  ) : step.stepStatus === 'done' ? (
                    <svg width="11" height="11" viewBox="0 0 12 12" fill="none">
                      <path
                        d="M2 6l3 3 5-5"
                        stroke="currentColor"
                        strokeWidth="0.75"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />
                    </svg>
                  ) : (
                    <svg width="11" height="11" viewBox="0 0 12 12" fill="none">
                      <path
                        d="M2 2l8 8M10 2l-8 8"
                        stroke="currentColor"
                        strokeWidth="0.75"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />
                    </svg>
                  )}
                </span>
                <span className="deck-progress-step-label">{step.label}</span>
              </div>
            ))}
          </div>
          {pages && pages.items.length > 0 && (
            <div className="deck-progress-pages">
              {pages.items.map((p, i) => (
                <div key={i} className={`deck-progress-page-item deck-progress-page-${p.status}`}>
                  <span className="deck-progress-page-icon">
                    {p.status === 'done'
                      ? '✓'
                      : p.status === 'error'
                        ? '✗'
                        : p.status === 'running'
                          ? '⋯'
                          : '·'}
                  </span>
                  <span className="deck-progress-page-title">{`${t('aiPageN', { n: i + 1 })}${p.title ? '·' + p.title : ''}`}</span>
                  {p.status === 'error' && p.error && (
                    <span className="deck-progress-page-error-text" title={p.error}>
                      {p.error}
                    </span>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

/** Survey card: options clickable per question (single/multi), with "decide for me" and "Other (fill in)". */
function ClarifyCard({
  questions,
  onSubmit,
  onSkip,
}: {
  questions: ClarifyQuestion[]
  onSubmit: (answers: string, qa: Array<{ q: string; a: string }>) => void
  onSkip: () => void
}) {
  const { t } = useI18n()
  // Per question: set of selected options + the "Other" free text
  const [picked, setPicked] = useState<Record<string, Set<string>>>({})
  const [other, setOther] = useState<Record<string, string>>({})
  // Latest selections for the delayed auto-submit (the timer closure would otherwise
  // read the pre-click state and drop the final pick)
  const pickedRef = useRef(picked)
  pickedRef.current = picked
  const otherRef = useRef(other)
  otherRef.current = other
  // Pager view-state (one question at a time): back-nav limited to visited range,
  // single-select auto-advances after a beat, typing cancels the pending advance
  const [qIdx, setQIdx] = useState(0)
  const [furthest, setFurthest] = useState(0)
  const [slideDir, setSlideDir] = useState<'next' | 'prev'>('next')
  const advanceTimerRef = useRef<number | null>(null)

  const cancelAdvance = () => {
    if (advanceTimerRef.current) {
      window.clearTimeout(advanceTimerRef.current)
      advanceTimerRef.current = null
    }
  }

  const goTo = (i: number) => {
    cancelAdvance()
    const clamped = Math.max(0, Math.min(i, questions.length - 1))
    setSlideDir(clamped >= qIdx ? 'next' : 'prev')
    setQIdx(clamped)
    setFurthest((f) => Math.max(f, clamped))
  }

  // Single-select: picking advances; picking on the LAST question submits after the
  // same beat (typing in "Other" or navigating cancels the pending action)
  const scheduleAdvance = (from: number) => {
    cancelAdvance()
    const last = from >= questions.length - 1
    advanceTimerRef.current = window.setTimeout(() => {
      advanceTimerRef.current = null
      if (last) submit()
      else goTo(from + 1)
    }, 250)
  }

  useEffect(() => cancelAdvance, [])

  const toggle = (qid: string, opt: string, multi?: boolean) => {
    setPicked((prev) => {
      const cur = new Set(prev[qid] ?? [])
      if (multi) {
        if (cur.has(opt)) cur.delete(opt)
        else cur.add(opt)
      } else {
        cur.clear()
        cur.add(opt)
      }
      return { ...prev, [qid]: cur }
    })
  }

  const submit = () => {
    const qa = questions.map((q) => {
      const chosen = [...(pickedRef.current[q.id] ?? [])]
      const ot = (otherRef.current[q.id] ?? '').trim()
      if (ot) chosen.push(ot)
      const ans = chosen.length ? chosen.join('、') : t('aiClarifyDecideAnswer')
      return { q: q.label, a: ans }
    })
    onSubmit(qa.map(({ q, a }) => `${q}: ${a}`).join('\n'), qa)
  }

  const q = questions[qIdx]
  const isLast = qIdx === questions.length - 1
  const selCount = picked[q.id]?.size ?? 0
  const decideAnswer = t('aiClarifyDecideAnswer')

  const renderOpt = (opt: string, label: string) => (
    <button
      key={opt}
      className={`ai-clarify-opt${q.multi ? ' multi' : ''}${picked[q.id]?.has(opt) ? ' ai-clarify-opt-on' : ''}`}
      role={q.multi ? 'checkbox' : 'radio'}
      aria-checked={picked[q.id]?.has(opt) ?? false}
      onClick={() => {
        toggle(q.id, opt, q.multi)
        if (!q.multi) scheduleAdvance(qIdx)
      }}
    >
      <span className="ai-clarify-opt-box" aria-hidden />
      <span className="ai-clarify-opt-label">{label}</span>
      {!q.multi && (
        <span className="ai-clarify-opt-arrow" aria-hidden>
          ›
        </span>
      )}
    </button>
  )

  return (
    <div className="ai-clarify-card">
      <div className="ai-clarify-head">
        <span className="ai-clarify-head-label">{t('aiClarifyTitle')}</span>
        <div className="ai-clarify-head-pager">
          <span className="ai-clarify-head-progress" aria-live="polite">
            {`${qIdx + 1} / ${questions.length}`}
          </span>
          <span className="ai-clarify-head-arrows">
            <button
              type="button"
              className="ai-clarify-head-arrow"
              disabled={qIdx === 0}
              onClick={() => goTo(qIdx - 1)}
              aria-label="‹"
            >
              ‹
            </button>
            <button
              type="button"
              className="ai-clarify-head-arrow"
              disabled={qIdx >= furthest}
              onClick={() => goTo(qIdx + 1)}
              aria-label="›"
            >
              ›
            </button>
          </span>
        </div>
      </div>
      <div key={q.id} className={`ai-clarify-q slide-${slideDir}`}>
        <div className="ai-clarify-label">{q.label}</div>
        {q.description && <div className="ai-clarify-desc">{q.description}</div>}
        <div className="ai-clarify-opts">
          {q.options.map((opt) => renderOpt(opt, opt))}
          {renderOpt(decideAnswer, t('aiClarifyDecide'))}
        </div>
        <input
          className="ai-clarify-other"
          placeholder={t('aiClarifyOther')}
          value={other[q.id] ?? ''}
          onChange={(e) => {
            cancelAdvance()
            setOther((p) => ({ ...p, [q.id]: e.target.value }))
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault()
              cancelAdvance()
              if (isLast) submit()
              else goTo(qIdx + 1)
            }
          }}
        />
      </div>
      <div className={`ai-clarify-actions${q.multi ? ' multi' : ''}`}>
        {q.multi && selCount > 0 && (
          <span className="ai-clarify-count">{t('aiClarifySelected', { n: selCount })}</span>
        )}
        <span className="ai-clarify-actions-btns">
          <button type="button" className="ai-clarify-skip" onClick={onSkip}>
            {t('aiClarifySkip')}
          </button>
          {isLast ? (
            <button type="button" className="ai-clarify-submit" onClick={submit}>
              {t('aiClarifySubmit')}
            </button>
          ) : (
            <button
              type="button"
              className="ai-clarify-next"
              onClick={() => goTo(qIdx + 1)}
              aria-label="›"
            >
              ›
            </button>
          )}
        </span>
      </div>
    </div>
  )
}
