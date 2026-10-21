import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { Block } from '@araoffice/docx-engine'
import { AgentLoop, composeSkills, type AgentImage } from '@araoffice/agent-core'
import type { AiSettings, AttachmentAddResult, AttachmentMeta, ToolDisplay } from '../../shared/ipc'
import { ATTACHMENT_IMAGE_EXTS } from '../../shared/ipc'
import type { PmNode } from '../editor/convert'
import { findNumId, type NumIds } from './protocol'
import { markDocSeen } from './tools'
import { createDocsSkill } from './docs-skill'
import { applyRevisionsBy } from '../editor/revisions'
import fileImageIcon from '../assets/file-image.png'
import { CONTINUE_INSTRUCTION } from '@araoffice/agent-core'
import { createFilesSkill } from './files-skill'
import { createElectronTransport } from './transport'
import { useI18n, t as tModule, aiLangDirective } from '../i18n/locale'
import {
  AiComposer,
  AraSpark,
  AiPanelHeader,
  AiPanelShell,
  AiProcessCard,
  AiSnapshotsPopover,
  AiTypingIndicator,
  AiUnconfiguredCard,
  AiWelcomeHub,
  Markdown,
  extractThinking,
  IconAttach,
  IconBulb,
  IconFileText,
  IconNewChat,
  IconPen,
  IconSidebarCollapse,
  IconSparkles,
  openExternalUrl,
  useAiPanelWidth,
  useAiReady,
} from '@araoffice/ui'
interface Snapshot {
  label: string
  time: string
  json: PmNode
}

interface ToolActivity {
  name: string
  summary: string
  /** still executing: rendered as a spinner chip, replaced in place when the tool finishes */
  running?: boolean
  isError?: boolean
  /** Tool output (truncated on the UI side); when set, the row can be expanded for details */
  output?: string
  display?: ToolDisplay
}

/** Max characters of tool output in the UI expansion panel */
const TOOL_OUTPUT_MAX_CHARS = 2000

/** Cap on tool args/output persisted in the transcript (the store layer has another 16k truncation fallback) */
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
  /** tool executions performed during this assistant turn */
  tools?: ToolActivity[]
  attachments?: AttachmentMeta[]
}

/** resizable panel width: persisted, clamped so neither pane collapses */
const PANEL_WIDTH_KEY = 'docs-ai-panel-width'

/** persisted UI preference: highlight AI edits in yellow and ask for confirmation */
const TRACK_CHANGES_KEY = 'ai-docs-track-changes'

/** Clipboard bitmap MIME → attachment extension (corresponds to ATTACHMENT_IMAGE_EXTS) */
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

/** author name on AI-generated tracked revisions (accept/reject via Review) */
export const AI_REVISION_AUTHOR = 'AI Assistant'

interface AiPanelProps {
  editor: Editor
  blocks: Block[]
  settings: AiSettings
  /** the document has no text yet — the empty-state copy offers drafting instead of editing */
  docEmpty?: boolean
  /** fallback numbering ids for documents created from the blank template */
  numIdFallback?: NumIds | null
  /** preset instruction pushed from the ribbon or start screen; autoRun sends it immediately */
  preset?: { text: string; nonce: number; autoRun?: boolean } | null
  /** false shows only the collapsed rail; the component stays mounted so panel state survives */
  open?: boolean
  /** expand from the collapsed rail */
  onExpand?: () => void
  /** collapse the panel to the sidebar rail */
  onCollapse?: () => void
  /** Absolute path of the currently open file (used for chat-history persistence) */
  filePath?: string | null
}

export function AiPanel({
  editor,
  blocks,
  settings,
  docEmpty,
  numIdFallback,
  preset,
  open = true,
  onExpand,
  onCollapse,
  filePath,
}: AiPanelProps) {
  const { t } = useI18n()
  const { isReady } = useAiReady({
    getAiSettings: window.desktop?.getAiSettings ? () => window.desktop.getAiSettings() : undefined,
    onAiSettingsChanged: window.desktop?.onAiSettingsChanged
      ? (h) => window.desktop.onAiSettingsChanged!(h)
      : undefined,
    initialSettings: settings,
  })
  const [input, setInput] = useState('')
  const [busy, setBusy] = useState(false)
  /** Wall-clock start of the current run, drives the elapsed badge */
  const runStartedAtRef = useRef(0)
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [snapshots, setSnapshots] = useState<Snapshot[]>([])
  const [trackChanges, setTrackChanges] = useState(
    () => localStorage.getItem(TRACK_CHANGES_KEY) === '1',
  )
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
  const { panelWidth, asideRef, resizing, startResize } = useAiPanelWidth(PANEL_WIDTH_KEY, 360)
  const [snapshotsOpen, setSnapshotsOpen] = useState(false)

  // The .ai-dock wrapper owns the animated width (Excel-parity 180ms slide);
  // it tracks the resizable panel width through this variable
  // `open` dep: the aside ref only exists while expanded
  useEffect(() => {
    const dock = asideRef.current?.closest('.ai-dock') as HTMLElement | null
    dock?.style.setProperty('--ai-panel-width', `${panelWidth}px`)
  }, [panelWidth, open])
  /** Past conversation restored from JSONL (read-only transcript, not fed to the model) */
  const [historicChat, setHistoricChat] = useState<ChatEntry[]>([])
  // bumped on selection/doc changes so the scope hint & quick actions stay fresh
  const [, setScopeTick] = useState(0)
  const logRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  /** false once the user scrolls up to read; re-arms near the bottom */
  const stickToBottomRef = useRef(true)
  /** projectId/chatId of the current chat */
  const chatRefIds = useRef<{ projectId: string; chatId: string } | null>(null)

  const editorRef = useRef(editor)
  editorRef.current = editor
  const settingsRef = useRef(settings)
  settingsRef.current = settings
  editorRef.current = editor
  const blocksRef = useRef(blocks)
  blocksRef.current = blocks
  const numIdFallbackRef = useRef(numIdFallback)
  numIdFallbackRef.current = numIdFallback
  const attachmentsRef = useRef(attachments)
  attachmentsRef.current = attachments
  const runAttachmentsRef = useRef<readonly AttachmentMeta[]>([])
  const mergeAttachments = (result: AttachmentAddResult | null) => {
    if (!result) return
    const current = attachmentsRef.current
    const accepted = result.accepted.filter((a) => !current.some((item) => item.path === a.path))
    if (accepted.length > 0) {
      if (current.length > 0) {
        setAttachNotice('Only one attachment is allowed per chat. Remove the current file first.')
        window.setTimeout(() => setAttachNotice(null), 5000)
      } else {
        const first = accepted[0]
        if (first) {
          attachmentsRef.current = [first]
          setAttachments([first])
        }
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
  const trackChangesRef = useRef(trackChanges)
  trackChangesRef.current = trackChanges

  /** drop every aiChanged flag; silent = skip undo history (auto-accept path) */
  const clearAiHighlights = (silent = false) => {
    const view = editorRef.current.view
    let tr = view.state.tr
    let touched = false
    view.state.doc.forEach((node, offset) => {
      if (node.attrs.aiChanged) {
        tr = tr.setNodeMarkup(offset, undefined, { ...node.attrs, aiChanged: false })
        touched = true
      }
    })
    if (silent) tr = tr.setMeta('addToHistory', false)
    if (touched) {
      view.dispatch(tr)
      // AI-pipeline housekeeping, not a user edit: keep the freshness baseline current
      markDocSeen(editorRef.current)
    }
  }
  /** instruction of the in-flight run, labels its rollback snapshot */
  const instructionRef = useRef('')
  /** last sent instruction, for one-click retry */
  const lastInstructionRef = useRef('')
  /** Tool activity of the whole run (with args/output, accumulated across turns) — for full
      transcript persistence, and so persisting needn't do side effects inside a setState updater */
  const runToolsRef = useRef<
    Array<{ name: string; summary: string; isError?: boolean; input?: string; output?: string }>
  >([])

  // ── Chat-history persistence ────────────────────────────────────────────
  useEffect(() => {
    const api = (window as Window & { projectApi?: typeof window.projectApi }).projectApi
    if (!api) return
    const tempChatId = `unsaved-${Date.now()}`
    void api
      .resolveChat({ filePath: filePath ?? null, tempChatId })
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
        // restore model context: follow-ups after reopening a file continue the previous conversation (only when the loop is idle with no history)
        loopRef.current?.restore(msgs.map((m) => ({ role: m.role, text: m.text })))
      })
      .catch(() => {
        /* history load failures are silent */
      })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  /** After an unsaved document's first save yields a real path, bind the unsaved-* history to that file (recoverable by path on reopen) */
  useEffect(() => {
    const ids = chatRefIds.current
    const api = (window as Window & { projectApi?: typeof window.projectApi }).projectApi
    if (!api || !ids || !filePath || !ids.chatId.startsWith('unsaved-')) return
    void api
      .rebindChat({ projectId: ids.projectId, tempChatId: ids.chatId, newFilePath: filePath })
      .then((r) => {
        if (r?.chatId) chatRefIds.current = r
      })
      .catch(() => {
        /* silent */
      })
  }, [filePath])

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
        /* silent */
      })
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

  const loopRef = useRef<AgentLoop<PmNode> | null>(null)
  if (!loopRef.current) {
    const numIds = (): NumIds => ({
      bullet: findNumId(blocksRef.current, 'bullet') ?? numIdFallbackRef.current?.bullet ?? null,
      ordered: findNumId(blocksRef.current, 'ordered') ?? numIdFallbackRef.current?.ordered ?? null,
    })
    loopRef.current = new AgentLoop<PmNode>({
      transport: createElectronTransport(() => settingsRef.current),
      systemSuffix: aiLangDirective,
      maxTurns: null,
      maxToolCalls: null,
      skill: composeSkills('docs+files', '', [
        createDocsSkill({
          getEditor: () => editorRef.current,
          getNumIds: numIds,
          getTrack: () => (trackChangesRef.current ? { author: AI_REVISION_AUTHOR } : undefined),
        }),
        createFilesSkill(() => [...runAttachmentsRef.current]),
      ]),
      captureSnapshot: () => editorRef.current.getJSON() as PmNode,
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
          patchLastAssistant((last) => ({
            tools: [
              ...(last.tools ?? []),
              { name: call.name, summary: call.name.replace(/[_-]+/g, ' '), running: true },
            ],
          }))
        },
        onToolExecuted: ({ call, execution, snapshotBefore }) => {
          if (snapshotBefore) {
            setSnapshots((prev) =>
              [
                {
                  label: instructionRef.current.slice(0, 40),
                  time: new Date().toLocaleTimeString(),
                  json: snapshotBefore,
                },
                ...prev,
              ].slice(0, 20),
            )
          }
          if (execution.mutated) {
            // tracking off: accept immediately (same tick, so the yellow never paints);
            // tracking on: revisions stay pending, handled in the Review tab
            if (!trackChangesRef.current) clearAiHighlights(true)
          }
          runToolsRef.current.push({
            name: call.name,
            summary: execution.summary,
            isError: execution.isError,
            input: safeJsonInput(call.input),
            output: execution.output
              ? execution.output.slice(0, PERSIST_TOOL_FIELD_MAX)
              : undefined,
          })
          patchLastAssistant((last) => {
            // Swap out the running placeholder pushed by onToolStart (parse-fail calls have none)
            const tools = [...(last.tools ?? [])]
            if (tools.at(-1)?.running) tools.pop()
            return {
              tools: [
                ...tools,
                {
                  name: call.name,
                  summary: execution.summary,
                  isError: execution.isError,
                  output: execution.output
                    ? execution.output.slice(0, TOOL_OUTPUT_MAX_CHARS)
                    : undefined,
                  display: execution.display,
                },
              ],
            }
          })
        },
        onTurnEnd: () => {
          // ponytail: keep tools and thinking in one unified assistant response rather than fragmenting into multiple empty bubbles
        },
        onDone: ({ text, cancelled, turnLimit, stuck, truncated }) => {
          // module-level t: the loop instance is created only once; the component's t goes stale with the first-render closure
          const baseText = stuck
            ? [text, tModule('aiStopped')].filter(Boolean).join('\n\n')
            : turnLimit
              ? [text, tModule('aiTurnLimit')].filter(Boolean).join('\n\n')
              : text || (cancelled ? tModule('aiStopped') : '')
          const finalText = truncated
            ? [baseText, tModule('aiTruncatedNote')].filter(Boolean).join('\n\n')
            : baseText
          patchLastAssistant((last) => ({
            streaming: false,
            turnLimit,
            stuck,
            text: finalText || (last.tools?.length ? last.text : tModule('aiNoReply')),
            // A stop mid-tool can leave a running placeholder behind — drop it
            tools: last.tools?.filter((tl) => !tl.running),
          }))
          setBusy(false)
          // persist outside the updater (a double-invoked updater would write history twice); tools stores the whole run's full activity.
          // Edits-only runs (tools ran, no text) persist too, or the whole turn vanishes from the restored transcript
          if (!cancelled && (finalText || runToolsRef.current.length > 0)) {
            persistMessage('assistant', finalText, runToolsRef.current)
          }
        },
        onError: (error) => {
          setChat((prev) => {
            const next = [...prev]
            // the loop rolled this run's user message out of the model context — surface that
            for (let i = next.length - 1; i >= 0; i--) {
              const entry = next[i]!
              if (entry.role === 'user') {
                next[i] = { ...entry, undelivered: true }
                break
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
          setBusy(false)
        },
      },
    })
  }

  const pendingAutoRunRef = useRef<string | null>(null)

  useEffect(() => {
    if (!preset) return
    if (preset.autoRun) {
      if (isReady) {
        runWith(preset.text)
      } else {
        pendingAutoRunRef.current = preset.text
      }
    } else {
      setInput(preset.text)
      inputRef.current?.focus()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [preset?.nonce])

  useEffect(() => {
    if (isReady && pendingAutoRunRef.current) {
      const text = pendingAutoRunRef.current
      pendingAutoRunRef.current = null
      runWith(text)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isReady])

  // keep the scope hint & quick actions in sync with the editor selection
  useEffect(() => {
    const bump = () => setScopeTick((t) => t + 1)
    editor.on('selectionUpdate', bump)
    editor.on('update', bump)
    return () => {
      editor.off('selectionUpdate', bump)
      editor.off('update', bump)
    }
  }, [editor])

  // follow the stream, but stop yanking once the user scrolls up to read;
  // `open` dep: re-expanding lands on messages streamed while collapsed
  useEffect(() => {
    if (stickToBottomRef.current) {
      logRef.current?.scrollTo({ top: logRef.current.scrollHeight })
    }
  }, [chat, open])

  const onLogScroll = () => {
    const el = logRef.current
    if (!el) return
    stickToBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 48
  }

  const run = () => runWith(input.trim())

  /** Image attachments are read as base64 and go multimodal with this user message (≤5MB per image, max 20) */
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
        failures.push(result.error ?? t('aiImageReadFail', { name: att.name }))
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

  const runWith = (instruction: string, displayInstruction = instruction) => {
    if (!isReady) return
    const loop = loopRef.current
    if (!instruction || !loop || loop.busy) return
    setInput('')
    instructionRef.current = instruction
    lastInstructionRef.current = instruction
    runToolsRef.current = []
    stickToBottomRef.current = true
    const turnAttachments = attachmentsRef.current
    runAttachmentsRef.current = turnAttachments
    setChat((prev) => [
      ...prev,
      { role: 'user', text: displayInstruction, attachments: turnAttachments },
      { role: 'assistant', text: '', streaming: true },
    ])
    attachmentsRef.current = []
    setAttachments([])
    runStartedAtRef.current = Date.now()
    setBusy(true)
    persistMessage('user', instruction, undefined, turnAttachments)
    // a rejected image read must not strand the run (busy would stay true forever): degrade to a no-image send
    void collectImageAttachments(turnAttachments)
      .catch((): AgentImage[] => {
        setAttachNotice(t('aiImagesSendFailed'))
        window.setTimeout(() => setAttachNotice(null), 5000)
        return []
      })
      .then((images) => loop.run(instruction, images))
  }

  const cancel = () => loopRef.current?.cancel()

  const retry = () => runWith(lastInstructionRef.current)

  const continueRun = () => runWith(CONTINUE_INSTRUCTION, t('aiContinue'))

  const newChat = () => {
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

  /** Files pasted into the input: local paths use the normal attachment flow; image bytes are saved temporarily first. */
  const onPasteFiles = async (files: File[]) => {
    const paths: string[] = []
    for (const f of files) {
      const path = window.desktop.getPathForFile(f)
      if (path) {
        paths.push(path)
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

  const acceptChanges = () => {
    applyRevisionsBy(editorRef.current, AI_REVISION_AUTHOR, 'accept')
    clearAiHighlights()
  }

  const toggleTrackChanges = () => {
    const next = !trackChanges
    setTrackChanges(next)
    localStorage.setItem(TRACK_CHANGES_KEY, next ? '1' : '0')
    // switching off keeps nothing pending: accept whatever is still highlighted
    if (!next) acceptChanges()
  }

  const rollback = (snapshot: Snapshot) => {
    editor.commands.setContent(snapshot.json as never)
    setSnapshots((prev) => prev.filter((s) => s !== snapshot))
  }

  // collapsed: rail only — after all hooks, so the instance and its state survive
  if (!open) {
    return (
      <button className="ai-rail" title={t('appExpandAiPanel')} onClick={onExpand}>
        <span className="ai-rail-icon" aria-hidden="true">
          <AraSpark size={21} />
        </span>
      </button>
    )
  }

  const docText = editor?.state?.doc?.textContent?.trim() ?? ''
  const wordCount = docText ? docText.split(/\s+/).filter(Boolean).length : 0
  const fileName = filePath
    ? filePath.replace(/\\/g, '/').split('/').pop() || t('aiDocUntitled')
    : t('aiDocUntitled')

  return (
    <AiPanelShell
      ref={asideRef}
      style={{ width: '100%' }}
      className={`ai-panel${dragOver ? ' ai-panel-dragover' : ''}${resizing ? ' ai-panel-resizing' : ''}`}
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
        showNewChat={chat.length > 0}
        onNewChat={newChat}
        onCollapse={onCollapse}
        newChatLabel={t('aiNewChatTitle')}
        collapseLabel={t('aiCollapseTitle')}
        newChatIcon={<IconNewChat size={16} />}
        collapseIcon={<IconSidebarCollapse size={15} />}
        snapshotsCount={snapshots.length}
        isSnapshotsOpen={snapshotsOpen}
        onToggleSnapshots={() => setSnapshotsOpen((v) => !v)}
        live={busy}
      >
        <AiSnapshotsPopover
          isOpen={snapshotsOpen}
          onClose={() => setSnapshotsOpen(false)}
          snapshots={snapshots}
          onRollback={(s) => rollback(s)}
          title={t('aiSnapshotsTitle')}
          rollbackLabel={t('aiRollback')}
          busy={busy}
        />
      </AiPanelHeader>

      <div ref={logRef} className="ai-chat" onScroll={onLogScroll}>
        {/* past conversation (read-only transcript, not fed to the model), shown continuously with the current turn */}
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
              <AiWelcomeHub
                description={
                  docEmpty
                    ? `${t('aiEmptyDraftBody1')} ${t('aiEmptyDraftBody2')}`
                    : `${t('aiEmptyBody1')} ${t('aiEmptyBody2')}`
                }
                contextBadge={{
                  icon: <IconFileText size={12} />,
                  label: docEmpty
                    ? `${fileName} · ${t('aiDocBlank')}`
                    : `${fileName} · ${t('aiDocWords', { count: wordCount })}`,
                  title: filePath ?? fileName,
                }}
                cards={
                  docEmpty
                    ? [
                        {
                          id: 'weekly-report',
                          icon: <IconFileText size={16} />,
                          title: t('aiStarterWeeklyReport'),
                          subtitle: t('aiStarterWeeklyReportSub'),
                          onClick: () => {
                            setInput(t('aiStarterWeeklyReport'))
                            inputRef.current?.focus()
                          },
                        },
                        {
                          id: 'launch-post',
                          icon: <IconSparkles size={16} />,
                          title: t('aiStarterLaunchPost'),
                          subtitle: t('aiStarterLaunchPostSub'),
                          onClick: () => {
                            setInput(t('aiStarterLaunchPost'))
                            inputRef.current?.focus()
                          },
                        },
                        {
                          id: 'event-outline',
                          icon: <IconBulb size={16} />,
                          title: t('aiStarterEventOutline'),
                          subtitle: t('aiStarterEventOutlineSub'),
                          onClick: () => {
                            setInput(t('aiStarterEventOutline'))
                            inputRef.current?.focus()
                          },
                        },
                      ]
                    : [
                        {
                          id: 'summarize',
                          icon: <IconFileText size={16} />,
                          title: t('aiStarterSummarize'),
                          subtitle: t('aiStarterSummarizeSub'),
                          onClick: () => {
                            setInput(t('aiStarterSummarize'))
                            inputRef.current?.focus()
                          },
                        },
                        {
                          id: 'polish',
                          icon: <IconPen size={16} />,
                          title: t('aiStarterPolishAll'),
                          subtitle: t('aiStarterPolishAllSub'),
                          onClick: () => {
                            setInput(t('aiStarterPolishAll'))
                            inputRef.current?.focus()
                          },
                        },
                        {
                          id: 'continue',
                          icon: <IconSparkles size={16} />,
                          title: t('aiStarterContinue'),
                          subtitle: t('aiStarterContinueSub'),
                          onClick: () => {
                            setInput(t('aiStarterContinue'))
                            inputRef.current?.focus()
                          },
                        },
                      ]
                }
              />
            ) : (
              <AiUnconfiguredCard onConfigure={() => void window.desktop?.configureAi?.()} />
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
                            <AiTypingIndicator label={t('aiWorking')} />
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
                <div className="ai-msg-error">{t('aiErrorPrefix', { error: entry.error })}</div>
              )}
              {showToolbar && (
                <div className="ai-msg-toolbar">
                  {entry.text && (
                    <button
                      className="ai-msg-tool-btn"
                      onClick={() => copyMessage(entry.text, i)}
                      aria-label={t('aiCopyReplyTitle')}
                      data-tip={t('aiCopyReplyTitle')}
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
                      aria-label={t('aiRegenerateTitle')}
                      data-tip={t('aiRegenerateTitle')}
                    >
                      {/* IconRefresh glyph, restated for this 20px slot: the button CSS
                          stretches svgs to 20px, and the shared icon's 12px-pinned stroke
                          painted ~1.8px here — 1 canvas unit paints 1.25px, matching IconCopy */}
                      <svg
                        width="20"
                        height="20"
                        viewBox="0 0 16 16"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="1"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        aria-hidden
                      >
                        <path d="M 12.68 6.65 a 4.86 4.86 0 0 0 -9 -1.08 M 3.32 9.35 a 4.86 4.86 0 0 0 9 1.08" />
                        <path d="M 12.95 3.05 v 2.7 h -2.7 M 3.05 12.95 v -2.7 h 2.7" />
                      </svg>
                    </button>
                  )}
                </div>
              )}
              {entry.turnLimit && !entry.stuck && isLast && !busy && (
                <button className="ai-continue-btn" onClick={continueRun}>
                  {t('aiContinue')}
                </button>
              )}
            </div>
          )
        })}
      </div>

      {/* Reclaimed bottom space: version snapshots now live in AiPanelHeader popover */}

      <div className="ai-composer">
        {attachNotice && <div className="ai-attach-notice">{attachNotice}</div>}
        <AiComposer
          disabled={!isReady}
          disclaimer={t('aiAdvisoryDisclaimer')}
          header={
            attachments.length > 0 && (
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
                        title={t('aiRemoveAttachmentTitle')}
                        aria-label={t('aiRemoveAttachmentTitle')}
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
                        title={t('aiRemoveAttachmentTitle')}
                        aria-label={t('aiRemoveAttachmentTitle')}
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
          authGate={undefined}
          value={input}
          busy={busy}
          ariaLabel={t('aiInputPlaceholder')}
          placeholder={t('aiInputPlaceholder')}
          hintIdle={t('aiHintIdle')}
          hintBusy={t('aiHintBusy')}
          sendLabel={t('aiSend')}
          stopLabel={t('aiStop')}
          iconOnly
          textareaRef={inputRef}
          onChange={setInput}
          onSend={run}
          onStop={cancel}
          onPasteFiles={(files) => void onPasteFiles(files)}
          footerStart={
            <>
              <button
                className="ai-attach-btn"
                onClick={pickAttachments}
                title={t('aiAttachTitle')}
              >
                <IconAttach size={18} />
              </button>
              <button
                className={`ai-track-btn${trackChanges ? ' on' : ''}`}
                onClick={toggleTrackChanges}
                title={trackChanges ? t('aiTrackOnTitle') : t('aiTrackOffTitle')}
              >
                <span className="ai-track-dot" aria-hidden />
                {t('aiTrackChanges')}
              </button>
            </>
          }
        />
      </div>
    </AiPanelShell>
  )
}
