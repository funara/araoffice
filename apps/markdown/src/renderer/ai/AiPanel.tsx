import { useEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { AgentLoop, composeSkills, type ToolDisplay } from '@araoffice/agent-core'
import type { AiSettings } from '@araoffice/ai-provider'
import {
  AiComposer,
  AiPanelHeader,
  AiPanelShell,
  AiProcessCard,
  AiSnapshotsPopover,
  AiTypingIndicator,
  AiUnconfiguredCard,
  AiWelcomeHub,
  extractThinking,
  IconCheck,
  IconCopy,
  IconFileText,
  IconNewChat,
  IconPen,
  IconRefresh,
  IconSidebarCollapse,
  IconSparkles,
  Markdown,
  useAiPanelWidth,
  useAiReady,
} from '@araoffice/ui'
import type { Editor } from '@tiptap/core'
import { aiLangDirective, t as tGlobal, useI18n } from '../i18n/locale'
import { CONTINUE_INSTRUCTION } from '@araoffice/agent-core'
import { clearAiHighlights } from '../editor/aiHighlight'
import { createMarkdownSkill } from './markdown-skill'
import { createSearchSkill } from './search-skill'
import { createElectronTransport } from './transport'

const PANEL_WIDTH_KEY = 'markdown-ai-panel-width'
const MAX_SNAPSHOTS = 20
const TOOL_OUTPUT_MAX_CHARS = 2000

interface ToolActivity {
  name: string
  summary: string
  /** still executing: rendered as a spinner chip, replaced in place when the tool finishes */
  running?: boolean
  isError?: boolean
  output?: string
  display?: ToolDisplay
}

interface ChatEntry {
  role: 'user' | 'assistant'
  text: string
  thinking?: string
  streaming?: boolean
  isError?: boolean
  turnLimit?: boolean
  stuck?: boolean
  /** the run failed and this user message was rolled back out of the model context */
  undelivered?: boolean
  tools?: ToolActivity[]
}

interface Snapshot {
  label: string
  time: string
  markdown: string
}

/** Ribbon preset instruction; a new nonce triggers one auto-send */
export interface AiPreset {
  text: string
  nonce: number
}

export interface MarkdownAiDeps {
  getEditor(): Editor | null
  /** full document body as markdown, for pre-mutation snapshots */
  getSnapshot(): string
  /** rollback: replace the document with a snapshot */
  restoreSnapshot(markdown: string): void
  /** fired when a run with at least one mutation finishes (auto-save hook) */
  onRunDone(mutated: boolean): void
}

export function AiPanel({
  deps,
  filePath,
  preset,
  onCollapse,
}: {
  deps: MarkdownAiDeps
  filePath: string | null
  preset?: AiPreset | null
  onCollapse: () => void
}): ReactElement {
  const { lang, t } = useI18n()
  const { isReady } = useAiReady({
    getAiSettings: window.markdownApi?.getAiSettings
      ? () => window.markdownApi.getAiSettings()
      : undefined,
    onAiSettingsChanged: window.markdownApi?.onAiSettingsChanged
      ? (h) => window.markdownApi.onAiSettingsChanged!(h)
      : undefined,
  })
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [prompt, setPrompt] = useState('')
  const [busy, setBusy] = useState(false)
  const [copiedIdx, setCopiedIdx] = useState<number | null>(null)
  const [snapshots, setSnapshots] = useState<Snapshot[]>([])
  const [snapshotsOpen, setSnapshotsOpen] = useState(false)
  const chatRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const stickToBottomRef = useRef(true)
  const { panelWidth, asideRef, resizing, startResize } = useAiPanelWidth(PANEL_WIDTH_KEY, 360)

  useEffect(() => {
    const dock = asideRef.current?.closest('.ai-dock') as HTMLElement | null
    dock?.style.setProperty('--ai-panel-width', `${panelWidth}px`)
  }, [panelWidth])

  const settingsRef = useRef<AiSettings | null>(null)
  const langRef = useRef(lang)
  langRef.current = lang
  const depsRef = useRef(deps)
  depsRef.current = deps
  const filePathRef = useRef(filePath)
  /** instruction of the in-flight run, labels its rollback snapshot */
  const runInstructionRef = useRef('')
  const runMutatedRef = useRef(false)
  /** tool activity of the whole run, for transcript persistence */
  const runToolsRef = useRef<ToolActivity[]>([])
  const chatIdsRef = useRef<{ projectId: string; chatId: string } | null>(null)
  /** messages sent before resolveChat returned, flushed once the chat id is known */
  const pendingPersistRef = useRef<
    Array<{ role: 'user' | 'assistant'; text: string; tools?: ToolActivity[] }>
  >([])

  const patchLast = (patch: Partial<ChatEntry> | ((last: ChatEntry) => Partial<ChatEntry>)) => {
    setChat((prev) => {
      const next = [...prev]
      const last = next[next.length - 1]
      if (!last || last.role !== 'assistant') return prev
      next[next.length - 1] = { ...last, ...(typeof patch === 'function' ? patch(last) : patch) }
      return next
    })
  }

  const persistMessage = (role: 'user' | 'assistant', text: string, tools?: ToolActivity[]) => {
    const ids = chatIdsRef.current
    if (!window.projectApi) return
    if (!ids) {
      pendingPersistRef.current.push({ role, text, tools })
      return
    }
    void window.projectApi
      .appendChat({
        projectId: ids.projectId,
        chatId: ids.chatId,
        role,
        text,
        ...(tools && tools.length > 0 ? { tools } : {}),
      })
      .catch(() => {
        /* persistence failures are silent */
      })
  }

  // The loop is built once; every mutable value goes through a ref getter
  const loopRef = useRef<AgentLoop<string> | null>(null)
  if (!loopRef.current) {
    loopRef.current = new AgentLoop<string>({
      transport: createElectronTransport(() => settingsRef.current!),
      maxTurns: null,
      maxToolCalls: null,
      skill: composeSkills('markdown+search', '', [
        createMarkdownSkill(() => depsRef.current.getEditor()),
        createSearchSkill(),
      ]),
      captureSnapshot: () => depsRef.current.getSnapshot(),
      systemSuffix: () => aiLangDirective(langRef.current),
      events: {
        onThinking: (chunk) => {
          patchLast((last) => ({
            thinking: (last.thinking ?? '') + chunk,
          }))
        },
        onText: (raw) => {
          const { thinking, text } = extractThinking(raw)
          patchLast((last) => ({
            text,
            ...(thinking ? { thinking: last.thinking ? last.thinking : thinking } : {}),
          }))
        },
        onToolStart: (call) => {
          // Live "running" chip: replaced in place by onToolExecuted
          patchLast((last) => ({
            tools: [
              ...(last.tools ?? []),
              { name: call.name, summary: call.name.replace(/[_-]+/g, ' '), running: true },
            ],
          }))
        },
        onToolExecuted: ({ call, execution, snapshotBefore }) => {
          if (execution.mutated) runMutatedRef.current = true
          if (snapshotBefore !== undefined) {
            const label = runInstructionRef.current.slice(0, 40)
            const time = new Date().toLocaleTimeString([], {
              hour: '2-digit',
              minute: '2-digit',
            })
            setSnapshots((prev) =>
              [...prev, { label, time, markdown: snapshotBefore }].slice(-MAX_SNAPSHOTS),
            )
          }
          const activity: ToolActivity = {
            name: call.name,
            summary: execution.summary,
            isError: execution.isError,
            output: execution.output?.slice(0, TOOL_OUTPUT_MAX_CHARS),
            display: execution.display,
          }
          runToolsRef.current.push(activity)
          patchLast((last) => {
            // Swap out the running placeholder pushed by onToolStart (parse-fail calls have none)
            const tools = [...(last.tools ?? [])]
            if (tools.at(-1)?.running) tools.pop()
            return { tools: [...tools, activity] }
          })
        },
        onTurnEnd: () => {
          // ponytail: keep tools and thinking in one unified assistant response rather than fragmenting into multiple empty bubbles
        },
        onDone: ({ text, cancelled, turnLimit, stuck }) => {
          const final = stuck
            ? [text, tGlobal('aiStopped')].filter(Boolean).join('\n\n')
            : turnLimit
              ? [text, tGlobal('aiTurnLimit')].filter(Boolean).join('\n\n')
              : text || (cancelled ? tGlobal('aiStopped') : '')
          patchLast((last) => ({
            streaming: false,
            turnLimit,
            stuck,
            text: final || (last.tools?.length ? last.text : tGlobal('aiNoReply')),
            // A stop mid-tool can leave a running placeholder behind — drop it
            tools: last.tools?.filter((tl) => !tl.running),
          }))
          persistMessage('assistant', final, runToolsRef.current)
          const editor = depsRef.current.getEditor()
          if (editor) clearAiHighlights(editor)
          depsRef.current.onRunDone(runMutatedRef.current)
          setBusy(false)
        },
        onError: (error) => {
          setChat((prev) => {
            const next = [...prev]
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
                text: error,
                isError: true,
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

  // ── chat-history persistence: bind to the file, restore prior transcript ──
  useEffect(() => {
    const api = window.projectApi
    if (!api) return
    const tempChatId = `unsaved-${Date.now()}`
    void api
      .resolveChat({ filePath: filePathRef.current ?? null, tempChatId })
      .then((ids) => {
        chatIdsRef.current = ids
        for (const msg of pendingPersistRef.current.splice(0)) {
          persistMessage(msg.role, msg.text, msg.tools)
        }
        return api.loadChat({ projectId: ids.projectId, chatId: ids.chatId, limit: 200 })
      })
      .then((msgs) => {
        if (msgs.length === 0) return
        // the user may have sent a message while history was loading — never
        // replace a live transcript (and don't clobber the loop context)
        let applied = false
        setChat((prev) => {
          if (prev.length > 0) return prev
          applied = true
          return msgs.map((m) => ({
            role: m.role,
            text: m.text,
            tools: m.tools?.map((tool) => ({
              name: tool.name,
              summary: tool.summary,
              isError: tool.isError,
              output: tool.output ? tool.output.slice(0, TOOL_OUTPUT_MAX_CHARS) : undefined,
            })),
          }))
        })
        if (applied && !loopRef.current?.busy) {
          loopRef.current?.restore(msgs.map((m) => ({ role: m.role, text: m.text })))
        }
      })
      .catch(() => {
        /* history load failures are silent */
      })
  }, [])

  /** after an untitled document's first save, bind the unsaved-* history to the real path */
  useEffect(() => {
    filePathRef.current = filePath
    const ids = chatIdsRef.current
    if (!window.projectApi || !ids || !filePath || !ids.chatId.startsWith('unsaved-')) return
    void window.projectApi
      .rebindChat({ projectId: ids.projectId, tempChatId: ids.chatId, newFilePath: filePath })
      .then((r) => {
        if (r?.chatId) chatIdsRef.current = r
      })
      .catch(() => {
        /* silent */
      })
  }, [filePath])

  useEffect(() => {
    if (stickToBottomRef.current) {
      chatRef.current?.scrollTo({ top: chatRef.current.scrollHeight })
    }
  }, [chat, busy])

  const onChatScroll = (): void => {
    const el = chatRef.current
    if (!el) return
    stickToBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 48
  }

  const send = (text: string): void => {
    if (!isReady) return
    const instruction = text.trim()
    const loop = loopRef.current
    if (!instruction || !loop || loop.busy) return
    stickToBottomRef.current = true
    runInstructionRef.current = instruction
    runMutatedRef.current = false
    runToolsRef.current = []
    setChat((prev) => [
      ...prev,
      { role: 'user', text: instruction },
      { role: 'assistant', text: '', streaming: true },
    ])
    setPrompt('')
    setBusy(true)
    persistMessage('user', instruction)
    void (async () => {
      try {
        const nextSettings = await window.markdownApi.getAiSettings()
        // The model is transport configuration, so stored settings pass through.
        settingsRef.current = nextSettings
        await loop.run(instruction)
      } catch (err) {
        patchLast({
          streaming: false,
          text: err instanceof Error ? err.message : String(err),
          isError: true,
        })
        setBusy(false)
      }
    })()
  }

  const stop = (): void => loopRef.current?.cancel()

  useEffect(() => {
    return () => {
      loopRef.current?.cancel()
      loopRef.current?.reset()
    }
  }, [])

  const retry = (): void => send(runInstructionRef.current)
  const continueRun = (): void => send(CONTINUE_INSTRUCTION)

  // ribbon presets auto-send; while a run is active they land in the composer instead
  const presetNonceRef = useRef(0)
  const pendingPresetRef = useRef<string | null>(null)

  useEffect(() => {
    if (!preset || preset.nonce === presetNonceRef.current) return
    presetNonceRef.current = preset.nonce
    if (loopRef.current?.busy) {
      setPrompt(preset.text)
    } else if (isReady) {
      send(preset.text)
    } else {
      pendingPresetRef.current = preset.text
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [preset, isReady])

  useEffect(() => {
    if (isReady && pendingPresetRef.current) {
      const text = pendingPresetRef.current
      pendingPresetRef.current = null
      send(text)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isReady])

  const copyMessage = (text: string, idx: number): void => {
    void navigator.clipboard.writeText(text)
    setCopiedIdx(idx)
    window.setTimeout(() => setCopiedIdx((cur) => (cur === idx ? null : cur)), 1200)
  }

  const rollback = (snapshot: Snapshot): void => {
    if (busy) return
    depsRef.current.restoreSnapshot(snapshot.markdown)
    setSnapshots((prev) => prev.filter((s) => s !== snapshot))
  }
  useEffect(() => {
    const onFocus = (): void => {
      window.setTimeout(() => inputRef.current?.focus(), 50)
    }
    window.addEventListener('ara:focus-ai-input', onFocus)
    return () => window.removeEventListener('ara:focus-ai-input', onFocus)
  }, [])

  return (
    <AiPanelShell
      ref={asideRef}
      className={`copilot${resizing ? ' ai-panel-resizing' : ''}`}
      style={{ width: '100%' }}
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
        onNewChat={() => {
          stop()
          loopRef.current?.reset()
          setBusy(false)
          setChat([])
          setPrompt('')
        }}
        onCollapse={onCollapse}
        newChatLabel={t('aiNewChat')}
        collapseLabel={t('aiCollapsePanel')}
        newChatIcon={<IconNewChat size={15} />}
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

      <div className="ai-chat" ref={chatRef} onScroll={onChatScroll}>
        {chat.length === 0 && (
          <div className="ai-chat-empty">
            {isReady ? (
              (() => {
                const mdText = deps.getSnapshot()?.trim() ?? ''
                const docEmpty = !mdText
                const fileName = filePath
                  ? filePath.replace(/\\/g, '/').split('/').pop() || 'Untitled.md'
                  : 'Untitled.md'
                return (
                  <AiWelcomeHub
                    description={t('aiEmptyBody')}
                    contextBadge={{
                      icon: <IconFileText size={12} />,
                      label: docEmpty
                        ? `${fileName} · Blank`
                        : `${fileName} · ${mdText.split(/\s+/).filter(Boolean).length} words`,
                      title: filePath ?? fileName,
                    }}
                    cards={[
                      {
                        id: 'md-research',
                        icon: <IconSparkles size={16} />,
                        title: t('aiQuickResearch'),
                        subtitle: t('aiQuickResearchPrompt'),
                        onClick: () => {
                          setPrompt(t('aiQuickResearchPrompt'))
                          inputRef.current?.focus()
                        },
                      },
                      {
                        id: 'md-draft',
                        icon: <IconFileText size={16} />,
                        title: t('aiQuickDraft'),
                        subtitle: t('aiQuickDraftPrompt'),
                        onClick: () => {
                          setPrompt(t('aiQuickDraftPrompt'))
                          inputRef.current?.focus()
                        },
                      },
                      {
                        id: 'md-polish',
                        icon: <IconPen size={16} />,
                        title: t('aiQuickPolish'),
                        subtitle: t('aiQuickPolishPrompt'),
                        onClick: () => {
                          setPrompt(t('aiQuickPolishPrompt'))
                          inputRef.current?.focus()
                        },
                      },
                    ]}
                  />
                )
              })()
            ) : (
              <AiUnconfiguredCard onConfigure={() => void window.markdownApi?.configureAi?.()} />
            )}
          </div>
        )}
        {chat.map((entry, i) => {
          if (entry.role === 'user') {
            return (
              <div key={i} className="ai-msg ai-msg-user">
                {entry.text}
                {entry.undelivered && (
                  <div className="ai-msg-undelivered">
                    {t('aiUndelivered')}
                    {!busy && (
                      <button className="ai-retry-btn" onClick={() => send(entry.text)}>
                        {t('aiRetry')}
                      </button>
                    )}
                  </div>
                )}
              </div>
            )
          }
          const hasTools = (entry.tools?.length ?? 0) > 0
          if (!entry.text && !entry.streaming && !hasTools) return null
          const isLast = i === chat.length - 1
          // Action row appears once per completed reply: on the turn's final segment only
          // (mid-turn segments have a following assistant entry; the live turn ends when !busy)
          const nextEntry = chat[i + 1]
          const turnEnded = nextEntry ? nextEntry.role === 'user' : !busy
          const showToolbar = !entry.streaming && turnEnded && !!entry.text && !entry.isError
          return (
            <div
              key={i}
              className={`ai-msg ai-msg-assistant${entry.isError ? ' ai-msg-error' : ''}${entry.streaming ? ' ai-msg-streaming' : ''}`}
            >
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
              {showToolbar && (
                <div className="ai-msg-toolbar">
                  <button
                    className="ai-msg-tool-btn"
                    onClick={() => copyMessage(entry.text, i)}
                    aria-label={t('aiCopyReplyTitle')}
                    data-tip={t('aiCopyReplyTitle')}
                  >
                    {copiedIdx === i ? <IconCheck size={14} /> : <IconCopy size={14} />}
                  </button>
                  {isLast && !busy && runInstructionRef.current && (
                    <button
                      className="ai-msg-tool-btn"
                      onClick={retry}
                      aria-label={t('aiRegenerateTitle')}
                      data-tip={t('aiRegenerateTitle')}
                    >
                      <IconRefresh size={14} />
                    </button>
                  )}
                </div>
              )}
              {entry.turnLimit && !entry.stuck && isLast && !busy && (
                <button className="ai-retry-btn" onClick={continueRun}>
                  {t('aiContinue')}
                </button>
              )}
            </div>
          )
        })}
      </div>

      {/* Reclaimed bottom space: version snapshots now live in AiPanelHeader popover */}

      <div className="ai-composer">
        <AiComposer
          authGate={undefined}
          disabled={!isReady}
          disclaimer="ARA Assistant suggestions are advisory, it can make mistakes."
          value={prompt}
          busy={busy}
          ariaLabel={t('aiComposerPlaceholder')}
          placeholder={t('aiComposerPlaceholder')}
          hintIdle={t('aiHintIdle')}
          hintBusy={t('aiHintBusy')}
          sendLabel={t('aiSend')}
          stopLabel={t('aiStop')}
          iconOnly
          textareaRef={inputRef}
          onChange={setPrompt}
          onSend={() => send(prompt)}
          onStop={stop}
        />
      </div>
    </AiPanelShell>
  )
}
