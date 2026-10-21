import { useEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { AgentLoop } from '@araoffice/agent-core'
import type { AiSettings } from '@araoffice/ai-provider'
import {
  AiComposer,
  AiPanelHeader,
  AiPanelShell,
  AiProcessCard,
  AiTypingIndicator,
  AiUnconfiguredCard,
  AiWelcomeHub,
  extractThinking,
  IconBulb,
  IconCheck,
  IconCopy,
  IconFileText,
  IconNewChat,
  IconSidebarCollapse,
  Markdown,
  useAiPanelWidth,
  useAiReady,
} from '@araoffice/ui'
import { CONTINUE_INSTRUCTION } from '@araoffice/agent-core'
import { createPdfSkill } from './pdf-skill'
import { createElectronTransport } from './transport'
import type { PdfSkillDeps } from './tools'
import { useI18n, t as tGlobal, aiLangDirective } from '../i18n/locale'
const PANEL_WIDTH_KEY = 'pdf-ai-panel-width'

interface ToolActivity {
  name: string
  summary: string
  running?: boolean
  isError?: boolean
  output?: string
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

type Phase = 'thinking' | 'replying' | 'working'

export function AiPanel({
  api,
  onCollapse,
  preset,
  onClearPreset,
}: {
  api: PdfSkillDeps
  onCollapse: () => void
  preset?: { text: string; nonce: number } | null
  onClearPreset?: () => void
}): ReactElement {
  const { lang, t } = useI18n()
  const { isReady } = useAiReady({
    getAiSettings: window.pdfApi?.getAiSettings ? () => window.pdfApi.getAiSettings() : undefined,
    onAiSettingsChanged: window.pdfApi?.onAiSettingsChanged
      ? (h) => window.pdfApi.onAiSettingsChanged!(h)
      : undefined,
  })
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [prompt, setPrompt] = useState('')
  const [busy, setBusy] = useState(false)
  const [phase, setPhase] = useState<Phase>('thinking')
  const chatRef = useRef<HTMLDivElement>(null)
  const stickToBottomRef = useRef(true)
  const { panelWidth, asideRef, resizing, startResize } = useAiPanelWidth(PANEL_WIDTH_KEY, 280)
  const [copiedIdx, setCopiedIdx] = useState<number | null>(null)

  // The .ai-dock wrapper owns the animated width (docs-style 180ms slide);
  // it tracks the resizable panel width through this variable
  useEffect(() => {
    const dock = asideRef.current?.closest('.ai-dock') as HTMLElement | null
    dock?.style.setProperty('--ai-panel-width', `${panelWidth}px`)
  }, [panelWidth])
  const settingsRef = useRef<AiSettings | null>(null)
  const langRef = useRef(lang)
  langRef.current = lang
  const apiRef = useRef(api)
  apiRef.current = api

  const patchLast = (patch: Partial<ChatEntry> | ((last: ChatEntry) => Partial<ChatEntry>)) => {
    setChat((prev) => {
      const next = [...prev]
      const last = next[next.length - 1]
      if (!last || last.role !== 'assistant') return prev
      next[next.length - 1] = { ...last, ...(typeof patch === 'function' ? patch(last) : patch) }
      return next
    })
  }

  // The loop is built once; every mutable value goes through a ref getter
  const loopRef = useRef<AgentLoop | null>(null)
  if (!loopRef.current) {
    const deps: PdfSkillDeps = {
      doc: () => apiRef.current.doc(),
      fileName: () => apiRef.current.fileName(),
      pageCount: () => apiRef.current.pageCount(),
      currentPage: () => apiRef.current.currentPage(),
      readOnly: () => apiRef.current.readOnly(),
      outline: () => apiRef.current.outline(),
      searchIndex: () => apiRef.current.searchIndex(),
      isDeleted: (i) => apiRef.current.isDeleted(i),
      gotoPage: (p) => apiRef.current.gotoPage(p),
      addMarkup: (type, idx, rects) => apiRef.current.addMarkup(type, idx, rects),
      formEdits: () => apiRef.current.formEdits(),
      applyFormEdit: (v) => apiRef.current.applyFormEdit(v),
      rotatePage: (idx, dir) => apiRef.current.rotatePage(idx, dir),
      deletePage: (idx) => apiRef.current.deletePage(idx),
    }
    loopRef.current = new AgentLoop({
      transport: createElectronTransport(() => settingsRef.current!),
      maxTurns: null,
      maxToolCalls: null,
      skill: createPdfSkill(deps),
      systemSuffix: () => aiLangDirective(langRef.current),
      events: {
        onThinking: (chunk) => {
          setPhase('replying')
          patchLast((last) => ({
            thinking: (last.thinking ?? '') + chunk,
          }))
        },
        onText: (raw) => {
          setPhase('replying')
          const { thinking, text } = extractThinking(raw)
          patchLast((last) => ({
            text,
            ...(thinking ? { thinking: last.thinking ? last.thinking : thinking } : {}),
          }))
        },
        onToolStart: (call) => {
          setPhase('working')
          patchLast((last) => ({
            tools: [
              ...(last.tools ?? []),
              { name: call.name, summary: call.name.replace(/[_-]+/g, ' '), running: true },
            ],
          }))
        },
        onToolExecuted: ({ call, execution }) => {
          setPhase('working')
          patchLast((last) => {
            const tools = [...(last.tools ?? [])]
            if (tools.at(-1)?.running) tools.pop()
            return {
              tools: [
                ...tools,
                {
                  name: call.name,
                  summary: execution.summary,
                  isError: execution.isError,
                  output: execution.output?.slice(0, 2000),
                },
              ],
            }
          })
        },
        onTurnEnd: () => {
          setPhase('thinking')
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
            tools: last.tools?.filter((tool) => !tool.running),
          }))
          setBusy(false)
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
              next[next.length - 1] = { ...last, streaming: false, text: error, isError: true }
            }
            return next
          })
          setBusy(false)
        },
      },
    })
  }

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
    setChat((prev) => [
      ...prev,
      { role: 'user', text: instruction },
      { role: 'assistant', text: '', streaming: true },
    ])
    setPrompt('')
    setBusy(true)
    setPhase('thinking')
    void (async () => {
      try {
        const nextSettings = await window.pdfApi.getAiSettings()
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
  const continueRun = (): void => send(CONTINUE_INSTRUCTION)

  useEffect(() => {
    if (preset?.text && isReady) {
      send(preset.text)
      onClearPreset?.()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [preset?.nonce, isReady])

  useEffect(() => {
    return () => {
      loopRef.current?.cancel()
      loopRef.current?.reset()
    }
  }, [])

  const typingLabel = phase === 'replying' ? t('aiReplying') : t('aiWorking')

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
        live={busy}
      />

      <div className="ai-chat" ref={chatRef} onScroll={onChatScroll}>
        {chat.length === 0 && (
          <div className="ai-chat-empty">
            {isReady ? (
              <AiWelcomeHub
                description={t('aiEmptyBody')}
                contextBadge={{
                  icon: <IconFileText size={12} />,
                  label: `${api.fileName() || 'Document.pdf'} · ${api.pageCount()} pages`,
                }}
                cards={[
                  {
                    id: 'pdf-summary',
                    icon: <IconFileText size={16} />,
                    title: t('aiQuickSummary'),
                    subtitle: t('aiQuickSummaryPrompt'),
                    onClick: () => send(t('aiQuickSummaryPrompt')),
                  },
                  {
                    id: 'pdf-key-points',
                    icon: <IconBulb size={16} />,
                    title: t('aiQuickKeyPoints'),
                    subtitle: t('aiQuickKeyPointsPrompt'),
                    onClick: () => send(t('aiQuickKeyPointsPrompt')),
                  },
                ]}
              />
            ) : (
              <AiUnconfiguredCard onConfigure={() => void window.pdfApi?.configureAi?.()} />
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
          const isLast = i === chat.length - 1
          return (
            <div
              key={i}
              className={`ai-msg ai-msg-assistant${entry.isError ? ' ai-msg-error' : ''}`}
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
                    {displayText && <Markdown text={displayText} />}
                    {!entry.streaming && displayText && (
                      <div className="ai-msg-toolbar">
                        <button
                          className="ai-msg-tool-btn"
                          onClick={() => {
                            void navigator.clipboard.writeText(displayText)
                            setCopiedIdx(i)
                            setTimeout(() => setCopiedIdx(null), 1500)
                          }}
                          aria-label="Copy reply"
                          title="Copy reply"
                        >
                          {copiedIdx === i ? <IconCheck size={14} /> : <IconCopy size={14} />}
                        </button>
                      </div>
                    )}
                  </>
                )
              })()}
              {entry.turnLimit && !entry.stuck && isLast && !busy && (
                <button className="ai-retry-btn" onClick={continueRun}>
                  {t('aiContinue')}
                </button>
              )}
            </div>
          )
        })}
        {/* In-progress state: a standalone three-dot row at the end of the stream, kept until done */}
        {busy && <AiTypingIndicator label={typingLabel} />}
      </div>

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
          onChange={setPrompt}
          onSend={() => send(prompt)}
          onStop={stop}
        />
      </div>
    </AiPanelShell>
  )
}
