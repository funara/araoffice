import { useCallback, useEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { AgentLoop, type AgentRunResult } from '@araoffice/agent-core'
import { defaultAiSettings } from '@araoffice/ai-provider/providers'
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
  IconBulb,
  IconCheck,
  IconCopy,
  IconFileText,
  IconSparkles,
  Markdown,
  useAiPanelWidth,
  useAiReady,
} from '@araoffice/ui'
import {
  checkPromptInjection,
  extractMermaidFromChat,
  validateAndSanitizeMermaid,
} from './diagram-guardrail'
import { createDiagramSkill } from './diagram-skill'
import { createElectronTransport } from './transport'
import type { DiagramSkillDeps } from './tools'

export interface DiagramSnapshot {
  id: string
  time: string
  label: string
  mode: 'visual' | 'mermaid'
  content: string
}

const PANEL_WIDTH_KEY = 'diagram-ai-panel-width'

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
  undelivered?: boolean
  tools?: ToolActivity[]
}

export function AiPanel({
  deps,
  filePath,
  onCollapse,
  presetPrompt,
  snapshots = [],
  onRollback,
}: {
  deps: DiagramSkillDeps
  filePath?: string | null
  onCollapse: () => void
  presetPrompt?: string | null
  snapshots?: DiagramSnapshot[]
  onRollback?: (snapshot: DiagramSnapshot) => void
}): ReactElement {
  const { isReady } = useAiReady({
    getAiSettings: window.diagramApi?.getAiSettings
      ? () => window.diagramApi.getAiSettings()
      : undefined,
    onAiSettingsChanged: window.diagramApi?.onAiSettingsChanged
      ? (h) => window.diagramApi.onAiSettingsChanged!(h)
      : undefined,
  })
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [prompt, setPrompt] = useState('')
  const [busy, setBusy] = useState(false)
  const [snapshotsOpen, setSnapshotsOpen] = useState(false)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const chatRef = useRef<HTMLDivElement>(null)
  const { panelWidth, asideRef, resizing, startResize } = useAiPanelWidth(PANEL_WIDTH_KEY, 280)
  const [copiedIdx, setCopiedIdx] = useState<number | null>(null)

  const chatIdsRef = useRef<{ projectId: string; chatId: string } | null>(null)
  const pendingPersistRef = useRef<
    Array<{ role: 'user' | 'assistant'; text: string; tools?: ToolActivity[] }>
  >([])
  const filePathRef = useRef<string | null>(filePath ?? null)

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
      .catch(() => {})
  }
  useEffect(() => {
    const dock = asideRef.current?.closest('.ai-dock') as HTMLElement | null
    dock?.style.setProperty('--ai-panel-width', `${panelWidth}px`)
  }, [panelWidth])

  const settingsRef = useRef<AiSettings | null>(null)

  useEffect(() => {
    void window.diagramApi
      ?.getAiSettings()
      .then((s) => {
        if (!s) return
        settingsRef.current = s
      })
      .catch(() => {})
  }, [])

  const depsRef = useRef(deps)
  depsRef.current = deps

  /** one-shot omnibar handoffs already consumed — prevents infinite auto-run loops */
  const consumedPresetRef = useRef<string | null>(null)

  const patchLast = (patch: Partial<ChatEntry> | ((last: ChatEntry) => Partial<ChatEntry>)) => {
    setChat((prev) => {
      const next = [...prev]
      const last = next[next.length - 1]
      if (!last || last.role !== 'assistant') return prev
      next[next.length - 1] = { ...last, ...(typeof patch === 'function' ? patch(last) : patch) }
      return next
    })
  }

  /** set when the turn mutated the canvas via a tool (generate_mermaid, add_shape, …) */
  const diagramToolRanRef = useRef(false)

  /** Fallback: the model answered with a fenced Mermaid block instead of calling the
   * generation tool — recover it onto the canvas so the diagram isn't lost in chat. */
  const applyMissedDiagram = (result: AgentRunResult) => {
    if (result.cancelled || diagramToolRanRef.current) return
    const deps = depsRef.current
    if (deps.getMode() !== 'mermaid') return
    const missed = extractMermaidFromChat(result.text)
    if (!missed) return
    void validateAndSanitizeMermaid(missed.code).then((guard) => {
      if (!guard.valid) return
      deps.captureSnapshot?.('Apply diagram from chat')
      deps.setMermaidCode(guard.code)
      patchLast((last) => {
        if (last.role !== 'assistant' || !last.text?.includes(missed.block)) return {}
        return { text: last.text.replace(missed.block, '*Diagram applied to the canvas.*').trim() }
      })
    })
  }

  const loopRef = useRef<AgentLoop | null>(null)
  if (!loopRef.current) {
    loopRef.current = new AgentLoop({
      transport: createElectronTransport(() => settingsRef.current ?? defaultAiSettings()),
      skill: createDiagramSkill(() => depsRef.current),
      maxTurns: null,
      maxToolCalls: null,
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
          patchLast((last) => ({
            tools: [
              ...(last.tools ?? []),
              { name: call.name, summary: call.name.replace(/[_-]+/g, ' '), running: true },
            ],
          }))
        },
        onToolExecuted: ({ call, execution }) => {
          if (execution.mutated) diagramToolRanRef.current = true
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
        onDone: (result) => {
          patchLast((last) => {
            if (last.role === 'assistant' && last.text) {
              persistMessage('assistant', last.text, last.tools)
            }
            return { streaming: false }
          })
          setBusy(false)
          applyMissedDiagram(result)
        },
        onError: (err: string) => {
          patchLast({
            isError: true,
            text: err || 'Error executing request.',
            streaming: false,
          })
          setBusy(false)
        },
      },
    })
  }

  useEffect(() => {
    return () => {
      loopRef.current?.cancel()
      loopRef.current?.reset()
    }
  }, [])

  const send = useCallback(
    (textToSend: string) => {
      const trimmed = textToSend.trim()
      if (!isReady || !trimmed || busy) return

      const check = checkPromptInjection(trimmed)
      if (!check.safe) {
        setChat((prev) => [
          ...prev,
          { role: 'user', text: trimmed },
          {
            role: 'assistant',
            text: check.reason || 'Security Notice: Prompt matches restricted pattern.',
            isError: true,
            streaming: false,
          },
        ])
        return
      }

      const loop = loopRef.current
      if (!loop) return

      setPrompt('')
      setBusy(true)

      // Same behavior as every other app: the handover envelope is both the
      // bubble and the model instruction.
      persistMessage('user', trimmed)

      setChat((prev) => [
        ...prev,
        { role: 'user', text: trimmed },
        { role: 'assistant', text: '', streaming: true },
      ])

      void (async () => {
        try {
          // The model is transport configuration, so stored settings pass through.
          settingsRef.current = (await window.diagramApi?.getAiSettings()) ?? defaultAiSettings()
          diagramToolRanRef.current = false
          await loop.run(trimmed)
        } catch (err) {
          patchLast({
            isError: true,
            text: err instanceof Error ? err.message : String(err),
            streaming: false,
          })
          setBusy(false)
        }
      })()
    },
    [isReady, busy],
  )

  const stop = () => {
    loopRef.current?.cancel()
    setBusy(false)
  }

  const pendingPresetRef = useRef<string | null>(null)

  useEffect(() => {
    if (!presetPrompt) return
    // Consume the one-shot handoff exactly once: without the guard, StrictMode
    // remounts and provider-ready flips re-send the same prompt in a loop.
    if (consumedPresetRef.current === presetPrompt) return
    consumedPresetRef.current = presetPrompt
    // Same behavior as every other app: the handover envelope is the
    // instruction and the bubble.
    if (isReady) {
      send(presetPrompt)
    } else {
      pendingPresetRef.current = presetPrompt
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [presetPrompt, isReady])

  useEffect(() => {
    if (isReady && pendingPresetRef.current) {
      const pending = pendingPresetRef.current
      pendingPresetRef.current = null
      send(pending)
    }
  }, [isReady, send])
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
        if (!msgs || msgs.length === 0) return
        let applied = false
        setChat((prev) => {
          if (prev.length > 0) return prev
          applied = true
          return msgs.map((m) => ({
            role: m.role as 'user' | 'assistant',
            text: m.text,
            tools: m.tools?.map((tool) => ({
              name: tool.name,
              summary: tool.summary,
              isError: tool.isError,
              output: tool.output ? tool.output.slice(0, 2000) : undefined,
            })),
          }))
        })
        if (applied && !loopRef.current?.busy) {
          loopRef.current?.restore(
            msgs.map((m) => ({ role: m.role as 'user' | 'assistant', text: m.text })),
          )
        }
      })
      .catch(() => {})
  }, [])

  /** after an untitled diagram's first save, bind the unsaved-* history to the real path */
  useEffect(() => {
    filePathRef.current = filePath ?? null
    const ids = chatIdsRef.current
    if (!window.projectApi || !ids || !filePath || !ids.chatId.startsWith('unsaved-')) return
    void window.projectApi
      .rebindChat({ projectId: ids.projectId, tempChatId: ids.chatId, newFilePath: filePath })
      .then((r) => {
        if (r?.chatId) chatIdsRef.current = r
      })
      .catch(() => {})
  }, [filePath])

  return (
    <AiPanelShell
      ref={asideRef}
      style={{ width: `${panelWidth}px` }}
      className={resizing ? 'ai-panel-resizing' : ''}
    >
      <div
        className="ai-panel-resizer"
        onPointerDown={startResize}
        role="separator"
        aria-orientation="vertical"
        aria-label="ARA Assistant Resizer"
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
        newChatLabel="New chat"
        collapseLabel="Collapse panel"
        snapshotsCount={snapshots.length}
        isSnapshotsOpen={snapshotsOpen}
        onToggleSnapshots={() => setSnapshotsOpen((v) => !v)}
        live={busy}
      >
        <AiSnapshotsPopover
          isOpen={snapshotsOpen}
          onClose={() => setSnapshotsOpen(false)}
          snapshots={snapshots}
          onRollback={(s) => {
            setSnapshotsOpen(false)
            onRollback?.(s as DiagramSnapshot)
          }}
          busy={busy}
        />
      </AiPanelHeader>

      <div className="ai-chat" ref={chatRef}>
        {chat.length === 0 && (
          <div className="ai-chat-empty">
            {isReady ? (
              (() => {
                const fileName = filePath
                  ? filePath.replace(/\\/g, '/').split('/').pop() || 'Untitled'
                  : 'Untitled'
                const modeLabel = deps.getMode() === 'mermaid' ? 'Mermaid Code' : 'Draw.io Canvas'
                return (
                  <AiWelcomeHub
                    description="Describe a diagram and ARA will draft the structure directly onto the canvas."
                    contextBadge={{
                      icon: <IconFileText size={12} />,
                      label: `${fileName} · ${modeLabel}`,
                      title: filePath ?? fileName,
                    }}
                    cards={[
                      {
                        id: 'flowchart',
                        icon: <IconSparkles size={16} />,
                        title: 'Draw a flowchart',
                        subtitle: 'Create a process flow with decision steps and branches',
                        onClick: () => {
                          setPrompt('Draw a simple flowchart')
                          inputRef.current?.focus()
                        },
                      },
                      {
                        id: 'org-chart',
                        icon: <IconBulb size={16} />,
                        title: 'Make an org chart',
                        subtitle: 'Hierarchical team structure with roles and departments',
                        onClick: () => {
                          setPrompt('Make a simple org chart')
                          inputRef.current?.focus()
                        },
                      },
                      {
                        id: 'sequence-diag',
                        icon: <IconFileText size={16} />,
                        title: 'Create a sequence diagram',
                        subtitle: 'Message interaction between services, users, and databases',
                        onClick: () => {
                          setPrompt('Create a simple login sequence diagram')
                          inputRef.current?.focus()
                        },
                      },
                    ]}
                  />
                )
              })()
            ) : (
              <AiUnconfiguredCard onConfigure={() => void window.diagramApi?.configureAi?.()} />
            )}
          </div>
        )}
        {chat.map((entry, i) => {
          if (entry.role === 'user') {
            return (
              <div key={i} className="ai-msg ai-msg-user">
                {entry.text}
              </div>
            )
          }
          const displayThinking = entry.thinking
          const hasProcess = Boolean(displayThinking || (entry.tools && entry.tools.length > 0))
          return (
            <div
              key={i}
              className={`ai-msg ai-msg-assistant${entry.isError ? ' ai-msg-error' : ''}`}
            >
              {hasProcess && (
                <AiProcessCard
                  thinking={displayThinking}
                  tools={entry.tools}
                  isStreaming={entry.streaming}
                />
              )}
              {!entry.text && entry.streaming && !hasProcess ? (
                <span className="ai-typing-row">
                  <AiTypingIndicator label="Working" />
                </span>
              ) : (
                entry.text && <Markdown text={entry.text} />
              )}
              {!entry.streaming && entry.text && (
                <div className="ai-msg-toolbar">
                  <button
                    className="ai-msg-tool-btn"
                    onClick={() => {
                      void navigator.clipboard.writeText(entry.text)
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
            </div>
          )
        })}
      </div>

      <div className="ai-composer">
        <AiComposer
          disabled={!isReady}
          disclaimer="ARA Assistant suggestions are advisory, it can make mistakes."
          textareaRef={inputRef}
          value={prompt}
          busy={busy}
          ariaLabel="Ask ARA"
          placeholder="Ask ARA"
          hintIdle="Enter to send, Shift+Enter for newline"
          hintBusy="Esc to stop"
          sendLabel="Send"
          stopLabel="Stop"
          iconOnly
          footerStart={<></>}
          onChange={setPrompt}
          onSend={() => void send(prompt)}
          onStop={stop}
        />
      </div>
    </AiPanelShell>
  )
}
