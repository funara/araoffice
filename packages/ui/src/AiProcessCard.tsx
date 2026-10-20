import type React from 'react'
import { useCallback, useEffect, useRef, useState } from 'react'

export interface ToolDisplay {
  kind: 'images' | 'links' | 'text'
  text?: string
  items?: Array<{ url: string; title?: string }>
}

export interface ToolActivityItem {
  name?: string | undefined
  summary: string
  running?: boolean | undefined
  isError?: boolean | undefined
  output?: string | undefined
  display?: ToolDisplay | undefined
}

export interface AiProcessCardProps {
  thinking?: string | undefined
  tools?: readonly ToolActivityItem[] | undefined
  isStreaming?: boolean | undefined
  durationSeconds?: number | undefined
  title?: string | undefined
  defaultExpanded?: boolean | undefined
  onOpenExternalUrl?: ((url: string) => void) | undefined
}

function getDomain(url: string): string {
  try {
    return new URL(url).hostname.replace(/^www\./, '')
  } catch {
    return ''
  }
}

const ALLOWED_EXTERNAL_PROTOCOLS: readonly string[] = ['http:', 'https:']

/**
 * Open a model-provided URL in the OS browser. Protocol allowlist is enforced
 * inline (same http/https gate as safeExternalUrl in electron-utils, which ui
 * cannot depend on); anything else returns before window.open.
 */
export function openExternalUrl(url: string): void {
  let protocol: string
  try {
    protocol = new URL(url).protocol
  } catch {
    return
  }
  if (!ALLOWED_EXTERNAL_PROTOCOLS.includes(protocol)) return
  window.open(url, '_blank', 'noreferrer')
}

function formatToolName(name?: string): string {
  if (!name) return 'Action'
  const FRIENDLY: Record<string, string> = {
    get_document_context: 'Reviewed document',
    read_blocks: 'Read document content',
    insert_content: 'Added content',
    replace_blocks: 'Updated content',
    insert_chart: 'Added chart',
    edit_chart: 'Edited chart',
    apply_commands: 'Formatted document',
    propose_operations: 'Updated spreadsheet',
    read_range: 'Checked cells',
    get_workbook_context: 'Checked workbook',
    generate_deck: 'Created presentation',
    edit_slide_script: 'Updated slide',
    web_search: 'Web search',
    fetch_url: 'Read webpage',
  }
  if (FRIENDLY[name]) return FRIENDLY[name]
  return name.replace(/[_-]+/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase())
}

export function AiProcessCard({
  thinking = '',
  tools = [],
  isStreaming = false,
  durationSeconds,
  title,
  defaultExpanded,
  onOpenExternalUrl,
}: AiProcessCardProps): React.JSX.Element | null {
  const hasThinking = Boolean(thinking && thinking.trim())
  const hasTools = tools.length > 0
  const anyRunning = isStreaming || tools.some((t) => t.running)

  // Auto-expand when execution starts; auto-collapse when execution completes
  const [expanded, setExpanded] = useState(defaultExpanded ?? anyRunning)
  const [expandedTools, setExpandedTools] = useState<Set<number>>(new Set())
  const [copiedThinking, setCopiedThinking] = useState(false)
  const prevRunningRef = useRef(anyRunning)
  const drawerRef = useRef<HTMLDivElement | null>(null)
  const thoughtRef = useRef<HTMLDivElement | null>(null)
  /** user scroll intent for the thought preview: scrolling up pauses the follow,
   * returning to the bottom resumes it (big streamed chunks would defeat a
   * near-bottom distance check, so the decision is made in onScroll instead) */
  const thoughtStickRef = useRef(true)

  useEffect(() => {
    if (anyRunning && !prevRunningRef.current) {
      setExpanded(true)
    } else if (!anyRunning && prevRunningRef.current) {
      setExpanded(false)
    }
    prevRunningRef.current = anyRunning
  }, [anyRunning])

  const onThoughtScroll = useCallback(() => {
    const el = thoughtRef.current
    if (!el) return
    thoughtStickRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 24
  }, [])

  // Autoscroll the drawer to the newest entry while the user stays near the
  // bottom; reading history above pauses the stickiness. The thought preview
  // follows the latest reasoning line while streaming unless the user scrolled up.
  useEffect(() => {
    if (!expanded) return
    const drawer = drawerRef.current
    if (drawer) {
      const nearBottom = drawer.scrollHeight - drawer.scrollTop - drawer.clientHeight < 80
      if (nearBottom) drawer.scrollTop = drawer.scrollHeight
    }
    const thought = thoughtRef.current
    if (thought && isStreaming && thoughtStickRef.current) {
      thought.scrollTop = thought.scrollHeight
    }
  }, [thinking, tools, isStreaming, expanded])

  const toggleTool = useCallback((idx: number) => {
    setExpandedTools((prev) => {
      const next = new Set(prev)
      if (next.has(idx)) next.delete(idx)
      else next.add(idx)
      return next
    })
  }, [])

  const copyReasoning = useCallback(
    (e: React.MouseEvent) => {
      e.stopPropagation()
      if (!thinking) return
      void navigator.clipboard.writeText(thinking).then(() => {
        setCopiedThinking(true)
        setTimeout(() => setCopiedThinking(false), 1800)
      })
    },
    [thinking],
  )

  if (!hasThinking && !hasTools && !isStreaming) return null

  // Bound the rendered thought: marathon sessions can stream tens of KB of
  // reasoning per turn, and keeping every turn's full text in the DOM compounds
  // renderer memory. The Copy button still copies the complete `thinking` string.
  const MAX_RENDERED_THOUGHT_CHARS = 6000
  const visibleThinking =
    thinking.length > MAX_RENDERED_THOUGHT_CHARS
      ? `…${thinking.slice(-MAX_RENDERED_THOUGHT_CHARS)}`
      : thinking

  const toolCount = tools.length
  const stepLabel = toolCount === 1 ? '1 step' : `${toolCount} steps`

  // Summary header label
  let headerLabel = title
  if (!headerLabel) {
    const errorCount = tools.filter((t) => t.isError).length
    const successCount = toolCount - errorCount
    if (anyRunning) {
      const runningTool = tools.find((t) => t.running)
      headerLabel = runningTool
        ? `Working: ${runningTool.summary || formatToolName(runningTool.name)}…`
        : 'Working…'
    } else if (errorCount === toolCount && toolCount > 0) {
      headerLabel = toolCount === 1 ? 'Action failed' : `${toolCount} actions failed`
    } else if (errorCount > 0) {
      headerLabel = `${successCount} of ${toolCount} steps completed`
    } else if (hasTools) {
      headerLabel = toolCount === 1 ? '1 step completed' : `${toolCount} steps completed`
    } else {
      headerLabel = 'Reasoning'
    }
  }

  return (
    <div className={`ai-process-box${expanded ? ' expanded' : ''}`}>
      <div
        className="ai-process-header"
        onClick={() => setExpanded((v) => !v)}
        role="button"
        tabIndex={0}
        aria-expanded={expanded}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault()
            setExpanded((v) => !v)
          }
        }}
      >
        <svg
          className={`ai-process-caret${expanded ? ' open' : ''}`}
          viewBox="0 0 16 16"
          width="12"
          height="12"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
          aria-hidden="true"
        >
          <polyline points="6 4 10 8 6 12" />
        </svg>
        {anyRunning && <span className="ai-process-spinner" aria-hidden="true" />}
        <span className="ai-process-title">{headerLabel}</span>
        {anyRunning && hasTools && <span className="ai-process-badge">{stepLabel}</span>}
        {typeof durationSeconds === 'number' && durationSeconds > 0 && (
          <span className="ai-process-duration">({durationSeconds.toFixed(1)}s)</span>
        )}
      </div>

      {/* Compact Action Trail: Glanceable chips showing tool statuses at a glance */}
      {hasTools && !expanded && (
        <div className="ai-action-trail">
          {tools.map((t, idx) => (
            <span
              key={idx}
              className={`ai-action-chip ${t.running ? 'is-running' : t.isError ? 'is-error' : 'is-done'}`}
              title={t.summary || formatToolName(t.name)}
              onClick={() => {
                setExpanded(true)
                toggleTool(idx)
              }}
            >
              {t.running ? (
                <span className="ai-action-chip-spinner" aria-hidden="true" />
              ) : t.isError ? (
                <span className="ai-action-chip-err" aria-hidden="true">
                  ✕
                </span>
              ) : (
                <svg
                  className="ai-action-chip-check"
                  width="11"
                  height="11"
                  viewBox="0 0 16 16"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2.2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  aria-hidden="true"
                >
                  <polyline points="3 8.5 6.5 12 13 4" />
                </svg>
              )}
              <span className="ai-action-chip-text">
                {t.summary ? t.summary : formatToolName(t.name)}
              </span>
            </span>
          ))}
        </div>
      )}

      <div className="ai-process-drawer" ref={drawerRef}>
        <div className="ai-process-timeline">
          {/* Thinking item */}
          {hasThinking && (
            <div className="ai-process-item is-thought">
              <div className="ai-process-item-content">
                <div className="ai-process-item-header">
                  <span className="ai-process-item-tag">Reasoning</span>
                  {thinking && (
                    <button
                      type="button"
                      className={`ai-process-copy-btn${copiedThinking ? ' copied' : ''}`}
                      onClick={copyReasoning}
                      title={copiedThinking ? 'Copied!' : 'Copy reasoning'}
                      aria-label="Copy reasoning"
                    >
                      {copiedThinking ? '✓ Copied' : 'Copy'}
                    </button>
                  )}
                </div>
                <div
                  ref={thoughtRef}
                  className="ai-process-thought-preview"
                  onScroll={onThoughtScroll}
                >
                  {visibleThinking}
                </div>
              </div>
            </div>
          )}

          {/* Tools items */}
          {tools.map((tool, idx) => {
            const hasDisplayData = Boolean(
              tool.display?.items?.length || (tool.display?.kind === 'text' && tool.display.text),
            )
            const hasOutput = !tool.running && (Boolean(tool.output) || hasDisplayData)
            const isToolOpen = expandedTools.has(idx)
            const statusClass = tool.running ? 'is-running' : tool.isError ? 'is-error' : 'is-tool'

            return (
              <div key={idx} className={`ai-process-item ${statusClass}`}>
                <div className="ai-process-item-content">
                  <div className="ai-process-item-header">
                    <span className="ai-process-tool-title">
                      {tool.summary || formatToolName(tool.name)}
                    </span>
                    {hasOutput && (
                      <button
                        type="button"
                        className="ai-process-toggle-btn"
                        aria-expanded={isToolOpen}
                        onClick={() => toggleTool(idx)}
                      >
                        {isToolOpen ? 'Hide' : 'Details'}
                      </button>
                    )}
                  </div>

                  {/* Expandable tool output / technical details */}
                  {hasOutput && isToolOpen && (
                    <div className="ai-process-tool-output">
                      {tool.name && (
                        <div className="ai-process-tool-meta">
                          <span className="ai-process-tool-code">{tool.name}</span>
                        </div>
                      )}
                      {tool.display?.kind === 'links' && tool.display.items && (
                        <div className="ai-tool-output-links">
                          {tool.display.items.map((item, linkIdx) => {
                            const domain = getDomain(item.url)
                            return (
                              <div key={linkIdx} className="ai-tool-output-link-row">
                                <button
                                  type="button"
                                  className="ai-tool-output-link-btn"
                                  onClick={(e) => {
                                    e.preventDefault()
                                    onOpenExternalUrl?.(item.url)
                                  }}
                                  title={`${item.title || item.url}\n${item.url}`}
                                >
                                  <span className="ai-tool-link-icon" aria-hidden="true">
                                    🌐
                                  </span>
                                  {domain && <span className="ai-tool-link-domain">{domain}</span>}
                                  <span className="ai-tool-link-title">
                                    {item.title || item.url}
                                  </span>
                                </button>
                              </div>
                            )
                          })}
                        </div>
                      )}

                      {tool.output && (
                        <pre className="ai-tool-output-pre">
                          {tool.output.slice(0, 1500)}
                          {tool.output.length > 1500 ? '…' : ''}
                        </pre>
                      )}
                    </div>
                  )}
                </div>
              </div>
            )
          })}

          {/* Active generation indicator while still streaming after tools have run */}
          {anyRunning && hasTools && !tools.some((t) => t.running) && (
            <div className="ai-process-item is-running">
              <div className="ai-process-item-content">
                <span className="ai-process-tool-summary ai-generating-glow">
                  Generating response…
                </span>
              </div>
            </div>
          )}

          {/* Completion state */}
          {!anyRunning && hasTools && (
            <div className="ai-process-item is-done">
              <div className="ai-process-item-content">
                <span className="ai-process-done-label">✓ Completed</span>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
