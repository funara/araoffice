import type React from 'react'
import { useEffect, useRef, useState } from 'react'

export interface AiThinkingBlockProps {
  thinking: string
  isStreaming?: boolean | undefined
  defaultExpanded?: boolean | undefined
}

export function extractThinking(text: string): { thinking: string; text: string } {
  if (!text) return { thinking: '', text: '' }
  const match = /<think>([\s\S]*?)(?:<\/think>|$)/i.exec(text)
  const rawText = text.replace(
    /<(?:tool_call|function_call)>[\s\S]*?(?:<\/(?:tool_call|function_call)>|$)/gi,
    '',
  )
  if (!match || match[1] === undefined) {
    return { thinking: '', text: rawText.trim() }
  }
  const thinking = match[1].trim()
  const cleanText = rawText.replace(/<think>[\s\S]*?(?:<\/think>|$)/i, '').trim()
  return { thinking, text: cleanText }
}

export function AiThinkingBlock({
  thinking,
  isStreaming = false,
  defaultExpanded = false,
}: AiThinkingBlockProps): React.JSX.Element | null {
  const [expanded, setExpanded] = useState(isStreaming || defaultExpanded)
  const [copied, setCopied] = useState(false)
  const prevStreamingRef = useRef(isStreaming)
  const textRef = useRef<HTMLDivElement | null>(null)

  // ponytail: auto-expand during streaming, auto-collapse when processing completes
  useEffect(() => {
    if (isStreaming && !prevStreamingRef.current) {
      setExpanded(true)
    } else if (!isStreaming && prevStreamingRef.current) {
      setExpanded(false)
    }
    prevStreamingRef.current = isStreaming
  }, [isStreaming])

  // auto-scroll to latest reasoning line while streaming
  useEffect(() => {
    if (isStreaming && expanded && textRef.current) {
      textRef.current.scrollTop = textRef.current.scrollHeight
    }
  }, [thinking, isStreaming, expanded])

  const copyReasoning = (e: React.MouseEvent) => {
    e.stopPropagation()
    if (!thinking) return
    void navigator.clipboard.writeText(thinking).then(() => {
      setCopied(true)
      setTimeout(() => setCopied(false), 1800)
    })
  }

  if (!thinking && !isStreaming) return null

  return (
    <div className={`ai-thinking-block${expanded ? ' ai-thinking-expanded' : ''}`}>
      <button
        type="button"
        className="ai-thinking-toggle"
        onClick={() => setExpanded((v) => !v)}
        aria-expanded={expanded}
      >
        <span className="ai-thinking-icon" aria-hidden>
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            <path d="M12 2a7 7 0 0 1 7 7c0 2.38-1.19 4.47-3 5.74V17a2 2 0 0 1-2 2H10a2 2 0 0 1-2-2v-2.26C6.19 13.47 5 11.38 5 9a7 7 0 0 1 7-7z" />
            <path d="M9 21h6" />
          </svg>
        </span>
        <span className="ai-thinking-title">
          {isStreaming && !thinking ? 'Thinking…' : 'Thinking'}
        </span>
        {thinking && (
          <span
            className={`ai-thinking-copy-btn${copied ? ' copied' : ''}`}
            onClick={copyReasoning}
            title={copied ? 'Copied!' : 'Copy reasoning'}
            aria-label={copied ? 'Copied!' : 'Copy reasoning'}
          >
            {copied ? (
              <svg
                width="12"
                height="12"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.5"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <polyline points="20 6 9 17 4 12" />
              </svg>
            ) : (
              <svg
                width="12"
                height="12"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <rect x="9" y="9" width="13" height="13" rx="2" ry="2" />
                <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
              </svg>
            )}
          </span>
        )}
        {isStreaming && <span className="ai-thinking-pulse" aria-hidden />}
        <span className="ai-thinking-chevron" aria-hidden>
          <svg
            width="12"
            height="12"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2.2"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            <polyline points="6 9 12 15 18 9" />
          </svg>
        </span>
      </button>
      {expanded && (
        <div className="ai-thinking-content">
          <div ref={textRef} className="ai-thinking-text">
            {thinking}
          </div>
        </div>
      )}
    </div>
  )
}
