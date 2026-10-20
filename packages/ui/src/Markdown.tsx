import { Fragment, memo, useState, type ReactNode } from 'react'
import { IconJump } from './icons'

/**
 * Robust, lightweight, streaming-tolerant markdown renderer for AI panels.
 * Supports:
 * - Fenced code blocks (```lang ... ```) with syntax/copy button
 * - Markdown tables (| Col 1 | Col 2 |)
 * - Blockquotes (> quote)
 * - Headings (# through ######)
 * - Lists (ul, ol)
 * - Horizontal rules (---)
 * - Inline formatting (**bold**, *italic*, `code`, [link](url), [target](cell:target))
 * Tolerates partial/streaming input without throwing.
 */

const INLINE_RE = /(`[^`\n]+`|\*\*[^*\n]+?\*\*|\*[^*\n]+?\*|\[[^\]\n]+\]\([^)\n]+\))/g

function renderInline(text: string, onJumpTarget?: (target: string) => void): ReactNode[] {
  const out: ReactNode[] = []
  let last = 0
  let key = 0
  for (const m of text.matchAll(INLINE_RE)) {
    const i = m.index ?? 0
    if (i > last) out.push(text.slice(last, i))
    const tok = m[0] ?? ''
    if (tok.startsWith('`')) {
      out.push(
        <code key={key++} className="ai-inline-code">
          {tok.slice(1, -1)}
        </code>,
      )
    } else if (tok.startsWith('**')) {
      out.push(<strong key={key++}>{tok.slice(2, -2)}</strong>)
    } else if (tok.startsWith('[')) {
      const linkMatch = /^\[([^\]]+)\]\(([^)]+)\)$/.exec(tok)
      if (linkMatch && linkMatch[1] !== undefined && linkMatch[2] !== undefined) {
        const label = linkMatch[1]
        const rawHref = linkMatch[2].trim()
        const isJump = /^(?:cell:|slide:|jump:|#jump-)/i.test(rawHref)
        if (isJump) {
          const target = rawHref.replace(/^(?:cell:|slide:|jump:|#jump-)/i, '')
          out.push(
            <button
              key={key++}
              type="button"
              className="ai-jump-badge"
              onClick={() => onJumpTarget?.(target)}
              title={`Go to ${target}`}
            >
              <IconJump size={12} />
              <span>{label}</span>
            </button>,
          )
        } else {
          const isSafe = /^(?:https?:\/\/|mailto:|#|\/)/i.test(rawHref)
          if (isSafe) {
            out.push(
              <a
                key={key++}
                href={rawHref}
                target="_blank"
                rel="noopener noreferrer"
                className="ai-link"
              >
                {label}
              </a>,
            )
          } else {
            out.push(label)
          }
        }
      } else {
        out.push(tok)
      }
    } else {
      out.push(<em key={key++}>{tok.slice(1, -1)}</em>)
    }
    last = i + tok.length
  }
  if (last < text.length) out.push(text.slice(last))
  return out
}

function CodeBlock({
  language,
  code,
}: {
  language: string
  code: string
}): React.JSX.Element | null {
  const [copied, setCopied] = useState(false)

  // Orphan fence remnant (e.g. a stray "```" with no content): render nothing
  // instead of an empty chrome block with header + copy button.
  if (!code.trim()) return null

  const handleCopy = () => {
    void navigator.clipboard.writeText(code).then(() => {
      setCopied(true)
      setTimeout(() => setCopied(false), 1800)
    })
  }

  return (
    <div className="ai-code-block">
      <div className="ai-code-header">
        <span className="ai-code-lang">{language || 'code'}</span>
        <button
          type="button"
          className={`ai-code-copy-btn ${copied ? 'copied' : ''}`}
          onClick={handleCopy}
          title={copied ? 'Copied to clipboard' : 'Copy code'}
        >
          {copied ? (
            <>
              <svg
                width="13"
                height="13"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.5"
              >
                <polyline points="20 6 9 17 4 12" />
              </svg>
              <span>Copied</span>
            </>
          ) : (
            <>
              <svg
                width="13"
                height="13"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
              >
                <rect x="9" y="9" width="13" height="13" rx="2" ry="2" />
                <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
              </svg>
              <span>Copy</span>
            </>
          )}
        </button>
      </div>
      <pre className="ai-code-pre">
        <code>{code}</code>
      </pre>
    </div>
  )
}

export type MdBlock =
  | { kind: 'p'; lines: string[] }
  | { kind: 'ul'; items: string[] }
  | { kind: 'ol'; start?: number; items: string[] }
  | { kind: 'h'; level: number; text: string }
  | { kind: 'code'; language: string; code: string }
  | { kind: 'quote'; lines: string[] }
  | { kind: 'hr' }
  | { kind: 'table'; headers: string[]; rows: string[][] }

function isTableStart(lines: string[], idx: number): boolean {
  const line = lines[idx]?.trimEnd() ?? ''
  return (
    line.includes('|') &&
    line.trim().startsWith('|') &&
    line.trim().endsWith('|') &&
    idx + 1 < lines.length &&
    /^\|(\s*:?-+:?\s*\|)+$/.test(lines[idx + 1]?.trim() ?? '')
  )
}
const TRAILING_QUESTION_RE =
  /^(.*?[a-zÀ-žА-я一-鿿,;:.!?…)])\s+((?:what|when|where|which|who|whom|whose|why|how|is|are|was|were|do|does|did|can|could|will|would|should|shall|may|might|have|has|had|apa|kapan|di\s?mana|bagaimana|mengapa|kenapa|siapa|berapa|apakah|bisa|bisakah|mau|adakah)\b[^?!]*[?!]\s*)$/i

const TRAILING_CTA_RE =
  /^(.*?[a-zÀ-žА-я一-鿿,;:.!?…)])\s+((?:(?:click|tap|select|open|press|please|see|check|gunakan|klik|silakan|pilih|buka)\b).*[.!?]\s*)$/i

// Trailing question or call-to-action sentence glued to the last list item
// (e.g. model omitted the blank line before "What type...?" or "Click the action button...")
// splits off as its own paragraph.
function splitTrailingProse(body: string): { body: string; trailing?: string } {
  const qm = TRAILING_QUESTION_RE.exec(body)
  if (qm) {
    const head = qm[1]?.trim() ?? ''
    const tail = qm[2]?.trim() ?? ''
    if (head.length >= 8 && tail.length >= 20 && tail.split(/\s+/).length >= 5) {
      return { body: head, trailing: tail }
    }
  }
  const cm = TRAILING_CTA_RE.exec(body)
  if (cm) {
    const head = cm[1]?.trim() ?? ''
    const tail = cm[2]?.trim() ?? ''
    if (head.length >= 8 && tail.length >= 15 && tail.split(/\s+/).length >= 4) {
      return { body: head, trailing: tail }
    }
  }
  return { body }
}

const MAX_BLOCKS = 2_000

export function parseBlocks(text: string): MdBlock[] {
  const blocks: MdBlock[] = []
  // Normalize glued heading delimiters like "... ### 5.1" where newlines were omitted
  const normalizedText = text.replace(/([^\n])\s+(#{1,6}\s+)/g, '$1\n\n$2')
  const lines = normalizedText.split('\n')
  let i = 0
  let runningOlIndex = 0
  while (i < lines.length && blocks.length < MAX_BLOCKS) {
    const raw = lines[i] ?? ''
    const line = raw.trimEnd()

    // 1. Fenced Code Block
    const codeMatch = /^```([a-zA-Z0-9_-]*)/.exec(line.trim())
    if (codeMatch) {
      const language = codeMatch[1] ?? ''
      const codeLines: string[] = []
      i++
      while (i < lines.length) {
        const cur = lines[i] ?? ''
        if (cur.trim().startsWith('```')) {
          i++
          break
        }
        codeLines.push(cur)
        i++
      }
      blocks.push({ kind: 'code', language, code: codeLines.join('\n') })
      runningOlIndex = 0
      continue
    }

    // Blank line
    if (!line.trim()) {
      i++
      continue
    }

    // 2. Horizontal Rule
    if (/^(\*\*\*|---|___)$/.test(line.trim())) {
      blocks.push({ kind: 'hr' })
      runningOlIndex = 0
      i++
      continue
    }

    // 3. Headings (# H1, ## H2, etc.)
    const hMatch = /^(#{1,6})\s+(.*)$/.exec(line.trim())
    if (hMatch) {
      const level = hMatch[1]?.length || 1
      let hText = hMatch[2] ?? ''
      let pText = ''

      if (hText.length > 70) {
        const colonIdx = hText.indexOf(': ')
        if (colonIdx > 0 && colonIdx <= 60) {
          pText = hText.slice(colonIdx + 2).trim()
          hText = hText.slice(0, colonIdx).trim()
        } else {
          const periodMatch = /^(.*?[a-zA-Z0-9]{3,}\.)\s+([A-Z].*)$/.exec(hText)
          if (periodMatch && periodMatch[1] && periodMatch[2] && periodMatch[1].length <= 60) {
            hText = periodMatch[1].trim()
            pText = periodMatch[2].trim()
          } else {
            const gluedMatch = /^(.+?(?:\b\d{4}|[a-z]{3,}))([A-Z][a-z]+(?:\s+[a-z]+){2,}.*)$/.exec(
              hText,
            )
            if (gluedMatch && gluedMatch[1] && gluedMatch[2] && gluedMatch[1].length <= 60) {
              hText = gluedMatch[1].trim()
              pText = gluedMatch[2].trim()
            } else {
              const spaceIdx = hText.indexOf(' ', 35)
              if (spaceIdx > 0 && spaceIdx <= 65) {
                pText = hText.slice(spaceIdx + 1).trim()
                hText = hText.slice(0, spaceIdx).trim()
              } else if (hText.length > 120) {
                blocks.push({ kind: 'p', lines: [line.trim()] })
                runningOlIndex = 0
                i++
                continue
              }
            }
          }
        }
      }

      blocks.push({ kind: 'h', level, text: hText })
      if (pText) {
        blocks.push({ kind: 'p', lines: [pText] })
      }
      runningOlIndex = 0
      i++
      continue
    }

    // 4. Blockquote (> quote)
    if (line.trim().startsWith('>')) {
      const quoteLines: string[] = []
      while (
        i < lines.length &&
        (lines[i]?.trim().startsWith('>') ||
          (lines[i]?.trim() &&
            quoteLines.length > 0 &&
            !/^#{1,6}\s+/.test(lines[i]?.trim() ?? '') &&
            !lines[i]?.trim().startsWith('-')))
      ) {
        const qLine = (lines[i] ?? '').replace(/^\s*>\s?/, '')
        quoteLines.push(qLine)
        i++
      }
      blocks.push({ kind: 'quote', lines: quoteLines })
      runningOlIndex = 0
      continue
    }

    // 5. Table (| Col 1 | Col 2 |)
    if (isTableStart(lines, i)) {
      const splitRow = (str: string) =>
        str
          .trim()
          .slice(1, -1)
          .split('|')
          .map((c) => c.trim())

      const headers = splitRow(line)
      i += 2 // skip header and separator
      const rows: string[][] = []
      while (
        i < lines.length &&
        (lines[i]?.trim().startsWith('|') ?? false) &&
        (lines[i]?.trim().endsWith('|') ?? false)
      ) {
        rows.push(splitRow(lines[i] ?? ''))
        i++
      }
      blocks.push({ kind: 'table', headers, rows })
      runningOlIndex = 0
      continue
    }

    // 6. Unordered List
    const ulMatch = /^\s*[-*•]\s+(.*)$/.exec(line)
    if (ulMatch) {
      const items: string[] = []
      let trailing: string | null = null
      while (i < lines.length) {
        const itemMatch = /^\s*[-*•]\s+(.*)$/.exec(lines[i] ?? '')
        if (!itemMatch) break
        let body = itemMatch[1] ?? ''
        if (!/^\s*[-*•]\s+/.test(lines[i + 1] ?? '')) {
          const split = splitTrailingProse(body)
          body = split.body
          if (split.trailing) trailing = split.trailing
        }
        items.push(body)
        i++
      }
      blocks.push({ kind: 'ul', items })
      if (trailing) blocks.push({ kind: 'p', lines: [trailing] })
      continue
    }

    // 6b. Glued "intro: - item" on one line (e.g. "... dalam: - Word: ..."):
    // split the intro off so the first bullet joins the list below instead of
    // sticking to the paragraph above.
    // ponytail: colon-gated only — bare "intro - item" prose stays inline; widen if models glue without colons.
    const glueUl = /^(.*[:;])\s+([-*•])\s+(.+)$/.exec(line)
    if (glueUl && glueUl[1] && glueUl[2] && glueUl[3]) {
      const intro = glueUl[1].trim()
      if (intro && intro.length <= 500) {
        blocks.push({ kind: 'p', lines: [intro] })
        lines[i] = `${glueUl[2]} ${glueUl[3]}`
        runningOlIndex = 0
        continue
      }
    }

    // 7. Glued prose + first numbered item on one line (colon not required):
    // split the intro off so item 1 joins the list below instead of sticking
    // to the paragraph above. Mirrors 6b; guards keep bare numbers and long
    // lines inline.
    const glueOl = /^(.*?[a-zA-Z0-9\u00C0-\u024F\u1E00-\u1EFF:;,.])\s+(\d+)[.、)]\s+(.+)$/.exec(
      line,
    )
    if (glueOl && glueOl[1] && glueOl[2] && glueOl[3]) {
      const intro = glueOl[1].trim()
      // Only split when the prefix is prose, not itself a list item number.
      if (intro && !/^\d+[.、)]$/.test(intro) && intro.length <= 500) {
        blocks.push({ kind: 'p', lines: [intro] })
        lines[i] = `${glueOl[2]}. ${glueOl[3]}`
        runningOlIndex = 0
        continue
      }
    }

    // 8. Ordered List
    const olMatch = /^\s*(\d+)[.、)]\s+(.*)$/.exec(line)
    if (olMatch) {
      const rawStart = parseInt(olMatch[1] ?? '1', 10)
      let start = rawStart
      if (rawStart === 1 && runningOlIndex > 0) {
        start = runningOlIndex + 1
      }
      const items: string[] = []
      let trailing: string | null = null
      while (i < lines.length) {
        const itemMatch = /^\s*(\d+)[.、)]\s+(.*)$/.exec(lines[i] ?? '')
        if (!itemMatch) break
        let body = itemMatch[2] ?? ''
        if (!/^\s*(\d+)[.、)]\s+/.test(lines[i + 1] ?? '')) {
          const split = splitTrailingProse(body)
          body = split.body
          if (split.trailing) trailing = split.trailing
        }
        items.push(body)
        i++
      }
      runningOlIndex = start + items.length - 1
      blocks.push({ kind: 'ol', start, items })
      if (trailing) blocks.push({ kind: 'p', lines: [trailing] })
      continue
    }

    // 9. Paragraph
    const pLines: string[] = []
    while (
      i < lines.length &&
      lines[i]?.trim() &&
      !/^#{1,6}\s+/.test(lines[i]?.trim() ?? '') &&
      !lines[i]?.trim().startsWith('```') &&
      !lines[i]?.trim().startsWith('>') &&
      !isTableStart(lines, i) &&
      !/^\s*[-*•]\s+/.test(lines[i] ?? '') &&
      !/^\s*\d+[.、)]\s+/.test(lines[i] ?? '') &&
      !/^(\*\*\*|---|___)$/.test(lines[i]?.trim() ?? '')
    ) {
      pLines.push(lines[i] ?? '')
      i++
    }
    if (pLines.length > 0) {
      blocks.push({ kind: 'p', lines: pLines })
      runningOlIndex = 0
    } else {
      blocks.push({ kind: 'p', lines: [lines[i] ?? ''] })
      i++
      runningOlIndex = 0
    }
  }

  return blocks
}

// memoized: during streaming, every token re-renders the chat list — finished
// messages keep an unchanged `text` and must not re-parse their markdown
export const Markdown = memo(function Markdown({
  text,
  className,
  onJumpTarget,
}: {
  readonly text: string
  readonly className?: string | undefined
  readonly onJumpTarget?: ((target: string) => void) | undefined
}): React.JSX.Element {
  return (
    <div className={`ai-md${className ? ` ${className}` : ''}`}>
      {parseBlocks(text).map((b, i) => {
        if (b.kind === 'code') {
          return <CodeBlock key={i} language={b.language} code={b.code} />
        }
        if (b.kind === 'h') {
          const Tag = b.level <= 2 ? 'h4' : b.level === 3 ? 'h5' : 'h6'
          return (
            <Tag key={i} className={`ai-md-h ai-md-h${b.level}`}>
              {renderInline(b.text, onJumpTarget)}
            </Tag>
          )
        }
        if (b.kind === 'quote') {
          return (
            <blockquote key={i} className="ai-md-quote">
              {b.lines.map((ln, j) => (
                <Fragment key={j}>
                  {j > 0 && <br />}
                  {renderInline(ln, onJumpTarget)}
                </Fragment>
              ))}
            </blockquote>
          )
        }
        if (b.kind === 'hr') {
          return <hr key={i} className="ai-md-hr" />
        }
        if (b.kind === 'table') {
          return (
            <div key={i} className="ai-md-table-wrap">
              <table className="ai-md-table">
                <thead>
                  <tr>
                    {b.headers.map((h, j) => (
                      <th key={j}>{renderInline(h, onJumpTarget)}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {b.rows.map((row, rIdx) => (
                    <tr key={rIdx}>
                      {row.map((cell, cIdx) => (
                        <td key={cIdx}>{renderInline(cell, onJumpTarget)}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )
        }
        if (b.kind === 'ul') {
          return (
            <ul key={i} className="ai-md-ul">
              {b.items.map((it, j) => (
                <li key={j}>{renderInline(it, onJumpTarget)}</li>
              ))}
            </ul>
          )
        }
        if (b.kind === 'ol') {
          return (
            <ol key={i} className="ai-md-ol" start={b.start ?? 1}>
              {b.items.map((it, j) => (
                <li key={j}>{renderInline(it, onJumpTarget)}</li>
              ))}
            </ol>
          )
        }
        return (
          <p key={i} className="ai-md-p">
            {b.lines.map((ln, j) => (
              <Fragment key={j}>
                {j > 0 && <br />}
                {renderInline(ln, onJumpTarget)}
              </Fragment>
            ))}
          </p>
        )
      })}
    </div>
  )
})
