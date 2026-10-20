import { sanitizeModelIdentity } from '@araoffice/agent-core'
import { extractThinking } from '@araoffice/ui'
import type { AiStreamRequest } from '@araoffice/ai-provider'
import type { ActiveTabContext } from '../../../shared/tabs-api'
import { buildOmnibarSystemPrompt, isIndonesianText } from './omnibar-prompt'

export interface OmnibarMessage {
  role: 'user' | 'assistant'
  text: string
}

export interface OmnibarAiResult {
  text: string
  isStreaming: boolean
  error?: string
  followUps?: string[]
}

export interface OmnibarStreamCallbacks {
  onUpdate: (result: OmnibarAiResult) => void
  onError: (error: string) => void
  onDone: (result: OmnibarAiResult) => void
}

let activeStreamHandle: { cancel: () => void } | null = null

/** Keep long omnibar threads from blowing the model context window. */
const MAX_HISTORY_MESSAGES = 20
const MAX_MESSAGE_CHARS = 4000

/**
 * One assistant persona; which endpoint serves it is transport configuration,
 * and the transport already walks its own fallback chain. The omnibar
 * therefore retries the request only once more, on the same settings, and
 * never switches models itself.
 *
 * Thinking chunks reset the countdown so active reasoning is never cut off.
 */
export const OMNIBAR_FIRST_CHUNK_TIMEOUT_MS = 15_000
export const OMNIBAR_MAX_ATTEMPTS = 2

/**
 * Build the outgoing thread when the user sends a follow-up while the
 * previous answer is still streaming: freeze in-flight bubbles (they carry
 * answered text) instead of dropping them, drop only empty ones.
 * Pure — unit-testable without React.
 */
export function appendFollowUp(
  history: Array<{ role: 'user' | 'assistant'; text: string; isStreaming?: boolean }>,
  userText: string,
): Array<{ role: 'user' | 'assistant'; text: string; isStreaming?: boolean }> {
  const frozen = history
    .filter((m) => m.role !== 'assistant' || !m.isStreaming || m.text?.trim())
    .map((m) => (m.isStreaming ? { ...m, isStreaming: false } : m))
  return [...frozen, { role: 'user', text: userText }]
}
export function cancelOmnibarStream(): void {
  if (activeStreamHandle) {
    try {
      activeStreamHandle.cancel()
    } catch {
      /* ignore cancel error */
    }
    activeStreamHandle = null
  }
}

export function extractFollowUps(rawText: string): { cleanText: string; followUps: string[] } {
  if (!rawText) return { cleanText: '', followUps: [] }

  // Closing ::: must be whitespace-separated (newline or space); a bare `::`
  // (C++/Rust scopes, URLs) must not end the block.
  const followUpRegex = /:::followups([\s\S]*?)(?:\s:::[ \t]*(?:\r?\n|$)|$)/i
  const match = followUpRegex.exec(rawText)

  if (!match) {
    return { cleanText: rawText.trim(), followUps: [] }
  }

  const cleanText = rawText.replace(followUpRegex, '').trim()
  const block = (match[1] || '').trim()
  const followUps = block
    .split('\n')
    .map((line) => {
      let item = line.trim()
      // Drop lines that are only leftover closing markers.
      if (/^:+$/.test(item)) return ''
      // Strip a trailing standalone closing marker ("option :::").
      item = item.replace(/\s:::\s*$/, '').trim()
      // Strip list prefixes ("-", "*", "•", "1.", "1)") — twice for "- 1. foo" nesting.
      item = item
        .replace(/^[-*•\d.)\]]+\s*/, '')
        .replace(/^[-*•\d.)\]]+\s*/, '')
        .trim()
      // Strip template scaffolding the model sometimes emits around options
      // ("[Use standard]", "[Saya akan ...]:::", "option: ...", quotes).
      item = item
        .replace(/\s*:+\s*$/, '')
        .replace(/^(?:option|opsi|choice)\s*[:\-–—]\s*/i, '')
        .replace(/^\[+/, '')
        .replace(/\]+$/, '')
        .replace(/^["'“”‘’`]+|["'“”‘’`]+$/g, '')
        .trim()
      return item
    })
    .map((line) => sanitizeModelIdentity(line, { allowFallback: false }).trim())
    .filter(
      (line) =>
        line.length > 0 &&
        line.length <= 120 &&
        !/^(?:open\s+in|buka\s+di)\s+ara/i.test(line) &&
        !line.includes('& Write'),
    )
    .slice(0, 3)

  return { cleanText, followUps }
}

/** Timeout message language: thread wins over the UI setting (guard shared with the prompt). */
export async function executeOmnibarQuery(
  input: string | OmnibarMessage[],
  context: ActiveTabContext,
  uiLang = 'en',
  callbacks: OmnibarStreamCallbacks,
): Promise<() => void> {
  cancelOmnibarStream()
  // Last user text gates the identity/catalog chunks: ordinary turns run on
  const messages: OmnibarMessage[] =
    typeof input === 'string'
      ? [{ role: 'user', text: input }]
      : input.length > 0
        ? input
        : [{ role: 'user', text: '' }]

  const trimmedMessages = messages
    .filter((m) => m.role === 'user' || m.role === 'assistant')
    .filter((m) => typeof m.text === 'string' && m.text.trim().length > 0)
    .slice(-MAX_HISTORY_MESSAGES)
    .map((m) => ({
      role: m.role,
      text: m.text.length > MAX_MESSAGE_CHARS ? `${m.text.slice(0, MAX_MESSAGE_CHARS)}…` : m.text,
    }))
  if (trimmedMessages.length === 0 || trimmedMessages[trimmedMessages.length - 1].role !== 'user') {
    callbacks.onError('Empty prompt — type a question first')
    return () => {}
  }
  const lastUserQuery = [...trimmedMessages].reverse().find((m) => m.role === 'user')?.text ?? ''

  try {
    if (!window.aiOfficeAi) {
      callbacks.onError('AI bridge not initialized')
      return () => {}
    }

    const settings = await window.aiOfficeAi.getSettings()
    const provider = settings.provider ?? 'google'
    const system = buildOmnibarSystemPrompt(context, uiLang, lastUserQuery, provider)
    const aiApp: AiStreamRequest['app'] = context.kind === 'home' ? 'docs' : context.kind
    for (let attempt = 1; attempt <= OMNIBAR_MAX_ATTEMPTS; attempt += 1) {
      const lastAttempt = attempt >= OMNIBAR_MAX_ATTEMPTS
      const settled = await runOmnibarAttempt(
        { app: aiApp, settings, system, messages: trimmedMessages },
        callbacks,
      )
      if (settled) return () => cancelOmnibarStream()
      if (lastAttempt) {
        const isIndo =
          uiLang.startsWith('id') ||
          isIndonesianText(lastUserQuery) ||
          (typeof navigator !== 'undefined' && navigator.language?.startsWith('id'))
        const serviceName = provider === 'custom' ? 'OpenAI-Compatible' : 'Gemini AI'
        callbacks.onError(
          isIndo
            ? `${serviceName} membutuhkan waktu lebih lama untuk merespons. Silakan coba kirim ulang pertanyaan Anda.`
            : `${serviceName} is taking longer than usual to respond. Please try asking again.`,
        )
        return () => cancelOmnibarStream()
      }
    }
    return () => cancelOmnibarStream()
  } catch (err) {
    callbacks.onError(String(err))
    return () => {}
  }
}

/**
 * One attempt: 10s to first activity (thinking or answer), then stream to done.
 * Active thinking events refresh the 10s timer so reasoning models aren't cut off.
 * Returns true when the attempt settled (done/error); false on timeout
 * so the caller advances to the next leg (smart retry, then swift).
 */
function runOmnibarAttempt(
  request: Omit<AiStreamRequest, 'requestId'>,
  callbacks: OmnibarStreamCallbacks,
): Promise<boolean> {
  return new Promise((resolve) => {
    let accumulatedRaw = ''
    let done = false
    let firstChunkTimer: ReturnType<typeof setTimeout> | undefined

    const resetTimer = () => {
      if (firstChunkTimer) clearTimeout(firstChunkTimer)
      firstChunkTimer = setTimeout(() => {
        cancelOmnibarStream()
        finish(false)
      }, OMNIBAR_FIRST_CHUNK_TIMEOUT_MS)
      firstChunkTimer.unref?.()
    }

    const finish = (settled: boolean) => {
      if (done) return
      done = true
      if (firstChunkTimer) {
        clearTimeout(firstChunkTimer)
        firstChunkTimer = undefined
      }
      resolve(settled)
    }

    // Start 10s countdown for first activity (thinking or answer text)
    resetTimer()

    const handleChunk = (chunk: { type: string; text?: string; error?: string }) => {
      if (chunk.type === 'delta' && chunk.text) {
        accumulatedRaw += chunk.text
        // Thinking is stripped, never surfaced: omnibar shows the answer only.
        const { text } = extractThinking(accumulatedRaw)
        const { cleanText, followUps } = extractFollowUps(text)
        if (cleanText.trim()) {
          if (firstChunkTimer) {
            clearTimeout(firstChunkTimer)
            firstChunkTimer = undefined
          }
          callbacks.onUpdate({
            text: sanitizeModelIdentity(cleanText),
            isStreaming: true,
            followUps: followUps.length > 0 ? followUps : undefined,
          })
        }
      } else if (chunk.type === 'thinking' && chunk.text) {
        // Model is actively reasoning — keep stream alive by resetting the 10s inactivity countdown
        resetTimer()
      } else if (chunk.type === 'error') {
        activeStreamHandle = null
        callbacks.onError(chunk.error || 'AI request failed')
        finish(true)
      } else if (chunk.type === 'done') {
        activeStreamHandle = null
        const { text } = extractThinking(accumulatedRaw)
        const { cleanText, followUps } = extractFollowUps(text)
        if (!cleanText.trim()) {
          // Empty done (e.g. cancelled): let the caller retry or surface timeout.
          finish(false)
          return
        }
        const draft = sanitizeModelIdentity(cleanText)
        callbacks.onDone({
          text: draft,
          isStreaming: false,
          followUps: followUps.length > 0 ? followUps : undefined,
        })
        finish(true)
      }
    }

    try {
      activeStreamHandle = window.aiOfficeAi.stream(request, { onChunk: handleChunk })
    } catch (err) {
      callbacks.onError(String(err))
      finish(true)
    }
  })
}
