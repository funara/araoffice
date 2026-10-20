import type { AgentMessage, AgentToolCall, AgentToolDef } from '@araoffice/agent-core'
import { httpBodyDetail } from './http-error'
import { readSseStream } from './sse'
import type { AiChatResponse, AiProviderConfig } from './types'

/**
 * 'google' provider: the Google AI Studio Gemini REST API, called directly
 * from the main process. No SDK, no proxy — native fetch and a plain SSE
 * reader are enough for both the one-shot and the streaming path.
 */

const GOOGLE_API_ORIGIN = 'https://generativelanguage.googleapis.com/v1beta/'

interface GooglePart {
  text?: string
  inlineData?: { mimeType: string; data: string }
  functionCall?: { name: string; args?: Record<string, unknown> }
  functionResponse?: { name: string; response?: Record<string, unknown> }
  /** Gemini marks reasoning parts with thought: true */
  thought?: boolean
}

interface GoogleContent {
  role: 'user' | 'model'
  parts: GooglePart[]
}

export interface GoogleRequestBody {
  systemInstruction?: { parts: Array<{ text: string }> }
  contents: GoogleContent[]
  tools?: Array<{ functionDeclarations: Array<Record<string, unknown>> }>
  generationConfig?: { maxOutputTokens?: number }
}

/** Map agent history onto Gemini contents. Tool results ride back as function responses. */
export function toGoogleContents(messages: AgentMessage[]): GoogleContent[] {
  const contents: GoogleContent[] = []
  for (const message of messages) {
    if (message.role === 'user') {
      const parts: GooglePart[] = [{ text: message.text }]
      for (const image of message.images ?? []) {
        parts.push({ inlineData: { mimeType: image.mime, data: image.base64 } })
      }
      contents.push({ role: 'user', parts })
    } else if (message.role === 'assistant') {
      const parts: GooglePart[] = []
      if (message.text) parts.push({ text: message.text })
      for (const call of message.toolCalls ?? []) {
        parts.push({ functionCall: { name: call.name, args: call.input } })
      }
      if (parts.length) contents.push({ role: 'model', parts })
    } else {
      const parts: GooglePart[] = message.results.map((result) => ({
        functionResponse: { name: result.name, response: { output: result.output } },
      }))
      if (parts.length) contents.push({ role: 'user', parts })
    }
  }
  return contents
}

export function toGoogleTools(tools: AgentToolDef[]): Array<{
  functionDeclarations: Array<Record<string, unknown>>
}> {
  if (!tools.length) return []
  return [
    {
      functionDeclarations: tools.map((tool) => ({
        name: tool.name,
        description: tool.description,
        parameters: tool.inputSchema,
      })),
    },
  ]
}

/** Assemble the full request body for one Gemini call. */
export function buildGoogleBody(
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens?: number,
): GoogleRequestBody {
  const body: GoogleRequestBody = { contents: toGoogleContents(messages) }
  if (system) body.systemInstruction = { parts: [{ text: system }] }
  const declarations = toGoogleTools(tools)
  if (declarations.length) body.tools = declarations
  if (maxTokens && maxTokens > 0) body.generationConfig = { maxOutputTokens: maxTokens }
  return body
}

/**
 * Gemini model ids: `gemini-<major>.<minor>[-<variant>]`. The id reaches a
 * URL path, so it is validated against this shape before use — a hostile
 * settings file cannot turn it into a path or host escape.
 */
const GOOGLE_MODEL_RE = /^gemini-[0-9]+\.[0-9]+(?:-[a-z0-9]+)*$/

export function assertGoogleModel(model: string): string {
  const id = model.trim()
  if (!GOOGLE_MODEL_RE.test(id)) {
    throw new Error(`Invalid Gemini model id: ${model}`)
  }
  return id
}

/**
 * The only host this client ever talks to. The request URL is derived from
 * this constant origin plus a validated model id, so a tampered settings file
 * can never redirect the request to another host.
 */
const GOOGLE_API_BASE = new URL(GOOGLE_API_ORIGIN)

/** Build the request URL: constant origin + validated model id. */
export function endpointFor(model: string, streaming: boolean): URL {
  const suffix = streaming ? ':streamGenerateContent?alt=sse' : ':generateContent'
  try {
    return new URL(`models/${assertGoogleModel(model)}${suffix}`, GOOGLE_API_BASE)
  } catch (err) {
    throw new Error(`Invalid Gemini request URL for model ${model}`, { cause: err })
  }
}

function headers(config: AiProviderConfig): Record<string, string> {
  return {
    'content-type': 'application/json',
    'x-goog-api-key': config.apiKey ?? '',
  }
}

/** POST the Gemini call; throw with the endpoint's own error detail on non-2xx. */
async function post(
  config: AiProviderConfig,
  body: GoogleRequestBody,
  streaming: boolean,
  signal: AbortSignal,
): Promise<Response> {
  let response: Response
  try {
    response = await fetch(endpointFor(config.model, streaming), {
      method: 'POST',
      headers: headers(config),
      body: JSON.stringify(body),
      signal,
    })
  } catch (err) {
    if (signal.aborted) throw err
    throw new Error(
      `Google AI endpoint unreachable (${err instanceof Error ? err.message : String(err)})`,
      { cause: err },
    )
  }
  if (!response.ok) {
    const detail = httpBodyDetail(await response.text().catch(() => ''))
    throw new Error(`Google AI request failed (${response.status}): ${detail}`)
  }
  return response
}

/** Gemini reports the cap as MAX_TOKENS; the agent loop watches for max_tokens. */
export function normalizeGoogleStopReason(finishReason: string): string {
  if (finishReason === 'MAX_TOKENS') return 'max_tokens'
  return finishReason
}

/**
 * One Gemini stream chunk reduced to callbacks. Pure, so the delta/tool-call
 * rules are testable without a network.
 */
export function absorbGoogleChunk(json: unknown): {
  deltas: string[]
  thinking: string[]
  toolCalls: AgentToolCall[]
  stopReason?: string
} {
  const deltas: string[] = []
  const thinking: string[] = []
  const toolCalls: AgentToolCall[] = []
  const payload = json as {
    candidates?: Array<{
      content?: { parts?: GooglePart[] }
      finishReason?: string
    }>
  }
  const candidate = payload.candidates?.[0]
  if (!candidate) return { deltas, thinking, toolCalls }
  for (const part of candidate.content?.parts ?? []) {
    if (part.functionCall?.name) {
      toolCalls.push({
        id: `call_${part.functionCall.name}`,
        name: part.functionCall.name,
        input: part.functionCall.args ?? {},
      })
      continue
    }
    if (typeof part.text !== 'string' || !part.text) continue
    if (part.thought) thinking.push(part.text)
    else deltas.push(part.text)
  }
  const stopReason = candidate.finishReason
    ? normalizeGoogleStopReason(candidate.finishReason)
    : undefined
  return { deltas, thinking, toolCalls, ...(stopReason !== undefined ? { stopReason } : {}) }
}

/** Non-streaming one-shot completion. */
export async function chatGoogle(
  config: AiProviderConfig,
  system: string,
  user: string,
  signal?: AbortSignal,
): Promise<AiChatResponse> {
  const response = await post(
    config,
    buildGoogleBody(system, [{ role: 'user', text: user }], []),
    false,
    signal ?? new AbortController().signal,
  )
  const json = (await response.json()) as {
    candidates?: Array<{ content?: { parts?: GooglePart[] } }>
  }
  const text = (json.candidates?.[0]?.content?.parts ?? [])
    .filter((part) => typeof part.text === 'string' && !part.thought)
    .map((part) => part.text as string)
    .join('')
  return { ok: true, content: text }
}

/** Streaming completion: forwards deltas/thinking/tool calls onto the callbacks. */
export async function streamGoogle(
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  cb: {
    signal: AbortSignal
    onDelta: (text: string) => void
    onThinking?: (text: string) => void
    onToolCall: (call: AgentToolCall) => void
    onStopReason?: (reason: string) => void
    onActivity?: () => void
  },
  maxTokens?: number,
): Promise<void> {
  const response = await post(
    config,
    buildGoogleBody(system, messages, tools, maxTokens),
    true,
    cb.signal,
  )
  await readSseStream(response, (data) => {
    if (data === '[DONE]') return
    let json: unknown
    try {
      json = JSON.parse(data)
    } catch {
      return // keepalives/comments from proxies
    }
    cb.onActivity?.()
    const absorbed = absorbGoogleChunk(json)
    for (const thinking of absorbed.thinking) cb.onThinking?.(thinking)
    for (const delta of absorbed.deltas) cb.onDelta(delta)
    for (const call of absorbed.toolCalls) cb.onToolCall(call)
    if (absorbed.stopReason) cb.onStopReason?.(absorbed.stopReason)
  })
}
