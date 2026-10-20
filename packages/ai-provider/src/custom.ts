import type { AgentMessage, AgentToolCall, AgentToolDef } from '@araoffice/agent-core'
import { httpBodyDetail } from './http-error'
import { readSseStream } from './sse'
import type { AiChatResponse, AiProviderConfig } from './types'

/**
 * 'custom' provider: a user-configured OpenAI-compatible endpoint hit directly
 * from the main process (Ollama, LM Studio, OpenAI, or any relay). No env
 * config, no proxy script, no usage reporting — what the endpoint accepts is
 * what you get, and failures surface as plain request errors.
 */

interface OpenAiToolCallAcc {
  id: string
  name: string
  args: string
  /** OpenAI tool_calls deltas are unordered; the index is the stream's own key */
  emitted: boolean
}

export interface OpenAiChunkState {
  acc: Map<number, OpenAiToolCallAcc>
}

/** Chat-completions message in OpenAI wire format (chat/v1 chat.completions). */
type OpenAiMessage = {
  role: 'system' | 'user' | 'assistant' | 'tool'
  content: string | Array<Record<string, unknown>> | null
  tool_calls?: Array<{
    id: string
    type: 'function'
    function: { name: string; arguments: string }
  }>
  tool_call_id?: string
}

/** Map agent history onto OpenAI messages; names/ids survive round-trips via tool_call_id. */
export function toOpenAiMessages(system: string, messages: AgentMessage[]): OpenAiMessage[] {
  const out: OpenAiMessage[] = [{ role: 'system', content: system }]
  for (const message of messages) {
    if (message.role === 'user') {
      if (message.images?.length) {
        out.push({
          role: 'user',
          content: [
            { type: 'text', text: message.text },
            ...message.images.map((image) => ({
              type: 'image_url' as const,
              image_url: { url: `data:${image.mime};base64,${image.base64}` },
            })),
          ],
        })
      } else {
        out.push({ role: 'user', content: message.text })
      }
    } else if (message.role === 'assistant') {
      const calls = message.toolCalls?.map((call) => ({
        id: call.id,
        type: 'function' as const,
        function: { name: call.name, arguments: JSON.stringify(call.input) },
      }))
      out.push({
        role: 'assistant',
        content: message.text,
        ...(calls?.length ? { tool_calls: calls } : {}),
      })
    } else {
      for (const result of message.results) {
        out.push({ role: 'tool', content: result.output, tool_call_id: result.id })
      }
    }
  }
  return out
}

export function toOpenAiTools(tools: AgentToolDef[]): Array<{
  type: 'function'
  function: { name: string; description: string; parameters: Record<string, unknown> }
}> {
  return tools.map((tool) => ({
    type: 'function',
    function: { name: tool.name, description: tool.description, parameters: tool.inputSchema },
  }))
}

function endpoint(config: AiProviderConfig): string {
  const base = (config.baseUrl ?? '').replace(/\/+$/, '')
  return `${base}/chat/completions`
}

function headers(config: AiProviderConfig): HeadersInit {
  const headers: Record<string, string> = { 'content-type': 'application/json' }
  if (config.apiKey) headers.authorization = `Bearer ${config.apiKey}`
  return headers
}

/** POST the chat-completions call; throw with the endpoint's own error detail on non-2xx. */
async function post(
  config: AiProviderConfig,
  body: Record<string, unknown>,
  signal: AbortSignal,
): Promise<Response> {
  let response: Response
  try {
    response = await fetch(endpoint(config), {
      method: 'POST',
      headers: headers(config),
      body: JSON.stringify(body),
      signal,
    })
  } catch (err) {
    if (signal.aborted) throw err
    throw new Error(
      `Custom AI endpoint unreachable (${err instanceof Error ? err.message : String(err)})`,
      {
        cause: err,
      },
    )
  }
  if (!response.ok) {
    const detail = httpBodyDetail(await response.text().catch(() => ''))
    throw new Error(`Custom AI request failed (${response.status}): ${detail}`)
  }
  return response
}

/** Non-streaming one-shot completion. */
export async function chatCustom(
  config: AiProviderConfig,
  system: string,
  user: string,
  signal?: AbortSignal,
): Promise<AiChatResponse> {
  const response = await post(
    config,
    { model: config.model, messages: toOpenAiMessages(system, [{ role: 'user', text: user }]) },
    signal ?? new AbortController().signal,
  )
  const json = (await response.json()) as {
    choices?: Array<{ message?: { content?: string } }>
  }
  return { ok: true, content: json.choices?.[0]?.message?.content ?? '' }
}

/**
 * One SSE event reduced to stream callbacks. Pure so the delta/tool-call rules
 * are testable without a network. Tool calls stream fragment by fragment under
 * an index; a call is emitted once its arguments parse as JSON, or flagged
 * truncated when the stream ends without a parsable payload.
 */
export function absorbOpenAiChunk(
  state: OpenAiChunkState,
  json: unknown,
): {
  deltas: string[]
  thinking: string[]
  toolCalls: AgentToolCall[]
  stopReason?: string
} {
  const deltas: string[] = []
  const thinking: string[] = []
  const toolCalls: AgentToolCall[] = []
  let stopReason: string | undefined
  const choice = (json as { choices?: Array<Record<string, unknown>> }).choices?.[0]
  if (!choice)
    return { deltas, thinking, toolCalls, ...(stopReason !== undefined ? { stopReason } : {}) }
  const delta = (choice.delta ?? {}) as {
    content?: string
    reasoning_content?: string
    tool_calls?: Array<{
      index?: number
      id?: string
      function?: { name?: string; arguments?: string }
    }>
  }
  if (typeof delta.content === 'string' && delta.content.length) deltas.push(delta.content)
  if (typeof delta.reasoning_content === 'string' && delta.reasoning_content.length)
    thinking.push(delta.reasoning_content)
  for (const fragment of delta.tool_calls ?? []) {
    const index = fragment.index ?? 0
    let acc = state.acc.get(index)
    if (!acc) {
      acc = { id: fragment.id ?? `call_${index}`, name: '', args: '', emitted: false }
      state.acc.set(index, acc)
    }
    if (fragment.id) acc.id = fragment.id
    if (fragment.function?.name) acc.name = fragment.function.name
    if (typeof fragment.function?.arguments === 'string') acc.args += fragment.function.arguments
    const complete = emitIfArgsParse(acc)
    if (complete) toolCalls.push(complete)
  }
  const finish = choice.finish_reason
  if (typeof finish === 'string' && finish.length) stopReason = finish
  // An exhausted stream can leave arguments half-written: report them truncated
  // so the agent loop asks for a split instead of retrying broken JSON forever.
  for (const acc of state.acc.values()) {
    if (!acc.emitted && acc.name && finish) {
      const call: AgentToolCall = {
        id: acc.id,
        name: acc.name,
        input: safeJsonParse(acc.args),
        inputError: /[{[]/.test(acc.args) ? undefined : 'The tool input arrived empty',
        truncated: true,
      }
      acc.emitted = true
      toolCalls.push(call)
    }
  }
  return { deltas, thinking, toolCalls, ...(stopReason !== undefined ? { stopReason } : {}) }
}

function safeJsonParse(text: string): Record<string, unknown> {
  try {
    const parsed = JSON.parse(text) as unknown
    return typeof parsed === 'object' && parsed !== null ? (parsed as Record<string, unknown>) : {}
  } catch {
    return {}
  }
}

function emitIfArgsParse(acc: OpenAiToolCallAcc): AgentToolCall | null {
  if (acc.emitted || !acc.name || !/[}\]]\s*$/.test(acc.args)) return null
  try {
    const input = JSON.parse(acc.args) as unknown
    if (typeof input !== 'object' || input === null) return null
    acc.emitted = true
    return { id: acc.id, name: acc.name, input: input as Record<string, unknown> }
  } catch {
    return null
  }
}

/** Streaming completion: forwards deltas/thinking/tool calls onto the callbacks. */
export async function streamCustom(
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
): Promise<void> {
  const body: Record<string, unknown> = {
    model: config.model,
    messages: toOpenAiMessages(system, messages),
    stream: true,
  }
  if (tools.length) body.tools = toOpenAiTools(tools)
  // No max_tokens, no timeout ladder: the user asked for a direct pipe; the
  // renderer-side cancel + silence watchdog remain the only guards.
  const response = await post(config, body, cb.signal)
  const state: OpenAiChunkState = { acc: new Map() }
  await readSseStream(response, (payload) => {
    if (payload === '[DONE]') return
    let json: unknown
    try {
      json = JSON.parse(payload)
    } catch {
      return // keepalives/comments from non-OpenAI relays
    }
    cb.onActivity?.()
    const absorbed = absorbOpenAiChunk(state, json)
    for (const thinking of absorbed.thinking) cb.onThinking?.(thinking)
    for (const delta of absorbed.deltas) cb.onDelta(delta)
    for (const call of absorbed.toolCalls) cb.onToolCall(call)
    if (absorbed.stopReason) cb.onStopReason?.(absorbed.stopReason)
  })
}

/**
 * Query an OpenAI-compatible endpoint's GET /models to discover available models.
 * Returns string[] of model IDs. Tolerant of missing, 404, or non-standard responses.
 */
export async function fetchCustomModels(
  config: AiProviderConfig,
  signal?: AbortSignal,
): Promise<string[]> {
  const base = (config.baseUrl ?? '').trim().replace(/\/+$/, '')
  if (!base) return []
  const url = `${base}/models`
  try {
    const res = await fetch(url, {
      method: 'GET',
      headers: headers(config),
      signal: signal ?? AbortSignal.timeout(8000),
    })
    if (!res.ok) return []
    const json = (await res.json()) as unknown
    const rawList = Array.isArray(json)
      ? json
      : typeof json === 'object' &&
          json !== null &&
          'data' in json &&
          Array.isArray((json as { data: unknown }).data)
        ? (json as { data: unknown[] }).data
        : typeof json === 'object' &&
            json !== null &&
            'models' in json &&
            Array.isArray((json as { models: unknown }).models)
          ? (json as { models: unknown[] }).models
          : []
    return rawList
      .map((item) => {
        if (typeof item === 'string') return item.trim()
        if (
          typeof item === 'object' &&
          item !== null &&
          'id' in item &&
          typeof (item as { id: unknown }).id === 'string'
        ) {
          return (item as { id: string }).id.trim()
        }
        return ''
      })
      .filter((id) => id.length > 0)
  } catch {
    return []
  }
}
