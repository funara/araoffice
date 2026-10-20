import type {
  AgentStreamRequest,
  AgentToolCall,
  AgentToolDef,
  AgentTransport,
  AgentMessage,
} from './types'

/**
 * One streamed chunk pushed back over an Electron IPC bridge. Structurally
 * identical to ai-provider's AiStreamChunk; declared here so this package
 * stays dependency-free.
 */
export interface IpcStreamChunk {
  requestId: string
  /** 'ping' = wire-level keepalive; re-arms the silence watchdog and carries no payload */
  type: 'delta' | 'thinking' | 'tool-call' | 'done' | 'error' | 'ping'
  text?: string
  toolCall?: AgentToolCall
  error?: string
  /** machine-readable error cause; maps to the localized timeout/credits message */
  errorCode?: 'timeout' | 'credits'
  /** normalized stop reason on 'done' ('max_tokens' = cut off by the token limit) */
  stopReason?: string
}

export interface IpcStreamStart<S, A extends string = string> {
  requestId: string
  app: A
  settings: S
  system: string
  messages: AgentMessage[]
  tools: AgentToolDef[]
}

/**
 * Renderer-side silence watchdog: the main process re-arms it with keepalive
 * pings on wire activity, so firing means the turn is dead (main-process stall,
 * lost chunks) and the run must fail instead of leaving the UI busy forever.
 * Longer than the main-process idle timeout (180s) so that one (localized) wins.
 */
export const IPC_STREAM_SILENCE_TIMEOUT_MS = 240_000

export interface IpcTransportOptions<S, A extends string = string> {
  /** application identity forwarded to usage tracking */
  app?: A
  /** subscribe to stream chunks; returns the unsubscribe function */
  onStream(listener: (chunk: IpcStreamChunk) => void): () => void
  /** forward the start request to the main process; a returned promise reports handler failure */
  start(request: IpcStreamStart<S, A>): void | Promise<unknown>
  /** abort the in-flight turn in the main process */
  cancel(requestId: string): void
  getSettings(): S
  /** localized fallback when an error chunk carries no message */
  unknownErrorText(): string
  /** localized message for timeouts (errorCode 'timeout' and the silence watchdog) */
  timeoutErrorText?(): string
  /** localized message for exhausted credits (errorCode 'credits') */
  creditsErrorText?(): string
}
export interface ElectronAiBridge<S, A extends string = string> {
  onAiStream(listener: (chunk: IpcStreamChunk) => void): () => void
  aiStream(request: IpcStreamStart<S, A>): void | Promise<unknown>
  aiStreamCancel(requestId: string): void | Promise<unknown>
}

export interface ElectronTransportOptions<S, A extends string = string> {
  app: A
  bridge: ElectronAiBridge<S, A>
  getSettings(): S
  unknownErrorText: string | (() => string)
  timeoutErrorText?: string | (() => string)
  creditsErrorText?: string | (() => string)
}

function resolveText(
  value: string | (() => string) | undefined,
  fallback: () => string,
): () => string {
  if (typeof value === 'function') return value
  if (value !== undefined) return () => value
  return fallback
}

/**
 * Shared preload-bridge wiring: apps pass only their bridge object and texts.
 * Text args stay lazy (string or thunk) so per-language `t()` calls resolve at
 * error time, not at transport creation.
 */
export function createElectronTransport<S, A extends string = string>(
  options: ElectronTransportOptions<S, A>,
): AgentTransport {
  const unknownText = resolveText(options.unknownErrorText, () => 'Unknown error')
  const ipcOptions: IpcTransportOptions<S, A> = {
    app: options.app,
    onStream: (listener) => options.bridge.onAiStream(listener),
    start: (request) => options.bridge.aiStream(request),
    cancel: (requestId) => void options.bridge.aiStreamCancel(requestId),
    getSettings: options.getSettings,
    unknownErrorText: unknownText,
  }
  if (options.timeoutErrorText !== undefined)
    ipcOptions.timeoutErrorText = resolveText(options.timeoutErrorText, unknownText)
  if (options.creditsErrorText !== undefined)
    ipcOptions.creditsErrorText = resolveText(options.creditsErrorText, unknownText)
  return createIpcTransport(ipcOptions)
}

function userFacingError(
  error: string,
  unknownErrorText: string,
  timeoutErrorText: string,
): string {
  const normalized = error.toLowerCase()
  if (normalized.includes('aborted') || normalized.includes('timeout after ')) {
    return timeoutErrorText
  }
  return error
}

/**
 * AgentTransport over an Electron IPC bridge: the main process talks to the
 * LLM providers (avoids renderer CORS) and streams chunks back per requestId.
 * Each app wires in its own preload bridge and i18n via the options.
 */
export function createIpcTransport<S, A extends string = string>(
  options: IpcTransportOptions<S, A>,
): AgentTransport {
  const timeoutText = () => options.timeoutErrorText?.() ?? options.unknownErrorText()
  return {
    stream(request: AgentStreamRequest, cb) {
      const requestId = crypto.randomUUID()
      let settled = false
      let silenceTimer: ReturnType<typeof setTimeout> | undefined
      const settle = () => {
        settled = true
        clearTimeout(silenceTimer)
        unsubscribe()
      }
      const fail = (error: string) => {
        if (settled) return
        settle()
        cb.onError(userFacingError(error, options.unknownErrorText(), timeoutText()))
      }
      const armSilence = () => {
        clearTimeout(silenceTimer)
        silenceTimer = setTimeout(() => {
          options.cancel(requestId)
          fail(timeoutText())
        }, IPC_STREAM_SILENCE_TIMEOUT_MS)
      }
      const unsubscribe = options.onStream((chunk) => {
        if (chunk.requestId !== requestId || settled) return
        if (chunk.type === 'ping') {
          armSilence()
        } else if (chunk.type === 'delta') {
          armSilence()
          cb.onDelta(chunk.text ?? '')
        } else if (chunk.type === 'thinking') {
          armSilence()
          cb.onThinking?.(chunk.text ?? '')
        } else if (chunk.type === 'tool-call') {
          armSilence()
          if (chunk.toolCall) cb.onToolCall(chunk.toolCall)
        } else if (chunk.type === 'done') {
          settle()
          if (chunk.stopReason) cb.onStopReason?.(chunk.stopReason)
          cb.onDone()
        } else {
          settle()
          let error: string
          if (chunk.errorCode === 'timeout') error = timeoutText()
          else if (chunk.errorCode === 'credits')
            error = options.creditsErrorText?.() ?? chunk.error ?? options.unknownErrorText()
          else error = chunk.error ?? options.unknownErrorText()
          cb.onError(userFacingError(error, options.unknownErrorText(), timeoutText()))
        }
      })
      armSilence()
      try {
        // a rejected/thrown start would otherwise leave the run pending until the watchdog
        Promise.resolve(
          options.start({
            requestId,
            app: (options.app ?? 'unknown') as A,
            settings: options.getSettings(),
            system: request.system,
            messages: request.messages,
            tools: request.tools,
          }),
        ).catch((err: unknown) => {
          fail(err instanceof Error ? err.message : options.unknownErrorText())
        })
      } catch (err) {
        fail(err instanceof Error ? err.message : options.unknownErrorText())
      }
      return { cancel: () => options.cancel(requestId) }
    },
  }
}
