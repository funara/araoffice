import type { AgentMessage, AgentToolCall, AgentToolDef } from '@araoffice/agent-core'

/**
 * Direct-connect providers only. ARAOffice talks to the model endpoint the
 * user configured — there is no proxy service, no account, no telemetry.
 */
export type AiProviderId = 'google' | 'custom'

export interface AiProviderConfig {
  model: string
  /** OpenAI-compatible API root, e.g. https://api.openai.com/v1 — 'custom' provider only */
  baseUrl?: string
  /** stored plaintext by design; never returned to the renderer (redacted reads) */
  apiKey?: string
  /** Indicates to the renderer that an API key is stored in settings without exposing the plaintext key */
  hasApiKey?: boolean
}

export interface AiModelMeta {
  id: string
  label: string
  description: string
  bestFor: string
  capabilities: string[]
}

export interface AiProviderMeta {
  id: AiProviderId
  label: string
  models: string[]
  modelDetails?: AiModelMeta[]
  defaultModel: string
}

export interface AiSettings {
  provider: AiProviderId
  providers: { google: AiProviderConfig; custom: AiProviderConfig }
}

export interface AiChatRequest {
  settings: AiSettings
  system: string
  user: string
}

export interface AiChatResponse {
  ok: boolean
  content?: string
  error?: string
}

export type AiApp = 'docs' | 'sheets' | 'slides' | 'pdf' | 'markdown' | 'diagram'

export interface AiStreamRequest {
  requestId: string
  app: AiApp
  settings: AiSettings
  system: string
  messages: AgentMessage[]
  tools?: AgentToolDef[]
  maxTokens?: number
}

export interface AiStreamChunk {
  requestId: string
  /** 'ping' = wire-level keepalive so the renderer can tell a live stream from a dead one */
  type: 'delta' | 'thinking' | 'tool-call' | 'done' | 'error' | 'ping'
  text?: string
  /** complete parsed tool call (emitted once its arguments finish streaming) */
  toolCall?: AgentToolCall
  error?: string
  /** machine-readable error cause ('timeout', exhausted 'credits'); lets the renderer localize the message */
  errorCode?: 'timeout' | 'credits'
  /** normalized stop reason carried on 'done' ('max_tokens' = output cut off by the token limit) */
  stopReason?: string
}
