import type { AgentMessage, AgentToolCall, AgentToolDef } from '@araoffice/agent-core'
import { streamCustom } from './custom'
import { streamGoogle } from './google'
import type { AiProviderConfig, AiProviderId } from './types'

export interface StreamCallbacks {
  onDelta: (text: string) => void
  onThinking?: (text: string) => void
  onToolCall: (call: AgentToolCall) => void
  onStopReason?: (reason: string) => void
  onActivity?: () => void
  signal: AbortSignal
}

export function streamForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  if (provider === 'google') return streamGoogle(config, system, messages, tools, cb, maxTokens)
  if (provider === 'custom') return streamCustom(config, system, messages, tools, cb)
  return Promise.reject(new Error(`Unknown provider: ${provider}`))
}
