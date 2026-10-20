import { chatCustom } from './custom'
import { chatGoogle } from './google'
import type { AiChatResponse, AiProviderConfig, AiProviderId } from './types'

export function chatForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  user: string,
  signal?: AbortSignal,
): Promise<AiChatResponse> {
  if (provider === 'google') return chatGoogle(config, system, user, signal)
  if (provider === 'custom') return chatCustom(config, system, user, signal)
  return Promise.resolve({ ok: false, error: `Unknown provider: ${provider}` })
}
