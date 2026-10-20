import type { AiSettings, AiStreamChunk, AiStreamRequest } from '@araoffice/ai-provider'

export interface ShellAiApi {
  getSettings(): Promise<AiSettings>
  setSettings(settings: AiSettings): Promise<void>
  fetchCustomModels?(override?: { baseUrl?: string; apiKey?: string }): Promise<{
    ok: boolean
    models: string[]
    error?: string
  }>
  stream(
    request: Omit<AiStreamRequest, 'requestId'>,
    callbacks: {
      onChunk: (chunk: AiStreamChunk) => void
    },
  ): {
    requestId: string
    cancel: () => void
  }
  cancel(requestId: string): Promise<void>
}
