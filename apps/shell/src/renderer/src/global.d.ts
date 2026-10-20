import type { ShellAiApi } from '../../shared/ai-api'

declare global {
  interface Window {
    aiOfficeAi: ShellAiApi
  }
}

export {}
