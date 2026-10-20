import { createElectronTransport as createSharedTransport } from '@araoffice/agent-core'
import type { AiSettings } from '@araoffice/ai-provider'

export function createElectronTransport(getSettings: () => AiSettings) {
  return createSharedTransport({
    app: 'diagram',
    bridge: window.diagramApi,
    getSettings,
    unknownErrorText: 'An unknown error occurred in ARA Assistant',
    timeoutErrorText: 'AI request timed out',
    creditsErrorText: 'AI credits exhausted',
  })
}
