import { createElectronTransport as createSharedTransport } from '@araoffice/agent-core'
import type { AiSettings } from '@araoffice/ai-provider'
import { t } from '../i18n/locale'

/** The shared IPC transport wired to the markdown preload bridge (window.markdownApi). */
export function createElectronTransport(getSettings: () => AiSettings) {
  return createSharedTransport({
    app: 'markdown',
    bridge: window.markdownApi,
    getSettings,
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
  })
}
