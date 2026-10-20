import { createElectronTransport as createSharedTransport } from '@araoffice/agent-core'
import type { AiSettings } from '@araoffice/ai-provider'
import { t } from '../i18n/locale'

/** The shared IPC transport wired to the sheets preload bridge (window.desktopApi). */
export function createElectronTransport(getSettings: () => AiSettings) {
  return createSharedTransport({
    app: 'sheets',
    bridge: window.desktopApi,
    getSettings,
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
  })
}
