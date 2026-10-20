import { createElectronTransport as createSharedTransport } from '@araoffice/agent-core'
import type { AiSettings } from '../../shared/ipc'
import { t } from '../i18n/locale'

/** The shared IPC transport wired to the slides preload bridge (window.slidesApi). */
export function createElectronTransport(getSettings: () => AiSettings) {
  return createSharedTransport({
    app: 'slides',
    bridge: window.slidesApi,
    getSettings,
    unknownErrorText: () => t('aiErrUnknown'),
    timeoutErrorText: () => t('aiErrStreamTimeout'),
    creditsErrorText: () => t('aiCreditsExhausted'),
  })
}
