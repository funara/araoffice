import { useEffect, useState } from 'react'

/**
 * Provider settings as the renderer sees them: the apiKey never leaves the
 * main process, so `hasApiKey` is the only signal that a key is stored.
 */
export interface AiReadySettings {
  provider?: string
  providers?: {
    google?: { model?: string; apiKey?: string; hasApiKey?: boolean }
    custom?: { model?: string; baseUrl?: string; apiKey?: string; hasApiKey?: boolean }
  }
}

export interface UseAiReadyOptions {
  readonly getAiSettings?: (() => Promise<AiReadySettings>) | undefined
  readonly onAiSettingsChanged?:
    ((handler: (settings: AiReadySettings) => void) => () => void) | undefined
  readonly initialSettings?: AiReadySettings | null | undefined
}

/**
 * Whether the AI panel can talk to a model.
 *
 * ARAOffice has no account system: the only question is whether the active
 * provider is configured — Gemini needs an API key, an OpenAI-compatible
 * endpoint needs a base URL.
 */
export function useAiReady(options: UseAiReadyOptions): { isReady: boolean } {
  const [settings, setSettings] = useState<AiReadySettings | null>(options.initialSettings ?? null)

  useEffect(() => {
    if (options.initialSettings) setSettings(options.initialSettings)
  }, [options.initialSettings])

  useEffect(() => {
    const getSettings = options.getAiSettings
    if (!getSettings) return
    let active = true
    const applySettings = (next?: AiReadySettings | null): void => {
      if (active && next) setSettings(next)
    }
    const poll = (): void => {
      try {
        const res = getSettings()
        if (res && typeof (res as Promise<unknown>).then === 'function') {
          void (res as Promise<AiReadySettings>).then(applySettings, () => {})
        }
      } catch {
        /* bridge unavailable: keep the last known state */
      }
    }
    poll()
    const unsub = options.onAiSettingsChanged?.((next) => applySettings(next))

    // The settings live in the main process; re-read on focus so a change
    // made in the shell's AI settings shows up without a reload.
    const onFocus = (): void => poll()
    window.addEventListener('focus', onFocus)

    return () => {
      active = false
      unsub?.()
      window.removeEventListener('focus', onFocus)
    }
  }, [options.getAiSettings, options.onAiSettingsChanged])

  const provider = settings?.provider ?? 'google'
  const providers = settings?.providers
  const isReady =
    provider === 'custom'
      ? Boolean(providers?.custom?.baseUrl?.trim())
      : Boolean(providers?.google?.hasApiKey || providers?.google?.apiKey?.trim())

  return { isReady }
}
