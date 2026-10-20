import { describe, expect, it } from 'vitest'
import {
  AI_PROVIDERS,
  DEFAULT_CUSTOM_MODEL,
  DEFAULT_GOOGLE_MODEL,
  defaultAiSettings,
  isProviderConfigured,
  redactAiSettings,
  resolveAiSettings,
} from '../src/providers'

describe('AI provider settings', () => {
  it('exposes Gemini and an OpenAI-compatible provider', () => {
    const settings = defaultAiSettings()
    expect(settings.provider).toBe('google')
    expect(AI_PROVIDERS.map((provider) => provider.id)).toEqual(['google', 'custom'])
    expect(settings.providers.google.model).toBe(DEFAULT_GOOGLE_MODEL)
    expect(settings.providers.custom.model).toBe(DEFAULT_CUSTOM_MODEL)
  })

  it('lists the supported Gemini models with flash recommended first', () => {
    const google = AI_PROVIDERS.find((entry) => entry.id === 'google')
    expect(google?.defaultModel).toBe('gemini-3.8-flash')
    expect(google?.models).toEqual(['gemini-3.8-flash', 'gemini-3.5-flash-lite', 'gemini-3.1-pro'])
    expect(google?.modelDetails).toHaveLength(3)
  })

  it('normalizes unknown and removed provider settings', () => {
    const resolved = resolveAiSettings(
      {
        provider: 'removed',
        providers: {
          google: { model: 'unknown', apiKey: '' },
          custom: { model: 'm', baseUrl: '', apiKey: '' },
        },
      } as never,
      defaultAiSettings(),
    )
    expect(resolved.provider).toBe('google')
    // Any model string is preserved: Gemini ships new models continuously and
    // the user may pick one the picker does not list yet.
    expect(resolved.providers.google.model).toBe('unknown')
  })

  it('falls back to the default model when none is stored', () => {
    const resolved = resolveAiSettings({ provider: 'google' }, defaultAiSettings())
    expect(resolved.providers.google.model).toBe(DEFAULT_GOOGLE_MODEL)
    expect(resolved.providers.custom.model).toBe(DEFAULT_CUSTOM_MODEL)
  })

  it('keeps a stored Gemini key and model across a resolve', () => {
    const resolved = resolveAiSettings(
      {
        provider: 'google',
        providers: {
          google: { model: 'gemini-3.1-pro', apiKey: 'key-123' },
          custom: { model: 'm', baseUrl: '', apiKey: '' },
        },
      },
      defaultAiSettings(),
    )
    expect(resolved.providers.google).toEqual({ model: 'gemini-3.1-pro', apiKey: 'key-123' })
  })

  it('treats a provider as configured only when its credentials are present', () => {
    const defaults = defaultAiSettings()
    expect(isProviderConfigured(defaults)).toBe(false)
    expect(
      isProviderConfigured({
        ...defaults,
        providers: { ...defaults.providers, google: { model: 'm', apiKey: 'k' } },
      }),
    ).toBe(true)
    expect(
      isProviderConfigured({
        ...defaults,
        provider: 'custom',
        providers: {
          ...defaults.providers,
          custom: { model: 'm', baseUrl: 'http://localhost:11434/v1', apiKey: '' },
        },
      }),
    ).toBe(true)
    expect(
      isProviderConfigured({
        ...defaults,
        provider: 'custom',
        providers: { ...defaults.providers, custom: { model: 'm', baseUrl: '', apiKey: '' } },
      }),
    ).toBe(false)
  })

  it('redacts both provider keys for renderer reads', () => {
    const redacted = redactAiSettings({
      provider: 'google',
      providers: {
        google: { model: 'gemini-3.8-flash', apiKey: 'secret' },
        custom: { model: 'm', baseUrl: 'http://x/v1', apiKey: 'secret2' },
      },
    })
    expect(redacted.providers.google.apiKey).toBe('')
    expect(redacted.providers.google.hasApiKey).toBe(true)
    expect(redacted.providers.custom.apiKey).toBe('')
    expect(redacted.providers.custom.hasApiKey).toBe(true)
  })
})
