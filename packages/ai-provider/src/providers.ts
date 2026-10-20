import type {
  AiModelMeta,
  AiProviderConfig,
  AiProviderId,
  AiProviderMeta,
  AiSettings,
} from './types'

/**
 * ARAOffice talks straight to the model endpoint the user configured:
 * Google AI Studio (Gemini) or any OpenAI-compatible endpoint (Ollama,
 * LM Studio, OpenAI, ...). No proxy service, no account, no telemetry.
 */

export const GOOGLE_PROVIDER: AiProviderId = 'google'
export const CUSTOM_PROVIDER: AiProviderId = 'custom'

/** Recommended Gemini model: fast reasoning, drafting, and document analysis. */
export const DEFAULT_GOOGLE_MODEL = 'gemini-3.8-flash'
/** Default model for an OpenAI-compatible endpoint; the UI can fetch the real list. */
export const DEFAULT_CUSTOM_MODEL = 'gpt-4o-mini'

export const GOOGLE_MODELS = [
  'gemini-3.8-flash',
  'gemini-3.5-flash-lite',
  'gemini-3.1-pro',
] as const

/** One-click base URLs for the well-known OpenAI-compatible servers. */
export const CUSTOM_ENDPOINT_PRESETS = [
  { id: 'ollama', label: 'Ollama', baseUrl: 'http://localhost:11434/v1' },
  { id: 'lmstudio', label: 'LM Studio', baseUrl: 'http://localhost:1234/v1' },
  { id: 'openai', label: 'OpenAI', baseUrl: 'https://api.openai.com/v1' },
] as const

const GOOGLE_MODEL_META: AiModelMeta[] = [
  {
    id: DEFAULT_GOOGLE_MODEL,
    label: 'Gemini 3.8 Flash',
    description: 'Fast, recommended for everyday office work.',
    bestFor: 'Drafting, editing, analysis, and multi-step work',
    capabilities: ['Deep analysis', 'Multi-step reasoning', 'Support vision'],
  },
  {
    id: 'gemini-3.5-flash-lite',
    label: 'Gemini 3.5 Flash Lite',
    description: 'Blazing fast with generous free-tier limits.',
    bestFor: 'Quick edits, summaries, and short answers',
    capabilities: ['Low latency', 'High volume', 'Support vision'],
  },
  {
    id: 'gemini-3.1-pro',
    label: 'Gemini 3.1 Pro',
    description: 'Advanced reasoning for complex synthesis.',
    bestFor: 'Deep research, audits, and complicated documents',
    capabilities: ['Advanced reasoning', 'Long context', 'Support vision'],
  },
]

export const AI_PROVIDERS: AiProviderMeta[] = [
  {
    id: GOOGLE_PROVIDER,
    label: 'Gemini AI',
    models: [...GOOGLE_MODELS],
    modelDetails: GOOGLE_MODEL_META,
    defaultModel: DEFAULT_GOOGLE_MODEL,
  },
  {
    id: CUSTOM_PROVIDER,
    label: 'OpenAI-Compatible',
    models: [DEFAULT_CUSTOM_MODEL],
    defaultModel: DEFAULT_CUSTOM_MODEL,
  },
]

/** Trim to a string or ''. Anything falsy/non-string becomes ''. */
function cleanStr(value: unknown): string {
  return typeof value === 'string' ? value.trim() : ''
}

/** OpenAI-compatible roots must be http(s) URLs (the regex only pre-filters; validity comes from the protocol check in custom.ts). */
export function sanitizeCustomBaseUrl(raw: unknown): string {
  const value = cleanStr(raw).replace(/\/+$/, '')
  return /^https?:\/\/\S+$/.test(value) ? value : ''
}

export function defaultAiSettings(): AiSettings {
  return {
    provider: GOOGLE_PROVIDER,
    providers: {
      google: { model: DEFAULT_GOOGLE_MODEL, apiKey: '' },
      custom: { model: DEFAULT_CUSTOM_MODEL, baseUrl: '', apiKey: '' },
    },
  }
}

/**
 * Normalize whatever is on disk into the current schema. Unknown providers
 * (including the retired 'ara' id) fall back to Google; every other field is
 * preserved so a half-configured second provider survives a restart.
 */
export function resolveAiSettings(stored: Partial<AiSettings>, defaults: AiSettings): AiSettings {
  const google = stored.providers?.google
  const custom = stored.providers?.custom
  return {
    provider: stored.provider === CUSTOM_PROVIDER ? CUSTOM_PROVIDER : GOOGLE_PROVIDER,
    providers: {
      google: {
        model: cleanStr(google?.model) || defaults.providers.google.model,
        apiKey: cleanStr(google?.apiKey),
      },
      custom: {
        model: cleanStr(custom?.model) || defaults.providers.custom.model,
        baseUrl: sanitizeCustomBaseUrl(custom?.baseUrl),
        apiKey: cleanStr(custom?.apiKey),
      },
    },
  }
}

/** Renderer-facing settings: an apiKey never leaves the main process. */
export function redactAiSettings(settings: AiSettings): AiSettings {
  const google = settings.providers.google
  const custom = settings.providers.custom
  return {
    provider: settings.provider,
    providers: {
      google: {
        model: google.model,
        apiKey: '',
        hasApiKey: Boolean(google.apiKey && google.apiKey.length > 0),
      },
      custom: {
        model: custom.model,
        baseUrl: custom.baseUrl ?? '',
        apiKey: '',
        hasApiKey: Boolean(custom.apiKey && custom.apiKey.length > 0),
      },
    },
  }
}

/**
 * Merge a renderer's settings into the stored ones.
 *
 * The renderer only ever sees redacted settings, so a key it did not touch
 * arrives as ''. An empty incoming key therefore means "keep the stored one" —
 * per provider, so clearing the Gemini key never wipes the custom one.
 */
export function mergeAiSettings(stored: AiSettings, incoming: Partial<AiSettings>): AiSettings {
  const merged = resolveAiSettings(incoming, defaultAiSettings())
  return {
    provider: merged.provider,
    providers: {
      google: {
        model: merged.providers.google.model,
        apiKey: merged.providers.google.apiKey || stored.providers.google.apiKey || '',
      },
      custom: {
        model: merged.providers.custom.model,
        baseUrl: merged.providers.custom.baseUrl ?? '',
        apiKey: merged.providers.custom.apiKey || stored.providers.custom.apiKey || '',
      },
    },
  }
}

/** The request handlers resolve provider config from one place so all of them share the same fallback rules. */
export function activeProviderConfig(settings: AiSettings): {
  provider: AiProviderId
  config: AiProviderConfig
} {
  const config =
    settings.provider === CUSTOM_PROVIDER ? settings.providers.custom : settings.providers.google
  return { provider: settings.provider, config: { ...config, apiKey: config.apiKey ?? '' } }
}

/**
 * True when the active provider has everything it needs to call the endpoint:
 * Google needs an API key, an OpenAI-compatible endpoint needs a base URL.
 */
export function isProviderConfigured(settings: AiSettings): boolean {
  const { provider, config } = activeProviderConfig(settings)
  if (!config.model) return false
  return provider === GOOGLE_PROVIDER ? Boolean(config.apiKey) : Boolean(config.baseUrl)
}
