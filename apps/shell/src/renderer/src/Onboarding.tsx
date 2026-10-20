import { useEffect, useRef, useState } from 'react'
import type { UiTheme } from '../../shared/home-api'
import type { AiProviderId, AiSettings } from '@araoffice/ai-provider'
import {
  AI_PROVIDERS,
  CUSTOM_ENDPOINT_PRESETS,
  DEFAULT_GOOGLE_MODEL,
} from '@araoffice/ai-provider/providers'
import { LANGUAGE_OPTIONS, useI18n } from './locale'
import type { StringKey } from './locale'
import { FileBadge } from './Home'
import './onboarding.css'

/** Gemini model list comes from the shared provider metadata so the UI never drifts. */
const GEMINI_MODELS = AI_PROVIDERS.find((provider) => provider.id === 'google')?.models ?? []

interface OnboardingProps {
  /** called when the user finishes the tour */
  onDone: () => void
  /** called when the user completes the tour and starts using the app */
  onStart: () => void
  isDarkTheme?: boolean
  onToggleTheme?: () => void
  initialTheme?: UiTheme
}

interface Slide {
  id: 'welcome' | 'apps' | 'free' | 'ai'
  titleKey: StringKey
  subtitleKey?: StringKey
  bodyKey?: StringKey
}

const SLIDES: readonly Slide[] = [
  {
    id: 'welcome',
    titleKey: 'onbTitle1',
    subtitleKey: 'onbSubtitle1',
  },
  {
    id: 'apps',
    titleKey: 'onbTitle2',
    subtitleKey: 'onbSubtitle2',
  },
  {
    id: 'free',
    titleKey: 'onbTitle4',
    subtitleKey: 'onbSubtitle4',
  },
  {
    id: 'ai',
    titleKey: 'onbTitle3',
    subtitleKey: 'onbSubtitle3',
  },
] as const

const APPS = [
  { ext: 'docx', key: 'newDoc' as const, tagKey: 'appTagDocx' as const },
  { ext: 'xlsx', key: 'newSheet' as const, tagKey: 'appTagXlsx' as const },
  { ext: 'pptx', key: 'newSlide' as const, tagKey: 'appTagPptx' as const },
  { ext: 'pdf', key: 'newPdf' as const, tagKey: 'appTagPdf' as const },
  { ext: 'md', key: 'newMarkdown' as const, tagKey: 'appTagMd' as const },
  { ext: 'drawio', key: 'newDiagram' as const, tagKey: 'appTagDiag' as const },
] as const

export function Onboarding({
  onDone,
  onStart,
  isDarkTheme: propIsDarkTheme,
  onToggleTheme: propOnToggleTheme,
  initialTheme = 'system',
}: OnboardingProps) {
  const { t, lang, setLang } = useI18n()
  const [index, setIndex] = useState(0)
  const [localTheme, setLocalTheme] = useState<UiTheme>(initialTheme)
  const [systemDark, setSystemDark] = useState(
    typeof window === 'undefined'
      ? false
      : window.matchMedia('(prefers-color-scheme: dark)').matches,
  )
  // AI provider flow on the last slide: pick Gemini (Google AI Studio key) or
  // any OpenAI-compatible endpoint. Writes the same ai-settings.json as the
  // Home switch. AI is entirely optional — the tour finishes without it.
  const [aiChoice, setAiChoice] = useState<AiProviderId>('google')
  const [showApiKey, setShowApiKey] = useState(false)
  const [providerSaved, setProviderSaved] = useState(false)
  const [aiNotice, setAiNotice] = useState<'none' | 'incomplete'>('none')
  const [googleModel, setGoogleModel] = useState(DEFAULT_GOOGLE_MODEL)
  const [googleKey, setGoogleKey] = useState('')
  const [aiEndpoint, setAiEndpoint] = useState('')
  const [aiModelId, setAiModelId] = useState('')
  const [aiKey, setAiKey] = useState('')
  const [availableModels, setAvailableModels] = useState<string[]>([])
  const [isFetchingModels, setIsFetchingModels] = useState(false)
  const [aiBusy, setAiBusy] = useState(false)
  /** what the main process already holds, so the tour never rewrites it unchanged */
  const [storedGoogleModel, setStoredGoogleModel] = useState(DEFAULT_GOOGLE_MODEL)
  const [storedGoogleKey, setStoredGoogleKey] = useState(false)
  const [storedCustom, setStoredCustom] = useState({ baseUrl: '', model: '' })
  const cardRef = useRef<HTMLDivElement>(null)
  const slide = SLIDES[index]
  const isLast = index === SLIDES.length - 1

  useEffect(() => {
    void window.aiOfficeAi?.getSettings?.().then((settings) => {
      if (!settings) return
      const storedModel = settings.providers.google?.model || DEFAULT_GOOGLE_MODEL
      setGoogleModel(storedModel)
      setStoredGoogleModel(storedModel)
      setStoredGoogleKey(Boolean(settings.providers.google?.hasApiKey))
      setStoredCustom({
        baseUrl: settings.providers.custom?.baseUrl ?? '',
        model: settings.providers.custom?.model ?? '',
      })
      if (settings.provider === 'custom') {
        setAiChoice('custom')
        setAiEndpoint(settings.providers.custom?.baseUrl ?? '')
        setAiModelId(settings.providers.custom?.model ?? '')
        setProviderSaved(true)
      }
    })
  }, [])

  useEffect(() => {
    const media = window.matchMedia('(prefers-color-scheme: dark)')
    const syncSystemTheme = () => setSystemDark(media.matches)
    syncSystemTheme()
    media.addEventListener('change', syncSystemTheme)
    return () => media.removeEventListener('change', syncSystemTheme)
  }, [])

  const effectiveIsDark =
    typeof propIsDarkTheme === 'boolean'
      ? propIsDarkTheme
      : localTheme === 'dark' || (localTheme === 'system' && systemDark)

  const handleToggleTheme = () => {
    if (propOnToggleTheme) {
      propOnToggleTheme()
    } else {
      const next: UiTheme = effectiveIsDark ? 'light' : 'dark'
      setLocalTheme(next)
      document.documentElement.setAttribute('data-theme', next)
      void window.aiOffice.setTheme(next).catch(() => {})
    }
  }

  const openApiPanel = async () => {
    // Prefill from the redacted read (apiKey stays '' — blank keeps the key).
    try {
      const settings = await window.aiOfficeAi?.getSettings()
      if (settings) {
        if (!aiEndpoint) setAiEndpoint(settings.providers.custom?.baseUrl ?? '')
        if (!aiModelId) setAiModelId(settings.providers.custom?.model ?? '')
        setAiKey('')
      }
    } catch {
      /* bridge unavailable: fields stay as typed */
    }
  }

  const fetchCustomModels = async () => {
    const endpoint = aiEndpoint.trim().replace(/\/+$/, '')
    if (!endpoint) {
      setAiNotice('incomplete')
      return
    }
    setIsFetchingModels(true)
    setAiNotice('none')
    try {
      const res = await window.aiOfficeAi?.fetchCustomModels?.({
        baseUrl: endpoint,
        apiKey: aiKey.trim() || undefined,
      })
      if (res?.ok && Array.isArray(res.models) && res.models.length > 0) {
        setAvailableModels(res.models)
        if (!aiModelId.trim()) setAiModelId(res.models[0])
      } else {
        setAiNotice('incomplete')
      }
    } catch {
      setAiNotice('incomplete')
    } finally {
      setIsFetchingModels(false)
    }
  }

  const saveProvider = async (): Promise<boolean> => {
    const endpoint = aiEndpoint.trim().replace(/\/+$/, '')
    if (aiChoice === 'google' && !googleModel.trim()) {
      setAiNotice('incomplete')
      return false
    }
    if (aiChoice === 'custom' && (!/^https?:\/\/\S+$/.test(endpoint) || !aiModelId.trim())) {
      setAiNotice('incomplete')
      return false
    }
    setAiBusy(true)
    try {
      const settings = await window.aiOfficeAi?.getSettings()
      if (!settings) throw new Error('AI bridge unavailable')
      const next: AiSettings = {
        provider: aiChoice,
        providers: {
          google:
            aiChoice === 'google'
              ? { model: googleModel.trim(), apiKey: googleKey }
              : settings.providers.google,
          custom:
            aiChoice === 'custom'
              ? { model: aiModelId.trim(), baseUrl: endpoint, apiKey: aiKey }
              : settings.providers.custom,
        },
      }
      await window.aiOfficeAi?.setSettings(next)
      setProviderSaved(true)
      setStoredGoogleModel(googleModel.trim())
      setStoredGoogleKey(storedGoogleKey || googleKey.trim().length > 0)
      setStoredCustom({ baseUrl: endpoint, model: aiModelId.trim() })
      setAiNotice('none')
      return true
    } catch {
      setAiNotice('incomplete')
      return false
    } finally {
      setAiBusy(false)
    }
  }

  const next = async () => {
    if (!isLast) {
      setIndex(index + 1)
      return
    }
    // The app is usable without AI: persist a provider only when the user
    // actually typed something that differs from the stored configuration.
    const customEndpoint = aiEndpoint.trim().replace(/\/+$/, '')
    const googleChanged =
      aiChoice === 'google' &&
      (googleKey.trim().length > 0 || googleModel.trim() !== storedGoogleModel || !storedGoogleKey)
    const customChanged =
      aiChoice === 'custom' &&
      (aiKey.trim().length > 0 ||
        customEndpoint !== storedCustom.baseUrl ||
        aiModelId.trim() !== storedCustom.model)
    if (!providerSaved && (googleChanged || customChanged)) {
      await saveProvider()
    }
    onStart()
  }

  // move focus into the dialog on mount so keyboard users start inside it
  useEffect(() => {
    cardRef.current?.focus({ preventScroll: true })
  }, [])

  // slide changes can strip focus from the active control — pull focus back onto the
  // card so it never drops to body
  useEffect(() => {
    const card = cardRef.current
    const active = document.activeElement
    if (card && (!(active instanceof HTMLElement) || !card.contains(active))) {
      card.focus({ preventScroll: true })
    }
  }, [index])

  // keyboard handling: Escape skips, Enter / ArrowRight advance, ArrowLeft goes
  // back, and Tab is trapped inside the dialog (aria-modal). Enter is ignored
  // when a button is focused so the native click doesn't double-fire.
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onDone()
        return
      }
      if (event.key === 'Tab') {
        const card = cardRef.current
        if (!card) return
        const focusables = Array.from(
          card.querySelectorAll<HTMLElement>('button, select, input'),
        ).filter((el) => !el.closest('[inert]'))
        if (focusables.length === 0) return
        const first = focusables[0]
        const last = focusables[focusables.length - 1]
        const active = document.activeElement
        const onButton = active instanceof HTMLElement && focusables.includes(active)
        if (!onButton) {
          event.preventDefault()
          ;(event.shiftKey ? last : first).focus()
        } else if (event.shiftKey && active === first) {
          event.preventDefault()
          last.focus()
        } else if (!event.shiftKey && active === last) {
          event.preventDefault()
          first.focus()
        }
        return
      }
      const target = event.target instanceof HTMLElement ? event.target : null
      const interactiveFocused = target?.closest('button, select, input') !== null
      if (
        (event.key === 'Enter' && !interactiveFocused) ||
        (event.key === 'ArrowRight' && !interactiveFocused)
      ) {
        next()
      }
      if (event.key === 'ArrowLeft' && !interactiveFocused) setIndex((i) => Math.max(0, i - 1))
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  })

  const footerNoticeKey = aiNotice === 'incomplete' ? 'onbAiIncomplete' : ''

  return (
    <div className="onb-overlay" role="dialog" aria-modal="true" aria-label={t(slide.titleKey)}>
      <div className="onb-card" ref={cardRef} tabIndex={-1}>
        <div className="onb-header-bar">
          <div className="onb-header-actions">
            <label className="onb-language" title={t('language')}>
              <span className="onb-visually-hidden">{t('language')}</span>
              <select
                value={lang}
                aria-label={t('language')}
                onChange={(event) => setLang(event.target.value as typeof lang)}
              >
                {LANGUAGE_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
              <svg
                className="onb-select-arrow"
                width="12"
                height="12"
                viewBox="0 0 24 24"
                fill="none"
                aria-hidden="true"
              >
                <path
                  d="M6 9l6 6 6-6"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            </label>
            <button
              className="onb-theme-toggle"
              type="button"
              onClick={handleToggleTheme}
              aria-label={effectiveIsDark ? t('themeLight') : t('themeDark')}
              title={effectiveIsDark ? t('themeLight') : t('themeDark')}
              aria-pressed={effectiveIsDark}
            >
              {effectiveIsDark ? (
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                  <circle cx="12" cy="12" r="4" stroke="currentColor" strokeWidth="1.8" />
                  <path
                    d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41"
                    stroke="currentColor"
                    strokeWidth="1.8"
                    strokeLinecap="round"
                  />
                </svg>
              ) : (
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                  <path
                    d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79Z"
                    stroke="currentColor"
                    strokeWidth="1.8"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                </svg>
              )}
            </button>
          </div>
        </div>

        <div className="onb-stage">
          <div
            className={`onb-morph-mascot onb-pos-s${index + 1}`}
            id="onbMorphMascot"
            aria-hidden="true"
          >
            <div className="onb-morph-inner">
              <img className="onb-mascot-image" src="./ARA_Assistant.png" alt="" />
            </div>
          </div>

          {SLIDES.map((s, i) => (
            <div
              key={s.id}
              className={`onb-slide${i === index ? ' active' : ''}${s.id === 'welcome' ? ' onb-slide-welcome' : ''}`}
              inert={i !== index}
            >
              {s.id === 'welcome' ? (
                <div className="s1-container">
                  <div className="welcome-pill">
                    <span aria-hidden="true">✨</span>
                    <span>{t('onbWelcomePill')}</span>
                  </div>

                  <h1 className="s1-title">{t(s.titleKey)}</h1>
                  {s.subtitleKey && <p className="s1-subtitle">{t(s.subtitleKey)}</p>}

                  <div className="s1-pills">
                    <div className="s1-pill">
                      <span aria-hidden="true">🔒</span>
                      <span>{t('onbPill1')}</span>
                    </div>
                    <div className="s1-pill">
                      <span aria-hidden="true">⚡</span>
                      <span>{t('onbPill2')}</span>
                    </div>
                    <div className="s1-pill">
                      <span aria-hidden="true">🤖</span>
                      <span>{t('onbPill3')}</span>
                    </div>
                  </div>
                </div>
              ) : (
                <>
                  <span className="onb-art onb-art-placeholder" aria-hidden="true" />
                  <h2 className="onb-title">{t(s.titleKey)}</h2>
                  {s.subtitleKey && <p className="onb-subtitle">{t(s.subtitleKey)}</p>}
                  {s.bodyKey && <p className="onb-body">{t(s.bodyKey)}</p>}

                  {s.id === 'apps' && (
                    <div className="onb-feature-grid">
                      {APPS.map((app) => (
                        <div className="onb-feature" key={app.ext}>
                          <span className="onb-feature-icon" aria-hidden="true">
                            <FileBadge ext={app.ext} size={36} />
                          </span>
                          <div className="onb-feature-meta">
                            <span className="onb-feature-title">{t(app.key)}</span>
                            <span className="onb-feature-tag">{t(app.tagKey)}</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}

                  {s.id === 'free' && (
                    <div className="onb-highlights-grid">
                      <div className="onb-highlight-card">
                        <div
                          className="onb-highlight-icon"
                          style={{ background: 'rgba(16, 185, 129, 0.12)', color: '#10b981' }}
                          aria-hidden="true"
                        >
                          <svg
                            width="18"
                            height="18"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <path d="M20 12V22H4V12" />
                            <rect x="2" y="7" width="20" height="5" rx="1" />
                            <path d="M12 22V7" />
                            <path d="M12 7H7.5a2.5 2.5 0 0 1 0-5C11 2 12 7 12 7z" />
                            <path d="M12 7h4.5a2.5 2.5 0 0 0 0-5C13 2 12 7 12 7z" />
                          </svg>
                        </div>
                        <div className="onb-highlight-text">
                          <strong>{t('onbPrivacyCard1Title')}</strong>
                          <span>{t('onbPrivacyCard1Sub')}</span>
                        </div>
                      </div>
                      <div className="onb-highlight-card">
                        <div
                          className="onb-highlight-icon onb-highlight-icon-shield"
                          aria-hidden="true"
                        >
                          <svg
                            width="18"
                            height="18"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
                          </svg>
                        </div>
                        <div className="onb-highlight-text">
                          <strong>{t('onbPrivacyCard2Title')}</strong>
                          <span>{t('onbPrivacyCard2Sub')}</span>
                        </div>
                      </div>
                      <div className="onb-highlight-card">
                        <div
                          className="onb-highlight-icon onb-highlight-icon-bolt"
                          aria-hidden="true"
                        >
                          <svg
                            width="18"
                            height="18"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2" />
                          </svg>
                        </div>
                        <div className="onb-highlight-text">
                          <strong>{t('onbPrivacyCard3Title')}</strong>
                          <span>{t('onbPrivacyCard3Sub')}</span>
                        </div>
                      </div>
                    </div>
                  )}

                  {s.id === 'ai' && (
                    <div className="onb-ai-slide-content">
                      <div
                        className="onb-provider-grid"
                        role="radiogroup"
                        aria-label={t('aiMenuTitle')}
                      >
                        <button
                          type="button"
                          role="radio"
                          aria-checked={aiChoice === 'google'}
                          className={`onb-provider-option${aiChoice === 'google' ? ' selected' : ''}`}
                          onClick={() => {
                            setAiChoice('google')
                            setAiNotice('none')
                          }}
                        >
                          <div className="onb-provider-option-top">
                            <div className="onb-provider-icon onb-icon-ara" aria-hidden="true">
                              ✦
                            </div>
                            <div className="onb-provider-meta">
                              <span className="onb-provider-name">{t('onbAiPickGemini')}</span>
                              <span className="onb-provider-sub">{t('onbAiPickGeminiSub')}</span>
                            </div>
                            <div
                              className={`onb-radio-dot${aiChoice === 'google' ? ' checked' : ''}`}
                              aria-hidden="true"
                            />
                          </div>
                        </button>

                        <button
                          type="button"
                          role="radio"
                          aria-checked={aiChoice === 'custom'}
                          className={`onb-provider-option${aiChoice === 'custom' ? ' selected' : ''}`}
                          onClick={() => {
                            setAiChoice('custom')
                            setAiNotice('none')
                            void openApiPanel()
                          }}
                        >
                          <div className="onb-provider-option-top">
                            <div className="onb-provider-icon onb-icon-key" aria-hidden="true">
                              🔑
                            </div>
                            <div className="onb-provider-meta">
                              <span className="onb-provider-name">{t('onbAiPickLocal')}</span>
                              <span className="onb-provider-sub">{t('onbAiPickLocalSub')}</span>
                            </div>
                            <div
                              className={`onb-radio-dot${aiChoice === 'custom' ? ' checked' : ''}`}
                              aria-hidden="true"
                            />
                          </div>
                        </button>
                      </div>

                      {aiChoice === 'google' ? (
                        <div className="onb-api-panel open">
                          <div className="onb-api-fields-grid">
                            <div className="onb-api-row">
                              <label htmlFor="onb-ai-google-model">
                                <span className="onb-field-label-text">
                                  <svg
                                    width="12"
                                    height="12"
                                    viewBox="0 0 24 24"
                                    fill="none"
                                    stroke="currentColor"
                                    strokeWidth="2"
                                    strokeLinecap="round"
                                    strokeLinejoin="round"
                                    aria-hidden="true"
                                  >
                                    <rect x="4" y="4" width="16" height="16" rx="2" />
                                    <rect x="9" y="9" width="6" height="6" />
                                  </svg>
                                  {t('aiMenuModelLabel')}
                                </span>
                              </label>
                              <select
                                id="onb-ai-google-model"
                                value={googleModel}
                                onChange={(e) => setGoogleModel(e.target.value)}
                              >
                                {GEMINI_MODELS.map((model) => (
                                  <option key={model} value={model}>
                                    {model}
                                  </option>
                                ))}
                              </select>
                            </div>
                            <div className="onb-api-row">
                              <label htmlFor="onb-ai-google-key">
                                <span className="onb-field-label-text">
                                  <svg
                                    width="12"
                                    height="12"
                                    viewBox="0 0 24 24"
                                    fill="none"
                                    stroke="currentColor"
                                    strokeWidth="2"
                                    strokeLinecap="round"
                                    strokeLinejoin="round"
                                    aria-hidden="true"
                                  >
                                    <path d="M21 2l-2 2m-1.5 1.5L16 7l-1.5-1.5M16 7l-4 4-2-2-4 4 2 2-4 4 2 2 4-4z" />
                                  </svg>
                                  {t('onbAiKey')}
                                </span>
                              </label>
                              <div className="onb-password-wrapper">
                                <input
                                  id="onb-ai-google-key"
                                  type={showApiKey ? 'text' : 'password'}
                                  autoComplete="off"
                                  value={googleKey}
                                  placeholder="AIza..."
                                  onChange={(e) => setGoogleKey(e.target.value)}
                                />
                                <button
                                  type="button"
                                  className="onb-pwd-toggle"
                                  onClick={() => setShowApiKey(!showApiKey)}
                                  tabIndex={-1}
                                  aria-label={showApiKey ? 'Hide key' : 'Show key'}
                                >
                                  {showApiKey ? (
                                    <svg
                                      width="14"
                                      height="14"
                                      viewBox="0 0 24 24"
                                      fill="none"
                                      stroke="currentColor"
                                      strokeWidth="2"
                                      strokeLinecap="round"
                                      strokeLinejoin="round"
                                    >
                                      <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24" />
                                      <line x1="1" y1="1" x2="23" y2="23" />
                                    </svg>
                                  ) : (
                                    <svg
                                      width="14"
                                      height="14"
                                      viewBox="0 0 24 24"
                                      fill="none"
                                      stroke="currentColor"
                                      strokeWidth="2"
                                      strokeLinecap="round"
                                      strokeLinejoin="round"
                                    >
                                      <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8z" />
                                      <circle cx="12" cy="12" r="3" />
                                    </svg>
                                  )}
                                </button>
                              </div>
                            </div>
                          </div>
                          <div className="onb-api-actions">
                            <p className="onb-api-note">{t('aiMenuStudioNote')}</p>
                            <div className="onb-api-action-buttons">
                              <button
                                type="button"
                                className="onb-fetch-models-link"
                                onClick={() =>
                                  void window.aiOfficeWindow.openExternal(
                                    'https://aistudio.google.com/apikey',
                                  )
                                }
                                rel="noreferrer"
                              >
                                {t('aiMenuStudioLink')}
                              </button>
                              <button
                                type="button"
                                className="onb-api-save"
                                onClick={() => void saveProvider()}
                                disabled={aiBusy}
                              >
                                {providerSaved ? `✓ ${t('onbAiSavedBadge')}` : t('onbAiKeySave')}
                              </button>
                            </div>
                          </div>
                        </div>
                      ) : (
                        <div className="onb-api-panel open">
                          <div className="onb-api-presets">
                            <span className="onb-field-label-text">{t('aiMenuPresets')}</span>
                            <div className="onb-api-preset-chips">
                              {CUSTOM_ENDPOINT_PRESETS.map((preset) => (
                                <button
                                  key={preset.id}
                                  type="button"
                                  className="onb-api-preset-chip"
                                  onClick={() => setAiEndpoint(preset.baseUrl)}
                                >
                                  {preset.label}
                                </button>
                              ))}
                            </div>
                          </div>
                          <div className="onb-api-fields-grid">
                            <div className="onb-api-row">
                              <label htmlFor="onb-ai-endpoint">
                                <span className="onb-field-label-text">
                                  <svg
                                    width="12"
                                    height="12"
                                    viewBox="0 0 24 24"
                                    fill="none"
                                    stroke="currentColor"
                                    strokeWidth="2"
                                    strokeLinecap="round"
                                    strokeLinejoin="round"
                                    aria-hidden="true"
                                  >
                                    <circle cx="12" cy="12" r="10" />
                                    <line x1="2" y1="12" x2="22" y2="12" />
                                    <path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z" />
                                  </svg>
                                  {t('onbAiEndpoint')}
                                </span>
                              </label>
                              <input
                                id="onb-ai-endpoint"
                                value={aiEndpoint}
                                placeholder="https://api.openai.com/v1"
                                onChange={(e) => setAiEndpoint(e.target.value)}
                              />
                            </div>
                            <div className="onb-api-row">
                              <label htmlFor="onb-ai-model">
                                <span className="onb-field-label-text">
                                  <svg
                                    width="12"
                                    height="12"
                                    viewBox="0 0 24 24"
                                    fill="none"
                                    stroke="currentColor"
                                    strokeWidth="2"
                                    strokeLinecap="round"
                                    strokeLinejoin="round"
                                    aria-hidden="true"
                                  >
                                    <rect x="4" y="4" width="16" height="16" rx="2" />
                                    <rect x="9" y="9" width="6" height="6" />
                                  </svg>
                                  {t('onbAiModel')}
                                </span>
                              </label>
                              <input
                                id="onb-ai-model"
                                list="onb-custom-models-list"
                                value={aiModelId}
                                placeholder="gpt-4o-mini"
                                onChange={(e) => setAiModelId(e.target.value)}
                              />
                              <datalist id="onb-custom-models-list">
                                {availableModels.map((model) => (
                                  <option key={model} value={model} />
                                ))}
                              </datalist>
                              <button
                                type="button"
                                className="onb-fetch-models-link"
                                onClick={() => void fetchCustomModels()}
                                disabled={isFetchingModels || !aiEndpoint.trim()}
                              >
                                {isFetchingModels ? '...' : t('aiMenuFetchModels')}
                              </button>
                            </div>
                          </div>
                          <div className="onb-api-row">
                            <label htmlFor="onb-ai-key">
                              <span className="onb-field-label-text">
                                <svg
                                  width="12"
                                  height="12"
                                  viewBox="0 0 24 24"
                                  fill="none"
                                  stroke="currentColor"
                                  strokeWidth="2"
                                  strokeLinecap="round"
                                  strokeLinejoin="round"
                                  aria-hidden="true"
                                >
                                  <path d="M21 2l-2 2m-1.5 1.5L16 7l-1.5-1.5M16 7l-4 4-2-2-4 4 2 2-4 4 2 2 4-4 2 2 4-4z" />
                                </svg>
                                {t('onbAiKey')}
                              </span>
                            </label>
                            <div className="onb-password-wrapper">
                              <input
                                id="onb-ai-key"
                                type={showApiKey ? 'text' : 'password'}
                                autoComplete="off"
                                value={aiKey}
                                placeholder="sk-..."
                                onChange={(e) => setAiKey(e.target.value)}
                              />
                              <button
                                type="button"
                                className="onb-pwd-toggle"
                                onClick={() => setShowApiKey(!showApiKey)}
                                tabIndex={-1}
                                aria-label={showApiKey ? 'Hide key' : 'Show key'}
                              >
                                {showApiKey ? (
                                  <svg
                                    width="14"
                                    height="14"
                                    viewBox="0 0 24 24"
                                    fill="none"
                                    stroke="currentColor"
                                    strokeWidth="2"
                                    strokeLinecap="round"
                                    strokeLinejoin="round"
                                  >
                                    <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24" />
                                    <line x1="1" y1="1" x2="23" y2="23" />
                                  </svg>
                                ) : (
                                  <svg
                                    width="14"
                                    height="14"
                                    viewBox="0 0 24 24"
                                    fill="none"
                                    stroke="currentColor"
                                    strokeWidth="2"
                                    strokeLinecap="round"
                                    strokeLinejoin="round"
                                  >
                                    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                                    <circle cx="12" cy="12" r="3" />
                                  </svg>
                                )}
                              </button>
                            </div>
                          </div>
                          <div className="onb-api-actions">
                            <button
                              type="button"
                              className="onb-api-save"
                              onClick={() => void saveProvider()}
                              disabled={aiBusy}
                            >
                              {providerSaved ? `✓ ${t('onbAiSavedBadge')}` : t('onbAiKeySave')}
                            </button>
                            <p className="onb-api-note">{t('onbAiKeyNote')}</p>
                          </div>
                        </div>
                      )}
                    </div>
                  )}
                </>
              )}
            </div>
          ))}
        </div>

        <div className="onb-footer">
          <div className="onb-footer-notice-wrap">
            {isLast && footerNoticeKey && (
              <span className="onb-footer-notice" role="status">
                <svg
                  width="13"
                  height="13"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  aria-hidden="true"
                >
                  <circle cx="12" cy="12" r="10" />
                  <line x1="12" y1="8" x2="12" y2="12" />
                  <line x1="12" y1="16" x2="12.01" y2="16" />
                </svg>
                {t(footerNoticeKey)}
              </span>
            )}
          </div>
          <div className="onb-dots">
            {SLIDES.map((s, i) => (
              <button
                key={s.id}
                className={`onb-dot${i === index ? ' active' : ''}`}
                aria-label={t('onbStepAria', { n: i + 1, total: SLIDES.length })}
                aria-current={i === index}
                onClick={() => setIndex(i)}
              />
            ))}
          </div>
          <div className="onb-nav">
            {index > 0 && (
              <button className="onb-back" onClick={() => setIndex(index - 1)}>
                {t('onbBack')}
              </button>
            )}
            <button className="onb-next" onClick={() => void next()}>
              <span>{isLast ? t('onbStart') : t('onbNext')}</span>
              <svg
                width="14"
                height="14"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.2"
                aria-hidden="true"
              >
                <path d="M5 12h14M12 5l7 7-7 7" />
              </svg>
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
