import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@araoffice/i18n'
import { ErrorBoundary } from '@araoffice/ui'
import { App } from './App'
import { AudienceView } from './components/AudienceView'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import './styles.css'
import '@araoffice/ui/ai-panel.css'

// ?mode=audience: the presenter view's external-screen audience show window (created by the main process)
const mode = new URLSearchParams(window.location.search).get('mode')

// macOS windows are created with vibrancy; let the thumbnail pane show it
// (the audience show window stays fully opaque)
if (mode !== 'audience' && navigator.platform.toLowerCase().includes('mac'))
  document.body.classList.add('vib')

async function bootstrap(): Promise<void> {
  let lang: Lang = 'zh'
  try {
    lang = await window.slidesApi.getLanguage()
  } catch {
    /* dev renderer without the preload handler */
  }
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <ErrorBoundary>
        <LocaleProvider initial={lang}>
          {mode === 'audience' ? <AudienceView /> : <App />}
        </LocaleProvider>
      </ErrorBoundary>
    </React.StrictMode>,
  )
}

void bootstrap()
