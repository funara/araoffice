import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@araoffice/i18n'
import { ErrorBoundary } from '@araoffice/ui'
import { App } from './App'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import './styles.css'
import '@araoffice/ui/ai-panel.css'
import './fonts/fonts.css'

async function bootstrap(): Promise<void> {
  let lang: Lang = 'zh'
  try {
    lang = await window.desktop.getLanguage()
  } catch {
    /* dev renderer without the preload handler */
  }
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  createRoot(document.getElementById('root')!).render(
    <ErrorBoundary>
      <LocaleProvider initial={lang}>
        <App />
      </LocaleProvider>
    </ErrorBoundary>,
  )
}

void bootstrap()
