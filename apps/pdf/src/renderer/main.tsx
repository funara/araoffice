import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@araoffice/i18n'
import { ErrorBoundary } from '@araoffice/ui'
import App from './App'
import { LocaleProvider } from './i18n/locale'
import './styles.css'
import '@araoffice/ui/ai-panel.css'

void (async () => {
  const lang: Lang = await window.pdfApi.getLanguage().catch(() => 'zh' as const)
  document.documentElement.lang = htmlLang(lang)
  createRoot(document.getElementById('root')!).render(
    <ErrorBoundary>
      <LocaleProvider initial={lang}>
        <App />
      </LocaleProvider>
    </ErrorBoundary>,
  )
})()
