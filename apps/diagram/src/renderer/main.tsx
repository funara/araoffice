import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@araoffice/i18n'
import { ErrorBoundary } from '@araoffice/ui'
import App from './App'
import './styles.css'
import '@araoffice/ui/ai-panel.css'

void (async () => {
  const lang: Lang = (await window.diagramApi?.getLanguage?.().catch(() => 'en' as const)) || 'en'
  document.documentElement.lang = htmlLang(lang)
  const rootEl = document.getElementById('root')
  if (rootEl) {
    createRoot(rootEl).render(
      <ErrorBoundary>
        <App initialLang={lang} />
      </ErrorBoundary>,
    )
  }
})()
