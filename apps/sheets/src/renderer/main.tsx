import ReactDOM from 'react-dom/client'
import { htmlLang, type Lang } from '@araoffice/i18n'
import { ErrorBoundary } from '@araoffice/ui'

import '@univerjs/preset-sheets-core/lib/index.css'

import { App } from './App'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import './styles.css'
import '@araoffice/ui/ai-panel.css'

if (import.meta.hot) {
  import.meta.hot.on('vite:beforeUpdate', ({ updates }) => {
    const replacesUniverRuntime = updates.some(
      ({ path }) => path.endsWith('/App.tsx') || path.endsWith('/univer-sync.ts'),
    )
    if (replacesUniverRuntime) window.location.reload()
  })
}

const root = document.getElementById('root')
if (!root) throw new Error('Missing application root.')

async function bootstrap(): Promise<void> {
  let lang: Lang = 'zh'
  try {
    lang = await window.desktopApi.getLanguage()
  } catch {
    /* dev renderer without the preload handler */
  }
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  ReactDOM.createRoot(root!).render(
    <ErrorBoundary>
      <LocaleProvider initial={lang}>
        <App />
      </LocaleProvider>
    </ErrorBoundary>,
  )
}

void bootstrap()
