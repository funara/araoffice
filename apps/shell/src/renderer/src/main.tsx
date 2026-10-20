import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlLang } from '@araoffice/i18n'
import { AppFrame } from './AppFrame'
import { LocaleProvider } from './locale'
import './home.css'
import './tabbar.css'

// macOS shell window is created with vibrancy; a transparent body lets the
// editor views' translucent regions (e.g. slides thumbnail pane) show it
if (navigator.platform.toLowerCase().includes('mac')) document.body.classList.add('vib')

// resolve persisted state before first paint so the onboarding overlay never
// flashes after the home screen
void Promise.all([
  window.aiOffice.getLanguage(),
  window.aiOffice.onboardingSeen().catch(() => true),
  window.aiOffice.getTheme().catch(() => 'system' as const),
]).then(([lang, onboardingSeen, theme]) => {
  document.documentElement.lang = htmlLang(lang)
  // resolve theme attribute before first paint to avoid flash
  const applyResolvedTheme = (t: string) => {
    if (t === 'system') {
      const isDark = window.matchMedia('(prefers-color-scheme: dark)').matches
      document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light')
    } else {
      document.documentElement.setAttribute('data-theme', t)
    }
  }
  applyResolvedTheme(theme)
  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
    void window.aiOffice.getTheme().then((current) => {
      if (current === 'system') {
        document.documentElement.setAttribute('data-theme', e.matches ? 'dark' : 'light')
      }
    })
  })
  createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <LocaleProvider initial={lang}>
        <AppFrame initialOnboardingSeen={onboardingSeen} initialTheme={theme} />
      </LocaleProvider>
    </React.StrictMode>,
  )
})
