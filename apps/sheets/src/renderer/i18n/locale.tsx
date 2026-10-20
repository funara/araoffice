import { createContext, useContext, useEffect, useState } from 'react'
import type { ReactNode } from 'react'
import {
  AI_LANG_DIRECTIVES,
  DATE_LOCALES,
  createI18n,
  htmlLang,
  type Lang,
  type Params,
} from '@araoffice/i18n'
import { strings } from './strings'

const translate = createI18n(strings)

export type StringKey = keyof typeof strings.zh
export type TFunc = (key: StringKey, params?: Params) => string

// mirror for non-React modules (edit-journal, lazy-plan, AI tools …);
// set before first render and on every language switch
let moduleLang: Lang = 'zh'
export const getLang = (): Lang => moduleLang
export const setModuleLang = (lang: Lang): void => {
  moduleLang = lang
}
/** module-level translator — components should prefer useI18n().t so they re-render on switch */
export const t: TFunc = (key, params) => translate(moduleLang, key, params)

const SHEETS_LANG_OVERRIDES: Partial<Record<Lang, string>> = {
  en: "\n\nSTRICT LANGUAGE RULE: Reply and write all thoughts, summaries, and messages in the exact same language as the user's message (e.g. Indonesian for Indonesian, English for English). NEVER output Chinese unless the user explicitly used Chinese.",
  id: '\n\nATURAN BAHASA KETAT: Balas dan tulis semua pemikiran, ringkasan operasi, dan pesan dalam bahasa yang sama dengan pesan pengguna (misalnya bahasa Indonesia). JANGAN PERNAH menampilkan bahasa Mandarin kecuali pengguna secara eksplisit memintanya.',
}

/** appended to the agent system prompt: replies follow the user's message language, falling back to the UI language */
export function aiLangDirective(): string {
  return SHEETS_LANG_OVERRIDES[moduleLang] ?? AI_LANG_DIRECTIVES[moduleLang]
}
export { DATE_LOCALES } from '@araoffice/i18n'

const LocaleContext = createContext<Lang>('zh')

export function LocaleProvider({ initial, children }: { initial: Lang; children: ReactNode }) {
  const [lang, setLang] = useState<Lang>(initial)
  setModuleLang(initial)
  useEffect(
    () =>
      window.desktopApi.onLanguageChanged((next) => {
        setModuleLang(next)
        document.documentElement.lang = htmlLang(next)
        setLang(next)
      }),
    [],
  )
  return <LocaleContext.Provider value={lang}>{children}</LocaleContext.Provider>
}

export interface I18n {
  lang: Lang
  t: TFunc
  /** BCP-47 locale for date/number formatting */
  dateLocale: string
}

export function useI18n(): I18n {
  const lang = useContext(LocaleContext)
  return {
    lang,
    t: (key, params) => translate(lang, key, params),
    dateLocale: DATE_LOCALES[lang],
  }
}
