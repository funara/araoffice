import type { ParsedDoc, Run } from '@araoffice/docx-engine'
import { canonicalFontFamily, cssFontFamily, isFontAvailable, textHasCjk } from './line-metrics'

/** bundled web fonts (fonts.css) — always "available" but they are the
 *  one-size-fits-all fallback, so landing on them still counts as a miss */
const BUNDLED = new Set([
  'Noto Sans CJK SC',
  'Noto Serif CJK SC',
  'Carlito',
  'Caladea',
  'Liberation Serif',
  'Liberation Sans',
  'Liberation Mono',
])

const GENERIC = new Set(['serif', 'sans-serif', 'monospace'])

/**
 * Every distinct font family the document actually renders with.
 *
 * Word's slots decide which declaration matters: a run's Latin text uses the
 * ascii/hAnsi slot (falling back to the inherited style font when only an East
 * Asian slot is declared), a run's CJK text uses the eastAsia slot. Fonts that no
 * text ever renders with — unused style definitions, East Asian slots on Latin-only
 * runs — are not collected, so they can't raise a "missing font" notice.
 */
export function collectDocFonts(parsed: ParsedDoc): string[] {
  const names = new Set<string>()
  const add = (font?: string) => {
    if (font) names.add(font)
  }
  const addRun = (r: Run) => {
    if (textHasCjk(r.text ?? '')) {
      // mixed/CJK text renders with the East Asian slot and the Latin slot
      add(r.font)
      add(r.fontAscii)
      return
    }
    // Latin-only text renders with the ascii/hAnsi slot; an eastAsia-only
    // declaration (font set, fontAscii unset) never renders anything
    add(r.fontAscii)
  }
  const usedStyleIds = new Set<string>()
  for (const b of parsed.blocks) {
    if (b.styleId) usedStyleIds.add(b.styleId)
    for (const r of b.runs ?? []) {
      if (r.styleId) usedStyleIds.add(r.styleId)
      addRun(r)
    }
    for (const tb of b.textboxes ?? []) {
      for (const p of tb.paras) for (const r of p.runs) addRun(r)
    }
  }
  add(parsed.docDefaults?.asciiFont)
  add(parsed.docDefaults?.eastAsiaFont)
  for (const id of usedStyleIds) add(parsed.styles.get(id)?.display?.font)
  return [...names]
}

// document.fonts.check is useless here: Chromium answers true for any family
// (system-fallback rendering counts) — delegate to line-metrics isFontAvailable,
// which measures against generic fallbacks with a result cache.

export interface FontSubstitution {
  name: string
  /** first available family in the fallback chain, null = only bundled fallback left */
  substitute: string | null
}

/**
 * Declared fonts that don't resolve on this system, with the family that
 * actually renders instead (substitution used to be silent). PostScript-style
 * names (TimesNewRomanPSMT) count as present when their family is installed —
 * that is what Word itself renders with.
 */
export function checkMissingFonts(names: string[]): FontSubstitution[] {
  if (typeof document === 'undefined') return []
  const out: FontSubstitution[] = []
  for (const name of names) {
    const canonical = canonicalFontFamily(name)
    if (isFontAvailable(canonical) || (canonical !== name && isFontAvailable(name))) continue
    const chain = cssFontFamily(canonical)
      .split(',')
      .map((s) => s.trim().replace(/^'|'$/g, ''))
      .filter((c) => c !== name && c !== canonical && !GENERIC.has(c))
    const substitute = chain.find((c) => !BUNDLED.has(c) && isFontAvailable(c)) ?? null
    out.push({ name, substitute })
  }
  return out
}
