/**
 * Where a line may break: Unicode UAX #14 line-break opportunities, then a kinsoku
 * post-filter — the same two-stage structure a word processor uses (and the same one
 * WordCraft's layout crate uses: a UAX #14 pass followed by `kinsoku::apply`).
 *
 * This replaces the ad-hoc "break at whitespace, else per CJK character" rule that the
 * old simulator used as it walked characters. The difference matters for the cases the
 * old rule got wrong: it broke between any two CJK characters (so a line could start
 * with a closing bracket or 、), it never broke after a hyphen or slash inside a Latin
 * word, and it treated a run of spaces as one indivisible unit.
 *
 * The opportunity list is deliberately a covered subset of UAX #14 rather than the full
 * pair table: it implements the classes that decide line breaking in the languages the
 * editor supports today (Latin, CJK, and the punctuation they mix). Unknown characters
 * default to the ideographic/graphic class, matching Word's behaviour of allowing a
 * break between two ideographs.
 */

/** Characters that may not begin a line (kinsoku shori: line-start prohibition). */
const NO_LINE_START =
  '、。，．・：；？！゛゜ヽヾゝゞ々ー』」）］｝〉》」』】〙〗〟’＂｠»' +
  ',.;:!?%)]}\u2019\u201d\u203a\u00bb\u2026\u2014\u2013\u30fb\uff0c\uff0e\uff1a\uff1b\uff01\uff1f'

/** Characters that may not end a line (kinsoku shori: line-end prohibition). */
const NO_LINE_END = '‘“（［｛〈《「『【〘〖〝＂｟«' + '([{\u2018\u201c\u2039\u00ab'

/** Opening brackets/quotes and their ASCII equivalents, for the pair rules above. */
function isNoLineStart(ch: string): boolean {
  return NO_LINE_START.includes(ch)
}
function isNoLineEnd(ch: string): boolean {
  return NO_LINE_END.includes(ch)
}

/** East Asian wide/fullwidth ranges whose characters may break on either side. */
export function isEastAsian(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x115f) || // Hangul Jamo
    (cp >= 0x2e80 && cp <= 0x303e) || // CJK radicals, Kangxi, CJK symbols
    (cp >= 0x3041 && cp <= 0x33ff) || // Hiragana, Katakana, Bopomofo, compat
    (cp >= 0x3400 && cp <= 0x4dbf) || // CJK Ext A
    (cp >= 0x4e00 && cp <= 0x9fff) || // CJK Unified
    (cp >= 0xa000 && cp <= 0xa4cf) || // Yi
    (cp >= 0xac00 && cp <= 0xd7a3) || // Hangul syllables
    (cp >= 0xf900 && cp <= 0xfaff) || // CJK compat ideographs
    (cp >= 0xfe30 && cp <= 0xfe4f) || // CJK compat forms
    (cp >= 0xff00 && cp <= 0xff60) || // Fullwidth forms
    (cp >= 0xffe0 && cp <= 0xffe6) ||
    (cp >= 0x20000 && cp <= 0x3fffd) // CJK Ext B+
  )
}

/** Identifiers/hyphenation-adjacent ASCII that must not be split (kept in one word). */
function isWordChar(ch: string): boolean {
  const cp = ch.codePointAt(0) ?? 0
  if (isEastAsian(cp)) return false
  // letters, digits, marks, apostrophes inside words, underscore
  return /[\p{L}\p{N}\p{M}_'\u2019]/u.test(ch)
}

/** Characters a break may follow inside an otherwise unbreakable Latin word. */
function allowsBreakAfter(ch: string): boolean {
  return ch === '-' || ch === '/' || ch === '\u2010' || ch === '\u2013' || ch === '\u2014'
}

/**
 * UAX #14 line-break opportunities: the indices at which a line may start.
 * `0` and `text.length` are never included; the result is ascending and unique.
 */
export function lineBreakOpportunities(text: string): number[] {
  const chars = [...text]
  const out: number[] = []
  // Offsets are in UTF-16 code units so callers can slice the original string.
  const offsets: number[] = []
  let off = 0
  for (const ch of chars) {
    offsets.push(off)
    off += ch.length
  }
  offsets.push(off)

  const push = (i: number) => {
    const at = offsets[i]
    if (at > 0 && at < text.length && out[out.length - 1] !== at) out.push(at)
  }

  for (let i = 1; i < chars.length; i++) {
    const prev = chars[i - 1]
    const cur = chars[i]
    const prevCp = prev.codePointAt(0) ?? 0
    const curCp = cur.codePointAt(0) ?? 0

    // A break after a space is always allowed (the space ends the line).
    if (prev === ' ' || prev === '\t') {
      push(i)
      continue
    }

    // An explicit hyphen/slash inside a word is a break opportunity, unless the next
    // character is a digit (Word does not break inside "-12" or "3/4").
    if (allowsBreakAfter(prev) && !/[0-9]/.test(cur)) {
      push(i)
      continue
    }

    // East Asian: a break is allowed on either side of an ideograph/kana/hangul
    // syllable; the kinsoku filter then removes the forbidden ones.
    if (isEastAsian(curCp) || isEastAsian(prevCp)) {
      // No break before a combining mark or a variation selector.
      if (curCp >= 0x300 && curCp <= 0x36f) continue
      push(i)
      continue
    }

    // Between two non-word characters (e.g. ")" then "(") a break is allowed.
    if (!isWordChar(prev) && !isWordChar(cur)) {
      push(i)
      continue
    }

    // Everything else (letter↔letter, digit↔letter, letter↔digit) stays together.
  }
  return out
}

/**
 * Kinsoku post-filter: drop opportunities that would strand a character that may not
 * begin a line, or leave one that may not end a line. With `wordWrap` off, Word also
 * allows a break between any two East Asian characters, which is added here.
 */
export function applyKinsoku(text: string, breaks: number[], wordWrap = true): number[] {
  const chars = [...text]
  const offsets: number[] = []
  let off = 0
  for (const ch of chars) {
    offsets.push(off)
    off += ch.length
  }
  const indexOf = new Map<number, number>()
  offsets.forEach((o, i) => indexOf.set(o, i))

  const allowed = new Set<number>(breaks)
  if (!wordWrap) {
    for (let i = 1; i < chars.length; i++) {
      const a = chars[i - 1].codePointAt(0) ?? 0
      const b = chars[i].codePointAt(0) ?? 0
      if (isEastAsian(a) && isEastAsian(b)) allowed.add(offsets[i])
    }
  }

  const out: number[] = []
  for (const at of [...allowed].sort((a, b) => a - b)) {
    const i = indexOf.get(at)
    if (i === undefined) continue
    const before = chars[i - 1]
    const after = chars[i]
    if (before && isNoLineEnd(before)) continue
    if (after && isNoLineStart(after)) continue
    out.push(at)
  }
  return out
}

/** Break opportunities for a paragraph: UAX #14, then kinsoku. */
export function breakOpportunities(text: string, wordWrap = true): number[] {
  return applyKinsoku(text, lineBreakOpportunities(text), wordWrap)
}
