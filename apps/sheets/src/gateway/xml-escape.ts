const NAMED_ENTITIES: Record<string, string> = {
  quot: '"',
  apos: "'",
  lt: '<',
  gt: '>',
  amp: '&',
}

/// Single-pass decode of the XML named entities plus numeric character
/// references (&#dd; / &#xhh;) — producers may encode non-ASCII sheet names
/// as numeric references, which sequential replaceAll would miss.
/// Canonical copy; was duplicated x4 across gateway files (keep this full decoder).
export function decodeEntities(input: string): string {
  return input.replace(
    /&(?:#x([0-9A-Fa-f]+)|#([0-9]+)|(quot|apos|lt|gt|amp));/g,
    (match, hex: string | undefined, dec: string | undefined, named: string | undefined) => {
      if (named !== undefined) return NAMED_ENTITIES[named] ?? match
      const code = hex === undefined ? Number(dec) : Number.parseInt(hex, 16)
      return code <= 0x10ffff ? String.fromCodePoint(code) : match
    },
  )
}

export function decodeAttribute(input: string): string {
  return decodeEntities(input)
}

export function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

export function escapeXmlAttribute(input: string): string {
  return escapeXmlText(input).replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

export function escapeRegExp(input: string): string {
  return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
