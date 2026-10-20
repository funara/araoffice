import type { AnnotMatch } from '../shared/ipc'
import type { StringKey } from './i18n/locale'

/**
 * The comments panel's source data: a flat row per annotation the viewer lists, merged with the
 * still-unsaved markup additions. Subtyping here decides two things:
 * - `Widget` and `Link` never become rows (form fields and navigation are not comments)
 * - file-side deletion matches on the decoded identity, because pdf.js row ids are synthesized
 */

/** One listed comment: what the user reads plus the identity the deleter matches on */
export interface CommentRow {
  /** Pending local markups only: the markups-state id, so row removal filters by it */
  id?: string
  /** 0-based original page index (the saved file's order) */
  pageIndex: number
  /** exact identity for pdf:delete-annotations */
  match: AnnotMatch
  /** author-then-text readout, trimmed for the one-line panel row */
  excerpt: string
  /** display label; a strings key when one exists, else the raw subtype */
  label: string
  /** true for the file's annotations, false for pending unsaved markups */
  saved: boolean
}

/** pdf.js subtype → existing strings key; absent subtypes fall back to their raw name */
const SUBTYPE_KEY: Record<string, StringKey> = {
  Text: 'drawNote',
  FreeText: 'drawNote',
  Highlight: 'highlight',
  Underline: 'underline',
  StrikeOut: 'strikeout',
  Squiggly: 'squiggly',
}

const EXCERPT_LEN = 80

const trim = (s: string): string => {
  const flat = s.replace(/\s+/g, ' ').trim()
  return flat.length > EXCERPT_LEN ? `${flat.slice(0, EXCERPT_LEN)}…` : flat
}

/**
 * pdf.js annotation objects → rows. `labelFor` resolves a subtype through the app strings so
 * the panel stays lenguaje-parallel with the draw tools; unknown subtypes show raw.
 */
export function annotationsToRows(
  annots: unknown[],
  pageIndex: number,
  labelOf: (subtype: string, key: StringKey) => string,
): CommentRow[] {
  const rows: CommentRow[] = []
  for (const a of annots) {
    // pdf.js annot data is a loose record; the fields we read are the ones FormLayer reads too
    const a2 = a as {
      subtype?: string
      rect?: number[]
      contents?: string
      title?: string
      name?: string
    }
    const subtype = a2.subtype ?? ''
    if (!subtype || subtype === 'Widget' || subtype === 'Link' || subtype === 'Popup') continue
    const text = (a2.contents ?? '').trim()
    const author = (a2.title ?? '').trim()
    const excerpt = trim(author ? (text ? `${author}: ${text}` : author) : text)
    // Known subtypes resolve through the app strings (same keys the draw tools use); the rest
    // show their raw PDF subtype, exactly as viewers do
    const labelKey = SUBTYPE_KEY[subtype]
    const label = labelKey ? labelOf(subtype, labelKey) : subtype
    rows.push({
      pageIndex,
      match: {
        pageIndex,
        subtype,
        rect: Array.isArray(a2.rect) ? a2.rect.map(Number) : [],
        contents: text,
        name: (a2.name ?? '').trim(),
      },
      excerpt: excerpt || label,
      label,
      saved: true,
    })
  }
  return rows
}
