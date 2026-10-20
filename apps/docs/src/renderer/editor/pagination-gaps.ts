import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'

const key = new PluginKey<DecorationSet>('paginationGaps')
const sheetsKey = new PluginKey<DecorationSet>('sectionSheets')
const masksKey = new PluginKey<DecorationSet>('sectionMasks')

/**
 * Always-on pagination in the canvas: renders a "page gap" widget before each
 * page-leading block (previous page's bottom margin + gray inter-page band +
 * next page's top margin). Decorations are visual only; document content and
 * saving are unaffected. Positions are driven by App's pagination measurement;
 * gaps map along transactions while editing and are rebuilt wholesale after a debounce.
 *
 * Second decoration layer (`setSectionSheets`): blocks whose section has a different
 * paper geometry than the canvas paper (mixed landscape/portrait documents) get their
 * own content width + centering + sheet background, so line wrapping, margins and
 * page size match Word per section instead of the cursor section's paper.
 */
export const PaginationGapsExtension = Extension.create({
  name: 'paginationGaps',

  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const next = tr.getMeta(key) as DecorationSet | undefined
            if (next) return next
            return set.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
        },
      }),
      new Plugin<DecorationSet>({
        key: sheetsKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const next = tr.getMeta(sheetsKey) as DecorationSet | undefined
            if (next) return next
            return set.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return sheetsKey.getState(state)
          },
        },
      }),
      new Plugin<DecorationSet>({
        key: masksKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const next = tr.getMeta(masksKey) as DecorationSet | undefined
            if (next) return next
            return set.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return masksKey.getState(state)
          },
        },
      }),
    ]
  },
})

/** A block's section-specific geometry (mixed-orientation canvas) */
export type SectionSheetSpec = {
  /** top-level node position */
  pos: number
  /** top-level node end position */
  end: number
  /** merged node attributes: class + inline content width / centering shift */
  attrs: Record<string, string>
}

/** Rebuild the section geometry node decorations (no-op when nothing changed); true when a dispatch happened */
export function setSectionSheets(view: EditorView, specs: SectionSheetSpec[]): boolean {
  const decos = specs.map((spec) => Decoration.node(spec.pos, spec.end, spec.attrs))
  const next = DecorationSet.create(view.state.doc, decos)
  const prev = sheetsKey.getState(view.state)
  if (prev && sameDecoStyles(prev, next)) return false
  view.dispatch(view.state.tr.setMeta(sheetsKey, next).setMeta('addToHistory', false))
  return true
}

function sameDecoStyles(a: DecorationSet, b: DecorationSet): boolean {
  const sig = (d: Decoration) =>
    `${d.from}:${d.to}:${String((d.spec as { attrs?: { style?: string } }).attrs?.style ?? '')}`
  const af = a.find()
  const bf = b.find()
  return af.length === bf.length && af.every((d, i) => sig(d) === sig(bf[i]))
}

/** One section's paper outline in a mixed-orientation canvas */
export type SectionMaskSpec = {
  /** top-level node position the (out of flow) mask widget hangs on */
  pos: number
  /** vertical extent in canvas coordinates, measured from the page content top */
  top: number
  height: number
  /** this section's paper box inside the canvas column */
  paperLeft: number
  paperWidth: number
  /** the canvas column (widest paper) width */
  columnWidth: number
}

/** Rebuild the section paper masks (no-op when nothing changed); true when a dispatch happened */
export function setSectionMasks(view: EditorView, specs: SectionMaskSpec[]): boolean {
  const decos = specs.map((spec) =>
    Decoration.widget(spec.pos, () => makeMaskEl(spec), {
      side: -2,
      key: `page-sheet-${spec.pos}-${spec.top}-${spec.height}-${spec.paperLeft}-${spec.paperWidth}`,
    }),
  )
  const next = DecorationSet.create(view.state.doc, decos)
  const prev = masksKey.getState(view.state)
  if (prev && sameDecoKeys(prev, next)) return false
  view.dispatch(view.state.tr.setMeta(masksKey, next).setMeta('addToHistory', false))
  return true
}

/**
 * The mask paints the canvas color outside a section's paper, so a narrower sheet inside
 * a wider canvas column reads as its own page size. It spans the section's whole height,
 * so page bodies and the page gaps between them stay covered.
 */
function makeMaskEl(spec: SectionMaskSpec): HTMLElement {
  const el = document.createElement('span')
  el.contentEditable = 'false'
  el.className = 'page-sheet-mask'
  el.style.top = `${spec.top}px`
  el.style.height = `${spec.height}px`
  el.style.width = `${spec.columnWidth}px`
  const right = spec.paperLeft + spec.paperWidth
  el.style.background =
    `linear-gradient(to right, var(--canvas) 0px ${spec.paperLeft}px,` +
    ` transparent ${spec.paperLeft}px ${right}px,` +
    ` var(--canvas) ${right}px 100%)`
  return el
}

function sameDecoKeys(a: DecorationSet, b: DecorationSet): boolean {
  const keyOf = (d: Decoration) => (d.spec as { key?: string }).key ?? ''
  const af = a.find()
  const bf = b.find()
  return af.length === bf.length && af.every((d, i) => keyOf(d) === keyOf(bf[i]))
}

export interface GapMetrics {
  marginTop: number
  marginBottom: number
  marginLeft: number
  marginRight: number
}

/** height of the gray inter-page band inside a page gap */
export const GAP_BAND = 28

function makeGapEl(
  m: GapMetrics,
  kind: 'block' | 'inline' | 'table' | 'cut',
  colSpan = 1,
): HTMLElement {
  const gap = document.createElement(
    kind === 'table' ? 'tr' : kind === 'inline' || kind === 'cut' ? 'span' : 'div',
  )
  gap.contentEditable = 'false'
  if (kind === 'cut') {
    // in-row table break cut point: zero-height dashed marker (takes no layout height, coordinate bookkeeping unchanged)
    gap.className = 'page-gap-cut'
    return gap
  }
  gap.style.height = `${m.marginBottom + GAP_BAND + m.marginTop}px`
  gap.style.setProperty('--gap-mb', `${m.marginBottom}px`)
  gap.style.setProperty('--gap-mt', `${m.marginTop}px`)
  gap.style.setProperty('--gap-ml', `${m.marginLeft}px`)
  gap.style.setProperty('--gap-mr', `${m.marginRight}px`)
  if (kind === 'table') {
    // A real spanning cell matching the table's exact column count is required here.
    // Setting colSpan=1000 created 1000 phantom columns in Chromium's fixed table layout,
    // crushing real columns into microscopic slivers on the left edge.
    gap.className = 'page-gap page-gap-inline page-gap-table'
    const cell = document.createElement('td')
    cell.colSpan = colSpan > 0 ? colSpan : 1
    cell.contentEditable = 'false'
    const fill = document.createElement('div')
    fill.className = 'page-gap-table-fill'
    cell.appendChild(fill)
    gap.appendChild(cell)
    return gap
  }
  gap.className = kind === 'inline' ? 'page-gap page-gap-inline' : 'page-gap'
  gap.style.marginLeft = `-${m.marginLeft}px`
  gap.style.marginRight = `-${m.marginRight}px`
  gap.style.width = `calc(100% + ${m.marginLeft + m.marginRight}px)`
  return gap
}

/** Page gap anchor: el = top-level page-leading block (inserted before it); pos = document position at a mid-paragraph/mid-table break */
export type PageGapSpec = {
  metrics: GapMetrics
  /** Number of table columns spanned by an in-table gap row */
  colSpan?: number
  /** Previous page's bottom footnote area (a ready-made positioned/sized element) and its content signature (change triggers rebuild) */
  notes?: HTMLElement
  notesKey?: string
  /** Previous page's footer / next page's header (ready-made positioned .page-gap-hf elements) and their content signature */
  hfEls?: HTMLElement[]
  hfKey?: string
  /** w:tblHeader repetition: cloned header rows rendered right after a table gap
   *  (the slicing engine already reserved their height on the new page) */
  repeatHeaderEls?: HTMLElement[]
  repeatHeaderKey?: string
} & ({ el: HTMLElement } | { pos: number; kind?: 'inline' | 'table' | 'cut' })

/** Rebuild all page gaps (an empty list clears them); each gap carries its own margins (sections differ) */
export function setPageGaps(view: EditorView, gaps: PageGapSpec[]): void {
  const decos: Decoration[] = []
  let ordinal = -1
  for (const gap of gaps) {
    ordinal++
    const { metrics, notes, hfEls } = gap
    let pos: number
    let kind: 'block' | 'inline' | 'table' | 'cut'
    if ('el' in gap) {
      kind = 'block'
      try {
        const inside = view.posAtDOM(gap.el, 0)
        pos = view.state.doc.resolve(inside).before(1)
      } catch {
        continue
      }
    } else {
      pos = gap.pos
      kind = gap.kind ?? 'inline'
    }
    const mKey = `${metrics.marginTop},${metrics.marginBottom},${metrics.marginLeft},${metrics.marginRight}`
    decos.push(
      Decoration.widget(
        pos,
        (editorView) => {
          let colSpan = gap.colSpan
          if (!colSpan && kind === 'table') {
            try {
              const $p = (editorView ?? view).state.doc.resolve(pos)
              for (let d = $p.depth; d > 0; d--) {
                const n = $p.node(d)
                if (n.type.name === 'docTableRow' || n.type.name === 'docTable') {
                  const row = n.type.name === 'docTableRow' ? n : n.firstChild
                  if (row) {
                    let count = 0
                    row.forEach((c) => {
                      count += Number(c.attrs.colspan) || 1
                    })
                    if (count > 0) colSpan = count
                  }
                  break
                }
              }
            } catch {
              /* ignore */
            }
          }
          const el = makeGapEl(metrics, kind, colSpan ?? 1)
          if (notes) el.appendChild(notes)
          if (hfEls && (kind === 'block' || kind === 'inline')) {
            for (const hf of hfEls) el.appendChild(hf)
          } else if (hfEls && kind === 'table') {
            // table-row gaps position their strips inside the absolutely-filled cell;
            // only zero-height cut markers still can't carry them
            const fill = el.querySelector('.page-gap-table-fill')
            if (fill) for (const hf of hfEls) fill.appendChild(hf)
          }
          return el
        },
        {
          side: -1,
          // keyed by page ordinal, NOT pos: edits above a gap shift its mapped
          // position without changing the page, so an ordinal key lets sameGaps
          // skip the dispatch entirely and lets PM reuse the widget DOM when a
          // dispatch does happen
          key: `page-gap-${kind[0]}-${ordinal}-${mKey}${gap.notesKey ? `-${gap.notesKey}` : ''}${gap.hfKey ? `-${gap.hfKey}` : ''}`,
        },
      ),
    )
    // repeated header rows (w:tblHeader) directly after the table gap: one widget per
    // cloned tr, side 0 so they land between the gap (side -1) and the split row
    if (kind === 'table' && gap.repeatHeaderEls?.length) {
      gap.repeatHeaderEls.forEach((rowEl, i) => {
        decos.push(
          Decoration.widget(pos, () => rowEl, {
            side: 0,
            key: `page-gap-rh-${pos}-${i}-${gap.repeatHeaderKey ?? ''}`,
          }),
        )
      })
    }
  }
  const next = DecorationSet.create(view.state.doc, decos)
  const prev = key.getState(view.state)
  if (prev && sameGaps(prev, next)) return
  view.dispatch(view.state.tr.setMeta(key, next).setMeta('addToHistory', false))
}

function sameGaps(a: DecorationSet, b: DecorationSet): boolean {
  const keyOf = (d: Decoration) => (d.spec as { key?: string }).key ?? ''
  const af = a.find()
  const bf = b.find()
  return (
    af.length === bf.length &&
    af.every((d, i) => d.from === bf[i].from && keyOf(d) === keyOf(bf[i]))
  )
}
