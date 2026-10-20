import type { Editor } from '@tiptap/core'

/** Word's ribbon indent step: half an inch, in twips */
const STEP = 720
/** Default hanging indent for list markers: a quarter inch, in twips */
const DEFAULT_LIST_HANG_TWIPS = 360

/** Ribbon increase/decrease indent: list items change level or align to margin, plain paragraphs and headings snap to the next half-inch stop */
export function stepParagraphIndent(editor: Editor, delta: 1 | -1): boolean {
  return editor
    .chain()
    .focus()
    .command(({ state, tr, dispatch }) => {
      const { from, to } = state.selection
      let changed = false
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (node.type.name === 'docListItem') {
          const ilvl = Number(node.attrs.ilvl) || 0
          const curIndent = node.attrs.indentLeft != null ? Number(node.attrs.indentLeft) : null

          if (delta > 0) {
            // Increase indent:
            if (curIndent === 0 && ilvl === 0) {
              tr.setNodeMarkup(pos, undefined, {
                ...node.attrs,
                indentLeft: DEFAULT_LIST_HANG_TWIPS,
              })
              changed = true
            } else if (curIndent === DEFAULT_LIST_HANG_TWIPS && ilvl === 0) {
              tr.setNodeMarkup(pos, undefined, { ...node.attrs, indentLeft: null })
              changed = true
            } else {
              // Advance list level (up to 8)
              const nextIlvl = Math.min(ilvl + 1, 8)
              if (nextIlvl !== ilvl) {
                tr.setNodeMarkup(pos, undefined, {
                  ...node.attrs,
                  ilvl: nextIlvl,
                  indentLeft: null,
                })
                changed = true
              }
            }
          } else {
            // Decrease indent:
            if (ilvl > 0) {
              // Sub-level item: decrement level
              tr.setNodeMarkup(pos, undefined, { ...node.attrs, ilvl: ilvl - 1 })
              changed = true
            } else {
              // Level 0:
              if (curIndent == null || curIndent > DEFAULT_LIST_HANG_TWIPS) {
                // Shift marker flush with left margin (360 twips)
                tr.setNodeMarkup(pos, undefined, {
                  ...node.attrs,
                  indentLeft: DEFAULT_LIST_HANG_TWIPS,
                })
                changed = true
              } else if (curIndent === DEFAULT_LIST_HANG_TWIPS) {
                // Shift both marker and text to margin (0 twips)
                tr.setNodeMarkup(pos, undefined, { ...node.attrs, indentLeft: 0 })
                changed = true
              }
              // When already at 0, stop at 0 and NEVER delete list numbering/bullets
            }
          }
          return
        }

        if (node.type.name !== 'docParagraph' && node.type.name !== 'docHeading') return
        const cur = Number(node.attrs.indentLeft) || 0
        const next =
          delta > 0
            ? Math.floor(cur / STEP) * STEP + STEP
            : Math.max(Math.ceil(cur / STEP) * STEP - STEP, 0)
        if (next === cur) return
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, indentLeft: next || null })
        changed = true
      })
      if (changed && dispatch) dispatch(tr)
      return changed
    })
    .run()
}
