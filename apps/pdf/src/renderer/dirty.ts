/** Structural inputs to the dirty predicate: lengths, sizes and nullability are all that matter */
export interface DirtyInputs {
  markups: readonly unknown[]
  drawings: readonly unknown[]
  stampCfg: unknown
  formEdits: { readonly size: number }
  rotations: { readonly size: number }
  deleted: { readonly size: number }
  order: readonly unknown[] | null
  metadata: unknown
}

/** Any unsaved edit present — single source of truth for the unsaved badge, autosave gate, close guard and open guard */
export function hasUnsavedEdits(d: DirtyInputs): boolean {
  return (
    d.markups.length > 0 ||
    d.drawings.length > 0 ||
    d.stampCfg !== null ||
    d.formEdits.size > 0 ||
    d.rotations.size > 0 ||
    d.deleted.size > 0 ||
    d.order !== null ||
    d.metadata !== null
  )
}
