/**
 * Flow for replacing a workbook with unsaved edits: prompt Save/Discard/Cancel,
 * save before opening when chosen, and treat a failed save as cancel. The UI
 * prompt lives in UnsavedOpenDialog; this module keeps the decision flow
 * testable without DOM.
 */
export type UnsavedChoice = 'save' | 'discard' | 'cancel'

export interface OpenGuardIo {
  /// Dirty check over the currently open workbook (the same journal signal
  /// save and the close prompt use).
  hasDirtyWorkbook: () => boolean
  /// Shows the Save/Discard/Cancel prompt and waits for the choice.
  prompt: () => Promise<UnsavedChoice>
  /// Runs the save; resolves false when it failed (stay on the workbook).
  save: () => Promise<boolean>
}

/** Resolves the guard, then opens. Returns false when the open was canceled. */
export async function openGuardFlow(io: OpenGuardIo, open: () => void): Promise<boolean> {
  if (io.hasDirtyWorkbook()) {
    const choice = await io.prompt()
    if (choice === 'cancel') return false
    if (choice === 'save' && !(await io.save())) return false
  }
  open()
  return true
}
