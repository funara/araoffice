import { describe, expect, it } from 'vitest'
import { hasUnsavedEdits } from '../src/renderer/dirty'

/** A freshly opened document: every edit channel empty */
const clean = {
  markups: [],
  drawings: [],
  stampCfg: null,
  formEdits: new Map(),
  rotations: new Map(),
  deleted: new Set(),
  order: null,
  metadata: null,
}

describe('hasUnsavedEdits', () => {
  it('a document with no edits is clean', () => {
    expect(hasUnsavedEdits(clean)).toBe(false)
  })

  it('each edit channel alone marks the document dirty', () => {
    expect(hasUnsavedEdits({ ...clean, markups: [{}] })).toBe(true)
    expect(hasUnsavedEdits({ ...clean, drawings: [{}] })).toBe(true)
    expect(hasUnsavedEdits({ ...clean, stampCfg: {} })).toBe(true)
    expect(hasUnsavedEdits({ ...clean, formEdits: new Map([['f', {}]]) })).toBe(true)
    expect(hasUnsavedEdits({ ...clean, rotations: new Map([[0, 90]]) })).toBe(true)
    expect(hasUnsavedEdits({ ...clean, deleted: new Set([0]) })).toBe(true)
    expect(hasUnsavedEdits({ ...clean, order: [1, 0] })).toBe(true)
    expect(hasUnsavedEdits({ ...clean, metadata: {} })).toBe(true)
  })
})
