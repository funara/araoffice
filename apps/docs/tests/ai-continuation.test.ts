import { describe, expect, it } from 'vitest'
import { CONTINUE_INSTRUCTION } from '@araoffice/agent-core'

describe('Docs AI continuation', () => {
  it('instructs continuation to preserve completed work', () => {
    expect(CONTINUE_INSTRUCTION).toContain('existing conversation and document state')
    expect(CONTINUE_INSTRUCTION).toContain('only the outstanding work')
    expect(CONTINUE_INSTRUCTION).toContain('Do not repeat edits')
  })
})
