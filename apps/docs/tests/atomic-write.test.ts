import { describe, expect, it } from 'vitest'
import { looksLikeZip } from '../src/main/atomic-write'

describe('looksLikeZip', () => {
  it('accepts a zip local-file header and rejects truncated or garbage bytes', () => {
    expect(looksLikeZip(Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00]))).toBe(true)
    expect(looksLikeZip(Buffer.from([0x50, 0x4b]))).toBe(false)
    expect(looksLikeZip(Buffer.from('not a zip'))).toBe(false)
    expect(looksLikeZip(Buffer.alloc(0))).toBe(false)
  })
})
