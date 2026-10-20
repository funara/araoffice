import { describe, expect, it, vi } from 'vitest'
import { mkdtempSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { isExternallyModified } from '../src/main/external-change'

// Slides' disk fingerprint rides on the shared pure helper; this pins the wiring
// contract the save handler relies on: stat-before-hash, and a missing file
// treated as no conflict.
describe('slides external-change wiring contract', () => {
  it('untouched file (same mtime+size) skips the hash read', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'slides-ext-'))
    const file = join(dir, 'deck.pptx')
    writeFileSync(file, 'aaa')
    const s = statSync(file)
    const readHash = vi.fn(() => null)
    await expect(
      isExternallyModified(
        { mtimeMs: s.mtimeMs, size: s.size, hash: 'h' },
        { mtimeMs: s.mtimeMs, size: s.size },
        readHash,
      ),
    ).resolves.toBe(false)
    expect(readHash).not.toHaveBeenCalled()
  })

  it('content rewrite by another program flips to modified', async () => {
    writeFileSync(join(mkdtempSync(join(tmpdir(), 'slides-ext-')), 'd.pptx'), 'aaa')
    await expect(
      isExternallyModified(
        { mtimeMs: 1000, size: 3, hash: 'h-old' },
        { mtimeMs: 2000, size: 9 },
        () => String('h-new'),
      ),
    ).resolves.toBe(true)
  })

  it('a deleted file is not a conflict (save recreates it)', async () => {
    await expect(
      isExternallyModified({ mtimeMs: 1, size: 1, hash: 'h' }, null, () => null),
    ).resolves.toBe(false)
  })
})
