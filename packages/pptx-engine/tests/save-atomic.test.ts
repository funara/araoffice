import { describe, it, expect, vi } from 'vitest'
import { mkdtempSync, readFileSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { openPptx, savePptxToFile } from '../src/index'
import { rename } from 'node:fs/promises'

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return { ...actual, rename: vi.fn(actual.rename) }
})

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('savePptxToFile atomic write', () => {
  it('keeps the original byte-identical and cleans the temp file when the rename fails', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const dir = mkdtempSync(join(tmpdir(), 'save-atomic-'))
    const target = join(dir, 'out.pptx')

    await savePptxToFile(opened, target)
    const before = readFileSync(target)

    vi.mocked(rename).mockRejectedValueOnce(new Error('disk full'))
    await expect(savePptxToFile(opened, target)).rejects.toThrow('disk full')

    expect(readFileSync(target).equals(before)).toBe(true)
    expect(readdirSync(dir).filter((name) => name.includes('.tmp-'))).toEqual([])
  })
})
