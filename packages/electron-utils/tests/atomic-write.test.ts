import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { rename } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { atomicWriteFile, renameWithRetry } from '../src/atomic-write'

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return { ...actual, rename: vi.fn(actual.rename) }
})

let dir: string

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
  vi.mocked(rename).mockClear()
})

const epermError = () =>
  Object.assign(new Error('EPERM: operation not permitted'), { code: 'EPERM' })

describe('atomicWriteFile', () => {
  it('replaces the target and leaves no temp file behind', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-utils-'))
    const target = join(dir, 'a.txt')
    writeFileSync(target, 'old')
    await atomicWriteFile(target, Buffer.from('new'))
    expect(readFileSync(target, 'utf-8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.txt'])
  })

  it('keeps the original intact and cleans the temp file when the write fails', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-utils-'))
    const target = join(dir, 'a.txt')
    writeFileSync(target, 'old')
    const bad = { byteLength: -1 } as unknown as Buffer
    await expect(atomicWriteFile(target, bad)).rejects.toThrow()
    expect(readFileSync(target, 'utf-8')).toBe('old')
    expect(readdirSync(dir)).toEqual(['a.txt'])
  })

  it('retries a transiently locked rename and still lands atomically', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-utils-'))
    const target = join(dir, 'a.txt')
    writeFileSync(target, 'old')
    vi.mocked(rename).mockRejectedValueOnce(epermError()).mockRejectedValueOnce(epermError())
    await atomicWriteFile(target, Buffer.from('new'))
    expect(readFileSync(target, 'utf-8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.txt'])
  })

  it('keeps the original intact when every rename attempt stays locked', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-utils-'))
    const target = join(dir, 'a.txt')
    writeFileSync(target, 'old')
    for (let i = 0; i < 5; i++) vi.mocked(rename).mockRejectedValueOnce(epermError())
    await expect(atomicWriteFile(target, Buffer.from('new'))).rejects.toMatchObject({
      code: 'EPERM',
    })
    expect(readFileSync(target, 'utf-8')).toBe('old')
    expect(readdirSync(dir)).toEqual(['a.txt'])
  })
})

describe('renameWithRetry', () => {
  it('renames successfully on the first attempt', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-rename-'))
    const src = join(dir, 'src.txt')
    const dst = join(dir, 'dst.txt')
    writeFileSync(src, 'content')
    await renameWithRetry(src, dst)
    expect(readFileSync(dst, 'utf-8')).toBe('content')
  })

  it('retries when transiently locked with EPERM', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-rename-'))
    const src = join(dir, 'src.txt')
    const dst = join(dir, 'dst.txt')
    writeFileSync(src, 'content')
    vi.mocked(rename).mockRejectedValueOnce(epermError())
    await renameWithRetry(src, dst)
    expect(readFileSync(dst, 'utf-8')).toBe('content')
  })
})
