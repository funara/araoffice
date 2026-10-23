import { afterEach, describe, expect, it, vi } from 'vitest'
import { mkdtempSync, readFileSync, readdirSync, renameSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return { ...actual, renameSync: vi.fn(actual.renameSync) }
})

import { ProjectStore } from '../src/store.js'

let tmpDir: string

afterEach(() => {
  rmSync(tmpDir, { recursive: true, force: true })
  vi.mocked(renameSync).mockClear()
})

describe('project JSON writes', () => {
  it('preserves the old file when all atomic renames fail', async () => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-write-'))
    const store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
    const project = store.createProject('Before')
    const projectDir = join(tmpDir, 'projects', project.id)
    const projectPath = join(projectDir, 'project.json')
    const before = readFileSync(projectPath, 'utf8')
    const actual = await vi.importActual<typeof import('node:fs')>('node:fs')
    vi.mocked(renameSync).mockImplementation((source, target) => {
      if (String(target) === projectPath) {
        throw Object.assign(new Error('target locked'), { code: 'EPERM' })
      }
      return actual.renameSync(source, target)
    })

    expect(() => store.renameProject(project.id, 'After')).toThrow('target locked')
    expect(readFileSync(projectPath, 'utf8')).toBe(before)
    expect(readdirSync(projectDir)).not.toContain('project.json.tmp')
  })
})
