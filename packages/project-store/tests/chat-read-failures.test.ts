import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const controls = vi.hoisted(() => ({
  sourceReadFailures: new Set<string>(),
  shortRead: { active: false, calls: 0, payload: '' },
}))

vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    readFileSync: vi.fn((filePath: string | URL, encoding?: BufferEncoding) => {
      if (controls.sourceReadFailures.has(String(filePath))) {
        throw Object.assign(new Error('source read failed'), { code: 'EIO' })
      }
      return encoding ? actual.readFileSync(filePath, encoding) : actual.readFileSync(filePath)
    }),
    readSync: vi.fn(
      (
        fd: number,
        buffer: ArrayBufferView,
        offset: number,
        length: number,
        position: number | null,
      ) => {
        const bytes = Buffer.from(buffer.buffer, buffer.byteOffset, buffer.byteLength)
        if (!controls.shortRead.active)
          return actual.readSync(fd, bytes as NodeJS.ArrayBufferView, offset, length, position)
        controls.shortRead.calls++
        if (controls.shortRead.calls > 1) return 0
        bytes.fill(0x20, offset, offset + length)
        const payload = Buffer.from(controls.shortRead.payload)
        payload.copy(bytes, offset)
        return payload.length
      },
    ),
  }
})

import { ProjectStore } from '../src/store.js'

let tmpDir: string
let store: ProjectStore

beforeEach(() => {
  tmpDir = mkdtempSync(join(tmpdir(), 'project-store-read-'))
  store = new ProjectStore(tmpDir)
  store.ensureDefaultProject()
  controls.sourceReadFailures.clear()
  controls.shortRead.active = false
  controls.shortRead.calls = 0
  controls.shortRead.payload = ''
})

afterEach(() => {
  controls.sourceReadFailures.clear()
  controls.shortRead.active = false
  rmSync(tmpDir, { recursive: true, force: true })
})

describe('chat read failures', () => {
  it('keeps the source file when a merge cannot read its history', async () => {
    const chatsDir = join(tmpDir, 'projects', 'default', 'chats')
    mkdirSync(chatsDir, { recursive: true })
    store.appendChatMessage('default', 'target-chat', { role: 'assistant', text: 'existing' })
    const sourcePath = join(chatsDir, 'unsaved-read-failure.jsonl')
    const source = `${JSON.stringify({ seq: 0, ts: '', role: 'user', text: 'keep me' })}\n`
    writeFileSync(sourcePath, source, 'utf8')
    controls.sourceReadFailures.add(sourcePath)
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    try {
      store.rebindChat('default', 'unsaved-read-failure', 'target-chat')
    } finally {
      warn.mockRestore()
    }

    controls.sourceReadFailures.clear()
    const actual = await vi.importActual<typeof import('node:fs')>('node:fs')
    expect(existsSync(sourcePath)).toBe(true)
    expect(actual.readFileSync(sourcePath, 'utf8')).toBe(source)
    expect(store.loadChat('default', 'target-chat').map((message) => message.text)).toEqual([
      'existing',
    ])
    store.appendChatMessage('default', 'target-chat', { role: 'user', text: 'after failure' })
    expect(store.loadChat('default', 'target-chat').map((message) => message.seq)).toEqual([0, 1])
  })

  it('does not parse a partial chunk after readSync reaches EOF', () => {
    const chatsDir = join(tmpDir, 'projects', 'default', 'chats')
    mkdirSync(chatsDir, { recursive: true })
    const chatPath = join(chatsDir, 'short-read.jsonl')
    const payload = JSON.stringify({ seq: 0, ts: '', role: 'user', text: 'partial' })
    writeFileSync(chatPath, payload + ' '.repeat(300), 'utf8')
    controls.shortRead.active = true
    controls.shortRead.payload = payload

    expect(store.loadChat('default', 'short-read', 1)).toEqual([])
    expect(controls.shortRead.calls).toBe(2)
  })
})
