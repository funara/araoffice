import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { readAppSettings, updateAppSetting, writeAppSetting } from '../src/main/app-settings'

/**
 * userData/app-settings.json helpers (src/main/app-settings.ts): a flat JSON
 * object shared by the language preference and the first-run onboarding flag.
 */

let dir: string
let settingsPath: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'app-settings-'))
  settingsPath = join(dir, 'app-settings.json')
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('readAppSettings', () => {
  it('returns an empty object when the file does not exist', async () => {
    expect(await readAppSettings(settingsPath)).toEqual({})
  })

  it('returns an empty object for invalid JSON', async () => {
    writeFileSync(settingsPath, 'not json')
    expect(await readAppSettings(settingsPath)).toEqual({})
  })

  it('returns an empty object when the JSON root is not an object', async () => {
    writeFileSync(settingsPath, '[1, 2]')
    expect(await readAppSettings(settingsPath)).toEqual({})
    writeFileSync(settingsPath, '"zh"')
    expect(await readAppSettings(settingsPath)).toEqual({})
  })

  it('parses a valid settings object', async () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'zh', onboardingSeen: true }))
    expect(await readAppSettings(settingsPath)).toEqual({ language: 'zh', onboardingSeen: true })
  })
})

describe('writeAppSetting', () => {
  it('creates the file with the single key on first write', async () => {
    await writeAppSetting(settingsPath, 'onboardingSeen', true)
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({ onboardingSeen: true })
  })

  it('preserves unrelated existing keys', async () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'ja' }))
    await writeAppSetting(settingsPath, 'onboardingSeen', true)
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({
      language: 'ja',
      onboardingSeen: true,
    })
  })

  it('overwrites the value of an existing key', async () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'ja' }))
    await writeAppSetting(settingsPath, 'language', 'en')
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({ language: 'en' })
  })

  it('recovers from a corrupt file by rewriting it', async () => {
    writeFileSync(settingsPath, '{broken')
    await writeAppSetting(settingsPath, 'language', 'en')
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({ language: 'en' })
  })

  it('updates a setting from its current value while preserving other keys', async () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'ja', windowState: { width: 100 } }))
    await updateAppSetting(settingsPath, 'windowState', (current) => ({
      ...(current as { width?: number }),
      height: 200,
    }))
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({
      language: 'ja',
      windowState: { width: 100, height: 200 },
    })
  })

  it('persists the onboardingSeen boolean flag', async () => {
    await writeAppSetting(settingsPath, 'onboardingSeen', true)
    expect(await readAppSettings(settingsPath)).toEqual({ onboardingSeen: true })
    await writeAppSetting(settingsPath, 'onboardingSeen', false)
    expect(await readAppSettings(settingsPath)).toEqual({ onboardingSeen: false })
  })
})
