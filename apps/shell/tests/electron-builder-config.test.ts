import { readFileSync } from 'node:fs'
import * as path from 'node:path'
import { fileURLToPath } from 'node:url'
import { runInNewContext } from 'node:vm'
import { describe, expect, it } from 'vitest'

const testDir = path.dirname(fileURLToPath(import.meta.url))
const configPath = path.resolve(testDir, '../electron-builder.cjs')
const configSource = readFileSync(configPath, 'utf8')

function loadConfig(env: Record<string, string> = {}) {
  const mockedFs = {
    existsSync: (filePath: string) => filePath.endsWith('LICENSES.chromium.html'),
    mkdirSync: () => undefined,
    readFileSync: () => {
      throw new Error('Unexpected file read in isolated packaging config test')
    },
    writeFileSync: () => undefined,
  }
  const module = { exports: {} as Record<string, unknown> }
  const sandbox = {
    __dirname: path.dirname(configPath),
    module,
    process: { env },
    require: (name: string) => {
      if (name === 'node:fs') return mockedFs
      if (name === 'node:path') return path
      if (name === 'electron/package.json') return { version: '43.3.0' }
      throw new Error(`Unexpected require: ${name}`)
    },
  }

  runInNewContext(configSource, sandbox, { filename: configPath })
  return JSON.parse(JSON.stringify(module.exports)) as { publish?: unknown[] }
}

describe('electron-builder release configuration', () => {
  it('defaults to the public GitHub Releases repository', () => {
    expect(loadConfig().publish).toEqual([
      { provider: 'github', owner: 'funara', repo: 'araoffice', releaseType: 'release' },
    ])
  })

  it('keeps the optional generic update feed override', () => {
    expect(loadConfig({ ARAOFFICE_UPDATE_URL: 'https://updates.example.test' }).publish).toEqual([
      { provider: 'generic', url: 'https://updates.example.test', channel: 'latest' },
    ])
  })
})
