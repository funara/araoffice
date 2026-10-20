#!/usr/bin/env node
/**
 * Run the shell (dev server or built app) with a clean Electron environment.
 *
 * ELECTRON_RUN_AS_NODE turns Electron into a plain Node runtime: require('electron')
 * then resolves to the executable path string instead of the API, and the main
 * process dies on the first ipcMain access. Tools that are themselves Electron
 * apps (editors, terminals, launchers) export it into the environment, where it
 * is inherited by everything they spawn - including this script.
 *
 *   node scripts/dev.cjs dev     electron-vite dev --watch
 *   node scripts/dev.cjs start   electron-vite build && electron .
 */

const { spawn } = require('node:child_process')
const { existsSync } = require('node:fs')
const { createRequire } = require('node:module')
const { join } = require('node:path')

const env = { ...process.env }
delete env.ELECTRON_RUN_AS_NODE

const packageDir = join(__dirname, '..')
const isWindows = process.platform === 'win32'
const suffix = isWindows ? '.cmd' : ''

/**
 * Locate a package's bin shim. This workspace hoists to the repo root, but
 * nested installs are possible, and npm sets INIT_CWD to the *invoking*
 * directory - which is the repo root when run as `npm run dev -w @araoffice/shell`.
 */
function resolveBin(packageName) {
  const name = packageName.replace(/^@[^/]+\//, '')
  const candidates = [
    join(packageDir, 'node_modules', '.bin', name + suffix),
    join(packageDir, '..', '..', 'node_modules', '.bin', name + suffix),
  ]
  for (const candidate of candidates) {
    if (existsSync(candidate)) return candidate
  }
  try {
    const require_ = createRequire(join(packageDir, 'package.json'))
    const manifest = require_(packageName + '/package.json')
    const bin = typeof manifest.bin === 'string' ? manifest.bin : manifest.bin && manifest.bin[name]
    if (bin) {
      const resolved = join(packageDir, 'node_modules', packageName, bin)
      if (existsSync(resolved)) return resolved
    }
  } catch {
    /* fall through to the spawn error below */
  }
  return candidates[0]
}

const mode = process.argv[2] === 'start' ? 'start' : 'dev'
const electronVite = resolveBin('electron-vite')
const electron = resolveBin('electron')
const steps =
  mode === 'start'
    ? [
        [electronVite, ['build']],
        [electron, ['.']],
      ]
    : [[electronVite, ['dev', '--watch']]]

function run(index) {
  if (index >= steps.length) return
  const [command, args] = steps[index]
  const child = spawn(command, args, { stdio: 'inherit', env, shell: isWindows })
  child.on('error', (error) => {
    console.error(`failed to start ${command}: ${error.message}`)
    process.exitCode = 1
  })
  for (const signal of ['SIGINT', 'SIGTERM']) {
    process.on(signal, () => child.kill(signal))
  }
  child.on('exit', (code, signal) => {
    if (code === 0 && index + 1 < steps.length) return run(index + 1)
    process.exitCode = code ?? (signal ? 1 : 0)
  })
}

run(0)
