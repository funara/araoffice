/**
 * Personal-data spill gate.
 *
 * The repo rule is "no private identity in source" (see apps/shell/src/main/
 * service-config.ts). This gate enforces it mechanically so it cannot regress
 * the way it did before: identity metadata must come from env (service-config /
 * electron-builder) or tests that quote a name to assert its absence.
 *
 *   node tools/check-personal-spill.mjs        # fail on any hit outside allowlist
 *   ARA_SPILL_NAME=... (repeatable, comma-sep) extra deny patterns
 *
 * Scope: code + prompt files under apps/, packages/, scripts/, tools/, e2e/.
 * Deliberately NOT scanned: .md docs at repo root, .env (gitignored), build
 * outputs (out/, release/, target/), node_modules.
 */

import { readdirSync, readFileSync, statSync } from 'node:fs'
import { extname, join, relative, sep } from 'node:path'

// Assembled indirectly so this file never contains its own trigger strings.
const BASE_PATTERNS = [
  'adhi' + 'rahmadian', // full personal name
  'fun' + 'ara', // org handle
  'adhir' + '\\.my' + '\\.id', // personal domain (suffix form my.id)
  '@adhi' + 'r', // telegram-style handle
]
const envExtras = (process.env.ARA_SPILL_NAME ?? '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean)
const PATTERNS = [RegExp(`(${[...BASE_PATTERNS, ...envExtras].join('|')})`, 'gi')]

const ALLOWED_FILE_SNIPPETS = [
  'identity.test',
  'omnibar-service.test',
  'service-config',
  'check-personal-spill.mjs',
]

const SCAN_ROOTS = ['apps', 'packages', 'scripts', 'tools', 'e2e']
// .json is included: workspace manifests are packaged into the shipped app
// (electron-builder copies them into the bundle and the Windows PE resource
// headers), so a personal name in "author" would ship in the binary.
const CODE_EXTS = new Set(['.ts', '.tsx', '.rs', '.js', '.mjs', '.cjs', '.css', '.html', '.json'])
const SKIP_DIRS = new Set([
  'node_modules',
  'out',
  'release',
  'target',
  'dist',
  'fixture',
  'fixtures',
  'i18n', // locale string tables are user-facing copy, identity gates cover them separately
])

function* walk(dir) {
  for (const name of readdirSync(dir)) {
    if (SKIP_DIRS.has(name)) continue
    const full = join(dir, name)
    let st
    try {
      st = statSync(full)
    } catch {
      continue
    }
    if (st.isDirectory()) yield* walk(full)
    else if (
      CODE_EXTS.has(extname(name)) ||
      (extname(name) === '.md' && full.includes('prompts'))
    ) {
      yield full
    }
  }
}

const hits = []
for (const root of SCAN_ROOTS) {
  for (const file of walk(root)) {
    if (ALLOWED_FILE_SNIPPETS.some((snippet) => file.replace(/\\/g, '/').includes(snippet))) {
      continue
    }
    let text
    try {
      text = readFileSync(file, 'utf8')
    } catch {
      continue
    }
    for (const re of PATTERNS) {
      re.lastIndex = 0
      let m
      while ((m = re.exec(text)) !== null) {
        const line = text.slice(0, m.index).split('\n').length
        hits.push(`${relative('.', file).split(sep).join('/')}:${line}: ${m[0]}`)
      }
    }
  }
}

if (hits.length > 0) {
  console.error(`personal-spill gate failed (${hits.length}):\n${hits.join('\n')}`)
  console.error('Move private identity into env (ARA_BRAND_VENDOR / ARA_GITHUB_*).')
  process.exit(1)
}
console.log(
  `personal-spill gate passed (${BASE_PATTERNS.length + envExtras.length} patterns scanned)`,
)
