import { existsSync } from 'node:fs'
import { delimiter, join } from 'node:path'
import { spawnSync } from 'node:child_process'

const args = process.argv.slice(2)
const mode = args.shift()

if (mode !== '--check' && mode !== '--write') {
  console.error('Usage: node tools/format-changed.mjs <--check|--write> [--base <git-ref>]')
  process.exit(2)
}

let baseRef = process.env.FORMAT_BASE_REF || ''
let baseFromArg = false
if (args[0] === '--base') {
  args.shift()
  baseRef = args.shift() || ''
  baseFromArg = true
}
// The all-zero id means "no previous commit" (the first push of a branch or
// repo); there is nothing to diff against, so only working-tree changes are checked.
if (/^0+$/.test(baseRef)) baseRef = ''
if (args.length > 0) {
  console.error(`Unexpected argument: ${args[0]}`)
  process.exit(2)
}

function git(commandArgs, { allowFailure = false } = {}) {
  const result = spawnSync('git', commandArgs, {
    encoding: 'buffer',
    stdio: ['ignore', 'pipe', 'inherit'],
  })
  if (result.status !== 0 && !allowFailure) {
    process.exit(result.status ?? 1)
  }
  return result
}

const repoRoot = git(['rev-parse', '--show-toplevel']).stdout.toString().trim()
const files = new Set()

function addNullSeparated(output) {
  for (const file of output.toString().split('\0')) {
    if (file && existsSync(join(repoRoot, file))) files.add(file)
  }
}

if (baseRef) {
  const verification = git(['rev-parse', '--verify', `${baseRef}^{commit}`], {
    allowFailure: true,
  })
  if (verification.status !== 0) {
    // An explicit --base must resolve; a CI-provided ref may legitimately be
    // gone (force-pushed/rewritten history), where no diff base exists and
    // only working-tree changes can be checked.
    if (baseFromArg) {
      console.error(`Formatting base is not a commit available in this checkout: ${baseRef}`)
      process.exit(2)
    }
    console.warn(
      `Formatting base ${baseRef} is not available in this checkout (rewritten history?); ` +
        'checking working-tree changes only.',
    )
  } else {
    addNullSeparated(
      git(['diff', '--name-only', '--diff-filter=ACMRT', '-z', `${baseRef}...HEAD`]).stdout,
    )
  }
}

addNullSeparated(git(['diff', '--name-only', '--diff-filter=ACMRT', '-z']).stdout)
addNullSeparated(git(['diff', '--cached', '--name-only', '--diff-filter=ACMRT', '-z']).stdout)
addNullSeparated(git(['ls-files', '--others', '--exclude-standard', '-z']).stdout)

const changedFiles = [...files].sort()
if (changedFiles.length === 0) {
  console.log('No changed files to format.')
  process.exit(0)
}

const prettierExecutable = join(
  repoRoot,
  'node_modules',
  '.bin',
  process.platform === 'win32' ? 'prettier.cmd' : 'prettier',
)
const prettierMode = mode === '--write' ? '--write' : '--check'
const prettierEnv = {
  ...process.env,
  PATH: `${join(repoRoot, 'node_modules', '.bin')}${delimiter}${process.env.PATH}`,
}

// Windows caps a command line at ~8k characters. A large working tree (or a
// first commit where every file is untracked) hands prettier thousands of
// paths at once and trips ENAMETOOLONG, so the file list is split into
// batches that stay well under the limit.
const COMMAND_LINE_BUDGET = 4000

/** @param {string[]} paths */
function batchPaths(paths) {
  const batches = []
  let batch = []
  let size = 0
  for (const path of paths) {
    const cost = path.length + 1
    if (batch.length > 0 && size + cost > COMMAND_LINE_BUDGET) {
      batches.push(batch)
      batch = []
      size = 0
    }
    batch.push(path)
    size += cost
  }
  if (batch.length > 0) batches.push(batch)
  return batches
}

let failed = false
for (const batch of batchPaths(changedFiles)) {
  const result = spawnSync(prettierExecutable, [prettierMode, '--ignore-unknown', '--', ...batch], {
    cwd: repoRoot,
    stdio: 'inherit',
    shell: process.platform === 'win32',
    env: prettierEnv,
  })

  if (result.error) {
    console.error(`Unable to run Prettier: ${result.error.message}`)
    process.exit(1)
  }
  if ((result.status ?? 1) !== 0) failed = true
}

process.exit(failed ? 1 : 0)
