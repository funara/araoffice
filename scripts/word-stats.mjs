#!/usr/bin/env node
/**
 * Word's own statistics for a document — the fidelity gold standard, read from Word
 * itself rather than inferred from a PDF.
 *
 * This is the Windows counterpart to the AppleScript layer in
 * `scripts/docs-word-fidelity.mjs` (`--engine word`), which only runs on macOS. On
 * Windows the equivalent is Word's COM automation; this script drives it read-only
 * against a temporary copy, so the original file is never opened for writing and the
 * user's recent-files list is left alone.
 *
 * Usage:
 *   node scripts/word-stats.mjs <file.docx> [--json] [--pages] [--stories] [--out <file>]
 *
 * Reports the same figures Word's Review ▸ Word Count dialog shows, after a forced
 * repagination:
 *   pages, words, characters, paragraphs, lines
 *
 * Extra capture layers for pagination-parity work:
 *   --pages    per-page first/last line text (same shape as the corpus baselines:
 *              `pageStarts` / `pageEnds`), so our engine's page boundaries can be
 *              diffed against Word's page by page
 *   --stories  per-story word counts (main text, text frames, footnotes, endnotes,
 *              headers/footers) — attributes exactly which text sources make up the
 *              document total, so a word-count gap can be pinned to a story
 *   --out      write the JSON report to a file (in addition to stdout)
 *
 * Why it matters: a page count from another renderer (LibreOffice, a PDF conversion,
 * an in-app model) is a proxy. Word's `Repaginate` + `ComputeStatistics(wdStatisticPages)`
 * is the number the document was authored against, so it is the only sound target for
 * pagination parity work.
 *
 * Notes:
 *   - Word must be installed. The script exits 2 with a message when it is not, so a
 *     caller can skip the check instead of failing.
 *   - If Word is already running, this attaches to that instance and never quits it
 *     (only the temporary copy is closed); a Word it started itself is quit again.
 *   - Text travels out of PowerShell base64-encoded on `key=value` lines: Windows
 *     PowerShell 5.1's ConvertTo-Json corrupts strings containing quotes and control
 *     characters (it emits them unescaped), so no PS-side JSON is trusted.
 *   - A 30 MB document takes a minute or two to open and paginate; the COM call is
 *     given a generous timeout and the copy is deleted afterwards in a `finally`.
 */
import { execFileSync } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, dirname, join, resolve } from 'node:path'
import { randomUUID } from 'node:crypto'

const WD_PAGES = 2
const WD_WORDS = 0
const WD_CHARACTERS = 3
const WD_CHARACTERS_WITH_SPACES = 5
const WD_PARAGRAPHS = 4
const WD_LINES = 1

const WD_GO_TO_PAGE = 1
const WD_GO_TO_ABSOLUTE = 1

/** WdStoryType → report key (only stories that carry countable text) */
const STORY_TYPES = [
  ['mainText', 1],
  ['footnotes', 2],
  ['endnotes', 3],
  ['comments', 4],
  ['textFrames', 5],
  ['evenPagesHeader', 6],
  ['primaryHeader', 7],
  ['evenPagesFooter', 8],
  ['primaryFooter', 9],
  ['firstPageHeader', 10],
  ['firstPageFooter', 11],
]

function psScript(copyPath, opts) {
  const lines = [
    "$ErrorActionPreference = 'Stop'",
    '$wasRunning = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue).Count -gt 0',
    '$word = New-Object -ComObject Word.Application',
    '$word.Visible = $false',
    '$word.DisplayAlerts = 0',
    'try {',
    '  $doc = $word.Documents.Open(' + JSON.stringify(copyPath) + ', $false, $true, $false)',
    '  $doc.Repaginate()',
    '  $pages = $doc.ComputeStatistics(' + WD_PAGES + ')',
    '  $report = New-Object System.Collections.Generic.List[string]',
    "  $report.Add('pages=' + $pages)",
    "  $report.Add('words=' + $doc.ComputeStatistics(" + WD_WORDS + '))',
    "  $report.Add('characters=' + $doc.ComputeStatistics(" + WD_CHARACTERS + '))',
    "  $report.Add('charactersWithSpaces=' + $doc.ComputeStatistics(" +
      WD_CHARACTERS_WITH_SPACES +
      '))',
    "  $report.Add('paragraphs=' + $doc.ComputeStatistics(" + WD_PARAGRAPHS + '))',
    "  $report.Add('lines=' + $doc.ComputeStatistics(" + WD_LINES + '))',
  ]
  if (opts.stories) {
    lines.push(
      '  foreach ($entry in @(' +
        STORY_TYPES.map(([key, type]) => `@('${key}', ${type})`).join(', ') +
        ')) {',
      '    try {',
      '      $range = $doc.StoryRanges.Item($entry[1])',
      '      if ($null -ne $range) {',
      "        $report.Add('story.' + $entry[0] + '.words=' + $range.ComputeStatistics(" +
        WD_WORDS +
        '))',
      "        $report.Add('story.' + $entry[0] + '.charactersWithSpaces=' + $range.ComputeStatistics(" +
        WD_CHARACTERS_WITH_SPACES +
        '))',
      '      }',
      '    } catch { }',
      '  }',
    )
  }
  if (opts.pages) {
    lines.push(
      '  $contentEnd = $doc.Content.End',
      '  for ($i = 1; $i -le $pages; $i++) {',
      '    $start = $doc.GoTo(' + WD_GO_TO_PAGE + ', ' + WD_GO_TO_ABSOLUTE + ', $i).Start',
      '    $startText = $doc.Range($start, [Math]::Min($start + 120, $contentEnd)).Text',
      '    if ($null -eq $startText) { $startText = "" }',
      "    $report.Add('pageStart.' + $i + '=' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($startText)))",
      '    $next = $contentEnd',
      '    if ($i -lt $pages) { $next = $doc.GoTo(' +
        WD_GO_TO_PAGE +
        ', ' +
        WD_GO_TO_ABSOLUTE +
        ', ($i + 1)).Start }',
      '    $endText = $doc.Range([Math]::Max(0, $next - 120), $next).Text',
      '    if ($null -eq $endText) { $endText = "" }',
      "    $report.Add('pageEnd.' + $i + '=' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($endText)))",
      '  }',
    )
  }
  lines.push(
    '  $report -join [char]10',
    '  $doc.Close($false)',
    '} finally {',
    '  if (-not $wasRunning) { $word.Quit() }',
    '}',
  )
  return lines.join('\n')
}

/** C0/C1 control marks Word embeds in range text (cell \a, paragraph \r, soft break \v…) */
const CONTROL_RE = /\p{Cc}/gu

/** Word range text → a single clean line (strip control chars, cell/paragraph marks) */
function normalizeRangeText(text, mode) {
  const cleaned = String(text ?? '')
    .replace(CONTROL_RE, ' ')
    .replace(/\s+/g, ' ')
    .trim()
  if (mode === 'start') return cleaned.slice(0, 40)
  return cleaned.slice(-40)
}

function parseReport(raw) {
  const stats = { stories: {} }
  const starts = []
  const ends = []
  for (const line of raw.split(/\r?\n/)) {
    const eq = line.indexOf('=')
    if (eq <= 0) continue
    const key = line.slice(0, eq)
    const value = line.slice(eq + 1)
    const storyMatch = /^story\.([A-Za-z]+)\.(words|charactersWithSpaces)$/.exec(key)
    if (storyMatch) {
      const entry = (stats.stories[storyMatch[1]] ??= {})
      entry[storyMatch[2]] = Number(value)
      continue
    }
    const pageStart = /^pageStart\.(\d+)$/.exec(key)
    if (pageStart) {
      starts[Number(pageStart[1]) - 1] = Buffer.from(value, 'base64').toString('utf8')
      continue
    }
    const pageEnd = /^pageEnd\.(\d+)$/.exec(key)
    if (pageEnd) {
      ends[Number(pageEnd[1]) - 1] = Buffer.from(value, 'base64').toString('utf8')
      continue
    }
    if (/^(pages|words|characters|charactersWithSpaces|paragraphs|lines)$/.test(key)) {
      stats[key] = Number(value)
    }
  }
  if (starts.length > 0) stats.pageStarts = starts.map((t) => normalizeRangeText(t, 'start'))
  if (ends.length > 0) stats.pageEnds = ends.map((t) => normalizeRangeText(t, 'end'))
  return stats
}

function main() {
  const args = process.argv.slice(2)
  const json = args.includes('--json')
  const withPages = args.includes('--pages')
  const withStories = args.includes('--stories')
  const outIndex = args.indexOf('--out')
  const outFile = outIndex >= 0 && args[outIndex + 1] ? resolve(args[outIndex + 1]) : null
  const target = args.find((a, i) => !a.startsWith('--') && args[i - 1] !== '--out')
  if (!target) {
    console.error(
      'usage: node scripts/word-stats.mjs <file.docx> [--json] [--pages] [--stories] [--out <file>]',
    )
    process.exit(2)
  }
  const source = resolve(target)
  if (!existsSync(source)) {
    console.error('no such file: ' + source)
    process.exit(2)
  }
  if (process.platform !== 'win32') {
    console.error(
      'word-stats.mjs drives Word through COM and only runs on Windows.\n' +
        'On macOS use: node scripts/docs-word-fidelity.mjs --engine word',
    )
    process.exit(2)
  }

  // Word refuses files carrying the Mark-of-the-Web (it opens them in Protected View,
  // which COM cannot read), and it should never be pointed at the user's own file.
  const copy = join(tmpdir(), 'word-stats-' + randomUUID() + '.docx')
  copyFileSync(source, copy)
  try {
    execFileSync(
      'powershell',
      [
        '-NoProfile',
        '-NonInteractive',
        '-Command',
        'Unblock-File -LiteralPath ' + JSON.stringify(copy),
      ],
      { timeout: 60_000 },
    )
  } catch {
    /* Unblock-File is best effort; an unblocked source works without it */
  }

  try {
    const raw = execFileSync(
      'powershell',
      [
        '-NoProfile',
        '-NonInteractive',
        '-Command',
        psScript(copy, { pages: withPages, stories: withStories }),
      ],
      { timeout: 900_000, encoding: 'utf8', windowsHide: true },
    ).trim()
    const stats = parseReport(raw)
    if (!stats.pages) {
      const dump = (outFile ?? join(tmpdir(), 'word-stats-raw')) + '.raw.txt'
      mkdirSync(dirname(dump), { recursive: true })
      writeFileSync(dump, raw)
      console.error('Word produced no report (raw output saved to ' + dump + ')')
      process.exit(1)
    }
    const report = { file: basename(source), engine: 'Microsoft Word', ...stats }
    if (outFile) {
      mkdirSync(dirname(outFile), { recursive: true })
      writeFileSync(outFile, JSON.stringify(report, null, 2) + '\n')
    }
    if (json || outFile) {
      console.log(JSON.stringify(report, null, 2))
    } else {
      console.log(basename(source) + '  (Microsoft Word)')
      for (const [k, v] of Object.entries(stats)) {
        if (typeof v === 'object') continue
        console.log('  ' + k.padEnd(21) + v)
      }
    }
  } catch (err) {
    const text = String(err?.stderr ?? err?.message ?? err)
    if (/New-Object|COM|80040154|not registered/i.test(text)) {
      console.error('Word does not appear to be installed (COM could not start Word).')
      process.exit(2)
    }
    console.error('Word failed: ' + text.slice(0, 800))
    process.exit(1)
  } finally {
    rmSync(copy, { force: true })
  }
}

main()
