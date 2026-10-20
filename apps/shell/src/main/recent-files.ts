import { stat } from 'node:fs/promises'
import { basename, extname } from 'node:path'
import type { RecentEntry, RecentPage, RecentQuery } from '../shared/home-api'

const RECENT_PAGE_DEFAULT = 50
const RECENT_PAGE_MAX = 200
// ponytail: fixed chunk size bounds fd usage on huge recent lists; raise if profiling cares
const STAT_CHUNK = 8

const DIAGRAM_EXTENSIONS: Record<string, true> = {
  drawio: true,
  dio: true,
  mmd: true,
  mermaid: true,
}

async function toRecentEntry(
  path: string,
  starredPaths: ReadonlySet<string>,
): Promise<RecentEntry | null> {
  try {
    const info = await stat(path)
    return {
      path,
      name: basename(path),
      ext: extname(path).slice(1).toLowerCase(),
      mtimeMs: info.mtimeMs,
      sizeBytes: info.size,
      starred: starredPaths.has(path),
    }
  } catch {
    return null
  }
}

export async function statExistingPaths(
  paths: readonly string[],
  starredPaths: ReadonlySet<string>,
): Promise<RecentEntry[]> {
  const entries: RecentEntry[] = []
  for (let i = 0; i < paths.length; i += STAT_CHUNK) {
    const chunk = await Promise.all(
      paths.slice(i, i + STAT_CHUNK).map((path) => toRecentEntry(path, starredPaths)),
    )
    for (const entry of chunk) if (entry !== null) entries.push(entry)
  }
  return entries
}

export function normalizeRecentQuery(
  raw: unknown,
): Required<Omit<RecentQuery, 'ext'>> & { ext?: string } {
  const query = (raw ?? {}) as RecentQuery
  const offset = Number.isFinite(query.offset) ? Math.max(0, Math.floor(query.offset!)) : 0
  const limit = Number.isFinite(query.limit)
    ? Math.min(RECENT_PAGE_MAX, Math.max(0, Math.floor(query.limit!)))
    : RECENT_PAGE_DEFAULT
  const ext = typeof query.ext === 'string' && query.ext ? query.ext.toLowerCase() : undefined
  return { offset, limit, ext }
}

export function matchesRecentFilter(entry: RecentEntry, ext?: string): boolean {
  return !ext || entry.ext === ext || (ext === 'diagram' && DIAGRAM_EXTENSIONS[entry.ext] === true)
}

/** Page over existing paths only, preserving the source's newest-first order. */
export async function pageRecentPaths(
  paths: readonly string[],
  raw: unknown,
  starredPaths: ReadonlySet<string>,
): Promise<RecentPage> {
  const { offset, limit, ext } = normalizeRecentQuery(raw)
  const all = await statExistingPaths(paths, starredPaths)
  const filtered = all.filter((entry) => matchesRecentFilter(entry, ext))
  return {
    entries: limit === 0 ? [] : filtered.slice(offset, offset + limit),
    total: filtered.length,
    totalAll: all.length,
  }
}
