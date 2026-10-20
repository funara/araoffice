/**
 * Public release + brand identity configured in the codebase.
 *
 * Core metadata defaults are defined directly in code and can be optionally
 * overridden via `process.env` (or a local .env file during development).
 *
 * Keys:
 * - ARA_GITHUB_OWNER / ARA_GITHUB_REPO: GitHub repository coordinates
 * - ARA_WEBSITE_URL / ARA_GITHUB_URL / ARA_INSTAGRAM_URL: Home social profile links
 * - ARA_BRAND_VENDOR: Vendor / author name
 */
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'

export const ARA_CONFIG_DEFAULTS = {
  vendor: 'Adhi Rahmadian',
  websiteUrl: 'https://adhirahmadian.id',
  githubUrl: 'https://github.com/funara/araoffice',
  instagramUrl: 'https://instagram.com/araoffice_ai',
  githubOwner: 'funara',
  githubRepo: 'araoffice',
} as const

/** Minimal .env parser: comments, blank lines, optional export/quotes. */
function parseEnvFile(text: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const raw of String(text ?? '').split(/\r?\n/)) {
    const line = raw.trim()
    if (!line || line.startsWith('#')) continue
    const body = line.startsWith('export ') ? line.slice(7).trim() : line
    const eq = body.indexOf('=')
    if (eq <= 0) continue
    const key = body.slice(0, eq).trim()
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) continue
    let value = body.slice(eq + 1).trim()
    const quote = value[0]
    if (quote === '"' || quote === "'") {
      if (value.length >= 2 && value.endsWith(quote)) {
        value = value.slice(1, -1)
        if (quote === '"') value = value.replace(/\\n/g, '\n')
      }
    } else {
      const hash = value.indexOf(' #')
      if (hash >= 0) value = value.slice(0, hash).trim()
    }
    if (value !== '') out[key] = value
  }
  return out
}

function envFileCandidates(): string[] {
  const resourcesPath = (process as NodeJS.Process & { resourcesPath?: string }).resourcesPath
  return [
    join(process.cwd(), '.env'),
    ...(resourcesPath ? [join(resourcesPath, '.env')] : []),
    join(dirname(process.execPath), '.env'),
  ]
}

let fileCache: Record<string, string> | null = null

function readEnvFile(): Record<string, string> {
  if (fileCache) return fileCache
  for (const candidate of envFileCandidates()) {
    try {
      if (existsSync(candidate)) {
        fileCache = parseEnvFile(readFileSync(candidate, 'utf8'))
        return fileCache
      }
    } catch {
      // unreadable candidate: try the next one
    }
  }
  fileCache = {}
  return fileCache
}

function lookup(key: string, fallback: string): string {
  const fromEnv = process.env[key]
  if (fromEnv !== undefined) return fromEnv.trim()
  const fromFile = readEnvFile()[key]
  if (fromFile !== undefined) return fromFile.trim()
  return fallback
}

/** Display name used in `author` / `supportDesc` copy. */
export function araBrandVendor(): string {
  return lookup('ARA_BRAND_VENDOR', ARA_CONFIG_DEFAULTS.vendor)
}

/** Public profile links shown on the home sidebar. */
export function araWebsiteUrl(): string {
  return lookup('ARA_WEBSITE_URL', ARA_CONFIG_DEFAULTS.websiteUrl)
}

export function araGithubUrl(): string {
  return lookup('ARA_GITHUB_URL', ARA_CONFIG_DEFAULTS.githubUrl)
}

export function araInstagramUrl(): string {
  return lookup('ARA_INSTAGRAM_URL', ARA_CONFIG_DEFAULTS.instagramUrl)
}

export function araGithubOwner(): string {
  return lookup('ARA_GITHUB_OWNER', ARA_CONFIG_DEFAULTS.githubOwner)
}

export function araGithubRepo(): string {
  return lookup('ARA_GITHUB_REPO', ARA_CONFIG_DEFAULTS.githubRepo)
}

/** Drop the cached .env read so the next lookup re-reads (tests only). */
export function resetServiceConfigForTests(): void {
  fileCache = null
}
