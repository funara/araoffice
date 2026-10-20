import { lookup } from 'node:dns/promises'
import { BlockList, isIP } from 'node:net'
import type { WebSearchResult } from './shared'
export type { WebSearchResult } from './shared'

export interface WebExtractResult {
  url: string
  title?: string
  content?: string
  excerpt?: string
  length?: number
  isPdf?: boolean
  pageCount?: number
  extractionQuality?: string
}

const DEFAULT_EXTRACT_CHARS = 8_000
const MAX_EXTRACT_CHARS = 50_000
const USER_AGENT =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36'

const blockedAddresses = (() => {
  const list = new BlockList()
  list.addSubnet('0.0.0.0', 8)
  list.addSubnet('10.0.0.0', 8)
  list.addSubnet('100.64.0.0', 10)
  list.addSubnet('127.0.0.0', 8)
  list.addSubnet('169.254.0.0', 16)
  list.addSubnet('172.16.0.0', 12)
  list.addSubnet('192.0.0.0', 24)
  list.addSubnet('192.168.0.0', 16)
  list.addSubnet('198.18.0.0', 15)
  list.addSubnet('224.0.0.0', 4)
  list.addSubnet('240.0.0.0', 4)
  list.addAddress('::', 'ipv6')
  list.addAddress('::1', 'ipv6')
  list.addSubnet('::', 96, 'ipv6')
  list.addSubnet('64:ff9b::', 96, 'ipv6')
  list.addSubnet('fc00::', 7, 'ipv6')
  list.addSubnet('fe80::', 10, 'ipv6')
  list.addSubnet('ff00::', 8, 'ipv6')
  return list
})()

const BLOCKED_HOST_SUFFIXES: readonly string[] = ['.localhost', '.local', '.internal']

export function isBlockedAddress(ip: string): boolean {
  const type = isIP(ip)
  return type === 0 || blockedAddresses.check(ip, type === 4 ? 'ipv4' : 'ipv6')
}

export async function isSafeRemoteUrl(raw: unknown): Promise<boolean> {
  if (typeof raw !== 'string') return false
  let parsed: URL
  try {
    parsed = new URL(raw)
  } catch {
    return false
  }
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') return false
  const host = parsed.hostname.replace(/^\[|\]$/g, '').toLowerCase()
  if (host === '') return false
  if (isIP(host)) return !isBlockedAddress(host)
  if (host === 'localhost' || BLOCKED_HOST_SUFFIXES.some((s) => host.endsWith(s))) return false
  try {
    const addrs = await lookup(host, { all: true })
    return addrs.length > 0 && addrs.every((a) => !isBlockedAddress(a.address))
  } catch {
    return false
  }
}

export function decodeHtmlEntities(text: string): string {
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCharCode(Number(dec)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)))
}

function stripHtmlTags(html: string): string {
  return html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, ' ')
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, ' ')
    .replace(/<noscript\b[^<]*(?:(?!<\/noscript>)<[^<]*)*<\/noscript>/gi, ' ')
    .replace(/<svg\b[^<]*(?:(?!<\/svg>)<[^<]*)*<\/svg>/gi, ' ')
    .replace(/<nav\b[^<]*(?:(?!<\/nav>)<[^<]*)*<\/nav>/gi, ' ')
    .replace(/<header\b[^<]*(?:(?!<\/header>)<[^<]*)*<\/header>/gi, ' ')
    .replace(/<footer\b[^<]*(?:(?!<\/footer>)<[^<]*)*<\/footer>/gi, ' ')
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<(?:br|p|div|h[1-6]|li|tr|blockquote)[^>]*>/gi, '\n')
    .replace(/<[^>]+>/g, '')
}

export function unwrapBingUrl(rawUrl: string): string {
  try {
    const clean = decodeHtmlEntities(rawUrl)
    const uMatch = /[?&]u=a1([a-zA-Z0-9+/=_-]+)/.exec(clean)
    if (uMatch && uMatch[1]) {
      let b64 = uMatch[1].replace(/-/g, '+').replace(/_/g, '/')
      while (b64.length % 4 !== 0) b64 += '='
      const decoded = Buffer.from(b64, 'base64').toString('utf8')
      if (decoded.startsWith('http')) return decoded
    }
    return clean
  } catch {
    return rawUrl
  }
}

export function unwrapYahooUrl(rawUrl: string): string {
  try {
    const ruMatch = /\/RU=([^/]+)\//.exec(rawUrl)
    if (ruMatch && ruMatch[1]) {
      const decoded = decodeURIComponent(ruMatch[1])
      if (decoded.startsWith('http')) return decoded
    }
    return rawUrl
  } catch {
    return rawUrl
  }
}

export function parseBraveResults(html: string, limit: number): WebSearchResult[] {
  const results: WebSearchResult[] = []
  const blocks = html.split(/<div class="result-body\b[^"]*"[^>]*>/i)
  for (let i = 1; i < blocks.length && results.length < limit; i++) {
    const block = blocks[i] ?? ''
    const linkMatch =
      /<a[^>]*href="([^"]+)"[^>]*>[\s\S]*?<div class="title search-snippet-title[^"]*"[^>]*>([\s\S]*?)<\/div>/i.exec(
        block,
      ) || /<a[^>]*href="([^"]+)"[^>]*title="([^"]+)"/i.exec(block)
    const snippetMatch =
      /<div class="content[^"]*t-primary[^"]*"[^>]*>([\s\S]*?)<\/div>/i.exec(block) ||
      /<div class="generic-snippet[^"]*"[^>]*>([\s\S]*?)<\/div>/i.exec(block)
    if (linkMatch) {
      const url = linkMatch[1] ?? ''
      const rawTitle = linkMatch[2] || url
      const title = decodeHtmlEntities(rawTitle.replace(/<[^>]+>/g, '').trim())
      const snippet = snippetMatch
        ? decodeHtmlEntities(snippetMatch[1]?.replace(/<[^>]+>/g, '').trim() ?? '')
        : ''
      if (url.startsWith('http') && title) {
        results.push({ title, url, snippet })
      }
    }
  }
  return results
}

function isJunkSearchResult(title: string, url: string, query = ''): boolean {
  const q = query.toLowerCase()
  const t = title.toLowerCase()
  const u = url.toLowerCase()
  if (
    q.includes('windows') ||
    q.includes('microsoft') ||
    q.includes('invoice') ||
    q.includes('請求書')
  ) {
    return false
  }
  if (
    u.includes('support.microsoft.com') ||
    u.includes('answers.microsoft.com') ||
    u.includes('windowsphone.com') ||
    u.includes('search.yahoo.com') ||
    u.includes('bing.com/aclick') ||
    u === 'https://www.yahoo.com' ||
    u === 'https://www.yahoo.com/' ||
    t.includes('get help in windows') ||
    t.includes('microsoft support') ||
    t.includes('windows 11 forum') ||
    t.includes('ヘルプについて') ||
    t.includes('請求書') ||
    t.includes('合計請求書') ||
    t === 'images' ||
    t === 'videos' ||
    t.includes('trending searches')
  ) {
    return true
  }
  return false
}
function hasQueryRelevance(results: WebSearchResult[], query: string): boolean {
  const words = query
    .toLowerCase()
    .replace(/[^\w\s]/g, ' ')
    .split(/\s+/)
    .filter((w) => w.length > 2 && !['and', 'for', 'the', 'with', 'from', 'about'].includes(w))
  if (words.length === 0) return true
  return results.some((r) => {
    const text = (r.title + ' ' + r.snippet).toLowerCase()
    return words.some((w) => text.includes(w))
  })
}

export function parseBingResults(html: string, limit: number, query = ''): WebSearchResult[] {
  const results: WebSearchResult[] = []
  const blocks = html.split(/<li class="b_algo[^"]*"[^>]*>/)
  for (let i = 1; i < blocks.length && results.length < limit; i++) {
    const block = blocks[i] ?? ''
    const linkMatch = /<h2[^>]*><a[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a><\/h2>/i.exec(block)
    const snippetMatch =
      /<p[^>]*class="b_lineclamp[^"]*"[^>]*>([\s\S]*?)<\/p>/i.exec(block) ||
      /<div class="b_caption">[\s\S]*?<p[^>]*>([\s\S]*?)<\/p>/i.exec(block) ||
      /<p[^>]*>([\s\S]*?)<\/p>/i.exec(block)

    if (linkMatch) {
      const rawUrl = linkMatch[1] ?? ''
      const url = unwrapBingUrl(rawUrl)
      const title = decodeHtmlEntities(linkMatch[2]?.replace(/<[^>]+>/g, '').trim() ?? '')
      const snippet = snippetMatch
        ? decodeHtmlEntities(snippetMatch[1]?.replace(/<[^>]+>/g, '').trim() ?? '')
        : ''

      if (url.startsWith('http') && !isJunkSearchResult(title, url, query)) {
        results.push({ title: title || url, url, snippet })
      }
    }
  }
  return results
}

export function parseYahooResults(html: string, limit: number, query = ''): WebSearchResult[] {
  const results: WebSearchResult[] = []
  const blocks = html.split(/<div class="[^"]*\balgo\b[^"]*"[^>]*>/i)
  for (let i = 1; i < blocks.length && results.length < limit; i++) {
    const block = blocks[i] ?? ''
    const linkMatch =
      /<a[^>]*href="([^"]+)"[^>]*>[\s\S]*?<h3[^>]*>([\s\S]*?)<\/h3>[\s\S]*?<\/a>/i.exec(block) ||
      /<h3[^>]*>[\s\S]*?<a[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>[\s\S]*?<\/h3>/i.exec(block) ||
      /<a[^>]*href="([^"]+)"[^>]*>[\s\S]*?<h3[^>]*>([\s\S]*?)<\/h3>/i.exec(block) ||
      /<h3[^>]*><a[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a><\/h3>/i.exec(block)
    const snippetMatch =
      /<div class="compText[^"]*"[^>]*>([\s\S]*?)<\/div>/i.exec(block) ||
      /<p class="[^"]*fc-dustygray[^"]*"[^>]*>([\s\S]*?)<\/p>/i.exec(block)
    if (linkMatch) {
      const rawUrl = linkMatch[1] ?? ''
      const url = unwrapYahooUrl(rawUrl)
      const title = decodeHtmlEntities(linkMatch[2]?.replace(/<[^>]+>/g, '').trim() ?? '')
      const snippet = snippetMatch
        ? decodeHtmlEntities(snippetMatch[1]?.replace(/<[^>]+>/g, '').trim() ?? '')
        : ''
      if (
        url.startsWith('http') &&
        title &&
        !isJunkSearchResult(title, url, query) &&
        !results.some((r) => r.url === url)
      ) {
        results.push({ title, url, snippet })
      }
    }
  }
  return results
}
/**
 * Searches the web using lightweight HTTP requests with automatic credential-free
 * fallback: Brave Search -> Bing Search -> Yahoo Search -> Wikipedia.
 * Completely dependency-free and fast (~200ms).
 */
export async function webSearch(
  query: string,
  maxResults = 6,
): Promise<{ results: WebSearchResult[]; method: string; answer?: string }> {
  const cleanQuery = query.trim()
  if (!cleanQuery) return { results: [], method: 'web-search' }
  const limit = Number.isFinite(maxResults) ? Math.min(20, Math.max(1, Math.floor(maxResults))) : 6

  // 1. Primary Credential-Free: Brave Web Search (clean, fast, unblocked globally)
  try {
    const res = await fetch(`https://search.brave.com/search?q=${encodeURIComponent(cleanQuery)}`, {
      headers: {
        'User-Agent': USER_AGENT,
        Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9',
      },
      signal: AbortSignal.timeout(6000),
    })
    if (res.ok) {
      const html = await res.text()
      const results = parseBraveResults(html, limit)
      if (results.length > 0 && hasQueryRelevance(results, cleanQuery)) {
        return { results, method: 'web-search' }
      }
    }
  } catch {
    // Fallthrough to Bing
  }

  // 2. Secondary: Bing Web Search
  try {
    const res = await fetch(
      `https://www.bing.com/search?q=${encodeURIComponent(cleanQuery)}&setlang=en`,
      {
        headers: {
          'User-Agent': USER_AGENT,
          Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
          'Accept-Language': 'en-US,en;q=0.9',
        },
        signal: AbortSignal.timeout(6000),
      },
    )
    if (res.ok) {
      const html = await res.text()
      const results = parseBingResults(html, limit, cleanQuery)
      if (results.length > 0 && hasQueryRelevance(results, cleanQuery)) {
        return { results, method: 'web-search' }
      }
    }
  } catch {
    // Fallthrough to Yahoo
  }

  // 3. Tertiary: Yahoo Web Search
  try {
    const res = await fetch(`https://search.yahoo.com/search?p=${encodeURIComponent(cleanQuery)}`, {
      headers: {
        'User-Agent': USER_AGENT,
        Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9',
      },
      signal: AbortSignal.timeout(6000),
    })
    if (res.ok) {
      const html = await res.text()
      const results = parseYahooResults(html, limit, cleanQuery)
      if (results.length > 0 && hasQueryRelevance(results, cleanQuery)) {
        return { results, method: 'web-search' }
      }
    }
  } catch {
    // Fallthrough to Wikipedia
  }

  // 4. Wikipedia Search (full-text search first, prefix opensearch fallback)
  try {
    const wikiQueryUrl = `https://en.wikipedia.org/w/api.php?action=query&list=search&srsearch=${encodeURIComponent(cleanQuery)}&srlimit=${limit}&format=json`
    const res = await fetch(wikiQueryUrl, {
      headers: { 'User-Agent': 'ARAOffice/2.0.0 (https://araoffice.app)' },
      signal: AbortSignal.timeout(4000),
    })
    if (res.ok) {
      const data = (await res.json()) as {
        query?: {
          search?: Array<{ title?: string; snippet?: string }>
        }
      }
      const hits = data.query?.search || []
      const results: WebSearchResult[] = []
      for (const h of hits) {
        if (h.title) {
          results.push({
            title: decodeHtmlEntities(h.title),
            url: `https://en.wikipedia.org/wiki/${encodeURIComponent(h.title.replace(/ /g, '_'))}`,
            snippet: decodeHtmlEntities((h.snippet || '').replace(/<[^>]+>/g, '').trim()),
          })
        }
      }
      if (results.length > 0) return { results, method: 'web-search' }
    }
  } catch {
    // Fallthrough to OpenSearch
  }

  try {
    const wikiUrl = `https://en.wikipedia.org/w/api.php?action=opensearch&search=${encodeURIComponent(cleanQuery)}&limit=${limit}&namespace=0&format=json`
    const res = await fetch(wikiUrl, {
      headers: { 'User-Agent': 'ARAOffice/2.0.0 (https://araoffice.app)' },
      signal: AbortSignal.timeout(4000),
    })
    if (res.ok) {
      const data = (await res.json()) as [string, string[], string[], string[]]
      const titles = data[1] || []
      const snippets = data[2] || []
      const urls = data[3] || []
      const results: WebSearchResult[] = []
      for (let i = 0; i < titles.length && results.length < limit; i++) {
        const u = urls[i]
        const t = titles[i]
        if (u && t) {
          results.push({
            title: String(t),
            url: String(u),
            snippet: String(snippets[i] || t),
          })
        }
      }
      if (results.length > 0) {
        return { results, method: 'web-search' }
      }
    }
  } catch {
    // Return empty results gracefully
  }

  return { results: [], method: 'web-search' }
}

/**
 * Extracts readable plain text content from a web URL using native HTTP fetch.
 */
export async function webExtract(
  url: string,
  maxChars?: number,
  mode?: 'full' | 'abstract' | 'metadata',
): Promise<WebExtractResult> {
  const limit =
    maxChars === undefined
      ? DEFAULT_EXTRACT_CHARS
      : Math.min(MAX_EXTRACT_CHARS, Math.max(200, maxChars))

  try {
    let currentUrl = url
    let res: Response | null = null
    for (let hop = 0; hop <= 5; hop++) {
      if (!(await isSafeRemoteUrl(currentUrl))) {
        return { url, excerpt: 'Private or unsafe network hosts are not accessible' }
      }
      res = await fetch(currentUrl, {
        headers: {
          'User-Agent': USER_AGENT,
          Accept: 'text/html,application/xhtml+xml,text/plain;q=0.9',
        },
        signal: AbortSignal.timeout(10000),
        redirect: 'manual',
      })
      const status = typeof res.status === 'number' ? res.status : res.ok ? 200 : 500
      if (status < 300 || status >= 400) break
      const location = res.headers.get('location')
      if (!location) {
        return { url, excerpt: 'Redirect missing location header' }
      }
      try {
        currentUrl = new URL(location, currentUrl).toString()
      } catch {
        return { url, excerpt: `Invalid redirect location: ${location}` }
      }
    }
    if (!res) {
      return { url, excerpt: 'Too many redirects' }
    }
    if (!res.ok) {
      return { url, excerpt: `HTTP ${res.status} ${res.statusText}` }
    }

    const contentType = res.headers.get('content-type') || ''
    if (contentType.includes('application/pdf')) {
      return { url, isPdf: true, excerpt: '[PDF Document]' }
    }

    const html = await res.text()

    // Extract title
    const titleMatch = /<title[^>]*>([\s\S]*?)<\/title>/i.exec(html)
    const title = titleMatch
      ? decodeHtmlEntities(titleMatch[1]?.trim() ?? '').slice(0, 512)
      : undefined

    if (mode === 'metadata') {
      return {
        url,
        ...(title !== undefined ? { title } : {}),
        length: html.length,
      }
    }

    // Strip HTML to clean readable text
    const cleanText = decodeHtmlEntities(stripHtmlTags(html))
      .replace(/[ \t]+/g, ' ')
      .replace(/\n\s*\n\s*\n+/g, '\n\n')
      .trim()

    const content = cleanText.slice(0, limit)
    const excerpt = content.slice(0, 200)

    return {
      url,
      ...(title !== undefined ? { title } : {}),
      content,
      excerpt,
      length: content.length,
      extractionQuality: 'http-reader',
    }
  } catch (err) {
    return {
      url,
      excerpt: `Failed to fetch URL: ${err instanceof Error ? err.message : String(err)}`,
    }
  }
}
