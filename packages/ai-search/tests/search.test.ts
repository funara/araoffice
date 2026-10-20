import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  decodeHtmlEntities,
  parseBingResults,
  parseBraveResults,
  parseYahooResults,
  unwrapBingUrl,
  unwrapYahooUrl,
  webExtract,
  webSearch,
} from '../src/index'

describe('Pure HTTP webSearch', () => {
  const originalFetch = globalThis.fetch

  beforeEach(() => {
    vi.restoreAllMocks()
  })

  afterEach(() => {
    globalThis.fetch = originalFetch
  })

  it('decodes HTML entities properly', () => {
    expect(decodeHtmlEntities('&amp; &lt; &gt; &quot; &#39; &nbsp; &#65; &#x42;')).toBe(
      '& < > " \'   A B',
    )
  })

  it('unwraps Bing base64 redirect URLs', () => {
    const raw = 'https://www.bing.com/ck/a?!&&u=a1aHR0cHM6Ly93d3cuZGVlcHNlZWsuY29tLw&ntb=1'
    expect(unwrapBingUrl(raw)).toBe('https://www.deepseek.com/')
    expect(unwrapBingUrl('https://example.com/direct')).toBe('https://example.com/direct')
  })

  it('unwraps Yahoo encoded redirect URLs', () => {
    const raw =
      'https://r.search.yahoo.com/_ylt=A2/RV=2/RE=123/RO=10/RU=https%3a%2f%2fwww.reuters.com%2ftechnology%2f/RK=2/RS=xyz'
    expect(unwrapYahooUrl(raw)).toBe('https://www.reuters.com/technology/')
    expect(unwrapYahooUrl('https://example.com/direct')).toBe('https://example.com/direct')
  })

  it('parses Brave HTML search results', () => {
    const braveHtml = `
      <div class="result-body svelte-1rq4ngz">
        <div class="result-content">
          <a href="https://www.reuters.com/technology/deepseek/" target="_self">
            <div class="title search-snippet-title">Reuters DeepSeek News</div>
          </a>
          <div class="content t-primary">DeepSeek price cut permanent.</div>
        </div>
      </div>
    `
    const parsed = parseBraveResults(braveHtml, 5)
    expect(parsed).toHaveLength(1)
    expect(parsed[0]).toEqual({
      title: 'Reuters DeepSeek News',
      url: 'https://www.reuters.com/technology/deepseek/',
      snippet: 'DeepSeek price cut permanent.',
    })
  })

  it('parses Yahoo HTML search results', () => {
    const yahooHtml = `
      <div class="dd algo algo-sr">
        <div class="compTitle">
          <a href="https://r.search.yahoo.com/RU=https%3a%2f%2fwww.reuters.com%2f/RK=2/">
            <h3 class="title">Reuters News</h3>
          </a>
        </div>
        <div class="compText"><p>Latest breaking world news.</p></div>
      </div>
    `
    const parsed = parseYahooResults(yahooHtml, 5)
    expect(parsed).toHaveLength(1)
    expect(parsed[0]).toEqual({
      title: 'Reuters News',
      url: 'https://www.reuters.com/',
      snippet: 'Latest breaking world news.',
    })
  })

  it('parses Bing HTML search results and unwraps redirect URLs', async () => {
    const mockHtml = `
      <ol id="b_results">
        <li class="b_algo">
          <h2><a href="https://www.bing.com/ck/a?!&amp;&amp;u=a1aHR0cHM6Ly93d3cuZGVlcHNlZWsuY29tLw&amp;ntb=1">DeepSeek | Into the Unknown</a></h2>
          <div class="b_caption"><p>DeepSeek is an AI research company focused on building world-leading general artificial intelligence.</p></div>
        </li>
      </ol>
    `
    const directParsed = parseBingResults(mockHtml, 5)
    expect(directParsed).toHaveLength(1)
    expect(directParsed[0].title).toBe('DeepSeek | Into the Unknown')
    globalThis.fetch = vi.fn().mockImplementation(async (url: string) => {
      if (String(url).includes('bing.com')) {
        return {
          ok: true,
          text: async () => mockHtml,
        }
      }
      throw new Error('Not Bing')
    })

    const res = await webSearch('DeepSeek', 5)
    expect(res.method).toBe('web-search')
    expect(res.results).toHaveLength(1)
    expect(res.results[0]).toEqual({
      title: 'DeepSeek | Into the Unknown',
      url: 'https://www.deepseek.com/',
      snippet:
        'DeepSeek is an AI research company focused on building world-leading general artificial intelligence.',
    })
  })

  it('falls back gracefully to Wikipedia API when search engines fail', async () => {
    globalThis.fetch = vi.fn().mockImplementation(async (url: string) => {
      if (String(url).includes('wikipedia.org')) {
        return {
          ok: true,
          json: async () => [
            'ARAOffice',
            ['ARAOffice'],
            ['AI-native office suite'],
            ['https://en.wikipedia.org/wiki/ARAOffice'],
          ],
        }
      }
      throw new Error('Search engine offline')
    })

    const res = await webSearch('ARAOffice', 5)
    expect(res.method).toBe('web-search')
    expect(res.results).toHaveLength(1)
    expect(res.results[0].title).toBe('ARAOffice')
    expect(res.results[0].url).toBe('https://en.wikipedia.org/wiki/ARAOffice')
  })

  it('returns empty results on complete network failure without throwing', async () => {
    globalThis.fetch = vi.fn().mockRejectedValue(new Error('Offline'))

    const res = await webSearch('test query')
    expect(res.method).toBe('web-search')
    expect(res.results).toEqual([])
  })
})

describe('Pure HTTP webExtract', () => {
  const originalFetch = globalThis.fetch

  beforeEach(() => {
    vi.restoreAllMocks()
  })

  afterEach(() => {
    globalThis.fetch = originalFetch
  })

  it.each([
    'http://localhost/admin',
    'http://127.0.0.1:8080/test',
    'http://10.0.0.1/private',
    'http://192.168.1.1/router',
    'http://172.16.0.1/internal',
    'ftp://example.com/file',
  ])('rejects unsafe URL %s', async (url) => {
    const res = await webExtract(url)
    expect(res.content).toBeUndefined()
    expect(res.excerpt).toBeDefined()
  })

  it('extracts and cleans readable HTML text', async () => {
    const mockHtml = `
      <!DOCTYPE html>
      <html>
        <head>
          <title>Example Domain &amp; Page</title>
          <style>body color: red; </style>
          <script>console.log('secret')</script>
        </head>
        <body>
          <nav><a href="/">Home</a></nav>
          <h1>Welcome to Example</h1>
          <p>This is a <strong>test</strong> paragraph with details.</p>
          <footer>Copyright 2026</footer>
        </body>
      </html>
    `
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      headers: new Headers({ 'content-type': 'text/html' }),
      text: async () => mockHtml,
    })

    const res = await webExtract('https://example.com/page', 500)
    expect(res.title).toBe('Example Domain & Page')
    expect(res.content).toContain('Welcome to Example')
    expect(res.content).toContain('This is a test paragraph with details.')
    expect(res.content).not.toContain('console.log')
    expect(res.content).not.toContain('body { color: red; }')
    expect(res.content).not.toContain('Copyright 2026')
  })
})
