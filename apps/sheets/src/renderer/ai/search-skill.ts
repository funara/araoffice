import type { AgentSkill } from '@araoffice/agent-core'
import {
  SEARCH_PROMPT,
  WEB_SEARCH_TOOLS,
  formatSearchResults,
  searchResultLinks,
} from '@araoffice/agent-core'
import { t } from '../i18n/locale'

/** Web-search AgentSkill: the main process returns titles, links, and snippets. */
export function createSearchSkill(): AgentSkill {
  return {
    id: 'search',
    systemPrompt: SEARCH_PROMPT,
    tools: [...WEB_SEARCH_TOOLS],
    executeTool: async (call) => {
      if (call.name !== 'web_search' && call.name !== 'web_fetch') {
        return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
      if (call.name === 'web_fetch') {
        const url = String(call.input.url ?? '').trim()
        if (!url) {
          return { output: 'url must not be empty', isError: true, summary: 'Web fetch' }
        }
        try {
          const r = await window.desktopApi.webFetch(url)
          if (r.isError) {
            return { output: 'web fetch failed', isError: true, summary: 'Web fetch' }
          }
          return {
            output: [
              `Title: ${r.title}`,
              `URL: ${r.url}`,
              '--- BEGIN UNTRUSTED WEB CONTENT ---',
              r.content,
              '--- END UNTRUSTED WEB CONTENT ---',
            ].join('\n'),
            mutated: false,
            summary: 'Web fetch',
          }
        } catch {
          return { output: 'web fetch failed', isError: true, summary: 'Web fetch' }
        }
      }
      const query = String(call.input.query ?? '').trim()
      if (!query) {
        return { output: 'query must not be empty', isError: true, summary: t('aiToolWebSearch') }
      }
      const r = await window.desktopApi.webSearch(query, Number(call.input.maxResults) || 6)
      if (r.method === 'error') {
        return {
          output: 'web search failed (service error, not an empty result — you may retry)',
          isError: true,
          summary: t('aiToolWebSearch'),
        }
      }
      return {
        output: formatSearchResults(r.results, r.answer),
        mutated: false,
        summary: t('aiToolWebSearchDone', { query, count: r.results.length }),
        display: {
          kind: 'links',
          items: searchResultLinks(r.results),
        },
      }
    },
  }
}
