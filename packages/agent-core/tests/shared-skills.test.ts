import { describe, expect, it } from 'vitest'
import {
  ARA_ASSISTANT_IDENTITY,
  ARA_ASSISTANT_SHARED_POLICY,
  buildAppSystemPrompt,
  CONTINUE_INSTRUCTION,
  FILES_PROMPT,
  SEARCH_PROMPT,
  READ_ATTACHMENT_TOOL,
  WEB_SEARCH_TOOLS,
  formatAttachmentList,
  formatSize,
  attachmentChunkHeader,
  formatSearchResults,
  searchResultLinks,
} from '../src'

describe('shared prompt texts', () => {
  it('pins the canonical continuation instruction', () => {
    expect(CONTINUE_INSTRUCTION).toBe(
      'Continue the current task using the existing conversation and document state. Finish only the outstanding work. Do not repeat edits or steps that are already complete.',
    )
  })

  it('builds system = identity + policy + section, in order', () => {
    expect(buildAppSystemPrompt('HELLO')).toBe(
      `${ARA_ASSISTANT_IDENTITY}\n\n${ARA_ASSISTANT_SHARED_POLICY}\n\nHELLO`,
    )
  })

  it('files prompt names the attachment contract', () => {
    expect(FILES_PROMPT).toContain('read_attachment')
    expect(FILES_PROMPT).toContain('attachment list')
  })

  it('search prompt names both web tools', () => {
    expect(SEARCH_PROMPT).toContain('web_search')
    expect(SEARCH_PROMPT).toContain('web_fetch')
  })

  it('shared tool shapes carry the canonical tool names', () => {
    expect(READ_ATTACHMENT_TOOL.name).toBe('read_attachment')
    expect(WEB_SEARCH_TOOLS.map((t) => t.name)).toEqual(['web_search', 'web_fetch'])
  })

  it('formats attachment lists, sizes, chunk headers, and search results', () => {
    expect(formatAttachmentList([])).toBe('')
    expect(formatAttachmentList([{ name: 'a.md', ext: 'md', sizeBytes: 2048 }])).toContain('2KB')
    expect(formatSize(512)).toBe('512B')
    expect(attachmentChunkHeader('a.md', 100, 0, 50)).toContain('continue with offset=50')
    expect(attachmentChunkHeader('a.md', 50, 0, 50)).toContain('end of file')
    expect(formatSearchResults([{ title: 't', url: 'u', snippet: 's' }])).toContain('1. t')
    expect(formatSearchResults([])).toBe('(no results)')
    expect(searchResultLinks([{ title: 't', url: 'u' }])).toEqual([{ url: 'u', title: 't' }])
  })
})
