import { describe, expect, it } from 'vitest'
import {
  appendFollowUp,
  executeOmnibarQuery,
  extractFollowUps,
  OMNIBAR_FIRST_CHUNK_TIMEOUT_MS,
  OMNIBAR_MAX_ATTEMPTS,
} from '../src/renderer/src/ai/omnibar-service'
import { buildOmnibarSystemPrompt } from '../src/renderer/src/ai/omnibar-prompt'

describe('omnibar-service', () => {
  describe('extractFollowUps', () => {
    it('returns empty list when no followups tag exists', () => {
      const raw = 'This is a normal response without followups.'
      const { cleanText, followUps } = extractFollowUps(raw)
      expect(cleanText).toBe(raw)
      expect(followUps).toEqual([])
    })

    it('extracts followups and cleans the body text', () => {
      const raw = `Here is the financial summary for Q3.

- Revenue is up 12%
- EBITDA margin expanded

:::followups
- Compare against Q2 numbers
- Break down operating expenses
- Generate presentation slide
:::
`
      const { cleanText, followUps } = extractFollowUps(raw)
      expect(cleanText).not.toContain(':::followups')
      expect(cleanText).toContain('Revenue is up 12%')
      expect(followUps).toEqual([
        'Compare against Q2 numbers',
        'Break down operating expenses',
        'Generate presentation slide',
      ])
    })

    it('handles bullet formats with asterisks and numbered lists', () => {
      const raw = `Done drafting the memo.
:::followups
* Check spelling
1. Add executive sign-off
- Export to Word
:::
`
      const { cleanText, followUps } = extractFollowUps(raw)
      expect(cleanText).toBe('Done drafting the memo.')
      expect(followUps).toEqual(['Check spelling', 'Add executive sign-off', 'Export to Word'])
    })

    it('strips template scaffolding from model-emitted options', () => {
      const raw = `Pick one:
:::followups
- [Daftar nama dan jabatan sudah tersedia]
- [Gunakan struktur standar DKM]
- [Saya akan mengirim daftar nama]:::
:::
`
      const { followUps } = extractFollowUps(raw)
      expect(followUps).toEqual([
        'Daftar nama dan jabatan sudah tersedia',
        'Gunakan struktur standar DKM',
        'Saya akan mengirim daftar nama',
      ])
    })

    it('filters out redundant app launcher actions from followups', () => {
      const raw = `Draft ready.
:::followups
- Open in ARA Docs & Write
- Help me design a 20-item survey
- Open in ARA Sheets & Write
- Turn this template into a journal article
:::
`
      const { followUps } = extractFollowUps(raw)
      expect(followUps).toEqual([
        'Help me design a 20-item survey',
        'Turn this template into a journal article',
      ])
    })

    it('handles empty input', () => {
      expect(extractFollowUps('')).toEqual({ cleanText: '', followUps: [] })
    })

    it('strips <think> blocks from streamed text before followup parsing', async () => {
      const seen: Array<{ text: string; isStreaming: boolean }> = []
      let done: { text: string } | null = null
      const fakeHandle = { cancel: () => {} }
      const g = globalThis as unknown as {
        window: {
          aiOfficeAi: {
            getSettings: () => Promise<unknown>
            stream: (
              req: unknown,
              cb: { onChunk: (c: { type: string; text?: string }) => void },
            ) => typeof fakeHandle
          }
        }
      }
      g.window = {
        aiOfficeAi: {
          getSettings: () => Promise.resolve({}),
          stream: (_req, cb) => {
            cb.onChunk({ type: 'thinking', text: 'planning the intro' })
            cb.onChunk({ type: 'delta', text: '<think>secret plan</think>Hello world' })
            cb.onChunk({ type: 'done' })
            return fakeHandle
          },
        },
      }
      try {
        await executeOmnibarQuery(
          'what model are you?',
          { id: 'home', kind: 'home', title: 'Home' },
          'en',
          {
            onUpdate: (r) => seen.push({ text: r.text, isStreaming: r.isStreaming }),
            onError: (e) => {
              throw new Error(e)
            },
            onDone: (r) => {
              done = { text: r.text }
            },
          },
        )
      } finally {
        delete (globalThis as unknown as Record<string, unknown>).window
      }
      expect(seen.length).toBe(1)
      expect(seen[0]?.text).toBe('Hello world')
      expect(done?.text).toBe('Hello world')
      expect('thinking' in (seen[0] ?? {})).toBe(false)
      expect('thinking' in (done ?? {})).toBe(false)
    })

    it('keeps stream alive and resets timer when thinking chunks arrive before answer text', async () => {
      let done: { text: string } | null = null
      const fakeHandle = { cancel: () => {} }
      const g = globalThis as unknown as {
        window: {
          aiOfficeAi: {
            getSettings: () => Promise<unknown>
            stream: (
              req: unknown,
              cb: { onChunk: (c: { type: string; text?: string }) => void },
            ) => typeof fakeHandle
          }
        }
      }
      g.window = {
        aiOfficeAi: {
          getSettings: () => Promise.resolve({}),
          stream: (_req, cb) => {
            cb.onChunk({ type: 'thinking', text: 'Step 1: Analyzing query...' })
            cb.onChunk({ type: 'thinking', text: 'Step 2: Synthesizing answer...' })
            cb.onChunk({ type: 'delta', text: 'Final Answer' })
            cb.onChunk({ type: 'done' })
            return fakeHandle
          },
        },
      }
      try {
        await executeOmnibarQuery(
          'can you solve this puzzle?',
          { id: 'home', kind: 'home', title: 'Home' },
          'en',
          {
            onUpdate: () => {},
            onError: (e) => {
              throw new Error(e)
            },
            onDone: (r) => {
              done = { text: r.text }
            },
          },
        )
      } finally {
        delete (globalThis as unknown as Record<string, unknown>).window
      }
      expect(done?.text).toBe('Final Answer')
    })
  })
  describe('appendFollowUp', () => {
    it('freezes in-flight answers instead of dropping them', () => {
      const thread = [
        { role: 'user' as const, text: 'buatkan outline' },
        { role: 'assistant' as const, text: 'partial answer...', isStreaming: true },
      ]
      const next = appendFollowUp(thread, 'lanjutkan')
      expect(next).toHaveLength(3)
      expect(next[1]).toMatchObject({ text: 'partial answer...', isStreaming: false })
      expect(next[2]).toMatchObject({ role: 'user', text: 'lanjutkan' })
    })
    it('drops empty streaming bubbles that carry nothing', () => {
      const thread = [{ role: 'assistant' as const, text: '', isStreaming: true }]
      expect(appendFollowUp(thread, 'halo')).toEqual([{ role: 'user', text: 'halo' }])
    })
  })

  describe('suggestion chips', () => {
    it('instructs the model not to emit suggestion chips or :::followups blocks', () => {
      const context: ActiveTabContext = {
        id: 't1',
        kind: 'docs',
        title: 'Report.docx',
      }
      const prompt = buildOmnibarSystemPrompt(context, 'en', 'Which topic do you want?')
      expect(prompt).toContain('Do NOT emit suggestion chips')
    })
  })

  describe('fallback ladder', () => {
    // One model, one 15s first-chunk budget, a single extra attempt. Which
    // backend actually serves the request is the transport's business.
    it('retries the same settings once with a 15s first-chunk budget', () => {
      expect(OMNIBAR_FIRST_CHUNK_TIMEOUT_MS).toBe(15_000)
      expect(OMNIBAR_MAX_ATTEMPTS).toBe(2)
    })
  })

  describe('omnibar-prompt', () => {
    it('builds system prompt with general conversation and no-chip instructions', () => {
      const context: ActiveTabContext = {
        id: 't1',
        kind: 'docs',
        title: 'Report.docx',
      }
      const prompt = buildOmnibarSystemPrompt(context, 'en', 'tolong ringkas dokumen ini')
      expect(prompt).toContain('ARA Assistant')
      expect(prompt).toContain('general conversation')
      expect(prompt).toContain('multi-turn conversation')
      expect(prompt).toContain('Do NOT emit suggestion chips')
      expect(prompt).toContain('Active Document: Word (.docx) - "Report.docx"')
    })

    it('instructs the model never to append closing call-to-action sentences', () => {
      const context: ActiveTabContext = { id: 't1', kind: 'docs', title: 'Report.docx' }
      const prompt = buildOmnibarSystemPrompt(context, 'en')
      expect(prompt).toContain('NEVER append closing call-to-action sentences')
    })

    it('keeps the handover dock contract out of chat', () => {
      const context: ActiveTabContext = { id: 'home', kind: 'home', title: 'Home' }
      const prompt = buildOmnibarSystemPrompt(context, 'en')
      expect(prompt).toContain('handover dock')
      expect(prompt).not.toContain('**Target Editor Prompt:**')
      expect(prompt).toContain("ARA Diagram's visual canvas (.diagram)")
    })

    it('runs ordinary turns on core identity without the defense/catalog dumps', () => {
      const context: ActiveTabContext = { id: 't1', kind: 'docs', title: 'Report.docx' }
      const prompt = buildOmnibarSystemPrompt(context, 'en', 'tolong ringkas dokumen ini')
      expect(prompt).not.toContain('Identity & Tone Guidelines')
      expect(prompt).not.toContain('Critical Format & Capabilities Invariants')
      expect(prompt).not.toContain('.diagram, .drawio')
      expect(prompt).toContain('name them only when asked')
    })

    it('answers who-are-you with rules but no catalog dump', () => {
      const context: ActiveTabContext = { id: 't1', kind: 'docs', title: 'Report.docx' }
      const prompt = buildOmnibarSystemPrompt(context, 'en', 'hi who are you?')
      expect(prompt).toContain('Answer only what was asked')
      expect(prompt).not.toContain('Critical Format & Capabilities Invariants')
      expect(prompt).not.toContain('.diagram, .drawio')
    })

    it('restores the full identity dump only for capability asks', () => {
      const context: ActiveTabContext = { id: 'home', kind: 'home', title: 'Home' }
      const prompt = buildOmnibarSystemPrompt(context, 'en', 'what can you do?')
      expect(prompt).toContain('Identity & Tone Guidelines')
      expect(prompt).toContain('Critical Format & Capabilities Invariants')
      expect(prompt).toContain('.diagram, .drawio')
    })

    it('builds neutral system prompt for custom AI provider without ARA identity rules or creator claims', () => {
      const context: ActiveTabContext = { id: 'home', kind: 'home', title: 'Home' }
      const prompt = buildOmnibarSystemPrompt(context, 'en', 'who are you?', 'custom')
      expect(prompt).toContain('the AI office copilot in ARAOffice')
      expect(prompt).not.toContain('Identity & Tone Guidelines')
      expect(prompt).not.toContain('Critical Format & Capabilities Invariants')
    })

    it('does not scrub third-party identity from custom AI responses', async () => {
      let done: { text: string } | null = null
      const fakeHandle = { cancel: () => {} }
      const g = globalThis as unknown as {
        window: {
          aiOfficeAi: {
            getSettings: () => Promise<unknown>
            stream: (
              req: unknown,
              cb: { onChunk: (c: { type: string; text?: string }) => void },
            ) => typeof fakeHandle
          }
        }
      }
      g.window = {
        aiOfficeAi: {
          getSettings: () => Promise.resolve({ provider: 'custom' }),
          stream: (_req, cb) => {
            cb.onChunk({
              type: 'delta',
              text: 'I am ChatGPT, an AI assistant developed by OpenAI.',
            })
            cb.onChunk({ type: 'done' })
            return fakeHandle
          },
        },
      }
      try {
        await executeOmnibarQuery(
          'who are you?',
          { id: 'home', kind: 'home', title: 'Home' },
          'en',
          {
            onUpdate: () => {},
            onError: (e) => {
              throw new Error(e)
            },
            onDone: (r) => {
              done = { text: r.text }
            },
          },
        )
      } finally {
        delete (globalThis as unknown as Record<string, unknown>).window
      }
      expect(done?.text).toBe('I am ChatGPT, an AI assistant developed by OpenAI.')
    })
  })
})
