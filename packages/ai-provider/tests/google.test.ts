import { describe, expect, it } from 'vitest'
import {
  absorbGoogleChunk,
  assertGoogleModel,
  buildGoogleBody,
  endpointFor,
  normalizeGoogleStopReason,
  toGoogleContents,
  toGoogleTools,
} from '../src/google'

describe('Gemini request translation', () => {
  it('keeps the system prompt out of the contents list', () => {
    const body = buildGoogleBody('You are ARA Assistant.', [{ role: 'user', text: 'hi' }], [])
    expect(body.systemInstruction).toEqual({ parts: [{ text: 'You are ARA Assistant.' }] })
    expect(body.contents).toEqual([{ role: 'user', parts: [{ text: 'hi' }] }])
  })

  it('maps user images to inline data parts', () => {
    const contents = toGoogleContents([
      {
        role: 'user',
        text: 'look',
        images: [{ mime: 'image/png', base64: 'AAA' }],
      },
    ])
    expect(contents[0].parts[1]).toEqual({ inlineData: { mimeType: 'image/png', data: 'AAA' } })
  })

  it('maps assistant tool calls to function calls and tool results to function responses', () => {
    const contents = toGoogleContents([
      {
        role: 'assistant',
        text: 'reading',
        toolCalls: [{ id: 'c1', name: 'read', input: { i: 1 } }],
      },
      { role: 'tool', results: [{ id: 'c1', name: 'read', output: 'body text' }] },
    ])
    expect(contents[0]).toEqual({
      role: 'model',
      parts: [{ text: 'reading' }, { functionCall: { name: 'read', args: { i: 1 } } }],
    })
    expect(contents[1]).toEqual({
      role: 'user',
      parts: [{ functionResponse: { name: 'read', response: { output: 'body text' } } }],
    })
  })

  it('maps tool definitions to Gemini function declarations', () => {
    const declarations = toGoogleTools([
      { name: 'read_blocks', description: 'read', inputSchema: { type: 'object' } },
    ])
    expect(declarations).toEqual([
      {
        functionDeclarations: [
          { name: 'read_blocks', description: 'read', parameters: { type: 'object' } },
        ],
      },
    ])
  })

  it('omits tools and generation config when they carry nothing', () => {
    const body = buildGoogleBody('sys', [{ role: 'user', text: 'hi' }], [])
    expect(body.tools).toBeUndefined()
    expect(body.generationConfig).toBeUndefined()
  })
})

describe('Gemini stream chunk absorption', () => {
  it('forwards text deltas, thought parts, and function calls', () => {
    const absorbed = absorbGoogleChunk({
      candidates: [
        {
          content: {
            parts: [
              { text: 'visible' },
              { text: 'reasoning', thought: true },
              { functionCall: { name: 'read_blocks', args: { index: 2 } } },
            ],
          },
        },
      ],
    })
    expect(absorbed.deltas).toEqual(['visible'])
    expect(absorbed.thinking).toEqual(['reasoning'])
    expect(absorbed.toolCalls).toEqual([
      { id: 'call_read_blocks', name: 'read_blocks', input: { index: 2 } },
    ])
  })

  it('maps the Gemini token cap onto the agent stop reason', () => {
    expect(normalizeGoogleStopReason('MAX_TOKENS')).toBe('max_tokens')
    expect(normalizeGoogleStopReason('STOP')).toBe('STOP')
    const absorbed = absorbGoogleChunk({ candidates: [{ finishReason: 'MAX_TOKENS' }] })
    expect(absorbed.stopReason).toBe('max_tokens')
  })

  it('returns empty results for a chunk without candidates', () => {
    const absorbed = absorbGoogleChunk({})
    expect(absorbed.deltas).toEqual([])
    expect(absorbed.toolCalls).toEqual([])
    expect(absorbed.stopReason).toBeUndefined()
  })
})

describe('Gemini model id validation', () => {
  it('accepts the supported Gemini model ids', () => {
    for (const id of ['gemini-3.8-flash', 'gemini-3.5-flash-lite', 'gemini-3.1-pro']) {
      expect(assertGoogleModel(id)).toBe(id)
    }
  })

  it('rejects an id that could escape the model path', () => {
    for (const id of ['../../evil', 'gemini/../x', 'gpt-4o', 'gemini-3.8-flash?x=1', '']) {
      expect(() => assertGoogleModel(id)).toThrow(/Invalid Gemini model id/)
    }
  })
})

describe('Gemini request endpoint', () => {
  it('always targets the Google AI Studio host', () => {
    for (const model of ['gemini-3.8-flash', 'gemini-3.1-pro']) {
      expect(endpointFor(model, false).hostname).toBe('generativelanguage.googleapis.com')
      expect(endpointFor(model, true).hostname).toBe('generativelanguage.googleapis.com')
    }
  })

  it('uses the documented generateContent and streamGenerateContent paths', () => {
    expect(endpointFor('gemini-3.8-flash', false).pathname).toBe(
      '/v1beta/models/gemini-3.8-flash:generateContent',
    )
    expect(endpointFor('gemini-3.8-flash', true).pathname).toBe(
      '/v1beta/models/gemini-3.8-flash:streamGenerateContent',
    )
  })

  it('keeps a hostile model id from escaping the models path', () => {
    expect(() => endpointFor('../../evil', false)).toThrow(/Invalid Gemini/)
  })
})
