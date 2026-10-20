import { describe, expect, it } from 'vitest'
import { absorbOpenAiChunk, toOpenAiMessages, toOpenAiTools } from '../src/custom'
import { defaultAiSettings, redactAiSettings, resolveAiSettings } from '../src/providers'
import type { OpenAiChunkState } from '../src/custom'

const emptyState = (): OpenAiChunkState => ({ acc: new Map() })

describe('custom provider message mapping', () => {
  it('puts the system prompt first and maps text turns verbatim', () => {
    const messages = toOpenAiMessages('You are ARA.', [
      { role: 'user', text: 'hi' },
      { role: 'assistant', text: 'hello' },
    ])
    expect(messages[0]).toEqual({ role: 'system', content: 'You are ARA.' })
    expect(messages[1]).toEqual({ role: 'user', content: 'hi' })
    expect(messages[2]).toEqual({ role: 'assistant', content: 'hello' })
  })

  it('sends inline images as multimodal content blocks', () => {
    const messages = toOpenAiMessages('sys', [
      { role: 'user', text: 'look', images: [{ base64: 'abc', mime: 'image/png' }] },
    ])
    expect(messages[1]).toEqual({
      role: 'user',
      content: [
        { type: 'text', text: 'look' },
        { type: 'image_url', image_url: { url: 'data:image/png;base64,abc' } },
      ],
    })
  })

  it('maps assistant tool calls and the following tool results with ids', () => {
    const messages = toOpenAiMessages('sys', [
      {
        role: 'assistant',
        text: '',
        toolCalls: [{ id: 'call_1', name: 'read_file', input: { path: 'a.txt' } }],
      },
      { role: 'tool', results: [{ id: 'call_1', name: 'read_file', output: 'contents' }] },
    ])
    expect(messages[1]).toEqual({
      role: 'assistant',
      content: '',
      tool_calls: [
        {
          id: 'call_1',
          type: 'function',
          function: { name: 'read_file', arguments: '{"path":"a.txt"}' },
        },
      ],
    })
    expect(messages[2]).toEqual({ role: 'tool', content: 'contents', tool_call_id: 'call_1' })
  })

  it('maps tool definitions to OpenAI function format', () => {
    expect(
      toOpenAiTools([{ name: 't1', description: 'd', inputSchema: { type: 'object' } }]),
    ).toEqual([
      {
        type: 'function',
        function: { name: 't1', description: 'd', parameters: { type: 'object' } },
      },
    ])
  })
})

describe('custom provider SSE chunk absorption', () => {
  it('collects deltas, thinking and passthrough stop reasons', () => {
    const absorbed1 = absorbOpenAiChunk(emptyState(), {
      choices: [{ delta: { content: 'he', reasoning_content: 'plan' } }],
    })
    const absorbed2 = absorbOpenAiChunk(emptyState(), {
      choices: [{ delta: { content: 'llo' }, finish_reason: 'stop' }],
    })
    expect(absorbed1.deltas).toEqual(['he'])
    expect(absorbed1.thinking).toEqual(['plan'])
    expect(absorbed2.deltas).toEqual(['llo'])
    expect(absorbed2.stopReason).toBe('stop')
  })

  it('accumulates fragmented tool calls by index and emits them once args parse', () => {
    const state = emptyState()
    absorbOpenAiChunk(state, {
      choices: [
        {
          delta: {
            tool_calls: [
              { index: 0, id: 'c1', function: { name: 'do_thing', arguments: '{"a":1' } },
            ],
          },
        },
      ],
    })
    // nothing emitted yet: args are incomplete
    expect(state.acc.get(0)?.emitted).toBe(false)
    const second = absorbOpenAiChunk(state, {
      choices: [
        {
          delta: {
            tool_calls: [
              {
                index: 0,
                function: { arguments: ',"b":2}' },
              },
            ],
          },
          finish_reason: 'tool_calls',
        },
      ],
    })
    expect(second.toolCalls).toHaveLength(1)
    expect(second.toolCalls[0]).toMatchObject({ id: 'c1', name: 'do_thing', input: { a: 1, b: 2 } })
    expect(second.toolCalls[0].inputError).toBeUndefined()
    // emitted only once
    const third = absorbOpenAiChunk(state, {
      choices: [{ delta: {}, finish_reason: 'tool_calls' }],
    })
    expect(third.toolCalls).toHaveLength(0)
  })

  it('flags leftover tool calls as truncated when the stream ends mid-args', () => {
    const state = emptyState()
    absorbOpenAiChunk(state, {
      choices: [
        { delta: { tool_calls: [{ index: 0, function: { name: 't', arguments: '{"a":1' } }] } },
      ],
    })
    const last = absorbOpenAiChunk(state, { choices: [{ delta: {}, finish_reason: 'length' }] })
    expect(last.toolCalls).toHaveLength(1)
    expect(last.toolCalls[0].truncated).toBe(true)
    // args had started but never parsed: the loop receives truncated=true and
    // asks the model to split the call instead of retrying broken JSON
    expect(last.toolCalls[0].inputError).toBeUndefined()
    expect(last.toolCalls[0].input).toEqual({})
  })
})

describe('custom provider settings', () => {
  it('keeps both providers and drops junk base urls on resolve', () => {
    const resolved = resolveAiSettings(
      {
        provider: 'custom',
        providers: {
          google: { model: 'gemini-3.8-flash', apiKey: 'g' },
          custom: { model: ' mi ', baseUrl: '  https://api.example.com/v1/// ', apiKey: ' k ' },
        },
      },
      defaultAiSettings(),
    )
    expect(resolved.provider).toBe('custom')
    expect(resolved.providers.google.model).toBe('gemini-3.8-flash')
    expect(resolved.providers.custom).toEqual({
      model: 'mi',
      baseUrl: 'https://api.example.com/v1',
      apiKey: 'k',
    })
  })

  it('falls back to google on junk provider but keeps custom fields for later', () => {
    const resolved = resolveAiSettings(
      {
        provider: 'removed' as never,
        providers: {
          google: { model: 'gemini-3.8-flash', apiKey: 'g' },
          custom: { model: 'm', baseUrl: 'not a url', apiKey: 'k' },
        },
      },
      defaultAiSettings(),
    )
    expect(resolved.provider).toBe('google')
    expect(resolved.providers.custom).toEqual({ model: 'm', baseUrl: '', apiKey: 'k' })
  })

  it('never leaks the stored key through a renderer-facing read', () => {
    const stored = resolveAiSettings(
      {
        provider: 'custom',
        providers: {
          google: { model: 'gemini-3.8-flash', apiKey: 'g' },
          custom: { model: 'm', baseUrl: 'https://api.example.com/v1', apiKey: 'secret' },
        },
      },
      defaultAiSettings(),
    )
    expect(redactAiSettings(stored).providers.custom.apiKey).toBe('')
    expect(redactAiSettings(stored).providers.custom.hasApiKey).toBe(true)
    expect(stored.providers.custom.apiKey).toBe('secret')

    const emptyKeyStored = resolveAiSettings(
      {
        provider: 'custom',
        providers: {
          google: { model: 'gemini-3.8-flash', apiKey: '' },
          custom: { model: 'm', baseUrl: 'https://api.example.com/v1', apiKey: '' },
        },
      },
      defaultAiSettings(),
    )
    expect(redactAiSettings(emptyKeyStored).providers.custom.hasApiKey).toBe(false)
    expect(redactAiSettings(emptyKeyStored).providers.google.hasApiKey).toBe(false)
  })

  it('fetchCustomModels parses OpenAI-compatible /models response', async () => {
    const originalFetch = globalThis.fetch
    globalThis.fetch = (async (url: string | URL | Request) => {
      expect(String(url)).toBe('https://api.example.com/v1/models')
      return {
        ok: true,
        json: async () => ({
          data: [{ id: 'deepseek-v4.1-flash-free' }, { id: 'gpt-4o' }, { id: '' }],
        }),
      } as unknown as Response
    }) as typeof fetch

    try {
      const { fetchCustomModels } = await import('../src/custom')
      const models = await fetchCustomModels({
        model: 'm',
        baseUrl: 'https://api.example.com/v1',
        apiKey: 'secret',
      })
      expect(models).toEqual(['deepseek-v4.1-flash-free', 'gpt-4o'])
    } finally {
      globalThis.fetch = originalFetch
    }
  })
})
