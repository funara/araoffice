import { describe, expect, it } from 'vitest'
import { streamForProvider } from '../src/stream'

describe('provider streaming', () => {
  it('rejects an unknown provider id', async () => {
    await expect(
      streamForProvider('unknown' as never, { model: 'm' }, 'sys', [], [], 100, {
        signal: new AbortController().signal,
        onDelta: () => undefined,
        onToolCall: () => undefined,
      }),
    ).rejects.toThrow(/Unknown provider/)
  })
})
