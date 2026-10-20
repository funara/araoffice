/**
 * Shared Server-Sent Events reader for the direct-connect providers.
 *
 * Both the OpenAI-compatible and the Gemini wire formats stream `data:` lines
 * separated by a blank line; only the payload shape differs. Reading the
 * transport once keeps the two providers substitutable (LSP).
 */

/** Extract the payload of one SSE `data:` line; null for comments/keepalives. */
export function sseData(line: string): string | null {
  if (!line.startsWith('data:')) return null
  const payload = line.slice(5).trim()
  return payload.length > 0 ? payload : null
}

/**
 * Read an SSE response body and hand every `data:` payload to `onData`.
 * Tolerates `\r\n` from proxies and a trailing event without a blank line.
 */
export async function readSseStream(
  response: Response,
  onData: (data: string) => void,
): Promise<void> {
  if (!response.body) throw new Error('AI response had no body')
  const reader = response.body.getReader()
  const decoder = new TextDecoder()
  let buffer = ''
  const drain = (text: string): void => {
    for (const line of text.split('\n')) {
      const data = sseData(line)
      if (data !== null) onData(data)
    }
  }
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    buffer += decoder.decode(value, { stream: true })
    let boundary = buffer.indexOf('\n\n')
    while (boundary >= 0) {
      const event = buffer.slice(0, boundary).replace(/\r/g, '')
      buffer = buffer.slice(boundary + 2)
      drain(event)
      boundary = buffer.indexOf('\n\n')
    }
  }
  buffer += decoder.decode()
  if (buffer.trim()) drain(buffer.trim())
}
