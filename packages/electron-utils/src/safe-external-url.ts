/// Single validation gate for every `shell.openExternal` call in the suite.
/// URLs reaching the main process (renderer IPC, window-open handlers, links
/// embedded in documents) are untrusted; anything that is not a well-formed
/// URL on the protocol allowlist is rejected so file:, javascript:, custom
/// app schemes, etc. can never reach the OS handler.
///
/// mailto: is permitted for user-activated document links suite-wide: a mailto
/// hyperlink that does nothing is a file-format fidelity break in a document
/// editor. file:, javascript:, and custom schemes remain rejected.

export interface SafeExternalUrlOptions {
  /** Protocol allowlist (with trailing colon). Defaults to http/https/mailto. */
  allowedProtocols?: readonly string[]
}

const DEFAULT_PROTOCOLS: readonly string[] = ['http:', 'https:', 'mailto:']

/**
 * Returns the URL string when it parses and its protocol is on the allowlist,
 * otherwise null. Callers must not fall back to opening the raw input.
 */
export function safeExternalUrl(url: unknown, options?: SafeExternalUrlOptions): string | null {
  if (typeof url !== 'string') return null
  let parsed: URL
  try {
    parsed = new URL(url)
  } catch {
    return null
  }
  const allowed = options?.allowedProtocols ?? DEFAULT_PROTOCOLS
  return allowed.includes(parsed.protocol) ? url : null
}
