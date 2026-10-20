import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { createInterface, type Interface } from 'node:readline'

const PROTOCOL_VERSION = 1
// ponytail: REQUEST_TIMEOUT_MS doubles as the health signal — it bounds every request, so a
// half-dead sidecar surfaces as 'XLSX sidecar request timed out.' with no ping/heartbeat.
// ceiling: dead and slow are indistinguishable; upgrade path: ping command or spawn watchdog,
// but only if that ambiguity ever matters in practice.
const REQUEST_TIMEOUT_MS = 30_000
// Archive commands stream whole workbooks; large files need more headroom.
const ARCHIVE_TIMEOUT_MS = 180_000
const MAX_STDERR_LENGTH = 8_192
// ponytail: fixed doubling, no jitter — enough for a single-client sidecar.
// upgrade path: jittered delay + failure metrics if respawn storms ever reappear.
const RESPAWN_BACKOFF_INITIAL_MS = 250
const RESPAWN_BACKOFF_MAX_MS = 4_000

interface PendingRequest {
  readonly resolve: (value: unknown) => void
  readonly reject: (error: Error) => void
  readonly timeout: NodeJS.Timeout
}

interface SidecarResponse {
  readonly version: number
  readonly requestId: string
  readonly ok: boolean
  readonly result?: unknown
  readonly error?: {
    readonly code: string
    readonly message: string
  }
}

/** Typed, retryable failure: the sidecar crashed, failed to spawn, or is in respawn cooldown. */
export class XlsxSidecarLostError extends Error {
  readonly code = 'XLSX_SIDECAR_LOST' as const
  constructor(message: string, options?: ErrorOptions) {
    super(message, options)
    this.name = 'XlsxSidecarLostError'
  }
}

export class XlsxSidecarClient {
  private process: ChildProcessWithoutNullStreams | null = null
  private lines: Interface | null = null
  private readonly pending = new Map<string, PendingRequest>()
  private stderr = ''
  private respawnBlockedUntil = 0
  private backoffMs = 0

  constructor(private readonly binaryPath: string) {}

  async open(path: string): Promise<unknown> {
    return this.request({ command: 'open', path })
  }

  async readRange(input: {
    readonly sessionId: string
    readonly sheetId: string
    readonly range: {
      readonly startRow: number
      readonly endRow: number
      readonly startColumn: number
      readonly endColumn: number
    }
  }): Promise<unknown> {
    return this.request({ command: 'read_range', ...input })
  }

  async readFormulaCells(input: {
    readonly sessionId: string
    readonly sheetId: string
  }): Promise<unknown> {
    return this.request({ command: 'read_formula_cells', ...input })
  }

  async close(sessionId: string): Promise<void> {
    await this.request({ command: 'close', sessionId })
  }

  async readMedia(input: {
    readonly sessionId: string
    readonly visualId: string
  }): Promise<unknown> {
    return this.request({ command: 'read_media', ...input })
  }

  async convertWorkbook(input: { path: string; targetPath: string }): Promise<unknown> {
    return this.request({ command: 'convert_workbook', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  async archiveManifest(path: string): Promise<unknown> {
    return this.request({ command: 'archive_manifest', path }, ARCHIVE_TIMEOUT_MS)
  }

  async readEntries(input: {
    readonly path: string
    readonly entries: readonly string[]
    readonly outputDir: string
  }): Promise<unknown> {
    return this.request({ command: 'read_entries', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  async scanEntries(input: {
    readonly path: string
    readonly entries: readonly string[]
    readonly needle: string
  }): Promise<unknown> {
    return this.request({ command: 'scan_entries', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  async saveArchive(input: {
    readonly sourcePath: string
    readonly targetPath: string
    readonly replacements: readonly { name: string; contentPath: string }[]
    readonly removals: readonly string[]
    readonly additions: readonly { name: string; contentPath: string }[]
  }): Promise<unknown> {
    return this.request({ command: 'save_archive', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  /** IronCalc-backed recalculation: load the file, apply the pending edits,
   * evaluate, and return the requested cells' computed values. Fail-soft —
   * callers keep cached values when this errors. */
  async recalcCells(input: {
    readonly path: string
    readonly edits: readonly {
      readonly sheet: string
      readonly row: number
      readonly column: number
      readonly input: string
    }[]
    readonly reads: readonly {
      readonly sheet: string
      readonly range: {
        readonly startRow: number
        readonly endRow: number
        readonly startColumn: number
        readonly endColumn: number
      }
    }[]
  }): Promise<unknown> {
    return this.request({ command: 'recalc_cells', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  /** spawn the sidecar ahead of the first request to hide process cold-start.
   * Warm-up only: during respawn cooldown it defers the typed failure to the first request. */
  start(): void {
    if (Date.now() < this.respawnBlockedUntil) return
    this.ensureStarted()
  }

  getProcessId(): number | null {
    return this.process?.pid ?? null
  }

  stop(): void {
    this.lines?.close()
    this.lines = null
    const child = this.process
    this.process = null // stale exit/error handlers no-op; explicit restart must be immediate
    this.respawnBlockedUntil = 0
    this.backoffMs = 0
    try {
      child?.kill()
    } catch {
      // kill throws EINVAL on Windows when the child never spawned (ENOENT) or is
      // already gone — best-effort, and it must not skip the rejectPending below.
    }
    this.rejectPending(new Error('XLSX sidecar stopped.'))
  }

  private noteSpawnFailure(): void {
    this.backoffMs = Math.min(
      this.backoffMs === 0 ? RESPAWN_BACKOFF_INITIAL_MS : this.backoffMs * 2,
      RESPAWN_BACKOFF_MAX_MS,
    )
    this.respawnBlockedUntil = Date.now() + this.backoffMs
  }

  private request(
    command: Readonly<Record<string, unknown>>,
    timeoutMs: number = REQUEST_TIMEOUT_MS,
  ): Promise<unknown> {
    const child = this.ensureStarted()
    const requestId = crypto.randomUUID()
    const payload = JSON.stringify({
      version: PROTOCOL_VERSION,
      requestId,
      ...command,
    })
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(requestId)
        reject(new Error('XLSX sidecar request timed out.'))
      }, timeoutMs)
      this.pending.set(requestId, { resolve, reject, timeout })
      child.stdin.write(`${payload}\n`, (error) => {
        if (!error) return
        const pending = this.pending.get(requestId)
        if (!pending) return
        clearTimeout(pending.timeout)
        this.pending.delete(requestId)
        // the transport to the sidecar broke: same typed, retryable failure as a crash
        pending.reject(new XlsxSidecarLostError(error.message, { cause: error }))
      })
    })
  }

  private ensureStarted(): ChildProcessWithoutNullStreams {
    if (this.process && !this.process.killed) return this.process
    if (Date.now() < this.respawnBlockedUntil) {
      throw new XlsxSidecarLostError('XLSX sidecar unavailable, retrying shortly.')
    }
    const child = spawn(this.binaryPath, [], {
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true,
    })
    this.process = child
    this.stderr = ''
    this.lines = createInterface({ input: child.stdout })
    this.lines.on('line', (line) => this.handleLine(line))
    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk: string) => {
      this.stderr = `${this.stderr}${chunk}`.slice(-MAX_STDERR_LENGTH)
    })
    child.once('error', (error) => {
      if (this.process !== child) return // stale event: explicit stop() or superseded spawn
      this.process = null
      this.lines?.close()
      this.lines = null
      this.noteSpawnFailure()
      this.rejectPending(new XlsxSidecarLostError(error.message, { cause: error }))
    })
    child.once('exit', (code, signal) => {
      if (this.process !== child) return // stale event: explicit stop() or superseded spawn
      this.process = null
      this.lines?.close()
      this.lines = null
      const detail = this.stderr.trim()
      const reason = detail
        ? `XLSX sidecar exited: ${detail}`
        : `XLSX sidecar exited with code ${String(code)} and signal ${String(signal)}.`
      this.noteSpawnFailure()
      this.rejectPending(new XlsxSidecarLostError(reason))
    })
    return child
  }

  private handleLine(line: string): void {
    const trimmed = line.trim()
    if (!trimmed) return
    let response: SidecarResponse
    try {
      response = JSON.parse(trimmed) as SidecarResponse
    } catch {
      // Non-JSON stdout line (e.g. tracing/log output); do not crash pending requests.
      return
    }
    if (!response || typeof response !== 'object' || typeof response.requestId !== 'string') {
      return
    }
    const pending = this.pending.get(response.requestId)
    if (!pending) return
    clearTimeout(pending.timeout)
    this.pending.delete(response.requestId)

    if (response.version !== PROTOCOL_VERSION || typeof response.ok !== 'boolean') {
      pending.reject(new Error('XLSX sidecar returned an invalid response envelope.'))
      return
    }
    if (response.ok) {
      // ponytail: healthy response resets escalation; cap + doubling above are the ceiling —
      // upgrade path: failure metrics if respawn storms ever reappear.
      this.backoffMs = 0
      pending.resolve(response.result)
      return
    }
    pending.reject(new Error(response.error?.message ?? 'XLSX sidecar request failed.'))
  }

  private rejectPending(error: Error): void {
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timeout)
      pending.reject(error)
    }
    this.pending.clear()
  }
}
