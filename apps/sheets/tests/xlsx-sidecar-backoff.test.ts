import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { XlsxSidecarClient, XlsxSidecarLostError } from '../src/main/xlsx-sidecar-client'

// A path that cannot exist: spawn fails with ENOENT, which drives real respawn failures
// without a protocol-speaking fake sidecar.
// ts-no-test-timers exception: the cooldown window is armed inside a real child_process
// 'error' handler, so deterministic time control cannot work here — faking timers would
// race the real spawn-failure event under test (faking only Date freezes the client's
// Date.now() window math while real time passes). The few short waits below only bound
// the handler delta (~ms) and keep the 250ms/500ms window assertions meaningful.
const MISSING_BINARY = join(tmpdir(), 'araoffice-missing-sidecar', 'xlsx-sidecar.exe')
const COOLDOWN_MESSAGE = 'XLSX sidecar unavailable, retrying shortly.'
const STOPPED_MESSAGE = 'XLSX sidecar stopped.'

/** Awaits a request that must fail and returns its rejection for assertions. */
async function captureFailure(request: Promise<unknown>): Promise<Error> {
  try {
    await request
  } catch (err) {
    return err instanceof Error ? err : new Error(String(err))
  }
  throw new Error('expected the sidecar request to fail')
}

describe('xlsx sidecar respawn backoff', () => {
  it('doubles the respawn window 250ms to 500ms and holds a typed cooldown error', async () => {
    const client = new XlsxSidecarClient(MISSING_BINARY)

    // failure #1 surfaces the spawn error itself; the first window is 250ms
    const first = await captureFailure(client.open('book.xlsx'))
    expect(first.message).not.toBe(COOLDOWN_MESSAGE)

    await new Promise((resolve) => setTimeout(resolve, 320))
    // +320ms clears the 250ms window: spawn is reached again (failure #2, window 500ms)
    const second = await captureFailure(client.open('book.xlsx'))
    expect(second.message).not.toBe(COOLDOWN_MESSAGE)

    await new Promise((resolve) => setTimeout(resolve, 320))
    // +320ms is still inside the doubled 500ms window: cooldown, spawn never reached
    const cooldown = await captureFailure(client.open('book.xlsx'))
    expect(cooldown.message).toBe(COOLDOWN_MESSAGE)
    expect(cooldown).toBeInstanceOf(XlsxSidecarLostError)

    // warm-up start() must defer the typed failure to the first request, never throw
    expect(() => client.start()).not.toThrow()
    expect(client.getProcessId()).toBeNull()
    client.stop()
    // ponytail: RESPAWN_BACKOFF_MAX_MS (4s) unasserted — proving the cap costs ~8s of
    // failure loops and the same Math.min doubling is covered above.
  })

  it('reports a lost sidecar as XlsxSidecarLostError with a stable code', async () => {
    const client = new XlsxSidecarClient(MISSING_BINARY)
    const err = await captureFailure(client.open('book.xlsx'))
    if (!(err instanceof XlsxSidecarLostError)) {
      throw new Error(`expected XlsxSidecarLostError, got: ${err.message}`)
    }
    expect(err.code).toBe('XLSX_SIDECAR_LOST')
    expect(err.name).toBe('XlsxSidecarLostError')
    expect(err.cause).toBeDefined()
    client.stop()
  })

  it('stop() clears the respawn cooldown so the next request reaches spawn', async () => {
    const client = new XlsxSidecarClient(MISSING_BINARY)
    const first = await captureFailure(client.open('book.xlsx'))
    expect(first.message).not.toBe(COOLDOWN_MESSAGE)
    client.stop()
    const afterStop = await captureFailure(client.open('book.xlsx'))
    expect(afterStop.message).not.toBe(COOLDOWN_MESSAGE)
    client.stop()
  })

  it('stop() rejects in-flight requests with a plain Error, not the lost-sidecar type', async () => {
    const client = new XlsxSidecarClient(MISSING_BINARY)
    const pending = client.open('book.xlsx')
    client.stop() // same tick as the request, so the shutdown rejection wins the race
    const err = await captureFailure(pending)
    expect(err.message).toBe(STOPPED_MESSAGE)
    // intentional shutdown is not "lost"
    expect(err).not.toBeInstanceOf(XlsxSidecarLostError)
    expect(client.getProcessId()).toBeNull()
    // ponytail: healthy-response backoff reset (handleLine) unasserted — it needs a live
    // protocol-speaking fake sidecar; the real binary is covered by the other xlsx tests.
  })
})
