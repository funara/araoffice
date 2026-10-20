import { describe, expect, it, vi } from 'vitest'

describe('slides handover launch & buffering', () => {
  it('buffers early menu commands before listener attaches and flushes on subscription', () => {
    let activeHandler: ((cmd: string, payload?: string) => void) | null = null
    const pendingCommands: Array<{ cmd: string; payload?: string }> = []

    // Simulate preload listener
    const simulateIpcReceive = (cmd: string, payload?: string) => {
      if (activeHandler) {
        activeHandler(cmd, payload)
      } else {
        pendingCommands.push({ cmd, payload })
      }
    }

    // Preload receives early IPC command while React hasn't mounted
    simulateIpcReceive('ai-prompt', 'Create presentation about quarterly sales')
    expect(pendingCommands.length).toBe(1)

    // React component mounts and subscribes
    const onMenuCommand = vi.fn()
    activeHandler = onMenuCommand

    while (pendingCommands.length > 0) {
      const item = pendingCommands.shift()
      if (item) activeHandler(item.cmd, item.payload)
    }

    expect(onMenuCommand).toHaveBeenCalledTimes(1)
    expect(onMenuCommand).toHaveBeenCalledWith(
      'ai-prompt',
      'Create presentation about quarterly sales',
    )
    expect(pendingCommands.length).toBe(0)
  })

  it('defers preset autoRun until providerReady is true', () => {
    let providerReady = false
    const runWith = vi.fn()
    let pendingAutoRun: { text: string } | null = null

    // Effect on preset
    const onPresetReceived = (preset: { text: string; autoRun?: boolean }) => {
      if (!preset) return
      if (preset.autoRun) {
        if (providerReady) {
          runWith(preset.text)
        } else {
          pendingAutoRun = { text: preset.text }
        }
      }
    }

    // Preset arrives on initial mount while providerReady is still false
    onPresetReceived({ text: 'Generate slide deck', autoRun: true })
    expect(runWith).not.toHaveBeenCalled()
    expect(pendingAutoRun).toEqual({ text: 'Generate slide deck' })

    // Later, provider status resolves to ready
    providerReady = true
    if (providerReady && pendingAutoRun) {
      const pending = pendingAutoRun as { text: string }
      pendingAutoRun = null
      runWith(pending.text)
    }

    expect(runWith).toHaveBeenCalledTimes(1)
    expect(runWith).toHaveBeenCalledWith('Generate slide deck')
    expect(pendingAutoRun).toBeNull()
  })
})
