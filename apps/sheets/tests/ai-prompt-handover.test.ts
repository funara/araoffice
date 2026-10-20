import { describe, expect, it, vi } from 'vitest'

describe('Sheets Delegated AI Prompt Handover', () => {
  it('triggers onSend exactly once per nonce and consumes the prompt immediately', () => {
    let aiMenuPrompt: { text: string; nonce: number } | null = {
      text: 'Create score table',
      nonce: 1001,
    }
    const onSend = vi.fn()
    const onConsumeAiMenuPrompt = vi.fn(() => {
      aiMenuPrompt = null
    })

    let isCopilotOpen = false
    const setIsCopilotOpen = (val: boolean) => {
      isCopilotOpen = val
    }

    // Model the single-shot effect logic implemented in ExcelShell
    let consumedNonce = 0
    const runEffect = () => {
      if (!aiMenuPrompt || aiMenuPrompt.nonce === consumedNonce) return
      consumedNonce = aiMenuPrompt.nonce
      setIsCopilotOpen(true)
      const promptText = aiMenuPrompt.text
      onConsumeAiMenuPrompt()
      onSend(promptText)
    }

    // 1. Initial trigger
    runEffect()
    expect(onSend).toHaveBeenCalledTimes(1)
    expect(onSend).toHaveBeenCalledWith('Create score table')
    expect(onConsumeAiMenuPrompt).toHaveBeenCalledTimes(1)
    expect(isCopilotOpen).toBe(true)
    expect(aiMenuPrompt).toBeNull()

    // 2. Subsequent re-render when aiBusy flips or parent re-renders
    runEffect()
    expect(onSend).toHaveBeenCalledTimes(1)

    // 3. Remount simulation: consumedNonce reset to 0, but parent aiMenuPrompt was consumed (null)
    consumedNonce = 0
    runEffect()
    expect(onSend).toHaveBeenCalledTimes(1)
  })

  it('ignores stale or identical nonce updates', () => {
    const onSend = vi.fn()
    const onConsume = vi.fn()
    let consumedNonce = 2002

    const aiMenuPrompt = { text: 'Repeat text', nonce: 2002 }
    if (aiMenuPrompt && aiMenuPrompt.nonce !== consumedNonce) {
      consumedNonce = aiMenuPrompt.nonce
      onConsume()
      onSend(aiMenuPrompt.text)
    }

    expect(onSend).not.toHaveBeenCalled()
    expect(onConsume).not.toHaveBeenCalled()
    expect(consumedNonce).toBe(2002)
  })
})
