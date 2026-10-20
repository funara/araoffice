import mermaid from 'mermaid'

let initialized = false

export function ensureMermaid(): void {
  if (initialized) return
  initialized = true
  mermaid.initialize({
    startOnLoad: false,
    securityLevel: 'strict',
    htmlLabels: false,
    theme: 'default',
    suppressErrorRendering: true,
  })
}

let renderSeq = 0

export async function renderMermaidDiagram(
  code: string,
): Promise<{ svg?: string; error?: string }> {
  const trimmed = code.trim()
  if (!trimmed) {
    return { svg: '' }
  }

  ensureMermaid()

  const id = `mermaid-render-${Date.now()}-${++renderSeq}`
  try {
    // Validate syntax first
    await mermaid.parse(trimmed)
    await document.fonts?.ready
    const { svg } = await mermaid.render(id, trimmed)
    return { svg }
  } catch (err) {
    // Clean up any lingering temporary error nodes injected by mermaid into document.body
    document.getElementById(id)?.remove()
    document.getElementById(`d${id}`)?.remove()
    // ponytail: purge dangling mermaid scratch containers to prevent DOM memory leaks
    document.querySelectorAll('div[id^="dmermaid-render-"]').forEach((el) => el.remove())
    const message = err instanceof Error ? err.message : String(err)
    return { error: message }
  }
}
