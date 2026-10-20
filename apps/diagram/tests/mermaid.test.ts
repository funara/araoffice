import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { beforeAll, describe, expect, it } from 'vitest'
import { validateAndSanitizeMermaid } from '../src/renderer/ai/diagram-guardrail'
import { mermaidViewport } from '../src/renderer/mermaid/MermaidEditor'
import { renderMermaidDiagram } from '../src/renderer/mermaid/renderMermaid'

describe('Mermaid Rendering and Error Handling', () => {
  beforeAll(() => {
    // Polyfill CSSStyleSheet for jsdom
    if (
      typeof globalThis.CSSStyleSheet === 'undefined' ||
      !globalThis.CSSStyleSheet.prototype.replaceSync
    ) {
      globalThis.CSSStyleSheet = class CSSStyleSheet {
        cssRules = []
        insertRule() {
          return 0
        }
        replaceSync() {}
      } as unknown as typeof globalThis.CSSStyleSheet
    }

    // Polyfill SVG measurement APIs for jsdom
    const svgProto = SVGElement.prototype as unknown as {
      getBBox?: () => DOMRect
      getComputedTextLength?: () => number
    }
    if (typeof SVGElement !== 'undefined' && !svgProto.getBBox) {
      svgProto.getBBox = () => ({
        x: 0,
        y: 0,
        width: 100,
        height: 50,
        top: 0,
        bottom: 50,
        left: 0,
        right: 100,
        toJSON: () => ({}),
      })
    }
    if (!svgProto.getComputedTextLength) svgProto.getComputedTextLength = () => 100
  })

  it('fits SVG content at 100% and expands outward when zooming out', () => {
    expect(mermaidViewport({ x: 10, y: 20, width: 100, height: 50 }, { x: 0, y: 0 }, 100)).toBe(
      '10 20 100 50',
    )
    expect(mermaidViewport({ x: 10, y: 20, width: 100, height: 50 }, { x: 0, y: 0 }, 50)).toBe(
      '-40 -5 200 100',
    )
  })

  it('applies pan at the default zoom level', () => {
    expect(mermaidViewport({ x: 10, y: 20, width: 100, height: 50 }, { x: 12, y: -8 }, 100)).toBe(
      '-2 28 100 50',
    )
  })

  it('renders valid simple-flow.mmd to SVG output', async () => {
    const fixturePath = join(__dirname, 'fixtures', 'simple-flow.mmd')
    const code = readFileSync(fixturePath, 'utf8')

    const result = await renderMermaidDiagram(code)
    expect(result.error).toBeUndefined()
    expect(result.svg).toBeDefined()
    expect(result.svg).toContain('<svg')
    expect(result.svg).toContain('Start')
    expect(result.svg).toContain('Process')
    expect(result.svg).toContain('End')
  })

  it('renders flowchart connectors and standalone class diagrams', async () => {
    const flowchart = await renderMermaidDiagram(`flowchart TD
  Start([Start]) --> Process[Step 1: Execute Task]
  Process --> Choice{Is Success?}
  Choice -- Yes --> Finish([Finish])
  Choice -- No --> Retry[Step 2: Retry]
  Retry --> Process`)
    const classDiagram = await renderMermaidDiagram(`classDiagram
  class BankAccount {
    +String owner
    +BigDecimal balance
    +deposit(amount)
    +withdraw(amount)
  }`)

    expect(flowchart.svg).toContain('flowchart-link')
    expect(flowchart.svg).toMatch(/d="M[^"]+"/)
    expect(flowchart.svg).not.toContain('foreignObject')
    expect(classDiagram.svg).toContain('BankAccount')
  })

  it('keeps SVG labels after AI validation', async () => {
    const code = 'flowchart TD\n  Start --> Finish'
    expect((await validateAndSanitizeMermaid(code)).valid).toBe(true)
    expect((await renderMermaidDiagram(code)).svg).not.toContain('foreignObject')
  })

  it('handles invalid Mermaid code gracefully without crashing', async () => {
    const invalidCode = 'flowchart TD\n  invalid [[[ syntax !!!'

    const result = await renderMermaidDiagram(invalidCode)
    expect(result.svg).toBeUndefined()
    expect(result.error).toBeDefined()
    expect(typeof result.error).toBe('string')
    expect(result.error!.length).toBeGreaterThan(0)
  })

  it('returns empty string for empty or blank input', async () => {
    const result1 = await renderMermaidDiagram('')
    expect(result1.svg).toBe('')
    expect(result1.error).toBeUndefined()

    const result2 = await renderMermaidDiagram('   \n  \t  ')
    expect(result2.svg).toBe('')
    expect(result2.error).toBeUndefined()
  })
})
