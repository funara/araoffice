import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { deflateRawSync } from 'node:zlib'
import { beforeEach, describe, expect, it } from 'vitest'
import { createDiagramGraph } from '../src/renderer/maxgraph/createGraph'
import { addRectangle, connectCells, setCellLabel } from '../src/renderer/maxgraph/diagramCommands'
import {
  extractGraphModelNode,
  formatStyleString,
  loadDrawioIntoGraph,
  serializeGraphToDrawio,
} from '../src/renderer/maxgraph/drawio'

describe('Draw.io Format Compatibility', () => {
  let container: HTMLDivElement

  beforeEach(() => {
    container = document.createElement('div')
    document.body.appendChild(container)
  })

  it('loads simple-flow.drawio fixture and verifies cells structure', async () => {
    const fixturePath = join(__dirname, 'fixtures', 'simple-flow.drawio')
    const xml = readFileSync(fixturePath, 'utf8')

    const { graph, destroy } = createDiagramGraph(container)
    try {
      const ok = await loadDrawioIntoGraph(graph, xml)
      expect(ok).toBe(true)

      const model = graph.getDataModel()
      const startCell = model.getCell('2')
      const processCell = model.getCell('3')
      const endCell = model.getCell('4')
      const edge1 = model.getCell('5')
      const edge2 = model.getCell('6')

      expect(startCell).toBeDefined()
      expect(startCell?.value).toBe('Start')
      expect(startCell?.vertex).toBeTruthy()
      expect(startCell?.geometry?.x).toBe(100)
      expect(startCell?.geometry?.y).toBe(100)

      expect(processCell).toBeDefined()
      expect(processCell?.value).toBe('Process')

      expect(endCell).toBeDefined()
      expect(endCell?.value).toBe('End')

      expect(edge1).toBeDefined()
      expect(edge1?.edge).toBeTruthy()
      expect(edge1?.source?.id).toBe('2')
      expect(edge1?.target?.id).toBe('3')

      expect(edge2).toBeDefined()
      expect(edge2?.edge).toBeTruthy()
      expect(edge2?.source?.id).toBe('3')
      expect(edge2?.target?.id).toBe('4')
    } finally {
      destroy()
    }
  })

  it('edits diagram and saves valid .drawio XML that round-trips cleanly', async () => {
    const fixturePath = join(__dirname, 'fixtures', 'simple-flow.drawio')
    const xml = readFileSync(fixturePath, 'utf8')

    const { graph, destroy } = createDiagramGraph(container)
    try {
      await loadDrawioIntoGraph(graph, xml)

      // Add a new Step 3 node and connect End to it
      const endCell = graph.getDataModel().getCell('4')
      expect(endCell).toBeDefined()
      const step3 = addRectangle(graph, {
        x: 650,
        y: 100,
        width: 100,
        height: 60,
        label: 'Archive',
      })
      connectCells(graph, endCell!, step3, 'archive')

      // Save to .drawio XML string
      const savedXml = serializeGraphToDrawio(graph)
      expect(savedXml).toContain('<mxfile')
      expect(savedXml).toContain('<diagram')
      expect(savedXml).toContain('<mxGraphModel')
      expect(savedXml).toContain('value="Archive"')
      expect(savedXml).toContain('value="archive"')

      // Re-load the saved XML into a fresh graph
      const container2 = document.createElement('div')
      document.body.appendChild(container2)
      const instance2 = createDiagramGraph(container2)
      try {
        const reloaded = await loadDrawioIntoGraph(instance2.graph, savedXml)
        expect(reloaded).toBe(true)

        const model2 = instance2.graph.getDataModel()
        expect(model2.getCell('2')?.value).toBe('Start')
        expect(model2.getCell('3')?.value).toBe('Process')
        expect(model2.getCell('4')?.value).toBe('End')
        expect(model2.getCell(step3.id!)?.value).toBe('Archive')
      } finally {
        instance2.destroy()
        container2.remove()
      }
    } finally {
      destroy()
    }
  })

  it('preserves mxfile pages, metadata, and unknown cell XML when saving', async () => {
    const xml = `
      <mxfile host="custom-host" custom-file="keep">
        <diagram id="main" name="Main">
          <mxGraphModel custom-model="keep">
            <root>
              <mxCell id="0"/>
              <mxCell id="1" parent="0"/>
              <mxCell id="2" value="Before" vertex="1" parent="1" custom-cell="keep">
                <mxGeometry x="40" y="40" width="100" height="60" as="geometry"/>
                <unknown-cell-node token="keep"/>
              </mxCell>
              <unknown-root-node token="keep"/>
            </root>
          </mxGraphModel>
        </diagram>
        <diagram id="other" name="Other">
          <mxGraphModel>
            <root>
              <mxCell id="0"/>
              <mxCell id="1" parent="0"/>
              <mxCell id="other-cell" value="Untouched" vertex="1" parent="1"/>
            </root>
          </mxGraphModel>
        </diagram>
        <metadata token="keep"/>
      </mxfile>
    `

    const { graph, destroy } = createDiagramGraph(container)
    try {
      await loadDrawioIntoGraph(graph, xml)
      const cell = graph.getDataModel().getCell('2')
      expect(cell).toBeDefined()
      setCellLabel(graph, cell!, 'After')

      const saved = serializeGraphToDrawio(graph)
      const savedDoc = new DOMParser().parseFromString(saved, 'text/xml')
      const mainModel = savedDoc.querySelector('diagram[id="main"] > mxGraphModel')
      expect(savedDoc.documentElement.getAttribute('custom-file')).toBe('keep')
      expect(savedDoc.querySelector('metadata[token="keep"]')).toBeTruthy()
      expect(savedDoc.querySelector('diagram[id="other"] [value="Untouched"]')).toBeTruthy()
      expect(mainModel?.getAttribute('custom-model')).toBe('keep')
      expect(mainModel?.querySelector('mxCell[id="2"]')?.getAttribute('value')).toBe('After')
      expect(mainModel?.querySelector('mxCell[id="2"]')?.getAttribute('custom-cell')).toBe('keep')
      expect(mainModel?.querySelector('unknown-cell-node[token="keep"]')).toBeTruthy()
      expect(mainModel?.querySelector('unknown-root-node[token="keep"]')).toBeTruthy()
    } finally {
      destroy()
    }
  })

  it('handles compressed base64 Draw.io diagrams', async () => {
    const rawModel =
      '<mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/><mxCell id="2" value="Compressed Node" vertex="1" parent="1"><mxGeometry x="50" y="50" width="120" height="60" as="geometry"/></mxCell></root></mxGraphModel>'

    const deflated = deflateRawSync(encodeURIComponent(rawModel)).toString('base64')
    const compressedDrawio = `<mxfile host="Electron"><diagram id="p1">${deflated}</diagram></mxfile>`

    const { graph, destroy } = createDiagramGraph(container)
    try {
      const ok = await loadDrawioIntoGraph(graph, compressedDrawio)
      expect(ok).toBe(true)

      const cell = graph.getDataModel().getCell('2')
      expect(cell).toBeDefined()
      expect(cell?.value).toBe('Compressed Node')
    } finally {
      destroy()
    }
  })

  it('formats style strings correctly', () => {
    expect(formatStyleString(undefined)).toBe('')
    expect(formatStyleString('rounded=1;')).toBe('rounded=1;')
    expect(
      formatStyleString({
        rounded: true,
        fillColor: '#ffffff',
        strokeWidth: 2,
      }),
    ).toBe('rounded;fillColor=#ffffff;strokeWidth=2;')
  })

  it('gracefully returns null for invalid XML', async () => {
    const node = await extractGraphModelNode('<<<invalid xml>>>')
    expect(node).toBeNull()
  })
})
