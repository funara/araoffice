import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { ConnectionHandler, Graph } from '@maxgraph/core'
import { createDiagramGraph } from '../src/renderer/maxgraph/createGraph'
import { exportGraphToSvg } from '../src/renderer/maxgraph/exportDiagram'
import {
  addCylinder,
  addDiamond,
  addEllipse,
  addRectangle,
  addRoundedRectangle,
  addShape,
  addShapeAtPoint,
  addText,
  alignSelection,
  autoLayout,
  connectCells,
  connectSelection,
  deleteSelection,
  distributeSelection,
  duplicateSelection,
  redo,
  undo,
} from '../src/renderer/maxgraph/diagramCommands'

describe('Diagram Commands', () => {
  let container: HTMLDivElement

  beforeEach(() => {
    container = document.createElement('div')
    document.body.appendChild(container)
  })

  it('opens a node picker instead of creating a dangling connection', () => {
    const onEmptyConnection = vi.fn()
    const { graph, destroy } = createDiagramGraph(container, onEmptyConnection)
    try {
      const source = addRectangle(graph)
      const handler = graph.getPlugin<ConnectionHandler>('ConnectionHandler')
      if (!handler) throw new Error('Missing connection handler')

      handler.connect(source, null, new MouseEvent('mouseup', { clientX: 240, clientY: 160 }), null)

      expect(onEmptyConnection).toHaveBeenCalledWith(source, 240, 160)
      expect(graph.getChildVertices(graph.getDefaultParent())).toHaveLength(1)
      expect(graph.getChildEdges(graph.getDefaultParent())).toHaveLength(0)
    } finally {
      destroy()
    }
  })

  it('inserts all standard shape types', () => {
    const { graph, destroy } = createDiagramGraph(container)
    try {
      const rect = addRectangle(graph, { label: 'Box' })
      const rounded = addRoundedRectangle(graph, { label: 'Round' })
      const ellipse = addEllipse(graph, { label: 'Circle' })
      const doubleEllipse = addShape(graph, 'doubleEllipse', { label: 'Double circle' })
      const diamond = addDiamond(graph, { label: 'Condition' })
      const cylinder = addCylinder(graph, { label: 'DB' })
      const text = addText(graph, { label: 'Note' })
      const triangle = addShape(graph, 'triangle', { label: 'Triangle' })
      const hexagon = addShape(graph, 'hexagon', { label: 'Hexagon' })
      const cloud = addShape(graph, 'cloud', { label: 'Cloud' })
      expect(rect.value).toBe('Box')
      expect(rect.vertex).toBeTruthy()

      expect(rounded.value).toBe('Round')
      expect(rounded.vertex).toBeTruthy()

      expect(ellipse.value).toBe('Circle')
      expect(doubleEllipse.value).toBe('Double circle')
      expect(doubleEllipse.vertex).toBeTruthy()
      expect(ellipse.vertex).toBeTruthy()

      expect(diamond.value).toBe('Condition')
      expect(diamond.vertex).toBeTruthy()

      expect(cylinder.value).toBe('DB')
      expect(cylinder.vertex).toBeTruthy()

      expect(text.value).toBe('Note')
      expect(text.vertex).toBeTruthy()
      expect(triangle.value).toBe('Triangle')
      expect(hexagon.value).toBe('Hexagon')
      expect(cloud.value).toBe('Cloud')

      const vertices = graph.getChildVertices(graph.getDefaultParent())
      expect(vertices.length).toBe(10)
    } finally {
      destroy()
    }
  })

  it('centers a shape at the requested canvas point', () => {
    const { graph, destroy } = createDiagramGraph(container)
    try {
      const cell = addShapeAtPoint(graph, 'doubleEllipse', 300, 200)

      expect(cell.geometry?.x).toBe(250)
      expect(cell.geometry?.y).toBe(160)
    } finally {
      destroy()
    }
  })

  it('preserves negative diagram bounds in SVG export', () => {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg')
    container.appendChild(svg)
    const graph = {
      container,
      getGraphBounds: () => ({ x: -80, y: -40, width: 120, height: 60 }),
    } as unknown as Graph

    expect(exportGraphToSvg(graph, 20)).toContain('viewBox="-100 -60 160 100"')
  })

  it('connects two vertices with an edge', () => {
    const { graph, destroy } = createDiagramGraph(container)
    try {
      const v1 = addRectangle(graph, { label: 'Source' })
      const v2 = addRectangle(graph, { label: 'Target', x: 300, y: 100 })
      const edge = connectCells(graph, v1, v2, 'leads to')

      expect(edge).toBeDefined()
      expect(edge.edge).toBeTruthy()
      expect(edge.source?.id).toBe(v1.id)
      expect(edge.target?.id).toBe(v2.id)
      expect(edge.value).toBe('leads to')
      expect(edge.style.labelBackgroundColor).toBe('#ffffff')
      expect(edge.style.labelBorderColor).toBe('#cbd5e1')
      expect(edge.style.labelPadding).toBe(4)

      const edges = graph.getChildEdges(graph.getDefaultParent())
      expect(edges.length).toBe(1)
    } finally {
      destroy()
    }
  })

  it('routes connected diagrams by flow direction and clears selection', () => {
    const { graph, destroy } = createDiagramGraph(container)
    try {
      const source = addRectangle(graph, { label: 'Source', x: 500, y: 320 })
      const middle = addRectangle(graph, { label: 'Middle', x: 80, y: 80 })
      const target = addRectangle(graph, { label: 'Target', x: 720, y: 20 })
      const firstEdge = connectCells(graph, source, middle)
      const secondEdge = connectCells(graph, middle, target)
      graph.setSelectionCell(secondEdge)

      autoLayout(graph, 'vertical')

      expect(source.geometry?.y).toBeLessThan(middle.geometry?.y ?? 0)
      expect(middle.geometry?.y).toBeLessThan(target.geometry?.y ?? 0)
      expect(firstEdge.geometry?.points?.length).toBeGreaterThan(0)
      expect(secondEdge.geometry?.points?.length).toBeGreaterThan(0)
      expect(graph.isSelectionEmpty()).toBe(true)

      autoLayout(graph, 'horizontal')

      expect(source.geometry?.x).toBeLessThan(middle.geometry?.x ?? 0)
      expect(middle.geometry?.x).toBeLessThan(target.geometry?.x ?? 0)
    } finally {
      destroy()
    }
  })

  it('keeps sibling branches on the same flow rank', () => {
    const { graph, destroy } = createDiagramGraph(container)
    try {
      const source = addRectangle(graph, { label: 'Source' })
      const branchA = addRectangle(graph, { label: 'Branch A' })
      const branchB = addRectangle(graph, { label: 'Branch B' })
      const target = addRectangle(graph, { label: 'Target' })
      connectCells(graph, source, branchA)
      connectCells(graph, source, branchB)
      connectCells(graph, branchA, target)
      connectCells(graph, branchB, target)

      autoLayout(graph, 'vertical')

      expect(source.geometry?.y).toBeLessThan(branchA.geometry?.y ?? 0)
      expect(branchA.geometry?.y).toBe(branchB.geometry?.y)
      expect(branchA.geometry?.y).toBeLessThan(target.geometry?.y ?? 0)
      expect(branchA.geometry?.x).not.toBe(branchB.geometry?.x)
      expect(graph.isSelectionEmpty()).toBe(true)
    } finally {
      destroy()
    }
  })
  it('connects exactly two selected vertices', () => {
    const { graph, destroy } = createDiagramGraph(container)
    try {
      const source = addRectangle(graph, { label: 'Source', x: 0, y: 0 })
      const target = addRectangle(graph, { label: 'Target', x: 240, y: 0 })
      graph.setSelectionCells([source, target])

      const edge = connectSelection(graph, 'links')

      expect(edge?.edge).toBeTruthy()
      expect(edge?.value).toBe('links')
      expect(graph.getChildEdges(graph.getDefaultParent())).toHaveLength(1)

      graph.setSelectionCells([source])
      expect(connectSelection(graph)).toBeNull()
      expect(graph.getChildEdges(graph.getDefaultParent())).toHaveLength(1)
    } finally {
      destroy()
    }
  })

  it('aligns and distributes selected vertices', () => {
    const { graph, destroy } = createDiagramGraph(container)
    try {
      const left = addRectangle(graph, { x: 20, y: 20, width: 80, height: 40 })
      const middle = addRectangle(graph, { x: 180, y: 90, width: 100, height: 50 })
      const right = addRectangle(graph, { x: 420, y: 160, width: 60, height: 30 })

      graph.setSelectionCells([left, middle, right])
      alignSelection(graph, 'top')
      expect(left.geometry?.y).toBe(20)
      expect(middle.geometry?.y).toBe(20)
      expect(right.geometry?.y).toBe(20)

      distributeSelection(graph, 'horizontal')
      expect(left.geometry?.x).toBe(20)
      expect(middle.geometry?.x).toBe(210)
      expect(right.geometry?.x).toBe(420)
    } finally {
      destroy()
    }
  })

  it('duplicates selected cells with offset', () => {
    const { graph, destroy } = createDiagramGraph(container)
    try {
      const v1 = addRectangle(graph, { label: 'Original', x: 50, y: 50 })
      graph.setSelectionCell(v1)

      const clones = duplicateSelection(graph)
      expect(clones.length).toBe(1)
      expect(clones[0].value).toBe('Original')
      expect(clones[0].geometry?.x).toBe(70)
      expect(clones[0].geometry?.y).toBe(70)
      expect(clones[0].id).not.toBe(v1.id)
    } finally {
      destroy()
    }
  })

  it('deletes selected cells', () => {
    const { graph, destroy } = createDiagramGraph(container)
    try {
      const v1 = addRectangle(graph, { label: 'ToDelete' })
      graph.setSelectionCell(v1)
      expect(graph.getChildVertices(graph.getDefaultParent()).length).toBe(1)

      deleteSelection(graph)
      expect(graph.getChildVertices(graph.getDefaultParent()).length).toBe(0)
    } finally {
      destroy()
    }
  })

  it('supports undo and redo', () => {
    const { graph, undoManager, destroy } = createDiagramGraph(container)
    try {
      expect(undoManager.canUndo()).toBe(false)
      expect(undoManager.canRedo()).toBe(false)

      const v1 = addRectangle(graph, { label: 'Step A' })
      expect(undoManager.canUndo()).toBe(true)

      undo(undoManager)
      expect(graph.getChildVertices(graph.getDefaultParent()).length).toBe(0)
      expect(undoManager.canRedo()).toBe(true)

      redo(undoManager)
      expect(graph.getChildVertices(graph.getDefaultParent()).length).toBe(1)
      expect(graph.getChildVertices(graph.getDefaultParent())[0].value).toBe(v1.value)
    } finally {
      destroy()
    }
  })

  it('pans with wheel and zooms with Ctrl/Cmd + wheel', () => {
    const { graph, destroy } = createDiagramGraph(container)
    try {
      const zoomEvent = new WheelEvent('wheel', {
        deltaY: -100,
        ctrlKey: true,
        cancelable: true,
      })
      container.dispatchEvent(zoomEvent)

      expect(graph.getView().scale).toBeCloseTo(1.2)
      expect(zoomEvent.defaultPrevented).toBe(true)

      const startY = graph.getView().translate.y
      const panEvent = new WheelEvent('wheel', {
        deltaY: 40,
        cancelable: true,
      })
      container.dispatchEvent(panEvent)

      expect(graph.getView().translate.y).toBeLessThan(startY)
      expect(panEvent.defaultPrevented).toBe(true)
    } finally {
      destroy()
    }
  })
})
