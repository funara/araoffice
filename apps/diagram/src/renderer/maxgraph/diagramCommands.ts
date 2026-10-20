import { HierarchicalLayout } from '@maxgraph/core'
import type { Cell, CellStateStyle, CellStyle, FitPlugin, Graph, UndoManager } from '@maxgraph/core'
export type DiagramShapeType =
  | 'rectangle'
  | 'rounded'
  | 'ellipse'
  | 'doubleEllipse'
  | 'diamond'
  | 'cylinder'
  | 'text'
  | 'triangle'
  | 'hexagon'
  | 'cloud'

export const DIAGRAM_SHAPE_DRAG_MIME = 'application/x-araoffice-diagram-shape'

const DIAGRAM_SHAPE_TYPES: readonly DiagramShapeType[] = [
  'rectangle',
  'rounded',
  'ellipse',
  'doubleEllipse',
  'diamond',
  'cylinder',
  'text',
  'triangle',
  'hexagon',
  'cloud',
]

export function isDiagramShapeType(value: unknown): value is DiagramShapeType {
  return typeof value === 'string' && DIAGRAM_SHAPE_TYPES.includes(value as DiagramShapeType)
}

export type DiagramShapeOptions = {
  x?: number
  y?: number
  width?: number
  height?: number
  label?: string
  style?: CellStyle
}

export type AlignDirection = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'
export type DistributeDirection = 'horizontal' | 'vertical'
export interface DiagramSelectionCommands {
  onDuplicate: () => void
  onDelete: () => void
  onSelectAll: () => void
  onConnect: () => void
  onAlign: (direction: AlignDirection) => void
  onDistribute: (direction: DistributeDirection) => void
  onOrder: (direction: 'front' | 'back') => void
}

/** Default sizes for initial shapes */
const SHAPE_DEFAULTS: Record<DiagramShapeType, { width: number; height: number }> = {
  rectangle: { width: 120, height: 60 },
  rounded: { width: 120, height: 60 },
  ellipse: { width: 100, height: 80 },
  doubleEllipse: { width: 100, height: 80 },
  diamond: { width: 100, height: 80 },
  cylinder: { width: 90, height: 80 },
  text: { width: 100, height: 40 },
  triangle: { width: 100, height: 80 },
  hexagon: { width: 120, height: 80 },
  cloud: { width: 130, height: 80 },
}

const SHAPE_FILL = '#ffffff'
const SHAPE_STROKE = '#334155'
const SHAPE_STROKE_WIDTH = 2

const SHAPE_PRESETS: Record<DiagramShapeType, { label: string; style: CellStyle }> = {
  rectangle: {
    label: 'Rectangle',
    style: {
      shape: 'rectangle',
      rounded: false,
      fillColor: SHAPE_FILL,
      strokeColor: SHAPE_STROKE,
      strokeWidth: SHAPE_STROKE_WIDTH,
    },
  },
  rounded: {
    label: 'Process',
    style: {
      shape: 'rectangle',
      rounded: true,
      arcSize: 12,
      fillColor: SHAPE_FILL,
      strokeColor: SHAPE_STROKE,
      strokeWidth: SHAPE_STROKE_WIDTH,
    },
  },
  ellipse: {
    label: 'Start / End',
    style: {
      shape: 'ellipse',
      perimeter: 'ellipsePerimeter',
      fillColor: SHAPE_FILL,
      strokeColor: SHAPE_STROKE,
      strokeWidth: SHAPE_STROKE_WIDTH,
    },
  },
  doubleEllipse: {
    label: 'Double ellipse',
    style: {
      shape: 'doubleEllipse',
      perimeter: 'ellipsePerimeter',
      fillColor: SHAPE_FILL,
      strokeColor: SHAPE_STROKE,
      strokeWidth: SHAPE_STROKE_WIDTH,
    },
  },
  diamond: {
    label: 'Decision',
    style: {
      shape: 'rhombus',
      perimeter: 'rhombusPerimeter',
      fillColor: SHAPE_FILL,
      strokeColor: SHAPE_STROKE,
      strokeWidth: SHAPE_STROKE_WIDTH,
    },
  },
  cylinder: {
    label: 'Database',
    style: {
      shape: 'cylinder',
      fillColor: SHAPE_FILL,
      strokeColor: SHAPE_STROKE,
      strokeWidth: SHAPE_STROKE_WIDTH,
    },
  },
  text: {
    label: 'Text',
    style: {
      shape: 'text',
      strokeColor: 'none',
      fillColor: 'none',
      align: 'center',
      verticalAlign: 'middle',
      fontSize: 14,
      fontColor: '#0f172a',
    },
  },
  triangle: {
    label: 'Triangle',
    style: {
      shape: 'triangle',
      perimeter: 'trianglePerimeter',
      direction: 'south',
      fillColor: SHAPE_FILL,
      strokeColor: SHAPE_STROKE,
      strokeWidth: SHAPE_STROKE_WIDTH,
    },
  },
  hexagon: {
    label: 'Hexagon',
    style: {
      shape: 'hexagon',
      perimeter: 'hexagonPerimeter',
      fillColor: SHAPE_FILL,
      strokeColor: SHAPE_STROKE,
      strokeWidth: SHAPE_STROKE_WIDTH,
    },
  },
  cloud: {
    label: 'Cloud',
    style: {
      shape: 'cloud',
      fillColor: SHAPE_FILL,
      strokeColor: SHAPE_STROKE,
      strokeWidth: SHAPE_STROKE_WIDTH,
    },
  },
}

function getInsertionCoordinates(graph: Graph): { x: number; y: number } {
  const container = graph.container
  const scrollLeft = container.scrollLeft || 0
  const scrollTop = container.scrollTop || 0
  const width = container.clientWidth || 800
  const height = container.clientHeight || 600

  const scale = graph.getView().scale || 1
  const translate = graph.getView().translate || { x: 0, y: 0 }

  // Place in the visible center of the canvas with subtle jitter so multiple clicks don't stack directly
  const jitterX = Math.floor(Math.random() * 5) * 20
  const jitterY = Math.floor(Math.random() * 4) * 20
  const centerX = (scrollLeft + width / 2) / scale - translate.x - 60 + jitterX
  const centerY = (scrollTop + height / 2) / scale - translate.y - 30 + jitterY

  return {
    x: Math.round(centerX / 10) * 10,
    y: Math.round(centerY / 10) * 10,
  }
}

function insertVertex(
  graph: Graph,
  defaultSize: { width: number; height: number },
  defaultLabel: string,
  defaultStyle: CellStyle,
  opts?: DiagramShapeOptions,
): Cell {
  const coords = getInsertionCoordinates(graph)
  const x = opts?.x ?? coords.x
  const y = opts?.y ?? coords.y
  const width = opts?.width ?? defaultSize.width
  const height = opts?.height ?? defaultSize.height
  const label = opts?.label ?? defaultLabel
  const style: CellStyle = { ...defaultStyle, ...opts?.style }

  let cell!: Cell
  graph.batchUpdate(() => {
    cell = graph.insertVertex(graph.getDefaultParent(), null, label, x, y, width, height, style)
    graph.setSelectionCell(cell)
  })
  return cell
}

export function addRectangle(graph: Graph, opts?: DiagramShapeOptions): Cell {
  return addShape(graph, 'rectangle', opts)
}

export function addRoundedRectangle(graph: Graph, opts?: DiagramShapeOptions): Cell {
  return addShape(graph, 'rounded', opts)
}

export function addEllipse(graph: Graph, opts?: DiagramShapeOptions): Cell {
  return addShape(graph, 'ellipse', opts)
}
export function addDoubleEllipse(graph: Graph, opts?: DiagramShapeOptions): Cell {
  return addShape(graph, 'doubleEllipse', opts)
}

export function addDiamond(graph: Graph, opts?: DiagramShapeOptions): Cell {
  return addShape(graph, 'diamond', opts)
}

export function addCylinder(graph: Graph, opts?: DiagramShapeOptions): Cell {
  return addShape(graph, 'cylinder', opts)
}

export function addText(graph: Graph, opts?: DiagramShapeOptions): Cell {
  return addShape(graph, 'text', opts)
}

export function addTriangle(graph: Graph, opts?: DiagramShapeOptions): Cell {
  return addShape(graph, 'triangle', opts)
}

export function addHexagon(graph: Graph, opts?: DiagramShapeOptions): Cell {
  return addShape(graph, 'hexagon', opts)
}

export function addCloud(graph: Graph, opts?: DiagramShapeOptions): Cell {
  return addShape(graph, 'cloud', opts)
}

export function addShape(graph: Graph, type: DiagramShapeType, opts?: DiagramShapeOptions): Cell {
  const preset = SHAPE_PRESETS[type]
  return insertVertex(graph, SHAPE_DEFAULTS[type], preset.label, preset.style, opts)
}

export function addShapeAtPoint(graph: Graph, type: DiagramShapeType, x: number, y: number): Cell {
  const { width, height } = SHAPE_DEFAULTS[type]
  return addShape(graph, type, { x: x - width / 2, y: y - height / 2 })
}

export function connectCells(
  graph: Graph,
  source: Cell,
  target: Cell,
  label = '',
  style?: CellStyle,
): Cell {
  let edge!: Cell
  graph.batchUpdate(() => {
    edge = graph.insertEdge(graph.getDefaultParent(), null, label, source, target, {
      edgeStyle: 'orthogonalEdgeStyle',
      rounded: true,
      strokeColor: '#475569',
      strokeWidth: 2,
      endArrow: 'classic',
      labelBackgroundColor: '#ffffff',
      labelBorderColor: '#cbd5e1',
      labelPadding: 4,
      ...style,
    })
    graph.setSelectionCell(edge)
  })
  return edge
}

export type DiagramLayoutDirection = 'horizontal' | 'vertical' | 'grid'

/**
 * Arrange a connected diagram by graph topology instead of insertion order.
 * HierarchicalLayout supplies separated ranks and explicit edge waypoints so
 * connectors do not run through unrelated vertices.
 */
export function autoLayout(graph: Graph, direction: DiagramLayoutDirection = 'vertical'): void {
  const parent = graph.getDefaultParent()
  const vertices = graph.getChildVertices(parent)
  const edges = graph.getChildEdges(parent)

  if (vertices.length < 2 || edges.length === 0) {
    graph.clearSelection()
    return
  }

  const layout = new HierarchicalLayout(graph, direction === 'horizontal' ? 'west' : 'north')
  layout.intraCellSpacing = 60
  layout.interRankCellSpacing = 90
  layout.interHierarchySpacing = 80
  layout.parallelEdgeSpacing = 24
  layout.fineTuning = true
  layout.execute(parent)
  graph.clearSelection()
  graph.refresh()
}
export function connectSelection(graph: Graph, label = ''): Cell | null {
  const cells = graph.getSelectionCells()
  if (cells.length !== 2 || !cells.every((cell) => cell.isVertex())) return null
  return connectCells(graph, cells[0], cells[1], label)
}

type VertexWithGeometry = Cell & { geometry: NonNullable<Cell['geometry']> }

function selectedVertices(graph: Graph): VertexWithGeometry[] {
  return graph
    .getSelectionCells()
    .filter((cell): cell is VertexWithGeometry => cell.isVertex() && Boolean(cell.geometry))
}

export function alignSelection(graph: Graph, direction: AlignDirection): void {
  const cells = selectedVertices(graph)
  if (cells.length < 2) return

  const geometries = cells.map((cell) => cell.geometry)
  const left = Math.min(...geometries.map((geometry) => geometry.x))
  const right = Math.max(...geometries.map((geometry) => geometry.x + geometry.width))
  const top = Math.min(...geometries.map((geometry) => geometry.y))
  const bottom = Math.max(...geometries.map((geometry) => geometry.y + geometry.height))
  const centerX = (left + right) / 2
  const centerY = (top + bottom) / 2

  graph.batchUpdate(() => {
    for (const cell of cells) {
      const geometry = cell.geometry.clone()
      switch (direction) {
        case 'left':
          geometry.x = left
          break
        case 'center':
          geometry.x = centerX - geometry.width / 2
          break
        case 'right':
          geometry.x = right - geometry.width
          break
        case 'top':
          geometry.y = top
          break
        case 'middle':
          geometry.y = centerY - geometry.height / 2
          break
        case 'bottom':
          geometry.y = bottom - geometry.height
          break
        default:
          break
      }
      graph.getDataModel().setGeometry(cell, geometry)
    }
  })
}

export function distributeSelection(graph: Graph, direction: DistributeDirection): void {
  const cells = selectedVertices(graph)
  if (cells.length < 3) return

  const sorted = [...cells].sort((a, b) => {
    const aGeometry = a.geometry
    const bGeometry = b.geometry
    return direction === 'horizontal' ? aGeometry.x - bGeometry.x : aGeometry.y - bGeometry.y
  })
  const first = sorted[0].geometry
  const last = sorted[sorted.length - 1].geometry
  const totalSize = sorted.reduce(
    (sum, cell) => sum + (direction === 'horizontal' ? cell.geometry.width : cell.geometry.height),
    0,
  )
  const span =
    direction === 'horizontal' ? last.x + last.width - first.x : last.y + last.height - first.y
  const gap = Math.max(0, (span - totalSize) / (sorted.length - 1))
  let position = direction === 'horizontal' ? first.x : first.y

  graph.batchUpdate(() => {
    for (const cell of sorted) {
      const geometry = cell.geometry.clone()
      if (direction === 'horizontal') {
        geometry.x = position
        position += geometry.width + gap
      } else {
        geometry.y = position
        position += geometry.height + gap
      }
      graph.getDataModel().setGeometry(cell, geometry)
    }
  })
}

export function orderSelection(graph: Graph, direction: 'front' | 'back'): void {
  const cells = graph.getSelectionCells()
  if (cells.length > 0) {
    graph.orderCells(direction === 'back', cells)
  }
}

export function deleteSelection(graph: Graph): void {
  const cells = graph.getSelectionCells()
  if (cells.length > 0) {
    graph.removeCells(cells)
  }
}

export function duplicateSelection(graph: Graph): Cell[] {
  const cells = graph.getSelectionCells()
  if (cells.length === 0) return []

  const clones = graph.cloneCells(cells)
  graph.batchUpdate(() => {
    for (const clone of clones) {
      if (clone.geometry) {
        clone.geometry.x = (clone.geometry.x || 0) + 20
        clone.geometry.y = (clone.geometry.y || 0) + 20
      }
      graph.addCell(clone, graph.getDefaultParent())
    }
  })
  graph.setSelectionCells(clones)
  return clones
}

export function selectAll(graph: Graph): void {
  graph.selectAll()
}

export function zoomIn(graph: Graph): void {
  graph.zoomIn()
}

export function zoomOut(graph: Graph): void {
  graph.zoomOut()
}

export function zoomActual(graph: Graph): void {
  graph.zoomActual()
}

export function zoomFit(graph: Graph): void {
  const fitPlugin = graph.getPlugin<FitPlugin>('fit')
  if (fitPlugin) {
    fitPlugin.fit()
  }
}

export function setCellLabel(graph: Graph, cell: Cell, label: string): void {
  graph.model.setValue(cell, label)
}

export function setSelectionStyle<K extends keyof CellStateStyle>(
  graph: Graph,
  key: K,
  value: CellStateStyle[K],
): void {
  const cells = graph.getSelectionCells()
  if (cells.length === 0) return
  graph.setCellStyles(key, value as string | number | boolean | null, cells)
}

export function undo(undoManager: UndoManager): void {
  if (undoManager.canUndo()) {
    undoManager.undo()
  }
}

export function redo(undoManager: UndoManager): void {
  if (undoManager.canRedo()) {
    undoManager.redo()
  }
}
