import {
  Graph,
  InternalEvent,
  type Cell,
  type ConnectionHandler,
  type PanningHandler,
  RubberBandHandler,
  UndoManager,
  getDefaultPlugins,
  registerAllCodecs,
  registerDefaultEdgeStyles,
  registerDefaultShapes,
} from '@maxgraph/core'

export interface GraphInstance {
  graph: Graph
  undoManager: UndoManager
  destroy: () => void
}

export type EmptyConnectionHandler = (source: Cell, clientX: number, clientY: number) => void

let registered = false
function ensureRegistered(): void {
  if (registered) return
  registered = true
  registerAllCodecs()
  registerDefaultShapes()
  registerDefaultEdgeStyles()
}

export function createDiagramGraph(
  container: HTMLElement,
  onEmptyConnection?: EmptyConnectionHandler,
): GraphInstance {
  ensureRegistered()

  // Initialize with all standard editor handlers (Panning, Connection, Selection, etc.)
  const graph = new Graph(container, undefined, getDefaultPlugins())

  // Canvas interaction rules
  graph.setConnectable(true)
  graph.setCellsEditable(true)
  graph.setCellsResizable(true)
  graph.setCellsMovable(true)
  graph.setAllowDanglingEdges(false)
  graph.setDisconnectOnMove(false)
  graph.setDropEnabled(true)
  graph.setSplitEnabled(false)

  const connectionHandler = graph.getPlugin<ConnectionHandler>('ConnectionHandler')
  if (connectionHandler && onEmptyConnection) {
    const connect = connectionHandler.connect.bind(connectionHandler)
    connectionHandler.connect = (source, target, event, dropTarget) => {
      if (source && !target && !dropTarget) {
        onEmptyConnection(source, event.clientX, event.clientY)
        return
      }
      connect(source, target, event, dropTarget)
    }
  }

  // Grid and snapping
  graph.setGridEnabled(true)
  graph.setGridSize(10)

  // Pan by dragging blank canvas space with the left mouse button.
  graph.setPanning(true)
  const panning = graph.getPlugin<PanningHandler>('PanningHandler')
  if (panning) {
    panning.useLeftButtonForPanning = true
    panning.panningEnabled = true
  }

  // Wheel scroll pans; Ctrl/Cmd + wheel zooms like draw.io.
  const wheelListener = (event: WheelEvent) => {
    const zooming = event.ctrlKey || event.metaKey
    const delta = event.deltaY || event.deltaX
    if (delta === 0) return

    event.preventDefault()
    if (zooming) {
      if (delta < 0) graph.zoomIn()
      else graph.zoomOut()
      return
    }

    const scale = graph.getView().scale || 1
    const translate = graph.getView().translate
    graph
      .getView()
      .setTranslate(translate.x - event.deltaX / scale, translate.y - event.deltaY / scale)
  }
  container.addEventListener('wheel', wheelListener, { passive: false })

  // Rubberband marquee multi-selection
  const rubberband = new RubberBandHandler(graph)

  // Style defaults matching ARAOffice modern palette
  const defaultVertexStyle = graph.getStylesheet().getDefaultVertexStyle()
  defaultVertexStyle.shape = 'rectangle'
  defaultVertexStyle.rounded = true
  defaultVertexStyle.arcSize = 8
  defaultVertexStyle.fillColor = '#ffffff'
  defaultVertexStyle.strokeColor = '#334155'
  defaultVertexStyle.strokeWidth = 2
  defaultVertexStyle.fontColor = '#0f172a'
  defaultVertexStyle.fontSize = 13
  defaultVertexStyle.fontFamily =
    '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", sans-serif'

  const defaultEdgeStyle = graph.getStylesheet().getDefaultEdgeStyle()
  defaultEdgeStyle.edgeStyle = 'orthogonalEdgeStyle'
  defaultEdgeStyle.rounded = true
  defaultEdgeStyle.strokeColor = '#475569'
  defaultEdgeStyle.strokeWidth = 2
  defaultEdgeStyle.endArrow = 'classic'
  defaultEdgeStyle.fontFamily = defaultVertexStyle.fontFamily
  defaultEdgeStyle.fontSize = 12
  defaultEdgeStyle.fontColor = '#334155'

  // ponytail: cap undo history to 40 edits to avoid unbounded memory retention
  const undoManager = new UndoManager(40)
  const undoListener = (_sender: unknown, evt: { getProperty: (key: string) => unknown }) => {
    const edit = evt.getProperty('edit')
    if (edit) {
      undoManager.undoableEditHappened(edit as Parameters<UndoManager['undoableEditHappened']>[0])
    }
  }

  graph.getDataModel().addListener(InternalEvent.UNDO, undoListener)

  const destroy = () => {
    try {
      graph.getDataModel().removeListener(undoListener)
      container.removeEventListener('wheel', wheelListener)
      rubberband.onDestroy()
      // ponytail: release undo history and cell models on unmount to prevent OOM
      undoManager.clear()
      graph.getDataModel().clear()
      graph.destroy()
    } catch {
      // ignore cleanup errors during unmount
    }
  }

  return { graph, undoManager, destroy }
}
