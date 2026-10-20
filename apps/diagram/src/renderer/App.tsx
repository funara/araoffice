import { useCallback, useEffect, useMemo, useRef, useState } from 'react'

import type { Dispatch, DragEvent, ReactElement, SetStateAction } from 'react'
import type { Cell, CellStateStyle } from '@maxgraph/core'
import { htmlLang, type Lang } from '@araoffice/i18n'
import { createDiagramGraph, type GraphInstance } from './maxgraph/createGraph'
import {
  addShape,
  addShapeAtPoint,
  autoLayout,
  DIAGRAM_SHAPE_DRAG_MIME,
  alignSelection,
  connectCells,
  connectSelection,
  deleteSelection,
  distributeSelection,
  duplicateSelection,
  isDiagramShapeType,
  orderSelection,
  redo,
  selectAll,
  setCellLabel,
  setSelectionStyle,
  undo,
  zoomActual,
  zoomFit,
  zoomIn,
  zoomOut,
} from './maxgraph/diagramCommands'
import type {
  AlignDirection,
  DiagramSelectionCommands,
  DiagramShapeType,
  DistributeDirection,
} from './maxgraph/diagramCommands'
import { loadDrawioIntoGraph, serializeGraphToDrawio } from './maxgraph/drawio'
import { quickAddCopy, shapeHelp } from './maxgraph/shape-help'
import {
  exportGraphToRaster,
  exportGraphToSvg,
  svgStringToRasterDataUrl,
} from './maxgraph/exportDiagram'
import { Ribbon } from './components/Ribbon'
import { IconProperties } from './components/icons'
import { PropertiesPanel } from './components/PropertiesPanel'
import { StatusBar } from './components/StatusBar'
import { clampMermaidZoom, MermaidEditor, MERMAID_ZOOM_STEP } from './mermaid/MermaidEditor'
import { renderMermaidDiagram } from './mermaid/renderMermaid'
import { AraSpark } from '@araoffice/ui'
import { AiPanel, type DiagramSnapshot } from './ai/AiPanel'
import type { DiagramSkillDeps } from './ai/tools'
import type { DiagramMode, ExportFormat, PendingDiagram, SaveMode } from '../shared/ipc'
import './styles.css'

const DEFAULT_MERMAID = `flowchart TD
    Start([Start]) --> Process[Step 1: Execute Task]
    Process --> Choice{Is Success?}
    Choice -- Yes --> Finish([Finish])
    Choice -- No --> Retry[Step 2: Retry]
    Retry --> Process`

type PendingConnection = {
  source: Cell
  clientX: number
  clientY: number
}

const QUICK_CONNECT_SHAPES: readonly DiagramShapeType[] = [
  'rounded',
  'diamond',
  'ellipse',
  'cylinder',
  'rectangle',
  'hexagon',
  'cloud',
  'triangle',
  'doubleEllipse',
  'text',
]

type SelectionType = 'vertex' | 'edge' | 'mixed' | undefined

function getCellType(cell: Cell): Exclude<SelectionType, 'mixed' | undefined> | undefined {
  if (cell.isVertex()) return 'vertex'
  if (cell.isEdge()) return 'edge'
  return undefined
}

function getSelectionType(cells: Cell[]): SelectionType {
  if (cells.length === 0) return undefined
  if (cells.length === 1) return getCellType(cells[0])

  const hasVertex = cells.some((cell) => cell.isVertex())
  const hasEdge = cells.some((cell) => cell.isEdge())
  if (hasVertex && hasEdge) return 'mixed'
  if (hasVertex) return 'vertex'
  return 'edge'
}

function createSelectionListener(
  graph: GraphInstance['graph'],
  setSelectedCount: (count: number) => void,
  setSelectedCell: (cell: Cell | null) => void,
  setSelectedType: (type: SelectionType) => void,
) {
  return () => {
    const cells = graph.getSelectionCells()
    setSelectedCount(cells.length)
    setSelectedCell(cells[0] ?? null)
    setSelectedType(getSelectionType(cells))
  }
}

const IS_MAC_PLATFORM = /mac/i.test(navigator.userAgent)

interface DiagramKeyDownOptions {
  graph: GraphInstance['graph']
  undoManager: GraphInstance['undoManager']
  setCanUndo: (value: boolean) => void
  setCanRedo: (value: boolean) => void
  setZoomPercent: (value: number) => void
}

function syncHistory(
  undoManager: GraphInstance['undoManager'],
  setCanUndo: (value: boolean) => void,
  setCanRedo: (value: boolean) => void,
) {
  setCanUndo(undoManager.canUndo())
  setCanRedo(undoManager.canRedo())
}

function handleZoomShortcut(
  event: KeyboardEvent,
  key: string,
  graph: GraphInstance['graph'],
  setZoomPercent: (value: number) => void,
) {
  if (key === '=' || key === '+') {
    event.preventDefault()
    zoomIn(graph)
    setZoomPercent((graph.getView().scale || 1) * 100)
    return true
  }
  if (key === '-') {
    event.preventDefault()
    zoomOut(graph)
    setZoomPercent((graph.getView().scale || 1) * 100)
    return true
  }
  if (key === '0') {
    event.preventDefault()
    zoomActual(graph)
    setZoomPercent(100)
    return true
  }
  return false
}

function handleDiagramKeyDown(event: KeyboardEvent, options: DiagramKeyDownOptions) {
  const { graph, undoManager, setCanUndo, setCanRedo, setZoomPercent } = options
  const target = event.target as HTMLElement | null
  if (target?.tagName === 'INPUT' || target?.tagName === 'TEXTAREA' || target?.isContentEditable) {
    return
  }

  if (event.key === 'Delete' || event.key === 'Backspace') {
    event.preventDefault()
    deleteSelection(graph)
    return
  }
  if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) {
    if (!graph.isEditing()) {
      const cells = graph.getSelectionCells()
      if (cells && cells.length > 0) {
        event.preventDefault()
        const step = event.altKey ? 1 : event.shiftKey ? 20 : graph.gridSize || 10
        let dx = 0
        let dy = 0
        if (event.key === 'ArrowUp') dy = -step
        else if (event.key === 'ArrowDown') dy = step
        else if (event.key === 'ArrowLeft') dx = -step
        else if (event.key === 'ArrowRight') dx = step
        graph.moveCells(cells, dx, dy)
        syncHistory(undoManager, setCanUndo, setCanRedo)
        return
      }
    }
  }

  const mod = IS_MAC_PLATFORM ? event.metaKey : event.ctrlKey
  if (!mod) return

  const key = event.key.toLowerCase()
  if (key === 'z' || key === 'y') {
    event.preventDefault()
    if (key === 'z' && !event.shiftKey) {
      undo(undoManager)
    } else {
      redo(undoManager)
    }
    syncHistory(undoManager, setCanUndo, setCanRedo)
    return
  }

  if (key === 'a') {
    event.preventDefault()
    selectAll(graph)
    return
  }

  if (key === 'd') {
    event.preventDefault()
    duplicateSelection(graph)
    return
  }

  handleZoomShortcut(event, key, graph, setZoomPercent)
}

interface DiagramCommandOptions {
  mode: DiagramMode
  graphInstance: GraphInstance | null
  graphRef: { current: GraphInstance | null }
  setZoomPercent: Dispatch<SetStateAction<number>>
  setDirty: (value: boolean) => void
  setMermaidCode: (value: string) => void
  setSnapshots: (value: DiagramSnapshot[]) => void
}

interface DiagramCommands extends DiagramZoomCommands, DiagramEditCommands {
  graph: GraphInstance['graph'] | null
}

interface DiagramZoomCommands {
  onZoomIn: () => void
  onZoomOut: () => void
  onZoomActual: () => void
  onZoomFit: () => void
  onZoomChange: (value: number) => void
}

function useDiagramZoomCommands(
  mode: DiagramMode,
  getGraph: () => GraphInstance['graph'] | null,
  setZoomPercent: Dispatch<SetStateAction<number>>,
): DiagramZoomCommands {
  const runZoom = useCallback(
    (command: (graph: GraphInstance['graph']) => void, fixedPercent?: number) => {
      const graph = getGraph()
      if (!graph) return
      command(graph)
      setZoomPercent(fixedPercent ?? (graph.getView().scale || 1) * 100)
    },
    [getGraph, setZoomPercent],
  )
  const updateMermaidZoom = useCallback(
    (delta: number) => {
      setZoomPercent((current) => clampMermaidZoom(current + delta))
    },
    [setZoomPercent],
  )
  const onZoomIn = useCallback(() => {
    if (mode === 'mermaid') {
      updateMermaidZoom(MERMAID_ZOOM_STEP)
      return
    }
    runZoom(zoomIn)
  }, [mode, runZoom, updateMermaidZoom])
  const onZoomOut = useCallback(() => {
    if (mode === 'mermaid') {
      updateMermaidZoom(-MERMAID_ZOOM_STEP)
      return
    }
    runZoom(zoomOut)
  }, [mode, runZoom, updateMermaidZoom])
  const onZoomActual = useCallback(() => {
    if (mode === 'mermaid') {
      setZoomPercent(100)
      return
    }
    runZoom(zoomActual, 100)
  }, [mode, runZoom, setZoomPercent])
  const onZoomFit = useCallback(() => {
    if (mode === 'mermaid') {
      setZoomPercent(100)
      return
    }
    runZoom(zoomFit)
  }, [mode, runZoom, setZoomPercent])
  const onZoomChange = useCallback(
    (value: number) => {
      if (mode === 'mermaid') {
        setZoomPercent(clampMermaidZoom(value))
        return
      }
      const graph = getGraph()
      if (!graph) return
      graph.zoomTo(value / 100)
      setZoomPercent(value)
    },
    [getGraph, mode, setZoomPercent],
  )

  return { onZoomIn, onZoomOut, onZoomActual, onZoomFit, onZoomChange }
}

interface DiagramEditCommands extends DiagramSelectionCommands {
  handleInsertShape: (type: DiagramShapeType) => void
  handleDropShape: (type: DiagramShapeType, clientX: number, clientY: number) => Cell | null
  handleClear: () => void
  onUndo: () => void
  onRedo: () => void
}

interface DiagramEditCommandOptions {
  mode: DiagramMode
  graphRef: { current: GraphInstance | null }
  getGraph: () => GraphInstance['graph'] | null
  getUndoManager: () => GraphInstance['undoManager'] | null
  setDirty: (value: boolean) => void
  setMermaidCode: (value: string) => void
  setSnapshots: (value: DiagramSnapshot[]) => void
}

function useDiagramEditCommands({
  mode,
  graphRef,
  getGraph,
  getUndoManager,
  setDirty,
  setMermaidCode,
  setSnapshots,
}: DiagramEditCommandOptions): DiagramEditCommands {
  const runGraphCommand = useCallback(
    (command: (graph: GraphInstance['graph']) => void) => {
      const graph = getGraph()
      if (graph) command(graph)
    },
    [getGraph],
  )
  const handleInsertShape = useCallback(
    (type: DiagramShapeType) => {
      const graph = graphRef.current?.graph
      if (!graph) return
      addShape(graph, type)
      setDirty(true)
      window.diagramApi?.setDirty(true)
    },
    [graphRef, setDirty],
  )
  const handleDropShape = useCallback(
    (type: DiagramShapeType, clientX: number, clientY: number): Cell | null => {
      const graph = graphRef.current?.graph
      if (!graph) return null

      const bounds = graph.container.getBoundingClientRect()
      const scale = graph.getView().scale || 1
      const translate = graph.getView().translate || { x: 0, y: 0 }
      const x = (clientX - bounds.left + graph.container.scrollLeft) / scale - translate.x
      const y = (clientY - bounds.top + graph.container.scrollTop) / scale - translate.y

      const cell = addShapeAtPoint(graph, type, x, y)
      setDirty(true)
      window.diagramApi?.setDirty(true)
      return cell
    },
    [graphRef, setDirty],
  )

  const handleClear = useCallback(() => {
    const graph = getGraph()
    if (mode === 'visual' && graph) {
      graph.removeCells(graph.getChildCells(graph.getDefaultParent()))
      graphRef.current?.undoManager.clear()
    } else {
      setMermaidCode('flowchart TD\n  Start([Start])')
    }
    setSnapshots([])
    setDirty(true)
    window.diagramApi?.setDirty(true)
  }, [getGraph, graphRef, mode, setDirty, setMermaidCode, setSnapshots])
  const onUndo = useCallback(() => {
    const undoManager = getUndoManager()
    if (undoManager) undo(undoManager)
  }, [getUndoManager])
  const onRedo = useCallback(() => {
    const undoManager = getUndoManager()
    if (undoManager) redo(undoManager)
  }, [getUndoManager])
  const onDuplicate = useCallback(() => runGraphCommand(duplicateSelection), [runGraphCommand])
  const onDelete = useCallback(() => runGraphCommand(deleteSelection), [runGraphCommand])
  const onSelectAll = useCallback(() => runGraphCommand(selectAll), [runGraphCommand])
  const onConnect = useCallback(() => runGraphCommand(connectSelection), [runGraphCommand])
  const onAlign = useCallback(
    (direction: AlignDirection) => runGraphCommand((graph) => alignSelection(graph, direction)),
    [runGraphCommand],
  )
  const onDistribute = useCallback(
    (direction: DistributeDirection) =>
      runGraphCommand((graph) => distributeSelection(graph, direction)),
    [runGraphCommand],
  )
  const onOrder = useCallback(
    (direction: 'front' | 'back') => runGraphCommand((graph) => orderSelection(graph, direction)),
    [runGraphCommand],
  )

  return {
    handleInsertShape,
    handleDropShape,
    handleClear,
    onUndo,
    onRedo,
    onDuplicate,
    onDelete,
    onSelectAll,
    onConnect,
    onAlign,
    onDistribute,
    onOrder,
  }
}

function useDiagramCommands({
  mode,
  graphInstance,
  graphRef,
  setZoomPercent,
  setDirty,
  setMermaidCode,
  setSnapshots,
}: DiagramCommandOptions): DiagramCommands {
  const getGraph = useCallback(
    () => graphInstance?.graph ?? graphRef.current?.graph ?? null,
    [graphInstance, graphRef],
  )
  const getUndoManager = useCallback(
    () => graphInstance?.undoManager ?? graphRef.current?.undoManager ?? null,
    [graphInstance, graphRef],
  )
  const zoomCommands = useDiagramZoomCommands(mode, getGraph, setZoomPercent)
  const editCommands = useDiagramEditCommands({
    mode,
    graphRef,
    getGraph,
    getUndoManager,
    setDirty,
    setMermaidCode,
    setSnapshots,
  })

  return {
    graph: getGraph(),
    ...zoomCommands,
    ...editCommands,
  }
}

interface DiagramUiActionOptions {
  handleSave: (saveMode?: SaveMode) => Promise<boolean>
  aiOpen: boolean
  setAiOpen: Dispatch<SetStateAction<boolean>>
  setAiPresetPrompt: Dispatch<SetStateAction<string | null>>
  setShowPropsPanel: Dispatch<SetStateAction<boolean>>
  setGridEnabled: Dispatch<SetStateAction<boolean>>
  setMermaidCode: Dispatch<SetStateAction<string>>
  setDirty: Dispatch<SetStateAction<boolean>>
  setOperationError: Dispatch<SetStateAction<string | null>>
}

interface DiagramUiActions {
  onSave: () => void
  onSaveAs: () => void
  onToggleAi: () => void
  onAiPrompt: (prompt: string) => void
  onTogglePropsPanel: () => void
  onToggleGrid: () => void
  onApplyTemplate: (template: string) => void
  onInsertMermaidSnippet: (snippet: string) => void
  onMermaidChange: (code: string) => void
  onDismissOperationError: () => void
  onOpenAi: () => void
  onCollapseAi: () => void
  onCollapseProps: () => void
  onExpandProps: () => void
}

function useDiagramUiActions({
  handleSave,
  aiOpen,
  setAiOpen,
  setAiPresetPrompt,
  setShowPropsPanel,
  setGridEnabled,
  setMermaidCode,
  setDirty,
  setOperationError,
}: DiagramUiActionOptions): DiagramUiActions {
  const markMermaidDirty = useCallback(() => {
    setDirty(true)
    window.diagramApi?.setDirty(true)
  }, [setDirty])
  const onSave = useCallback(() => void handleSave('save'), [handleSave])
  const onSaveAs = useCallback(() => void handleSave('saveAs'), [handleSave])
  const onToggleAi = useCallback(() => setAiOpen((value) => !value), [setAiOpen])
  const onAiPrompt = useCallback(
    (prompt: string) => {
      setAiPresetPrompt(prompt)
      if (!aiOpen) setAiOpen(true)
    },
    [aiOpen, setAiOpen, setAiPresetPrompt],
  )
  const onTogglePropsPanel = useCallback(
    () => setShowPropsPanel((value) => !value),
    [setShowPropsPanel],
  )
  const onToggleGrid = useCallback(() => setGridEnabled((value) => !value), [setGridEnabled])
  const onApplyTemplate = useCallback(
    (template: string) => {
      setMermaidCode(template)
      markMermaidDirty()
    },
    [markMermaidDirty, setMermaidCode],
  )
  const onInsertMermaidSnippet = useCallback(
    (snippet: string) => {
      setMermaidCode((current) => {
        const base = current.trimEnd()
        return base ? `${base}\n${snippet}` : snippet
      })
      markMermaidDirty()
    },
    [markMermaidDirty, setMermaidCode],
  )
  const onMermaidChange = useCallback(
    (code: string) => {
      setMermaidCode(code)
      markMermaidDirty()
    },
    [markMermaidDirty, setMermaidCode],
  )

  return {
    onSave,
    onSaveAs,
    onToggleAi,
    onAiPrompt,
    onTogglePropsPanel,
    onToggleGrid,
    onApplyTemplate,
    onInsertMermaidSnippet,
    onMermaidChange,
    onDismissOperationError: () => setOperationError(null),
    onOpenAi: () => setAiOpen(true),
    onCollapseAi: () => setAiOpen(false),
    onCollapseProps: () => setShowPropsPanel(false),
    onExpandProps: () => setShowPropsPanel(true),
  }
}
interface DiagramGraphStateOptions {
  setCanUndo: (value: boolean) => void
  setCanRedo: (value: boolean) => void
  setDirty: (value: boolean) => void
  setSelectedCount: (value: number) => void
  setSelectedCell: (cell: Cell | null) => void
  setSelectedType: (type: SelectionType) => void
  setZoomPercent: (value: number) => void
}

interface DiagramGraphLifecycleOptions extends DiagramGraphStateOptions {
  initialLoadReady: boolean
  mode: DiagramMode
  gridEnabled: boolean
  canvasRef: { current: HTMLDivElement | null }
  graphRef: { current: GraphInstance | null }
  initialDrawioContentRef: { current: string | null }
  initialDrawioLoadStateRef: { current: 'ready' | 'loading' | 'failed' }
  graphInstance: GraphInstance | null
  setGraphInstance: (instance: GraphInstance | null) => void
  setOperationError: (value: string) => void
  onEmptyConnection: (source: Cell, clientX: number, clientY: number) => void
}

interface DiagramGraphListenerOptions extends DiagramGraphStateOptions {
  graph: GraphInstance['graph']
  undoManager: GraphInstance['undoManager']
}

function attachHistoryListener({
  graph,
  undoManager,
  setCanUndo,
  setCanRedo,
  setDirty,
}: DiagramGraphListenerOptions): () => void {
  const onChange = () => {
    setCanUndo(undoManager.canUndo())
    setCanRedo(undoManager.canRedo())
    setDirty(true)
    window.diagramApi?.setDirty(true)
  }
  const model = graph.getDataModel()
  model.addListener('change', onChange)
  return () => model.removeListener(onChange)
}

function attachSelectionListener({
  graph,
  setSelectedCount,
  setSelectedCell,
  setSelectedType,
}: DiagramGraphListenerOptions): () => void {
  const selectionModel = graph.getSelectionModel()
  const onSelectionChange = createSelectionListener(
    graph,
    setSelectedCount,
    setSelectedCell,
    setSelectedType,
  )
  selectionModel.addListener('change', onSelectionChange)
  return () => selectionModel.removeListener(onSelectionChange)
}

function attachScaleListener({ graph, setZoomPercent }: DiagramGraphListenerOptions): () => void {
  const view = graph.getView()
  const onScaleChange = () => {
    setZoomPercent((view.scale || 1) * 100)
  }
  view.addListener('scale', onScaleChange)
  view.addListener('scaleAndTranslate', onScaleChange)
  return () => {
    view.removeListener(onScaleChange)
  }
}

function attachKeyboardListener({
  graph,
  undoManager,
  setCanUndo,
  setCanRedo,
  setZoomPercent,
}: DiagramGraphListenerOptions): () => void {
  const onKeyDown = (event: KeyboardEvent) => {
    handleDiagramKeyDown(event, {
      graph,
      undoManager,
      setCanUndo,
      setCanRedo,
      setZoomPercent,
    })
  }
  window.addEventListener('keydown', onKeyDown)
  return () => window.removeEventListener('keydown', onKeyDown)
}

function useDiagramGraphLifecycle({
  initialLoadReady,
  mode,
  gridEnabled,
  canvasRef,
  graphRef,
  initialDrawioContentRef,
  initialDrawioLoadStateRef,
  graphInstance,
  setGraphInstance,
  setCanUndo,
  setCanRedo,
  setDirty,
  setSelectedCount,
  setSelectedCell,
  setSelectedType,
  setZoomPercent,
  setOperationError,
  onEmptyConnection,
}: DiagramGraphLifecycleOptions): void {
  useEffect(() => {
    if (!initialLoadReady || mode !== 'visual' || !canvasRef.current || graphRef.current) return

    let disposed = false
    const instance = createDiagramGraph(canvasRef.current, onEmptyConnection)
    graphRef.current = instance
    setGraphInstance(instance)
    const { graph, undoManager } = instance
    const listenerOptions: DiagramGraphListenerOptions = {
      graph,
      undoManager,
      setCanUndo,
      setCanRedo,
      setDirty,
      setSelectedCount,
      setSelectedCell,
      setSelectedType,
      setZoomPercent,
    }
    const cleanupListeners = [
      attachHistoryListener(listenerOptions),
      attachSelectionListener(listenerOptions),
      attachScaleListener(listenerOptions),
      attachKeyboardListener(listenerOptions),
    ]

    const initialXml = initialDrawioContentRef.current
    initialDrawioContentRef.current = null
    if (initialXml !== null) {
      initialDrawioLoadStateRef.current = 'loading'
      const handleFailedInitialLoad = () => {
        initialDrawioLoadStateRef.current = 'failed'
        setOperationError('Could not load the diagram.')
      }
      void loadDrawioIntoGraph(graph, initialXml)
        .then((loaded) => {
          if (disposed) return
          if (!loaded) {
            handleFailedInitialLoad()
            return
          }
          initialDrawioLoadStateRef.current = 'ready'
          undoManager.clear()
          setDirty(false)
          window.diagramApi?.setDirty(false)
        })
        .catch(() => {
          if (!disposed) handleFailedInitialLoad()
        })
    }

    return () => {
      disposed = true
      cleanupListeners.forEach((cleanup) => cleanup())
      instance.destroy()
      graphRef.current = null
      setGraphInstance(null)
    }
  }, [
    canvasRef,
    initialDrawioLoadStateRef,
    initialLoadReady,
    mode,
    setCanRedo,
    setCanUndo,
    setDirty,
    setOperationError,
    onEmptyConnection,
    setGraphInstance,
    setSelectedCell,
    setSelectedCount,
    setSelectedType,
    setZoomPercent,
  ])

  useEffect(() => {
    graphRef.current?.graph.setGridEnabled(gridEnabled)
  }, [graphInstance, graphRef, gridEnabled])
}

interface PendingDiagramLoaderOptions {
  setInitialLoadReady: (value: boolean) => void
  setMode: (value: DiagramMode) => void
  setFilePath: (value: string) => void
  setMermaidCode: (value: string) => void
  setDirty: (value: boolean) => void
  setSnapshots: (value: DiagramSnapshot[]) => void
  setOperationError: (value: string) => void
  initialDrawioContentRef: { current: string | null }
}

function usePendingDiagramLoader({
  setInitialLoadReady,
  setMode,
  setFilePath,
  setMermaidCode,
  setDirty,
  setSnapshots,
  setOperationError,
  initialDrawioContentRef,
}: PendingDiagramLoaderOptions): void {
  const isMountedRef = useRef(false)

  useEffect(() => {
    if (isMountedRef.current) return
    isMountedRef.current = true

    const loadInitialFile = async () => {
      if (!window.diagramApi) {
        setInitialLoadReady(true)
        return
      }
      try {
        const pending: PendingDiagram = await window.diagramApi.consumePending()
        const initialPath = pending.path
        setMode(pending.mode)
        if (initialPath) {
          setFilePath(initialPath)
          const content = await window.diagramApi.readFile(initialPath)
          if (pending.mode === 'mermaid') {
            setMermaidCode(content)
          } else {
            initialDrawioContentRef.current = content
          }
          setDirty(false)
          window.diagramApi.setDirty(false)
          setSnapshots([])
        }
      } catch {
        setOperationError('Could not load the diagram.')
      } finally {
        setInitialLoadReady(true)
      }
    }
    void loadInitialFile()
  }, [
    initialDrawioContentRef,
    setDirty,
    setFilePath,
    setInitialLoadReady,
    setMermaidCode,
    setMode,
    setOperationError,
    setSnapshots,
  ])
}

interface DiagramSaveOptions {
  mode: DiagramMode
  mermaidCode: string
  filePath: string | null
  autoSave: boolean
  dirty: boolean
  graphRef: { current: GraphInstance | null }
  initialDrawioLoadStateRef: { current: 'ready' | 'loading' | 'failed' }
  setFilePath: (value: string) => void
  setDirty: (value: boolean) => void
  setOperationError: (value: string) => void
}

function useDiagramSave({
  mode,
  mermaidCode,
  filePath,
  autoSave,
  dirty,
  graphRef,
  initialDrawioLoadStateRef,
  setFilePath,
  setDirty,
  setOperationError,
}: DiagramSaveOptions): (saveMode?: SaveMode) => Promise<boolean> {
  const handleSave = useCallback(
    async (saveMode: SaveMode = 'save') => {
      if (!window.diagramApi) return false

      if (mode === 'visual' && initialDrawioLoadStateRef.current !== 'ready') {
        setOperationError(
          initialDrawioLoadStateRef.current === 'loading'
            ? 'The diagram is still loading.'
            : 'Could not load the diagram.',
        )
        return false
      }

      try {
        let content = ''
        const format = mode === 'mermaid' ? 'mermaid' : 'drawio'

        if (mode === 'mermaid') {
          content = mermaidCode
        } else if (graphRef.current) {
          content = serializeGraphToDrawio(graphRef.current.graph)
        } else {
          return false
        }

        const res = await window.diagramApi.save({
          content,
          mode: saveMode,
          format,
          suggestedName: filePath ? undefined : 'diagram',
        })

        if (res.ok && 'path' in res && res.path) {
          setFilePath(res.path)
          setDirty(false)
          window.diagramApi.setDirty(false)
          return true
        }
        if (!res.ok) setOperationError('Could not save the diagram.')
      } catch {
        setOperationError('Could not save the diagram.')
      }

      return false
    },
    [
      filePath,
      graphRef,
      initialDrawioLoadStateRef,
      mermaidCode,
      mode,
      setDirty,
      setFilePath,
      setOperationError,
    ],
  )

  useEffect(() => {
    window.localStorage.setItem('ara-diagram.autoSave', autoSave ? '1' : '0')
  }, [autoSave])

  const handleSaveRef = useRef(handleSave)
  handleSaveRef.current = handleSave
  const dirtyRef = useRef(dirty)
  dirtyRef.current = dirty

  useEffect(() => {
    if (!autoSave || !filePath) return

    let saving = false
    const saveIfDirty = () => {
      if (!dirtyRef.current || saving) return
      saving = true
      void handleSaveRef.current('save').finally(() => {
        saving = false
      })
    }
    const intervalId = window.setInterval(saveIfDirty, 30_000)
    window.addEventListener('blur', saveIfDirty)
    return () => {
      window.clearInterval(intervalId)
      window.removeEventListener('blur', saveIfDirty)
    }
  }, [autoSave, filePath])

  return handleSave
}

interface DiagramExportOptions {
  mode: DiagramMode
  filePath: string | null
  mermaidCode: string
  graphRef: { current: GraphInstance | null }
  setOperationError: (value: string) => void
}

interface DiagramExports {
  handleExportSvg: () => Promise<void>
  handleExportImage: (format: 'png' | 'jpeg') => Promise<void>
}

function useDiagramExports({
  mode,
  filePath,
  mermaidCode,
  graphRef,
  setOperationError,
}: DiagramExportOptions): DiagramExports {
  const handleExportSvg = useCallback(async () => {
    if (!window.diagramApi) return

    try {
      let svgString = ''
      if (mode === 'mermaid') {
        const container = document.querySelector('.mermaid-svg-container svg')
        if (container) svgString = new XMLSerializer().serializeToString(container)
      } else if (graphRef.current) {
        svgString = exportGraphToSvg(graphRef.current.graph)
      }

      if (svgString) {
        await window.diagramApi.exportSvg({
          data: svgString,
          format: 'svg',
          suggestedName: filePath
            ? filePath
                .split(/[/\\]/)
                .pop()
                ?.replace(/\.[^.]+$/, '')
            : 'diagram',
        })
      }
    } catch {
      setOperationError('Could not export the SVG diagram.')
    }
  }, [filePath, graphRef, mode, setOperationError])

  const handleExportImage = useCallback(
    async (format: 'png' | 'jpeg') => {
      if (!window.diagramApi) return
      try {
        let dataUrl: string
        if (mode === 'mermaid') {
          // Render fresh from source: the live preview SVG carries the current pan/zoom viewBox
          const rendered = await renderMermaidDiagram(mermaidCode)
          if (!rendered.svg) throw new Error(rendered.error || 'nothing to export')
          dataUrl = await svgStringToRasterDataUrl(rendered.svg, {
            mime: format === 'jpeg' ? 'image/jpeg' : 'image/png',
          })
        } else {
          if (!graphRef.current) return
          dataUrl = await exportGraphToRaster(
            graphRef.current.graph,
            format === 'jpeg' ? 'image/jpeg' : 'image/png',
          )
        }
        await window.diagramApi.exportImage({
          data: dataUrl,
          format,
          suggestedName: filePath
            ? filePath
                .split(/[/\\]/)
                .pop()
                ?.replace(/\.[^.]+$/, '')
            : 'diagram',
        })
      } catch {
        setOperationError(`Could not export the ${format.toUpperCase()} image.`)
      }
    },
    [filePath, graphRef, mermaidCode, mode, setOperationError],
  )

  return { handleExportSvg, handleExportImage }
}

interface DiagramIpcEventsOptions {
  handleSave: (saveMode?: SaveMode) => Promise<boolean>
  handleExportSvg: () => Promise<void>
  handleExportImage: (format: 'png' | 'jpeg') => Promise<void>
  setFilePath: (value: string) => void
}

function useDiagramIpcEvents({
  handleSave,
  handleExportSvg,
  handleExportImage,
  setFilePath,
}: DiagramIpcEventsOptions): void {
  const handleSaveRef = useRef(handleSave)
  handleSaveRef.current = handleSave
  const handleExportSvgRef = useRef(handleExportSvg)
  handleExportSvgRef.current = handleExportSvg
  const handleExportImageRef = useRef(handleExportImage)
  handleExportImageRef.current = handleExportImage

  useEffect(() => {
    if (!window.diagramApi) return

    const unsubSave = window.diagramApi.onSaveRequest((reqMode) => {
      void handleSaveRef.current(reqMode).then(
        (ok) => {
          if (!ok) window.diagramApi?.sendSaveRequestAck(false)
        },
        () => window.diagramApi?.sendSaveRequestAck(false),
      )
    })

    const unsubCloseSave = window.diagramApi.onCloseSaveRequest(() => {
      void handleSaveRef.current('save').then(
        (ok) => window.diagramApi?.sendCloseSaveResult(ok),
        () => window.diagramApi?.sendCloseSaveResult(false),
      )
    })

    const unsubRename = window.diagramApi.onFileRenamed((newPath) => {
      setFilePath(newPath)
    })

    const unsubExport = window.diagramApi.onExportRequest((fmt: ExportFormat) => {
      if (fmt === 'svg') void handleExportSvgRef.current()
      else void handleExportImageRef.current(fmt)
    })

    return () => {
      unsubSave()
      unsubCloseSave()
      unsubRename()
      unsubExport()
    }
  }, [setFilePath])
}

interface DiagramAiDependenciesOptions {
  mode: DiagramMode
  filePath: string | null
  mermaidCode: string
  graphRef: { current: GraphInstance | null }
  setMermaidCode: (value: string) => void
  setZoomPercent: (value: number) => void
  setDirty: (value: boolean) => void
  handleClear: () => void
  captureSnapshot: (label: string) => void
  handleExportSvg: () => Promise<void>
  handleExportImage: (format: 'png' | 'jpeg') => Promise<void>
}

function useDiagramAiDependencies({
  mode,
  filePath,
  mermaidCode,
  graphRef,
  setMermaidCode,
  setZoomPercent,
  setDirty,
  handleClear,
  captureSnapshot,
  handleExportSvg,
  handleExportImage,
}: DiagramAiDependenciesOptions): DiagramSkillDeps {
  return useMemo<DiagramSkillDeps>(
    () => ({
      getMode: () => mode,
      getFileName: () => (filePath ? filePath.split(/[/\\]/).pop() || 'Untitled' : 'Untitled'),
      getMermaidCode: () => mermaidCode,
      setMermaidCode: (code) => {
        setZoomPercent(100)
        setMermaidCode(code)
        setDirty(true)
        window.diagramApi?.setDirty(true)
      },
      addVisualShape: (shape) => {
        if (!graphRef.current) return { id: '0', label: shape.label }
        const { graph } = graphRef.current
        const cell = addShape(graph, shape.type, {
          label: shape.label,
          x: shape.x,
          y: shape.y,
          width: shape.width,
          height: shape.height,
        })
        if (shape.fillColor) setSelectionStyle(graph, 'fillColor', shape.fillColor)
        if (shape.strokeColor) setSelectionStyle(graph, 'strokeColor', shape.strokeColor)
        setDirty(true)
        window.diagramApi?.setDirty(true)
        return { id: String(cell.id), label: shape.label }
      },
      connectVisualShapes: (sourceId, targetId, label) => {
        if (!graphRef.current) return { id: '0' }
        const { graph } = graphRef.current
        const model = graph.getDataModel()
        const source = model.getCell(sourceId)
        const target = model.getCell(targetId)
        if (!source || !target) return { id: '0' }
        const edge = connectCells(graph, source, target, label)
        setDirty(true)
        window.diagramApi?.setDirty(true)
        return { id: String(edge.id) }
      },

      layoutVisualDiagram: (direction) => {
        if (!graphRef.current) return
        autoLayout(graphRef.current.graph, direction)
        setDirty(true)
        window.diagramApi?.setDirty(true)
      },
      getVisualSummary: () => {
        if (!graphRef.current) return { vertexCount: 0, edgeCount: 0, nodes: [] }
        const { graph } = graphRef.current
        const parent = graph.getDefaultParent()
        const vertices = graph.getChildVertices(parent)
        const edges = graph.getChildEdges(parent)
        return {
          vertexCount: vertices.length,
          edgeCount: edges.length,
          nodes: vertices.map((v) => ({ id: String(v.id), label: String(v.value || '') })),
        }
      },
      clearDiagram: handleClear,
      captureSnapshot,
      exportDiagram: async (fmt) => {
        if (fmt === 'svg') await handleExportSvg()
        else await handleExportImage(fmt)
      },
    }),
    [
      captureSnapshot,
      filePath,
      graphRef,
      handleClear,
      handleExportImage,
      handleExportSvg,
      mermaidCode,
      mode,
      setDirty,
      setMermaidCode,
      setZoomPercent,
    ],
  )
}

function ConnectionQuickAdd({
  pending,
  onSelect,
  onDismiss,
  lang,
}: {
  pending: PendingConnection
  onSelect: (type: DiagramShapeType) => void
  onDismiss: () => void
  lang: Lang
}): ReactElement {
  const [query, setQuery] = useState('')
  const [hoveredShape, setHoveredShape] = useState<DiagramShapeType | null>(null)
  const pickerRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  const copy = quickAddCopy(lang)
  const shapes = query
    ? QUICK_CONNECT_SHAPES.filter((type) =>
        shapeHelp(type, lang).label.toLowerCase().includes(query.toLowerCase()),
      )
    : QUICK_CONNECT_SHAPES.slice(0, 4)
  const position = {
    left: Math.max(8, Math.min(pending.clientX + 12, window.innerWidth - 252)),
    top: Math.max(8, Math.min(pending.clientY + 12, window.innerHeight - 218)),
  }

  useEffect(() => {
    inputRef.current?.focus()
    const onMouseDown = (event: MouseEvent) => {
      if (event.button !== 0 || pickerRef.current?.contains(event.target as Node)) return
      event.preventDefault()
      event.stopPropagation()
      onDismiss()
    }
    document.addEventListener('mousedown', onMouseDown, true)
    return () => document.removeEventListener('mousedown', onMouseDown, true)
  }, [onDismiss])

  return (
    <div
      ref={pickerRef}
      className="connection-quick-add"
      role="dialog"
      aria-label={copy.title}
      style={position}
    >
      <div className="connection-quick-add-title">{copy.title}</div>
      <input
        ref={inputRef}
        type="search"
        value={query}
        placeholder={copy.search}
        aria-label={copy.search}
        onChange={(event) => setQuery(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === 'Escape') onDismiss()
          if (event.key === 'Enter' && shapes[0]) onSelect(shapes[0])
        }}
      />
      <div className="connection-quick-add-grid">
        {shapes.map((type) => {
          const shape = shapeHelp(type, lang)
          return (
            <button
              key={type}
              type="button"
              title={shape.description}
              onMouseEnter={() => setHoveredShape(type)}
              onFocus={() => setHoveredShape(type)}
              onClick={() => onSelect(type)}
            >
              <span className={`connection-quick-add-icon ${type}`} aria-hidden="true" />
              {shape.label}
            </button>
          )
        })}
      </div>
      {shapes.length === 0 && <div className="connection-quick-add-empty">{copy.noMatch}</div>}
      <div className="connection-quick-add-hint">
        {hoveredShape ? shapeHelp(hoveredShape, lang).description : copy.hint}
      </div>
    </div>
  )
}

interface DiagramWorkspaceProps {
  ribbonProps: Parameters<typeof Ribbon>[0]
  lang: Lang
  operationError: string | null
  onDismissOperationError: () => void
  mode: DiagramMode
  visualProperties: ReactElement | null
  canvasRef: { current: HTMLDivElement | null }
  onDropShape: (type: DiagramShapeType, clientX: number, clientY: number) => void
  pendingConnection: PendingConnection | null
  onAddConnectedShape: (type: DiagramShapeType) => void
  onDismissConnectedShape: () => void
  gridEnabled: boolean
  mermaidCode: string
  onMermaidChange: (code: string) => void
  selectedCount: number
  selectedType: SelectionType
  zoomPercent: number
  onZoomIn: () => void
  onZoomOut: () => void
  onZoomChange: (value: number) => void
  aiOpen: boolean
  onOpenAi: () => void
  aiDeps: DiagramSkillDeps
  filePath: string | null
  onCollapseAi: () => void
  aiPresetPrompt: string | null
  snapshots: DiagramSnapshot[]
  onRollback: (snapshot: DiagramSnapshot) => void
}

function DiagramWorkspace({
  ribbonProps,
  lang,
  operationError,
  onDismissOperationError,
  mode,
  visualProperties,
  canvasRef,
  onDropShape,
  pendingConnection,
  onAddConnectedShape,
  onDismissConnectedShape,
  gridEnabled,
  mermaidCode,
  onMermaidChange,
  selectedCount,
  selectedType,
  zoomPercent,
  onZoomIn,
  onZoomOut,
  onZoomChange,
  aiOpen,
  onOpenAi,
  aiDeps,
  filePath,
  onCollapseAi,
  aiPresetPrompt,
  snapshots,
  onRollback,
}: DiagramWorkspaceProps): ReactElement {
  const [shapeDropActive, setShapeDropActive] = useState(false)

  const isShapeDrag = (event: DragEvent<HTMLDivElement>) =>
    Array.from(event.dataTransfer.types).includes(DIAGRAM_SHAPE_DRAG_MIME)

  const handleCanvasDragOver = (event: DragEvent<HTMLDivElement>) => {
    if (!isShapeDrag(event)) return
    event.preventDefault()
    event.dataTransfer.dropEffect = 'copy'
    setShapeDropActive(true)
  }

  const handleCanvasDragLeave = (event: DragEvent<HTMLDivElement>) => {
    const relatedTarget = event.relatedTarget
    if (relatedTarget instanceof Node && event.currentTarget.contains(relatedTarget)) return
    setShapeDropActive(false)
  }

  const handleCanvasDrop = (event: DragEvent<HTMLDivElement>) => {
    if (!isShapeDrag(event)) return
    event.preventDefault()
    setShapeDropActive(false)

    const type = event.dataTransfer.getData(DIAGRAM_SHAPE_DRAG_MIME)
    if (isDiagramShapeType(type)) onDropShape(type, event.clientX, event.clientY)
  }

  const canvasClassName = `diagram-canvas-container${gridEnabled ? '' : ' grid-hidden'}`

  return (
    <div className="diagram-app">
      <Ribbon {...ribbonProps} />
      {operationError && (
        <div className="diagram-operation-error" role="alert">
          <span>{operationError}</span>
          <button type="button" onClick={onDismissOperationError}>
            Dismiss
          </button>
        </div>
      )}
      <div className="diagram-main-layout">
        <div className="diagram-content-column">
          <main className="diagram-body">
            {mode === 'visual' ? (
              <>
                {visualProperties}
                <div
                  className={`diagram-canvas-viewport${shapeDropActive ? ' is-drop-target' : ''}`}
                  onDragOver={handleCanvasDragOver}
                  onDragLeave={handleCanvasDragLeave}
                  onDrop={handleCanvasDrop}
                >
                  <div ref={canvasRef} className={canvasClassName} />
                  {pendingConnection && (
                    <ConnectionQuickAdd
                      pending={pendingConnection}
                      onSelect={onAddConnectedShape}
                      onDismiss={onDismissConnectedShape}
                      lang={lang}
                    />
                  )}
                </div>
              </>
            ) : (
              <MermaidEditor
                code={mermaidCode}
                onChange={onMermaidChange}
                zoom={zoomPercent}
                onZoomChange={onZoomChange}
              />
            )}
          </main>

          <StatusBar
            selectedCount={selectedCount}
            selectedType={selectedType}
            zoomPercent={zoomPercent}
            mode={mode}
            gridEnabled={gridEnabled}
            onZoomIn={onZoomIn}
            onZoomOut={onZoomOut}
            onZoomChange={onZoomChange}
          />
        </div>

        <div className={`ai-dock${aiOpen ? '' : ' collapsed'}`}>
          {!aiOpen && (
            <button type="button" className="ai-rail" title="Open ARA Assistant" onClick={onOpenAi}>
              <span className="ai-rail-icon" aria-hidden="true">
                <AraSpark size={21} />
              </span>
            </button>
          )}
          <AiPanel
            deps={aiDeps}
            filePath={filePath}
            onCollapse={onCollapseAi}
            presetPrompt={aiPresetPrompt}
            snapshots={snapshots}
            onRollback={onRollback}
          />
        </div>
      </div>
    </div>
  )
}

interface DiagramPropertiesProps {
  showPropsPanel: boolean
  selectedCell: Cell | null
  graph: GraphInstance['graph'] | null
  setSelectedCell: (cell: Cell | null) => void
  onCollapse: () => void
  onExpand: () => void
}

function DiagramProperties({
  showPropsPanel,
  selectedCell,
  graph,
  setSelectedCell,
  onCollapse,
  onExpand,
}: DiagramPropertiesProps): ReactElement {
  if (showPropsPanel) {
    return (
      <PropertiesPanel
        selectedCell={selectedCell}
        onUpdateLabel={(val) => {
          if (graph && selectedCell) {
            setCellLabel(graph, selectedCell, val)
            // ponytail: clone keeps the Cell prototype (isEdge/isVertex) —
            // object spread strips it and crashes the panel on next render.
            setSelectedCell(
              Object.assign(Object.create(Object.getPrototypeOf(selectedCell)), selectedCell, {
                value: val,
              }) as Cell,
            )
          }
        }}
        onUpdateStyle={(key, val) => {
          if (graph) {
            setSelectionStyle(
              graph,
              key as keyof CellStateStyle,
              val as CellStateStyle[keyof CellStateStyle],
            )
            if (selectedCell) {
              setSelectedCell(
                Object.assign(
                  Object.create(Object.getPrototypeOf(selectedCell)),
                  selectedCell,
                ) as Cell,
              )
            }
          }
        }}
        onCollapse={onCollapse}
      />
    )
  }

  return (
    <button
      type="button"
      className="diagram-props-rail"
      title="Expand properties panel"
      aria-label="Expand properties panel"
      onClick={onExpand}
    >
      <IconProperties size={16} />
      <span>Properties</span>
    </button>
  )
}

export default function App({ initialLang = 'en' }: { initialLang?: Lang }): ReactElement {
  const [lang, setLang] = useState<Lang>(initialLang)
  const [mode, setMode] = useState<DiagramMode>('visual')
  const [filePath, setFilePath] = useState<string | null>(null)
  const [dirty, setDirty] = useState(false)
  const [autoSave, setAutoSave] = useState(
    () => window.localStorage.getItem('ara-diagram.autoSave') === '1',
  )
  const [pendingConnection, setPendingConnection] = useState<PendingConnection | null>(null)
  const [canUndo, setCanUndo] = useState(false)
  const [canRedo, setCanRedo] = useState(false)
  const [zoomPercent, setZoomPercent] = useState(100)
  const [gridEnabled, setGridEnabled] = useState(true)
  const [selectedCell, setSelectedCell] = useState<Cell | null>(null)
  const [selectedCount, setSelectedCount] = useState(0)
  const [selectedType, setSelectedType] = useState<'vertex' | 'edge' | 'mixed' | undefined>()
  const [mermaidCode, setMermaidCode] = useState(DEFAULT_MERMAID)
  const [showPropsPanel, setShowPropsPanel] = useState(false)
  const [aiOpen, setAiOpen] = useState(() => localStorage.getItem('ai-diagram-open') !== '0')
  const [aiPresetPrompt, setAiPresetPrompt] = useState<string | null>(null)
  const [snapshots, setSnapshots] = useState<DiagramSnapshot[]>([])
  const [operationError, setOperationError] = useState<string | null>(null)
  useEffect(() => {
    document.documentElement.lang = htmlLang(lang)
  }, [lang])
  useEffect(() => window.diagramApi?.onLanguageChanged(setLang), [])

  // Delegated by ARA Omni (shell Command Center): open the AI panel; AiPanel auto-sends the preset
  useEffect(
    () =>
      window.diagramApi?.onMenuCommand((command, payload) => {
        if (command === 'ai-prompt' && payload) {
          setAiOpen(true)
          // Same behavior as every other app: the handover envelope is the
          // instruction and the bubble — no display/model split.
          setAiPresetPrompt(payload)
        }
      }) ?? undefined,
    [],
  )

  useEffect(() => {
    localStorage.setItem('ai-diagram-open', aiOpen ? '1' : '0')
  }, [aiOpen])

  useEffect(() => {
    if (selectedCount > 0) setShowPropsPanel(true)
  }, [selectedCount])

  const captureSnapshot = useCallback(
    (label: string) => {
      const content =
        mode === 'visual'
          ? graphRef.current
            ? serializeGraphToDrawio(graphRef.current.graph)
            : ''
          : mermaidCode
      if (!content) return
      const snap: DiagramSnapshot = {
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        time: new Date().toLocaleTimeString([], {
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
        }),
        label,
        mode,
        content,
      }
      // ponytail: cap snapshot history to 5 items to avoid OOM from full XML retention
      setSnapshots((prev) => [snap, ...prev].slice(0, 5))
    },
    [mermaidCode, mode],
  )

  const handleRollback = useCallback(async (snapshot: DiagramSnapshot) => {
    if (snapshot.mode === 'visual' && graphRef.current) {
      await loadDrawioIntoGraph(graphRef.current.graph, snapshot.content)
      // ponytail: clear undo history on rollback so detached cell trees do not accumulate
      graphRef.current.undoManager.clear()
    } else if (snapshot.mode === 'mermaid') {
      setMermaidCode(snapshot.content)
    }
    setDirty(true)
    window.diagramApi?.setDirty(true)
    setSnapshots((prev) => prev.filter((s) => s.id !== snapshot.id))
  }, [])
  const handleEmptyConnection = useCallback((source: Cell, clientX: number, clientY: number) => {
    setPendingConnection({ source, clientX, clientY })
  }, [])
  const canvasRef = useRef<HTMLDivElement>(null)
  const graphRef = useRef<GraphInstance | null>(null)
  const initialDrawioLoadStateRef = useRef<'ready' | 'loading' | 'failed'>('ready')
  const [graphInstance, setGraphInstance] = useState<GraphInstance | null>(null)
  const [initialLoadReady, setInitialLoadReady] = useState(false)
  const initialDrawioContentRef = useRef<string | null>(null)
  useDiagramGraphLifecycle({
    initialLoadReady,
    mode,
    onEmptyConnection: handleEmptyConnection,
    gridEnabled,
    canvasRef,
    graphRef,
    initialDrawioContentRef,
    initialDrawioLoadStateRef,
    graphInstance,
    setGraphInstance,
    setCanUndo,
    setCanRedo,
    setDirty,
    setSelectedCount,
    setSelectedCell,
    setSelectedType,
    setZoomPercent,
    setOperationError,
  })
  usePendingDiagramLoader({
    setInitialLoadReady,
    setMode,
    setFilePath,
    setMermaidCode,
    setDirty,
    setSnapshots,
    setOperationError,
    initialDrawioContentRef,
  })

  const handleSave = useDiagramSave({
    mode,
    mermaidCode,
    filePath,
    autoSave,
    dirty,
    graphRef,
    initialDrawioLoadStateRef,
    setFilePath,
    setDirty,
    setOperationError,
  })
  const { handleExportSvg, handleExportImage } = useDiagramExports({
    mode,
    filePath,
    mermaidCode,
    graphRef,
    setOperationError,
  })
  useDiagramIpcEvents({
    handleSave,
    handleExportSvg,
    handleExportImage,
    setFilePath,
  })
  const {
    graph: g,
    handleInsertShape,
    handleClear,
    onUndo,
    onRedo,
    onZoomIn,
    onZoomOut,
    onZoomActual,
    handleDropShape,
    onZoomFit,
    onZoomChange,
    onDuplicate,
    onDelete,
    onSelectAll,
    onConnect,
    onAlign,
    onDistribute,
    onOrder,
  } = useDiagramCommands({
    mode,
    graphInstance,
    graphRef,
    setZoomPercent,
    setDirty,
    setMermaidCode,
    setSnapshots,
  })

  const handleAddConnectedShape = useCallback(
    (type: DiagramShapeType) => {
      const pending = pendingConnection
      const graph = graphRef.current?.graph
      if (pending && graph && pending.source.getParent()) {
        graph.batchUpdate(() => {
          const target = handleDropShape(type, pending.clientX, pending.clientY)
          if (target) connectCells(graph, pending.source, target)
        })
      }
      setPendingConnection(null)
    },
    [graphRef, handleDropShape, pendingConnection],
  )

  const aiDeps = useDiagramAiDependencies({
    mode,
    filePath,
    mermaidCode,
    graphRef,
    setMermaidCode,
    setZoomPercent,
    setDirty,
    handleClear,
    captureSnapshot,
    handleExportSvg,
    handleExportImage,
  })
  const uiActions = useDiagramUiActions({
    handleSave,
    aiOpen,
    setAiOpen,
    setAiPresetPrompt,
    setShowPropsPanel,
    setGridEnabled,
    setMermaidCode,
    setDirty,
    setOperationError,
  })
  const visualProperties = (
    <DiagramProperties
      showPropsPanel={showPropsPanel}
      selectedCell={selectedCell}
      graph={g ?? null}
      setSelectedCell={setSelectedCell}
      onCollapse={uiActions.onCollapseProps}
      onExpand={uiActions.onExpandProps}
    />
  )
  const ribbonProps: Parameters<typeof Ribbon>[0] = {
    mode,
    canUndo,
    canRedo,
    onUndo,
    onRedo,
    onZoomIn,
    onZoomOut,
    onZoomActual,
    onZoomFit,
    zoomPercent,
    onDuplicate,
    onDelete,
    onSelectAll,
    onConnect,
    onAlign,
    onDistribute,
    onOrder,
    selectionCount: selectedCount,
    onClear: handleClear,
    hasSelection: selectedCount > 0,
    onExportSvg: handleExportSvg,
    onExportImage: handleExportImage,
    onSave: uiActions.onSave,
    onSaveAs: uiActions.onSaveAs,
    dirty,
    filePath,
    autoSave,
    onAutoSaveChange: setAutoSave,
    aiOpen,
    onToggleAi: uiActions.onToggleAi,
    onAiPrompt: uiActions.onAiPrompt,
    showPropsPanel,
    onTogglePropsPanel: uiActions.onTogglePropsPanel,
    gridEnabled,
    onToggleGrid: uiActions.onToggleGrid,
    onInsertShape: handleInsertShape,
    onApplyTemplate: uiActions.onApplyTemplate,
    onInsertMermaidSnippet: uiActions.onInsertMermaidSnippet,
    lang,
  }

  return (
    <DiagramWorkspace
      ribbonProps={ribbonProps}
      operationError={operationError}
      onDismissOperationError={uiActions.onDismissOperationError}
      mode={mode}
      visualProperties={visualProperties}
      canvasRef={canvasRef}
      onDropShape={handleDropShape}
      pendingConnection={pendingConnection}
      onAddConnectedShape={handleAddConnectedShape}
      onDismissConnectedShape={() => setPendingConnection(null)}
      lang={lang}
      gridEnabled={gridEnabled}
      mermaidCode={mermaidCode}
      onMermaidChange={uiActions.onMermaidChange}
      selectedCount={selectedCount}
      selectedType={selectedType}
      zoomPercent={zoomPercent}
      onZoomIn={onZoomIn}
      onZoomOut={onZoomOut}
      onZoomChange={onZoomChange}
      aiOpen={aiOpen}
      onOpenAi={uiActions.onOpenAi}
      aiDeps={aiDeps}
      filePath={filePath}
      onCollapseAi={uiActions.onCollapseAi}
      aiPresetPrompt={aiPresetPrompt}
      snapshots={snapshots}
      onRollback={handleRollback}
    />
  )
}
