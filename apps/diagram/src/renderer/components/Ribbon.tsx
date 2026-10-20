import { useEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import type { Lang } from '@araoffice/i18n'
import type {
  AlignDirection,
  DiagramSelectionCommands,
  DiagramShapeType,
  DistributeDirection,
} from '../maxgraph/diagramCommands'
import { DIAGRAM_SHAPE_DRAG_MIME } from '../maxgraph/diagramCommands'
import { shapeHelp } from '../maxgraph/shape-help'
import {
  AssistantMark,
  IconCode,
  IconAlignBottom,
  IconAlignLeft,
  IconAlignCenter,
  IconAlignMiddle,
  IconAlignRight,
  IconAlignTop,
  IconAutoLayout,
  IconBringFront,
  IconClear,
  IconCloud,
  IconConnector,
  IconCylinder,
  IconDelete,
  IconDiamond,
  IconDistributeHorizontal,
  IconDistributeVertical,
  IconDownload,
  IconDuplicate,
  IconEllipse,
  IconDoubleEllipse,
  IconFit,
  IconFlowchart,
  IconGrid,
  IconHexagon,
  IconProperties,
  IconRectangle,
  IconRedo,
  IconRoundedRectangle,
  IconSave,
  IconSchema,
  IconSelectAll,
  IconSendBack,
  IconSequence,
  IconSparkles,
  IconText,
  IconTriangle,
  IconUndo,
  IconZoomIn,
  IconZoomOut,
} from './icons'

export interface DiagramRibbonProps extends DiagramSelectionCommands {
  mode: 'visual' | 'mermaid'
  canUndo: boolean
  canRedo: boolean
  onUndo: () => void
  onRedo: () => void
  onZoomIn: () => void
  onZoomOut: () => void
  onZoomActual: () => void
  onZoomFit: () => void
  zoomPercent: number
  selectionCount: number
  onClear: () => void
  hasSelection: boolean
  onExportSvg: () => void
  onExportImage: (format: 'png' | 'jpeg') => void
  onSave: () => void
  onSaveAs: () => void
  dirty: boolean
  filePath?: string | null
  autoSave: boolean
  onAutoSaveChange: (enabled: boolean) => void
  aiOpen: boolean
  onToggleAi: () => void
  onAiPrompt?: (prompt: string) => void
  showPropsPanel: boolean
  onTogglePropsPanel: () => void
  gridEnabled: boolean
  onToggleGrid: () => void
  onInsertShape?: (type: DiagramShapeType) => void
  lang: Lang
  onApplyTemplate?: (template: string) => void
  onInsertMermaidSnippet?: (snippet: string) => void
}

const IS_MAC = /mac/i.test(navigator.userAgent)
const IN_TAB = new URLSearchParams(window.location.search).get('mode') === 'tab'

const SHAPE_BUTTONS: Array<{ type: DiagramShapeType; icon: ReactElement }> = [
  { type: 'rectangle', icon: <IconRectangle /> },
  { type: 'rounded', icon: <IconRoundedRectangle /> },
  { type: 'ellipse', icon: <IconEllipse /> },
  { type: 'doubleEllipse', icon: <IconDoubleEllipse /> },
  { type: 'diamond', icon: <IconDiamond /> },
  { type: 'cylinder', icon: <IconCylinder /> },
  { type: 'triangle', icon: <IconTriangle /> },
  { type: 'hexagon', icon: <IconHexagon /> },
  { type: 'cloud', icon: <IconCloud /> },
  { type: 'text', icon: <IconText /> },
]
const MERMAID_TEMPLATES: Array<{
  label: string
  code: string
  icon: ReactElement
}> = [
  {
    label: 'Flowchart',
    code: `flowchart TD
    A[Start] --> B[Process Task]
    B --> C{Success?}
    C -- Yes --> D[Done]
    C -- No --> E[Retry]`,
    icon: <IconFlowchart size={15} />,
  },
  {
    label: 'Sequence',
    code: `sequenceDiagram
    autonumber
    actor User
    participant App
    participant Server
    User->>App: Click Action
    App->>Server: POST /api/v1/run
    Server-->>App: 200 OK`,
    icon: <IconSequence size={15} />,
  },
  {
    label: 'ER Diagram',
    code: `erDiagram
    USER ||--o{ ORDER : places
    ORDER ||--|{ LINE_ITEM : contains
    PRODUCT ||--o{ LINE_ITEM : includes`,
    icon: <IconSchema size={15} />,
  },
  {
    label: 'Class',
    code: `classDiagram
    class BankAccount {
        +String owner
        +BigDecimal balance
        +deposit(amount)
        +withdraw(amount)
    }`,
    icon: <IconCode size={15} />,
  },
  {
    label: 'State',
    code: `stateDiagram-v2
    [*] --> Idle
    Idle --> Running : start
    Running --> Done : complete
    Done --> [*]`,
    icon: <IconCode size={15} />,
  },
  {
    label: 'Git Graph',
    code: `gitGraph
    commit
    branch develop
    checkout develop
    commit
    checkout main
    merge develop`,
    icon: <IconCode size={15} />,
  },
  {
    label: 'Pie Chart',
    code: `pie title Workload
    "Build" : 45
    "Review" : 30
    "Research" : 25`,
    icon: <IconCode size={15} />,
  },
  {
    label: 'Mind map',
    code: `mindmap
    root((Project))
        Planning
        Build
        Launch`,
    icon: <IconSparkles size={15} />,
  },
]

const ALIGN_BUTTONS: Array<{
  direction: AlignDirection
  label: string
  icon: ReactElement
}> = [
  { direction: 'left', label: 'Left', icon: <IconAlignLeft /> },
  { direction: 'center', label: 'Center', icon: <IconAlignCenter /> },
  { direction: 'right', label: 'Right', icon: <IconAlignRight /> },
  { direction: 'top', label: 'Top', icon: <IconAlignTop /> },
  { direction: 'middle', label: 'Middle', icon: <IconAlignMiddle /> },
  { direction: 'bottom', label: 'Bottom', icon: <IconAlignBottom /> },
]

const DISTRIBUTE_BUTTONS: Array<{
  direction: DistributeDirection
  label: string
  icon: ReactElement
}> = [
  { direction: 'horizontal', label: 'Horizontal', icon: <IconDistributeHorizontal /> },
  { direction: 'vertical', label: 'Vertical', icon: <IconDistributeVertical /> },
]
type RibbonTopbarProps = Pick<
  DiagramRibbonProps,
  | 'mode'
  | 'canUndo'
  | 'canRedo'
  | 'onUndo'
  | 'onRedo'
  | 'onSave'
  | 'onSaveAs'
  | 'onExportSvg'
  | 'onExportImage'
  | 'onClear'
  | 'autoSave'
  | 'onAutoSaveChange'
> & {
  saveDisabled: boolean
}

function RibbonTopbar({
  mode,
  canUndo,
  canRedo,
  onUndo,
  onRedo,
  onSave,
  onSaveAs,
  onExportSvg,
  onExportImage,
  onClear,
  autoSave,
  onAutoSaveChange,
  saveDisabled,
}: RibbonTopbarProps): ReactElement {
  const [fileMenuOpen, setFileMenuOpen] = useState(false)
  const fileMenuRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const onDocClick = (e: MouseEvent) => {
      if (fileMenuRef.current && !fileMenuRef.current.contains(e.target as Node)) {
        setFileMenuOpen(false)
      }
    }
    window.addEventListener('mousedown', onDocClick)
    return () => window.removeEventListener('mousedown', onDocClick)
  }, [])

  let topbarPlatformClass = ''
  if (!IN_TAB) {
    topbarPlatformClass = 'ribbon-topbar-win'
    if (IS_MAC) {
      topbarPlatformClass = 'ribbon-topbar-mac'
    }
  }

  return (
    <div className={`ribbon-topbar ${topbarPlatformClass}`} aria-label="Diagram file commands">
      {!IS_MAC && (
        <div className="file-tab-wrap" ref={fileMenuRef}>
          <button
            type="button"
            className={`ribbon-file-button ${mode === 'mermaid' ? 'diagram-file-code' : 'diagram-file-visual'} ${fileMenuOpen ? 'open' : ''}`}
            onClick={() => setFileMenuOpen((v) => !v)}
          >
            File
          </button>
          {fileMenuOpen && (
            <div className="file-menu">
              <button
                type="button"
                disabled={saveDisabled}
                onClick={() => {
                  setFileMenuOpen(false)
                  onSave()
                }}
              >
                <span>Save</span>
                <span className="file-menu-key">Ctrl+S</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  setFileMenuOpen(false)
                  onSaveAs()
                }}
              >
                <span>Save As</span>
                <span className="file-menu-key">Ctrl+Shift+S</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  setFileMenuOpen(false)
                  onExportSvg()
                }}
              >
                <span>Export SVG</span>
                <span className="file-menu-key">Vector</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  setFileMenuOpen(false)
                  onExportImage('png')
                }}
              >
                <span>Export PNG</span>
                <span className="file-menu-key">2× bitmap</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  setFileMenuOpen(false)
                  onExportImage('jpeg')
                }}
              >
                <span>Export JPEG</span>
                <span className="file-menu-key">2× bitmap</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  setFileMenuOpen(false)
                  onClear()
                }}
              >
                <span>Clear Canvas</span>
                <span className="file-menu-key">Delete all</span>
              </button>
            </div>
          )}
        </div>
      )}

      <button
        type="button"
        className="qa-btn"
        title="Save (Ctrl+S)"
        disabled={saveDisabled}
        onClick={onSave}
        aria-label="Save"
      >
        <IconSave size={16} />
      </button>
      <button
        type="button"
        className="qa-btn"
        disabled={!canUndo}
        title="Undo (Ctrl+Z)"
        onClick={onUndo}
        aria-label="Undo"
      >
        <IconUndo size={16} />
      </button>
      <button
        type="button"
        className="qa-btn"
        disabled={!canRedo}
        title="Redo (Ctrl+Y)"
        onClick={onRedo}
        aria-label="Redo"
      >
        <IconRedo size={16} />
      </button>
      <label className={`autosave-toggle ${autoSave ? 'on' : ''}`} title="Toggle AutoSave">
        <span className="autosave-knob" />
        <span className="autosave-text">AutoSave</span>
        <input
          type="checkbox"
          checked={autoSave}
          onChange={(e) => onAutoSaveChange(e.target.checked)}
        />
      </label>
      <span className="ribbon-topbar-spacer" />
    </div>
  )
}

type RibbonExportGroupProps = Pick<DiagramRibbonProps, 'onExportSvg' | 'onExportImage'>

function RibbonExportGroup({ onExportSvg, onExportImage }: RibbonExportGroupProps): ReactElement {
  const [exportMenuOpen, setExportMenuOpen] = useState(false)
  const exportMenuRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const onDocClick = (e: MouseEvent) => {
      if (exportMenuRef.current && !exportMenuRef.current.contains(e.target as Node)) {
        setExportMenuOpen(false)
      }
    }
    window.addEventListener('mousedown', onDocClick)
    return () => window.removeEventListener('mousedown', onDocClick)
  }, [])

  return (
    <div className="ribbon-group">
      <span className="rb-group-caption">Export</span>
      <div className="ribbon-group-items" ref={exportMenuRef}>
        <div className="rb-dropdown-container">
          <button
            type="button"
            className={`rb-big ${exportMenuOpen ? 'active' : ''}`}
            title="Export diagram"
            aria-haspopup="menu"
            aria-expanded={exportMenuOpen}
            onClick={() => setExportMenuOpen((v) => !v)}
          >
            <div className="rb-big-icon">
              <IconDownload size={20} />
            </div>
            <span>Export</span>
          </button>
          {exportMenuOpen && (
            <div className="rb-dropdown-menu">
              <button
                type="button"
                className="rb-dropdown-item"
                onClick={() => {
                  setExportMenuOpen(false)
                  onExportSvg()
                }}
              >
                <span>Export SVG</span>
                <span className="rb-hint">Vector</span>
              </button>
              <button
                type="button"
                className="rb-dropdown-item"
                onClick={() => {
                  setExportMenuOpen(false)
                  onExportImage('png')
                }}
              >
                <span>Export PNG</span>
                <span className="rb-hint">2× bitmap</span>
              </button>
              <button
                type="button"
                className="rb-dropdown-item"
                onClick={() => {
                  setExportMenuOpen(false)
                  onExportImage('jpeg')
                }}
              >
                <span>Export JPEG</span>
                <span className="rb-hint">2× bitmap</span>
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

function RibbonAiGroup({
  aiOpen,
  onToggleAi,
}: Pick<DiagramRibbonProps, 'aiOpen' | 'onToggleAi'>): ReactElement {
  return (
    <div className="ribbon-group ribbon-ai-group">
      <div className="ribbon-group-items">
        <button
          type="button"
          className={`rb-big ai-entry ${aiOpen ? 'active' : ''}`}
          title="Toggle ARA Assistant"
          aria-pressed={aiOpen}
          onClick={onToggleAi}
        >
          <span className="rb-big-icon">
            <AssistantMark size={26} />
          </span>
          <span>ARA Assistant</span>
        </button>
      </div>
    </div>
  )
}

type VisualRibbonGroupsProps = Pick<
  DiagramRibbonProps,
  | 'onSelectAll'
  | 'onDuplicate'
  | 'onDelete'
  | 'onClear'
  | 'hasSelection'
  | 'onInsertShape'
  | 'onConnect'
  | 'onOrder'
  | 'onAlign'
  | 'onDistribute'
  | 'selectionCount'
  | 'onAiPrompt'
  | 'gridEnabled'
  | 'showPropsPanel'
  | 'onTogglePropsPanel'
  | 'onToggleGrid'
  | 'onZoomOut'
  | 'onZoomActual'
  | 'onZoomIn'
  | 'onZoomFit'
  | 'zoomPercent'
  | 'aiOpen'
  | 'onToggleAi'
  | 'onExportSvg'
  | 'onExportImage'
  | 'lang'
>

function VisualRibbonGroups({
  onSelectAll,
  onDuplicate,
  onDelete,
  onClear,
  hasSelection,
  onInsertShape,
  onConnect,
  onOrder,
  onAlign,
  onDistribute,
  selectionCount,
  onAiPrompt,
  gridEnabled,
  showPropsPanel,
  onTogglePropsPanel,
  onToggleGrid,
  onZoomOut,
  onZoomActual,
  onZoomIn,
  onZoomFit,
  zoomPercent,
  aiOpen,
  onToggleAi,
  onExportSvg,
  onExportImage,
  lang,
}: VisualRibbonGroupsProps): ReactElement {
  const canConnect = selectionCount === 2
  const canAlign = selectionCount >= 2
  const canDistribute = selectionCount >= 3
  const gridButtonClass = gridEnabled ? 'rb-small active' : 'rb-small'
  const gridButtonLabel = gridEnabled ? 'Grid On' : 'Grid Off'
  const propsButtonClass = showPropsPanel ? 'rb-small active' : 'rb-small'

  return (
    <div className="ribbon-toolbar-row">
      <RibbonAiGroup aiOpen={aiOpen} onToggleAi={onToggleAi} />
      <div className="ribbon-sep" />
      <div className="ribbon-group">
        <span className="rb-group-caption">Edit</span>
        <div className="ribbon-group-items rb-col">
          <div className="rb-row">
            <button
              type="button"
              className="rb-small"
              title="Select all (Ctrl+A)"
              onClick={onSelectAll}
            >
              <IconSelectAll size={15} />
              <span>Select All</span>
            </button>
            <button
              type="button"
              className="rb-small"
              disabled={!hasSelection}
              title="Duplicate selection (Ctrl+D)"
              onClick={onDuplicate}
            >
              <IconDuplicate size={15} />
              <span>Duplicate</span>
            </button>
          </div>
          <div className="rb-row">
            <button
              type="button"
              className="rb-small rb-danger"
              disabled={!hasSelection}
              title="Delete selection (Del / Backspace)"
              onClick={onDelete}
            >
              <IconDelete size={15} />
              <span>Delete</span>
            </button>
            <button type="button" className="rb-small" title="Clear canvas" onClick={onClear}>
              <IconClear size={15} />
              <span>Clear</span>
            </button>
          </div>
        </div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <span className="rb-group-caption">Insert</span>
        <div className="ribbon-group-items">
          <div className="rb-shape-grid">
            {SHAPE_BUTTONS.map((shape) => {
              const help = shapeHelp(shape.type, lang)
              return (
                <button
                  key={shape.type}
                  type="button"
                  className="rb-shape-quick-btn"
                  title={`${help.label}: ${help.description}`}
                  aria-label={`${help.label}: ${help.description}`}
                  draggable
                  onDragStart={(event) => {
                    event.dataTransfer.effectAllowed = 'copy'
                    event.dataTransfer.setData(DIAGRAM_SHAPE_DRAG_MIME, shape.type)
                  }}
                  onClick={() => onInsertShape?.(shape.type)}
                >
                  {shape.icon}
                </button>
              )
            })}
          </div>
          <button
            type="button"
            className="rb-small"
            disabled={!canConnect}
            title="Connect the two selected shapes"
            onClick={onConnect}
          >
            <IconConnector size={15} />
            <span>Connect</span>
          </button>
        </div>
      </div>
      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <span className="rb-group-caption">Arrange</span>
        <div className="ribbon-group-items rb-col">
          <div className="rb-row">
            <button
              type="button"
              className="rb-small"
              title="Auto-layout the current diagram"
              onClick={() =>
                onAiPrompt?.(
                  'Organize and auto-layout the current diagram nodes neatly without overlapping',
                )
              }
            >
              <IconAutoLayout size={15} />
              <span>Auto-layout</span>
            </button>
            <button
              type="button"
              className="rb-small"
              disabled={!hasSelection}
              title="Bring selection to front"
              onClick={() => onOrder('front')}
            >
              <IconBringFront size={15} />
              <span>Front</span>
            </button>
            <button
              type="button"
              className="rb-small"
              disabled={!hasSelection}
              title="Send selection to back"
              onClick={() => onOrder('back')}
            >
              <IconSendBack size={15} />
              <span>Back</span>
            </button>
          </div>
          <div className="rb-row">
            {ALIGN_BUTTONS.map((item) => (
              <button
                key={item.direction}
                type="button"
                className="rb-icon-btn"
                disabled={!canAlign}
                title={`Align ${item.label.toLowerCase()}`}
                aria-label={`Align ${item.label.toLowerCase()}`}
                onClick={() => onAlign(item.direction)}
              >
                {item.icon}
              </button>
            ))}
            {DISTRIBUTE_BUTTONS.map((item) => (
              <button
                key={item.direction}
                type="button"
                className="rb-icon-btn"
                disabled={!canDistribute}
                title={`Distribute ${item.label.toLowerCase()}`}
                aria-label={`Distribute ${item.label.toLowerCase()}`}
                onClick={() => onDistribute(item.direction)}
              >
                {item.icon}
              </button>
            ))}
          </div>
        </div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <span className="rb-group-caption">View</span>
        <div className="ribbon-group-items">
          <div className="rb-col">
            <button
              type="button"
              className={gridButtonClass}
              title="Toggle grid and snap"
              aria-pressed={gridEnabled}
              onClick={onToggleGrid}
            >
              <IconGrid size={15} />
              <span>{gridButtonLabel}</span>
            </button>
            <button
              type="button"
              className={propsButtonClass}
              title="Toggle properties panel"
              aria-pressed={showPropsPanel}
              onClick={onTogglePropsPanel}
            >
              <IconProperties size={15} />
              <span>Properties</span>
            </button>
          </div>
          <div className="rb-zoom-grid">
            <button type="button" className="rb-icon-btn" title="Zoom out" onClick={onZoomOut}>
              <IconZoomOut size={15} />
            </button>
            <button type="button" className="rb-icon-btn" title="Zoom in" onClick={onZoomIn}>
              <IconZoomIn size={15} />
            </button>
            <button
              type="button"
              className="rb-small"
              title="Reset zoom to 100%"
              onClick={onZoomActual}
            >
              <span>{Math.round(zoomPercent)}%</span>
            </button>
            <button
              type="button"
              className="rb-icon-btn"
              title="Fit diagram to window"
              onClick={onZoomFit}
            >
              <IconFit size={15} />
            </button>
          </div>
        </div>
      </div>

      <div className="ribbon-sep" />

      <RibbonExportGroup onExportSvg={onExportSvg} onExportImage={onExportImage} />
    </div>
  )
}

type MermaidRibbonGroupsProps = Pick<
  DiagramRibbonProps,
  | 'onApplyTemplate'
  | 'onInsertMermaidSnippet'
  | 'aiOpen'
  | 'onToggleAi'
  | 'onExportSvg'
  | 'onExportImage'
>

function MermaidRibbonGroups({
  onApplyTemplate,
  onInsertMermaidSnippet,
  aiOpen,
  onToggleAi,
  onExportSvg,
  onExportImage,
}: MermaidRibbonGroupsProps): ReactElement {
  return (
    <div className="ribbon-toolbar-row">
      <RibbonAiGroup aiOpen={aiOpen} onToggleAi={onToggleAi} />
      <div className="ribbon-sep" />
      <div className="ribbon-group">
        <span className="rb-group-caption">Templates</span>
        <div className="ribbon-group-items">
          <div className="rb-template-grid">
            {MERMAID_TEMPLATES.map((template) => (
              <button
                key={template.label}
                type="button"
                className="rb-small"
                title={`Use ${template.label} template`}
                onClick={() => onApplyTemplate?.(template.code)}
              >
                {template.icon}
                <span>{template.label}</span>
              </button>
            ))}
          </div>
        </div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <span className="rb-group-caption">Insert</span>
        <div className="ribbon-group-items rb-syntax-grid">
          <button
            type="button"
            className="rb-small"
            title="Insert a Mermaid node"
            disabled={!onInsertMermaidSnippet}
            onClick={() => onInsertMermaidSnippet?.('    A[New node]')}
          >
            <IconCode size={15} />
            <span>Node</span>
          </button>
          <button
            type="button"
            className="rb-small"
            title="Insert a Mermaid connection"
            disabled={!onInsertMermaidSnippet}
            onClick={() => onInsertMermaidSnippet?.('    A --> B')}
          >
            <IconConnector size={15} />
            <span>Link</span>
          </button>
          <button
            type="button"
            className="rb-small"
            title="Insert a Mermaid subgraph"
            disabled={!onInsertMermaidSnippet}
            onClick={() => onInsertMermaidSnippet?.('    subgraph Group\n        A --> B\n    end')}
          >
            <IconFlowchart size={15} />
            <span>Subgraph</span>
          </button>
          <button
            type="button"
            className="rb-small"
            title="Insert a Mermaid comment"
            disabled={!onInsertMermaidSnippet}
            onClick={() => onInsertMermaidSnippet?.('    %% Add a note')}
          >
            <IconText size={15} />
            <span>Comment</span>
          </button>
        </div>
      </div>
      <div className="ribbon-sep" />

      <RibbonExportGroup onExportSvg={onExportSvg} onExportImage={onExportImage} />
    </div>
  )
}

export function Ribbon({
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
  selectionCount,
  onClear,
  hasSelection,
  onExportSvg,
  onExportImage,
  onSave,
  onSaveAs,
  dirty,
  filePath,
  autoSave,
  onAutoSaveChange,
  aiOpen,
  onToggleAi,
  onAiPrompt,
  showPropsPanel,
  onTogglePropsPanel,
  gridEnabled,
  onToggleGrid,
  onInsertShape,
  onApplyTemplate,
  onInsertMermaidSnippet,
  lang,
}: DiagramRibbonProps): ReactElement {
  const saveDisabled = Boolean(filePath) && !dirty
  const visualGroupProps: VisualRibbonGroupsProps = {
    onSelectAll,
    onDuplicate,
    onDelete,
    onClear,
    hasSelection,
    onInsertShape,
    onConnect,
    onOrder,
    onAlign,
    onDistribute,
    selectionCount,
    onAiPrompt,
    gridEnabled,
    showPropsPanel,
    onTogglePropsPanel,
    onToggleGrid,
    onZoomOut,
    onZoomActual,
    onZoomIn,
    onZoomFit,
    zoomPercent,
    aiOpen,
    onToggleAi,
    onExportSvg,
    onExportImage,
    lang,
  }
  const mermaidGroupProps: MermaidRibbonGroupsProps = {
    onApplyTemplate,
    onInsertMermaidSnippet,
    aiOpen,
    onToggleAi,
    onExportSvg,
    onExportImage,
  }

  return (
    <div className="ribbon">
      <RibbonTopbar
        mode={mode}
        canUndo={canUndo}
        canRedo={canRedo}
        onUndo={onUndo}
        onRedo={onRedo}
        onSave={onSave}
        onSaveAs={onSaveAs}
        onExportSvg={onExportSvg}
        onExportImage={onExportImage}
        onClear={onClear}
        autoSave={autoSave}
        onAutoSaveChange={onAutoSaveChange}
        saveDisabled={saveDisabled}
      />

      <div className="ribbon-body">
        {mode === 'visual' ? (
          <VisualRibbonGroups {...visualGroupProps} />
        ) : (
          <MermaidRibbonGroups {...mermaidGroupProps} />
        )}
      </div>
    </div>
  )
}
