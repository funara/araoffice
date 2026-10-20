import type { Cell } from '@maxgraph/core'

export interface PropertiesPanelProps {
  selectedCell: Cell | null
  onUpdateLabel: (label: string) => void
  onUpdateStyle: (key: string, value: string | number | boolean) => void
  onCollapse: () => void
}

const FILL_PALETTE = [
  '#ffffff',
  '#f8fafc',
  '#e0f2fe',
  '#dcfce7',
  '#fef3c7',
  '#fee2e2',
  '#f3e8ff',
  '#e2e8f0',
]

const STROKE_PALETTE = [
  '#334155',
  '#0f172a',
  '#0284c7',
  '#16a34a',
  '#d97706',
  '#dc2626',
  '#9333ea',
  '#64748b',
]

type PropertyValue = string | number

interface ColorPaletteProps {
  colors: readonly string[]
  activeColor: string
  onSelect: (color: string) => void
}

function ColorPalette({ colors, activeColor, onSelect }: ColorPaletteProps) {
  return (
    <div className="diagram-color-palette">
      {colors.map((color) => (
        <button
          key={color}
          type="button"
          className={`diagram-color-swatch swatch-${color.slice(1)} ${activeColor === color ? 'active' : ''}`}
          onClick={() => onSelect(color)}
          title={color}
        />
      ))}
    </div>
  )
}

interface OptionGroupProps {
  options: readonly { value: PropertyValue; label: string }[]
  activeValue: PropertyValue
  onSelect: (value: PropertyValue) => void
}

function OptionGroup({ options, activeValue, onSelect }: OptionGroupProps) {
  return (
    <div className="diagram-btn-group">
      {options.map((option) => (
        <button
          key={String(option.value)}
          type="button"
          className={`diagram-opt-btn ${activeValue === option.value ? 'active' : ''}`}
          onClick={() => onSelect(option.value)}
        >
          {option.label}
        </button>
      ))}
    </div>
  )
}

const STROKE_WIDTH_OPTIONS = [
  { value: 1, label: '1px' },
  { value: 2, label: '2px' },
  { value: 3, label: '3px' },
  { value: 4, label: '4px' },
]

const FONT_SIZE_OPTIONS = [10, 12, 13, 14, 16, 18, 20, 24].map((size) => ({
  value: size,
  label: `${size}px`,
}))

const ALIGN_OPTIONS = [
  { value: 'left', label: 'Left' },
  { value: 'center', label: 'Center' },
  { value: 'right', label: 'Right' },
]

const EDGE_STYLE_OPTIONS = [
  { value: 'orthogonalEdgeStyle', label: 'Elbow' },
  { value: 'none', label: 'Straight' },
  { value: 'entityRelationEdgeStyle', label: 'Curved' },
]

const ARROW_OPTIONS = [
  { value: 'classic', label: 'Classic' },
  { value: 'block', label: 'Block' },
  { value: 'oval', label: 'Oval' },
  { value: 'none', label: 'None' },
]

type SelectedPropertiesProps = Pick<
  PropertiesPanelProps,
  'onUpdateLabel' | 'onUpdateStyle' | 'onCollapse'
> & {
  selectedCell: Cell
}

interface SelectedPropertiesFieldsProps {
  isEdge: boolean
  isVertex: boolean
  label: string
  currentFill: string
  currentStroke: string
  currentStrokeWidth: number
  currentFontSize: number
  currentFontWeight: string
  currentAlign: string
  currentEdgeStyle: string
  currentEndArrow: string
  onUpdateLabel: (label: string) => void
  onUpdateStyle: (key: string, value: string | number | boolean) => void
}

function SelectedPropertiesFields({
  isEdge,
  isVertex,
  label,
  currentFill,
  currentStroke,
  currentStrokeWidth,
  currentFontSize,
  currentFontWeight,
  currentAlign,
  currentEdgeStyle,
  currentEndArrow,
  onUpdateLabel,
  onUpdateStyle,
}: SelectedPropertiesFieldsProps) {
  return (
    <>
      <div className="diagram-prop-field">
        <label className="diagram-prop-label">Label</label>
        <input
          type="text"
          className="diagram-prop-input"
          value={label}
          onChange={(e) => onUpdateLabel(e.target.value)}
          placeholder="Element label..."
        />
      </div>

      {isVertex && (
        <div className="diagram-prop-field">
          <label className="diagram-prop-label">Fill Color</label>
          <ColorPalette
            colors={FILL_PALETTE}
            activeColor={currentFill}
            onSelect={(color) => onUpdateStyle('fillColor', color)}
          />
        </div>
      )}

      <div className="diagram-prop-field">
        <label className="diagram-prop-label">Line Color</label>
        <ColorPalette
          colors={STROKE_PALETTE}
          activeColor={currentStroke}
          onSelect={(color) => onUpdateStyle('strokeColor', color)}
        />
      </div>

      <div className="diagram-prop-field">
        <label className="diagram-prop-label">Line Width</label>
        <OptionGroup
          options={STROKE_WIDTH_OPTIONS}
          activeValue={currentStrokeWidth}
          onSelect={(value) => onUpdateStyle('strokeWidth', value)}
        />
      </div>

      <div className="diagram-prop-section-title">Typography</div>

      <div className="diagram-prop-row">
        <div className="diagram-prop-field flex-1">
          <label className="diagram-prop-label">Size</label>
          <select
            className="diagram-prop-select"
            value={currentFontSize}
            onChange={(e) => onUpdateStyle('fontSize', Number(e.target.value))}
          >
            {FONT_SIZE_OPTIONS.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>

        <div className="diagram-prop-field flex-1">
          <label className="diagram-prop-label">Weight</label>
          <button
            type="button"
            className={`diagram-opt-btn ${currentFontWeight === 'bold' ? 'active' : ''}`}
            onClick={() => {
              const isBold = currentFontWeight === 'bold'
              onUpdateStyle('fontStyle', isBold ? 0 : 1)
              onUpdateStyle('fontWeight', isBold ? 'normal' : 'bold')
            }}
          >
            Bold
          </button>
        </div>
      </div>

      <div className="diagram-prop-field">
        <label className="diagram-prop-label">Text Align</label>
        <OptionGroup
          options={ALIGN_OPTIONS}
          activeValue={currentAlign}
          onSelect={(value) => onUpdateStyle('align', value)}
        />
      </div>

      {isEdge && (
        <>
          <div className="diagram-prop-section-title">Connector Routing</div>
          <div className="diagram-prop-field">
            <label className="diagram-prop-label">Style</label>
            <OptionGroup
              options={EDGE_STYLE_OPTIONS}
              activeValue={currentEdgeStyle}
              onSelect={(value) => onUpdateStyle('edgeStyle', value)}
            />
          </div>

          <div className="diagram-prop-field">
            <label className="diagram-prop-label">Arrow Head</label>
            <OptionGroup
              options={ARROW_OPTIONS}
              activeValue={currentEndArrow}
              onSelect={(value) => onUpdateStyle('endArrow', value)}
            />
          </div>
        </>
      )}
    </>
  )
}

function SelectedProperties({
  selectedCell,
  onUpdateLabel,
  onUpdateStyle,
  onCollapse,
}: SelectedPropertiesProps) {
  const isEdge = Boolean(selectedCell.isEdge())
  const isVertex = Boolean(selectedCell.isVertex())
  const style = (selectedCell.style || {}) as Record<string, unknown>
  const label = typeof selectedCell.value === 'string' ? selectedCell.value : ''
  const currentFill = String(style.fillColor || '#ffffff')
  const currentStroke = String(style.strokeColor || '#334155')
  const currentStrokeWidth = Number(style.strokeWidth || 2)
  const currentFontSize = Number(style.fontSize || 13)
  const currentFontWeight = style.fontStyle === 1 || style.fontWeight === 'bold' ? 'bold' : 'normal'
  const currentAlign = String(style.align || 'center')
  const currentEdgeStyle = String(style.edgeStyle || 'orthogonalEdgeStyle')
  const currentEndArrow = String(style.endArrow || 'classic')
  let panelTitle = 'Element Properties'
  if (isVertex) {
    panelTitle = 'Shape Properties'
  } else if (isEdge) {
    panelTitle = 'Connector Properties'
  }

  return (
    <aside className="diagram-props-panel">
      <div className="diagram-props-header">
        <span>{panelTitle}</span>
        <button
          type="button"
          className="diagram-panel-collapse"
          title="Collapse properties panel"
          aria-label="Collapse properties panel"
          onClick={onCollapse}
        >
          ‹
        </button>
      </div>
      <SelectedPropertiesFields
        isEdge={isEdge}
        isVertex={isVertex}
        label={label}
        currentFill={currentFill}
        currentStroke={currentStroke}
        currentStrokeWidth={currentStrokeWidth}
        currentFontSize={currentFontSize}
        currentFontWeight={currentFontWeight}
        currentAlign={currentAlign}
        currentEdgeStyle={currentEdgeStyle}
        currentEndArrow={currentEndArrow}
        onUpdateLabel={onUpdateLabel}
        onUpdateStyle={onUpdateStyle}
      />
    </aside>
  )
}
export function PropertiesPanel({
  selectedCell,
  onUpdateLabel,
  onUpdateStyle,
  onCollapse,
}: PropertiesPanelProps) {
  if (!selectedCell) {
    return (
      <aside className="diagram-props-panel">
        <div className="diagram-props-header">
          <span>Canvas Properties</span>
          <button
            type="button"
            className="diagram-panel-collapse"
            title="Collapse properties panel"
            aria-label="Collapse properties panel"
            onClick={onCollapse}
          >
            ‹
          </button>
        </div>
        <div className="diagram-props-empty-hint">
          Select a shape or connector to edit its styles and label.
        </div>
      </aside>
    )
  }

  return (
    <SelectedProperties
      selectedCell={selectedCell}
      onUpdateLabel={onUpdateLabel}
      onUpdateStyle={onUpdateStyle}
      onCollapse={onCollapse}
    />
  )
}
