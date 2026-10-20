export interface StatusBarProps {
  selectedCount: number
  gridEnabled: boolean
  selectedType?: 'vertex' | 'edge' | 'mixed'
  zoomPercent: number
  mode: 'visual' | 'mermaid'
  onZoomIn?: () => void
  onZoomOut?: () => void
  onZoomChange?: (zoom: number) => void
}

export function StatusBar({
  selectedCount,
  selectedType,
  zoomPercent,
  mode: _mode,
  gridEnabled,
  onZoomIn,
  onZoomOut,
  onZoomChange,
}: StatusBarProps) {
  let statusText = 'Canvas Ready'
  if (selectedCount === 1) {
    statusText =
      selectedType === 'vertex'
        ? '1 Shape Selected'
        : selectedType === 'edge'
          ? '1 Connector Selected'
          : '1 Item Selected'
  } else if (selectedCount > 1) {
    statusText = `${selectedCount} Elements Selected`
  }

  return (
    <footer className="status-bar">
      <div className="status-left">
        <span className="diagram-status-dot" />
        <span className="status-item">{statusText}</span>
        <span className="status-item diagram-subtle-info">
          Grid: {gridEnabled ? '10px' : 'Off'}
        </span>
        <span className="status-item diagram-subtle-info">Snap: {gridEnabled ? 'On' : 'Off'}</span>
      </div>

      <div className="status-right">
        {onZoomOut && (
          <button
            type="button"
            className="zoom-btn"
            onClick={onZoomOut}
            title="Zoom Out"
            aria-label="Zoom Out"
          >
            -
          </button>
        )}
        {onZoomChange && (
          <input
            className="zoom-slider"
            type="range"
            min={25}
            max={200}
            step={5}
            value={Math.round(zoomPercent)}
            onChange={(e) => onZoomChange(Number(e.target.value))}
            aria-label="Zoom Level"
          />
        )}
        {onZoomIn && (
          <button
            type="button"
            className="zoom-btn"
            onClick={onZoomIn}
            title="Zoom In"
            aria-label="Zoom In"
          >
            +
          </button>
        )}
        <span className="zoom-value">{Math.round(zoomPercent)}%</span>
      </div>
    </footer>
  )
}
