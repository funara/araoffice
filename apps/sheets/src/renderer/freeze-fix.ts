import type { IRenderManagerService } from '@univerjs/engine-render'
import { IRenderManagerService as IRenderManagerServiceToken } from '@univerjs/engine-render'
import { HeaderFreezeRenderController, ISheetSelectionRenderService } from '@univerjs/sheets-ui'
import type { UniverRuntime } from './univer-state'

// ============================================================================
// Monkey-patch HeaderFreezeRenderController.prototype at module load time.
// This guarantees that across every sheet, workbook, and skeleton re-render:
// 1. When unfrozen (default), the thick corner drag handles are NEVER created.
// 2. Interactive canvas dragging to freeze rows/columns is 100% disabled.
// 3. The top-left corner intersection remains clean, fixed, and static (Excel-parity).
// ============================================================================

interface FreezeConfig {
  startRow: number
  startColumn: number
  xSplit?: number
  ySplit?: number
}

interface FreezeControllerInternal {
  _getFreeze(): FreezeConfig | null | undefined
  _createFreeze(type?: number, config?: unknown): unknown
  _eventBinding(type?: number): void
  _freezeDown(...args: unknown[]): boolean
}

const proto = HeaderFreezeRenderController.prototype as unknown as FreezeControllerInternal

// 1. Permanently disable canvas drag interaction on freeze handles
proto._freezeDown = (): boolean => false

// 2. Do not bind any pointer down / hover listeners to freeze handles
proto._eventBinding = (): void => {
  // no-op: disable interactive dragging on freeze handles completely
}

// 3. Only draw freeze lines when there are actual frozen panes (e.g. set from View tab).
// When unfrozen (freezeRow <= 0, freezeColumn <= 0), do NOT draw the thick corner handles at all.
const origCreateFreeze = proto._createFreeze
proto._createFreeze = function (freezeDirectionType = 0, freezeConfig?: unknown): unknown {
  const config = (freezeConfig as FreezeConfig | null | undefined) ?? this._getFreeze()
  if (!config) return null
  const { startRow: freezeRow, startColumn: freezeColumn } = config
  if (freezeDirectionType === 0 && (freezeRow === -1 || freezeRow === 0)) {
    return null
  }
  if (freezeDirectionType === 1 && (freezeColumn === -1 || freezeColumn === 0)) {
    return null
  }
  return origCreateFreeze.call(this, freezeDirectionType, freezeConfig)
}

/**
 * Installs additional runtime guards for cell selection drag and render events.
 */
export function installFreezeHandleFix(runtime: UniverRuntime): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const renderManager = injector.get(IRenderManagerServiceToken) as IRenderManagerService

  // Disable dragging/moving cell ranges via border drag
  try {
    const selectionRenderService = injector.get(ISheetSelectionRenderService)
    if (selectionRenderService) {
      selectionRenderService.interceptor.intercept(
        selectionRenderService.interceptor.getInterceptPoints().RANGE_MOVE_PERMISSION_CHECK,
        {
          handler: () => false,
        },
      )
    }
  } catch {
    // ignore if selection render service not registered yet
  }

  const freezeRowHeaderName = '__SpreadsheetFreezeRowHeaderName__'
  const freezeColHeaderName = '__SpreadsheetFreezeColumnHeaderName__'

  const applyToRender = (unitId: string): void => {
    const render = renderManager.getRenderById(unitId)
    if (!render) return

    const freezeController = render.with(HeaderFreezeRenderController)
    if (freezeController) {
      freezeController.interceptor.intercept(
        freezeController.interceptor.getInterceptPoints().FREEZE_PERMISSION_CHECK,
        {
          handler: () => false,
        },
      )
    }

    const cleanScene = (): void => {
      try {
        const scene = render.scene
        if (!scene) return
        const activeWorkbook = runtime.univerAPI.getActiveWorkbook()
        const activeSheet = activeWorkbook?.getActiveSheet()
        const freeze = activeSheet?.getFreeze?.()
        const isUnfrozen = !freeze?.xSplit && !freeze?.ySplit

        if (isUnfrozen) {
          const rH = scene.getObject(freezeRowHeaderName) as {
            hide?(): void
            setProps?(p: Record<string, unknown>): void
          } | null
          const cH = scene.getObject(freezeColHeaderName) as {
            hide?(): void
            setProps?(p: Record<string, unknown>): void
          } | null
          rH?.hide?.()
          cH?.hide?.()
          rH?.setProps?.({ visible: false })
          cH?.setProps?.({ visible: false })
        }
      } catch {
        // ignore if objects not ready
      }
    }

    cleanScene()
    setTimeout(cleanScene, 50)
    setTimeout(cleanScene, 200)
  }

  const sub = renderManager.createRender$?.subscribe((unitId: string) => {
    if (unitId) {
      applyToRender(unitId)
    }
  })

  const activeWorkbook = runtime.univerAPI.getActiveWorkbook()
  if (activeWorkbook) {
    applyToRender(activeWorkbook.getId())
  }

  return {
    dispose() {
      sub?.unsubscribe()
    },
  }
}
