import { beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * TabManager (src/main/tab-manager.ts): tab list state, activation,
 * close guards, and view lifecycle inside the shell's single window.
 * Electron and the per-module main entrypoints are mocked; only the
 * manager's own observable behavior is asserted.
 */

interface FakeWebContents {
  id: number
  on: ReturnType<typeof vi.fn>
  removeListener: ReturnType<typeof vi.fn>
  close: ReturnType<typeof vi.fn>
  forcefullyCrashRenderer: ReturnType<typeof vi.fn>
  isDestroyed: ReturnType<typeof vi.fn>
  setBackgroundThrottling: ReturnType<typeof vi.fn>
  loadURL: ReturnType<typeof vi.fn>
  invalidate?: ReturnType<typeof vi.fn>
  isLoading: ReturnType<typeof vi.fn>
  focus: ReturnType<typeof vi.fn>
  send: ReturnType<typeof vi.fn>
  capturePage?: ReturnType<typeof vi.fn>
}

interface FakeView {
  webContents: FakeWebContents
  setVisible: ReturnType<typeof vi.fn>
  setBounds: ReturnType<typeof vi.fn>
}

let nextWebContentsId = 1

function makeFakeView(): FakeView {
  const listeners = new Map<string, (...args: unknown[]) => void>()
  return {
    webContents: {
      id: nextWebContentsId++,
      listeners,
      on: vi.fn((event: string, handler: (...args: unknown[]) => void) => {
        listeners.set(event, handler)
      }),
      removeListener: vi.fn((event: string) => {
        listeners.delete(event)
      }),
      close: vi.fn(),
      forcefullyCrashRenderer: vi.fn(),
      isDestroyed: vi.fn(() => false),
      setBackgroundThrottling: vi.fn(),
      loadURL: vi.fn(() => Promise.resolve()),
      invalidate: vi.fn(),
      isLoading: vi.fn(() => false),
      focus: vi.fn(),
      send: vi.fn(),
      capturePage: vi.fn(() =>
        Promise.resolve({
          isEmpty: () => false,
          getSize: () => ({ width: 1920, height: 1080 }),
          resize: vi.fn(() => ({
            toJPEG: () => Buffer.from('mock-jpeg'),
          })),
          toJPEG: vi.fn(() => Buffer.from('mock-jpeg')),
        }),
      ),
    },
    setVisible: vi.fn(),
    setBounds: vi.fn(),
  }
}

vi.mock('electron', () => ({
  BrowserWindow: class {},
  app: { setName: vi.fn(), getName: vi.fn(() => 'ARA Office') },
  powerMonitor: { on: vi.fn(), removeListener: vi.fn() },
}))

const createDiagramView = vi.fn(() => makeFakeView())
const diagramIsDirty = vi.fn(() => false)
const requestDiagramClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../diagram/src/main/diagram-main', () => ({
  createDiagramView: (...args: unknown[]) => createDiagramView(...(args as [])),
  diagramIsDirty: (...args: unknown[]) => diagramIsDirty(...(args as [])),
  requestDiagramClose: (...args: unknown[]) => requestDiagramClose(...(args as [])),
}))

const createDocsView = vi.fn(() => makeFakeView())
const docsQueryDirty = vi.fn(() => Promise.resolve(false))
const markDocsNewBlank = vi.fn()
const requestDocsClose = vi.fn(() => Promise.resolve(true))
const setActiveDocsResolver = vi.fn()
const teardownDocsRenderer = vi.fn()

vi.mock('../../docs/src/main/docs-main', () => ({
  createDocsView: (...args: unknown[]) => createDocsView(...(args as [])),
  docsQueryDirty: (...args: unknown[]) => docsQueryDirty(...(args as [])),
  markDocsNewBlank: (...args: unknown[]) => markDocsNewBlank(...args),
  requestDocsClose: (...args: unknown[]) => requestDocsClose(...(args as [])),
  setActiveDocsResolver: (...args: unknown[]) => setActiveDocsResolver(...args),
  teardownDocsRenderer: (...args: unknown[]) => teardownDocsRenderer(...args),
}))

const createPdfView = vi.fn(() => makeFakeView())
const pdfIsDirty = vi.fn(() => false)
const requestPdfClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../pdf/src/main/pdf-main', () => ({
  createPdfView: (...args: unknown[]) => createPdfView(...(args as [])),
  pdfIsDirty: (...args: unknown[]) => pdfIsDirty(...(args as [])),
  requestPdfClose: (...args: unknown[]) => requestPdfClose(...(args as [])),
}))

const createSheetsView = vi.fn(() => makeFakeView())
const queueSheetsLaunch = vi.fn()
const requestSheetsClose = vi.fn(() => Promise.resolve(true))
const setActiveSheetsWebContents = vi.fn()
const sheetsPendingEditCount = vi.fn(() => 0)

vi.mock('../../sheets/src/main/sheets-main', () => ({
  createSheetsView: (...args: unknown[]) => createSheetsView(...(args as [])),
  queueSheetsLaunch: (...args: unknown[]) => queueSheetsLaunch(...args),
  requestSheetsClose: (...args: unknown[]) => requestSheetsClose(...(args as [])),
  setActiveSheetsWebContents: (...args: unknown[]) => setActiveSheetsWebContents(...args),
  sheetsPendingEditCount: (...args: unknown[]) => sheetsPendingEditCount(...(args as [])),
}))

const createSlidesView = vi.fn(() => makeFakeView())
const requestSlidesClose = vi.fn(() => Promise.resolve(true))
const setActiveSlidesWebContents = vi.fn()
const slidesIsDirty = vi.fn(() => false)

vi.mock('../../slides/src/main/slides-main', () => ({
  createSlidesView: (...args: unknown[]) => createSlidesView(...(args as [])),
  requestSlidesClose: (...args: unknown[]) => requestSlidesClose(...(args as [])),
  setActiveSlidesWebContents: (...args: unknown[]) => setActiveSlidesWebContents(...args),
  slidesIsDirty: (...args: unknown[]) => slidesIsDirty(...(args as [])),
}))

import { TabManager } from '../src/main/tab-manager'

const TAB_STRIP_HEIGHT = 40
const WINDOW_WIDTH = 800
const WINDOW_HEIGHT = 600

interface FakeShellWindow {
  on: ReturnType<typeof vi.fn>
  isDestroyed: ReturnType<typeof vi.fn>
  getContentBounds: () => { x: number; y: number; width: number; height: number }
  contentView: {
    addChildView: ReturnType<typeof vi.fn>
    removeChildView: ReturnType<typeof vi.fn>
  }
}

function makeShellWindow(): FakeShellWindow {
  return {
    on: vi.fn(),
    isDestroyed: vi.fn(() => false),
    getContentBounds: () => ({ x: 0, y: 0, width: WINDOW_WIDTH, height: WINDOW_HEIGHT }),
    contentView: { addChildView: vi.fn(), removeChildView: vi.fn() },
  }
}

let shellWindow: FakeShellWindow
let onChanged: ReturnType<typeof vi.fn>
let applyMenuFor: ReturnType<typeof vi.fn>
let manager: TabManager

function lastCreatedView(factory: ReturnType<typeof vi.fn>): FakeView {
  return factory.mock.results.at(-1)!.value as FakeView
}

beforeEach(() => {
  vi.clearAllMocks()
  nextWebContentsId = 1
  docsQueryDirty.mockImplementation(() => Promise.resolve(false))
  requestDocsClose.mockImplementation(() => Promise.resolve(true))
  pdfIsDirty.mockImplementation(() => false)
  sheetsPendingEditCount.mockImplementation(() => 0)
  slidesIsDirty.mockImplementation(() => false)
  diagramIsDirty.mockImplementation(() => false)
  requestDiagramClose.mockImplementation(() => Promise.resolve(true))
  shellWindow = makeShellWindow()
  onChanged = vi.fn()
  applyMenuFor = vi.fn()
  manager = new TabManager(
    shellWindow as never,
    () => onChanged(),
    (kind) => applyMenuFor(kind),
  )
})

describe('initial state', () => {
  it('starts with only the non-closable, active Home tab', () => {
    expect(manager.list()).toEqual([
      { id: 'home', kind: 'home', title: 'Home', closable: false, active: true },
    ])
  })
})

describe('opening tabs', () => {
  it('opens a docs tab, activates it, and attaches its view to the window', () => {
    const id = manager.openDocsTab()
    const tabs = manager.list()
    expect(tabs).toHaveLength(2)
    expect(tabs[1]).toMatchObject({
      id,
      kind: 'docs',
      title: 'ARAOffice Docs',
      closable: true,
      active: true,
    })
    expect(tabs[0].active).toBe(false)
    expect(shellWindow.contentView.addChildView).toHaveBeenCalledTimes(1)
    expect(applyMenuFor).toHaveBeenLastCalledWith('docs')
    expect(onChanged).toHaveBeenCalled()
  })

  it('titles file-backed tabs with the file basename', () => {
    manager.openDocsTab('/tmp/report.docx')
    manager.openSheetsTab('/tmp/budget.xlsx')
    manager.openSlidesTab('/tmp/deck.pptx')
    manager.openPdfTab('/tmp/scan.pdf')
    expect(manager.list().map((t) => t.title)).toEqual([
      'Home',
      'report.docx',
      'budget.xlsx',
      'deck.pptx',
      'scan.pdf',
    ])
  })

  it('uses module default titles for pathless tabs', () => {
    manager.openSheetsTab()
    manager.openSlidesTab()
    expect(manager.list().map((t) => t.title)).toEqual(['Home', 'ARA Sheets', 'ARA Slides'])
  })

  it('forwards the selected mode and titles when opening diagram tabs', () => {
    const codeId = manager.openDiagramTab(undefined, 'mermaid')
    expect(createDiagramView).toHaveBeenCalledWith(undefined, 'mermaid')
    const codeTab = manager.list().find((t) => t.id === codeId)
    expect(codeTab?.title).toBe('Untitled Diagram Code')
    expect(codeTab?.diagramMode).toBe('mermaid')

    const visualId = manager.openDiagramTab(undefined, 'visual')
    expect(createDiagramView).toHaveBeenCalledWith(undefined, 'visual')
    const visualTab = manager.list().find((t) => t.id === visualId)
    expect(visualTab?.title).toBe('Untitled Diagram')
    expect(visualTab?.diagramMode).toBe('visual')
  })

  it('assigns unique, monotonic tab ids', () => {
    const a = manager.openDocsTab()
    const b = manager.openSheetsTab()
    expect(a).not.toBe(b)
    expect(a).toBe('t1')
    expect(b).toBe('t2')
  })

  it('queues each Sheets launch on its own renderer', () => {
    manager.openSheetsTab('/tmp/first.xlsx')
    const firstView = lastCreatedView(createSheetsView)
    manager.openSheetsTab('/tmp/second.xlsx')
    const secondView = lastCreatedView(createSheetsView)

    expect(queueSheetsLaunch).toHaveBeenNthCalledWith(1, firstView.webContents, {
      path: '/tmp/first.xlsx',
    })
    expect(queueSheetsLaunch).toHaveBeenNthCalledWith(2, secondView.webContents, {
      path: '/tmp/second.xlsx',
    })
  })

  it('forwards a new sheet save-as launch to its Sheets renderer', () => {
    manager.openDocsTab(undefined, { newBlank: true })
    expect(markDocsNewBlank).toHaveBeenCalledTimes(1)
    manager.openSheetsTab(undefined, {
      newBlank: true,
      queuedOpenPath: '/tmp/new-sheet.xlsx',
      suggestSaveAs: '/tmp/Untitled.xlsx',
    })
    expect(queueSheetsLaunch).toHaveBeenLastCalledWith(
      lastCreatedView(createSheetsView).webContents,
      {
        path: '/tmp/new-sheet.xlsx',
        suggestSaveAs: '/tmp/Untitled.xlsx',
        newBlank: true,
      },
    )
  })
})

describe('activation', () => {
  it('shows only the activated tab view and lays it out below the tab strip', () => {
    const docsId = manager.openDocsTab()
    const docsView = lastCreatedView(createDocsView)
    manager.openSheetsTab()
    const sheetsView = lastCreatedView(createSheetsView)

    manager.activateTab(docsId)
    expect(docsView.setVisible).toHaveBeenLastCalledWith(true)
    expect(sheetsView.setVisible).toHaveBeenLastCalledWith(false)
    expect(docsView.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT - TAB_STRIP_HEIGHT,
    })
    expect(manager.list().find((t) => t.id === docsId)?.active).toBe(true)
  })

  it('ignores activation of unknown tab ids', () => {
    onChanged.mockClear()
    manager.activateTab('nope')
    expect(onChanged).not.toHaveBeenCalled()
    expect(manager.list()[0].active).toBe(true)
  })

  it('routes the active webContents to the matching module', () => {
    manager.openSheetsTab()
    const sheetsView = lastCreatedView(createSheetsView)
    expect(setActiveSheetsWebContents).toHaveBeenLastCalledWith(sheetsView.webContents)

    manager.openSlidesTab()
    const slidesView = lastCreatedView(createSlidesView)
    expect(setActiveSlidesWebContents).toHaveBeenLastCalledWith(slidesView.webContents)
  })

  it('lets the active view cover the tab strip during HTML fullscreen', () => {
    manager.openSlidesTab()
    const view = lastCreatedView(createSlidesView)
    view.webContents.listeners.get('enter-html-full-screen')!()
    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: 0,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT,
    })
    view.webContents.listeners.get('leave-html-full-screen')!()
    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT - TAB_STRIP_HEIGHT,
    })
  })
})

describe('window resize layout', () => {
  function resizeHandler(): () => void {
    const call = shellWindow.on.mock.calls.find((c) => c[0] === 'resize')
    expect(call).toBeDefined()
    return call![1] as () => void
  }

  it('re-lays out after resize bounds settle (Linux/X11 stale getContentBounds)', async () => {
    // On X11, `resize` fires before the WM applies maximize bounds, so the first
    // layout still sees the pre-maximize size. The deferred layout must pick up
    // the real size on the next turn (see issue #15).
    manager.openSheetsTab()
    const view = lastCreatedView(createSheetsView)
    view.setBounds.mockClear()

    let width = WINDOW_WIDTH
    let height = WINDOW_HEIGHT
    shellWindow.getContentBounds = () => ({ x: 0, y: 0, width, height })

    resizeHandler()()
    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT - TAB_STRIP_HEIGHT,
    })

    // Bounds update after the synchronous layout, as on X11 maximize.
    width = 1920
    height = 1080
    await new Promise<void>((resolve) => setImmediate(resolve))

    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: 1920,
      height: 1080 - TAB_STRIP_HEIGHT,
    })
    expect(view.setBounds).toHaveBeenCalledTimes(2)
  })

  it('skips deferred layout after the shell window is destroyed', async () => {
    manager.openSheetsTab()
    const view = lastCreatedView(createSheetsView)
    view.setBounds.mockClear()

    resizeHandler()()
    expect(view.setBounds).toHaveBeenCalledTimes(1)

    shellWindow.isDestroyed.mockReturnValue(true)
    await new Promise<void>((resolve) => setImmediate(resolve))
    expect(view.setBounds).toHaveBeenCalledTimes(1)
  })

  it('forces visibility toggle and repaint on focus to recover dropped swapchains', () => {
    manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    view.setVisible.mockClear()
    view.webContents.invalidate?.mockClear()

    const focusHandler = shellWindow.on.mock.calls.find((c) => c[0] === 'focus')?.[1] as () => void
    expect(focusHandler).toBeDefined()
    focusHandler()

    expect(view.setVisible).toHaveBeenCalledWith(false)
    expect(view.setVisible).toHaveBeenCalledWith(true)
    expect(view.webContents.invalidate).toHaveBeenCalled()
  })
})

describe('closing tabs', () => {
  it('never closes the Home tab', async () => {
    await manager.closeTab('home')
    expect(manager.list()).toHaveLength(1)

    manager.openHomeTab()
    manager.closeActiveTab()
    await Promise.resolve()
    expect(manager.list()).toHaveLength(1)
  })

  it('removes a clean tab and falls back to the previous tab', async () => {
    manager.openSheetsTab()
    const sheetsView = lastCreatedView(createSheetsView)
    const slidesId = manager.openSlidesTab()

    await manager.closeTab(slidesId)
    const tabs = manager.list()
    expect(tabs.map((t) => t.id)).toEqual(['home', 't1'])
    expect(tabs[1].active).toBe(true)
    expect(sheetsView.setVisible).toHaveBeenLastCalledWith(true)
  })

  it('keeps the current tab active when closing a background tab', async () => {
    const sheetsId = manager.openSheetsTab()
    const slidesId = manager.openSlidesTab()
    await manager.closeTab(sheetsId)
    expect(manager.list().find((t) => t.id === slidesId)?.active).toBe(true)
  })

  it('detaches and destroys non-docs views on close', async () => {
    const id = manager.openSheetsTab()
    const view = lastCreatedView(createSheetsView)
    await manager.closeTab(id)
    expect(shellWindow.contentView.removeChildView).toHaveBeenCalledWith(view)
    expect(view.webContents.close).toHaveBeenCalledTimes(1)
  })

  it('detaches docs views without destroying the webContents (freeze workaround)', async () => {
    const id = manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    await manager.closeTab(id)
    expect(shellWindow.contentView.removeChildView).toHaveBeenCalledWith(view)
    expect(view.webContents.close).not.toHaveBeenCalled()
    // the orphaned renderer must be told to go inert (recovery-copy resurrection guard)
    expect(teardownDocsRenderer).toHaveBeenCalledWith(view.webContents)
    expect(view.webContents.forcefullyCrashRenderer).toHaveBeenCalledTimes(1)
  })

  it('closes a clean docs tab after the async dirty query says clean', async () => {
    const id = manager.openDocsTab()
    await manager.closeTab(id)
    expect(docsQueryDirty).toHaveBeenCalledTimes(1)
    expect(requestDocsClose).not.toHaveBeenCalled()
    expect(manager.list()).toHaveLength(1)
  })

  it('keeps a dirty docs tab open when the user cancels the close guard', async () => {
    docsQueryDirty.mockImplementation(() => Promise.resolve(true))
    requestDocsClose.mockImplementation(() => Promise.resolve(false))
    const id = manager.openDocsTab()
    await manager.closeTab(id)
    expect(requestDocsClose).toHaveBeenCalledTimes(1)
    expect(manager.list().map((t) => t.id)).toEqual(['home', id])
  })

  it('activates a dirty background tab before showing its close guard', async () => {
    sheetsPendingEditCount.mockImplementation(() => 1)
    requestSheetsClose.mockImplementation(() => Promise.resolve(false))
    const sheetsId = manager.openSheetsTab()
    manager.openSlidesTab()

    await manager.closeTab(sheetsId)
    expect(requestSheetsClose).toHaveBeenCalledTimes(1)
    // the guarded tab was brought into view for the prompt
    expect(manager.list().find((t) => t.id === sheetsId)?.active).toBe(true)
  })

  it('closes a dirty sheets tab when the guard resolves true', async () => {
    sheetsPendingEditCount.mockImplementation(() => 1)
    requestSheetsClose.mockImplementation(() => Promise.resolve(true))
    const id = manager.openSheetsTab()
    await manager.closeTab(id)
    expect(manager.list()).toHaveLength(1)
  })

  it('does not stack close guards while one prompt is pending', async () => {
    sheetsPendingEditCount.mockImplementation(() => 1)
    let resolveGuard!: (ok: boolean) => void
    requestSheetsClose.mockImplementation(
      () =>
        new Promise<boolean>((resolve) => {
          resolveGuard = resolve
        }),
    )
    const id = manager.openSheetsTab()
    const first = manager.closeTab(id)
    const second = manager.closeTab(id)
    resolveGuard(true)
    await Promise.all([first, second])
    expect(requestSheetsClose).toHaveBeenCalledTimes(1)
    expect(manager.list()).toHaveLength(1)
  })
})

describe('file path bookkeeping', () => {
  it('updates the tab title when a module opens a file in an existing tab', () => {
    manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    manager.setTabFileFor(view.webContents.id, '/tmp/final.docx')
    expect(manager.list()[1].title).toBe('final.docx')
    expect(manager.findDocsTabByPath('/tmp/final.docx')).toBe('t1')
  })

  it('ignores setTabFileFor for unknown webContents', () => {
    onChanged.mockClear()
    manager.setTabFileFor(999, '/tmp/x.docx')
    expect(onChanged).not.toHaveBeenCalled()
  })

  it('renames matching tabs and reports the affected views', () => {
    manager.openDocsTab('/tmp/old.docx')
    const view = lastCreatedView(createDocsView)
    const affected = manager.renameTabFile('/tmp/old.docx', '/tmp/new.docx')
    expect(affected).toEqual([{ kind: 'docs', webContents: view.webContents }])
    expect(manager.list()[1].title).toBe('new.docx')
    expect(manager.findDocsTabByPath('/tmp/new.docx')).toBe('t1')
    expect(manager.findDocsTabByPath('/tmp/old.docx')).toBeUndefined()
  })

  it('returns no affected views when nothing matches a rename', () => {
    onChanged.mockClear()
    expect(manager.renameTabFile('/tmp/none.docx', '/tmp/new.docx')).toEqual([])
    expect(onChanged).not.toHaveBeenCalled()
  })
})

describe('dirty-tab queries (shell close guard)', () => {
  it('lists only tabs whose module reports unsaved changes', () => {
    const dirtySheetsId = manager.openSheetsTab()
    const dirtyView = lastCreatedView(createSheetsView)
    manager.openSheetsTab()
    sheetsPendingEditCount.mockImplementation((id: number) =>
      id === dirtyView.webContents.id ? 3 : 0,
    )

    const dirty = manager.dirtySheetsTabs()
    expect(dirty).toEqual([{ id: dirtySheetsId, webContents: dirtyView.webContents }])
  })

  it('lists dirty pdf and slides tabs', () => {
    const pdfId = manager.openPdfTab('/tmp/c.pdf')
    const slidesId = manager.openSlidesTab()
    expect(manager.dirtyPdfTabs()).toEqual([])
    expect(manager.dirtySlidesTabs()).toEqual([])
    pdfIsDirty.mockImplementation(() => true)
    slidesIsDirty.mockImplementation(() => true)
    expect(manager.dirtyPdfTabs().map((t) => t.id)).toEqual([pdfId])
    expect(manager.dirtySlidesTabs().map((t) => t.id)).toEqual([slidesId])
  })

  it('lists every live docs tab for the async dirtiness sweep', () => {
    manager.openDocsTab()
    manager.openSheetsTab()
    manager.openDocsTab('/tmp/a.docx')
    expect(manager.docsTabs().map((t) => t.id)).toEqual(['t1', 't3'])
  })
})

describe('overlay visibility management', () => {
  it('hides active WebContentsView and returns snapshot when overlay is visible, and restores view when overlay closes', async () => {
    manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    expect(view.setVisible).toHaveBeenLastCalledWith(true)

    // Modal/overlay opened: active view hidden and snapshot captured
    const snapshot = await manager.setOverlayVisible(true)
    expect(view.setVisible).toHaveBeenLastCalledWith(false)
    expect(snapshot).toContain('data:image/jpeg;base64,')

    // Modal/overlay closed: active view restored and returns null
    const closed = await manager.setOverlayVisible(false)
    expect(view.setVisible).toHaveBeenLastCalledWith(true)
    expect(closed).toBeNull()
  })
})

describe('active context and tab actions', () => {
  it('returns home context when no document is active', async () => {
    const ctx = await manager.getActiveContext()
    expect(ctx.id).toBe('home')
    expect(ctx.kind).toBe('home')
    expect(ctx.title).toBe('Home')
  })

  it('returns active editor context with selection text', async () => {
    manager.openDocsTab('/tmp/Report.docx')
    const view = lastCreatedView(createDocsView)
    view.webContents.executeJavaScript = vi.fn().mockResolvedValue('Sample selected text')

    const ctx = await manager.getActiveContext()
    expect(ctx.kind).toBe('docs')
    expect(ctx.title).toBe('Report.docx')
    expect(ctx.selectionText).toBe('Sample selected text')
  })

  it('dispatches insert-text action into active editor view', async () => {
    manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    view.webContents.insertText = vi.fn()

    const res = await manager.dispatchTabAction({
      type: 'insert-text',
      text: 'Generated AI Summary',
    })
    expect(res.ok).toBe(true)
    expect(view.webContents.insertText).toHaveBeenCalledWith('Generated AI Summary')
  })

  it('dispatches create-doc action to open a new tab', async () => {
    const res = await manager.dispatchTabAction({ type: 'create-doc', kind: 'slides' })
    expect(res.ok).toBe(true)
    expect(manager.list().some((t) => t.kind === 'slides')).toBe(true)
  })

  it('opens orchestrated diagram tabs in visual mode and delegates the prompt to them', async () => {
    vi.useFakeTimers()
    try {
      const res = await manager.dispatchTabAction({
        type: 'create-doc',
        kind: 'diagram',
        prompt: 'Draw the checkout flow on the visual canvas',
      })
      expect(res.ok).toBe(true)
      expect(createDiagramView).toHaveBeenCalledWith(undefined, 'visual')

      await vi.advanceTimersByTimeAsync(300)
      const view = lastCreatedView(createDiagramView)
      expect(view.webContents.send).toHaveBeenCalledWith(
        'menu:command',
        'ai-prompt',
        'Draw the checkout flow on the visual canvas',
      )
    } finally {
      vi.useRealTimers()
    }
  })

  it('backfills missing outline stages without duplicating the Stages heading', async () => {
    vi.useFakeTimers()
    try {
      const prompt =
        'Task: Buatkan saya "flowchart pendaftaran" di editor.\nLanguage: Indonesian\nDetails:\n- Kebutuhan: flowchart pendaftaran.\nStages:\n1. Masukkan data\n2. Proses data'
      const res = await manager.dispatchTabAction({
        type: 'create-doc',
        kind: 'docs',
        prompt,
        outline: [
          { label: 'Masukkan data', depth: 0 },
          { label: 'Proses data', depth: 0 },
          { label: 'Tampilkan hasil', depth: 0 },
        ],
      })
      expect(res.ok).toBe(true)
      await vi.advanceTimersByTimeAsync(300)
      const view = lastCreatedView(createDocsView)
      // docs is a handshake-gated kind: the prompt lands once the mounted editor
      // announces itself (content assertions below are about the backfilled text)
      view.webContents.listeners.get('ipc-message')!(null, 'editor:ready')
      await vi.advanceTimersByTimeAsync(0)
      expect(view.webContents.send).toHaveBeenCalledWith(
        'menu:command',
        'ai-prompt',
        `${prompt}\n3. Tampilkan hasil`,
      )
    } finally {
      vi.useRealTimers()
    }
  })

  it('anchors backfilled stages under Stages when the prompt has no section', async () => {
    vi.useFakeTimers()
    try {
      const res = await manager.dispatchTabAction({
        type: 'create-doc',
        kind: 'docs',
        prompt: 'Buatkan laporan evaluasi.',
        outline: [{ label: 'Kumpulkan data', depth: 0 }],
      })
      expect(res.ok).toBe(true)
      await vi.advanceTimersByTimeAsync(300)
      const view = lastCreatedView(createDocsView)
      // docs is a handshake-gated kind: the prompt lands once the mounted editor
      // announces itself (content assertions below are about the backfilled text)
      view.webContents.listeners.get('ipc-message')!(null, 'editor:ready')
      await vi.advanceTimersByTimeAsync(0)
      expect(view.webContents.send).toHaveBeenCalledWith(
        'menu:command',
        'ai-prompt',
        'Buatkan laporan evaluasi.\nStages:\n1. Kumpulkan data',
      )
    } finally {
      vi.useRealTimers()
    }
  })
})

describe('editor readiness handshake', () => {
  it('holds queued insertText until the renderer reports its editor mounted', async () => {
    vi.useFakeTimers()
    try {
      await manager.dispatchTabAction({ type: 'create-doc', kind: 'docs', content: 'Hello world' })
      const view = lastCreatedView(createDocsView)
      view.webContents.insertText = vi.fn()
      await vi.advanceTimersByTimeAsync(300)
      expect(view.webContents.insertText).not.toHaveBeenCalled()
      view.webContents.listeners.get('ipc-message')!(null, 'editor:ready')
      await vi.advanceTimersByTimeAsync(0)
      expect(view.webContents.insertText).toHaveBeenCalledWith('Hello world')
    } finally {
      vi.useRealTimers()
    }
  })

  it('writes anyway after the bounded handshake timeout', async () => {
    vi.useFakeTimers()
    try {
      await manager.dispatchTabAction({ type: 'create-doc', kind: 'docs', content: 'Hello world' })
      const view = lastCreatedView(createDocsView)
      view.webContents.insertText = vi.fn()
      await vi.advanceTimersByTimeAsync(3000)
      expect(view.webContents.insertText).toHaveBeenCalledWith('Hello world')
    } finally {
      vi.useRealTimers()
    }
  })
})

describe('tab path identity', () => {
  it('finds a tab by any path that normalizes to the same file', () => {
    const id = manager.openDocsTab('/tmp/report.docx')
    expect(manager.findDocsTabByPath('/tmp/./report.docx')).toBe(id)
    if (process.platform === 'win32') {
      // path comparison is case-insensitive only on win32 (normalizePathKey)
      expect(manager.findDocsTabByPath('/tmp/REPORT.DOCX')).toBe(id)
    }
  })
})

describe('isTabWebContents', () => {
  it('tracks tab webContents id when opened and untracks when closed', async () => {
    expect(manager.isTabWebContents(999)).toBe(false)
    const id = manager.openDocsTab('/tmp/report.docx')
    const view = lastCreatedView(createDocsView)
    expect(manager.isTabWebContents(view.webContents.id)).toBe(true)
    await manager.closeTab(id)
    expect(manager.isTabWebContents(view.webContents.id)).toBe(false)
  })
})
