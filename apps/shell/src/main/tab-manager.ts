import { basename, resolve } from 'node:path'
import type { BrowserWindow, Rectangle, WebContents, WebContentsView } from 'electron'
import { powerMonitor } from 'electron'

import {
  createDocsView,
  docsQueryDirty,
  markDocsNewBlank,
  requestDocsClose,
  setActiveDocsResolver,
  teardownDocsRenderer,
} from '../../../docs/src/main/docs-main'
import {
  createMarkdownView,
  markdownIsDirty,
  requestMarkdownClose,
} from '../../../markdown/src/main/markdown-main'
import {
  createDiagramView,
  diagramIsDirty,
  requestDiagramClose,
} from '../../../diagram/src/main/diagram-main'
import { createPdfView, pdfIsDirty, requestPdfClose } from '../../../pdf/src/main/pdf-main'
import {
  createSheetsView,
  queueSheetsLaunch,
  requestSheetsClose,
  setActiveSheetsWebContents,
  sheetsPendingEditCount,
} from '../../../sheets/src/main/sheets-main'
import {
  createSlidesView,
  requestSlidesClose,
  setActiveSlidesWebContents,
  slidesIsDirty,
} from '../../../slides/src/main/slides-main'
import type { DiagramMode } from '../shared/home-api'
import type { ActiveTabContext, TabAction, TabKind, TabSummary } from '../shared/tabs-api'
import { TAB_STRIP_HEIGHT } from '../shared/tabs-api'

/**
 * Comparison-only path identity for tab dedupe: `path.resolve` + lowercase on win32.
 * `TabRecord.filePath` / `TabSummary` keep the raw display string — canonicalizing
 * storage would lowercase recents/titles on win32 and fork identity on case-sensitive volumes.
 * Main-process only: `shared/` is bundled into the sandboxed renderer, where `node:` imports
 * are externalized to empty stubs — Node-only helpers must not live there.
 */
export function normalizePathKey(p: string): string {
  const resolved = resolve(p)
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved
}

interface TabRecord {
  id: string
  kind: TabKind
  /** null for the Home tab — it's rendered by the shell window's own webContents */
  view: WebContentsView | null
  title: string
  filePath?: string
  diagramMode?: DiagramMode
  /** resolves when the renderer announces its mounted text editor ('editor:ready') */
  editorReady?: Promise<void>
  disposables?: Array<() => void>
}

const HOME_ID = 'home'
/** how long create-doc waits for the 'editor:ready' handshake before writing anyway */
const EDITOR_READY_TIMEOUT_MS = 2500
/** renderers that announce 'editor:ready' once their real editor is mounted. Other
 *  kinds never signal, so gating them would only add EDITOR_READY_TIMEOUT_MS of dead
 *  latency to every create-doc handover — they resolve immediately instead. */
const EDITOR_READY_KINDS: Partial<Record<TabKind, true>> = {
  docs: true,
  markdown: true,
  sheets: true,
}

/**
 * Owns every open tab (Home + docs + sheets) inside the shell's single
 * BrowserWindow. Docs/sheets tabs are WebContentsView children of that
 * window; only the active one is visible at a time. Home has no view of its
 * own — hiding every other tab reveals the shell window's own content.
 */
export class TabManager {
  private readonly tabs: TabRecord[] = [{ id: HOME_ID, kind: 'home', view: null, title: 'Home' }]
  private activeId: string = HOME_ID
  private nextId = 1
  /** tab whose page entered HTML fullscreen (e.g. slides slideshow) — its view covers the tab strip */
  private htmlFullScreenId: string | null = null
  /** tabs mid unsaved-changes prompt, so a second close click doesn't stack dialogs */
  private readonly closingIds = new Set<string>()
  private keepaliveTimer: NodeJS.Timeout | null = null
  private overlayVisible = false
  private readonly tabWebContentsIds = new Set<number>()

  isTabWebContents(webContentsId: number): boolean {
    if (this.tabWebContentsIds.has(webContentsId)) return true
    return this.tabs.some((t) => t.view?.webContents?.id === webContentsId)
  }

  private registerTabView(view: WebContentsView): void {
    const id = view.webContents?.id
    if (typeof id === 'number') {
      this.tabWebContentsIds.add(id)
      view.webContents?.once?.('destroyed', () => {
        this.tabWebContentsIds.delete(id)
      })
    }
  }

  constructor(
    private readonly shellWindow: BrowserWindow,
    private readonly onChanged: () => void,
    private readonly applyMenuFor: (kind: TabKind) => void,
    /** localized placeholder title for a tab that has no file yet */
    private readonly untitledTitleFor?: (kind: TabKind, subKind?: string) => string,
  ) {
    // Layout once synchronously for macOS/Windows (bounds are already correct),
    // then once more on the next tick. On Linux/X11, `resize` fires before the
    // window manager applies the new size, so getContentBounds() is still the
    // pre-maximize size inside the handler and a follow-up layout is required.
    shellWindow.on('resize', () => {
      this.layout()
      setImmediate(() => this.layout())
    })
    shellWindow.on('focus', () => this.layout(true))
    shellWindow.on('show', () => this.layout(true))
    shellWindow.on('restore', () => this.layout(true))

    const onPowerResume = () => this.layout(true)
    powerMonitor?.on?.('resume', onPowerResume)
    powerMonitor?.on?.('unlock-screen', onPowerResume)

    shellWindow.on('closed', () => {
      powerMonitor?.removeListener?.('resume', onPowerResume)
      powerMonitor?.removeListener?.('unlock-screen', onPowerResume)
      if (this.keepaliveTimer) {
        clearInterval(this.keepaliveTimer)
        this.keepaliveTimer = null
      }
    })

    this.startKeepalive()
  }

  private untitled(kind: TabKind, fallback: string, subKind?: string): string {
    return this.untitledTitleFor?.(kind, subKind) ?? fallback
  }

  private contentBounds(): Rectangle {
    const { width, height } = this.shellWindow.getContentBounds()
    if (this.htmlFullScreenId !== null && this.htmlFullScreenId === this.activeId) {
      return { x: 0, y: 0, width, height }
    }
    return { x: 0, y: TAB_STRIP_HEIGHT, width, height: Math.max(0, height - TAB_STRIP_HEIGHT) }
  }

  /**
   * When a tab's page enters HTML fullscreen (the slides slideshow calls requestFullscreen),
   * grow its view over the tab strip so nothing of the shell chrome shows;
   * restore the normal bounds on leave.
   */
  private trackHtmlFullScreen(id: string, view: WebContentsView, record: TabRecord): void {
    const onEnter = () => {
      this.htmlFullScreenId = id
      this.layout()
    }
    const onLeave = () => {
      if (this.htmlFullScreenId === id) this.htmlFullScreenId = null
      this.layout()
    }
    view.webContents.on('enter-html-full-screen', onEnter)
    view.webContents.on('leave-html-full-screen', onLeave)

    const onRenderProcessGone = (
      _event: Electron.Event,
      details: Electron.RenderProcessGoneDetails,
    ) => {
      console.warn(
        `[tab-manager] tab ${id} (${record.kind}) renderer gone: ${details.reason} (exitCode: ${details.exitCode})`,
      )
      // A crashed page is no longer in HTML fullscreen, but the matching
      // leave-html-full-screen never fires — without this the stale flag pins the
      // view over the tab strip for the tab's remaining life.
      if (this.htmlFullScreenId === id) this.htmlFullScreenId = null
      if (details.reason !== 'clean-exit' && !this.shellWindow.isDestroyed()) {
        try {
          if (!view.webContents.isDestroyed()) {
            view.webContents.reload()
          }
        } catch (err) {
          console.error(`[tab-manager] failed to reload crashed tab ${id}:`, err)
        }
      }
    }
    view.webContents.on('render-process-gone', onRenderProcessGone)

    const wcId = view.webContents?.id
    if (typeof wcId === 'number') {
      this.tabWebContentsIds.add(wcId)
    }

    record.disposables = [
      () => {
        if (typeof wcId === 'number') {
          this.tabWebContentsIds.delete(wcId)
        }
        try {
          view.webContents.removeListener('enter-html-full-screen', onEnter)
          view.webContents.removeListener('leave-html-full-screen', onLeave)
          view.webContents.removeListener('render-process-gone', onRenderProcessGone)
        } catch {
          /* ignore listener removal if webContents already destroyed */
        }
      },
    ]

    // editor:ready handshake: resolves when the renderer announces its mounted
    // editor; bounded so a silent renderer never wedges the queued handover payload.
    // Kinds that never announce resolve immediately (see EDITOR_READY_KINDS).
    record.editorReady = EDITOR_READY_KINDS[record.kind]
      ? new Promise<void>((resolve) => {
          const timer = setTimeout(resolve, EDITOR_READY_TIMEOUT_MS)
          timer.unref?.()
          const onMessage = (_event: Electron.IpcMainEvent, channel: string) => {
            if (channel !== 'editor:ready') return
            clearTimeout(timer)
            resolve()
            view.webContents.removeListener('ipc-message', onMessage)
          }
          view.webContents.on('ipc-message', onMessage)
          record.disposables?.push(() => {
            clearTimeout(timer)
            try {
              view.webContents.removeListener('ipc-message', onMessage)
            } catch {
              /* ignore listener removal if webContents already destroyed */
            }
          })
        })
      : Promise.resolve()
  }

  private startKeepalive(): void {
    // 2-minute keepalive resets Chromium's ~5min FrameEvictor timer before frames can be discarded
    this.keepaliveTimer = setInterval(() => {
      if (this.shellWindow.isDestroyed()) return
      const active = this.tabs.find((t) => t.id === this.activeId)
      if (active?.view && !active.view.webContents.isDestroyed?.()) {
        try {
          const bounds = this.contentBounds()
          active.view.setBounds({ ...bounds, width: Math.max(0, bounds.width - 1) })
          active.view.setBounds(bounds)
        } catch {
          /* ignore */
        }
      }
    }, 120_000)
    this.keepaliveTimer.unref?.()
  }

  /** re-fit every tab view after a window resize; background tabs keep their bounds
   *  too, so a resize-then-switch never flashes a stale-sized view */
  layout(force = false): void {
    // Deferred resize layouts can land after the shell window was closed.
    if (this.shellWindow.isDestroyed()) return
    const bounds = this.contentBounds()
    for (const t of this.tabs) {
      if (!t.view || t.view.webContents.isDestroyed?.()) continue
      try {
        t.view.setBounds(bounds)
      } catch {
        /* ignore webContents errors during layout */
      }
    }
    const active = this.tabs.find((t) => t.id === this.activeId)
    if (active?.view && !active.view.webContents.isDestroyed?.()) {
      try {
        if (force) {
          this.attachView(active.view)
          active.view.setVisible(false)
        }
        active.view.setVisible(!this.overlayVisible)
        if (force) {
          active.view.webContents.invalidate?.()
          // Nudge bounds by 1px to force DelegatedFrameHost to reset its evicted frame buffer
          active.view.setBounds({ ...bounds, width: Math.max(0, bounds.width - 1) })
          active.view.setBounds(bounds)
        }
      } catch {
        /* ignore webContents errors during layout */
      }
    }
  }

  /** raise `view` to the top of the shell's view stack without re-adding an
   *  already-attached view (addChildView on a current child hides z-order bugs) */
  private attachView(view: WebContentsView): void {
    const stack = this.shellWindow.contentView
    try {
      stack.removeChildView(view)
    } catch {
      /* not currently a child */
    }
    stack.addChildView(view)
  }

  list(): TabSummary[] {
    return this.tabs.map((t) => ({
      id: t.id,
      kind: t.kind,
      title: t.title,
      closable: t.id !== HOME_ID,
      active: t.id === this.activeId,
      diagramMode: t.diagramMode,
    }))
  }

  openHomeTab(): void {
    this.activateTab(HOME_ID)
  }

  /** Shared body of every open<Kind>Tab: create view, register, prep, attach, activate. */
  private openTab(
    kind: TabKind,
    createView: () => WebContentsView,
    prep: (id: string, view: WebContentsView) => Pick<TabRecord, 'title'> & Partial<TabRecord>,
  ): string {
    const view = createView()
    this.registerTabView(view)
    const id = `t${this.nextId++}`
    const record: TabRecord = { id, kind, view, ...prep(id, view) }
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view, record)
    this.tabs.push(record)
    this.activateTab(id)
    return id
  }

  openDocsTab(openPath?: string, options?: { newBlank?: boolean }): string {
    return this.openTab(
      'docs',
      () => createDocsView(openPath),
      (_id, view) => {
        if (options?.newBlank) markDocsNewBlank(view.webContents.id)
        return {
          title: openPath ? basename(openPath) : this.untitled('docs', 'ARAOffice Docs'),
          filePath: openPath,
        }
      },
    )
  }

  openSheetsTab(
    openPath?: string,
    options?: { newBlank?: boolean; queuedOpenPath?: string; suggestSaveAs?: string },
  ): string {
    return this.openTab(
      'sheets',
      () => createSheetsView({ includeAiHandlers: false }),
      (_id, view) => {
        const queuedOpenPath = options?.queuedOpenPath ?? openPath
        if (queuedOpenPath || options?.newBlank) {
          queueSheetsLaunch(view.webContents, {
            ...(queuedOpenPath ? { path: queuedOpenPath } : {}),
            ...(options?.suggestSaveAs ? { suggestSaveAs: options.suggestSaveAs } : {}),
            ...(options?.newBlank ? { newBlank: true } : {}),
          })
        }
        return {
          title: openPath ? basename(openPath) : this.untitled('sheets', 'ARA Sheets'),
          filePath: openPath,
        }
      },
    )
  }

  openSlidesTab(openPath?: string): string {
    return this.openTab(
      'slides',
      () => createSlidesView(openPath),
      () => ({
        title: openPath ? basename(openPath) : this.untitled('slides', 'ARA Slides'),
        filePath: openPath,
      }),
    )
  }

  openPdfTab(openPath: string): string {
    return this.openTab(
      'pdf',
      () => createPdfView(openPath),
      () => ({
        title: basename(openPath),
        filePath: openPath,
      }),
    )
  }

  openMarkdownTab(openPath?: string): string {
    return this.openTab(
      'markdown',
      () => createMarkdownView(openPath),
      () => ({
        title: openPath ? basename(openPath) : this.untitled('markdown', 'ARA Notes'),
        filePath: openPath,
      }),
    )
  }

  openDiagramTab(openPath?: string, initialMode?: DiagramMode): string {
    return this.openTab(
      'diagram',
      () => createDiagramView(openPath, initialMode),
      () => {
        const isCode =
          initialMode === 'mermaid' || (openPath ? /\.(mmd|mermaid)$/i.test(openPath) : false)
        return {
          title: openPath
            ? basename(openPath)
            : this.untitled(
                'diagram',
                isCode ? 'Untitled Diagram Code' : 'Untitled Diagram',
                isCode ? 'mermaid' : 'visual',
              ),
          filePath: openPath,
          diagramMode: isCode ? 'mermaid' : 'visual',
        }
      },
    )
  }

  activateTab(id: string): void {
    const target = this.tabs.find((t) => t.id === id)
    if (!target) return
    for (const t of this.tabs) {
      const isCurrent = t.id === id && !this.overlayVisible
      t.view?.setVisible(isCurrent)
    }
    if (target.view) {
      if (this.activeId !== HOME_ID && this.activeId !== id) {
        this.attachView(target.view)
      }
      const bounds = this.contentBounds()
      target.view.setBounds({ ...bounds, width: Math.max(0, bounds.width - 1) })
      target.view.setBounds(bounds)
    }
    this.activeId = id
    setActiveDocsResolver(target.kind === 'docs' ? () => target.view!.webContents : () => null)
    if (target.kind === 'sheets' && target.view) setActiveSheetsWebContents(target.view.webContents)
    if (target.kind === 'slides' && target.view) setActiveSlidesWebContents(target.view.webContents)
    this.applyMenuFor(target.kind)
    this.onChanged()
  }

  async setOverlayVisible(visible: boolean): Promise<string | null> {
    this.overlayVisible = visible
    const active = this.tabs.find((t) => t.id === this.activeId)
    if (!active?.view) {
      return null
    }

    if (visible) {
      // A native WebContentsView always paints over the window's DOM, so the view
      // MUST be hidden for the modal to be usable. capturePage() snapshots the frame
      // at call time, so start the capture first but never await it before hiding —
      // gating the hide on the capture left the modal obscured behind the view for
      // up to the capture timeout. The snapshot is only a cosmetic backdrop.
      let snapshot: string | null = null
      let capture: Promise<Electron.NativeImage | null> = Promise.resolve(null)
      if (!active.view.webContents.isDestroyed?.()) {
        capture = active.view.webContents.capturePage().catch(() => null)
      }
      active.view.setVisible(false)
      try {
        const image = await Promise.race([
          capture,
          new Promise<null>((resolve) => setTimeout(() => resolve(null), 600)),
        ])
        if (image && !image.isEmpty()) {
          const size = image.getSize()
          const targetWidth = Math.min(size.width, 1280)
          const targetHeight = Math.round(size.height * (targetWidth / Math.max(1, size.width)))
          const resized =
            targetWidth < size.width
              ? image.resize({ width: targetWidth, height: targetHeight, quality: 'good' })
              : image
          const jpegBuffer = resized.toJPEG(75)
          snapshot = `data:image/jpeg;base64,${jpegBuffer.toString('base64')}`
        }
      } catch (err) {
        console.warn('Failed to capture overlay snapshot:', err)
      }
      return snapshot
    } else {
      active.view.setVisible(true)
      return null
    }
  }

  async getActiveContext(): Promise<ActiveTabContext> {
    const active = this.tabs.find((t) => t.id === this.activeId)
    if (!active || active.id === HOME_ID || !active.view) {
      return {
        id: HOME_ID,
        kind: 'home',
        title: 'Home',
      }
    }

    let selectionText = ''
    try {
      if (!active.view.webContents.isDestroyed()) {
        const sel = await Promise.race([
          active.view.webContents.executeJavaScript(
            '(function(){ try { return (window.getSelection ? window.getSelection().toString() : "").slice(0, 2000); } catch(e){ return ""; } })()',
            true,
          ),
          new Promise<string>((resolve) => setTimeout(() => resolve(''), 300)),
        ])
        if (typeof sel === 'string') selectionText = sel.trim()
      }
    } catch {
      /* ignore selection read failure */
    }

    return {
      id: active.id,
      kind: active.kind,
      title: active.title,
      filePath: active.filePath,
      selectionText: selectionText || undefined,
    }
  }

  async dispatchTabAction(
    action: TabAction,
  ): Promise<{ ok: boolean; error?: string; tabId?: string; tabTitle?: string }> {
    const active = this.tabs.find((t) => t.id === this.activeId)
    if (action.type === 'insert-text') {
      if (!active || !active.view || active.view.webContents.isDestroyed()) {
        return { ok: false, error: 'No active editor document' }
      }
      // webContents.insertText only lands in DOM text fields (Docs/Markdown). Canvas
      // editors (Sheets/Slides/Diagram) and the read-only PDF skill cannot receive it.
      if (active.kind === 'sheets' || active.kind === 'slides' || active.kind === 'diagram') {
        return {
          ok: false,
          error: 'Direct insert is not supported in this editor — use "Open in editor" instead',
        }
      }
      if (active.kind === 'pdf' || active.kind === 'home') {
        return {
          ok: false,
          error:
            active.kind === 'pdf'
              ? 'PDF is read-only — copy the text instead'
              : 'No active editor document',
        }
      }
      try {
        active.view.webContents.insertText(action.text)
        return { ok: true, tabId: active.id, tabTitle: active.title }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    }

    if (action.type === 'create-doc') {
      try {
        let tabId = ''
        if (action.kind === 'docs') {
          tabId = this.openDocsTab(undefined, { newBlank: true })
        } else if (action.kind === 'sheets') {
          tabId = this.openSheetsTab(undefined, { newBlank: true })
        } else if (action.kind === 'slides') {
          tabId = this.openSlidesTab(undefined)
        } else if (action.kind === 'markdown') {
          tabId = this.openMarkdownTab(undefined)
        } else if (action.kind === 'diagram') {
          tabId = this.openDiagramTab(undefined, action.diagramMode ?? 'visual')
        }

        const createdTab = this.tabs.find((t) => t.id === tabId)
        if (createdTab?.view && (action.content || action.prompt)) {
          const view = createdTab.view
          // Diagram tabs have no text field: content always flows through the
          // delegated ai-prompt, never insertText.
          const textToWrite = action.kind === 'diagram' ? '' : action.content || ''
          // Backfill outline stages missing from the prompt so the target app
          // never receives a generic label without its steps. Envelope prompts
          // already carry a Stages: section — continue its numbering instead of
          // opening a second heading; plain prompts get the Stages: anchor.
          let machinePrompt = action.prompt ?? ''
          const outlineLines = Array.isArray(action.outline)
            ? action.outline
                .map((s) => (typeof s === 'string' ? s : s?.label))
                .map((s) => String(s ?? '').trim())
                .filter(Boolean)
            : []
          const missingStages = outlineLines.filter(
            (s) => machinePrompt && !machinePrompt.toLowerCase().includes(s.toLowerCase()),
          )
          if (missingStages.length > 0) {
            const hasStagesSection = /^(Stages|Tahapan):/m.test(machinePrompt)
            if (hasStagesSection) {
              const existingCount = machinePrompt.match(/^\d+\./gm)?.length ?? 0
              machinePrompt = `${machinePrompt}\n${missingStages.map((s, i) => `${existingCount + i + 1}. ${s}`).join('\n')}`
            } else {
              machinePrompt = `${machinePrompt}\nStages:\n${missingStages.map((s, i) => `${i + 1}. ${s}`).join('\n')}`
            }
          }
          const deliveredPrompt = machinePrompt
          const maxAttempts = 20

          const setupTabContent = (attempt = 0) => {
            if (view.webContents.isDestroyed()) return
            // Wait for the mounted-editor handshake before handing over: an
            // unmounted editor silently drops insertText, and the ai-prompt it
            // drives runs tool calls against an empty document. EDITOR_READY_TIMEOUT_MS
            // bounds the wait so a renderer that never reports ready still gets its
            // prompt (today's fire-and-forget behaviour) instead of none.
            const deliver = () => {
              if (view.webContents.isDestroyed()) return
              void (createdTab.editorReady ?? Promise.resolve()).then(() => {
                if (view.webContents.isDestroyed()) return
                try {
                  view.webContents.focus()
                  if (textToWrite) view.webContents.insertText(textToWrite)
                  if (deliveredPrompt)
                    view.webContents.send('menu:command', 'ai-prompt', deliveredPrompt)
                } catch (err) {
                  console.warn('[tab-manager] Failed to initialize content/prompt in new tab:', err)
                }
              })
            }
            if (
              typeof view.webContents.executeJavaScript !== 'function' ||
              attempt >= maxAttempts
            ) {
              deliver()
              return
            }
            void view.webContents
              .executeJavaScript(
                '(function(){ try { return document.readyState === "complete" && !!document.body; } catch(e){ return false; } })()',
                true,
              )
              .then((ready) => {
                if (view.webContents.isDestroyed()) return
                if (ready) {
                  setTimeout(deliver, 50)
                } else {
                  setTimeout(() => setupTabContent(attempt + 1), 250)
                }
              })
              .catch(() => {
                if (!view.webContents.isDestroyed()) {
                  setTimeout(() => setupTabContent(attempt + 1), 250)
                }
              })
          }

          if (view.webContents.isLoading()) {
            view.webContents.once('did-finish-load', () => {
              setTimeout(() => setupTabContent(), 450)
            })
          } else {
            setTimeout(() => setupTabContent(), 250)
          }
        }

        return {
          ok: true,
          tabId,
          tabTitle: createdTab?.title || 'Untitled',
        }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    }

    return { ok: false, error: 'Unsupported action type' }
  }

  /** move a tab to a new index in the strip; Home is pinned at index 0 */
  reorderTab(id: string, toIndex: number): void {
    if (id === HOME_ID) return
    const fromIndex = this.tabs.findIndex((t) => t.id === id)
    if (fromIndex < 0) return
    const clamped = Math.min(Math.max(Math.trunc(toIndex), 1), this.tabs.length - 1)
    if (clamped === fromIndex) return
    const [moved] = this.tabs.splice(fromIndex, 1)
    this.tabs.splice(clamped, 0, moved)
    this.onChanged()
  }

  /** a module opened a file inside an existing tab (⌘O / queued path) — sync title + dedupe path */
  setTabFileFor(webContentsId: number, filePath: string): void {
    const tab = this.tabs.find((t) => t.view?.webContents.id === webContentsId)
    if (!tab) return
    tab.filePath = filePath
    tab.title = basename(filePath)
    if (tab.kind === 'diagram') {
      tab.diagramMode = /\.(mmd|mermaid)$/i.test(filePath) ? 'mermaid' : 'visual'
    }
    this.onChanged()
  }

  /** a file was renamed on disk (rename from the Home list) — sync any open tab's title/path;
   *  returns the affected views so callers can notify the embedded editors */
  renameTabFile(
    oldPath: string,
    newPath: string,
  ): Array<{ kind: TabKind; webContents: WebContents }> {
    const affected: Array<{ kind: TabKind; webContents: WebContents }> = []
    for (const tab of this.tabs) {
      if (
        tab.filePath === undefined ||
        normalizePathKey(tab.filePath) !== normalizePathKey(oldPath)
      )
        continue
      tab.filePath = newPath
      tab.title = basename(newPath)
      if (tab.kind === 'diagram') {
        tab.diagramMode = /\.(mmd|mermaid)$/i.test(newPath) ? 'mermaid' : 'visual'
      }
      if (tab.view) affected.push({ kind: tab.kind, webContents: tab.view.webContents })
    }
    if (affected.length > 0) this.onChanged()
    return affected
  }

  /** sheets tabs whose renderer reports unsaved journal edits (shell-close guard) */
  dirtySheetsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter(
        (t) => t.kind === 'sheets' && t.view && sheetsPendingEditCount(t.view.webContents.id) > 0,
      )
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** pdf tabs whose renderer reports unsaved markups/form edits (shell-close guard) */
  dirtyPdfTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'pdf' && t.view && pdfIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** markdown tabs whose renderer reports unsaved edits (shell-close guard) */
  dirtyMarkdownTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'markdown' && t.view && markdownIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }
  /** diagram tabs whose renderer reports unsaved edits (shell-close guard) */
  dirtyDiagramTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'diagram' && t.view && diagramIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** slides tabs whose main-process session has unsaved edits (shell-close guard) */
  dirtySlidesTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'slides' && t.view && slidesIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** all live docs tabs — dirtiness lives renderer-side, caller queries async (shell-close guard) */
  docsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'docs' && t.view)
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** closes whichever tab is currently active; no-op for Home (Cmd+W target) */
  closeActiveTab(): void {
    void this.closeTab(this.activeId)
  }

  async closeTab(id: string): Promise<void> {
    if (id === HOME_ID) return
    const tab = this.tabs.find((t) => t.id === id)
    if (!tab || this.closingIds.has(id)) return
    let closeGuard =
      tab.view &&
      (tab.kind === 'sheets' && sheetsPendingEditCount(tab.view.webContents.id) > 0
        ? requestSheetsClose
        : tab.kind === 'pdf' && pdfIsDirty(tab.view.webContents.id)
          ? requestPdfClose
          : tab.kind === 'markdown' && markdownIsDirty(tab.view.webContents.id)
            ? requestMarkdownClose
            : tab.kind === 'diagram' && diagramIsDirty(tab.view.webContents.id)
              ? requestDiagramClose
              : tab.kind === 'slides' && slidesIsDirty(tab.view.webContents.id)
                ? requestSlidesClose
                : null)
    // docs dirty state lives in the renderer and needs an async query; skip the guard when clean (avoids a flash activation)
    if (!closeGuard && tab.kind === 'docs' && tab.view) {
      this.closingIds.add(id)
      try {
        if (await docsQueryDirty(tab.view.webContents)) closeGuard = requestDocsClose
      } finally {
        this.closingIds.delete(id)
      }
    }
    if (closeGuard && tab.view) {
      // Bring the tab into view so the save prompt has visible context.
      if (this.activeId !== id) this.activateTab(id)
      this.closingIds.add(id)
      try {
        if (!(await closeGuard(tab.view.webContents, this.shellWindow))) return
      } finally {
        this.closingIds.delete(id)
      }
    }
    const idx = this.tabs.findIndex((t) => t.id === id)
    if (idx < 0) return
    if (this.htmlFullScreenId === id) this.htmlFullScreenId = null
    const [removed] = this.tabs.splice(idx, 1)
    if (this.activeId === id) {
      const fallback = this.tabs[idx - 1] ?? this.tabs[0]
      this.activateTab(fallback.id)
    } else {
      this.onChanged()
    }
    if (removed.disposables) {
      for (const dispose of removed.disposables) {
        try {
          dispose()
        } catch {
          /* ignore disposal errors */
        }
      }
    }
    if (removed.view) {
      if (typeof removed.view.webContents?.id === 'number') {
        this.tabWebContentsIds.delete(removed.view.webContents.id)
      }
      removed.view.setVisible(false)
      this.shellWindow.contentView.removeChildView(removed.view)
      if (removed.kind === 'docs') {
        // Closing this WebContents wedges Electron's UI thread, so the docs tab is
        // torn down instead: teardownDocsRenderer reclaims every per-wcId registry
        // (writable paths, disk states, attachments, close waiters) and crashing the
        // detached renderer releases that process. The residue is the WebContentsView
        // handle itself, which Electron keeps alive until app exit because close() is
        // unsafe — small and O(1) per closed docs tab, and unfixable from here without
        // re-introducing the UI-thread wedge.
        // ponytail: per-tab WebContents handle until app exit; release it if Electron
        // ever exposes a safe destroy that does not wedge the UI thread.
        teardownDocsRenderer(removed.view.webContents)
        try {
          removed.view.webContents.setBackgroundThrottling(true)
          void removed.view.webContents.loadURL('about:blank')
          removed.view.webContents.forcefullyCrashRenderer()
        } catch {
          /* ignore teardown errors on detached view */
        }
      } else {
        removed.view.webContents.close()
      }
    }
  }
  getTab(id: string): TabRecord | undefined {
    return this.tabs.find((t) => t.id === id)
  }

  async closeOtherTabs(id: string): Promise<void> {
    const targets = this.tabs.filter((t) => t.id !== HOME_ID && t.id !== id).map((t) => t.id)
    for (const targetId of targets) {
      await this.closeTab(targetId)
    }
  }

  async closeTabsToRight(id: string): Promise<void> {
    const index = this.tabs.findIndex((t) => t.id === id)
    if (index < 0) return
    const targets = this.tabs
      .slice(index + 1)
      .filter((t) => t.id !== HOME_ID)
      .map((t) => t.id)
    for (const targetId of targets) {
      await this.closeTab(targetId)
    }
  }

  private findTabByPath(kind: TabKind, path: string): string | undefined {
    const key = normalizePathKey(path)
    return this.tabs.find(
      (t) => t.kind === kind && t.filePath !== undefined && normalizePathKey(t.filePath) === key,
    )?.id
  }

  findDocsTabByPath(path: string): string | undefined {
    return this.findTabByPath('docs', path)
  }

  findSheetsTabByPath(path: string): string | undefined {
    return this.findTabByPath('sheets', path)
  }

  findSlidesTabByPath(path: string): string | undefined {
    return this.findTabByPath('slides', path)
  }

  findPdfTabByPath(path: string): string | undefined {
    return this.findTabByPath('pdf', path)
  }

  findMarkdownTabByPath(path: string): string | undefined {
    return this.findTabByPath('markdown', path)
  }
  findDiagramTabByPath(path: string): string | undefined {
    return this.findTabByPath('diagram', path)
  }
}
