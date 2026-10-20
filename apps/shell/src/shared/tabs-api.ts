/** the tab strip's rendered height; applied inline in TabBar.tsx (overrides tabbar.css) */
export const TAB_STRIP_HEIGHT = 40

export type TabKind = 'home' | 'docs' | 'sheets' | 'slides' | 'pdf' | 'markdown' | 'diagram'

/** one open tab in the top tab strip; Home is always id 'home' and not closable */
export interface TabSummary {
  id: string
  kind: TabKind
  title: string
  closable: boolean
  active: boolean
  diagramMode?: 'visual' | 'mermaid'
}

export interface ActiveTabContext {
  id: string
  kind: TabKind
  title: string
  filePath?: string
  selectionText?: string
}

export type TabAction =
  | { type: 'insert-text'; text: string }
  | {
      type: 'create-doc'
      kind: 'docs' | 'sheets' | 'slides' | 'markdown' | 'diagram'
      content?: string
      prompt?: string
      summary?: string
      /** orchestrated diagram handoff: visual canvas default, mermaid only on explicit code request */
      diagramMode?: 'visual' | 'mermaid'
      /** ordered outline stages parsed from the omnibar answer; depth preserves
       * hierarchy (0 = top, higher = child). Also embedded in prompt. */
      outline?: Array<{ label: string; depth: number }>
    }

export interface TabsApi {
  list(): Promise<TabSummary[]>
  activate(id: string): Promise<void>
  close(id: string): Promise<void>
  /**
   * Temporarily hide active editor WebContentsView so shell modal overlays
   * (e.g. Command Center, Account Dialog) can render without being obscured.
   * Returns a lightweight snapshot data URL if an active editor was captured, or null.
   */
  setOverlayVisible(visible: boolean): Promise<string | null>
  /**
   * Get context snapshot of currently active tab (title, kind, selection)
   * to enrich AI assistance and command execution.
   */
  getActiveContext(): Promise<ActiveTabContext>
  /**
   * Dispatch an approved action into the active tab (e.g. insert text, create doc).
   */
  dispatchAction(
    action: TabAction,
  ): Promise<{ ok: boolean; error?: string; tabId?: string; tabTitle?: string }>
  /**
   * pop up the native "all tabs" list menu at (x, y) in window CSS coordinates.
   * Native because the content area below the strip is a WebContentsView that
   * would cover any DOM dropdown rendered by the shell.
   */
  showMenu(x: number, y: number): Promise<void>
  /**
   * pop up the native "+" new-file menu (new doc/sheet/slides, open local
   * file) at (x, y) in window CSS coordinates. Native for the same reason
   * as showMenu.
   */
  showNewMenu(x: number, y: number): Promise<void>
  /** pop up the native context menu for a specific tab at (x, y) */
  showContextMenu(id: string, x: number, y: number): Promise<void>
  /** move a tab to a new index in the strip; Home stays pinned at index 0 */
  reorder(id: string, toIndex: number): Promise<void>
  /** subscribe to tab list changes (open/close/activate/title updates); returns unsubscribe */
  onChanged(handler: (tabs: TabSummary[]) => void): () => void
}

export const TABS_CHANNELS = {
  list: 'tabs:list',
  activate: 'tabs:activate',
  close: 'tabs:close',
  setOverlayVisible: 'tabs:set-overlay-visible',
  getActiveContext: 'tabs:get-active-context',
  dispatchAction: 'tabs:dispatch-action',
  showMenu: 'tabs:show-menu',
  showNewMenu: 'tabs:show-new-menu',
  showContextMenu: 'tabs:show-context-menu',
  reorder: 'tabs:reorder',
  changed: 'tabs:changed',
} as const
