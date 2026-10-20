import { contextBridge, ipcRenderer } from 'electron'
import type { IpcRendererEvent } from 'electron'
import type { AppUpdateState, ChangelogEntry } from '../shared/update-api'
import type {
  HomeApi,
  RecentEntry,
  RecentPage,
  RenameResult,
  ProjectHomeApi,
  ProjectSummaryEntry,
  UiLanguage,
} from '../shared/home-api'
import { HOME_CHANNELS, PROJECT_CHANNELS } from '../shared/home-api'
import type { ActiveTabContext, TabsApi, TabSummary } from '../shared/tabs-api'
import { TABS_CHANNELS } from '../shared/tabs-api'
import type { ShellAiApi } from '../shared/ai-api'
import type { AiStreamChunk } from '@araoffice/ai-provider'
import type { WindowApi } from '../shared/window-api'
import { WINDOW_CHANNELS } from '../shared/window-api'

const UI_LANGUAGES: readonly UiLanguage[] = [
  'zh',
  'en',
  'ja',
  'ko',
  'fr',
  'de',
  'es',
  'th',
  'id',
  'ru',
  'ar',
  'pt',
  'it',
  'pl',
  'nl',
  'ms',
  'he',
  'hi',
  'zh-TW',
]

function isUiLanguage(value: unknown): value is UiLanguage {
  return UI_LANGUAGES.includes(value as UiLanguage)
}

const EMPTY_PAGE: RecentPage = { entries: [], total: 0, totalAll: 0 }

function asRecentPage(result: unknown): RecentPage {
  if (result && typeof result === 'object' && Array.isArray((result as RecentPage).entries)) {
    return result as RecentPage
  }
  return EMPTY_PAGE
}

const homeApi: HomeApi = {
  async recents(query) {
    return asRecentPage(await ipcRenderer.invoke(HOME_CHANNELS.recents, query))
  },
  async starred(query) {
    return asRecentPage(await ipcRenderer.invoke(HOME_CHANNELS.starred, query))
  },
  async statPaths(paths) {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.statPaths, paths)
    return Array.isArray(result) ? (result as RecentEntry[]) : []
  },
  async toggleStar(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.toggleStar, path)
  },
  async openPath(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.openPath, path)
  },
  async browse() {
    await ipcRenderer.invoke(HOME_CHANNELS.browse)
  },
  async newDoc(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newDoc, opts)
  },
  async newSheet(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newSheet, opts)
  },
  async newSlide(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newSlide, opts)
  },
  async newMarkdown(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newMarkdown, opts)
  },
  async newDiagram(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newDiagram, opts)
  },
  async openPdf() {
    await ipcRenderer.invoke(HOME_CHANNELS.openPdf)
  },
  async removeRecent(paths) {
    await ipcRenderer.invoke(HOME_CHANNELS.removeRecent, paths)
  },
  async revealPath(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.revealPath, path)
  },
  async renameFile(path, newName) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.renameFile, path, newName)
    return (result ?? { ok: false, error: 'Rename failed' }) as RenameResult
  },
  async duplicateFile(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.duplicateFile, path)
  },
  async deleteFiles(paths) {
    await ipcRenderer.invoke(HOME_CHANNELS.deleteFiles, paths)
  },
  async openTrash() {
    await ipcRenderer.invoke(HOME_CHANNELS.openTrash)
  },
  async getLanguage() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getLanguage)
    return isUiLanguage(result) ? result : 'zh'
  },
  async setLanguage(lang) {
    if (!isUiLanguage(lang)) throw new Error('Invalid language.')
    await ipcRenderer.invoke(HOME_CHANNELS.setLanguage, lang)
  },
  async getUpdateChannel() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getUpdateChannel)
    return result === 'beta' ? 'beta' : 'stable'
  },
  async setUpdateChannel(channel) {
    // validated inline: a runtime import from ../shared/update-api would be
    // shared with the update.ts preload entry and get split into a chunk,
    // which sandboxed preload scripts cannot load (window.aiOffice would
    // silently disappear). Preload entries must stay single-file bundles.
    if (channel !== 'stable' && channel !== 'beta') throw new Error('Invalid update channel.')
    await ipcRenderer.invoke(HOME_CHANNELS.setUpdateChannel, channel)
  },
  async getAppVersion() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getAppVersion)
    return typeof result === 'string' ? result : ''
  },
  async onboardingSeen() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.onboardingSeen)
    return result === true
  },
  async setOnboardingSeen() {
    await ipcRenderer.invoke(HOME_CHANNELS.setOnboardingSeen)
  },
  async getTheme() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getTheme)
    return result === 'dark' || result === 'light' ? result : 'system'
  },
  async setTheme(theme) {
    if (theme !== 'light' && theme !== 'dark' && theme !== 'system')
      throw new Error('Invalid theme.')
    await ipcRenderer.invoke(HOME_CHANNELS.setTheme, theme)
  },
  async getUpdateState() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getUpdateState)
    return (
      result && typeof result === 'object' ? result : { status: 'idle', currentVersion: '' }
    ) as AppUpdateState
  },
  async checkForUpdates() {
    await ipcRenderer.invoke(HOME_CHANNELS.checkForUpdates)
  },
  async restartAndInstall() {
    await ipcRenderer.invoke(HOME_CHANNELS.restartAndInstall)
  },
  async getChangelog() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getChangelog)
    return Array.isArray(result) ? (result as ChangelogEntry[]) : []
  },
  async getJustUpdated() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getJustUpdated)
    return result === true
  },
  async clearJustUpdated() {
    await ipcRenderer.invoke(HOME_CHANNELS.clearJustUpdated)
  },
  async getSupportLinks() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getSupportLinks)
    const empty = {
      issuesUrl: null,
      websiteUrl: null,
      githubUrl: null,
      instagramUrl: null,
      vendorName: null,
    }
    if (result && typeof result === 'object') {
      const value = result as Record<string, unknown>
      const asUrl = (key: string): string | null =>
        typeof value[key] === 'string' ? (value[key] as string) : null
      const asName = (key: string): string | null =>
        typeof value[key] === 'string' && (value[key] as string).trim() !== ''
          ? (value[key] as string)
          : null
      return {
        issuesUrl: asUrl('issuesUrl'),
        websiteUrl: asUrl('websiteUrl'),
        githubUrl: asUrl('githubUrl'),
        instagramUrl: asUrl('instagramUrl'),
        vendorName: asName('vendorName'),
      }
    }
    return empty
  },
  onUpdateStateChanged(handler) {
    const listener = (_event: IpcRendererEvent, state: AppUpdateState) => {
      handler(state)
    }
    ipcRenderer.on(HOME_CHANNELS.updateStateChanged, listener)
    return () => {
      ipcRenderer.removeListener(HOME_CHANNELS.updateStateChanged, listener)
    }
  },
  onConfigureAi(handler) {
    const listener = () => handler()
    ipcRenderer.on(HOME_CHANNELS.configureAi, listener)
    return () => {
      ipcRenderer.removeListener(HOME_CHANNELS.configureAi, listener)
    }
  },
}
contextBridge.exposeInMainWorld('aiOffice', homeApi)

const projectApi: ProjectHomeApi = {
  async listProjects() {
    const result: unknown = await ipcRenderer.invoke(PROJECT_CHANNELS.list)
    return Array.isArray(result) ? (result as ProjectSummaryEntry[]) : []
  },
  async listFiles(projectId) {
    const result: unknown = await ipcRenderer.invoke(PROJECT_CHANNELS.files, { projectId })
    return Array.isArray(result)
      ? result.filter((path): path is string => typeof path === 'string')
      : []
  },
  async createProject(name) {
    const result: unknown = await ipcRenderer.invoke(PROJECT_CHANNELS.create, { name })
    return result as ProjectSummaryEntry
  },
  async renameProject(id, name) {
    await ipcRenderer.invoke(PROJECT_CHANNELS.rename, { id, name })
  },
  async deleteProject(id) {
    await ipcRenderer.invoke(PROJECT_CHANNELS.delete, { id })
  },
  async moveFile(filePath, projectId) {
    await ipcRenderer.invoke(PROJECT_CHANNELS.moveFile, { filePath, projectId })
  },
}

contextBridge.exposeInMainWorld('aiOfficeProject', projectApi)

const tabsApi: TabsApi = {
  async list() {
    const result: unknown = await ipcRenderer.invoke(TABS_CHANNELS.list)
    return Array.isArray(result) ? (result as TabSummary[]) : []
  },
  async activate(id) {
    await ipcRenderer.invoke(TABS_CHANNELS.activate, id)
  },
  async close(id) {
    await ipcRenderer.invoke(TABS_CHANNELS.close, id)
  },
  async setOverlayVisible(visible) {
    const result: unknown = await ipcRenderer.invoke(TABS_CHANNELS.setOverlayVisible, visible)
    return typeof result === 'string' ? result : null
  },
  async getActiveContext() {
    const result: unknown = await ipcRenderer.invoke(TABS_CHANNELS.getActiveContext)
    return (result as ActiveTabContext) ?? { id: 'home', kind: 'home', title: 'Home' }
  },
  async dispatchAction(action) {
    const result: unknown = await ipcRenderer.invoke(TABS_CHANNELS.dispatchAction, action)
    return (
      (result as { ok: boolean; error?: string; tabId?: string; tabTitle?: string }) ?? {
        ok: false,
      }
    )
  },
  async showMenu(x, y) {
    await ipcRenderer.invoke(TABS_CHANNELS.showMenu, x, y)
  },
  async showNewMenu(x, y) {
    await ipcRenderer.invoke(TABS_CHANNELS.showNewMenu, x, y)
  },
  async showContextMenu(id, x, y) {
    await ipcRenderer.invoke(TABS_CHANNELS.showContextMenu, id, x, y)
  },
  async reorder(id, toIndex) {
    await ipcRenderer.invoke(TABS_CHANNELS.reorder, id, toIndex)
  },
  onChanged(handler) {
    const listener = (_event: IpcRendererEvent, tabs: TabSummary[]) => handler(tabs)
    ipcRenderer.on(TABS_CHANNELS.changed, listener)
    return () => ipcRenderer.removeListener(TABS_CHANNELS.changed, listener)
  },
}

contextBridge.exposeInMainWorld('aiOfficeTabs', tabsApi)

const windowApi: WindowApi = {
  isWindows: navigator.userAgent.includes('Windows'),
  async minimize() {
    await ipcRenderer.invoke(WINDOW_CHANNELS.minimize)
  },
  async toggleMaximize() {
    await ipcRenderer.invoke(WINDOW_CHANNELS.toggleMaximize)
  },
  async isMaximized() {
    return (await ipcRenderer.invoke(WINDOW_CHANNELS.isMaximized)) === true
  },
  async close() {
    await ipcRenderer.invoke(WINDOW_CHANNELS.close)
  },
  async openExternal(url) {
    await ipcRenderer.invoke(WINDOW_CHANNELS.openExternal, url)
  },
  async openThirdPartyNotices() {
    await ipcRenderer.invoke(WINDOW_CHANNELS.openThirdPartyNotices)
  },
  onMaximizedChanged(handler) {
    const listener = (_event: IpcRendererEvent, maximized: boolean) => handler(maximized)
    ipcRenderer.on(WINDOW_CHANNELS.stateChanged, listener)
    return () => ipcRenderer.removeListener(WINDOW_CHANNELS.stateChanged, listener)
  },
}

contextBridge.exposeInMainWorld('aiOfficeWindow', windowApi)

const aiApi: ShellAiApi = {
  async getSettings() {
    return ipcRenderer.invoke('ai:get-settings')
  },
  async setSettings(settings) {
    await ipcRenderer.invoke('ai:set-settings', settings)
  },
  async fetchCustomModels(override) {
    return ipcRenderer.invoke('ai:fetch-custom-models', override)
  },
  stream(request, callbacks) {
    const requestId = crypto.randomUUID()
    const listener = (_event: IpcRendererEvent, chunk: AiStreamChunk) => {
      if (chunk.requestId === requestId) {
        callbacks.onChunk(chunk)
        if (chunk.type === 'done' || chunk.type === 'error') {
          ipcRenderer.removeListener('ai:stream-chunk', listener)
        }
      }
    }
    ipcRenderer.on('ai:stream-chunk', listener)
    void ipcRenderer.invoke('ai:stream', { ...request, requestId }).catch((err: unknown) => {
      callbacks.onChunk({
        requestId,
        type: 'error',
        error: String(err),
      })
      ipcRenderer.removeListener('ai:stream-chunk', listener)
    })
    return {
      requestId,
      cancel: () => {
        ipcRenderer.removeListener('ai:stream-chunk', listener)
        void ipcRenderer.invoke('ai:cancel', requestId)
      },
    }
  },
  async cancel(requestId) {
    await ipcRenderer.invoke('ai:cancel', requestId)
  },
}

contextBridge.exposeInMainWorld('aiOfficeAi', aiApi)
