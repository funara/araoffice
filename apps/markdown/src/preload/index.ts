import { contextBridge, ipcRenderer } from 'electron'
import type { Lang } from '@araoffice/i18n'
import type { AiSettings, AiStreamChunk } from '@araoffice/ai-provider'
import type { ProjectApi } from '@araoffice/project-store'
import { AI_CHANNELS, MARKDOWN_CHANNELS } from '../shared/ipc'
import type { MarkdownApi } from '../shared/ipc'

let activeMarkdownMenuHandler: ((command: 'ai-prompt', payload?: string) => void) | null = null
const pendingMarkdownMenuCommands: Array<{ command: 'ai-prompt'; payload?: string }> = []

const handleMarkdownMenuIpc = (
  _e: Electron.IpcRendererEvent,
  command: string,
  payload?: string,
) => {
  if (activeMarkdownMenuHandler) {
    activeMarkdownMenuHandler(command as 'ai-prompt', payload)
  } else {
    pendingMarkdownMenuCommands.push({ command: command as 'ai-prompt', payload })
  }
}
ipcRenderer.on('menu:command', handleMarkdownMenuIpc)

const api: MarkdownApi = {
  consumePending: () => ipcRenderer.invoke(MARKDOWN_CHANNELS.consumePending),
  assetTicket: () => ipcRenderer.invoke(MARKDOWN_CHANNELS.assetTicket),
  readFile: (path) => ipcRenderer.invoke(MARKDOWN_CHANNELS.readFile, path),
  save: (request) => ipcRenderer.invoke(MARKDOWN_CHANNELS.save, request),
  setDirty: (dirty) => ipcRenderer.send(MARKDOWN_CHANNELS.dirtyChanged, dirty),
  onCloseSaveRequest: (handler) => {
    const listener = () => handler()
    ipcRenderer.on(MARKDOWN_CHANNELS.closeSaveRequest, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.closeSaveRequest, listener)
  },
  sendCloseSaveResult: (ok) => ipcRenderer.send(MARKDOWN_CHANNELS.closeSaveResult, ok),
  onFileRenamed: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, newPath: string) => handler(newPath)
    ipcRenderer.on(MARKDOWN_CHANNELS.fileRenamed, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.fileRenamed, listener)
  },
  onMenuCommand: (handler) => {
    activeMarkdownMenuHandler = handler
    while (pendingMarkdownMenuCommands.length > 0) {
      const item = pendingMarkdownMenuCommands.shift()
      if (item) handler(item.command, item.payload)
    }
    return () => {
      if (activeMarkdownMenuHandler === handler) {
        activeMarkdownMenuHandler = null
      }
    }
  },
  pickImage: () => ipcRenderer.invoke(MARKDOWN_CHANNELS.pickImage),
  saveImage: (data) => ipcRenderer.invoke(MARKDOWN_CHANNELS.saveImage, data),
  readImage: (src) => ipcRenderer.invoke(MARKDOWN_CHANNELS.readImage, src),
  exportDocx: (request) => ipcRenderer.invoke(MARKDOWN_CHANNELS.exportDocx, request),
  exportPdf: (request) => ipcRenderer.invoke(MARKDOWN_CHANNELS.exportPdf, request),
  getLanguage: () => ipcRenderer.invoke(MARKDOWN_CHANNELS.getLanguage),
  onLanguageChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, lang: Lang) => handler(lang)
    ipcRenderer.on(MARKDOWN_CHANNELS.languageChanged, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.languageChanged, listener)
  },
  getAiSettings: () => ipcRenderer.invoke(AI_CHANNELS.getSettings),
  configureAi: () => ipcRenderer.invoke('ai:configure'),
  onAiSettingsChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, s: AiSettings) => handler(s)
    ipcRenderer.on('ai:settings-changed', listener)
    return () => ipcRenderer.removeListener('ai:settings-changed', listener)
  },
  aiStream: (request) => ipcRenderer.invoke(AI_CHANNELS.stream, request),
  aiStreamCancel: (requestId) => ipcRenderer.invoke(AI_CHANNELS.streamCancel, requestId),
  onAiStream: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, chunk: AiStreamChunk) => handler(chunk)
    ipcRenderer.on(AI_CHANNELS.streamChunk, listener)
    return () => ipcRenderer.removeListener(AI_CHANNELS.streamChunk, listener)
  },
  webSearch: (query, maxResults) => ipcRenderer.invoke(AI_CHANNELS.webSearch, query, maxResults),
  webFetch: (url) => ipcRenderer.invoke(AI_CHANNELS.webFetch, url),
  reportEditorReady: () => ipcRenderer.send('editor:ready'),
}

/** Chat persistence: the shared project:* handlers are registered once by the shell (docs-main registerProjectIpc) */
const projectApi: Pick<ProjectApi, 'resolveChat' | 'appendChat' | 'loadChat' | 'rebindChat'> = {
  resolveChat: (args) => ipcRenderer.invoke('project:resolveChat', args),
  appendChat: (args) => ipcRenderer.invoke('project:appendChat', args),
  loadChat: (args) => ipcRenderer.invoke('project:loadChat', args),
  rebindChat: (args) => ipcRenderer.invoke('project:rebindChat', args),
}

contextBridge.exposeInMainWorld('markdownApi', api)
contextBridge.exposeInMainWorld('projectApi', projectApi)
