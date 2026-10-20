import { contextBridge, ipcRenderer } from 'electron'
import type { Lang } from '@araoffice/i18n'
import type { AiSettings, AiStreamChunk } from '@araoffice/ai-provider'
import type { ProjectApi } from '@araoffice/project-store'
import { AI_CHANNELS, DIAGRAM_CHANNELS } from '../shared/ipc'
import type { DiagramApi, ExportFormat, SaveMode } from '../shared/ipc'

let activeDiagramMenuHandler: ((command: 'ai-prompt', payload?: string) => void) | null = null
const pendingDiagramMenuCommands: Array<{
  command: 'ai-prompt'
  payload?: string
}> = []

const handleDiagramMenuIpc = (_e: Electron.IpcRendererEvent, command: string, payload?: string) => {
  if (activeDiagramMenuHandler) {
    activeDiagramMenuHandler(command as 'ai-prompt', payload)
  } else {
    pendingDiagramMenuCommands.push({ command: command as 'ai-prompt', payload })
  }
}
ipcRenderer.on('menu:command', handleDiagramMenuIpc)

const api: DiagramApi = {
  consumePending: () => ipcRenderer.invoke(DIAGRAM_CHANNELS.consumePending),
  readFile: (path) => ipcRenderer.invoke(DIAGRAM_CHANNELS.readFile, path),
  save: (request) => ipcRenderer.invoke(DIAGRAM_CHANNELS.save, request),
  setDirty: (dirty) => ipcRenderer.send(DIAGRAM_CHANNELS.dirtyChanged, dirty),
  onSaveRequest: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, mode: SaveMode) => handler(mode)
    ipcRenderer.on(DIAGRAM_CHANNELS.saveRequest, listener)
    return () => ipcRenderer.removeListener(DIAGRAM_CHANNELS.saveRequest, listener)
  },
  onCloseSaveRequest: (handler) => {
    const listener = () => handler()
    ipcRenderer.on(DIAGRAM_CHANNELS.closeSaveRequest, listener)
    return () => ipcRenderer.removeListener(DIAGRAM_CHANNELS.closeSaveRequest, listener)
  },
  sendCloseSaveResult: (ok) => ipcRenderer.send(DIAGRAM_CHANNELS.closeSaveResult, ok),
  sendSaveRequestAck: (ok) => ipcRenderer.send(DIAGRAM_CHANNELS.saveRequestAck, ok),
  onFileRenamed: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, newPath: string) => handler(newPath)
    ipcRenderer.on(DIAGRAM_CHANNELS.fileRenamed, listener)
    return () => ipcRenderer.removeListener(DIAGRAM_CHANNELS.fileRenamed, listener)
  },
  onExportRequest: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, format: ExportFormat) => handler(format)
    ipcRenderer.on(DIAGRAM_CHANNELS.exportRequest, listener)
    return () => ipcRenderer.removeListener(DIAGRAM_CHANNELS.exportRequest, listener)
  },
  onMenuCommand: (handler) => {
    activeDiagramMenuHandler = handler
    while (pendingDiagramMenuCommands.length > 0) {
      const item = pendingDiagramMenuCommands.shift()
      if (item) handler(item.command, item.payload)
    }
    return () => {
      if (activeDiagramMenuHandler === handler) {
        activeDiagramMenuHandler = null
      }
    }
  },
  exportSvg: (request) => ipcRenderer.invoke(DIAGRAM_CHANNELS.exportSvg, request),
  exportImage: (request) => ipcRenderer.invoke(DIAGRAM_CHANNELS.exportImage, request),
  getLanguage: () => ipcRenderer.invoke(DIAGRAM_CHANNELS.getLanguage),
  onLanguageChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, lang: Lang) => handler(lang)
    ipcRenderer.on(DIAGRAM_CHANNELS.languageChanged, listener)
    return () => ipcRenderer.removeListener(DIAGRAM_CHANNELS.languageChanged, listener)
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
}

/** Chat persistence: the shared project:* handlers are registered once by the shell (docs-main registerProjectIpc) */
const projectApi: Pick<ProjectApi, 'resolveChat' | 'appendChat' | 'loadChat' | 'rebindChat'> = {
  resolveChat: (args) => ipcRenderer.invoke('project:resolveChat', args),
  appendChat: (args) => ipcRenderer.invoke('project:appendChat', args),
  loadChat: (args) => ipcRenderer.invoke('project:loadChat', args),
  rebindChat: (args) => ipcRenderer.invoke('project:rebindChat', args),
}

contextBridge.exposeInMainWorld('diagramApi', api)
contextBridge.exposeInMainWorld('projectApi', projectApi)
