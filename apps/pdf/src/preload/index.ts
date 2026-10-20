import { contextBridge, ipcRenderer } from 'electron'
import type { Lang } from '@araoffice/i18n'
import type { AiSettings, AiStreamChunk } from '@araoffice/ai-provider'
import { AI_CHANNELS, PDF_CHANNELS } from '../shared/ipc'
import type { PdfApi } from '../shared/ipc'

const api: PdfApi = {
  consumePending: () => ipcRenderer.invoke(PDF_CHANNELS.consumePending),
  readFile: (path) => ipcRenderer.invoke(PDF_CHANNELS.readFile, path),
  save: (request) => ipcRenderer.invoke(PDF_CHANNELS.save, request),
  extractPages: (request) => ipcRenderer.invoke(PDF_CHANNELS.extractPages, request),
  insertPdf: (request) => ipcRenderer.invoke(PDF_CHANNELS.insertPdf, request),
  exportImages: (request) => ipcRenderer.invoke(PDF_CHANNELS.exportImages, request),
  exportText: (request) => ipcRenderer.invoke(PDF_CHANNELS.exportText, request),
  exportEmbeddedImages: (request) => ipcRenderer.invoke(PDF_CHANNELS.exportEmbeddedImages, request),
  setOutline: (request) => ipcRenderer.invoke(PDF_CHANNELS.setOutline, request),
  splitPdf: (request) => ipcRenderer.invoke(PDF_CHANNELS.splitPdf, request),
  combinePdfs: (request) => ipcRenderer.invoke(PDF_CHANNELS.combinePdfs, request),
  sanitizePdf: (request) => ipcRenderer.invoke(PDF_CHANNELS.sanitizePdf, request),
  sanitizePreview: (request) => ipcRenderer.invoke(PDF_CHANNELS.sanitizePreview, request),
  deleteAnnotations: (request) => ipcRenderer.invoke(PDF_CHANNELS.deleteAnnotations, request),
  a11yCheck: (request) => ipcRenderer.invoke(PDF_CHANNELS.a11yCheck, request),
  a11yFix: (request) => ipcRenderer.invoke(PDF_CHANNELS.a11yFix, request),
  setDirty: (dirty) => ipcRenderer.send(PDF_CHANNELS.dirtyChanged, dirty),
  onCloseSaveRequest: (handler) => {
    const listener = () => handler()
    ipcRenderer.on(PDF_CHANNELS.closeSaveRequest, listener)
    return () => ipcRenderer.removeListener(PDF_CHANNELS.closeSaveRequest, listener)
  },
  sendCloseSaveResult: (ok) => ipcRenderer.send(PDF_CHANNELS.closeSaveResult, ok),
  getLanguage: () => ipcRenderer.invoke(PDF_CHANNELS.getLanguage),
  onLanguageChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, lang: Lang) => handler(lang)
    ipcRenderer.on(PDF_CHANNELS.languageChanged, listener)
    return () => ipcRenderer.removeListener(PDF_CHANNELS.languageChanged, listener)
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

contextBridge.exposeInMainWorld('pdfApi', api)
