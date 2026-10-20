import type { Lang } from '@araoffice/i18n'
import type { AiSettings, AiStreamChunk, AiStreamRequest } from '@araoffice/ai-provider'

export const AI_CHANNELS = {
  getSettings: 'ai:get-settings',
  stream: 'ai:stream',
  streamChunk: 'ai:stream-chunk',
  streamCancel: 'ai:stream-cancel',
} as const
export const DIAGRAM_CHANNELS = {
  consumePending: 'diagram:consume-pending',
  readFile: 'diagram:read-file',
  save: 'diagram:save',
  saveRequest: 'diagram:save-request',
  saveRequestAck: 'diagram:save-request-ack',
  dirtyChanged: 'diagram:dirty-changed',
  closeSaveRequest: 'diagram:close-save-request',
  closeSaveResult: 'diagram:close-save-result',
  fileRenamed: 'diagram:file-renamed',
  exportRequest: 'diagram:export-request',
  exportSvg: 'diagram:export-svg',
  exportImage: 'diagram:export-image',
  getLanguage: 'app:get-language',
  languageChanged: 'app:language-changed',
} as const

export type SaveMode = 'save' | 'saveAs'
export type DiagramMode = 'visual' | 'mermaid'

export interface PendingDiagram {
  path: string | null
  mode: DiagramMode
}

export interface SaveDiagramRequest {
  content: string
  mode: SaveMode
  format?: 'drawio' | 'mermaid'
  suggestedName?: string
}

export type SaveDiagramResult =
  { ok: true; path: string } | { ok: true; canceled: true } | { ok: false; error: string }

export type ExportFormat = 'svg' | 'png' | 'jpeg'

export interface ExportFileRequest {
  data: string // SVG string, or base64 PNG/JPEG data URL matching the export format
  format: ExportFormat
  suggestedName?: string
}

export type ExportResult =
  { ok: true; path: string } | { ok: true; canceled: true } | { ok: false; error: string }

/** API exposed by preload to the renderer (window.diagramApi) */
export interface DiagramApi {
  consumePending: () => Promise<PendingDiagram>
  readFile: (path: string) => Promise<string>
  save: (request: SaveDiagramRequest) => Promise<SaveDiagramResult>
  setDirty: (dirty: boolean) => void
  onSaveRequest: (handler: (mode: SaveMode) => void) => () => void
  onCloseSaveRequest: (handler: () => void) => () => void
  sendCloseSaveResult: (ok: boolean) => void
  sendSaveRequestAck: (ok: boolean) => void
  onFileRenamed: (handler: (newPath: string) => void) => () => void
  onExportRequest: (handler: (format: ExportFormat) => void) => () => void
  /** Shell menu commands (ARA Omni delegation); payload carries the synthesized prompt */
  onMenuCommand: (handler: (command: 'ai-prompt', payload?: string) => void) => () => void
  exportSvg: (request: ExportFileRequest) => Promise<ExportResult>
  /** PNG/JPEG raster export; data must be a data URL matching `format` */
  exportImage: (request: ExportFileRequest) => Promise<ExportResult>
  getLanguage: () => Promise<Lang>
  onLanguageChanged: (handler: (lang: Lang) => void) => () => void
  getAiSettings: () => Promise<AiSettings>
  /** open the shell's AI provider settings (no accounts: configuring a provider is the only setup) */
  configureAi: () => Promise<void>
  onAiSettingsChanged?: (handler: (settings: AiSettings) => void) => () => void
  aiStream: (request: AiStreamRequest) => Promise<void>
  aiStreamCancel: (requestId: string) => Promise<void>
  onAiStream: (handler: (chunk: AiStreamChunk) => void) => () => void
}
