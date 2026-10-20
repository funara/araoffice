import { existsSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'
import { BrowserWindow, WebContentsView, app, dialog, ipcMain, shell } from 'electron'
import type { WebContents } from 'electron'
import {
  atomicWriteFile,
  contextMenuLabels,
  installContextMenu,
  installNavigationGuard,
  safeExternalUrl,
  showSaveDialogWithMemory,
} from '@araoffice/electron-utils'
import { createI18n, getUiLang } from '@araoffice/i18n'
import { DIAGRAM_CHANNELS } from '../shared/ipc'
import type {
  ExportFileRequest,
  ExportFormat,
  ExportResult,
  SaveDiagramRequest,
  SaveDiagramResult,
  SaveMode,
  DiagramMode,
} from '../shared/ipc'

const tDlg = createI18n({
  zh: {
    dlgSaveTitle: '保存图表',
    filterDrawio: 'Draw.io 图表 (*.drawio)',
    filterMermaid: 'Mermaid 图表 (*.mmd)',
    filterSvg: 'SVG 矢量图 (*.svg)',
    filterPng: 'PNG 图片 (*.png)',
    filterJpeg: 'JPEG 图片 (*.jpg)',
    untitledFile: '未命名图表',
    closeUnsavedMsg: '此图表有未保存的更改。',
    closeUnsavedDetail: '关闭前是否保存？',
    btnSave: '保存',
    btnDontSave: '不保存',
    btnCancel: '取消',
  },
  en: {
    dlgSaveTitle: 'Save Diagram',
    filterDrawio: 'Draw.io Diagram (*.drawio)',
    filterMermaid: 'Mermaid Diagram (*.mmd)',
    filterSvg: 'SVG Image (*.svg)',
    filterPng: 'PNG Image (*.png)',
    filterJpeg: 'JPEG Image (*.jpg)',
    untitledFile: 'Untitled Diagram',
    closeUnsavedMsg: 'This diagram has unsaved changes.',
    closeUnsavedDetail: 'Do you want to save before closing?',
    btnSave: 'Save',
    btnDontSave: "Don't Save",
    btnCancel: 'Cancel',
  },
  ja: {
    dlgSaveTitle: '図面を保存',
    filterDrawio: 'Draw.io 図面 (*.drawio)',
    filterMermaid: 'Mermaid 図面 (*.mmd)',
    filterSvg: 'SVG 画像 (*.svg)',
    filterPng: 'PNG 画像 (*.png)',
    filterJpeg: 'JPEG 画像 (*.jpg)',
    untitledFile: '無題の図面',
    closeUnsavedMsg: 'この図面に未保存の変更があります。',
    closeUnsavedDetail: '閉じる前に保存しますか？',
    btnSave: '保存',
    btnDontSave: '保存しない',
    btnCancel: 'キャンセル',
  },
  ko: {
    dlgSaveTitle: '다이어그램 저장',
    filterDrawio: 'Draw.io 다이어그램 (*.drawio)',
    filterMermaid: 'Mermaid 다이어그램 (*.mmd)',
    filterSvg: 'SVG 이미지 (*.svg)',
    filterPng: 'PNG 이미지 (*.png)',
    filterJpeg: 'JPEG 이미지 (*.jpg)',
    untitledFile: '제목 없는 다이어그램',
    closeUnsavedMsg: '저장되지 않은 변경 사항이 있습니다.',
    closeUnsavedDetail: '닫기 전에 저장하시겠습니까?',
    btnSave: '저장',
    btnDontSave: '저장 안 함',
    btnCancel: '취소',
  },
  fr: {
    dlgSaveTitle: 'Enregistrer le diagramme',
    filterDrawio: 'Diagramme Draw.io (*.drawio)',
    filterMermaid: 'Diagramme Mermaid (*.mmd)',
    filterSvg: 'Image SVG (*.svg)',
    filterPng: 'Image PNG (*.png)',
    filterJpeg: 'Image JPEG (*.jpg)',
    untitledFile: 'Diagramme sans titre',
    closeUnsavedMsg: 'Ce diagramme contient des modifications non enregistrées.',
    closeUnsavedDetail: 'Voulez-vous les enregistrer avant de fermer ?',
    btnSave: 'Enregistrer',
    btnDontSave: 'Ne pas enregistrer',
    btnCancel: 'Annuler',
  },
  de: {
    dlgSaveTitle: 'Diagramm speichern',
    filterDrawio: 'Draw.io-Diagramm (*.drawio)',
    filterMermaid: 'Mermaid-Diagramm (*.mmd)',
    filterSvg: 'SVG-Grafik (*.svg)',
    filterPng: 'PNG-Bild (*.png)',
    filterJpeg: 'JPEG-Bild (*.jpg)',
    untitledFile: 'Unbenanntes Diagramm',
    closeUnsavedMsg: 'Dieses Diagramm enthält ungespeicherte Änderungen.',
    closeUnsavedDetail: 'Vor dem Schließen speichern?',
    btnSave: 'Speichern',
    btnDontSave: 'Nicht speichern',
    btnCancel: 'Abbrechen',
  },
  es: {
    dlgSaveTitle: 'Guardar diagrama',
    filterDrawio: 'Diagrama Draw.io (*.drawio)',
    filterMermaid: 'Diagrama Mermaid (*.mmd)',
    filterSvg: 'Imagen SVG (*.svg)',
    filterPng: 'Imagen PNG (*.png)',
    filterJpeg: 'Imagen JPEG (*.jpg)',
    untitledFile: 'Diagrama sin título',
    closeUnsavedMsg: 'Este diagrama tiene cambios sin guardar.',
    closeUnsavedDetail: '¿Quieres guardarlos antes de cerrar?',
    btnSave: 'Guardar',
    btnDontSave: 'No guardar',
    btnCancel: 'Cancelar',
  },
  th: {
    dlgSaveTitle: 'บันทึกไดอะแกรม',
    filterDrawio: 'ไดอะแกรม Draw.io (*.drawio)',
    filterMermaid: 'ไดอะแกรม Mermaid (*.mmd)',
    filterSvg: 'รูปภาพ SVG (*.svg)',
    filterPng: 'รูปภาพ PNG (*.png)',
    filterJpeg: 'รูปภาพ JPEG (*.jpg)',
    untitledFile: 'ไดอะแกรมไม่มีชื่อ',
    closeUnsavedMsg: 'ไดอะแกรมนี้มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก',
    closeUnsavedDetail: 'ต้องการบันทึกก่อนปิดหรือไม่?',
    btnSave: 'บันทึก',
    btnDontSave: 'ไม่บันทึก',
    btnCancel: 'ยกเลิก',
  },
  id: {
    dlgSaveTitle: 'Simpan Diagram',
    filterDrawio: 'Diagram Draw.io (*.drawio)',
    filterMermaid: 'Diagram Mermaid (*.mmd)',
    filterSvg: 'Gambar SVG (*.svg)',
    filterPng: 'Gambar PNG (*.png)',
    filterJpeg: 'Gambar JPEG (*.jpg)',
    untitledFile: 'Diagram Tanpa Judul',
    closeUnsavedMsg: 'Diagram ini memiliki perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
  },
  ru: {
    dlgSaveTitle: 'Сохранить диаграмму',
    filterDrawio: 'Диаграмма Draw.io (*.drawio)',
    filterMermaid: 'Диаграмма Mermaid (*.mmd)',
    filterSvg: 'Векторное изображение SVG (*.svg)',
    filterPng: 'Изображение PNG (*.png)',
    filterJpeg: 'Изображение JPEG (*.jpg)',
    untitledFile: 'Диаграмма без названия',
    closeUnsavedMsg: 'В этой диаграмме есть несохранённые изменения.',
    closeUnsavedDetail: 'Сохранить их перед закрытием?',
    btnSave: 'Сохранить',
    btnDontSave: 'Не сохранять',
    btnCancel: 'Отмена',
  },
  ar: {
    dlgSaveTitle: 'حفظ المخطط',
    filterDrawio: 'مخطط Draw.io (*.drawio)',
    filterMermaid: 'مخطط Mermaid (*.mmd)',
    filterSvg: 'صورة SVG (*.svg)',
    filterPng: 'صورة PNG (*.png)',
    filterJpeg: 'صورة JPEG (*.jpg)',
    untitledFile: 'مخطط بدون عنوان',
    closeUnsavedMsg: 'يحتوي هذا المخطط على تغييرات غير محفوظة.',
    closeUnsavedDetail: 'هل تريد حفظها قبل الإغلاق؟',
    btnSave: 'حفظ',
    btnDontSave: 'عدم الحفظ',
    btnCancel: 'إلغاء',
  },
  pt: {
    dlgSaveTitle: 'Salvar diagrama',
    filterDrawio: 'Diagrama Draw.io (*.drawio)',
    filterMermaid: 'Diagrama Mermaid (*.mmd)',
    filterSvg: 'Imagem SVG (*.svg)',
    filterPng: 'Imagem PNG (*.png)',
    filterJpeg: 'Imagem JPEG (*.jpg)',
    untitledFile: 'Diagrama sem título',
    closeUnsavedMsg: 'Este diagrama tem alterações não salvas.',
    closeUnsavedDetail: 'Deseja salvá-las antes de fechar?',
    btnSave: 'Salvar',
    btnDontSave: 'Não Salvar',
    btnCancel: 'Cancelar',
  },
  it: {
    dlgSaveTitle: 'Salva diagramma',
    filterDrawio: 'Diagramma Draw.io (*.drawio)',
    filterMermaid: 'Diagramma Mermaid (*.mmd)',
    filterSvg: 'Immagine SVG (*.svg)',
    filterPng: 'Immagine PNG (*.png)',
    filterJpeg: 'Immagine JPEG (*.jpg)',
    untitledFile: 'Diagramma senza titolo',
    closeUnsavedMsg: 'Questo diagramma contiene modifiche non salvate.',
    closeUnsavedDetail: 'Vuoi salvarle prima di chiudere?',
    btnSave: 'Salva',
    btnDontSave: 'Non salvare',
    btnCancel: 'Annulla',
  },
  pl: {
    dlgSaveTitle: 'Zapisz diagram',
    filterDrawio: 'Diagram Draw.io (*.drawio)',
    filterMermaid: 'Diagram Mermaid (*.mmd)',
    filterSvg: 'Obraz SVG (*.svg)',
    filterPng: 'Obraz PNG (*.png)',
    filterJpeg: 'Obraz JPEG (*.jpg)',
    untitledFile: 'Diagram bez tytułu',
    closeUnsavedMsg: 'Ten diagram ma niezapisane zmiany.',
    closeUnsavedDetail: 'Czy zapisać je przed zamknięciem?',
    btnSave: 'Zapisz',
    btnDontSave: 'Nie zapisuj',
    btnCancel: 'Anuluj',
  },
  nl: {
    dlgSaveTitle: 'Diagram opslaan',
    filterDrawio: 'Draw.io-diagram (*.drawio)',
    filterMermaid: 'Mermaid-diagram (*.mmd)',
    filterSvg: 'SVG-afbeelding (*.svg)',
    filterPng: 'PNG-afbeelding (*.png)',
    filterJpeg: 'JPEG-afbeelding (*.jpg)',
    untitledFile: 'Naamloos diagram',
    closeUnsavedMsg: 'Dit diagram bevat niet-opgeslagen wijzigingen.',
    closeUnsavedDetail: 'Wilt u ze opslaan voordat u sluit?',
    btnSave: 'Opslaan',
    btnDontSave: 'Niet opslaan',
    btnCancel: 'Annuleren',
  },
  ms: {
    dlgSaveTitle: 'Simpan Rajah',
    filterDrawio: 'Rajah Draw.io (*.drawio)',
    filterMermaid: 'Rajah Mermaid (*.mmd)',
    filterSvg: 'Imej SVG (*.svg)',
    filterPng: 'Imej PNG (*.png)',
    filterJpeg: 'Imej JPEG (*.jpg)',
    untitledFile: 'Rajah Tanpa Tajuk',
    closeUnsavedMsg: 'Rajah ini mempunyai perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
  },
  he: {
    dlgSaveTitle: 'שמירת תרשים',
    filterDrawio: 'תרשים Draw.io (*.drawio)',
    filterMermaid: 'תרשים Mermaid (*.mmd)',
    filterSvg: 'תמונת SVG (*.svg)',
    filterPng: 'תמונת PNG (*.png)',
    filterJpeg: 'תמונת JPEG (*.jpg)',
    untitledFile: 'תרשים ללא שם',
    closeUnsavedMsg: 'בתרשים זה יש שינויים שלא נשמרו.',
    closeUnsavedDetail: 'האם לשמור אותם לפני הסגירה?',
    btnSave: 'שמירה',
    btnDontSave: 'אל תשמור',
    btnCancel: 'ביטול',
  },
  hi: {
    dlgSaveTitle: 'आरेख सहेजें',
    filterDrawio: 'Draw.io आरेख (*.drawio)',
    filterMermaid: 'Mermaid आरेख (*.mmd)',
    filterSvg: 'SVG छवि (*.svg)',
    filterPng: 'PNG छवि (*.png)',
    filterJpeg: 'JPEG छवि (*.jpg)',
    untitledFile: 'शीर्षकहीन आरेख',
    closeUnsavedMsg: 'इस आरेख में सहेजे नहीं गए परिवर्तन हैं।',
    closeUnsavedDetail: 'क्या बंद करने से पहले उन्हें सहेजना चाहते हैं?',
    btnSave: 'सहेजें',
    btnDontSave: 'न सहेजें',
    btnCancel: 'रद्द करें',
  },
  'zh-TW': {
    dlgSaveTitle: '儲存圖表',
    filterDrawio: 'Draw.io 圖表 (*.drawio)',
    filterMermaid: 'Mermaid 圖表 (*.mmd)',
    filterSvg: 'SVG 向量圖 (*.svg)',
    filterPng: 'PNG 圖片 (*.png)',
    filterJpeg: 'JPEG 圖片 (*.jpg)',
    untitledFile: '未命名圖表',
    closeUnsavedMsg: '此圖表有未儲存的變更。',
    closeUnsavedDetail: '關閉前是否儲存？',
    btnSave: '儲存',
    btnDontSave: '不儲存',
    btnCancel: '取消',
  },
})

type DlgKey =
  | 'dlgSaveTitle'
  | 'filterDrawio'
  | 'filterMermaid'
  | 'filterSvg'
  | 'filterPng'
  | 'filterJpeg'
  | 'untitledFile'
  | 'closeUnsavedMsg'
  | 'closeUnsavedDetail'
  | 'btnSave'
  | 'btnDontSave'
  | 'btnCancel'

const tm = (key: DlgKey) => tDlg(getUiLang(), key)

interface RuntimePaths {
  preloadPath: string
  rendererUrl?: string
  rendererFile?: string
}

let runtime: RuntimePaths = { preloadPath: '' }

export function configureDiagramRuntime(paths: RuntimePaths): void {
  runtime = paths
}

const openPathByWc = new Map<number, string>()
const allowedByWc = new Map<number, Set<string>>()
const savePathByWc = new Map<number, string>()
const dirtyByWc = new Set<number>()
const diagramModeByWc = new Map<number, DiagramMode>()

function modeForPath(path?: string | null): DiagramMode {
  return path && /\.(mmd|mermaid)$/i.test(path) ? 'mermaid' : 'visual'
}

function isSaveDiagramRequest(value: unknown): value is SaveDiagramRequest {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false
  const request = value as Record<string, unknown>
  return (
    typeof request.content === 'string' &&
    (request.mode === 'save' || request.mode === 'saveAs') &&
    (request.format === undefined || request.format === 'drawio' || request.format === 'mermaid') &&
    (request.suggestedName === undefined || typeof request.suggestedName === 'string')
  )
}

const RASTER_DATA_PREFIX: Record<'png' | 'jpeg', string> = {
  png: 'data:image/png;base64,',
  jpeg: 'data:image/jpeg;base64,',
}

function isExportFileRequest(value: unknown, format: ExportFormat): value is ExportFileRequest {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false
  const request = value as Record<string, unknown>
  if (request.format !== format || typeof request.data !== 'string' || !request.data) return false
  if (format !== 'svg' && !request.data.startsWith(RASTER_DATA_PREFIX[format])) return false
  return request.suggestedName === undefined || typeof request.suggestedName === 'string'
}

function safeFileStem(value: string | undefined, fallback: string): string {
  const safe = (value || fallback)
    .replace(/[/\\:*?"<>|]/g, '_')
    .slice(0, 80)
    .trim()
  return safe || fallback
}

const closeSaveWaiters = new Map<number, (ok: boolean) => void>()
const saveWaiters = new Map<number, (ok: boolean) => void>()

let fileSavedHook: ((wc: WebContents, path: string) => void) | null = null

export function setDiagramFileSavedHook(hook: (wc: WebContents, path: string) => void): void {
  fileSavedHook = hook
}

export function sendDiagramExportRequest(contents: WebContents, format: ExportFormat): void {
  if (!contents.isDestroyed()) contents.send(DIAGRAM_CHANNELS.exportRequest, format)
}

export function diagramIsDirty(webContentsId: number): boolean {
  return dirtyByWc.has(webContentsId)
}

export function diagramFilePath(webContentsId: number): string | undefined {
  return savePathByWc.get(webContentsId)
}

export function diagramFileRenamed(contents: WebContents, oldPath: string, newPath: string): void {
  const wcId = contents.id
  if (savePathByWc.get(wcId) === oldPath) savePathByWc.set(wcId, newPath)
  if (openPathByWc.get(wcId) === oldPath) openPathByWc.set(wcId, newPath)
  const allowed = allowedByWc.get(wcId)
  if (allowed?.has(oldPath)) allowed.add(newPath)
  if (!contents.isDestroyed()) contents.send(DIAGRAM_CHANNELS.fileRenamed, newPath)
}

export async function requestDiagramClose(
  contents: WebContents,
  parent?: BrowserWindow | null,
): Promise<boolean> {
  if (!dirtyByWc.has(contents.id) || contents.isDestroyed()) return true
  const options = {
    type: 'warning' as const,
    message: tm('closeUnsavedMsg'),
    detail: tm('closeUnsavedDetail'),
    buttons: [tm('btnSave'), tm('btnDontSave'), tm('btnCancel')],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  }
  const win = parent ?? BrowserWindow.fromWebContents(contents) ?? undefined
  const choice = win
    ? (await dialog.showMessageBox(win, options)).response
    : (await dialog.showMessageBox(options)).response

  if (choice === 1) return true // Don't Save
  if (choice === 2) return false // Cancel

  // Save flow
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      closeSaveWaiters.delete(contents.id)
      resolve(false)
    }, 60_000)
    closeSaveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(DIAGRAM_CHANNELS.closeSaveRequest)
  })
}

export function requestDiagramSave(contents: WebContents, mode: SaveMode): Promise<boolean> {
  if (contents.isDestroyed()) return Promise.resolve(false)
  if (mode === 'save' && !dirtyByWc.has(contents.id) && savePathByWc.has(contents.id)) {
    return Promise.resolve(true)
  }
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      saveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    saveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(DIAGRAM_CHANNELS.saveRequest, mode)
  })
}

async function resolveSaveTarget(
  e: Electron.IpcMainInvokeEvent,
  mode: SaveMode,
  format: 'drawio' | 'mermaid',
  suggestedName?: string,
): Promise<string | null | 'canceled'> {
  const current = savePathByWc.get(e.sender.id)
  if (mode === 'save' && current) return current

  const ext = format === 'mermaid' ? 'mmd' : 'drawio'
  const filterName = format === 'mermaid' ? tm('filterMermaid') : tm('filterDrawio')

  const win =
    BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
  const defaultPath = current
    ? join(dirname(current), basename(current))
    : join(app.getPath('documents'), `${safeFileStem(suggestedName, tm('untitledFile'))}.${ext}`)

  const picked = await showSaveDialogWithMemory(dialog, win, {
    title: tm('dlgSaveTitle'),
    defaultPath,
    filters: [{ name: filterName, extensions: [ext] }],
  })
  if (picked.canceled || !picked.filePath) return 'canceled'
  return picked.filePath
}

let ipcRegistered = false

export function registerDiagramIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  ipcMain.handle(DIAGRAM_CHANNELS.consumePending, (e) => {
    const wcId = e.sender.id
    return {
      path: openPathByWc.get(wcId) ?? null,
      mode: diagramModeByWc.get(wcId) ?? 'visual',
    }
  })

  ipcMain.handle(DIAGRAM_CHANNELS.readFile, async (e, path: unknown) => {
    if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
      throw new Error('diagram: path not granted to this view')
    }
    return readFile(path, 'utf8')
  })

  ipcMain.handle(DIAGRAM_CHANNELS.save, async (e, request: unknown): Promise<SaveDiagramResult> => {
    const waiter = saveWaiters.get(e.sender.id)
    const closeWaiter = closeSaveWaiters.get(e.sender.id)
    saveWaiters.delete(e.sender.id)
    closeSaveWaiters.delete(e.sender.id)

    const fail = (error: string): SaveDiagramResult => {
      waiter?.(false)
      closeWaiter?.(false)
      return { ok: false, error }
    }
    if (!isSaveDiagramRequest(request)) return fail('diagram: invalid save request')

    const format = request.format || (request.content.trim().startsWith('<') ? 'drawio' : 'mermaid')
    let target: string | null | 'canceled'
    try {
      target = await resolveSaveTarget(e, request.mode, format, request.suggestedName)
    } catch (err) {
      return fail((err as Error).message)
    }

    if (target === 'canceled' || !target) {
      waiter?.(false)
      closeWaiter?.(false)
      return { ok: true, canceled: true }
    }

    try {
      await atomicWriteFile(target, Buffer.from(request.content, 'utf8'))
      savePathByWc.set(e.sender.id, target)
      openPathByWc.set(e.sender.id, target)
      let allowed = allowedByWc.get(e.sender.id)
      if (!allowed) {
        allowed = new Set<string>()
        allowedByWc.set(e.sender.id, allowed)
      }
      allowed.add(target)
      dirtyByWc.delete(e.sender.id)

      fileSavedHook?.(e.sender, target)
      waiter?.(true)
      closeWaiter?.(true)
      return { ok: true, path: target }
    } catch (err) {
      return fail((err as Error).message)
    }
  })

  ipcMain.on(DIAGRAM_CHANNELS.dirtyChanged, (e, dirty: unknown) => {
    if (dirty === true) dirtyByWc.add(e.sender.id)
    else dirtyByWc.delete(e.sender.id)
  })

  ipcMain.on(DIAGRAM_CHANNELS.closeSaveResult, (e, ok: unknown) => {
    const waiter = closeSaveWaiters.get(e.sender.id)
    closeSaveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  ipcMain.on(DIAGRAM_CHANNELS.saveRequestAck, (e, ok: unknown) => {
    const waiter = saveWaiters.get(e.sender.id)
    saveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  ipcMain.handle(DIAGRAM_CHANNELS.exportSvg, async (e, request: unknown): Promise<ExportResult> => {
    if (!isExportFileRequest(request, 'svg')) {
      return { ok: false, error: 'diagram: invalid SVG export request' }
    }
    const safeName = safeFileStem(request.suggestedName, 'diagram')
    const win =
      BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
    const defaultPath = join(app.getPath('documents'), `${safeName}.svg`)
    const picked = await showSaveDialogWithMemory(dialog, win, {
      title: tm('filterSvg'),
      defaultPath,
      filters: [{ name: tm('filterSvg'), extensions: ['svg'] }],
    })
    if (picked.canceled || !picked.filePath) return { ok: true, canceled: true }
    try {
      await atomicWriteFile(picked.filePath, Buffer.from(request.data, 'utf8'))
      return { ok: true, path: picked.filePath }
    } catch (err) {
      return { ok: false, error: (err as Error).message }
    }
  })

  ipcMain.handle(
    DIAGRAM_CHANNELS.exportImage,
    async (e, request: unknown): Promise<ExportResult> => {
      const format = (request as { format?: unknown } | null)?.format
      if ((format !== 'png' && format !== 'jpeg') || !isExportFileRequest(request, format)) {
        return { ok: false, error: 'diagram: invalid image export request' }
      }
      const filterKey = format === 'jpeg' ? 'filterJpeg' : 'filterPng'
      const extension = format === 'jpeg' ? 'jpg' : format
      const safeName = safeFileStem(request.suggestedName, 'diagram')
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const defaultPath = join(app.getPath('documents'), `${safeName}.${extension}`)
      const picked = await showSaveDialogWithMemory(dialog, win, {
        title: tm(filterKey),
        defaultPath,
        filters: [{ name: tm(filterKey), extensions: [extension] }],
      })
      if (picked.canceled || !picked.filePath) return { ok: true, canceled: true }
      try {
        const base64Data = request.data.replace(/^data:image\/\w+;base64,/, '')
        await atomicWriteFile(picked.filePath, Buffer.from(base64Data, 'base64'))
        return { ok: true, path: picked.filePath }
      } catch (err) {
        return { ok: false, error: (err as Error).message }
      }
    },
  )

  ipcMain.removeHandler(DIAGRAM_CHANNELS.getLanguage)
  ipcMain.handle(DIAGRAM_CHANNELS.getLanguage, () => getUiLang())
}

function grantAndTrack(wc: WebContents, openPath?: string | null, initialMode?: DiagramMode): void {
  const wcId = wc.id
  diagramModeByWc.set(wcId, openPath ? modeForPath(openPath) : (initialMode ?? 'visual'))
  if (openPath && existsSync(openPath)) {
    openPathByWc.set(wcId, openPath)
    savePathByWc.set(wcId, openPath)
    allowedByWc.set(wcId, new Set([openPath]))
  }
  wc.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url)
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })
  wc.once('destroyed', () => {
    openPathByWc.delete(wcId)
    allowedByWc.delete(wcId)
    diagramModeByWc.delete(wcId)
    savePathByWc.delete(wcId)
    dirtyByWc.delete(wcId)
    closeSaveWaiters.get(wcId)?.(false)
    closeSaveWaiters.delete(wcId)
    saveWaiters.get(wcId)?.(false)
    saveWaiters.delete(wcId)
  })
}

export function createDiagramView(
  openPath?: string | null,
  initialMode?: DiagramMode,
): WebContentsView {
  registerDiagramIpc()
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  grantAndTrack(view.webContents, openPath, initialMode)
  if (runtime.rendererUrl) void view.webContents.loadURL(runtime.rendererUrl)
  else if (runtime.rendererFile) void view.webContents.loadFile(runtime.rendererFile)
  return view
}

export function startDiagramStandalone(): void {
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  configureDiagramRuntime({
    preloadPath: join(__dirname, '../preload/index.js'),
    rendererUrl: process.env.ELECTRON_RENDERER_URL,
    rendererFile: join(__dirname, '../renderer/index.html'),
  })
  void app.whenReady().then(() => {
    registerDiagramIpc()
    const win = new BrowserWindow({
      width: 1200,
      height: 850,
      webPreferences: {
        preload: runtime.preloadPath,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
      },
    })
    const argPath = process.argv
      .slice(1)
      .find((a) => /\.(drawio|dio|mmd|mermaid)$/i.test(a) && existsSync(a))
    grantAndTrack(win.webContents, argPath)
    if (runtime.rendererUrl) void win.loadURL(runtime.rendererUrl)
    else if (runtime.rendererFile) void win.loadFile(runtime.rendererFile)
  })
  app.on('window-all-closed', () => app.quit())
}
