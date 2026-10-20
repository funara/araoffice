import { existsSync } from 'node:fs'
import { readFile, unlink, writeFile } from 'node:fs/promises'
import { dirname, join, basename, extname } from 'node:path'
import { BrowserWindow, WebContentsView, app, dialog, ipcMain, shell } from 'electron'

app?.setName?.('ARA Office')
import type { WebContents } from 'electron'
import {
  atomicWriteFile,
  contextMenuLabels,
  installContextMenu,
  installNavigationGuard,
  safeExternalUrl,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
} from '@araoffice/electron-utils'
import { createI18n, getUiLang } from '@araoffice/i18n'
import { PDFDocument } from 'pdf-lib'
import { PDF_CHANNELS } from '../shared/ipc'
import type {
  A11yCheckRequest,
  A11yFixRequest,
  A11yFixResult,
  A11yCheckResult,
  CombinePdfsRequest,
  CombinePdfsResult,
  ExportEmbeddedImagesRequest,
  ExportEmbeddedImagesResult,
  ExportImagesRequest,
  ExportImagesResult,
  ExportTextRequest,
  ExportTextResult,
  ExtractPagesRequest,
  ExtractPagesResult,
  InsertPdfRequest,
  InsertPdfResult,
  SavePdfRequest,
  SavePdfResult,
  DeleteAnnotationsRequest,
  DeleteAnnotationsResult,
  SanitizePdfRequest,
  SanitizePreviewRequest,
  SanitizePreviewResult,
  SanitizeResult,
  SetOutlineRequest,
  SetOutlineResult,
  SplitPdfRequest,
  SplitPdfResult,
} from '../shared/ipc'
import {
  combinePdfsBytes,
  extractPagesBytes,
  insertPdfBytes,
  savePdfToPath,
  sanitizePdfBytes,
  scanSanitizeCounts,
  splitPdfBytes,
} from './save-pdf'
import { applyOutline } from './outline-writer'
import { collectA11yReport, fixA11yMetadataBytes } from './a11y-check'
import { deleteAnnotsBytes } from './delete-annots'
import { collectEmbeddedImages } from './extract-images'

const tDlg = createI18n({
  zh: {
    dlgExportImages: '导出图片到文件夹',
    dlgExtract: '抽取页面为 PDF',
    dlgInsert: '选择要插入的 PDF',
    dlgSplit: '拆分 PDF 到文件夹',
    dlgCombine: '合并 PDF 文件',
    filterPdf: 'PDF 文档',
    filterText: '文本文件',
    dlgExportText: '导出文本',
    closeUnsavedMsg: '此 PDF 有未保存的更改。',
    closeUnsavedDetail: '关闭前是否保存？',
    btnSave: '保存',
    btnDontSave: '不保存',
    btnCancel: '取消',
  },
  en: {
    dlgExportImages: 'Export Images to Folder',
    dlgExtract: 'Extract Pages as PDF',
    dlgInsert: 'Choose a PDF to Insert',
    dlgSplit: 'Split PDF into Folder',
    dlgCombine: 'Combine PDF Files',
    filterPdf: 'PDF Documents',
    filterText: 'Text Files',
    dlgExportText: 'Export Text',
    closeUnsavedMsg: 'This PDF has unsaved changes.',
    closeUnsavedDetail: 'Do you want to save them before closing?',
    btnSave: 'Save',
    btnDontSave: "Don't Save",
    btnCancel: 'Cancel',
  },
  ja: {
    dlgExportImages: '画像をフォルダに書き出す',
    dlgExtract: 'ページを PDF として抽出',
    dlgInsert: '挿入する PDF を選択',
    dlgSplit: 'PDF をフォルダに分割',
    dlgCombine: 'PDF ファイルを結合',
    filterPdf: 'PDF ドキュメント',
    filterText: 'テキストファイル',
    dlgExportText: 'テキストを書き出し',
    closeUnsavedMsg: 'この PDF に未保存の変更があります。',
    closeUnsavedDetail: '閉じる前に保存しますか？',
    btnSave: '保存',
    btnDontSave: '保存しない',
    btnCancel: 'キャンセル',
  },
  ko: {
    dlgExportImages: '이미지를 폴더로 내보내기',
    dlgExtract: '페이지를 PDF로 추출',
    dlgInsert: '삽입할 PDF 선택',
    dlgSplit: 'PDF를 폴더로 분할',
    dlgCombine: 'PDF 파일 병합',
    filterPdf: 'PDF 문서',
    filterText: '텍스트 파일',
    dlgExportText: '텍스트 내보내기',
    closeUnsavedMsg: '이 PDF에 저장하지 않은 변경 사항이 있습니다.',
    closeUnsavedDetail: '닫기 전에 저장하시겠습니까?',
    btnSave: '저장',
    btnDontSave: '저장 안 함',
    btnCancel: '취소',
  },
  fr: {
    dlgExportImages: 'Exporter les images vers un dossier',
    dlgExtract: 'Extraire les pages en PDF',
    dlgInsert: 'Choisir un PDF à insérer',
    dlgSplit: 'Diviser le PDF dans un dossier',
    dlgCombine: 'Combiner des fichiers PDF',
    filterPdf: 'Documents PDF',
    filterText: 'Fichiers texte',
    dlgExportText: 'Exporter le texte',
    closeUnsavedMsg: 'Ce PDF contient des modifications non enregistrées.',
    closeUnsavedDetail: 'Voulez-vous les enregistrer avant de fermer ?',
    btnSave: 'Enregistrer',
    btnDontSave: 'Ne pas enregistrer',
    btnCancel: 'Annuler',
  },
  de: {
    dlgExportImages: 'Bilder in Ordner exportieren',
    dlgExtract: 'Seiten als PDF extrahieren',
    dlgInsert: 'Einzufügendes PDF wählen',
    dlgSplit: 'PDF in Ordner aufteilen',
    dlgCombine: 'PDF-Dateien zusammenführen',
    filterPdf: 'PDF-Dokumente',
    filterText: 'Textdateien',
    dlgExportText: 'Text exportieren',
    closeUnsavedMsg: 'Dieses PDF enthält ungespeicherte Änderungen.',
    closeUnsavedDetail: 'Vor dem Schließen speichern?',
    btnSave: 'Speichern',
    btnDontSave: 'Nicht speichern',
    btnCancel: 'Abbrechen',
  },
  es: {
    dlgExportImages: 'Exportar imágenes a una carpeta',
    dlgExtract: 'Extraer páginas como PDF',
    dlgInsert: 'Elegir un PDF para insertar',
    dlgSplit: 'Dividir PDF en carpeta',
    dlgCombine: 'Combinar archivos PDF',
    filterPdf: 'Documentos PDF',
    filterText: 'Archivos de texto',
    dlgExportText: 'Exportar texto',
    closeUnsavedMsg: 'Este PDF tiene cambios sin guardar.',
    closeUnsavedDetail: '¿Quieres guardarlos antes de cerrar?',
    btnSave: 'Guardar',
    btnDontSave: 'No guardar',
    btnCancel: 'Cancelar',
  },
  th: {
    dlgExportImages: 'ส่งออกรูปภาพไปยังโฟลเดอร์',
    dlgExtract: 'แยกหน้าเป็น PDF',
    dlgInsert: 'เลือก PDF ที่จะแทรก',
    dlgSplit: 'แยก PDF ไปยังโฟลเดอร์',
    dlgCombine: 'รวมไฟล์ PDF',
    filterPdf: 'เอกสาร PDF',
    filterText: 'ไฟล์ข้อความ',
    dlgExportText: 'ส่งออกข้อความ',
    closeUnsavedMsg: 'PDF นี้มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก',
    closeUnsavedDetail: 'ต้องการบันทึกก่อนปิดหรือไม่?',
    btnSave: 'บันทึก',
    btnDontSave: 'ไม่บันทึก',
    btnCancel: 'ยกเลิก',
  },
  id: {
    dlgExportImages: 'Ekspor gambar ke folder',
    dlgExtract: 'Ekstrak halaman sebagai PDF',
    dlgInsert: 'Pilih PDF untuk disisipkan',
    dlgSplit: 'Pisah PDF ke Folder',
    dlgCombine: 'Gabung Berkas PDF',
    filterPdf: 'Dokumen PDF',
    filterText: 'File Teks',
    dlgExportText: 'Ekspor Teks',
    closeUnsavedMsg: 'PDF ini memiliki perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
  },
  ru: {
    dlgExportImages: 'Экспорт изображений в папку',
    dlgExtract: 'Извлечь страницы в PDF',
    dlgInsert: 'Выберите PDF для вставки',
    dlgSplit: 'Разделить PDF в папку',
    dlgCombine: 'Объединить файлы PDF',
    filterPdf: 'Документы PDF',
    filterText: 'Текстовые файлы',
    dlgExportText: 'Экспорт текста',
    closeUnsavedMsg: 'В этом PDF есть несохранённые изменения.',
    closeUnsavedDetail: 'Сохранить их перед закрытием?',
    btnSave: 'Сохранить',
    btnDontSave: 'Не сохранять',
    btnCancel: 'Отмена',
  },
  ar: {
    dlgExportImages: 'تصدير الصور إلى مجلد',
    dlgExtract: 'استخراج الصفحات كملف PDF',
    dlgInsert: 'اختر PDF للإدراج',
    dlgSplit: 'تقسيم PDF إلى مجلد',
    dlgCombine: 'دمج ملفات PDF',
    filterPdf: 'مستندات PDF',
    filterText: 'ملفات نصية',
    dlgExportText: 'تصدير النص',
    closeUnsavedMsg: 'يحتوي هذا الـ PDF على تغييرات غير محفوظة.',
    closeUnsavedDetail: 'هل تريد حفظها قبل الإغلاق؟',
    btnSave: 'حفظ',
    btnDontSave: 'عدم الحفظ',
    btnCancel: 'إلغاء',
  },
  pt: {
    dlgExportImages: 'Exportar imagens para pasta',
    dlgExtract: 'Extrair páginas como PDF',
    dlgInsert: 'Escolher um PDF para inserir',
    dlgSplit: 'Dividir PDF em pasta',
    dlgCombine: 'Combinar arquivos PDF',
    filterPdf: 'Documentos PDF',
    filterText: 'Arquivos de texto',
    dlgExportText: 'Exportar texto',
    closeUnsavedMsg: 'Este PDF tem alterações não salvas.',
    closeUnsavedDetail: 'Deseja salvá-las antes de fechar?',
    btnSave: 'Salvar',
    btnDontSave: 'Não Salvar',
    btnCancel: 'Cancelar',
  },
  it: {
    dlgExportImages: 'Esporta immagini in una cartella',
    dlgExtract: 'Estrai pagine come PDF',
    dlgInsert: 'Scegli un PDF da inserire',
    dlgSplit: 'Dividi PDF in cartella',
    dlgCombine: 'Combina file PDF',
    filterPdf: 'Documenti PDF',
    filterText: 'File di testo',
    dlgExportText: 'Esporta testo',
    closeUnsavedMsg: 'Questo PDF contiene modifiche non salvate.',
    closeUnsavedDetail: 'Vuoi salvarle prima di chiudere?',
    btnSave: 'Salva',
    btnDontSave: 'Non salvare',
    btnCancel: 'Annulla',
  },
  pl: {
    dlgExportImages: 'Eksportuj obrazy do folderu',
    dlgExtract: 'Wyodrębnij strony jako PDF',
    dlgInsert: 'Wybierz PDF do wstawienia',
    dlgSplit: 'Podziel PDF do folderu',
    dlgCombine: 'Połącz pliki PDF',
    filterPdf: 'Dokumenty PDF',
    filterText: 'Pliki tekstowe',
    dlgExportText: 'Eksportuj tekst',
    closeUnsavedMsg: 'Ten PDF ma niezapisane zmiany.',
    closeUnsavedDetail: 'Czy zapisać je przed zamknięciem?',
    btnSave: 'Zapisz',
    btnDontSave: 'Nie zapisuj',
    btnCancel: 'Anuluj',
  },
  nl: {
    dlgExportImages: 'Afbeeldingen naar map exporteren',
    dlgExtract: "Pagina's extraheren als PDF",
    dlgInsert: 'Kies een PDF om in te voegen',
    dlgSplit: 'PDF splitsen naar map',
    dlgCombine: 'PDF-bestanden samenvoegen',
    filterPdf: 'PDF-documenten',
    filterText: 'Tekstbestanden',
    dlgExportText: 'Tekst exporteren',
    closeUnsavedMsg: 'Deze PDF bevat niet-opgeslagen wijzigingen.',
    closeUnsavedDetail: 'Wilt u ze opslaan voordat u sluit?',
    btnSave: 'Opslaan',
    btnDontSave: 'Niet opslaan',
    btnCancel: 'Annuleren',
  },
  ms: {
    dlgExportImages: 'Eksport imej ke folder',
    dlgExtract: 'Ekstrak halaman sebagai PDF',
    dlgInsert: 'Pilih PDF untuk disisipkan',
    dlgSplit: 'Pisah PDF ke folder',
    dlgCombine: 'Gabung fail PDF',
    filterPdf: 'Dokumen PDF',
    filterText: 'Fail teks',
    dlgExportText: 'Eksport teks',
    closeUnsavedMsg: 'PDF ini mempunyai perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
  },
  he: {
    dlgExportImages: 'ייצוא תמונות לתיקייה',
    dlgExtract: 'חילוץ עמודים כ-PDF',
    dlgInsert: 'בחרו PDF להוספה',
    dlgSplit: 'פיצול PDF לתיקייה',
    dlgCombine: 'מיזוג קובצי PDF',
    filterPdf: 'מסמכי PDF',
    filterText: 'קובצי טקסט',
    dlgExportText: 'ייצוא טקסט',
    closeUnsavedMsg: 'ב-PDF הזה יש שינויים שלא נשמרו.',
    closeUnsavedDetail: 'האם לשמור אותם לפני הסגירה?',
    btnSave: 'שמירה',
    btnDontSave: 'אל תשמור',
    btnCancel: 'ביטול',
  },
  hi: {
    dlgExportImages: 'चित्र फ़ोल्डर में निर्यात करें',
    dlgExtract: 'पृष्ठों को PDF के रूप में निकालें',
    dlgInsert: 'सम्मिलित करने के लिए PDF चुनें',
    dlgSplit: 'PDF को फ़ोल्डर में विभाजित करें',
    dlgCombine: 'PDF फ़ाइलें संयोजित करें',
    filterPdf: 'PDF दस्तावेज़',
    filterText: 'टेक्स्ट फ़ाइलें',
    dlgExportText: 'टेक्स्ट निर्यात करें',
    closeUnsavedMsg: 'इस PDF में सहेजे नहीं गए परिवर्तन हैं।',
    closeUnsavedDetail: 'क्या बंद करने से पहले उन्हें सहेजना चाहते हैं?',
    btnSave: 'सहेजें',
    btnDontSave: 'न सहेजें',
    btnCancel: 'रद्द करें',
  },
  'zh-TW': {
    dlgExportImages: '匯出圖片到資料夾',
    dlgExtract: '擷取頁面為 PDF',
    dlgInsert: '選擇要插入的 PDF',
    dlgSplit: '拆分 PDF 到資料夾',
    dlgCombine: '合併 PDF 檔案',
    filterPdf: 'PDF 文件',
    filterText: '文字檔',
    dlgExportText: '匯出文字',
    closeUnsavedMsg: '此 PDF 有未儲存的變更。',
    closeUnsavedDetail: '關閉前是否儲存？',
    btnSave: '儲存',
    btnDontSave: '不儲存',
    btnCancel: '取消',
  },
})
type DlgKey =
  | 'dlgExportImages'
  | 'dlgExportText'
  | 'dlgExtract'
  | 'dlgInsert'
  | 'dlgSplit'
  | 'dlgCombine'
  | 'filterPdf'
  | 'filterText'
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

export function configurePdfRuntime(paths: RuntimePaths): void {
  runtime = paths
}

/** Open path per view, queued at tab creation; the renderer consumes it after mount
 * (avoids did-finish-load races). Kept until the view is destroyed: a reload
 * (View > Reload) remounts the renderer and consumes again — a one-shot entry
 * would strand the tab on "No file to open". */
const openPathByWc = new Map<number, string>()
/** File paths granted to each view — readFile only allows these */
const allowedByWc = new Map<number, Set<string>>()
/** Unsaved-changes flags mirrored from the renderer; drives the save prompt before closing a tab/window */
const dirtyByWc = new Set<number>()
const closeSaveWaiters = new Map<number, (ok: boolean) => void>()

export function pdfIsDirty(webContentsId: number): boolean {
  return dirtyByWc.has(webContentsId)
}

/**
 * Close guard for the pdf renderer: true means proceed with closing.
 * Clean → true; dirty → Save / Don't Save / Cancel. On Save, ask the renderer to
 * write to disk and await the result; on failure or timeout stay open (renderer
 * has already shown the error).
 */
export async function requestPdfClose(
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
  const { response } =
    parent && !parent.isDestroyed()
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
  if (response === 2) return false
  if (response === 1) return true
  return await requestRendererSave(contents)
}

function requestRendererSave(contents: WebContents): Promise<boolean> {
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      closeSaveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    closeSaveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(PDF_CHANNELS.closeSaveRequest)
  })
}

let ipcRegistered = false

function registerPdfIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  ipcMain.handle(PDF_CHANNELS.consumePending, (e) => openPathByWc.get(e.sender.id) ?? null)

  ipcMain.handle(PDF_CHANNELS.readFile, async (e, path: unknown) => {
    if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
      throw new Error('pdf: path not granted to this view')
    }
    const buf = await readFile(path)
    return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)
  })

  ipcMain.handle(PDF_CHANNELS.save, async (e, request: SavePdfRequest): Promise<SavePdfResult> => {
    const path = request?.path
    if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
      return { ok: false, error: 'pdf: path not granted to this view' }
    }
    const target = typeof request.targetPath === 'string' ? request.targetPath : path
    try {
      await savePdfToPath(path, target, request)
      return { ok: true }
    } catch (err) {
      return { ok: false, error: err instanceof Error ? err.message : String(err) }
    }
  })

  ipcMain.handle(
    PDF_CHANNELS.extractPages,
    async (e, request: ExtractPagesRequest): Promise<ExtractPagesResult> => {
      const { path, pages, suggestedName } = request ?? {}
      if (
        typeof path !== 'string' ||
        !allowedByWc.get(e.sender.id)?.has(path) ||
        !Array.isArray(pages)
      ) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showSaveDialogWithMemory(dialog, win, {
        title: tm('dlgExtract'),
        defaultPath: join(dirname(path), String(suggestedName || 'pages.pdf')),
        filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
      })
      if (picked.canceled || !picked.filePath) return { ok: true, canceled: true }
      try {
        const bytes = await extractPagesBytes(new Uint8Array(await readFile(path)), pages)
        await writeFile(picked.filePath, bytes)
        return { ok: true, savedPath: picked.filePath }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.insertPdf,
    async (e, request: InsertPdfRequest): Promise<InsertPdfResult> => {
      const { path, afterPageIndex } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showOpenDialogWithMemory(dialog, win, {
        title: tm('dlgInsert'),
        filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
        properties: ['openFile'],
      })
      const other = picked.filePaths[0]
      if (picked.canceled || !other) return { ok: true, canceled: true }
      try {
        const { merged, count } = await insertPdfBytes(
          new Uint8Array(await readFile(path)),
          new Uint8Array(await readFile(other)),
          typeof afterPageIndex === 'number' ? afterPageIndex : -1,
        )
        await atomicWriteFile(path, merged)
        return { ok: true, insertedCount: count }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.splitPdf,
    async (e, request: SplitPdfRequest): Promise<SplitPdfResult> => {
      const { path, ranges, baseName } = request ?? {}
      if (
        typeof path !== 'string' ||
        !allowedByWc.get(e.sender.id)?.has(path) ||
        !Array.isArray(ranges) ||
        ranges.length === 0
      ) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showOpenDialogWithMemory(dialog, win, {
        title: tm('dlgSplit'),
        properties: ['openDirectory', 'createDirectory'],
      })
      const dir = picked.filePaths[0]
      if (picked.canceled || !dir) return { ok: true, canceled: true }

      const safeBase = String(baseName || 'part').replace(/[/\\:*?"<>|]/g, '_')
      const width = String(ranges.length).length
      const targets = ranges.map((_, i) =>
        join(dir, `${safeBase}-${String(i + 1).padStart(width, '0')}.pdf`),
      )
      // Refuse rather than overwrite: a split that silently replaced earlier output would be
      // unrecoverable. The click is cheap to repeat in another folder.
      const clash = targets.find((t) => existsSync(t))
      if (clash) return { ok: false, error: `pdf: file already exists: ${clash}` }

      try {
        const parts = await splitPdfBytes(new Uint8Array(await readFile(path)), ranges)
        for (const [i, part] of parts.entries()) await atomicWriteFile(targets[i]!, part)
        return { ok: true, savedDir: dir, count: parts.length }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.combinePdfs,
    async (e, request: CombinePdfsRequest): Promise<CombinePdfsResult> => {
      const { bookmarkPerFile } = request ?? {}
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showOpenDialogWithMemory(dialog, win, {
        title: tm('dlgCombine'),
        filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
        properties: ['openFile', 'multiSelections'],
      })
      const files = picked.filePaths
      if (picked.canceled || files.length === 0) return { ok: true, canceled: true }
      const saved = await showSaveDialogWithMemory(dialog, win, {
        title: tm('dlgCombine'),
        defaultPath: join(dirname(files[0]!), 'combined.pdf'),
        filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
      })
      if (saved.canceled || !saved.filePath) return { ok: true, canceled: true }
      // Refuse rather than clobber an existing file, same rule as split
      if (files.includes(saved.filePath)) {
        return { ok: false, error: 'pdf: the output must not be one of the inputs' }
      }
      try {
        const sources = await Promise.all(files.map(async (f) => new Uint8Array(await readFile(f))))
        // Bookmark titles come from the file names, which is the only sensible label the picker gives us
        const titles = files.map((f) => basename(f, extname(f)))
        const merged = await combinePdfsBytes(sources, bookmarkPerFile ? titles : null)
        await atomicWriteFile(saved.filePath, merged)
        return { ok: true, savedPath: saved.filePath, fileCount: files.length }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.sanitizePdf,
    async (e, request: SanitizePdfRequest): Promise<SanitizeResult> => {
      const { path, ...categories } = request ?? {}
      if (
        typeof path !== 'string' ||
        !allowedByWc.get(e.sender.id)?.has(path) ||
        !categories ||
        !Object.values(categories).some(Boolean)
      ) {
        return { ok: false, error: 'pdf: wrong sanitize request for this view' }
      }
      try {
        const result = await sanitizePdfBytes(new Uint8Array(await readFile(path)), categories)
        if (!result) return { ok: false, error: 'pdf: nothing selected to remove' }
        // In place on the granted path, atomically: the renderer reloads the document afterwards,
        // the same way an inserted merge does
        await atomicWriteFile(path, result.bytes)
        return { ok: true, report: result.report }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.sanitizePreview,
    async (e, request: SanitizePreviewRequest): Promise<SanitizePreviewResult> => {
      // Read-only census on the granted path: shapes the dialog's badges, never writes
      const { path } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: wrong sanitize preview request for this view' }
      }
      try {
        const counts = await scanSanitizeCounts(new Uint8Array(await readFile(path)))
        return { ok: true, counts }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.deleteAnnotations,
    async (e, request: DeleteAnnotationsRequest): Promise<DeleteAnnotationsResult> => {
      const { path, targets } = request ?? {}
      if (
        typeof path !== 'string' ||
        !allowedByWc.get(e.sender.id)?.has(path) ||
        !Array.isArray(targets) ||
        targets.some((t) => !t || typeof t.pageIndex !== 'number' || !Array.isArray(t.rect))
      ) {
        return { ok: false, error: 'pdf: wrong delete-annotations request for this view' }
      }
      try {
        // In place on the granted path; the renderer reloads the document afterwards, just
        // like sanitize does, so the panel refetches what the file now holds
        const bytes = await deleteAnnotsBytes(new Uint8Array(await readFile(path)), targets)
        await atomicWriteFile(path, bytes)
        return { ok: true }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.setOutline,
    async (e, request: SetOutlineRequest): Promise<SetOutlineResult> => {
      const { path, nodes } = request ?? {}
      if (
        typeof path !== 'string' ||
        !allowedByWc.get(e.sender.id)?.has(path) ||
        !Array.isArray(nodes) ||
        nodes.some(
          (n) =>
            !n ||
            typeof n.title !== 'string' ||
            typeof n.pageIndex !== 'number' ||
            !Number.isInteger(n.pageIndex) ||
            (n.children !== undefined && !Array.isArray(n.children)),
        )
      ) {
        return { ok: false, error: 'pdf: wrong outline request for this view' }
      }
      try {
        const pdfDoc = await PDFDocument.load(new Uint8Array(await readFile(path)), {
          updateMetadata: false,
        })
        // An empty list clears the sidebar entirely, which is what deleting every bookmark wants
        applyOutline(pdfDoc, nodes)
        await atomicWriteFile(path, await pdfDoc.save({ useObjectStreams: false }))
        return { ok: true }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.a11yCheck,
    async (e, request: A11yCheckRequest): Promise<A11yCheckResult> => {
      const { path } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      try {
        // Read-only: the report never writes, so no flush and no atomic write here
        const report = await collectA11yReport(new Uint8Array(await readFile(path)))
        return { ok: true, report }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.a11yFix,
    async (e, request: A11yFixRequest): Promise<A11yFixResult> => {
      const { path, lang } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      // A tag, not free text: letters/digits/8-char subtags guarded to 32 chars overall
      if (typeof lang !== 'string' || !/^[A-Za-z]{2,3}([A-Za-z0-9-]{1,8}){0,3}$/.test(lang)) {
        return { ok: false, error: 'pdf: language tag to use for /Lang must look like en or zh-CN' }
      }
      try {
        // In place on the granted path; the renderer reloads and re-runs the report after it applies
        const result = await fixA11yMetadataBytes(
          new Uint8Array(await readFile(path)),
          // The file name is the honest default for a missing title, no extension
          path.replace(/^.*[/\\]/, '').replace(/\.pdf$/i, ''),
          lang,
        )
        // Nothing was missing: do not rewrite the file just to spin it
        if (result.fixed === 0) return { ok: true, fixed: 0 }
        await atomicWriteFile(path, result.bytes)
        return { ok: true, fixed: result.fixed }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.exportEmbeddedImages,
    async (e, request: ExportEmbeddedImagesRequest): Promise<ExportEmbeddedImagesResult> => {
      const { path } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      // Declared out here so the catch below can unwind them: files that landed before a
      // mid-write failure are removed instead of left half-exported
      const written: string[] = []
      try {
        const scan = await collectEmbeddedImages(new Uint8Array(await readFile(path)))
        if (scan.images.length === 0) {
          return {
            ok: false,
            error: `pdf: no exportable images (${scan.skipped} skipped as undecodable)`,
          }
        }
        const picked = await showOpenDialogWithMemory(dialog, win, {
          title: tm('dlgExportImages'),
          properties: ['openDirectory', 'createDirectory'],
        })
        const dir = picked.filePaths[0]
        if (picked.canceled || !dir) return { ok: true, canceled: true }

        const width = String(scan.images.length).length
        const targets = scan.images.map((image, i) =>
          join(dir, `img-${String(i + 1).padStart(width, '0')}-p${image.pages[0]}.${image.ext}`),
        )
        // Refuse rather than overwrite, the same rule as split: an existing img-01.jpg must not
        // be silently replaced by a different file's image
        const clash = targets.find((t) => existsSync(t))
        if (clash) return { ok: false, error: `pdf: file already exists: ${clash}` }

        for (const [i, image] of scan.images.entries()) {
          written.push(targets[i]!)
          await atomicWriteFile(targets[i]!, Buffer.from(image.data))
        }
        return {
          ok: true,
          savedDir: dir,
          count: scan.images.length,
          skipped: scan.skipped,
        }
      } catch (err) {
        for (const w of written) {
          try {
            await unlink(w)
          } catch {
            // Preserve the original error; cleanup is best effort
          }
        }
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.exportText,
    async (e, request: ExportTextRequest): Promise<ExportTextResult> => {
      const { text, suggestedName } = request ?? {}
      // The text already lives in the view's search index and no source file is read, so there is
      // no path grant to check here — only that the payload is what we expect
      if (typeof text !== 'string' || text.trim().length === 0) {
        return { ok: false, error: 'pdf: no text to export' }
      }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showSaveDialogWithMemory(dialog, win, {
        title: tm('dlgExportText'),
        defaultPath: String(suggestedName || 'document.txt'),
        filters: [{ name: tm('filterText'), extensions: ['txt'] }],
      })
      if (picked.canceled || !picked.filePath) return { ok: true, canceled: true }
      try {
        await atomicWriteFile(picked.filePath, Buffer.from(text, 'utf8'))
        return { ok: true, savedPath: picked.filePath }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.exportImages,
    async (e, request: ExportImagesRequest): Promise<ExportImagesResult> => {
      const { images, pageNumbers, baseName } = request ?? {}
      if (!Array.isArray(images) || images.length === 0)
        return { ok: false, error: 'pdf: no images' }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showOpenDialogWithMemory(dialog, win, {
        title: tm('dlgExportImages'),
        properties: ['openDirectory', 'createDirectory'],
      })
      const dir = picked.filePaths[0]
      if (picked.canceled || !dir) return { ok: true, canceled: true }
      const safeBase = String(baseName || 'page').replace(/[/\\:*?"<>|]/g, '_')
      const paths = images.map((_, i) => join(dir, `${safeBase}-p${pageNumbers?.[i] ?? i + 1}.png`))
      const originals = new Map<string, Buffer | undefined>()
      const written = new Set<string>()
      try {
        for (const path of new Set(paths)) {
          try {
            originals.set(path, await readFile(path))
          } catch (err) {
            if ((err as NodeJS.ErrnoException).code !== 'ENOENT') throw err
            originals.set(path, undefined)
          }
        }
        for (const [i, b64] of images.entries()) {
          written.add(paths[i])
          await atomicWriteFile(paths[i], Buffer.from(b64, 'base64'))
        }
        return { ok: true, savedDir: dir, count: images.length }
      } catch (err) {
        for (const path of written) {
          try {
            const original = originals.get(path)
            if (original === undefined) await unlink(path)
            else await atomicWriteFile(path, original)
          } catch {
            // Preserve the export error; cleanup is best effort.
          }
        }
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )
  ipcMain.on(PDF_CHANNELS.dirtyChanged, (e, dirty: unknown) => {
    if (dirty === true) dirtyByWc.add(e.sender.id)
    else dirtyByWc.delete(e.sender.id)
  })

  ipcMain.on(PDF_CHANNELS.closeSaveResult, (e, ok: unknown) => {
    const waiter = closeSaveWaiters.get(e.sender.id)
    closeSaveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  // Language channel shared with markdown and other modules (all use 'app:get-language'
  // for the same value); registering twice would throw a duplicate-handler error.
  try {
    ipcMain.handle(PDF_CHANNELS.getLanguage, () => getUiLang())
  } catch {
    /* already registered by another module */
  }
}

function grantAndTrack(wc: WebContents, openPath?: string | null): void {
  const wcId = wc.id
  if (openPath && existsSync(openPath)) {
    openPathByWc.set(wcId, openPath)
    allowedByWc.set(wcId, new Set([openPath]))
  }
  // External links inside the PDF (Link annots with target=_blank) go to the system browser
  wc.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url)
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })
  wc.once('destroyed', () => {
    openPathByWc.delete(wcId)
    allowedByWc.delete(wcId)
    dirtyByWc.delete(wcId)
    closeSaveWaiters.get(wcId)?.(false)
    closeSaveWaiters.delete(wcId)
  })
}

export function createPdfView(openPath?: string | null): WebContentsView {
  registerPdfIpc()
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      // parity with docs/sheets: AI streaming and canvas work must not be
      // throttled when the tab is briefly occluded
      backgroundThrottling: false,
    },
  })
  grantAndTrack(view.webContents, openPath)
  if (runtime.rendererUrl) void view.webContents.loadURL(runtime.rendererUrl)
  else if (runtime.rendererFile) void view.webContents.loadFile(runtime.rendererFile)
  return view
}

/** Standalone window mode: `npm run dev -w @araoffice/pdf`, pdf path passed via argv */
export function startPdfStandalone(): void {
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  configurePdfRuntime({
    preloadPath: join(__dirname, '../preload/index.js'),
    rendererUrl: process.env.ELECTRON_RENDERER_URL,
    rendererFile: join(__dirname, '../renderer/index.html'),
  })
  void app.whenReady().then(() => {
    registerPdfIpc()
    const win = new BrowserWindow({
      width: 1200,
      height: 850,
      webPreferences: {
        preload: runtime.preloadPath,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
        backgroundThrottling: false,
      },
    })
    const argPath = process.argv.slice(1).find((a) => /\.pdf$/i.test(a) && existsSync(a))
    grantAndTrack(win.webContents, argPath)
    if (runtime.rendererUrl) void win.loadURL(runtime.rendererUrl)
    else if (runtime.rendererFile) void win.loadFile(runtime.rendererFile)
  })
  app.on('window-all-closed', () => app.quit())
}
