import { spawn } from 'node:child_process'
import { cpSync, existsSync, readFileSync, readdirSync } from 'node:fs'
import { copyFile, mkdir, stat, writeFile } from 'node:fs/promises'
import { randomUUID } from 'node:crypto'
import { basename, dirname, extname, join } from 'node:path'
import {
  BrowserWindow,
  Menu,
  app,
  clipboard,
  dialog,
  ipcMain,
  nativeImage,
  nativeTheme,
  screen,
  shell,
  webContents,
} from 'electron'
import type { MenuItemConstructorOptions, NativeImage } from 'electron'
import menuDocxIcon1x from './assets/menu-docx.png?asset'
import menuDocxIcon2x from './assets/menu-docx@2x.png?asset'
import menuXlsxIcon1x from './assets/menu-xlsx.png?asset'
import menuXlsxIcon2x from './assets/menu-xlsx@2x.png?asset'
import menuPptxIcon1x from './assets/menu-pptx.png?asset'
import menuPptxIcon2x from './assets/menu-pptx@2x.png?asset'
import menuPdfIcon1x from './assets/menu-pdf.png?asset'
import menuPdfIcon2x from './assets/menu-pdf@2x.png?asset'
import menuMdIcon1x from './assets/menu-md.png?asset'
import menuMdIcon2x from './assets/menu-md@2x.png?asset'
import menuDiagramIcon1x from './assets/menu-diagram.png?asset'
import menuDiagramIcon2x from './assets/menu-diagram@2x.png?asset'
import menuDiagramCodeIcon1x from './assets/menu-diagram-code.png?asset'
import menuDiagramCodeIcon2x from './assets/menu-diagram-code@2x.png?asset'
import menuHomeIcon1x from './assets/menu-home.png?asset'
import menuHomeIcon2x from './assets/menu-home@2x.png?asset'
import { createI18n, isLang, normalizeLang, setUiLang, type Lang } from '@araoffice/i18n'
import {
  contextMenuLabels,
  installContextMenu,
  installNavigationGuard,
  renameWithRetry,
  showOpenDialogWithMemory,
} from '@araoffice/electron-utils'
import { readAppSettings, updateAppSetting, writeAppSetting } from './app-settings'
import {
  araBrandVendor,
  araGithubOwner,
  araGithubRepo,
  araGithubUrl,
  araInstagramUrl,
  araWebsiteUrl,
} from './service-config'

app?.setName?.('ARA Office')
import { ProjectStore } from '@araoffice/project-store'
import { safeExternalUrl } from '@araoffice/electron-utils'
import { WINDOW_CHANNELS } from '../shared/window-api'

import {
  buildDocsMenu,
  configureDocsRuntime,
  docsFileRenamed,
  docsQueryDirty,
  requestDocsClose,
  readRecentFiles,
  readStarredFiles,
  recordRecentFile,
  removeRecentFiles,
  replaceRecentFile,
  registerAiIpc,
  setAiProviderConfigurator,
  registerProjectIpc,
  toggleStarredFile,
  registerDocsIpc,
  setDocsExtraFileMenuItems,
  setDocsMenuGate,
  setDocsShellHooks,
  projectFileRenamed,
  setDocsShellWindow,
  setDocsFileSavedHook,
  setSessionPathResolver,
  defaultSaveDir,
  uniquePathIn,
} from '../../../docs/src/main/docs-main'
import { blankXlsxBuffer } from '../../../sheets/src/gateway/csv-import'
import {
  configureSheetsRuntime,
  getActiveSheetsWebContents,
  hasQueuedWorkbook,
  installSheetsMenu,
  markSheetsShuttingDown,
  requestSheetsClose,
  resolveSheetsSessionPath,
  markSheetsUntitledPath,
  sendSheetsMenuAction,
  sheetsFileRenamed,
  setSheetsCloseTabHook,
  setSheetsExtraFileMenuItems,
  setSheetsShellWindow,
  setSheetsWorkbookOpenedHook,
  stopSheetsSidecar,
} from '../../../sheets/src/main/sheets-main'
import {
  configureSlidesRuntime,
  installSlidesMenu,
  replaceSlidesRecentFile,
  requestSlidesClose,
  setSlidesCloseTabHook,
  setSlidesExtraFileMenuItems,
  setSlidesOpenedHook,
  setSlidesShellWindow,
  slidesFileRenamed,
} from '../../../slides/src/main/slides-main'
import { configurePdfRuntime, requestPdfClose } from '../../../pdf/src/main/pdf-main'
import {
  configureMarkdownRuntime,
  markdownFileRenamed,
  requestMarkdownClose,
  setMarkdownDocxExportedHook,
  setMarkdownFileSavedHook,
} from '../../../markdown/src/main/markdown-main'
import {
  configureDiagramRuntime,
  diagramFileRenamed,
  setDiagramFileSavedHook,
} from '../../../diagram/src/main/diagram-main'
import type {
  DiagramMode,
  RecentEntry,
  RecentPage,
  RenameResult,
  UiTheme,
} from '../shared/home-api'
import { HOME_CHANNELS } from '../shared/home-api'
import type { TabAction, TabKind } from '../shared/tabs-api'
import { TABS_CHANNELS } from '../shared/tabs-api'
import {
  matchesRecentFilter,
  normalizeRecentQuery,
  pageRecentPaths,
  statExistingPaths,
} from './recent-files'
import { normalizePathKey, TabManager } from './tab-manager'
import {
  applyUpdateChannel,
  checkForUpdates,
  clearJustUpdated,
  fetchChangelog,
  getAppUpdateState,
  getJustUpdated,
  initAutoUpdater,
  restartAndInstall,
} from './updater'
import { isUpdateChannel, type UpdateChannel } from '../shared/update-api'

/**
 * ARAOffice unified shell: ONE Electron app, ONE BrowserWindow, hosting the
 * docs and sheets modules as WebContentsView tabs behind a WPS-style tab
 * strip. The shell owns the lifecycle — single-instance lock, file-
 * association routing by extension, and per-active-tab menu switching.
 * Renderers load from each module's build output (apps/docs/out,
 * apps/sheets/out), so build those before running the shell.
 */

// ANY unpacked run (`npm run shell`, `npm run dev`, `npx electron .`) must not
// share the installed app's userData or single-instance lock — otherwise a dev
// run silently quits and forwards its argv to the running installed ARAOffice.
// ARAOFFICE_USER_DATA: test drivers point this at a scratch dir so an
// automated instance can run alongside the dev instance (separate lock).
if (!app.isPackaged)
  app.setPath(
    'userData',
    process.env.ARAOFFICE_USER_DATA ?? join(app.getPath('appData'), 'ARAOffice Dev'),
  )

// The product rename from "AI Office" to ARAOffice changed the userData path; migrate old user data once
if (app.isPackaged) {
  const oldDir = join(app.getPath('appData'), 'AI Office')
  const newDir = app.getPath('userData')
  const newEmpty = !existsSync(newDir) || readdirSync(newDir).length === 0
  if (newEmpty && existsSync(oldDir)) cpSync(oldDir, newDir, { recursive: true })
}

// module build outputs: packaged builds carry them as extraResources
// (resources/modules/*, resources/native/*); dev/unpacked resolves them
// relative to apps/shell in the monorepo layout.
const SIDECAR_EXE = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
const APPS_ROOT = join(app.getAppPath(), '..')
const DOCS_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'docs')
  : join(APPS_ROOT, 'docs', 'out')
const SHEETS_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'sheets')
  : join(APPS_ROOT, 'sheets', 'out')
const SLIDES_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'slides')
  : join(APPS_ROOT, 'slides', 'out')
const PDF_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'pdf')
  : join(APPS_ROOT, 'pdf', 'out')
const MARKDOWN_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'markdown')
  : join(APPS_ROOT, 'markdown', 'out')
const DIAGRAM_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'diagram')
  : join(APPS_ROOT, 'diagram', 'out')
const SIDECAR_BIN = app.isPackaged
  ? join(process.resourcesPath, 'native', SIDECAR_EXE)
  : join(APPS_ROOT, 'sheets', 'native', 'xlsx-engine', 'target', 'release', SIDECAR_EXE)

/**
 * Renderer origin for an unpacked (dev) run.
 *
 * Set explicitly by <APP>_RENDERER_URL, else the documented dev port. The
 * fallback exists because the POSIX-only `VAR=value npm run ...` form used by
 * the root dev scripts does not parse under cmd.exe, which left every module
 * without a URL and made `npm run dev` unusable on Windows.
 *
 * Returns undefined - so the runtime loads its built index.html - when packaged
 * or when that build already exists. `npm run shell` builds before launching
 * and starts no dev servers, so guessing a dev port there would replace a
 * working built renderer with a dead localhost URL.
 */
function devUrl(envName: string, port: number, builtFile: string): string | undefined {
  if (app.isPackaged) return undefined
  const configured = process.env[envName]
  if (configured) return configured
  return existsSync(builtFile) ? undefined : `http://localhost:${port}`
}

configureDocsRuntime({
  preloadPath: join(DOCS_OUT, 'preload', 'index.js'),
  rendererUrl: devUrl('DOCS_RENDERER_URL', 5173, join(DOCS_OUT, 'renderer', 'index.html')),
  rendererFile: join(DOCS_OUT, 'renderer', 'index.html'),
})
configureSheetsRuntime({
  preloadPath: join(SHEETS_OUT, 'preload', 'index.js'),
  rendererUrl: devUrl('SHEETS_RENDERER_URL', 5174, join(SHEETS_OUT, 'renderer', 'index.html')),
  rendererFile: join(SHEETS_OUT, 'renderer', 'index.html'),
  sidecarPath: SIDECAR_BIN,
})
configureSlidesRuntime({
  preloadPath: join(SLIDES_OUT, 'preload', 'index.js'),
  rendererDevUrl: devUrl('SLIDES_RENDERER_URL', 5175, join(SLIDES_OUT, 'renderer', 'index.html')),
  rendererFilePath: join(SLIDES_OUT, 'renderer', 'index.html'),
})
configurePdfRuntime({
  preloadPath: join(PDF_OUT, 'preload', 'index.js'),
  rendererUrl: devUrl('PDF_RENDERER_URL', 5176, join(PDF_OUT, 'renderer', 'index.html')),
  rendererFile: join(PDF_OUT, 'renderer', 'index.html'),
})
configureMarkdownRuntime({
  preloadPath: join(MARKDOWN_OUT, 'preload', 'index.js'),
  rendererUrl: devUrl('MARKDOWN_RENDERER_URL', 5177, join(MARKDOWN_OUT, 'renderer', 'index.html')),
  rendererFile: join(MARKDOWN_OUT, 'renderer', 'index.html'),
})
configureDiagramRuntime({
  preloadPath: join(DIAGRAM_OUT, 'preload', 'index.js'),
  rendererUrl: devUrl('DIAGRAM_RENDERER_URL', 5178, join(DIAGRAM_OUT, 'renderer', 'index.html')),
  rendererFile: join(DIAGRAM_OUT, 'renderer', 'index.html'),
})

// ---- UI language ----
// Persisted in userData/app-settings.json so the editor modules can read the
// same file when they pick up i18n later. ARAOFFICE_LANG overrides for tests.

const APP_SETTINGS_PATH = () => join(app.getPath('userData'), 'app-settings.json')

let uiLang: Lang | null = null

function currentLang(): Lang {
  if (uiLang) return uiLang
  if (process.env.ARAOFFICE_LANG) {
    uiLang = normalizeLang(process.env.ARAOFFICE_LANG)
    setUiLang(uiLang)
    return uiLang
  }
  const lang = uiLang ?? normalizeLang(app.getLocale())
  uiLang = lang
  setUiLang(lang)
  return lang
}

function persistLang(lang: Lang): void {
  uiLang = lang
  setUiLang(lang)
  void writeAppSetting(APP_SETTINGS_PATH(), 'language', lang).catch((err) =>
    console.error('[shell] language write failed:', err),
  )
}

let cachedUpdateChannel: UpdateChannel | null = null

function currentUpdateChannel(): UpdateChannel {
  cachedUpdateChannel ??= 'stable'
  return cachedUpdateChannel
}

let cachedTheme: UiTheme | null = null

function currentTheme(): UiTheme {
  cachedTheme ??= 'system'
  return cachedTheme
}

const tMain = createI18n({
  zh: {
    menuFile: '文件',
    menuSectionNew: '新建',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: '未命名表格',
    untitledDoc: '未命名文档',
    untitledDeck: '未命名演示文稿',
    untitledMarkdown: '未命名 Markdown',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: '导出为 PDF…',
    menuOpenInDocs: '转换为 Docs 文档并打开',
    menuOpen: '打开…',
    menuSave: '保存',
    menuSaveAs: '另存为…',
    menuClose: '关闭',
    menuEdit: '编辑',
    menuWindow: '窗口',
    menuHome: '首页',
    backToHome: '返回首页',
    dlgOpenTitle: '打开文件',
    filterSupported: '支持的文件',
    filterWord: 'Word 文档',
    filterExcel: 'Excel 工作簿',
    filterPpt: 'PowerPoint 演示文稿',
    filterMarkdown: 'Markdown 文档',
    filterPdf: 'PDF 文档',
    errBadArgs: '参数无效',
    errBadName: '文件名不合法',
    errMissing: '文件不存在',
    errExists: '同名文件已存在',
    errRenameFailed: '重命名失败',
    errNewTabFailed: '新建文档失败',
    errUnsupportedExt: '暂不支持 .{ext} 类型',
    copySuffix: '副本',
    menuHelp: '帮助',
    thirdPartyNotices: '第三方软件声明',
    menuExportDocx: '导出为 Word…',
    btnCancel: '取消',
  },
  en: {
    menuFile: 'File',
    menuSectionNew: 'New',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'Untitled Spreadsheet',
    untitledDoc: 'Untitled Document',
    untitledDeck: 'Untitled Presentation',
    untitledMarkdown: 'Untitled Note',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'Export as PDF…',
    menuOpenInDocs: 'Convert & Open in ARA Docs',
    menuOpen: 'Open…',
    menuSave: 'Save',
    menuSaveAs: 'Save As…',
    menuClose: 'Close',
    menuEdit: 'Edit',
    menuWindow: 'Window',
    menuHome: 'Home',
    backToHome: 'Back to Home',
    dlgOpenTitle: 'Open File',
    filterSupported: 'Supported Files',
    filterWord: 'Word Documents',
    filterExcel: 'Excel Workbooks',
    filterPpt: 'PowerPoint Presentations',
    filterMarkdown: 'Markdown Documents',
    filterPdf: 'PDF Documents',
    errBadArgs: 'Invalid arguments',
    errBadName: 'Invalid file name',
    errMissing: 'File not found',
    errExists: 'A file with that name already exists',
    errRenameFailed: 'Rename failed',
    errNewTabFailed: 'Could not create the new document',
    errUnsupportedExt: '.{ext} files are not supported',
    copySuffix: 'copy',
    menuHelp: 'Help',
    thirdPartyNotices: 'Third-Party Notices',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'Cancel',
  },
  ja: {
    menuFile: 'ファイル',
    menuSectionNew: '新規作成',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: '無題のスプレッドシート',
    untitledDoc: '無題のドキュメント',
    untitledDeck: '無題のプレゼンテーション',
    untitledMarkdown: '無題の Markdown',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'PDF として書き出す…',
    menuOpenInDocs: 'Docs 文書に変換して開く',
    menuOpen: '開く…',
    menuSave: '保存',
    menuSaveAs: '名前を付けて保存…',
    menuClose: '閉じる',
    menuEdit: '編集',
    menuWindow: 'ウィンドウ',
    menuHome: 'ホーム',
    backToHome: 'ホームに戻る',
    dlgOpenTitle: 'ファイルを開く',
    filterSupported: '対応ファイル',
    filterWord: 'Word 文書',
    filterExcel: 'Excel ブック',
    filterPpt: 'PowerPoint プレゼンテーション',
    filterMarkdown: 'Markdown ドキュメント',
    filterPdf: 'PDF ドキュメント',
    errBadArgs: '引数が無効です',
    errBadName: 'ファイル名が無効です',
    errMissing: 'ファイルが見つかりません',
    errExists: '同名のファイルが既に存在します',
    errRenameFailed: '名前の変更に失敗しました',
    errNewTabFailed: '新規ドキュメントを作成できませんでした',
    errUnsupportedExt: '.{ext} 形式には対応していません',
    copySuffix: 'コピー',
    menuHelp: 'ヘルプ',
    thirdPartyNotices: 'サードパーティソフトウェアに関する通知',
    menuExportDocx: 'Word として書き出す…',
    btnCancel: 'キャンセル',
  },
  ko: {
    menuFile: '파일',
    menuSectionNew: '새로 만들기',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: '제목 없는 스프레드시트',
    untitledDoc: '제목 없는 문서',
    untitledDeck: '제목 없는 프레젠테이션',
    untitledMarkdown: '제목 없는 Markdown',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'PDF로 내보내기…',
    menuOpenInDocs: 'Docs 문서로 변환하여 열기',
    menuOpen: '열기…',
    menuSave: '저장',
    menuSaveAs: '다른 이름으로 저장…',
    menuClose: '닫기',
    menuEdit: '편집',
    menuWindow: '창',
    menuHome: '홈',
    backToHome: '홈으로 돌아가기',
    dlgOpenTitle: '파일 열기',
    filterSupported: '지원되는 파일',
    filterWord: 'Word 문서',
    filterExcel: 'Excel 통합 문서',
    filterPpt: 'PowerPoint 프레젠테이션',
    filterMarkdown: 'Markdown 문서',
    filterPdf: 'PDF 문서',
    errBadArgs: '잘못된 인수입니다',
    errBadName: '파일 이름이 잘못되었습니다',
    errMissing: '파일을 찾을 수 없습니다',
    errExists: '같은 이름의 파일이 이미 있습니다',
    errRenameFailed: '이름 바꾸기에 실패했습니다',
    errNewTabFailed: '새 문서를 만들지 못했습니다',
    errUnsupportedExt: '.{ext} 형식은 지원되지 않습니다',
    copySuffix: '복사본',
    menuHelp: '도움말',
    thirdPartyNotices: '타사 소프트웨어 고지',
    menuExportDocx: 'Export as Word…',
    btnCancel: '취소',
  },
  fr: {
    menuFile: 'Fichier',
    menuSectionNew: 'Nouveau',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'Feuille de calcul sans titre',
    untitledDoc: 'Document sans titre',
    untitledDeck: 'Présentation sans titre',
    untitledMarkdown: 'Markdown sans titre',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'Exporter en PDF…',
    menuOpenInDocs: 'Convertir et ouvrir dans ARA Docs',
    menuOpen: 'Ouvrir…',
    menuSave: 'Enregistrer',
    menuSaveAs: 'Enregistrer sous…',
    menuClose: 'Fermer',
    menuEdit: 'Édition',
    menuWindow: 'Fenêtre',
    menuHome: 'Accueil',
    backToHome: "Retour à l'accueil",
    dlgOpenTitle: 'Ouvrir un fichier',
    filterSupported: 'Fichiers pris en charge',
    filterWord: 'Documents Word',
    filterExcel: 'Classeurs Excel',
    filterPpt: 'Présentations PowerPoint',
    filterMarkdown: 'Documents Markdown',
    filterPdf: 'Documents PDF',
    errBadArgs: 'Arguments non valides',
    errBadName: 'Nom de fichier non valide',
    errMissing: 'Fichier introuvable',
    errExists: 'Un fichier du même nom existe déjà',
    errRenameFailed: 'Échec du renommage',
    errNewTabFailed: 'Impossible de créer le nouveau document',
    errUnsupportedExt: 'les fichiers .{ext} ne sont pas pris en charge',
    copySuffix: 'copie',
    menuHelp: 'Aide',
    thirdPartyNotices: 'Mentions relatives aux logiciels tiers',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'Annuler',
  },
  de: {
    menuFile: 'Datei',
    menuSectionNew: 'Neu',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'Unbenannte Tabelle',
    untitledDoc: 'Unbenanntes Dokument',
    untitledDeck: 'Unbenannte Präsentation',
    untitledMarkdown: 'Unbenanntes Markdown',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'Als PDF exportieren…',
    menuOpenInDocs: 'In ARA Docs umwandeln und öffnen',
    menuOpen: 'Öffnen…',
    menuSave: 'Speichern',
    menuSaveAs: 'Speichern unter…',
    menuClose: 'Schließen',
    menuEdit: 'Bearbeiten',
    menuWindow: 'Fenster',
    menuHome: 'Startseite',
    backToHome: 'Zurück zur Startseite',
    dlgOpenTitle: 'Datei öffnen',
    filterSupported: 'Unterstützte Dateien',
    filterWord: 'Word-Dokumente',
    filterExcel: 'Excel-Arbeitsmappen',
    filterPpt: 'PowerPoint-Präsentationen',
    filterMarkdown: 'Markdown-Dokumente',
    filterPdf: 'PDF-Dokumente',
    errBadArgs: 'Ungültige Argumente',
    errBadName: 'Ungültiger Dateiname',
    errMissing: 'Datei nicht gefunden',
    errExists: 'Eine Datei mit diesem Namen existiert bereits',
    errRenameFailed: 'Umbenennen fehlgeschlagen',
    errNewTabFailed: 'Neues Dokument konnte nicht erstellt werden',
    errUnsupportedExt: '.{ext}-Dateien werden nicht unterstützt',
    copySuffix: 'Kopie',
    menuHelp: 'Hilfe',
    thirdPartyNotices: 'Hinweise zu Drittanbietersoftware',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'Abbrechen',
  },
  es: {
    menuFile: 'Archivo',
    menuSectionNew: 'Nuevo',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'Hoja de cálculo sin título',
    untitledDoc: 'Documento sin título',
    untitledDeck: 'Presentación sin título',
    untitledMarkdown: 'Markdown sin título',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'Exportar como PDF…',
    menuOpenInDocs: 'Convertir y abrir en ARA Docs',
    menuOpen: 'Abrir…',
    menuSave: 'Guardar',
    menuSaveAs: 'Guardar como…',
    menuClose: 'Cerrar',
    menuEdit: 'Edición',
    menuWindow: 'Ventana',
    menuHome: 'Inicio',
    backToHome: 'Volver al inicio',
    dlgOpenTitle: 'Abrir archivo',
    filterSupported: 'Archivos compatibles',
    filterWord: 'Documentos de Word',
    filterExcel: 'Libros de Excel',
    filterPpt: 'Presentaciones de PowerPoint',
    filterMarkdown: 'Documentos Markdown',
    filterPdf: 'Documentos PDF',
    errBadArgs: 'Argumentos no válidos',
    errBadName: 'Nombre de archivo no válido',
    errMissing: 'Archivo no encontrado',
    errExists: 'Ya existe un archivo con ese nombre',
    errRenameFailed: 'No se pudo cambiar el nombre',
    errNewTabFailed: 'No se pudo crear el nuevo documento',
    errUnsupportedExt: 'los archivos .{ext} no son compatibles',
    copySuffix: 'copia',
    menuHelp: 'Ayuda',
    thirdPartyNotices: 'Avisos de software de terceros',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'Cancelar',
  },
  th: {
    menuFile: 'ไฟล์',
    menuSectionNew: 'สร้างใหม่',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'สเปรดชีตไม่มีชื่อ',
    untitledDoc: 'เอกสารไม่มีชื่อ',
    untitledDeck: 'งานนำเสนอไม่มีชื่อ',
    untitledMarkdown: 'Markdown ไม่มีชื่อ',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'ส่งออกเป็น PDF…',
    menuOpenInDocs: 'แปลงและเปิดใน ARA Docs',
    menuOpen: 'เปิด…',
    menuSave: 'บันทึก',
    menuSaveAs: 'บันทึกเป็น…',
    menuClose: 'ปิด',
    menuEdit: 'แก้ไข',
    menuWindow: 'หน้าต่าง',
    menuHome: 'หน้าแรก',
    backToHome: 'กลับไปหน้าแรก',
    dlgOpenTitle: 'เปิดไฟล์',
    filterSupported: 'ไฟล์ที่รองรับ',
    filterWord: 'เอกสาร Word',
    filterExcel: 'เวิร์กบุ๊ก Excel',
    filterPpt: 'งานนำเสนอ PowerPoint',
    filterMarkdown: 'เอกสาร Markdown',
    filterPdf: 'เอกสาร PDF',
    errBadArgs: 'อาร์กิวเมนต์ไม่ถูกต้อง',
    errBadName: 'ชื่อไฟล์ไม่ถูกต้อง',
    errMissing: 'ไม่พบไฟล์',
    errExists: 'มีไฟล์ชื่อเดียวกันอยู่แล้ว',
    errRenameFailed: 'เปลี่ยนชื่อไม่สำเร็จ',
    errNewTabFailed: 'สร้างเอกสารใหม่ไม่สำเร็จ',
    errUnsupportedExt: 'ไม่รองรับไฟล์ .{ext}',
    copySuffix: 'สำเนา',
    menuHelp: 'วิธีใช้',
    thirdPartyNotices: 'ประกาศเกี่ยวกับซอฟต์แวร์ของบุคคลที่สาม',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'ยกเลิก',
  },
  id: {
    menuFile: 'File',
    menuSectionNew: 'Baru',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'Spreadsheet tanpa judul',
    untitledDoc: 'Dokumen tanpa judul',
    untitledDeck: 'Presentasi tanpa judul',
    untitledMarkdown: 'Markdown tanpa judul',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'Ekspor sebagai PDF…',
    menuOpenInDocs: 'Konversi & buka di ARA Docs',
    menuOpen: 'Buka…',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuClose: 'Tutup',
    menuEdit: 'Edit',
    menuWindow: 'Jendela',
    menuHome: 'Beranda',
    backToHome: 'Kembali ke Beranda',
    dlgOpenTitle: 'Buka File',
    filterSupported: 'File yang Didukung',
    filterWord: 'Dokumen Word',
    filterExcel: 'Buku Kerja Excel',
    filterPpt: 'Presentasi PowerPoint',
    filterMarkdown: 'Dokumen Markdown',
    filterPdf: 'Dokumen PDF',
    errBadArgs: 'Argumen tidak valid',
    errBadName: 'Nama file tidak valid',
    errMissing: 'File tidak ditemukan',
    errExists: 'File dengan nama tersebut sudah ada',
    errRenameFailed: 'Gagal mengganti nama',
    errNewTabFailed: 'Gagal membuat dokumen baru',
    errUnsupportedExt: 'file .{ext} tidak didukung',
    copySuffix: 'salinan',
    menuHelp: 'Bantuan',
    thirdPartyNotices: 'Pemberitahuan Perangkat Lunak Pihak Ketiga',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'Batal',
  },
  ru: {
    menuFile: 'Файл',
    menuSectionNew: 'Создать',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'Таблица без названия',
    untitledDoc: 'Документ без названия',
    untitledDeck: 'Презентация без названия',
    untitledMarkdown: 'Markdown без названия',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'Экспортировать в PDF…',
    menuOpenInDocs: 'Преобразовать и открыть в ARA Docs',
    menuOpen: 'Открыть…',
    menuSave: 'Сохранить',
    menuSaveAs: 'Сохранить как…',
    menuClose: 'Закрыть',
    menuEdit: 'Правка',
    menuWindow: 'Окно',
    menuHome: 'Главная',
    backToHome: 'Вернуться на главную',
    dlgOpenTitle: 'Открытие файла',
    filterSupported: 'Поддерживаемые файлы',
    filterWord: 'Документы Word',
    filterExcel: 'Книги Excel',
    filterPpt: 'Презентации PowerPoint',
    filterMarkdown: 'Документы Markdown',
    filterPdf: 'Документы PDF',
    errBadArgs: 'Недопустимые аргументы',
    errBadName: 'Недопустимое имя файла',
    errMissing: 'Файл не найден',
    errExists: 'Файл с таким именем уже существует',
    errRenameFailed: 'Не удалось переименовать',
    errNewTabFailed: 'Не удалось создать новый документ',
    errUnsupportedExt: 'файлы .{ext} не поддерживаются',
    copySuffix: 'копия',
    menuHelp: 'Справка',
    thirdPartyNotices: 'Уведомления о стороннем ПО',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'Отмена',
  },
  ar: {
    menuFile: 'ملف',
    menuSectionNew: 'جديد',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'جدول بيانات بلا عنوان',
    untitledDoc: 'مستند بدون عنوان',
    untitledDeck: 'عرض تقديمي بدون عنوان',
    untitledMarkdown: 'Markdown بدون عنوان',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'تصدير بتنسيق PDF…',
    menuOpenInDocs: 'التحويل والفتح في ARA Docs',
    menuOpen: 'فتح…',
    menuSave: 'حفظ',
    menuSaveAs: 'حفظ باسم…',
    menuClose: 'إغلاق',
    menuEdit: 'تحرير',
    menuWindow: 'نافذة',
    menuHome: 'الصفحة الرئيسية',
    backToHome: 'العودة إلى الصفحة الرئيسية',
    dlgOpenTitle: 'فتح ملف',
    filterSupported: 'الملفات المدعومة',
    filterWord: 'مستندات Word',
    filterExcel: 'مصنفات Excel',
    filterPpt: 'عروض PowerPoint التقديمية',
    filterMarkdown: 'مستندات Markdown',
    filterPdf: 'مستندات PDF',
    errBadArgs: 'وسيطات غير صالحة',
    errBadName: 'اسم ملف غير صالح',
    errMissing: 'الملف غير موجود',
    errExists: 'يوجد ملف بالاسم نفسه بالفعل',
    errRenameFailed: 'فشلت إعادة التسمية',
    errNewTabFailed: 'تعذّر إنشاء المستند الجديد',
    errUnsupportedExt: 'ملفات .{ext} غير مدعومة',
    copySuffix: 'نسخة',
    menuHelp: 'تعليمات',
    thirdPartyNotices: 'إشعارات برامج الجهات الخارجية',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'إلغاء',
  },
  pt: {
    menuFile: 'Arquivo',
    menuSectionNew: 'Novo',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'Planilha sem título',
    untitledDoc: 'Documento sem título',
    untitledDeck: 'Apresentação sem título',
    untitledMarkdown: 'Markdown sem título',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'Exportar como PDF…',
    menuOpenInDocs: 'Converter e abrir no ARA Docs',
    menuOpen: 'Abrir…',
    menuSave: 'Salvar',
    menuSaveAs: 'Salvar Como…',
    menuClose: 'Fechar',
    menuEdit: 'Editar',
    menuWindow: 'Janela',
    menuHome: 'Início',
    backToHome: 'Voltar ao início',
    dlgOpenTitle: 'Abrir arquivo',
    filterSupported: 'Arquivos compatíveis',
    filterWord: 'Documentos do Word',
    filterExcel: 'Pastas de trabalho do Excel',
    filterPpt: 'Apresentações do PowerPoint',
    filterMarkdown: 'Documentos Markdown',
    filterPdf: 'Documentos PDF',
    errBadArgs: 'Argumentos inválidos',
    errBadName: 'Nome de arquivo inválido',
    errMissing: 'Arquivo não encontrado',
    errExists: 'Já existe um arquivo com esse nome',
    errRenameFailed: 'Falha ao renomear',
    errNewTabFailed: 'Falha ao criar o novo documento',
    errUnsupportedExt: 'arquivos .{ext} não são suportados',
    copySuffix: 'cópia',
    menuHelp: 'Ajuda',
    thirdPartyNotices: 'Avisos de software de terceiros',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'Cancelar',
  },
  it: {
    menuFile: 'File',
    menuSectionNew: 'Nuovo',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'Foglio di calcolo senza titolo',
    untitledDoc: 'Documento senza titolo',
    untitledDeck: 'Presentazione senza titolo',
    untitledMarkdown: 'Markdown senza titolo',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'Esporta come PDF…',
    menuOpenInDocs: 'Converti e apri in ARA Docs',
    menuOpen: 'Apri…',
    menuSave: 'Salva',
    menuSaveAs: 'Salva con nome…',
    menuClose: 'Chiudi',
    menuEdit: 'Modifica',
    menuWindow: 'Finestra',
    menuHome: 'Home',
    backToHome: 'Torna alla Home',
    dlgOpenTitle: 'Apri file',
    filterSupported: 'File supportati',
    filterWord: 'Documenti Word',
    filterExcel: 'Cartelle di lavoro Excel',
    filterPpt: 'Presentazioni PowerPoint',
    filterMarkdown: 'Documenti Markdown',
    filterPdf: 'Documenti PDF',
    errBadArgs: 'Argomenti non validi',
    errBadName: 'Nome file non valido',
    errMissing: 'File non trovato',
    errExists: 'Esiste già un file con questo nome',
    errRenameFailed: 'Impossibile rinominare',
    errNewTabFailed: 'Impossibile creare il nuovo documento',
    errUnsupportedExt: 'i file .{ext} non sono supportati',
    copySuffix: 'copia',
    menuHelp: 'Aiuto',
    thirdPartyNotices: 'Note sul software di terze parti',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'Annulla',
  },
  pl: {
    menuFile: 'Plik',
    menuSectionNew: 'Nowy',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'Arkusz bez tytułu',
    untitledDoc: 'Dokument bez tytułu',
    untitledDeck: 'Prezentacja bez tytułu',
    untitledMarkdown: 'Markdown bez tytułu',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'Eksportuj jako PDF…',
    menuOpenInDocs: 'Konwertuj i otwórz w ARA Docs',
    menuOpen: 'Otwórz…',
    menuSave: 'Zapisz',
    menuSaveAs: 'Zapisz jako…',
    menuClose: 'Zamknij',
    menuEdit: 'Edycja',
    menuWindow: 'Okno',
    menuHome: 'Strona główna',
    backToHome: 'Wróć do strony głównej',
    dlgOpenTitle: 'Otwieranie pliku',
    filterSupported: 'Obsługiwane pliki',
    filterWord: 'Dokumenty programu Word',
    filterExcel: 'Skoroszyty programu Excel',
    filterPpt: 'Prezentacje programu PowerPoint',
    filterMarkdown: 'Dokumenty Markdown',
    filterPdf: 'Dokumenty PDF',
    errBadArgs: 'Nieprawidłowe argumenty',
    errBadName: 'Nieprawidłowa nazwa pliku',
    errMissing: 'Nie znaleziono pliku',
    errExists: 'Plik o tej nazwie już istnieje',
    errRenameFailed: 'Nie udało się zmienić nazwy',
    errNewTabFailed: 'Nie udało się utworzyć nowego dokumentu',
    errUnsupportedExt: 'pliki .{ext} nie są obsługiwane',
    copySuffix: 'kopia',
    menuHelp: 'Pomoc',
    thirdPartyNotices: 'Informacje o oprogramowaniu innych firm',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'Anuluj',
  },
  nl: {
    menuFile: 'Bestand',
    menuSectionNew: 'Nieuw',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'Naamloze spreadsheet',
    untitledDoc: 'Naamloos document',
    untitledDeck: 'Naamloze presentatie',
    untitledMarkdown: 'Naamloos Markdown',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'Exporteren als PDF…',
    menuOpenInDocs: 'Converteren en openen in ARA Docs',
    menuOpen: 'Openen…',
    menuSave: 'Opslaan',
    menuSaveAs: 'Opslaan als…',
    menuClose: 'Sluiten',
    menuEdit: 'Bewerken',
    menuWindow: 'Venster',
    menuHome: 'Start',
    backToHome: 'Terug naar start',
    dlgOpenTitle: 'Bestand openen',
    filterSupported: 'Ondersteunde bestanden',
    filterWord: 'Word-documenten',
    filterExcel: 'Excel-werkmappen',
    filterPpt: 'PowerPoint-presentaties',
    filterMarkdown: 'Markdown-documenten',
    filterPdf: 'PDF-documenten',
    errBadArgs: 'Ongeldige argumenten',
    errBadName: 'Ongeldige bestandsnaam',
    errMissing: 'Bestand niet gevonden',
    errExists: 'Er bestaat al een bestand met die naam',
    errRenameFailed: 'Naam wijzigen mislukt',
    errNewTabFailed: 'Kan het nieuwe document niet maken',
    errUnsupportedExt: '.{ext}-bestanden worden niet ondersteund',
    copySuffix: 'kopie',
    menuHelp: 'Help',
    thirdPartyNotices: 'Kennisgevingen over software van derden',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'Annuleren',
  },
  ms: {
    menuFile: 'Fail',
    menuSectionNew: 'Baharu',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'Hamparan tanpa tajuk',
    untitledDoc: 'Dokumen tanpa tajuk',
    untitledDeck: 'Persembahan tanpa tajuk',
    untitledMarkdown: 'Markdown tanpa tajuk',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'Eksport sebagai PDF…',
    menuOpenInDocs: 'Tukar & buka dalam ARA Docs',
    menuOpen: 'Buka…',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuClose: 'Tutup',
    menuEdit: 'Edit',
    menuWindow: 'Tetingkap',
    menuHome: 'Laman Utama',
    backToHome: 'Kembali ke Laman Utama',
    dlgOpenTitle: 'Buka Fail',
    filterSupported: 'Fail yang Disokong',
    filterWord: 'Dokumen Word',
    filterExcel: 'Buku Kerja Excel',
    filterPpt: 'Persembahan PowerPoint',
    filterMarkdown: 'Dokumen Markdown',
    filterPdf: 'Dokumen PDF',
    errBadArgs: 'Argumen tidak sah',
    errBadName: 'Nama fail tidak sah',
    errMissing: 'Fail tidak ditemui',
    errExists: 'Fail dengan nama yang sama sudah wujud',
    errRenameFailed: 'Gagal menamakan semula',
    errNewTabFailed: 'Gagal mencipta dokumen baharu',
    errUnsupportedExt: 'fail .{ext} tidak disokong',
    copySuffix: 'salinan',
    menuHelp: 'Bantuan',
    thirdPartyNotices: 'Notis Perisian Pihak Ketiga',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'Batal',
  },
  he: {
    menuFile: 'קובץ',
    menuSectionNew: 'חדש',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'גיליון אלקטרוני ללא שם',
    untitledDoc: 'מסמך ללא שם',
    untitledDeck: 'מצגת ללא שם',
    untitledMarkdown: 'Markdown ללא שם',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'ייצוא כ-PDF…',
    menuOpenInDocs: 'המרה ופתיחה ב-ARA Docs',
    menuOpen: 'פתיחה…',
    menuSave: 'שמירה',
    menuSaveAs: 'שמירה בשם…',
    menuClose: 'סגירה',
    menuEdit: 'עריכה',
    menuWindow: 'חלון',
    menuHome: 'דף הבית',
    backToHome: 'חזרה לדף הבית',
    dlgOpenTitle: 'פתיחת קובץ',
    filterSupported: 'קבצים נתמכים',
    filterWord: 'מסמכי Word',
    filterExcel: 'חוברות עבודה של Excel',
    filterPpt: 'מצגות PowerPoint',
    filterMarkdown: 'מסמכי Markdown',
    filterPdf: 'מסמכי PDF',
    errBadArgs: 'ארגומנטים לא חוקיים',
    errBadName: 'שם קובץ לא חוקי',
    errMissing: 'הקובץ לא נמצא',
    errExists: 'כבר קיים קובץ באותו שם',
    errRenameFailed: 'שינוי השם נכשל',
    errNewTabFailed: 'יצירת המסמך החדש נכשלה',
    errUnsupportedExt: 'קובצי .{ext} אינם נתמכים',
    copySuffix: 'עותק',
    menuHelp: 'עזרה',
    thirdPartyNotices: 'הודעות על תוכנות צד שלישי',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'ביטול',
  },
  hi: {
    menuFile: 'फ़ाइल',
    menuSectionNew: 'नया',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: 'शीर्षकहीन स्प्रेडशीट',
    untitledDoc: 'बिना शीर्षक दस्तावेज़',
    untitledDeck: 'बिना शीर्षक प्रस्तुति',
    untitledMarkdown: 'अनाम Markdown',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: 'PDF के रूप में निर्यात…',
    menuOpenInDocs: 'ARA Docs में बदलें और खोलें',
    menuOpen: 'खोलें…',
    menuSave: 'सहेजें',
    menuSaveAs: 'इस रूप में सहेजें…',
    menuClose: 'बंद करें',
    menuEdit: 'संपादन',
    menuWindow: 'विंडो',
    menuHome: 'होम',
    backToHome: 'होम पर वापस जाएँ',
    dlgOpenTitle: 'फ़ाइल खोलें',
    filterSupported: 'समर्थित फ़ाइलें',
    filterWord: 'Word दस्तावेज़',
    filterExcel: 'Excel वर्कबुक',
    filterPpt: 'PowerPoint प्रस्तुतियाँ',
    filterMarkdown: 'Markdown दस्तावेज़',
    filterPdf: 'PDF दस्तावेज़',
    errBadArgs: 'अमान्य आर्ग्युमेंट',
    errBadName: 'अमान्य फ़ाइल नाम',
    errMissing: 'फ़ाइल नहीं मिली',
    errExists: 'इस नाम की फ़ाइल पहले से मौजूद है',
    errRenameFailed: 'नाम बदलने में विफल',
    errNewTabFailed: 'नया दस्तावेज़ बनाने में विफल',
    errUnsupportedExt: '.{ext} फ़ाइलें समर्थित नहीं हैं',
    copySuffix: 'प्रतिलिपि',
    menuHelp: 'सहायता',
    thirdPartyNotices: 'तृतीय-पक्ष सॉफ़्टवेयर सूचनाएँ',
    menuExportDocx: 'Export as Word…',
    btnCancel: 'रद्द करें',
  },
  'zh-TW': {
    menuFile: '檔案',
    menuSectionNew: '新增',
    menuNewDoc: 'ARA Docs',
    menuNewSheet: 'ARA Sheets',
    untitledSheet: '未命名試算表',
    untitledDoc: '未命名文件',
    untitledDeck: '未命名簡報',
    untitledMarkdown: '未命名 Markdown',
    menuNewSlide: 'ARA Slides',
    menuNewMarkdown: 'ARA Notes',
    menuExportPdf: '匯出為 PDF…',
    menuOpenInDocs: '轉換為 Docs 文件並開啟',
    menuOpen: '開啟…',
    menuSave: '儲存',
    menuSaveAs: '另存新檔…',
    menuClose: '關閉',
    menuEdit: '編輯',
    menuWindow: '視窗',
    menuHome: '首頁',
    backToHome: '返回首頁',
    dlgOpenTitle: '開啟檔案',
    filterSupported: '支援的檔案',
    filterWord: 'Word 文件',
    filterExcel: 'Excel 活頁簿',
    filterPpt: 'PowerPoint 簡報',
    filterMarkdown: 'Markdown 文件',
    filterPdf: 'PDF 文件',
    errBadArgs: '參數無效',
    errBadName: '檔案名稱不合法',
    errMissing: '檔案不存在',
    errExists: '同名檔案已存在',
    errRenameFailed: '重新命名失敗',
    errNewTabFailed: '新建文件失敗',
    errUnsupportedExt: '暫不支援 .{ext} 類型',
    copySuffix: '副本',
    menuHelp: '說明',
    thirdPartyNotices: '第三方軟體聲明',
    menuExportDocx: '匯出為 Word…',
    btnCancel: '取消',
  },
})

const tm = (key: Parameters<typeof tMain>[1], params?: Parameters<typeof tMain>[2]) =>
  tMain(currentLang(), key, params)

// ---- the shell window + its tab manager (recreated if the user closes it on macOS) ----

let shellWindow: BrowserWindow | null = null
let tabManager: TabManager | null = null
const trustedShellSenders = new Set<number>()

function trustShellSender(contents: Electron.WebContents): void {
  const id = contents.id
  trustedShellSenders.add(id)
  contents.once('destroyed', () => trustedShellSenders.delete(id))
}

function assertTrustedShellSender(
  event: Electron.IpcMainEvent | Electron.IpcMainInvokeEvent,
): void {
  const senderId = event.sender?.id
  if (typeof senderId !== 'number') throw new Error('Untrusted IPC sender.')
  if (trustedShellSenders.has(senderId)) return
  if (shellWindow && !shellWindow.isDestroyed() && shellWindow.webContents.id === senderId) return
  throw new Error('Untrusted IPC sender.')
}
function trustedHandle<Args extends unknown[], Result>(
  channel: string,
  handler: (event: Electron.IpcMainInvokeEvent, ...args: Args) => Result,
): void {
  ipcMain.handle(channel, (event, ...args) => {
    assertTrustedShellSender(event)
    return handler(event, ...(args as Args))
  })
}

function shellIconPath(): string {
  // The app theme belongs to renderer chrome; the window icon follows the OS theme.
  const followsSystemDark = nativeTheme.shouldUseDarkColorsForSystemIntegratedUI
  const filename = followsSystemDark ? 'ARA_dark.png' : 'ARA_light.png'
  return app.isPackaged
    ? join(process.resourcesPath, filename)
    : join(app.getAppPath(), '../../public', filename)
}

function updateShellIcon(win: BrowserWindow | null = shellWindow): void {
  const image = nativeImage.createFromPath(shellIconPath())
  if (image.isEmpty()) return
  if (process.platform === 'darwin') app.dock?.setIcon(image)
  else if (win && !win.isDestroyed()) win.setIcon(image)
}

/**
 * When the user creates a file from a specific project view, remember which
 * project the next save should belong to. key: 'doc' | 'sheet' | 'slide', value: projectId.
 * Consumed by each app's saveHook once the file first hits disk (P1 item 3).
 */
const pendingNewFileProject = new Map<string, string>()

/**
 * P1: after a file first hits disk, if a pending project was set earlier via
 * "create from project view", move the new file into that project automatically.
 * Called from createShellWindow's opened/saved hooks.
 */
function applyPendingProject(filePath: string): void {
  const ext = extname(filePath).slice(1).toLowerCase()
  let key: string | undefined
  if (ext === 'docx') key = 'doc'
  else if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') key = 'sheet'
  else if (ext === 'pptx') key = 'slide'
  else if (ext === 'md' || ext === 'markdown') key = 'markdown'
  if (!key) return
  const projectId = pendingNewFileProject.get(key)
  if (!projectId) return
  pendingNewFileProject.delete(key)
  try {
    const store = new ProjectStore(app.getPath('userData'))
    store.ensureDefaultProject()
    store.resolveProjectForFile(filePath) // assign to default first (idempotent)
    store.moveFileToProject(filePath, projectId)
  } catch (err) {
    console.warn('[shell] applyPendingProject failed:', err)
  }
}

function hideApplicationMenu(): void {
  Menu.setApplicationMenu(null)
}

function registerWindowIpc(): void {
  const targetWindow = (event: Electron.IpcMainInvokeEvent): BrowserWindow | null =>
    BrowserWindow.fromWebContents(event.sender) ?? shellWindow
  const broadcastWindowState = (win: BrowserWindow): void => {
    if (!win.isDestroyed()) win.webContents.send(WINDOW_CHANNELS.stateChanged, win.isMaximized())
  }
  trustedHandle(WINDOW_CHANNELS.minimize, (event) => targetWindow(event)?.minimize())
  trustedHandle(WINDOW_CHANNELS.toggleMaximize, (event) => {
    const win = targetWindow(event)
    if (!win) return
    if (win.isMaximized()) win.restore()
    else win.maximize()
    broadcastWindowState(win)
  })
  trustedHandle(WINDOW_CHANNELS.isMaximized, (event) => targetWindow(event)?.isMaximized() ?? false)
  trustedHandle(WINDOW_CHANNELS.close, (event) => targetWindow(event)?.close())
  trustedHandle(WINDOW_CHANNELS.openExternal, (_event, url: unknown) => {
    const safeUrl = safeExternalUrl(url)
    if (safeUrl) void shell.openExternal(safeUrl)
  })
  trustedHandle(WINDOW_CHANNELS.openThirdPartyNotices, (_event) => {
    void openThirdPartyNotices()
  })
}

function applyMenuFor(kind: TabKind): void {
  hideApplicationMenu()
  switch (kind) {
    case 'docs':
      buildDocsMenu()
      break
    case 'sheets':
      installSheetsMenu()
      break
    case 'slides':
      installSlidesMenu()
      break
    case 'pdf':
      buildPdfMenu()
      break
    case 'markdown':
      buildMarkdownMenu()
      break
    default:
      buildHomeMenu()
  }
}
interface SavedWindowState {
  x?: number
  y?: number
  width?: number
  height?: number
  isMaximized?: boolean
}

// windowState + UI prefs snapshot loaded at startup: createShellWindow must stay synchronous
// (second-instance/open-file create a window on demand and dispatch a file into it immediately).
let initialWindowState: SavedWindowState | undefined

async function primeUiSettings(): Promise<void> {
  const settings = await readAppSettings(APP_SETTINGS_PATH())
  const saved = settings.windowState
  if (saved && typeof saved === 'object' && !Array.isArray(saved)) {
    initialWindowState = saved as SavedWindowState
  }
  const lang = settings.language
  if (isLang(lang)) {
    uiLang = lang
    setUiLang(lang)
  }
  const channel = settings.updateChannel
  if (isUpdateChannel(channel)) cachedUpdateChannel = channel
  const theme = settings.theme
  if (theme === 'light' || theme === 'dark') cachedTheme = theme
}

function getInitialWindowState(): {
  x?: number
  y?: number
  width: number
  height: number
  isMaximized: boolean
} {
  const fallback = { width: 1360, height: 900, isMaximized: true }
  try {
    const saved = initialWindowState
    if (saved && typeof saved.width === 'number' && typeof saved.height === 'number') {
      const width = Math.max(980, Math.min(saved.width, 3840))
      const height = Math.max(600, Math.min(saved.height, 2160))
      if (typeof saved.x === 'number' && typeof saved.y === 'number') {
        const bounds = { x: saved.x, y: saved.y, width, height }
        const matchingDisplay = screen.getDisplayMatching(bounds)
        if (matchingDisplay) {
          return {
            x: saved.x,
            y: saved.y,
            width,
            height,
            isMaximized: saved.isMaximized ?? false,
          }
        }
      }
      return { width, height, isMaximized: saved.isMaximized ?? true }
    }
  } catch {
    /* ignore corrupted state */
  }
  return fallback
}

async function saveWindowState(win: BrowserWindow): Promise<void> {
  try {
    if (win.isDestroyed()) return
    const isMaximized = win.isMaximized()
    const bounds = isMaximized ? undefined : win.getBounds()
    await updateAppSetting(APP_SETTINGS_PATH(), 'windowState', (current) => {
      const prev = (current as SavedWindowState | undefined) ?? {}
      return {
        ...prev,
        isMaximized,
        ...(bounds
          ? {
              x: bounds.x,
              y: bounds.y,
              width: bounds.width,
              height: bounds.height,
            }
          : {}),
      }
    })
  } catch {
    /* ignore window state write errors */
  }
}

function createShellWindow(): void {
  const appIconPath = shellIconPath()
  const initial = getInitialWindowState()
  const win = new BrowserWindow({
    ...(typeof initial.x === 'number' ? { x: initial.x } : {}),
    ...(typeof initial.y === 'number' ? { y: initial.y } : {}),
    width: initial.width,
    height: initial.height,
    icon: appIconPath,
    ...(process.platform === 'win32' ? { frame: false as const } : {}),
    minWidth: 980,
    minHeight: 600,
    title: 'ARAOffice',
    // vibrancy: editor modules punch translucent regions (e.g. the slides
    // thumbnail pane) through to the desktop
    ...(process.platform === 'darwin'
      ? { titleBarStyle: 'hiddenInset' as const, vibrancy: 'sidebar' as const }
      : {}),
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  if (initial.isMaximized) {
    win.maximize()
  }
  let saveTimer: NodeJS.Timeout | null = null
  const scheduleSaveState = () => {
    if (saveTimer) clearTimeout(saveTimer)
    saveTimer = setTimeout(() => {
      void saveWindowState(win)
    }, 400)
  }

  win.on('maximize', () => {
    scheduleSaveState()
    win.webContents.send(WINDOW_CHANNELS.stateChanged, true)
  })
  win.on('unmaximize', () => {
    scheduleSaveState()
    win.webContents.send(WINDOW_CHANNELS.stateChanged, false)
  })
  win.on('resize', scheduleSaveState)
  win.on('move', scheduleSaveState)
  win.on('close', () => void saveWindowState(win))

  shellWindow = win
  trustShellSender(win.webContents)
  updateShellIcon(win)

  const manager = new TabManager(
    win,
    () => win.webContents.send(TABS_CHANNELS.changed, manager.list()),
    applyMenuFor,
    // no extension: these tabs have no file on disk yet; the title becomes the
    // real filename (the localized untitled default + .docx etc.) once the first save lands
    (kind, subKind) =>
      kind === 'docs'
        ? tm('untitledDoc')
        : kind === 'slides'
          ? tm('untitledDeck')
          : kind === 'markdown'
            ? tm('untitledMarkdown')
            : kind === 'diagram'
              ? subKind === 'mermaid' || subKind === 'code'
                ? 'Untitled Diagram Code'
                : 'Untitled Diagram'
              : tm('untitledSheet'),
  )
  tabManager = manager

  // pushRecent-triggered docs menu rebuilds must not clobber the active tab's menu
  setDocsMenuGate(() => manager.list().some((t) => t.active && t.kind === 'docs'))

  setDocsShellWindow(win)
  setSheetsShellWindow(win)
  setSlidesShellWindow(win)
  setDocsShellHooks({
    openTab: (openPath, options) => manager.openDocsTab(openPath, options),
    listTabs: () =>
      manager
        .list()
        .filter((t) => t.kind === 'docs')
        .map((t) => ({ id: t.id, title: t.title, focused: t.active })),
    focusTab: (id) => manager.activateTab(id),
    closeActiveTab: () => manager.closeActiveTab(),
  })
  setSheetsCloseTabHook(() => manager.closeActiveTab())
  setSlidesCloseTabHook(() => manager.closeActiveTab())
  // When ⌘O opens a file inside a tab, sync the tab title/path (used for de-dup by path) and record it as recent.
  // The first save / save-as fires this too, so applyPendingProject also runs here.
  setSheetsWorkbookOpenedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    if (pendingNewSheetSavePaths.delete(normalizePathKey(path))) markSheetsUntitledPath(path)
    trackRecentFile(path)
    applyPendingProject(path)
  })
  setSlidesOpenedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    trackRecentFile(path)
    applyPendingProject(path)
  })
  // docs' save-as / silent first save lands on a new path → sync the tab title too
  setDocsFileSavedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    trackRecentFile(path)
    applyPendingProject(path)
  })
  // markdown untitled first save / Save As lands on a new path
  setMarkdownFileSavedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    trackRecentFile(path)
    applyPendingProject(path)
  })
  // diagram untitled first save / Save As lands on a new path
  setDiagramFileSavedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    trackRecentFile(path)
    applyPendingProject(path)
  })
  // markdown "convert & open in Docs" → route the fresh .docx to a docs tab
  setMarkdownDocxExportedHook((path) => {
    void openDocumentPath(path).catch((err) => console.error('[shell] open failed:', err))
  })

  // Closing the whole window walks every dirty sheets/pdf/slides/docs tab through
  // the same save/don't-save/cancel prompt; any cancel aborts the close.
  // docs dirtiness lives renderer-side, so any live docs tab forces the async path
  // and gets queried there (clean tabs pass through without activation).
  let closeConfirmed = false
  win.on('close', (event) => {
    if (closeConfirmed) return
    const dirtySheets = manager.dirtySheetsTabs()
    const dirtyPdf = manager.dirtyPdfTabs()
    const dirtyMarkdown = manager.dirtyMarkdownTabs()
    const dirtySlides = manager.dirtySlidesTabs()
    const docsTabs = manager.docsTabs()
    if (
      dirtySheets.length === 0 &&
      dirtyPdf.length === 0 &&
      dirtyMarkdown.length === 0 &&
      dirtySlides.length === 0 &&
      docsTabs.length === 0
    )
      return
    event.preventDefault()
    void (async () => {
      for (const tab of dirtySheets) {
        manager.activateTab(tab.id)
        if (!(await requestSheetsClose(tab.webContents, win))) return
      }
      for (const tab of dirtyPdf) {
        manager.activateTab(tab.id)
        if (!(await requestPdfClose(tab.webContents, win))) return
      }
      for (const tab of dirtyMarkdown) {
        manager.activateTab(tab.id)
        if (!(await requestMarkdownClose(tab.webContents, win))) return
      }
      for (const tab of dirtySlides) {
        manager.activateTab(tab.id)
        if (!(await requestSlidesClose(tab.webContents, win))) return
      }
      for (const tab of docsTabs) {
        if (!(await docsQueryDirty(tab.webContents))) continue
        manager.activateTab(tab.id)
        if (!(await requestDocsClose(tab.webContents, win))) return
      }
      closeConfirmed = true
      if (!win.isDestroyed()) win.close()
    })()
  })

  win.on('closed', () => {
    if (shellWindow === win) shellWindow = null
    if (tabManager === manager) tabManager = null
  })

  if (process.env.ELECTRON_RENDERER_URL) {
    void win.loadURL(process.env.ELECTRON_RENDERER_URL)
  } else {
    void win.loadFile(join(__dirname, '../renderer/index.html'))
  }
}

// ---- routing: one dispatch function for every open path ----

const DOCX_RE = /\.docx$/i
const XLSX_RE = /\.(xlsx|xls|csv)$/i
const PPTX_RE = /\.pptx$/i
const PDF_RE = /\.pdf$/i
const MD_RE = /\.(md|markdown)$/i
const DIAGRAM_RE = /\.(drawio|dio|mmd|mermaid)$/i

/** document formats we recognize but don't open — surfaced as a dialog, not silently dropped */
const UNSUPPORTED_DOC_RE = /\.(doc|rtf|odt|ppt|pps|odp|ods|xlsm|xlsb|pages|key|numbers)$/i

/**
 * Single source of truth for the open-dialog filter. Includes the
 * legacy .doc/.ppt binaries so they are selectable and surface the explicit
 * "not supported" dialog via openDocumentPath instead of being grayed out.
 */
const OPEN_DIALOG_EXTENSIONS = [
  'docx',
  'doc',
  'xlsx',
  'xls',
  'csv',
  'pptx',
  'ppt',
  'pdf',
  'md',
  'markdown',
  'drawio',
  'dio',
  'mmd',
  'mermaid',
]
function cleanArgPath(arg: string): string {
  return arg ? arg.replace(/^["']|["']$/g, '').trim() : ''
}

function supportedFileIn(argv: string[]): string | null {
  for (const raw of argv) {
    const arg = cleanArgPath(raw)
    if (
      arg &&
      (DOCX_RE.test(arg) ||
        XLSX_RE.test(arg) ||
        PPTX_RE.test(arg) ||
        PDF_RE.test(arg) ||
        MD_RE.test(arg) ||
        DIAGRAM_RE.test(arg)) &&
      existsSync(arg)
    ) {
      return arg
    }
  }
  return null
}

function unsupportedFileIn(argv: string[]): string | null {
  for (const raw of argv) {
    const arg = cleanArgPath(raw)
    if (arg && UNSUPPORTED_DOC_RE.test(arg) && existsSync(arg)) {
      return arg
    }
  }
  return null
}

function notifyUnsupportedFile(filePath: string): void {
  const clean = cleanArgPath(filePath)
  const ext = extname(clean).slice(1).toLowerCase() || basename(clean)
  const options = {
    type: 'warning' as const,
    message: tm('errUnsupportedExt', { ext }),
  }
  if (shellWindow) {
    shellWindow.show()
    shellWindow.focus()
    void dialog.showMessageBox(shellWindow, options)
  } else {
    void dialog.showMessageBox(options)
  }
}

async function pathExists(path: string): Promise<boolean> {
  try {
    await stat(path)
    return true
  } catch {
    return false
  }
}

function trackRecentFile(path: string): void {
  void recordRecentFile(path).catch((err) => console.error('[shell] recent write failed:', err))
  try {
    app.addRecentDocument(path)
  } catch {
    /* ignore OS recent documents registration failure */
  }
}

/** the single router: extension decides which module owns the file; false = nothing opened */
async function openDocumentPath(rawPath: string): Promise<boolean> {
  const filePath = cleanArgPath(rawPath)
  if (!filePath || !tabManager || !(await pathExists(filePath))) return false
  if (DOCX_RE.test(filePath)) {
    trackRecentFile(filePath)
    const existing = tabManager.findDocsTabByPath(filePath)
    if (existing) tabManager.activateTab(existing)
    else tabManager.openDocsTab(filePath)
    return true
  }
  if (XLSX_RE.test(filePath)) {
    trackRecentFile(filePath)
    const existing = tabManager.findSheetsTabByPath(filePath)
    if (existing) {
      tabManager.activateTab(existing)
    } else {
      tabManager.openSheetsTab(filePath)
      startQueuedWorkbookNudge()
    }
    return true
  }
  if (PPTX_RE.test(filePath)) {
    trackRecentFile(filePath)
    const existing = tabManager.findSlidesTabByPath(filePath)
    if (existing) {
      tabManager.activateTab(existing)
    } else {
      // For a new tab the path goes through the pending queue; the renderer consumes it after mounting
      tabManager.openSlidesTab(filePath)
    }
    return true
  }
  if (PDF_RE.test(filePath)) {
    trackRecentFile(filePath)
    const existing = tabManager.findPdfTabByPath(filePath)
    if (existing) tabManager.activateTab(existing)
    else tabManager.openPdfTab(filePath)
    return true
  }
  if (MD_RE.test(filePath)) {
    trackRecentFile(filePath)
    const existing = tabManager.findMarkdownTabByPath(filePath)
    if (existing) tabManager.activateTab(existing)
    else tabManager.openMarkdownTab(filePath)
    return true
  }
  if (DIAGRAM_RE.test(filePath)) {
    trackRecentFile(filePath)
    const existing = tabManager.findDiagramTabByPath(filePath)
    if (existing) tabManager.activateTab(existing)
    else tabManager.openDiagramTab(filePath)
    return true
  }
  notifyUnsupportedFile(filePath)
  return false
}
const pendingNewSheetSavePaths = new Set<string>()
const sheetsTempDir = (): string => join(app.getPath('temp'), 'araoffice-sheets')

/**
 * New Sheets tabs use a private temporary workbook so the project file is not
 * created until the user completes the first Save dialog.
 */
async function newSheetTab(): Promise<void> {
  try {
    const saveAsPath = await uniquePathIn(await defaultSaveDir(), `${tm('untitledSheet')}.xlsx`)
    const tempDir = sheetsTempDir()
    const tempPath = join(tempDir, `${randomUUID()}.xlsx`)
    await mkdir(tempDir, { recursive: true })
    await writeFile(tempPath, await blankXlsxBuffer())
    // normalizePathKey: the workbook-opened hook reports the OS-cased path from the save dialog
    pendingNewSheetSavePaths.add(normalizePathKey(saveAsPath))
    tabManager?.openSheetsTab(undefined, {
      newBlank: true,
      queuedOpenPath: tempPath,
      suggestSaveAs: saveAsPath,
    })
    startQueuedWorkbookNudge()
  } catch (err) {
    console.warn('[shell] blank workbook create failed, opening in-memory blank tab:', err)
    try {
      tabManager?.openSheetsTab(undefined, { newBlank: true })
    } catch (fallbackErr) {
      surfaceNewTabError(fallbackErr)
    }
  }
}

/**
 * A throw anywhere in the create-tab path (view creation, sidecar resolution,
 * renderer load) used to be swallowed by `void`-ed promises and ipc-invoke
 * rejections, so the click looked like a pure no-op — the exact "AI Sheets /
 * AI Slides do nothing" alpha report. Surface the failure instead.
 */
function surfaceNewTabError(err: unknown): void {
  console.error('[shell] new tab failed:', err)
  const win = shellWindow
  const options = {
    type: 'error' as const,
    message: tm('errNewTabFailed'),
    detail: err instanceof Error ? err.message : String(err),
  }
  void (win && !win.isDestroyed()
    ? dialog.showMessageBox(win, options)
    : dialog.showMessageBox(options))
}

function newDocTab(): void {
  try {
    tabManager?.openDocsTab(undefined, { newBlank: true })
  } catch (err) {
    surfaceNewTabError(err)
  }
}

function newSlideTab(): void {
  try {
    tabManager?.openSlidesTab()
  } catch (err) {
    surfaceNewTabError(err)
  }
}

function newMarkdownTab(): void {
  try {
    tabManager?.openMarkdownTab()
  } catch (err) {
    surfaceNewTabError(err)
  }
}
function newDiagramTab(mode: DiagramMode = 'visual'): void {
  try {
    tabManager?.openDiagramTab(undefined, mode)
  } catch (err) {
    surfaceNewTabError(err)
  }
}

/**
 * The sheets renderer subscribes to menu actions only after Univer finishes
 * mounting (seconds on cold start), so a single 'open' can fire into the
 * void. Re-send to each queued renderer until it consumes its own path.
 */
const workbookNudgeTimers = new Map<number, NodeJS.Timeout>()

function startQueuedWorkbookNudge(): void {
  const target = getActiveSheetsWebContents()
  if (!target || target.isDestroyed()) return
  const existing = workbookNudgeTimers.get(target.id)
  if (existing) clearInterval(existing)
  const startedAt = Date.now()
  sendSheetsMenuAction('open', target)
  const timer = setInterval(() => {
    if (target.isDestroyed() || !hasQueuedWorkbook(target) || Date.now() - startedAt > 30_000) {
      clearInterval(timer)
      workbookNudgeTimers.delete(target.id)
      return
    }
    sendSheetsMenuAction('open', target)
  }, 700)
  workbookNudgeTimers.set(target.id, timer)
}

// ---- home IPC ----

async function statEntries(paths: string[]): Promise<RecentEntry[]> {
  return statExistingPaths(paths, new Set(await readStarredFiles()))
}

function registerHomeIpc(): void {
  trustedHandle(HOME_CHANNELS.getAppVersion, (_event): string => app.getVersion())

  trustedHandle(HOME_CHANNELS.recents, async (_event, query: unknown): Promise<RecentPage> =>
    pageRecentPaths(await readRecentFiles(), query, new Set(await readStarredFiles())),
  )

  // Starred files sort by mtime, which requires stat-ing them all first; they are hand-picked and few, so this is fine
  trustedHandle(HOME_CHANNELS.starred, async (_event, query: unknown): Promise<RecentPage> => {
    const { offset, limit, ext } = normalizeRecentQuery(query)
    const all = (await statEntries(await readStarredFiles())).sort((a, b) => b.mtimeMs - a.mtimeMs)
    const filtered = all.filter((entry) => matchesRecentFilter(entry, ext))
    return {
      entries: limit === 0 ? [] : filtered.slice(offset, offset + limit),
      total: filtered.length,
      totalAll: all.length,
    }
  })

  trustedHandle(HOME_CHANNELS.statPaths, async (_event, paths: unknown): Promise<RecentEntry[]> =>
    statEntries(stringPaths(paths)),
  )

  trustedHandle(HOME_CHANNELS.toggleStar, async (_event, path: unknown) => {
    if (typeof path === 'string') await toggleStarredFile(path)
  })

  trustedHandle(HOME_CHANNELS.openPath, async (_event, path: unknown) => {
    if (typeof path === 'string') await openDocumentPath(path)
  })

  trustedHandle(HOME_CHANNELS.browse, async (event) => {
    const win = BrowserWindow.fromWebContents(event.sender) ?? shellWindow
    if (!win) return
    const result = await showOpenDialogWithMemory(dialog, win, {
      title: tm('dlgOpenTitle'),
      filters: [
        { name: tm('filterSupported'), extensions: OPEN_DIALOG_EXTENSIONS },
        { name: tm('filterWord'), extensions: ['docx', 'doc'] },
        { name: tm('filterExcel'), extensions: ['xlsx', 'xls', 'csv'] },
        { name: tm('filterPpt'), extensions: ['pptx', 'ppt'] },
        { name: tm('filterPdf'), extensions: ['pdf'] },
        { name: tm('filterMarkdown'), extensions: ['md', 'markdown'] },
      ],
      properties: ['openFile'],
    })
    if (!result.canceled && result.filePaths[0]) await openDocumentPath(result.filePaths[0])
  })

  trustedHandle(HOME_CHANNELS.openPdf, async (event) => {
    const win = BrowserWindow.fromWebContents(event.sender) ?? shellWindow
    if (!win) return
    const result = await showOpenDialogWithMemory(dialog, win, {
      title: tm('filterPdf'),
      filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
      properties: ['openFile'],
    })
    if (!result.canceled && result.filePaths[0]) await openDocumentPath(result.filePaths[0])
  })

  trustedHandle(HOME_CHANNELS.newDoc, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('doc', opts.projectId)
    }
    newDocTab()
  })

  trustedHandle(HOME_CHANNELS.newSheet, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('sheet', opts.projectId)
    }
    void newSheetTab()
  })

  trustedHandle(HOME_CHANNELS.newSlide, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('slide', opts.projectId)
    }
    newSlideTab()
  })

  trustedHandle(HOME_CHANNELS.newMarkdown, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('markdown', opts.projectId)
    }
    newMarkdownTab()
  })
  trustedHandle(
    HOME_CHANNELS.newDiagram,
    (_event, opts?: { projectId?: string; mode?: DiagramMode }) => {
      if (opts?.projectId && opts.projectId !== 'default') {
        pendingNewFileProject.set('diagram', opts.projectId)
      }
      newDiagramTab(opts?.mode === 'mermaid' ? 'mermaid' : 'visual')
    },
  )

  trustedHandle(HOME_CHANNELS.removeRecent, async (_event, paths: unknown) => {
    await removeRecentFiles(stringPaths(paths))
  })

  trustedHandle(HOME_CHANNELS.revealPath, async (_event, path: unknown) => {
    if (typeof path === 'string' && (await pathExists(path))) shell.showItemInFolder(path)
  })

  trustedHandle(
    HOME_CHANNELS.renameFile,
    async (_event, path: unknown, newName: unknown): Promise<RenameResult> => {
      if (typeof path !== 'string' || typeof newName !== 'string')
        return { ok: false, error: tm('errBadArgs') }
      const name = newName.trim()
      if (!name || /[\\/:]/.test(name)) return { ok: false, error: tm('errBadName') }
      if (!(await pathExists(path))) return { ok: false, error: tm('errMissing') }
      const target = join(dirname(path), name)
      // normalizePathKey: a case-only rename is the same file on win32 ("a.docx" -> "A.docx")
      if (normalizePathKey(target) === normalizePathKey(path)) return { ok: true, path }
      if (await pathExists(target)) return { ok: false, error: tm('errExists') }
      try {
        await renameWithRetry(path, target)
      } catch (err) {
        return {
          ok: false,
          error: err instanceof Error ? err.message : tm('errRenameFailed'),
        }
      }
      await replaceRecentFile(path, target)
      // project-store's fileMap/chatIdByPath re-key too, so AI chat history follows the file
      projectFileRenamed(path, target)
      // the slides module's own recent list switches to the new path as well (used by the start screen)
      if (/\.pptx$/i.test(target)) void replaceSlidesRecentFile(path, target)
      // open tabs sync their title/path; each editor then syncs its internal save path and title bar
      const affected = tabManager?.renameTabFile(path, target) ?? []
      for (const t of affected) {
        if (t.kind === 'slides') slidesFileRenamed(t.webContents, path, target)
        else if (t.kind === 'docs') docsFileRenamed(t.webContents, path, target)
        else if (t.kind === 'sheets') sheetsFileRenamed(t.webContents, path, target)
        else if (t.kind === 'markdown') markdownFileRenamed(t.webContents, path, target)
        else if (t.kind === 'diagram') diagramFileRenamed(t.webContents, path, target)
      }
      return { ok: true, path: target }
    },
  )

  trustedHandle(HOME_CHANNELS.duplicateFile, async (_event, path: unknown) => {
    if (typeof path !== 'string' || !(await pathExists(path))) return
    const ext = extname(path)
    const base = basename(path, ext)
    const dir = dirname(path)
    for (let i = 1; ; i++) {
      const target = join(dir, `${base} ${tm('copySuffix')}${i === 1 ? '' : ` ${i}`}${ext}`)
      if (await pathExists(target)) continue
      await copyFile(path, target)
      await recordRecentFile(target)
      return
    }
  })

  trustedHandle(HOME_CHANNELS.deleteFiles, async (_event, paths: unknown) => {
    const list = stringPaths(paths)
    for (const p of list) {
      try {
        await shell.trashItem(p)
      } catch {
        // file already gone or trash unavailable; still drop it from the list
      }
    }
    await removeRecentFiles(list)
  })

  trustedHandle(HOME_CHANNELS.openTrash, (_event) => {
    if (process.platform === 'darwin') {
      void shell.openPath(join(app.getPath('home'), '.Trash'))
    } else if (process.platform === 'win32') {
      spawn('explorer.exe', ['shell:RecycleBin'], { detached: true }).unref()
    } else {
      void shell.openPath(join(app.getPath('home'), '.local', 'share', 'Trash', 'files'))
    }
  })

  trustedHandle(HOME_CHANNELS.getLanguage, (_event): Lang => currentLang())

  trustedHandle(HOME_CHANNELS.setLanguage, (_event, lang: unknown) => {
    if (!isLang(lang) || lang === currentLang()) return
    persistLang(lang)
    // the switcher lives on the home page, so the home menu is the active one
    buildHomeMenu()
    installDockMenu()
    installBackToHomeItems()
    for (const wc of webContents.getAllWebContents()) {
      if (!wc.isDestroyed()) {
        try {
          wc.send('app:language-changed', lang)
        } catch {
          /* frame disposed */
        }
      }
    }
  })

  trustedHandle(HOME_CHANNELS.getUpdateChannel, (_event): UpdateChannel => currentUpdateChannel())

  trustedHandle(HOME_CHANNELS.setUpdateChannel, async (_event, channel: unknown) => {
    if (!isUpdateChannel(channel) || channel === currentUpdateChannel()) return
    cachedUpdateChannel = channel
    await writeAppSetting(APP_SETTINGS_PATH(), 'updateChannel', channel)
    applyUpdateChannel(channel)
  })

  trustedHandle(HOME_CHANNELS.onboardingSeen, async (_event) => {
    const settings = await readAppSettings(APP_SETTINGS_PATH())
    return settings.onboardingSeen === true
  })

  trustedHandle(HOME_CHANNELS.setOnboardingSeen, async (_event) => {
    await writeAppSetting(APP_SETTINGS_PATH(), 'onboardingSeen', true)
  })

  trustedHandle(HOME_CHANNELS.getTheme, (_event): UiTheme => currentTheme())

  trustedHandle(HOME_CHANNELS.setTheme, async (_event, theme: unknown) => {
    if (theme !== 'light' && theme !== 'dark' && theme !== 'system') return
    if (theme === currentTheme()) return
    cachedTheme = theme
    nativeTheme.themeSource = theme
    await writeAppSetting(APP_SETTINGS_PATH(), 'theme', theme)
    updateShellIcon()
    for (const wc of webContents.getAllWebContents()) {
      if (!wc.isDestroyed()) {
        try {
          wc.send('app:theme-changed', theme)
        } catch {
          /* frame disposed */
        }
      }
    }
  })

  trustedHandle(HOME_CHANNELS.getUpdateState, (_event) => getAppUpdateState())
  trustedHandle(HOME_CHANNELS.checkForUpdates, (_event) => checkForUpdates())
  trustedHandle(HOME_CHANNELS.restartAndInstall, (_event) => restartAndInstall())
  trustedHandle(HOME_CHANNELS.getChangelog, (_event) => fetchChangelog())
  trustedHandle(HOME_CHANNELS.getJustUpdated, (_event) => getJustUpdated())
  trustedHandle(HOME_CHANNELS.clearJustUpdated, (_event) => clearJustUpdated())
  trustedHandle(HOME_CHANNELS.getSupportLinks, (_event) => {
    const owner = araGithubOwner()
    const repo = araGithubRepo()
    const asUrl = (value: string): string | null =>
      /^https:\/\/[^\s]+$/.test(value) ? value : null
    const vendor = araBrandVendor()
    return {
      issuesUrl: owner && repo ? `https://github.com/${owner}/${repo}/issues` : null,
      websiteUrl: asUrl(araWebsiteUrl()),
      githubUrl: asUrl(araGithubUrl()),
      instagramUrl: asUrl(araInstagramUrl()),
      vendorName: vendor ? vendor : null,
    }
  })
  // Feedback goes to the public issue tracker; there is no feedback service.
}

function stringPaths(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((p): p is string => typeof p === 'string') : []
}

// electron-vite emits ?asset files under hashed names, which breaks nativeImage's
// automatic `@2x` sibling lookup — attach the retina representation by hand
function loadMenuIcon(path1x: string, path2x: string): NativeImage {
  const icon = nativeImage.createFromPath(path1x)
  icon.addRepresentation({ scaleFactor: 2, buffer: readFileSync(path2x) })
  return icon
}

// loaded once, not on every menu open
interface MenuIconSet {
  docx: NativeImage
  xlsx: NativeImage
  pptx: NativeImage
  pdf: NativeImage
  md: NativeImage
  home: NativeImage
  diagram: NativeImage
  diagramCode: NativeImage
}
let menuIconCache: MenuIconSet | null = null
function menuIcons(): MenuIconSet {
  menuIconCache ??= {
    docx: loadMenuIcon(menuDocxIcon1x, menuDocxIcon2x),
    xlsx: loadMenuIcon(menuXlsxIcon1x, menuXlsxIcon2x),
    pptx: loadMenuIcon(menuPptxIcon1x, menuPptxIcon2x),
    pdf: loadMenuIcon(menuPdfIcon1x, menuPdfIcon2x),
    md: loadMenuIcon(menuMdIcon1x, menuMdIcon2x),
    home: loadMenuIcon(menuHomeIcon1x, menuHomeIcon2x),
    diagram: loadMenuIcon(menuDiagramIcon1x, menuDiagramIcon2x),
    diagramCode: loadMenuIcon(menuDiagramCodeIcon1x, menuDiagramCodeIcon2x),
  }
  return menuIconCache
}

const TAB_MENU_ICON: Record<TabKind, keyof MenuIconSet> = {
  home: 'home',
  docs: 'docx',
  sheets: 'xlsx',
  slides: 'pptx',
  pdf: 'pdf',
  markdown: 'md',
  diagram: 'diagram',
}
function registerTabsIpc(): void {
  trustedHandle(TABS_CHANNELS.list, (_event) => tabManager?.list() ?? [])
  trustedHandle(TABS_CHANNELS.activate, (_event, id: string) => tabManager?.activateTab(id))
  trustedHandle(TABS_CHANNELS.close, (_event, id: string) => tabManager?.closeTab(id))
  trustedHandle(TABS_CHANNELS.setOverlayVisible, async (_event, visible: unknown) => {
    return (await tabManager?.setOverlayVisible(Boolean(visible))) ?? null
  })
  trustedHandle(TABS_CHANNELS.getActiveContext, async (_event) => {
    if (!tabManager) return { id: 'home', kind: 'home', title: 'Home' }
    return tabManager.getActiveContext()
  })
  trustedHandle(TABS_CHANNELS.dispatchAction, async (_event, action: unknown) => {
    if (!tabManager || typeof action !== 'object' || !action) {
      return { ok: false, error: 'No tab manager' }
    }
    return tabManager.dispatchTabAction(action as TabAction)
  })
  trustedHandle(TABS_CHANNELS.reorder, (_event, id: string, toIndex: number) => {
    if (typeof id === 'string' && Number.isInteger(toIndex)) tabManager?.reorderTab(id, toIndex)
  })
  // "all tabs" overflow menu — native popup because the editors' WebContentsView
  // would cover any DOM dropdown the shell renderer draws below the tab strip
  trustedHandle(TABS_CHANNELS.showMenu, (_event, x: unknown, y: unknown) => {
    if (!tabManager || !shellWindow) return
    const menu = Menu.buildFromTemplate(
      tabManager.list().map((tab) => ({
        label: tab.title,
        type: 'checkbox' as const,
        checked: tab.active,
        icon:
          tab.kind === 'diagram'
            ? tab.title.includes('Diagram Code') ||
              tab.title.endsWith('.mmd') ||
              tab.title.endsWith('.mermaid')
              ? menuIcons().diagramCode
              : menuIcons().diagram
            : menuIcons()[TAB_MENU_ICON[tab.kind]],
        click: () => tabManager?.activateTab(tab.id),
      })),
    )
    menu.popup({
      window: shellWindow,
      ...(typeof x === 'number' && typeof y === 'number'
        ? { x: Math.round(x), y: Math.round(y) }
        : {}),
    })
  })
  // "+" new-file menu — native for the same reason as the tab list above
  trustedHandle(TABS_CHANNELS.showNewMenu, (_event, x: unknown, y: unknown) => {
    if (!tabManager || !shellWindow) return
    const menu = Menu.buildFromTemplate([
      // enabled:false so pre-Sonoma macOS / Windows (no 'header' support) degrade
      // to an inert label instead of a clickable no-op item
      { label: tm('menuSectionNew'), type: 'header', enabled: false },
      {
        label: tm('menuNewDoc'),
        icon: menuIcons().docx,
        click: () => newDocTab(),
      },
      {
        label: tm('menuNewSheet'),
        icon: menuIcons().xlsx,
        click: () => void newSheetTab(),
      },
      {
        label: tm('menuNewSlide'),
        icon: menuIcons().pptx,
        click: () => newSlideTab(),
      },
      {
        label: tm('menuNewMarkdown'),
        icon: menuIcons().md,
        click: () => newMarkdownTab(),
      },
      {
        label: 'ARA Diagram',
        icon: menuIcons().diagram,
        click: () => newDiagramTab('visual'),
      },
      {
        label: 'ARA Diagram Code',
        icon: menuIcons().diagramCode,
        click: () => newDiagramTab('mermaid'),
      },
      { type: 'separator' },
      { label: tm('menuOpen'), click: () => void openFileViaDialog() },
    ])
    menu.popup({
      window: shellWindow,
      ...(typeof x === 'number' && typeof y === 'number'
        ? { x: Math.round(x), y: Math.round(y) }
        : {}),
    })
  })
  trustedHandle(TABS_CHANNELS.showContextMenu, (_event, id: string, x: unknown, y: unknown) => {
    if (!tabManager || !shellWindow || typeof id !== 'string' || id === 'home') return
    const tab = tabManager.getTab(id)
    if (!tab) return
    const template: MenuItemConstructorOptions[] = [
      {
        label: tm('menuClose'),
        click: () => void tabManager?.closeTab(id),
      },
      {
        label: 'Close Other Tabs',
        click: () => void tabManager?.closeOtherTabs(id),
      },
      {
        label: 'Close Tabs to the Right',
        click: () => void tabManager?.closeTabsToRight(id),
      },
    ]
    if (tab.filePath) {
      template.push({ type: 'separator' })
      template.push({
        label: process.platform === 'darwin' ? 'Reveal in Finder' : 'Show in File Explorer',
        click: () => {
          if (tab.filePath) shell.showItemInFolder(tab.filePath)
        },
      })
      template.push({
        label: 'Copy Path',
        click: () => {
          if (tab.filePath) clipboard.writeText(tab.filePath)
        },
      })
    }
    const menu = Menu.buildFromTemplate(template)
    menu.popup({
      window: shellWindow,
      ...(typeof x === 'number' && typeof y === 'number'
        ? { x: Math.round(x), y: Math.round(y) }
        : {}),
    })
  })
}

// ---- home menu ----

async function openFileViaDialog(): Promise<void> {
  const win = shellWindow ?? BrowserWindow.getFocusedWindow()
  if (!win) return
  const result = await showOpenDialogWithMemory(dialog, win, {
    filters: [{ name: tm('filterSupported'), extensions: OPEN_DIALOG_EXTENSIONS }],
    properties: ['openFile'],
  })
  if (!result.canceled && result.filePaths[0]) await openDocumentPath(result.filePaths[0])
}

function buildHomeMenu(): void {
  Menu.setApplicationMenu(null)
}

// ---- pdf menu (pdf-main has no menu of its own; the shell owns pdf tabs, so it builds one) ----

function buildPdfMenu(): void {
  Menu.setApplicationMenu(null)
}

// ---- markdown menu (markdown-main has no menu of its own; the shell owns markdown tabs) ----

function buildMarkdownMenu(): void {
  Menu.setApplicationMenu(null)
}

function openThirdPartyNotices(): Promise<string> {
  const path = app.isPackaged
    ? join(process.resourcesPath, 'THIRD-PARTY-NOTICES.txt')
    : join(app.getAppPath(), 'build', 'THIRD-PARTY-NOTICES.txt')
  return shell.openPath(path)
}

/** every module's File menu gets a way back to the launcher */
function installBackToHomeItems(): void {
  const backToHomeItem: MenuItemConstructorOptions = {
    label: tm('backToHome'),
    accelerator: 'Shift+CmdOrCtrl+H',
    click: () => tabManager?.openHomeTab(),
  }
  setDocsExtraFileMenuItems([backToHomeItem])
  setSheetsExtraFileMenuItems([backToHomeItem])
  setSlidesExtraFileMenuItems([backToHomeItem])
}

function installDockMenu(): void {
  if (process.platform === 'darwin') {
    app.dock?.setMenu(
      Menu.buildFromTemplate([
        { label: tm('menuHome'), click: () => tabManager?.openHomeTab() },
        { label: tm('menuNewDoc'), click: () => newDocTab() },
        { label: tm('menuNewSheet'), click: () => void newSheetTab() },
        { label: tm('menuNewSlide'), click: () => newSlideTab() },
        { label: tm('menuNewMarkdown'), click: () => newMarkdownTab() },
        { label: 'ARA Diagram', click: () => newDiagramTab('visual') },
        { label: 'ARA Diagram Code', click: () => newDiagramTab('mermaid') },
      ]),
    )
  } else if (process.platform === 'win32') {
    try {
      const resolveResourceIcon = (filename: string) => {
        if (app.isPackaged) {
          const p = join(process.resourcesPath, filename)
          if (existsSync(p)) return p
        } else {
          const devPath = join(__dirname, '../../build', filename)
          if (existsSync(devPath)) return devPath
        }
        return process.execPath
      }

      app.setUserTasks([
        {
          program: process.execPath,
          arguments: '--new-doc',
          title: tm('menuNewDoc') || 'ARA Docs',
          description: 'Create a new Word Document',
          iconPath: resolveResourceIcon('docx.ico'),
          iconIndex: 0,
        },
        {
          program: process.execPath,
          arguments: '--new-sheet',
          title: tm('menuNewSheet') || 'ARA Sheets',
          description: 'Create a new Spreadsheet',
          iconPath: resolveResourceIcon('xlsx.ico'),
          iconIndex: 0,
        },
        {
          program: process.execPath,
          arguments: '--new-slide',
          title: tm('menuNewSlide') || 'ARA Slides',
          description: 'Create a new Presentation',
          iconPath: resolveResourceIcon('pptx.ico'),
          iconIndex: 0,
        },
        {
          program: process.execPath,
          arguments: '--new-markdown',
          title: tm('menuNewMarkdown') || 'ARA Notes',
          description: 'Create a new Note',
          iconPath: resolveResourceIcon('md.ico'),
          iconIndex: 0,
        },
        {
          program: process.execPath,
          arguments: '--new-diagram',
          title: 'ARA Diagrams',
          description: 'Create a new Diagram (Draw.io)',
          iconPath: resolveResourceIcon('diagram.ico'),
          iconIndex: 0,
        },
        {
          program: process.execPath,
          arguments: '--new-diagram-code',
          title: 'ARA Diagram Code',
          description: 'Create a new Mermaid Diagram',
          iconPath: resolveResourceIcon('diagram-code.ico'),
          iconIndex: 0,
        },
      ])
    } catch {}
  }
}
async function installMainProcessProxy(): Promise<void> {
  const proxyUrl = [
    process.env.HTTPS_PROXY,
    process.env.https_proxy,
    process.env.HTTP_PROXY,
    process.env.http_proxy,
    process.env.ALL_PROXY,
    process.env.all_proxy,
  ].find((value) => value && /^https?:\/\//.test(value))
  if (!proxyUrl) return
  try {
    const { ProxyAgent, setGlobalDispatcher } = await import('undici')
    setGlobalDispatcher(new ProxyAgent(proxyUrl))
  } catch {
    /* optional proxy support is unavailable */
  }
}

// ---- lifecycle (the shell is the only owner) ----

let pendingLaunchPath = supportedFileIn(process.argv) ?? unsupportedFileIn(process.argv)

// show() does not un-minimize, and on macOS ⌘W destroys the shell window while the
// app keeps running — either way a file opened from Finder would land out of sight.
function revealShellWindow(): void {
  if (!shellWindow) createShellWindow()
  if (shellWindow?.isMinimized()) shellWindow.restore()
  shellWindow?.show()
  shellWindow?.focus()
}

// On macOS a file opened from Finder is not in argv; it arrives via the open-file event (before ready).
// If another instance already holds the lock, this process exits, and the path must ride along in
// the lock request's additionalData to the surviving instance — so the lock request is deferred
// until ready, after the path is known.
app.on('open-file', async (event, filePath) => {
  event.preventDefault()
  if (!app.isReady()) {
    pendingLaunchPath = filePath
    return
  }
  revealShellWindow()
  if (!(await openDocumentPath(filePath))) tabManager?.openHomeTab()
})

function handleCustomArgv(argv: string[]): boolean {
  if (argv.includes('--new-doc')) {
    newDocTab()
    return true
  }
  if (argv.includes('--new-sheet')) {
    void newSheetTab()
    return true
  }
  if (argv.includes('--new-slide')) {
    newSlideTab()
    return true
  }
  if (argv.includes('--new-markdown')) {
    newMarkdownTab()
    return true
  }
  if (argv.includes('--new-diagram')) {
    newDiagramTab('visual')
    return true
  }
  if (argv.includes('--new-diagram-code')) {
    newDiagramTab('mermaid')
    return true
  }
  return false
}

app.on('second-instance', async (_event, argv, _cwd, additionalData) => {
  revealShellWindow()
  if (handleCustomArgv(argv)) return
  const file =
    supportedFileIn(argv) ??
    unsupportedFileIn(argv) ??
    (additionalData as { launchPath?: string } | null)?.launchPath
  if (!file || !(await openDocumentPath(file))) tabManager?.openHomeTab()
})

installNavigationGuard(app)
installContextMenu(app, () => contextMenuLabels(currentLang()))
registerAiIpc()
// An editor's "Configure AI Provider" card asks for the shell's provider
// dialog: ARAOffice has no accounts, so that one UI is the only setup path.
setAiProviderConfigurator(() => {
  tabManager?.openHomeTab()
  if (shellWindow && !shellWindow.isDestroyed()) {
    try {
      shellWindow.webContents.send('home:configure-ai')
    } catch {
      /* frame disposed */
    }
  }
})
registerProjectIpc()
registerDocsIpc()
registerHomeIpc()
registerTabsIpc()
registerWindowIpc()

// sheets' project:resolveChat goes through the handler registered by docs-main; the sessionId reverse lookup hooks in here
setSessionPathResolver(resolveSheetsSessionPath)
// Prevent Chromium on Windows from discarding DirectComposition swapchains
// or pausing rendering when a WebContentsView sits idle or is marked occluded.
if (process.platform === 'win32') {
  app.commandLine.appendSwitch('disable-direct-composition')
  app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion')
}
app.commandLine.appendSwitch('disable-backgrounding-occluded-windows')

// RCA aids for native renderer deaths (e.g. PDF tab, exitCode -36861 = sandbox-fatal):
// ARAOFFICE_CRASH_LOG=1 captures Chromium-level crash reasons (sandbox failures,
// OOM terminations) into chromium-debug.log inside the profile's user data dir.
// ARAOFFICE_DISABLE_GPU=1 rules GPU-driver involvement in or out by comparison.
if (process.env.ARAOFFICE_CRASH_LOG) {
  const logFile = join(app.getPath('userData'), 'chromium-debug.log')
  app.commandLine.appendSwitch('enable-logging', 'file')
  app.commandLine.appendSwitch('log-file', logFile)
  app.commandLine.appendSwitch('v', '1')
  console.log(`[crash-rca] Chromium logging active → ${logFile}`)
}
if (process.env.ARAOFFICE_DISABLE_GPU) {
  app.disableHardwareAcceleration()
}

app.whenReady().then(async () => {
  const hasLock = app.requestSingleInstanceLock(
    pendingLaunchPath ? { launchPath: pendingLaunchPath } : {},
  )
  if (!hasLock) {
    app.quit()
    return
  }

  if (process.platform === 'win32') {
    app.setAppUserModelId('com.araoffice.app')
  }
  void installMainProcessProxy()
  app.setAccessibilitySupportEnabled(true)
  await primeUiSettings()
  currentLang()
  nativeTheme.themeSource = currentTheme()
  nativeTheme.on('updated', () => updateShellIcon())
  createShellWindow()
  // deferred to ready: labels need currentLang(), which reads app.getLocale()
  installBackToHomeItems()
  installDockMenu()
  void initAutoUpdater(() => shellWindow, currentUpdateChannel())
  if (handleCustomArgv(process.argv)) {
    // handled custom argv flag
  } else if (!pendingLaunchPath || !(await openDocumentPath(pendingLaunchPath))) {
    tabManager?.openHomeTab()
  }
  pendingLaunchPath = null
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createShellWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

app.on('before-quit', () => {
  // No close prompt may fall through to "Save" during shutdown
  markSheetsShuttingDown()
  stopSheetsSidecar()
})
