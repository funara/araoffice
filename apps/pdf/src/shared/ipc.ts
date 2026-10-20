import type { Lang } from '@araoffice/i18n'
import type { AiSettings, AiStreamChunk, AiStreamRequest } from '@araoffice/ai-provider'

/**
 * Accessibility report, drawn from the file only: automated verdicts plus honest "manual" marks
 * where the answer is not in the file (contrast, reading order, screen flicker). Report only.
 */
export type A11yRuleId =
  'title' | 'language' | 'tagged' | 'altFigures' | 'bookmarksLong' | 'headings' | 'tables' | 'forms'

export interface A11yRuleResult {
  rule: A11yRuleId
  verdict: 'pass' | 'fail' | 'manual'
  /** The renderer localizes when the verdict needs a number on screen */
  count?: number
}

export interface A11yReport {
  pageCount: number
  rules: A11yRuleResult[]
}

export const PDF_CHANNELS = {
  consumePending: 'pdf:consume-pending',
  readFile: 'pdf:read-file',
  save: 'pdf:save',
  extractPages: 'pdf:extract-pages',
  insertPdf: 'pdf:insert-pdf',
  exportImages: 'pdf:export-images',
  exportText: 'pdf:export-text',
  exportEmbeddedImages: 'pdf:export-embedded-images',
  setOutline: 'pdf:set-outline',
  a11yCheck: 'pdf:a11y-check',
  a11yFix: 'pdf:a11y-fix',
  splitPdf: 'pdf:split-pdf',
  combinePdfs: 'pdf:combine-pdfs',
  sanitizePdf: 'pdf:sanitize-pdf',
  sanitizePreview: 'pdf:sanitize-preview',
  deleteAnnotations: 'pdf:delete-annotations',
  dirtyChanged: 'pdf:dirty-changed',
  closeSaveRequest: 'pdf:close-save-request',
  closeSaveResult: 'pdf:close-save-result',
  getLanguage: 'app:get-language',
  languageChanged: 'app:language-changed',
} as const

/**
 * What sanitize may remove, one flag per category. Each is safe to rewrite with pdf-lib; the
 * categories that need content-stream surgery (invisible text, hidden layers) are absent
 * deliberately rather than half-implemented.
 */
export interface SanitizeRequest {
  /** /Info fields (blanked) and the XMP /Metadata stream (dropped) */
  metadata: boolean
  /** Catalog /OpenAction, /AA and /Names /JavaScript */
  javascript: boolean
  /** Embedded files (/Names /EmbeddedFiles, /Collection) */
  attachments: boolean
  /** Page /Annots — markup comments, links and their actions; widget annotations stay */
  comments: boolean
  /** Page /Thumb */
  thumbnails: boolean
  /** /PieceInfo, in the catalog and on each page */
  privateData: boolean
}

/** How much of each category a sanitize run removed, for reporting after the fact */
export type SanitizeReport = Record<keyof SanitizeRequest, number>

export interface SanitizePdfRequest extends SanitizeRequest {
  /** The path rewritten in place; the same grant check every other write goes through */
  path: string
}

export type SanitizeResult = { ok: true; report: SanitizeReport } | { ok: false; error: string }

/** Read-only census for the sanitize dialog: runs the removal walk without applying it */
export interface SanitizePreviewRequest {
  path: string
}

export type SanitizePreviewResult =
  { ok: true; counts: SanitizeReport } | { ok: false; error: string }

/** Identity of one annotation as the renderer saw it; pdf.js ids are synthesized and unstable,
 * so the file-side match keys on the decoded fields instead. Shared here because both sides
 * speak it. */
export interface AnnotMatch {
  /** 0-based page index in the saved file */
  pageIndex: number
  /** decoded /Subtype, e.g. Text, Highlight, FreeText */
  subtype: string
  /** /Rect as [x1, y1, x2, y2] in the annotation's user space */
  rect: number[]
  /** decoded /Contents, normalized: absent and empty are the same */
  contents: string
  /** decoded /NM when present */
  name: string
}

/** Comments-panel deletion: refusals are whole-batch so the file never half-matches the panel */
export interface DeleteAnnotationsRequest {
  path: string
  targets: AnnotMatch[]
}

export type DeleteAnnotationsResult = { ok: true } | { ok: false; error: string }

/** Report-only accessibility scan; no file writes, same grant model as the other reads */
export interface A11yCheckRequest {
  path: string
}

export type A11yCheckResult = { ok: true; report: A11yReport } | { ok: false; error: string }

/** Fill the report's two metadata gaps: /Info title (from the file name) and catalogue /Lang */
export interface A11yFixRequest {
  /** The granted path the fix rewrites in place */
  path: string
  /** BCP-47-ish tag written when the catalogue has no /Lang */
  lang: string
}

export type A11yFixResult = { ok: true; fixed: number } | { ok: false; error: string }

export type MarkupType = 'highlight' | 'underline' | 'strikeout' | 'squiggly'

/** Minimum coordinate count in DrawingInput.points: 2 points open, 3 points closed.
    Shared so the renderer stops collecting and the writer refuses malformed input on the same rule. */
export const MIN_POLY_POINTS: Record<'polyline' | 'polygon', number> = {
  polyline: 4,
  polygon: 6,
}

/** A text markup to write; quads are 4-point groups in PDF coords (y up) [x1,yTop,x2,yTop,x1,yBottom,x2,yBottom] */
export interface MarkupInput {
  pageIndex: number
  type: MarkupType
  /** rgb normalized to 0-1 */
  color: [number, number, number]
  quads: number[][]
}

/** Drawing annotations (all coords in PDF user space, y up).
    One union member per kind; a union-literal kind would break TS narrowing. */
interface DrawBase {
  pageIndex: number
  color: [number, number, number]
  width: number
}

export type DrawingInput =
  | (DrawBase & {
      kind: 'ink'
      /** Each stroke as [x1,y1,x2,y2,...] */
      paths: number[][]
    })
  | (DrawBase & { kind: 'rect'; rect: [number, number, number, number] })
  | (DrawBase & { kind: 'ellipse'; rect: [number, number, number, number] })
  | (DrawBase & { kind: 'line'; from: [number, number]; to: [number, number] })
  | (DrawBase & {
      /** Open connected lines; flat [x1,y1,x2,y2,...], at least 2 points */
      kind: 'polyline'
      points: number[]
    })
  | (DrawBase & {
      /** Closed polygon; flat [x1,y1,x2,y2,...], at least 3 points */
      kind: 'polygon'
      points: number[]
    })
  | {
      /** Insertion-point caret, as a replace-text proposal uses */
      kind: 'caret'
      pageIndex: number
      color: [number, number, number]
      /** Insertion point at the base of the caret, PDF user space y up */
      at: [number, number]
      /** Line height the caret spans */
      size: number
      contents: string
    }
  | (DrawBase & { kind: 'arrow'; from: [number, number]; to: [number, number] })
  | {
      /** Image signature/stamp placed by the user; written as a Stamp annotation */
      kind: 'image'
      pageIndex: number
      /** base64 PNG, without the data: prefix */
      image: string
      /** PDF user space [x1,y1,x2,y2] */
      rect: [number, number, number, number]
    }
  | {
      /** Typewriter text; written as a FreeText annotation whose appearance is the rasterized
          glyphs, so CJK works despite pdf-lib's WinAnsi-only built-in fonts */
      kind: 'freetext'
      pageIndex: number
      /** base64 PNG, without the data: prefix */
      image: string
      /** PDF user space [x1,y1,x2,y2] */
      rect: [number, number, number, number]
      /** The typed text, kept in /Contents so the comment stays listable and searchable */
      contents: string
    }
  | {
      kind: 'note'
      pageIndex: number
      color: [number, number, number]
      at: [number, number]
      contents: string
    }

/**
 * Stamp layer (watermark/header/footer/page numbers all go through it).
 * The renderer rasterizes the bitmap via canvas (with rotation and fonts, bypassing
 * pdf-lib's lack of CJK support); the main process only embeds and positions it.
 */
export interface StampInput {
  pageIndex: number
  /** base64 PNG, without the data: prefix */
  image: string
  /** PDF user space [x1,y1,x2,y2] */
  rect: [number, number, number, number]
  opacity?: number
}

/** Document info; an empty string clears the field */
export interface MetadataInput {
  title?: string
  author?: string
  subject?: string
  keywords?: string
}

export interface FormValueInput {
  name: string
  kind: 'text' | 'checkbox' | 'radio' | 'choice'
  /** For radio: selected exportValue; for choice: selected option; empty string clears selection */
  value?: string
  checked?: boolean
}

export interface SavePdfRequest {
  path: string
  /**
   * Save As destination. When set, `path` is only read (source bytes) and the edited PDF
   * is written to this path instead — the original file must never be mutated.
   * Must match the target granted to the view by the main process (save dialog pick).
   */
  targetPath?: string
  markups: MarkupInput[]
  drawings: DrawingInput[]
  formValues: FormValueInput[]
  stamps: StampInput[]
  /** Page rotation deltas (original page index → multiple of 90 clockwise) */
  rotations?: { pageIndex: number; delta: number }[]
  /** Pages to delete (original page indices) */
  deletedPages?: number[]
  /** New page order (array of original page indices, excluding deleted); omitted if unreordered */
  pageOrder?: number[]
  metadata?: MetadataInput
}

export type SavePdfResult = { ok: true } | { ok: false; error: string }

/** Extract pages into a new PDF: main process shows a save dialog; cancel returns canceled */
export interface ExtractPagesRequest {
  path: string
  /** Original page indices */
  pages: number[]
  suggestedName: string
}

export type ExtractPagesResult =
  { ok: true; savedPath: string } | { ok: true; canceled: true } | { ok: false; error: string }

/** Insert (merge) another PDF after a page of the current file: main process shows a picker and writes back immediately */
export interface InsertPdfRequest {
  path: string
  /** Insert after this original page index; -1 means front of the document */
  afterPageIndex: number
}

export type InsertPdfResult =
  { ok: true; insertedCount: number } | { ok: true; canceled: true } | { ok: false; error: string }

/** Export pages as PNG: renderer rasterizes the bitmaps, main process shows a dialog and writes to disk */
export interface ExportImagesRequest {
  /** base64 PNGs (without the data: prefix), in page order */
  images: string[]
  /** 1-based page numbers, same length as images, used for file names */
  pageNumbers: number[]
  baseName: string
}

export type ExportImagesResult =
  | { ok: true; savedDir: string; count: number }
  | { ok: true; canceled: true }
  | { ok: false; error: string }

/** Export the document's text: main process shows a save dialog and writes it */
export interface ExportTextRequest {
  /** Plain text, page-marker separated */
  text: string
  suggestedName: string
}

export type ExportTextResult =
  { ok: true; savedPath: string } | { ok: true; canceled: true } | { ok: false; error: string }

/** Split into one file per range; main process shows a directory picker and writes them there */
export interface SplitPdfRequest {
  path: string
  /** One entry per output file, as original 0-based page indices */
  ranges: number[][]
  /** File name stem; each part gets a -01/-02 suffix */
  baseName: string
}

export type SplitPdfResult =
  | { ok: true; savedDir: string; count: number }
  | { ok: true; canceled: true }
  | { ok: false; error: string }

/** Merge several PDFs; the main process picks the files (multi-select), merges and saves */
export interface CombinePdfsRequest {
  /** Add one bookmark per source file, titled with its name */
  bookmarkPerFile: boolean
}

export type CombinePdfsResult =
  | { ok: true; savedPath: string; fileCount: number }
  | { ok: true; canceled: true }
  | { ok: false; error: string }

/** Export the file's embedded images: main process shows a folder picker and writes them there */
export interface ExportEmbeddedImagesRequest {
  path: string
}

export type ExportEmbeddedImagesResult =
  | { ok: true; savedDir: string; count: number; skipped: number }
  | { ok: true; canceled: true }
  | { ok: false; error: string }

/** Replace the document's outline with this tree; empty nodes clears it entirely */
export interface OutlineEditNode {
  title: string
  /** 0-based page index the bookmark jumps to; ignored when uri is set */
  pageIndex: number
  /** Kept for web and mailto bookmarks, which get a /URI action instead of a destination */
  uri?: string
  children?: OutlineEditNode[]
}

export interface SetOutlineRequest {
  path: string
  nodes: OutlineEditNode[]
}

export type SetOutlineResult = { ok: true } | { ok: false; error: string }

/** AI channels are app-wide shared ipcMain handlers (shell registers via docs-main registerAiIpc); pass-through only */
export const AI_CHANNELS = {
  getSettings: 'ai:get-settings',
  stream: 'ai:stream',
  streamChunk: 'ai:stream-chunk',
  streamCancel: 'ai:stream-cancel',
} as const

/** API exposed by preload to the renderer (window.pdfApi) */
export interface PdfApi {
  /** Take the pdf path pending for this view (queued at tab creation); null if none */
  consumePending(): Promise<string | null>
  /** Read pdf bytes. Only paths granted to this view are allowed */
  readFile(path: string): Promise<ArrayBuffer>
  /** Write markups/form values/page ops back to the original file (pdf-lib, content streams untouched); path grants same as readFile. With targetPath set (Save As), the original is only read and the result goes to targetPath */
  save(request: SavePdfRequest): Promise<SavePdfResult>
  extractPages(request: ExtractPagesRequest): Promise<ExtractPagesResult>
  insertPdf(request: InsertPdfRequest): Promise<InsertPdfResult>
  exportImages(request: ExportImagesRequest): Promise<ExportImagesResult>
  exportText(request: ExportTextRequest): Promise<ExportTextResult>
  exportEmbeddedImages(request: ExportEmbeddedImagesRequest): Promise<ExportEmbeddedImagesResult>
  setOutline(request: SetOutlineRequest): Promise<SetOutlineResult>
  a11yCheck(request: A11yCheckRequest): Promise<A11yCheckResult>
  a11yFix(request: A11yFixRequest): Promise<A11yFixResult>
  splitPdf(request: SplitPdfRequest): Promise<SplitPdfResult>
  combinePdfs(request: CombinePdfsRequest): Promise<CombinePdfsResult>
  sanitizePdf(request: SanitizePdfRequest): Promise<SanitizeResult>
  sanitizePreview(request: SanitizePreviewRequest): Promise<SanitizePreviewResult>
  deleteAnnotations(request: DeleteAnnotationsRequest): Promise<DeleteAnnotationsResult>
  /** Mirror unsaved-changes state to the main process; drives the save prompt before closing a tab/window */
  setDirty(dirty: boolean): void
  /** Main process picked "Save" in the close prompt → renderer saves and replies via sendCloseSaveResult */
  onCloseSaveRequest(handler: () => void): () => void
  sendCloseSaveResult(ok: boolean): void
  getLanguage(): Promise<Lang>
  onLanguageChanged(handler: (lang: Lang) => void): () => void
  getAiSettings(): Promise<AiSettings>
  /** open the shell's AI provider settings (no accounts: configuring a provider is the only setup) */
  configureAi: () => Promise<void>
  onAiSettingsChanged?(handler: (settings: AiSettings) => void): () => void
  aiStream(request: AiStreamRequest): Promise<void>
  aiStreamCancel(requestId: string): Promise<void>
  onAiStream(handler: (chunk: AiStreamChunk) => void): () => void
}
