import { useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import type { Editor } from '@tiptap/core'
import { useEditorState } from '@tiptap/react'
import { useI18n } from '../i18n/locale'
import type { StringKey } from '../i18n/locale'
import { liftFromList } from '../editor/slashCommand'
import {
  AssistantMark,
  IconBlockquote,
  IconBullets,
  IconCodeBlock,
  IconExport,
  IconHighlight,
  IconHr,
  IconIndentDec,
  IconIndentInc,
  IconInlineCode,
  IconLink,
  IconNumbered,
  IconPicture,
  IconPrint,
  IconProperties,
  IconRedo,
  IconTable,
  IconTaskList,
  IconUndo,
} from './icons'

interface Props {
  editor: Editor | null
  disabled: boolean
  imageEnabled: boolean
  onInsertImage: () => void
  frontmatterOpen: boolean
  onToggleFrontmatter: () => void
  aiOpen: boolean
  onToggleAi: () => void
  onAiPreset?: (instruction: string) => void
  onExport?: (format: 'pdf' | 'docx' | 'docs') => void
  onPrint?: () => void
}

type BlockStyle =
  | 'paragraph'
  | 'h1'
  | 'h2'
  | 'h3'
  | 'h4'
  | 'h5'
  | 'h6'
  | 'quote'
  | 'codeBlock'
  | 'callout'
  | 'toggle'

const STYLE_LABEL: Record<BlockStyle, StringKey> = {
  paragraph: 'styleParagraph',
  h1: 'styleH1',
  h2: 'styleH2',
  h3: 'styleH3',
  h4: 'styleH4',
  h5: 'styleH5',
  h6: 'styleH6',
  quote: 'styleQuote',
  codeBlock: 'styleCodeBlock',
  callout: 'insertCallout',
  toggle: 'insertToggle',
}

function applyBlockStyle(editor: Editor, style: BlockStyle): void {
  // block-type conversions are illegal inside list items — leave the list first
  liftFromList(editor)
  const chain = editor.chain().focus()
  // unwrap wrappers the target style does not keep
  if (style !== 'callout' && editor.isActive('callout')) chain.unsetCallout()
  if (style !== 'toggle' && editor.isActive('toggle')) chain.unsetToggle()
  switch (style) {
    case 'paragraph':
      chain.setParagraph().run()
      break
    case 'quote':
      chain.setParagraph().setBlockquote().run()
      break
    case 'codeBlock':
      chain.setCodeBlock().run()
      break
    case 'callout':
      chain.setCallout().run()
      break
    case 'toggle':
      chain.setToggle().run()
      break
    default:
      chain.setHeading({ level: Number(style.slice(1)) as 1 | 2 | 3 | 4 | 5 | 6 }).run()
  }
}

function IconBtn({
  title,
  active,
  disabled,
  onClick,
  children,
}: {
  title: string
  active?: boolean
  disabled?: boolean
  onClick?: () => void
  children: ReactNode
}) {
  return (
    <button
      type="button"
      className={`rb-btn${active ? ' active' : ''}`}
      title={title}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  )
}

function RibbonGroup({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="ribbon-group">
      <div className="ribbon-group-items">{children}</div>
      <span className="rb-group-caption">{label}</span>
    </div>
  )
}

export function Ribbon({
  editor,
  disabled,
  imageEnabled,
  onInsertImage,
  frontmatterOpen,
  onToggleFrontmatter,
  aiOpen,
  onToggleAi,
  onAiPreset,
  onExport,
  onPrint,
}: Props) {
  const { t } = useI18n()
  const [linkOpen, setLinkOpen] = useState(false)
  const [linkUrl, setLinkUrl] = useState('')
  const [exportOpen, setExportOpen] = useState(false)
  const [aiQuickOpen, setAiQuickOpen] = useState(false)
  const linkInputRef = useRef<HTMLInputElement>(null)
  const aiQuickRef = useRef<HTMLSpanElement>(null)

  const state = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      if (!e) return null
      const style: BlockStyle = e.isActive('callout')
        ? 'callout'
        : e.isActive('toggle')
          ? 'toggle'
          : e.isActive('codeBlock')
            ? 'codeBlock'
            : e.isActive('blockquote')
              ? 'quote'
              : e.isActive('heading')
                ? (`h${e.getAttributes('heading').level}` as BlockStyle)
                : 'paragraph'
      return {
        style,
        empty: e.isEmpty,
        bold: e.isActive('bold'),
        italic: e.isActive('italic'),
        strike: e.isActive('strike'),
        underline: e.isActive('underline'),
        highlight: e.isActive('highlight'),
        code: e.isActive('code'),
        link: e.isActive('link'),
        bullet: e.isActive('bulletList'),
        ordered: e.isActive('orderedList'),
        task: e.isActive('taskList'),
        inList: e.isActive('listItem') || e.isActive('taskItem'),
        canUndo: e.can().undo(),
        canRedo: e.can().redo(),
      }
    },
  })

  useEffect(() => {
    if (linkOpen) linkInputRef.current?.focus()
  }, [linkOpen])

  useEffect(() => {
    const closeAiQuickActions = (event: MouseEvent) => {
      if (aiQuickRef.current && !aiQuickRef.current.contains(event.target as Node)) {
        setAiQuickOpen(false)
      }
    }
    window.addEventListener('mousedown', closeAiQuickActions)
    return () => window.removeEventListener('mousedown', closeAiQuickActions)
  }, [])

  const off = disabled || !editor || !state

  const openLink = () => {
    if (!editor) return
    setLinkUrl(String(editor.getAttributes('link').href ?? ''))
    setLinkOpen((v) => !v)
  }

  const applyLink = () => {
    if (!editor) return
    const url = linkUrl.trim()
    const chain = editor.chain().focus().extendMarkRange('link')
    if (url) chain.setLink({ href: url }).run()
    else chain.unsetLink().run()
    setLinkOpen(false)
  }

  const listItemName = state?.task ? 'taskItem' : 'listItem'
  const ICON = 16

  return (
    <div className="ribbon" aria-label="Markdown editing tools">
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            type="button"
            className={`rb-big ai-entry${aiOpen ? ' active' : ''}`}
            title={t('aiOpenAssistant')}
            disabled={disabled}
            onMouseDown={(e) => e.preventDefault()}
            onClick={onToggleAi}
          >
            <span className="rb-big-icon">
              <AssistantMark size={26} />
            </span>
            <span>ARA Assistant</span>
          </button>
          <span className="rb-link-anchor" ref={aiQuickRef}>
            <button
              type="button"
              className={`rb-big ai-entry${aiQuickOpen ? ' active' : ''}`}
              title="AI quick actions"
              aria-haspopup="menu"
              aria-expanded={aiQuickOpen}
              disabled={disabled}
              onClick={() => setAiQuickOpen((open) => !open)}
              onKeyDown={(event) => {
                if (event.key === 'Escape') setAiQuickOpen(false)
              }}
            >
              <span className="rb-big-icon">
                <IconProperties size={22} />
              </span>
              <span>Quick</span>
            </button>
            {aiQuickOpen && (
              <div className="rb-export-menu rb-ai-quick-menu" role="menu">
                <button
                  type="button"
                  className="rb-export-item"
                  role="menuitem"
                  onClick={() => {
                    setAiQuickOpen(false)
                    onAiPreset?.(
                      'Draft a comprehensive, structured technical specification or PRD for the current topic with overview, architecture, and requirements',
                    )
                  }}
                >
                  Draft specification
                </button>
                <button
                  type="button"
                  className="rb-export-item"
                  role="menuitem"
                  onClick={() => {
                    setAiQuickOpen(false)
                    onAiPreset?.(
                      'Generate an actionable GitHub-flavored task checklist (- [ ]) for the active section',
                    )
                  }}
                >
                  Generate tasks
                </button>
                <button
                  type="button"
                  className="rb-export-item"
                  role="menuitem"
                  onClick={() => {
                    setAiQuickOpen(false)
                    onAiPreset?.(
                      'Insert a helpful callout admonition (:::callout {type="tip"}) highlighting key technical notes',
                    )
                  }}
                >
                  Insert callout
                </button>
                <button
                  type="button"
                  className="rb-export-item"
                  role="menuitem"
                  onClick={() => {
                    setAiQuickOpen(false)
                    onAiPreset?.(
                      'Proofread and polish the technical terminology, grammar, and flow in this document',
                    )
                  }}
                >
                  Polish prose
                </button>
              </div>
            )}
          </span>
        </div>
        <span className="rb-group-caption">ARA Assistant</span>
      </div>

      <RibbonGroup label="History">
        <IconBtn
          title={t('undo')}
          disabled={off || !state?.canUndo}
          onClick={() => editor?.chain().focus().undo().run()}
        >
          <IconUndo size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('redo')}
          disabled={off || !state?.canRedo}
          onClick={() => editor?.chain().focus().redo().run()}
        >
          <IconRedo size={ICON} />
        </IconBtn>
      </RibbonGroup>

      <RibbonGroup label="Block">
        <select
          className="rb-style"
          value={state?.style ?? 'paragraph'}
          disabled={off}
          onChange={(e) => editor && applyBlockStyle(editor, e.target.value as BlockStyle)}
        >
          {(Object.keys(STYLE_LABEL) as BlockStyle[]).map((s) => (
            <option key={s} value={s}>
              {t(STYLE_LABEL[s])}
            </option>
          ))}
        </select>
      </RibbonGroup>

      <RibbonGroup label="Font">
        <IconBtn
          title={t('bold')}
          active={state?.bold}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleBold().run()}
        >
          <span className="rb-glyph" style={{ fontWeight: 800 }}>
            B
          </span>
        </IconBtn>
        <IconBtn
          title={t('italic')}
          active={state?.italic}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleItalic().run()}
        >
          <span className="rb-glyph" style={{ fontStyle: 'italic' }}>
            I
          </span>
        </IconBtn>
        <IconBtn
          title={t('underline')}
          active={state?.underline}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleUnderline().run()}
        >
          <span className="rb-glyph" style={{ textDecoration: 'underline' }}>
            U
          </span>
        </IconBtn>
        <IconBtn
          title={t('strike')}
          active={state?.strike}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleStrike().run()}
        >
          <span className="rb-glyph" style={{ textDecoration: 'line-through' }}>
            S
          </span>
        </IconBtn>
        <IconBtn
          title={t('highlight')}
          active={state?.highlight}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleHighlight().run()}
        >
          <IconHighlight size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('inlineCode')}
          active={state?.code}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleCode().run()}
        >
          <IconInlineCode size={ICON} />
        </IconBtn>
        <span className="rb-link-anchor">
          <IconBtn title={t('link')} active={state?.link} disabled={off} onClick={openLink}>
            <IconLink size={ICON} />
          </IconBtn>
          {linkOpen && (
            <span className="rb-link-pop" onMouseDown={(e) => e.stopPropagation()}>
              <input
                ref={linkInputRef}
                value={linkUrl}
                placeholder={t('linkPlaceholder')}
                onChange={(e) => setLinkUrl(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') applyLink()
                  if (e.key === 'Escape') setLinkOpen(false)
                }}
              />
              <button type="button" onClick={applyLink}>
                {t('linkApply')}
              </button>
            </span>
          )}
        </span>
      </RibbonGroup>

      <RibbonGroup label="Lists">
        <IconBtn
          title={t('bulletList')}
          active={state?.bullet}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleBulletList().run()}
        >
          <IconBullets size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('orderedList')}
          active={state?.ordered}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleOrderedList().run()}
        >
          <IconNumbered size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('taskList')}
          active={state?.task}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleTaskList().run()}
        >
          <IconTaskList size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('outdent')}
          disabled={off || !state?.inList}
          onClick={() => editor?.chain().focus().liftListItem(listItemName).run()}
        >
          <IconIndentDec size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('indent')}
          disabled={off || !state?.inList}
          onClick={() => editor?.chain().focus().sinkListItem(listItemName).run()}
        >
          <IconIndentInc size={ICON} />
        </IconBtn>
      </RibbonGroup>

      <RibbonGroup label="Insert">
        <IconBtn
          title={t('styleQuote')}
          active={state?.style === 'quote'}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleBlockquote().run()}
        >
          <IconBlockquote size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('styleCodeBlock')}
          active={state?.style === 'codeBlock'}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleCodeBlock().run()}
        >
          <IconCodeBlock size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('insertTable')}
          disabled={off}
          onClick={() =>
            editor?.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()
          }
        >
          <IconTable size={ICON} />
        </IconBtn>
        <IconBtn title={t('insertImage')} disabled={off || !imageEnabled} onClick={onInsertImage}>
          <IconPicture size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('insertHr')}
          disabled={off}
          onClick={() => editor?.chain().focus().setHorizontalRule().run()}
        >
          <IconHr size={ICON} />
        </IconBtn>
      </RibbonGroup>

      <div className="rb-spacer" />

      <RibbonGroup label="File">
        <IconBtn title={`${t('print')} (⌘P)`} disabled={off} onClick={() => onPrint?.()}>
          <IconPrint size={ICON} />
        </IconBtn>
        <span className="rb-link-anchor">
          <IconBtn
            title={t('export')}
            active={exportOpen}
            disabled={off}
            onClick={() => setExportOpen((v) => !v)}
          >
            <IconExport size={ICON} />
          </IconBtn>
          {exportOpen && (
            <div className="rb-export-menu" onMouseDown={(e) => e.stopPropagation()}>
              <button
                type="button"
                className="rb-export-item"
                onClick={() => {
                  setExportOpen(false)
                  onExport?.('pdf')
                }}
              >
                {t('exportPdf')}
              </button>
              <button
                type="button"
                className="rb-export-item"
                onClick={() => {
                  setExportOpen(false)
                  onExport?.('docx')
                }}
              >
                {t('exportDocx')}
              </button>
            </div>
          )}
        </span>
        <IconBtn
          title={t('fmProperties')}
          active={frontmatterOpen}
          disabled={disabled}
          onClick={onToggleFrontmatter}
        >
          <IconProperties size={ICON} />
        </IconBtn>
      </RibbonGroup>
    </div>
  )
}
