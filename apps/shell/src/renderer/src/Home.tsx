import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type {
  HomeApi,
  ProjectHomeApi,
  ProjectSummaryEntry,
  RecentEntry,
  DiagramMode,
} from '../../shared/home-api'
import type { Lang } from '@araoffice/i18n'
import type { AiProviderId, AiSettings } from '@araoffice/ai-provider'
import {
  AI_PROVIDERS,
  CUSTOM_ENDPOINT_PRESETS,
  DEFAULT_GOOGLE_MODEL,
} from '@araoffice/ai-provider/providers'
import { fileCountKey, visiblePageCount } from './counts'
import { LANGUAGE_OPTIONS, useI18n } from './locale'
import type { I18n, StringKey } from './locale'
import { SocialLinks } from './ShellPanels'
import { ChangelogModal } from './ChangelogModal'
import { UpdateRestartModal } from './UpdateRestartModal'
import { DonationModal } from './DonationModal'
import { FeedbackModal } from './FeedbackModal'
import { ShortcutsModal } from './ShortcutsModal'
import type { AppUpdateState } from '../../shared/update-api'
declare global {
  interface Window {
    aiOffice: HomeApi
    aiOfficeProject?: ProjectHomeApi
  }
}

/** page size of the home list; scrolling to the bottom auto-loads the next page */
const PAGE_SIZE = 50

/** Gemini model list comes from the shared provider metadata so the UI never drifts. */
const GEMINI_MODELS = AI_PROVIDERS.find((provider) => provider.id === 'google')?.models ?? []

/** Local OpenAI-compatible servers (Ollama, LM Studio) usually need no API key. */
function isLocalhostEndpoint(baseUrl: string): boolean {
  return /^https?:\/\/(localhost|127\.0\.0\.1|\[::1\])(:|\/|$)/i.test(baseUrl.trim())
}

const FILE_BADGES: Record<string, { label: string; fill: string; fontSize: number }> = {
  docx: { label: 'W', fill: '#3276CD', fontSize: 17 },
  doc: { label: 'W', fill: '#3276CD', fontSize: 17 },
  xlsx: { label: 'X', fill: '#4FA16B', fontSize: 17 },
  xls: { label: 'X', fill: '#4FA16B', fontSize: 17 },
  csv: { label: 'C', fill: '#10b981', fontSize: 17 },
  pptx: { label: 'S', fill: '#e0442e', fontSize: 17 },
  ppt: { label: 'S', fill: '#e0442e', fontSize: 17 },
  pdf: { label: 'P', fill: '#dc2626', fontSize: 17 },
  md: { label: 'N', fill: '#7C3AED', fontSize: 17 },
  markdown: { label: 'N', fill: '#7C3AED', fontSize: 17 },
  mmd: { label: 'D', fill: '#ec4899', fontSize: 17 },
  mermaid: { label: 'D', fill: '#ec4899', fontSize: 17 },
  drawio: { label: 'D', fill: '#d97706', fontSize: 17 },
  sqlite: { label: 'Q', fill: '#0891b2', fontSize: 17 },
  query: { label: 'Q', fill: '#0891b2', fontSize: 17 },
  db: { label: 'Q', fill: '#0891b2', fontSize: 17 },
  sav: { label: 'S', fill: '#65a30d', fontSize: 17 },
  stats: { label: 'S', fill: '#65a30d', fontSize: 17 },
  spss: { label: 'S', fill: '#65a30d', fontSize: 17 },
}

export function FileBadge({ ext, size }: { ext: string; size: number }) {
  const badge = FILE_BADGES[ext] ?? {
    label: ext ? ext[0].toUpperCase() : '?',
    fill: '#98a2b3',
    fontSize: 17,
  }
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" aria-hidden="true">
      <rect width="32" height="32" rx="8" fill={badge.fill} />
      <text
        x="16"
        y="16"
        textAnchor="middle"
        dominantBaseline="central"
        fill="#fff"
        fontSize={badge.fontSize}
        fontWeight="700"
        fontFamily="'Segoe UI', Arial, sans-serif"
      >
        {badge.label}
      </text>
    </svg>
  )
}

function formatModified(mtimeMs: number, i18n: I18n): string {
  const date = new Date(mtimeMs)
  const now = new Date()
  const startOfDay = (d: Date) => new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime()
  const days = Math.round((startOfDay(now) - startOfDay(date)) / 86400000)
  if (days <= 0) {
    return `${i18n.t('today')} · ${date.toLocaleTimeString(i18n.dateLocale, { hour: '2-digit', minute: '2-digit' })}`
  }
  if (days === 1) return i18n.t('yesterday')
  return date.toLocaleDateString(i18n.dateLocale, { month: 'short', day: 'numeric' })
}

function formatSize(bytes: number): string {
  if (bytes >= 1048576) return `${(bytes / 1048576).toFixed(1)} MB`
  return `${Math.max(1, Math.round(bytes / 1024))} KB`
}

function parentDir(path: string): string {
  const parts = path.split(/[\\/]/).filter(Boolean)
  return parts[parts.length - 2] ?? ''
}

function fileName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path
}

function baseName(entry: RecentEntry): string {
  return entry.ext ? entry.name.slice(0, -(entry.ext.length + 1)) : entry.name
}

// ── Project hooks ─────────────────────────────────────────

/** whether we are inside the shell (aiOfficeProject API available) */
function hasProjectApi(): boolean {
  return typeof window.aiOfficeProject !== 'undefined'
}

const FILTERS: { key: string; label: StringKey }[] = [
  { key: 'all', label: 'filterAll' },
  { key: 'docx', label: 'filterDocs' },
  { key: 'xlsx', label: 'filterSheets' },
  { key: 'pptx', label: 'filterSlides' },
  { key: 'pdf', label: 'filterPdf' },
  { key: 'md', label: 'filterMd' },
  { key: 'diagram', label: 'filterDiagram' },
]

// ── Project sidebar component ────────────────────────────

function LanguageMenu({
  lang,
  setLang,
  label,
}: {
  lang: Lang
  setLang: (lang: Lang) => void
  label: string
}) {
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement | null>(null)
  const current = LANGUAGE_OPTIONS.find((option) => option.value === lang) ?? LANGUAGE_OPTIONS[0]
  useEffect(() => {
    if (!open) return
    const onPointerDown = (event: PointerEvent) => {
      if (wrapRef.current && !wrapRef.current.contains(event.target as Node)) setOpen(false)
    }
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setOpen(false)
    }
    document.addEventListener('pointerdown', onPointerDown)
    document.addEventListener('keydown', onKeyDown)
    return () => {
      document.removeEventListener('pointerdown', onPointerDown)
      document.removeEventListener('keydown', onKeyDown)
    }
  }, [open])
  return (
    <div className="language-selector-select-wrap" ref={wrapRef}>
      <button
        type="button"
        className="language-selector-btn"
        aria-label={label}
        aria-haspopup="listbox"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
      >
        <span className="language-selector-btn-label">{current?.label}</span>
        <svg
          className={`language-selector-chevron${open ? ' open' : ''}`}
          width="10"
          height="10"
          viewBox="0 0 16 16"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
          aria-hidden="true"
        >
          <path d="m4 6 4 4 4-4" />
        </svg>
      </button>
      {open && (
        <ul className="language-selector-menu" role="listbox" aria-label={label}>
          {LANGUAGE_OPTIONS.map((option) => (
            <li key={option.value} role="presentation">
              <button
                type="button"
                role="option"
                aria-selected={option.value === lang}
                className={`language-selector-option${option.value === lang ? ' selected' : ''}`}
                onClick={() => {
                  setLang(option.value)
                  setOpen(false)
                }}
              >
                {option.label}
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}

interface ProjectPanelProps {
  projects: ProjectSummaryEntry[]
  selectedId: string | null
  onSelect: (id: string | null) => void
  onRefresh: () => void
}

function ProjectPanel({ projects, selectedId, onSelect, onRefresh }: ProjectPanelProps) {
  const { t } = useI18n()
  const [creating, setCreating] = useState(false)
  const [newName, setNewName] = useState('')
  // open menu id + fixed-position anchor (viewport coords), so the popup can
  // escape the scrollable project list without the list losing overflow-y
  const [projMenu, setProjMenu] = useState<{ id: string; top: number; right: number } | null>(null)
  const [renaming, setRenaming] = useState<{ id: string; value: string } | null>(null)
  const newInputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    if (creating && newInputRef.current) newInputRef.current.focus()
  }, [creating])

  // close the menu on outside click or any scroll (the fixed-position popup
  // would otherwise detach from its row while the list scrolls)
  useEffect(() => {
    if (!projMenu) return
    const handler = (e: PointerEvent) => {
      const target = e.target as Element | null
      if (!target?.closest?.('.proj-menu-wrap')) setProjMenu(null)
    }
    const close = () => setProjMenu(null)
    window.addEventListener('pointerdown', handler)
    window.addEventListener('scroll', close, true)
    return () => {
      window.removeEventListener('pointerdown', handler)
      window.removeEventListener('scroll', close, true)
    }
  }, [projMenu])

  const commitCreate = async () => {
    const name = newName.trim()
    setCreating(false)
    setNewName('')
    if (!name) return
    await window.aiOfficeProject?.createProject(name)
    onRefresh()
  }

  const commitRename = async () => {
    if (!renaming) return
    const name = renaming.value.trim()
    const id = renaming.id
    setRenaming(null)
    if (!name) return
    await window.aiOfficeProject?.renameProject(id, name)
    onRefresh()
  }

  // in-app confirm dialog (same style as the delete-files modal), not window.confirm
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null)

  const doDelete = (id: string) => {
    setProjMenu(null)
    setConfirmDeleteId(id)
  }

  const confirmDeleteNow = async () => {
    const id = confirmDeleteId
    setConfirmDeleteId(null)
    if (!id) return
    await window.aiOfficeProject?.deleteProject(id)
    if (selectedId === id) onSelect(null)
    onRefresh()
  }

  useEffect(() => {
    if (!confirmDeleteId) return
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setConfirmDeleteId(null)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [confirmDeleteId])

  return (
    <div className="proj-panel">
      <div className="proj-panel-head">
        <span className="proj-panel-title">{t('projects')}</span>
        <button
          className="proj-add-btn"
          title={t('newProject')}
          onClick={() => setCreating(true)}
          aria-label={t('newProject')}
        >
          <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
            <path
              d="M7 1v12M1 7h12"
              stroke="currentColor"
              strokeWidth="1.7"
              strokeLinecap="round"
            />
          </svg>
        </button>
      </div>

      {creating && (
        <div className="proj-new-row">
          <label className="proj-new-label" htmlFor="proj-new-input">
            {t('projectName')}
          </label>
          <input
            id="proj-new-input"
            ref={newInputRef}
            className="proj-rename-input"
            aria-label={t('projectName')}
            placeholder={t('projectName')}
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            onBlur={() => void commitCreate()}
            onKeyDown={(e) => {
              if (e.key === 'Enter') void commitCreate()
              if (e.key === 'Escape') {
                setCreating(false)
                setNewName('')
              }
            }}
          />
        </div>
      )}

      <ul className="proj-list">
        {projects.map((proj) => {
          const isActive = selectedId === proj.id
          const isRenaming = renaming?.id === proj.id
          return (
            <li key={proj.id} className={`proj-item${isActive ? ' active' : ''}`}>
              <div
                className="proj-item-main"
                role="button"
                tabIndex={0}
                onClick={() => onSelect(isActive ? null : proj.id)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') onSelect(isActive ? null : proj.id)
                }}
              >
                <span className="proj-item-icon" aria-hidden="true">
                  <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                    <path
                      d="M1.5 4A1.5 1.5 0 0 1 3 2.5h3.1c.44 0 .85.19 1.13.52L8.4 4.4H13A1.5 1.5 0 0 1 14.5 5.9v5.6A1.5 1.5 0 0 1 13 13H3a1.5 1.5 0 0 1-1.5-1.5V4z"
                      stroke="currentColor"
                      strokeWidth="1.2"
                      strokeLinejoin="round"
                    />
                  </svg>
                </span>
                {isRenaming ? (
                  <input
                    className="proj-rename-input inline"
                    value={renaming.value}
                    autoFocus
                    onFocus={(e) => e.target.select()}
                    onClick={(e) => e.stopPropagation()}
                    onChange={(e) => setRenaming({ id: proj.id, value: e.target.value })}
                    onBlur={() => void commitRename()}
                    onKeyDown={(e) => {
                      e.stopPropagation()
                      if (e.key === 'Enter') void commitRename()
                      if (e.key === 'Escape') setRenaming(null)
                    }}
                  />
                ) : (
                  <span className="proj-item-name">
                    {proj.isDefault ? t('defaultProject') : proj.name}
                  </span>
                )}
                <span className="proj-item-meta">
                  <span className="proj-item-count">{proj.fileCount}</span>
                </span>
              </div>

              {!proj.isDefault && (
                <div className="proj-menu-wrap">
                  <button
                    className="proj-more-btn"
                    aria-label={t('projMoreActions', { name: proj.name })}
                    aria-expanded={projMenu?.id === proj.id}
                    onClick={(e) => {
                      e.stopPropagation()
                      if (projMenu?.id === proj.id) {
                        setProjMenu(null)
                        return
                      }
                      const rect = e.currentTarget.getBoundingClientRect()
                      setProjMenu({
                        id: proj.id,
                        top: rect.bottom + 4,
                        right: window.innerWidth - rect.right,
                      })
                    }}
                  >
                    <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">
                      <circle cx="3.2" cy="8" r="1.35" fill="currentColor" />
                      <circle cx="8" cy="8" r="1.35" fill="currentColor" />
                      <circle cx="12.8" cy="8" r="1.35" fill="currentColor" />
                    </svg>
                  </button>
                  {projMenu?.id === proj.id && (
                    <div
                      className="proj-menu"
                      role="menu"
                      style={{ top: projMenu.top, right: projMenu.right }}
                    >
                      <button
                        role="menuitem"
                        onClick={(e) => {
                          e.stopPropagation()
                          setProjMenu(null)
                          setRenaming({ id: proj.id, value: proj.name })
                        }}
                      >
                        {t('rename')}
                      </button>
                      <div className="row-menu-divider" />
                      <button
                        role="menuitem"
                        className="danger"
                        onClick={(e) => {
                          e.stopPropagation()
                          doDelete(proj.id)
                        }}
                      >
                        {t('deleteProject')}
                      </button>
                    </div>
                  )}
                </div>
              )}
            </li>
          )
        })}
      </ul>

      {confirmDeleteId &&
        (() => {
          // locale string is "title?\nbody" — split it across the dialog
          const [confirmTitle, ...confirmBody] = t('deleteProjectConfirm').split('\n')
          return (
            <div className="modal-overlay" onClick={() => setConfirmDeleteId(null)}>
              <div
                className="modal"
                role="dialog"
                aria-modal="true"
                aria-label={confirmTitle}
                onClick={(event) => event.stopPropagation()}
              >
                <h3>{confirmTitle}</h3>
                <p>{confirmBody.join('\n')}</p>
                <div className="modal-buttons">
                  <button
                    className="btn btn-secondary"
                    autoFocus
                    onClick={() => setConfirmDeleteId(null)}
                  >
                    {t('cancel')}
                  </button>
                  <button className="btn btn-danger" onClick={() => void confirmDeleteNow()}>
                    {t('delete')}
                  </button>
                </div>
              </div>
            </div>
          )
        })()}
    </div>
  )
}

// ── Main component ──────────────────────────────────────

export function Home({
  onOpenFeatures,
  onOpenCommandCenter: _onOpenCommandCenter,
  openAiSettings,
}: {
  onOpenFeatures: () => void
  onOpenCommandCenter?: () => void
  /** open the AI provider dialog on request (an editor's setup card asked for it) */
  openAiSettings?: number
}) {
  const i18n = useI18n()
  const { t, lang, setLang } = i18n
  // ── Paged list state (rows loaded for the current view + filter) ──
  const [entries, setEntries] = useState<RecentEntry[]>([])
  /** total count under the current view + filter (not just the loaded rows) */
  const [listTotal, setListTotal] = useState(0)
  /** sidebar Recent / Starred counts under the active type filter */
  const [navCounts, setNavCounts] = useState({ recent: 0, starred: 0 })
  const [loadingMore, setLoadingMore] = useState(false)
  const [view, setView] = useState<'recent' | 'starred'>('recent')
  const [filter, setFilter] = useState('all')
  const [rowMenu, setRowMenu] = useState<string | null>(null)
  const [selected, setSelected] = useState<ReadonlySet<string>>(new Set())
  const [renaming, setRenaming] = useState<{ path: string; value: string } | null>(null)
  const [confirmDelete, setConfirmDelete] = useState<string[] | null>(null)
  const [isDraggingFile, setIsDraggingFile] = useState(false)

  const [appVersion, setAppVersion] = useState('1.0.0')

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'o') {
        e.preventDefault()
        void window.aiOffice.browse()
      } else if (((e.ctrlKey || e.metaKey) && (e.key === '/' || e.key === '?')) || e.key === 'F1') {
        e.preventDefault()
        setShowShortcuts((prev) => !prev)
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [])

  const [updateState, setUpdateState] = useState<AppUpdateState>({
    status: 'idle',
    currentVersion: '',
  })
  const [showChangelog, setShowChangelog] = useState(false)
  const [showUpdateRestart, setShowUpdateRestart] = useState(false)
  const [showDonation, setShowDonation] = useState(false)
  const [showFeedback, setShowFeedback] = useState(false)
  const [showShortcuts, setShowShortcuts] = useState(false)
  const [diagramModeModal, setDiagramModeModal] = useState(false)
  const [websiteUrl, setWebsiteUrl] = useState<string | null>(null)
  const [vendor, setVendor] = useState('ARAOffice')

  // ── AI provider switch (hero button): load, edit locally, save once ──
  const [activeAiProvider, setActiveAiProvider] = useState<AiProviderId>('google')
  const [activeModel, setActiveModel] = useState('')
  const [isFetchingActiveModel, setIsFetchingActiveModel] = useState(false)
  const [isAiConnected, setIsAiConnected] = useState(false)

  const [aiMenuOpen, setAiMenuOpen] = useState(false)
  const [aiChoice, setAiChoice] = useState<AiProviderId>('google')
  const [googleModel, setGoogleModel] = useState(DEFAULT_GOOGLE_MODEL)
  const [googleKey, setGoogleKey] = useState('')
  const [hasSavedGoogleKey, setHasSavedGoogleKey] = useState(false)
  const [aiEndpoint, setAiEndpoint] = useState('')
  const [aiModelId, setAiModelId] = useState('')
  const [aiKey, setAiKey] = useState('')
  const [hasSavedKey, setHasSavedKey] = useState(false)
  const [showAiKey, setShowAiKey] = useState(false)
  const [aiMenuError, setAiMenuError] = useState('')
  const [availableModels, setAvailableModels] = useState<string[]>([])
  const [isFetchingModels, setIsFetchingModels] = useState(false)
  const [modelDropdownOpen, setModelDropdownOpen] = useState(false)
  const [modelFilter, setModelFilter] = useState('')
  const modelComboboxRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!modelDropdownOpen) return
    const handleClickOutside = (e: MouseEvent) => {
      if (modelComboboxRef.current && !modelComboboxRef.current.contains(e.target as Node)) {
        setModelDropdownOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [modelDropdownOpen])

  const filteredModels = useMemo(() => {
    if (!modelFilter.trim()) return availableModels
    const q = modelFilter.toLowerCase().trim()
    return availableModels.filter((m) => m.toLowerCase().includes(q))
  }, [availableModels, modelFilter])

  const loadAiProviderStatus = useCallback(async () => {
    try {
      const settings = await window.aiOfficeAi?.getSettings()
      if (!settings) {
        setIsAiConnected(false)
        return
      }
      const provider = settings.provider === 'custom' ? 'custom' : 'google'
      setActiveAiProvider(provider)
      const configuredModel =
        (provider === 'custom'
          ? settings.providers.custom?.model
          : settings.providers.google?.model
        )?.trim() ?? ''
      setActiveModel(configuredModel)

      if (provider === 'custom') {
        const customUrl = settings.providers.custom?.baseUrl?.trim()
        if (!customUrl) {
          setIsAiConnected(false)
        } else {
          setIsFetchingActiveModel(true)
          try {
            const res = await window.aiOfficeAi?.fetchCustomModels?.()
            if (res?.ok && Array.isArray(res.models) && res.models.length > 0) {
              setAvailableModels(res.models)
              if (!configuredModel) setActiveModel(res.models[0])
              setIsAiConnected(true)
            } else {
              setIsAiConnected(false)
            }
          } catch {
            setIsAiConnected(false)
          } finally {
            setIsFetchingActiveModel(false)
          }
        }
      } else {
        const hasKey = Boolean(settings.providers.google?.hasApiKey)
        setHasSavedGoogleKey(hasKey)
        setIsAiConnected(hasKey)
      }
    } catch {
      setIsAiConnected(false)
      /* bridge unavailable */
    }
  }, [])

  useEffect(() => {
    void loadAiProviderStatus()
    const onFocus = () => void loadAiProviderStatus()
    window.addEventListener('focus', onFocus)
    return () => window.removeEventListener('focus', onFocus)
  }, [loadAiProviderStatus])

  const openAiMenu = useCallback(async () => {
    setAiMenuError('')
    setShowAiKey(false)
    setModelDropdownOpen(false)
    setModelFilter('')
    setAiMenuOpen(true)
    try {
      const settings = await window.aiOfficeAi?.getSettings()
      if (!settings) return
      const provider = settings.provider === 'custom' ? 'custom' : 'google'
      setAiChoice(provider)
      setGoogleModel(settings.providers.google?.model || DEFAULT_GOOGLE_MODEL)
      setGoogleKey('') // stored keys never reach the renderer; blank keeps them
      setHasSavedGoogleKey(Boolean(settings.providers.google?.hasApiKey))
      const endpoint = settings.providers.custom?.baseUrl ?? ''
      const model = settings.providers.custom?.model ?? ''
      setAiEndpoint(endpoint)
      setAiModelId(model)
      setAiKey('')
      setHasSavedKey(Boolean(settings.providers.custom?.hasApiKey))

      if (endpoint) {
        setIsFetchingModels(true)
        try {
          const res = await window.aiOfficeAi?.fetchCustomModels?.()
          if (res?.ok && Array.isArray(res.models) && res.models.length > 0) {
            setAvailableModels(res.models)
            if (!model && res.models[0]) {
              setAiModelId(res.models[0])
            }
          }
        } finally {
          setIsFetchingModels(false)
        }
      }
    } catch {
      /* bridge unavailable: fields stay empty */
    }
  }, [])

  // An editor's "Configure AI Provider" card asks the shell to open this
  // dialog; the shell activates the Home tab and bumps this counter.
  const openAiMenuRef = useRef(openAiMenu)
  openAiMenuRef.current = openAiMenu
  useEffect(() => {
    if (openAiSettings) void openAiMenuRef.current()
  }, [openAiSettings])

  const handleFetchModels = useCallback(async () => {
    const endpoint = aiEndpoint.trim().replace(/\/+$/, '')
    if (!endpoint) {
      setAiMenuError(t('aiMenuIncomplete'))
      return
    }
    setIsFetchingModels(true)
    setAiMenuError('')
    try {
      const res = await window.aiOfficeAi?.fetchCustomModels?.({
        baseUrl: endpoint,
        apiKey: aiKey.trim() || undefined,
      })
      if (res?.ok && Array.isArray(res.models) && res.models.length > 0) {
        setAvailableModels(res.models)
        setModelDropdownOpen(true)
        if (!aiModelId.trim() && res.models[0]) {
          setAiModelId(res.models[0])
        }
      } else if (res?.error) {
        setAiMenuError(res.error)
      } else {
        setAiMenuError('No models returned from endpoint')
      }
    } catch (err) {
      setAiMenuError(err instanceof Error ? err.message : String(err))
    } finally {
      setIsFetchingModels(false)
    }
  }, [aiEndpoint, aiKey, aiModelId, t])

  const saveAiMenu = useCallback(async () => {
    const endpoint = aiEndpoint.trim().replace(/\/+$/, '')
    if (aiChoice === 'custom' && (!/^https?:\/\/\S+$/.test(endpoint) || !aiModelId.trim())) {
      setAiMenuError(t('aiMenuIncomplete'))
      return
    }
    if (aiChoice === 'google' && !googleModel.trim()) {
      setAiMenuError(t('aiMenuIncomplete'))
      return
    }
    try {
      // Reads are redacted (apiKey ''), so pass the typed key through untouched
      // and let the main process keep the stored key when the field is left blank.
      const settings = await window.aiOfficeAi?.getSettings()
      if (!settings) throw new Error('AI bridge unavailable')
      const next: AiSettings = {
        provider: aiChoice,
        providers: {
          google:
            aiChoice === 'google'
              ? { model: googleModel.trim(), apiKey: googleKey }
              : settings.providers.google,
          custom:
            aiChoice === 'custom'
              ? { model: aiModelId.trim(), baseUrl: endpoint, apiKey: aiKey }
              : settings.providers.custom,
        },
      }
      await window.aiOfficeAi?.setSettings(next)
      setAiMenuOpen(false)

      setActiveAiProvider(aiChoice)
      if (aiChoice === 'custom') {
        setActiveModel(aiModelId.trim())
        setHasSavedKey(Boolean(hasSavedKey || aiKey.trim()))
        if (endpoint) {
          void window.aiOfficeAi
            ?.fetchCustomModels?.()
            .then((res) => {
              if (res?.ok && Array.isArray(res.models) && res.models.length > 0) {
                setAvailableModels(res.models)
                setIsAiConnected(true)
                if (!aiModelId.trim() && res.models[0]) {
                  setActiveModel(res.models[0])
                }
              } else {
                setIsAiConnected(false)
              }
            })
            .catch(() => setIsAiConnected(false))
        } else {
          setIsAiConnected(false)
        }
      } else {
        const hasKey = Boolean(hasSavedGoogleKey || googleKey.trim())
        setActiveModel(googleModel.trim())
        setHasSavedGoogleKey(hasKey)
        setIsAiConnected(hasKey)
      }
    } catch (err) {
      setAiMenuError(err instanceof Error ? err.message : String(err))
    }
  }, [
    aiChoice,
    aiEndpoint,
    aiModelId,
    aiKey,
    hasSavedKey,
    googleKey,
    googleModel,
    hasSavedGoogleKey,
    t,
  ])
  useEffect(() => {
    void window.aiOffice
      .getSupportLinks()
      .then((links) => {
        setWebsiteUrl(links.websiteUrl)
        if (links.vendorName) setVendor(links.vendorName)
      })
      .catch(() => {})
  }, [])

  useEffect(() => {
    if (!diagramModeModal) return
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setDiagramModeModal(false)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [diagramModeModal])
  useEffect(() => {
    void window.aiOffice.getAppVersion().then((version) => {
      if (version) setAppVersion(version)
    })
    void window.aiOffice.getUpdateState().then((state) => {
      if (state) setUpdateState(state)
    })
    void window.aiOffice.getJustUpdated().then((justUpdated) => {
      if (justUpdated) {
        setShowChangelog(true)
        void window.aiOffice.clearJustUpdated()
      }
    })
    const unsubscribe = window.aiOffice.onUpdateStateChanged((state) => {
      setUpdateState(state)
      if (state.status === 'downloaded') {
        setShowChangelog(false)
        setShowUpdateRestart(true)
      }
    })
    return () => {
      unsubscribe()
    }
  }, [])
  // ── Project state ──
  const [projects, setProjects] = useState<ProjectSummaryEntry[]>([])
  const [selectedProjectId, setSelectedProjectId] = useState<string | null>(null)

  const projectMode = hasProjectApi()

  // ── Paged loading ──
  // stale responses are dropped via a request sequence number (when views/filters switch quickly)
  const requestSeq = useRef(0)
  const entriesLen = useRef(0)
  entriesLen.current = entries.length

  /** reload the list; keepCount keeps the loaded row count (refresh), otherwise back to page one */
  const reload = (keepCount: boolean) => {
    const seq = ++requestSeq.current
    const ext = filter === 'all' ? undefined : filter
    const limit = keepCount ? Math.max(entriesLen.current, PAGE_SIZE) : PAGE_SIZE
    const primary = view === 'recent' ? window.aiOffice.recents : window.aiOffice.starred
    const secondary = view === 'recent' ? window.aiOffice.starred : window.aiOffice.recents
    void primary({ offset: 0, limit, ext }).then((page) => {
      if (seq !== requestSeq.current) return
      setEntries(page.entries)
      setListTotal(page.total)
      setNavCounts((prev) =>
        view === 'recent'
          ? { ...prev, recent: visiblePageCount(page) }
          : { ...prev, starred: visiblePageCount(page) },
      )
    })
    // The other view fetches only its count under the same active filter.
    void secondary({ offset: 0, limit: 0, ext }).then((page) => {
      if (seq !== requestSeq.current) return
      setNavCounts((prev) =>
        view === 'recent'
          ? { ...prev, starred: visiblePageCount(page) }
          : { ...prev, recent: visiblePageCount(page) },
      )
    })
    if (projectMode) {
      void window.aiOfficeProject!.listProjects().then(setProjects)
    }
  }
  const reloadRef = useRef(reload)
  reloadRef.current = reload

  // refresh signal for project-view data (re-pull file stats after file changes)
  const [projectTick, setProjectTick] = useState(0)

  const refresh = () => {
    reloadRef.current(true)
    setProjectTick((n) => n + 1)
  }

  useEffect(() => {
    reloadRef.current(false)
  }, [view, filter])

  useEffect(() => {
    const onFocus = () => {
      reloadRef.current(true)
      setProjectTick((n) => n + 1)
    }
    window.addEventListener('focus', onFocus)
    return () => window.removeEventListener('focus', onFocus)
  }, [])

  const hasMore = entries.length < listTotal

  const loadMore = () => {
    if (loadingMore || !hasMore) return
    setLoadingMore(true)
    const seq = requestSeq.current
    const ext = filter === 'all' ? undefined : filter
    const api = view === 'recent' ? window.aiOffice.recents : window.aiOffice.starred
    void api({ offset: entriesLen.current, limit: PAGE_SIZE, ext }).then((page) => {
      setLoadingMore(false)
      if (seq !== requestSeq.current) return
      setEntries((prev) => [...prev, ...page.entries])
      setListTotal(page.total)
    })
  }
  const loadMoreRef = useRef(loadMore)
  loadMoreRef.current = loadMore

  // Load the next page once the bottom sentinel enters the viewport (240px early);
  // depending on entries.length rebuilds the observer after each page — observe fires an immediate
  // callback, so while the sentinel stays in view we keep loading until full or exhausted
  const sentinelRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const el = sentinelRef.current
    if (!el) return
    const observer = new IntersectionObserver(
      (records) => {
        if (records.some((r) => r.isIntersecting)) loadMoreRef.current()
      },
      { rootMargin: '240px' },
    )
    observer.observe(el)
    return () => observer.disconnect()
  }, [hasMore, entries.length])

  useEffect(() => {
    if (rowMenu === null && confirmDelete === null) return
    const onPointerDown = (event: PointerEvent) => {
      const target = event.target as Element | null
      if (rowMenu !== null && !target?.closest?.('.recent-actions')) setRowMenu(null)
    }
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setRowMenu(null)
        setConfirmDelete(null)
      }
    }
    window.addEventListener('pointerdown', onPointerDown)
    window.addEventListener('keydown', onKeyDown)
    return () => {
      window.removeEventListener('pointerdown', onPointerDown)
      window.removeEventListener('keydown', onKeyDown)
    }
  }, [rowMenu, confirmDelete])

  // ── Project files state ────────────────────────────────

  const [projectFileEntries, setProjectFileEntries] = useState<RecentEntry[]>([])
  const [moveFileMenu, setMoveFileMenu] = useState<string | null>(null)
  // submenu opens rightward by default; flips left when the window edge is too close
  const [moveMenuFlip, setMoveMenuFlip] = useState(false)
  // hover-open/close delays: avoid flashing the submenu while the pointer passes
  // through, and keep it open while crossing the 4px gap into it
  const moveMenuTimers = useRef<{ open: number | null; close: number | null }>({
    open: null,
    close: null,
  })

  const openMoveMenu = (path: string) => {
    setMoveMenuFlip(false)
    setMoveFileMenu(path)
  }

  // ref runs pre-paint, so measuring the real width (long project names exceed
  // the min-width) and flipping never flashes; once flipped the check no longer hits
  const measureSubmenu = (el: HTMLDivElement | null) => {
    if (el && el.getBoundingClientRect().right > document.documentElement.clientWidth - 8) {
      setMoveMenuFlip(true)
    }
  }

  const clearMoveMenuTimer = (kind: 'open' | 'close') => {
    const timers = moveMenuTimers.current
    if (timers[kind] !== null) {
      window.clearTimeout(timers[kind])
      timers[kind] = null
    }
  }
  const [bulkMoveMenu, setBulkMoveMenu] = useState(false)

  useEffect(() => {
    if (!projectMode || !selectedProjectId) {
      setProjectFileEntries([])
      return
    }
    let active = true
    const api = window.aiOfficeProject!
    void api.listFiles(selectedProjectId).then(async (paths) => {
      const stats = await window.aiOffice.statPaths(paths)
      if (!active) return
      setProjectFileEntries(stats.sort((a, b) => b.mtimeMs - a.mtimeMs))
    })
    return () => {
      active = false
    }
  }, [projectMode, selectedProjectId, projectTick])

  // the submenu lives inside the row menu: when that closes, drop the stale
  // submenu state and any pending hover timers so it doesn't reopen expanded
  useEffect(() => {
    if (rowMenu === null) {
      clearMoveMenuTimer('open')
      clearMoveMenuTimer('close')
      setMoveFileMenu(null)
    }
  }, [rowMenu])

  // close the move-file menu
  useEffect(() => {
    if (!moveFileMenu) return
    const handler = (e: PointerEvent) => {
      const target = e.target as Element | null
      if (!target?.closest?.('.move-menu-wrap')) setMoveFileMenu(null)
    }
    window.addEventListener('pointerdown', handler)
    return () => window.removeEventListener('pointerdown', handler)
  }, [moveFileMenu])

  // close the bulk move-to-project menu in the selection bar
  useEffect(() => {
    if (!bulkMoveMenu) return
    const handler = (e: PointerEvent) => {
      const target = e.target as Element | null
      if (!target?.closest?.('.selection-move-wrap')) setBulkMoveMenu(false)
    }
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setBulkMoveMenu(false)
    }
    window.addEventListener('pointerdown', handler)
    window.addEventListener('keydown', onKeyDown)
    return () => {
      window.removeEventListener('pointerdown', handler)
      window.removeEventListener('keydown', onKeyDown)
    }
  }, [bulkMoveMenu])

  // ── Plain view (no project selected): filtering runs in the main process; entries is the visible list ──
  const selectedPaths = entries.filter((e) => selected.has(e.path)).map((e) => e.path)
  const allSelected = entries.length > 0 && selectedPaths.length === entries.length

  // project view shares the same `selected` set (keyed by path)
  const projSelectedPaths = projectFileEntries
    .filter((e) => selected.has(e.path))
    .map((e) => e.path)
  const projAllSelected =
    projectFileEntries.length > 0 && projSelectedPaths.length === projectFileEntries.length

  const changeView = (next: 'recent' | 'starred') => {
    setView(next)
    setSelected(new Set())
    setRowMenu(null)
  }

  const changeFilter = (key: string) => {
    setFilter(key)
    setSelected(new Set())
    setRowMenu(null)
  }

  const toggleSelect = (path: string, on: boolean) => {
    setSelected((prev) => {
      const next = new Set(prev)
      if (on) next.add(path)
      else next.delete(path)
      return next
    })
  }

  const toggleSelectAll = () => {
    setSelected(allSelected ? new Set() : new Set(entries.map((e) => e.path)))
  }

  const toggleSelectAllProject = () => {
    setSelected(projAllSelected ? new Set() : new Set(projectFileEntries.map((e) => e.path)))
  }

  const toggleStar = (path: string) => {
    void window.aiOffice.toggleStar(path).then(refresh)
  }

  const removeRecent = (paths: string[]) => {
    setRowMenu(null)
    setSelected(new Set())
    void window.aiOffice.removeRecent(paths).then(refresh)
  }

  const deleteFiles = (paths: string[]) => {
    setRowMenu(null)
    setConfirmDelete(paths)
  }

  const confirmDeleteNow = () => {
    const paths = confirmDelete ?? []
    setConfirmDelete(null)
    setSelected(new Set())
    void window.aiOffice.deleteFiles(paths).then(refresh)
  }

  const duplicateFile = (path: string) => {
    setRowMenu(null)
    void window.aiOffice.duplicateFile(path).then(refresh)
  }

  const startRename = (entry: RecentEntry) => {
    setRowMenu(null)
    setRenaming({ path: entry.path, value: baseName(entry) })
  }

  const commitRename = (entry: RecentEntry) => {
    const value = renaming?.value.trim() ?? ''
    setRenaming(null)
    if (!value || value === baseName(entry)) return
    const newName = entry.ext ? `${value}.${entry.ext}` : value
    void window.aiOffice.renameFile(entry.path, newName).then((result) => {
      if (!result.ok) window.alert(result.error ?? t('renameFailed'))
      refresh()
    })
  }

  const moveFileTo = async (filePath: string, targetProjectId: string) => {
    setMoveFileMenu(null)
    setRowMenu(null)
    await window.aiOfficeProject?.moveFile(filePath, targetProjectId)
    refresh()
    if (selectedProjectId) {
      setProjectFileEntries((prev) => prev.filter((e) => e.path !== filePath))
    }
  }

  const moveFilesTo = async (paths: string[], targetProjectId: string) => {
    setBulkMoveMenu(false)
    setSelected(new Set())
    // drop moved rows immediately (same as moveFileTo) so they cannot be
    // re-selected or re-moved while the sequential IPC loop is in flight
    const moved = new Set(paths)
    setProjectFileEntries((prev) => prev.filter((e) => !moved.has(e.path)))
    for (const path of paths) {
      await window.aiOfficeProject?.moveFile(path, targetProjectId)
    }
    refresh()
  }

  // ── New file (passes projectId when a project is selected) ──
  const handleNewDoc = useCallback(() => {
    void window.aiOffice.newDoc(selectedProjectId ? { projectId: selectedProjectId } : undefined)
  }, [selectedProjectId])

  const handleNewSheet = () => {
    void window.aiOffice.newSheet(selectedProjectId ? { projectId: selectedProjectId } : undefined)
  }

  const handleNewSlide = () => {
    void window.aiOffice.newSlide(selectedProjectId ? { projectId: selectedProjectId } : undefined)
  }

  const handleNewMarkdown = () => {
    void window.aiOffice.newMarkdown(
      selectedProjectId ? { projectId: selectedProjectId } : undefined,
    )
  }
  const handleNewDiagram = () => {
    setDiagramModeModal(true)
  }

  const openNewDiagram = (mode: DiagramMode) => {
    setDiagramModeModal(false)
    void window.aiOffice.newDiagram(
      selectedProjectId ? { projectId: selectedProjectId, mode } : { mode },
    )
  }

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const active = document.activeElement
      if (
        active instanceof HTMLInputElement ||
        active instanceof HTMLTextAreaElement ||
        active?.getAttribute('contenteditable') === 'true'
      ) {
        return
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'o') {
        e.preventDefault()
        void window.aiOffice.browse()
        return
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'n') {
        e.preventDefault()
        handleNewDoc()
        return
      }
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [handleNewDoc])

  const FORMAT_ITEMS: Array<{ ext: string; title: string; sub: string; action: () => void }> = [
    { ext: 'docx', title: t('newDoc'), sub: '.docx', action: handleNewDoc },
    { ext: 'xlsx', title: t('newSheet'), sub: '.xlsx', action: handleNewSheet },
    { ext: 'pptx', title: t('newSlide'), sub: '.pptx', action: handleNewSlide },
    {
      ext: 'pdf',
      title: t('newPdf'),
      sub: '.pdf',
      action: () => window.aiOffice.openPdf(),
    },
    { ext: 'md', title: t('newMarkdown'), sub: '.md', action: handleNewMarkdown },
    { ext: 'drawio', title: t('newDiagram'), sub: '.drawio / .mmd', action: handleNewDiagram },
  ]

  function renderQuickCards() {
    return (
      <div className="quick-start-grid">
        <div className="quick-cards">
          {FORMAT_ITEMS.map((item) => (
            <button
              key={item.ext}
              className="quick-card"
              onClick={() => item.action()}
              aria-label={`${item.title} (${item.sub})`}
              title={`${item.title} (${item.sub})`}
            >
              <span className="quick-icon" aria-hidden="true">
                <FileBadge ext={item.ext} size={36} />
              </span>
              <span className="quick-text">
                <span className="quick-title-row">
                  <span className="quick-title">{item.title}</span>
                </span>
                <span className="quick-sub">{item.sub}</span>
              </span>
            </button>
          ))}
        </div>
      </div>
    )
  }

  // ── File row rendering (shared by the plain view and the project files view) ──

  function renderFileRow(entry: RecentEntry, context: 'global' | 'project') {
    const isRenaming = renaming?.path === entry.path
    const otherProjects = projects.filter(
      (p) => p.id !== (context === 'project' ? selectedProjectId : undefined),
    )
    return (
      <li className="recent-row" key={entry.path}>
        <div
          className="recent-item"
          role="button"
          tabIndex={0}
          onClick={() => {
            if (!isRenaming) void window.aiOffice.openPath(entry.path)
          }}
          onKeyDown={(event) => {
            if (event.key === 'Enter' && event.target === event.currentTarget) {
              void window.aiOffice.openPath(entry.path)
            }
          }}
        >
          <span className="col-check" onClick={(event) => event.stopPropagation()}>
            <input
              type="checkbox"
              className="row-check"
              checked={selected.has(entry.path)}
              onChange={(event) => toggleSelect(entry.path, event.target.checked)}
              aria-label={t('selectFile', { name: entry.name })}
            />
          </span>
          <span className="recent-icon">
            <FileBadge ext={entry.ext} size={24} />
          </span>
          {isRenaming ? (
            <input
              className="rename-input"
              value={renaming.value}
              autoFocus
              onFocus={(event) => event.target.select()}
              onClick={(event) => event.stopPropagation()}
              onChange={(event) => setRenaming({ path: entry.path, value: event.target.value })}
              onBlur={() => commitRename(entry)}
              onKeyDown={(event) => {
                event.stopPropagation()
                if (event.key === 'Enter') commitRename(entry)
                if (event.key === 'Escape') setRenaming(null)
              }}
            />
          ) : (
            <span className="recent-name">{entry.name}</span>
          )}
          <span className="recent-path">{parentDir(entry.path)}</span>
          <span className="recent-time">{formatModified(entry.mtimeMs, i18n)}</span>
          <span className="recent-size">{formatSize(entry.sizeBytes)}</span>
          <button
            className={`star-btn${entry.starred ? ' starred' : ''}`}
            aria-label={entry.starred ? t('unstar') : t('star')}
            onClick={(event) => {
              event.stopPropagation()
              toggleStar(entry.path)
            }}
          >
            <svg width="15" height="15" viewBox="0 0 16 16" aria-hidden="true">
              <path
                d="M8 1.9l1.9 3.85 4.25.62-3.07 3 .72 4.23L8 11.6l-3.8 2 .72-4.23-3.07-3 4.25-.62z"
                fill={entry.starred ? '#f5a623' : 'none'}
                stroke={entry.starred ? '#f5a623' : 'currentColor'}
                strokeWidth="1.2"
                strokeLinejoin="round"
              />
            </svg>
          </button>
          <span className="recent-actions" onClick={(event) => event.stopPropagation()}>
            <button
              className="more-btn"
              aria-label={t('moreActions')}
              aria-expanded={rowMenu === entry.path}
              onClick={() => setRowMenu(rowMenu === entry.path ? null : entry.path)}
            >
              <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
                <circle cx="3.2" cy="8" r="1.4" fill="currentColor" />
                <circle cx="8" cy="8" r="1.4" fill="currentColor" />
                <circle cx="12.8" cy="8" r="1.4" fill="currentColor" />
              </svg>
            </button>
            {rowMenu === entry.path && (
              <div className="row-menu" role="menu">
                <button
                  role="menuitem"
                  onClick={() => {
                    setRowMenu(null)
                    void window.aiOffice.openPath(entry.path)
                  }}
                >
                  {t('open')}
                </button>
                <button
                  role="menuitem"
                  onClick={() => {
                    setRowMenu(null)
                    void window.aiOffice.revealPath(entry.path)
                  }}
                >
                  {t('revealInFolder')}
                </button>
                <button
                  role="menuitem"
                  onClick={() => {
                    setRowMenu(null)
                    void navigator.clipboard.writeText(entry.path)
                  }}
                >
                  {t('copyPath')}
                </button>
                {projectMode && otherProjects.length > 0 && (
                  <>
                    <div className="row-menu-divider" />
                    <div
                      className="move-menu-wrap"
                      onMouseEnter={() => {
                        clearMoveMenuTimer('close')
                        if (moveFileMenu === entry.path) return
                        clearMoveMenuTimer('open')
                        moveMenuTimers.current.open = window.setTimeout(
                          () => openMoveMenu(entry.path),
                          160,
                        )
                      }}
                      onMouseLeave={() => {
                        clearMoveMenuTimer('open')
                        clearMoveMenuTimer('close')
                        moveMenuTimers.current.close = window.setTimeout(
                          () => setMoveFileMenu(null),
                          140,
                        )
                      }}
                    >
                      <button
                        role="menuitem"
                        className="submenu-trigger"
                        onClick={(e) => {
                          e.stopPropagation()
                          clearMoveMenuTimer('open')
                          clearMoveMenuTimer('close')
                          if (moveFileMenu === entry.path) setMoveFileMenu(null)
                          else openMoveMenu(entry.path)
                        }}
                      >
                        {t('moveToProject')}
                        <svg
                          width="11"
                          height="11"
                          viewBox="0 0 12 12"
                          aria-hidden="true"
                          style={{ marginLeft: 'auto' }}
                        >
                          <path
                            d="M4.5 2.5l4 3.5-4 3.5"
                            stroke="currentColor"
                            strokeWidth="1.3"
                            strokeLinecap="round"
                            fill="none"
                          />
                        </svg>
                      </button>
                      {moveFileMenu === entry.path && (
                        <div
                          className={`submenu${moveMenuFlip ? ' submenu-left' : ''}`}
                          role="menu"
                          ref={measureSubmenu}
                        >
                          {otherProjects.map((p) => (
                            <button
                              key={p.id}
                              role="menuitem"
                              onClick={() => void moveFileTo(entry.path, p.id)}
                            >
                              {p.isDefault ? t('defaultProject') : p.name}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  </>
                )}
                <div className="row-menu-divider" />
                <button role="menuitem" onClick={() => startRename(entry)}>
                  {t('rename')}
                </button>
                <button role="menuitem" onClick={() => duplicateFile(entry.path)}>
                  {t('duplicate')}
                </button>
                {context === 'global' && selectedPaths.length === 0 && (
                  <>
                    <div className="row-menu-divider" />
                    <button role="menuitem" onClick={() => removeRecent([entry.path])}>
                      {t('removeFromList')}
                    </button>
                    <button
                      role="menuitem"
                      className="danger"
                      onClick={() => deleteFiles([entry.path])}
                    >
                      {t('deleteFiles')}
                    </button>
                  </>
                )}
              </div>
            )}
          </span>
        </div>
      </li>
    )
  }

  // ── Project files view ────────────────────────────────

  function renderProjectContent() {
    const proj = projects.find((p) => p.id === selectedProjectId)
    if (!proj) return null
    const otherProjects = projects.filter((p) => p.id !== proj.id)

    return (
      <main className="content">
        <section className="quick-start" aria-label={t('secQuickStart')}>
          <div className="section-head">
            <span className="section-label">{t('secQuickStart')}</span>
          </div>
          {renderQuickCards()}
        </section>

        <section className="recents" aria-label={t('secProjectFiles')}>
          <div className="recents-toolbar">
            <div className="recents-heading">
              <span className="section-label">{t('secProjectFiles')}</span>
              <span className="file-count">
                {t(fileCountKey(projectFileEntries.length), { n: projectFileEntries.length })}
              </span>
            </div>
            {projSelectedPaths.length > 0 && (
              <div className="selection-bar">
                <span className="selection-count">
                  {t('selectedCount', { n: projSelectedPaths.length })}
                </span>
                {otherProjects.length > 0 && (
                  <span className="selection-move-wrap">
                    <button
                      className="selection-action"
                      aria-expanded={bulkMoveMenu}
                      onClick={() => setBulkMoveMenu((open) => !open)}
                    >
                      {t('moveToProject')}
                    </button>
                    {bulkMoveMenu && (
                      <div className="selection-move-menu" role="menu">
                        {otherProjects.map((p) => (
                          <button
                            key={p.id}
                            role="menuitem"
                            onClick={() => void moveFilesTo(projSelectedPaths, p.id)}
                          >
                            {p.isDefault ? t('defaultProject') : p.name}
                          </button>
                        ))}
                      </div>
                    )}
                  </span>
                )}
                <button
                  className="selection-action danger"
                  onClick={() => deleteFiles(projSelectedPaths)}
                >
                  {t('deleteFiles')}
                </button>
                <button className="selection-action" onClick={() => setSelected(new Set())}>
                  {t('cancel')}
                </button>
              </div>
            )}
          </div>

          {projectFileEntries.length === 0 ? (
            <p className="empty proj-empty">
              <svg
                className="proj-empty-icon"
                width="48"
                height="48"
                viewBox="0 0 24 24"
                fill="none"
                aria-hidden="true"
              >
                <path
                  d="M6.29297 3.75H14.1729C14.4927 3.75 14.7979 3.88392 15.0146 4.11914L18.5566 7.96387C18.7512 8.17512 18.8593 8.45208 18.8594 8.73926V19.1055C18.8593 19.7376 18.346 20.25 17.7139 20.25H6.29297C5.66091 20.2499 5.14855 19.7375 5.14844 19.1055V4.89453C5.14855 4.26247 5.66091 3.75011 6.29297 3.75Z"
                  stroke="currentColor"
                  strokeWidth="1.5"
                />
                <path
                  d="M13.8984 4V7.11C13.8984 8.15382 14.7446 9 15.7884 9H18.8984"
                  stroke="currentColor"
                  strokeWidth="1.5"
                />
              </svg>
              <span className="empty-hint">{t('projEmptyHint')}</span>
            </p>
          ) : (
            <div className="recent-table">
              <div className="recent-columns">
                <span className="col-check">
                  <input
                    type="checkbox"
                    checked={projAllSelected}
                    onChange={toggleSelectAllProject}
                    aria-label={t('selectAll')}
                  />
                </span>
                <span className="col-name">{t('colName')}</span>
                <span>{t('colLocation')}</span>
                <span>{t('colModified')}</span>
                <span className="col-size">{t('colSize')}</span>
                <span />
                <span />
              </div>
              <ul className="recent-list">
                {projectFileEntries.map((entry) => renderFileRow(entry, 'project'))}
              </ul>
            </div>
          )}
        </section>
      </main>
    )
  }

  // ── Plain view ────────────────────────────────────────

  function renderGlobalContent() {
    const now = new Date()
    const hour = now.getHours()
    const greetKey =
      hour < 6
        ? 'greetEvening'
        : hour < 12
          ? 'greetMorning'
          : hour < 18
            ? 'greetAfternoon'
            : 'greetEvening'
    const cjk = lang === 'zh' || lang === 'zh-TW' || lang === 'ja'
    const greeting = `${t(greetKey)}${cjk ? '。' : '.'}`
    return (
      <main className="content">
        <section className="quick-start" aria-label={t('secQuickStart')}>
          <div className="home-hero">
            <h1 className="hero-title">
              {greeting}
              <span className="hero-ask">{t('greetAsk1')}</span>
            </h1>
            <div className="hero-actions">
              <button
                className="open-file-button"
                onClick={() => void window.aiOffice.browse()}
                title={t('openLocal')}
              >
                <svg width="18" height="18" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                  <path
                    d="M1.5 4A1.5 1.5 0 0 1 3 2.5h3.1c.44 0 .85.19 1.13.52L8.4 4.4H13A1.5 1.5 0 0 1 14.5 5.9v5.6A1.5 1.5 0 0 1 13 13H3a1.5 1.5 0 0 1-1.5-1.5V4z"
                    stroke="currentColor"
                    strokeWidth="1.3"
                    strokeLinejoin="round"
                  />
                </svg>
                <span>{t('openLocal')}</span>
              </button>
              <button
                className={`open-file-button ai-provider-button${isAiConnected ? ' is-connected' : ' is-disconnected'}`}
                onClick={() => void openAiMenu()}
                title={`${activeModel || (activeAiProvider === 'custom' ? t('aiMenuCustom') : t('aiMenuGoogle'))} (${
                  activeAiProvider === 'custom' ? t('aiMenuCustom') : t('aiMenuGoogle')
                })`}
                aria-label={t('aiMenuTitle')}
                aria-haspopup="dialog"
                aria-expanded={aiMenuOpen}
              >
                <span
                  className={`ai-provider-button-dot ${isAiConnected ? 'connected' : 'disconnected'}`}
                  aria-hidden="true"
                />
                <span className="ai-provider-button-model" title={activeModel || undefined}>
                  {isFetchingActiveModel && !activeModel
                    ? '...'
                    : activeModel || (activeAiProvider === 'custom' ? 'OpenAI' : 'Gemini')}
                </span>
                <span className="ai-provider-button-chevron" aria-hidden="true">
                  ▾
                </span>
              </button>
            </div>
          </div>
          {renderQuickCards()}
        </section>

        <section
          className="recents"
          aria-label={view === 'recent' ? t('secRecent') : t('secStarred')}
        >
          <div className="recents-toolbar">
            <div className="recents-heading">
              <span className="section-label">
                {view === 'recent' ? t('secRecent') : t('secStarred')}
              </span>
              <span className="file-count">{t(fileCountKey(listTotal), { n: listTotal })}</span>
            </div>
            {selectedPaths.length > 0 ? (
              <div className="selection-bar">
                <span className="selection-count">
                  {t('selectedCount', { n: selectedPaths.length })}
                </span>
                <button className="selection-action" onClick={() => removeRecent(selectedPaths)}>
                  {t('removeFromList')}
                </button>
                <button
                  className="selection-action danger"
                  onClick={() => deleteFiles(selectedPaths)}
                >
                  {t('deleteFiles')}
                </button>
                <button className="selection-action" onClick={() => setSelected(new Set())}>
                  {t('cancel')}
                </button>
              </div>
            ) : (
              <div className="filter-pills" role="tablist" aria-label={t('filterAria')}>
                {FILTERS.map((f) => (
                  <button
                    key={f.key}
                    className={`filter-pill${filter === f.key ? ' active' : ''}`}
                    onClick={() => changeFilter(f.key)}
                  >
                    {t(f.label)}
                  </button>
                ))}
              </div>
            )}
          </div>

          {entries.length === 0 ? (
            <p className="empty proj-empty">
              <svg
                className="proj-empty-icon"
                width="48"
                height="48"
                viewBox="0 0 24 24"
                fill="none"
                aria-hidden="true"
              >
                <path
                  d="M6.29297 3.75H14.1729C14.4927 3.75 14.7979 3.88392 15.0146 4.11914L18.5566 7.96387C18.7512 8.17512 18.8593 8.45208 18.8594 8.73926V19.1055C18.8593 19.7376 18.346 20.25 17.7139 20.25H6.29297C5.66091 20.2499 5.14855 19.7375 5.14844 19.1055V4.89453C5.14855 4.26247 5.66091 3.75011 6.29297 3.75Z"
                  stroke="currentColor"
                  strokeWidth="1.5"
                />
                <path
                  d="M13.8984 4V7.11C13.8984 8.15382 14.7446 9 15.7884 9H18.8984"
                  stroke="currentColor"
                  strokeWidth="1.5"
                />
              </svg>
              <span className="empty-hint">
                {view === 'starred'
                  ? t('emptyStarred')
                  : navCounts.recent === 0
                    ? t('emptyRecent')
                    : t('emptyFiltered')}
              </span>
            </p>
          ) : (
            <div className={`recent-table${selectedPaths.length > 0 ? ' has-selection' : ''}`}>
              <div className="recent-columns">
                <span className="col-check">
                  <input
                    type="checkbox"
                    checked={allSelected}
                    onChange={toggleSelectAll}
                    aria-label={t('selectAll')}
                  />
                </span>
                <span className="col-name">{t('colName')}</span>
                <span>{t('colLocation')}</span>
                <span>{t('colModified')}</span>
                <span className="col-size">{t('colSize')}</span>
                <span />
                <span />
              </div>
              <ul className="recent-list">
                {entries.map((entry) => renderFileRow(entry, 'global'))}
              </ul>
              {hasMore && (
                <div ref={sentinelRef} className="load-more" aria-hidden="true">
                  <span className="load-more-spinner" />
                </div>
              )}
            </div>
          )}
        </section>
      </main>
    )
  }

  return (
    <div
      className={`home${isDraggingFile ? ' home-dragover' : ''}`}
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes('Files')) {
          e.preventDefault()
          e.stopPropagation()
          setIsDraggingFile(true)
        }
      }}
      onDragLeave={(e) => {
        if (!e.currentTarget.contains(e.relatedTarget as Node | null)) {
          setIsDraggingFile(false)
        }
      }}
      onDrop={(e) => {
        if (e.dataTransfer.types.includes('Files')) {
          e.preventDefault()
          e.stopPropagation()
          setIsDraggingFile(false)
          const files = Array.from(e.dataTransfer.files)
          for (const file of files) {
            const filePath = 'path' in file && typeof file.path === 'string' ? file.path : null
            if (filePath) {
              void window.aiOffice.openPath(filePath)
            }
          }
        }
      }}
    >
      {isDraggingFile && (
        <div className="home-drop-overlay" aria-hidden="true">
          <div className="home-drop-modal">
            <svg
              width="48"
              height="48"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.6"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="M4 14.899A7 7 0 1 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242" />
              <path d="M12 12v9" />
              <path d="m16 16-4-4-4 4" />
            </svg>
            <h3>Drop files to open</h3>
            <p>.docx, .xlsx, .pptx, .pdf, .md</p>
          </div>
        </div>
      )}
      <aside className="sidebar">
        <nav className="sidebar-nav">
          <button
            className={`nav-item${view === 'recent' && !selectedProjectId ? ' active' : ''}`}
            onClick={() => {
              changeView('recent')
              setSelectedProjectId(null)
            }}
          >
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
              <circle cx="8" cy="8" r="6.2" stroke="currentColor" strokeWidth="1.3" />
              <path
                d="M8 4.8V8l2.2 1.6"
                stroke="currentColor"
                strokeWidth="1.3"
                strokeLinecap="round"
              />
            </svg>
            <span className="nav-label">{t('navRecent')}</span>
            <span className="nav-count">{navCounts.recent}</span>
          </button>
          <button
            className={`nav-item${view === 'starred' && !selectedProjectId ? ' active' : ''}`}
            onClick={() => {
              changeView('starred')
              setSelectedProjectId(null)
            }}
          >
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
              <path
                d="M8 1.9l1.9 3.85 4.25.62-3.07 3 .72 4.23L8 11.6l-3.8 2 .72-4.23-3.07-3 4.25-.62z"
                stroke="currentColor"
                strokeWidth="1.3"
                strokeLinejoin="round"
              />
            </svg>
            <span className="nav-label">{t('navStarred')}</span>
            <span className="nav-count">{navCounts.starred}</span>
          </button>
        </nav>

        {/* project sidebar */}
        {projectMode && (
          <>
            <div className="sidebar-divider" />
            <ProjectPanel
              projects={projects}
              selectedId={selectedProjectId}
              onSelect={(id) => {
                setSelectedProjectId(id)
                // reset list-selection state on any project switch (paths are
                // shared between the plain view and project views)
                setSelected(new Set())
                setRowMenu(null)
              }}
              onRefresh={refresh}
            />
          </>
        )}
        <div className="sidebar-about">
          <div className="sidebar-brand-row">
            <div className="sidebar-brand-left">
              <button
                type="button"
                className="sidebar-brand"
                onClick={onOpenFeatures}
                aria-haspopup="dialog"
                aria-label={`${t('onbTitle1')} · ${t('appVersion', { v: appVersion })}`}
                title={`${t('onbTitle1')} · ${t('appVersion', { v: appVersion })}`}
              >
                ARAOffice
              </button>
              <button
                type="button"
                className={`sidebar-version-btn ${
                  updateState.status === 'downloading' ||
                  updateState.status === 'downloaded' ||
                  updateState.status === 'available'
                    ? 'has-update'
                    : ''
                }`}
                onClick={() => {
                  if (updateState.status === 'downloaded') {
                    setShowUpdateRestart(true)
                  } else {
                    setShowChangelog(true)
                  }
                }}
                title={
                  updateState.status === 'downloaded'
                    ? t('updateReadyTitle')
                    : updateState.status === 'downloading'
                      ? t('updateDownloading', { p: updateState.percent ?? 0 })
                      : updateState.status === 'available'
                        ? t('updateAvailable')
                        : t('changelogTitle')
                }
                aria-label={`${t('versionLabel')}: v${appVersion}`}
              >
                <span>v{appVersion}</span>
                {(updateState.status === 'downloading' ||
                  updateState.status === 'downloaded' ||
                  updateState.status === 'available') && (
                  <span
                    className={`version-pulse-dot ${
                      updateState.status === 'downloading'
                        ? 'downloading'
                        : updateState.status === 'downloaded'
                          ? 'ready'
                          : 'available'
                    }`}
                  />
                )}
              </button>
            </div>
            <button
              type="button"
              className="sidebar-feedback-btn"
              onClick={() => setShowFeedback(true)}
              title={t('feedbackTitle')}
              aria-label={t('feedbackTitle')}
            >
              <svg
                width="14"
                height="14"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.2"
                strokeLinecap="round"
                strokeLinejoin="round"
                aria-hidden="true"
              >
                <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
                <line x1="9" y1="10" x2="15" y2="10" />
                <line x1="12" y1="7" x2="12" y2="13" />
              </svg>
            </button>
          </div>
          <div className="sidebar-meta">
            <span className="sidebar-tagline">{t('tagline')}</span>
            <span className="sidebar-author">
              {websiteUrl ? (
                <button
                  type="button"
                  className="sidebar-author-btn"
                  onClick={() => void window.aiOfficeWindow.openExternal(websiteUrl)}
                  title={t('author', { vendor })}
                  aria-label={t('author', { vendor })}
                >
                  {t('author', { vendor })}
                </button>
              ) : (
                <span className="sidebar-author-btn" aria-hidden="true">
                  {t('author', { vendor })}
                </span>
              )}
            </span>
          </div>
          <div className="language-selector">
            <span className="language-selector-label">
              <svg
                className="language-selector-icon"
                width="13"
                height="13"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
                aria-hidden="true"
              >
                <circle cx="12" cy="12" r="10" />
                <path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1 4-10z" />
                <path d="M2 12h20" />
              </svg>
              <span>{t('language')}</span>
            </span>
            <LanguageMenu lang={lang} setLang={setLang} label={t('language')} />
          </div>
          <SocialLinks onOpenDonation={() => setShowDonation(true)} />
        </div>
      </aside>

      {selectedProjectId ? renderProjectContent() : renderGlobalContent()}

      {diagramModeModal && (
        <div className="modal-overlay" onClick={() => setDiagramModeModal(false)}>
          <div
            className="modal diagram-mode-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="diagram-mode-title"
            onClick={(event) => event.stopPropagation()}
          >
            <h3 id="diagram-mode-title">{t('newDiagram')}</h3>
            <p>{t('diagramModePrompt')}</p>
            <div className="diagram-mode-options">
              <button
                type="button"
                className="diagram-mode-option"
                autoFocus
                aria-label={`${t('diagramModeVisual')} (.drawio)`}
                onClick={() => openNewDiagram('visual')}
              >
                <FileBadge ext="drawio" size={32} />
                <span className="diagram-mode-copy">
                  <strong>{t('diagramModeVisual')}</strong>
                  <small>.drawio</small>
                </span>
              </button>
              <button
                type="button"
                className="diagram-mode-option"
                aria-label={`${t('diagramModeCode')} (.mmd)`}
                onClick={() => openNewDiagram('mermaid')}
              >
                <FileBadge ext="mmd" size={32} />
                <span className="diagram-mode-copy">
                  <strong>{t('diagramModeCode')}</strong>
                  <span className="diagram-mode-meta">
                    <small>.mmd</small>
                    <span className="diagram-mode-badge">beta</span>
                  </span>
                </span>
              </button>
            </div>
            <div className="modal-buttons">
              <button className="btn btn-secondary" onClick={() => setDiagramModeModal(false)}>
                {t('cancel')}
              </button>
            </div>
          </div>
        </div>
      )}

      {aiMenuOpen && (
        <div className="modal-overlay ai-menu-overlay" onClick={() => setAiMenuOpen(false)}>
          <div
            className="modal ai-menu-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="ai-menu-title"
            onClick={(event) => event.stopPropagation()}
          >
            {/* Modal Header */}
            <div className="ai-menu-header">
              <div className="ai-menu-header-left">
                <h3 id="ai-menu-title" className="ai-menu-title">
                  {t('aiMenuTitle')}
                </h3>
                <p className="ai-menu-subtitle">
                  {aiChoice === 'google' ? t('aiMenuGoogleDesc') : t('aiMenuCustomDesc')}
                </p>
              </div>
              <button
                type="button"
                className="ai-menu-close-btn"
                onClick={() => setAiMenuOpen(false)}
                aria-label={t('cancel')}
                title={t('cancel')}
              >
                <svg
                  width="15"
                  height="15"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2.2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <line x1="18" y1="6" x2="6" y2="18" />
                  <line x1="6" y1="6" x2="18" y2="18" />
                </svg>
              </button>
            </div>

            {/* Segmented Provider Tabs */}
            <div className="ai-menu-segmented" role="tablist" aria-label={t('aiMenuTitle')}>
              <button
                type="button"
                role="tab"
                aria-selected={aiChoice === 'custom'}
                className={`ai-menu-tab${aiChoice === 'custom' ? ' active' : ''}`}
                onClick={() => {
                  setAiChoice('custom')
                  setAiMenuError('')
                }}
              >
                <span className="ai-menu-tab-icon" aria-hidden="true">
                  ⚡
                </span>
                <span>{t('aiMenuCustom')}</span>
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={aiChoice === 'google'}
                className={`ai-menu-tab${aiChoice === 'google' ? ' active' : ''}`}
                onClick={() => {
                  setAiChoice('google')
                  setAiMenuError('')
                }}
              >
                <span className="ai-menu-tab-icon ara" aria-hidden="true">
                  ✦
                </span>
                <span>{t('aiMenuGoogle')}</span>
              </button>
            </div>

            {/* Google Gemini Panel */}
            {aiChoice === 'google' && (
              <div className="ai-menu-fields">
                <div className="ai-menu-field-block">
                  <div className="ai-field-label-row">
                    <label htmlFor="ai-menu-google-model" className="ai-field-label">
                      {t('aiMenuModelLabel')}
                    </label>
                  </div>
                  <select
                    id="ai-menu-google-model"
                    value={googleModel}
                    onChange={(event) => setGoogleModel(event.target.value)}
                  >
                    {GEMINI_MODELS.map((model) => (
                      <option key={model} value={model}>
                        {model}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="ai-menu-field-block">
                  <div className="ai-field-label-row">
                    <label htmlFor="ai-menu-google-key" className="ai-field-label">
                      {t('aiMenuKey')}
                    </label>
                    <button
                      type="button"
                      className="ai-menu-studio-link"
                      onClick={() =>
                        void window.aiOfficeWindow.openExternal(
                          'https://aistudio.google.com/apikey',
                        )
                      }
                      rel="noreferrer"
                    >
                      {t('aiMenuStudioLink')} ↗
                    </button>
                  </div>
                  <div className="ai-key-input-wrap">
                    <input
                      id="ai-menu-google-key"
                      type={showAiKey ? 'text' : 'password'}
                      autoComplete="off"
                      value={googleKey}
                      placeholder={
                        hasSavedGoogleKey ? '•••••••••••••••••••• (leave blank to keep)' : 'AIza…'
                      }
                      onChange={(event) => setGoogleKey(event.target.value)}
                    />
                    <button
                      type="button"
                      className="ai-pwd-toggle"
                      onClick={() => setShowAiKey(!showAiKey)}
                      tabIndex={-1}
                      aria-label={showAiKey ? 'Hide key' : 'Show key'}
                    >
                      {showAiKey ? (
                        <svg
                          width="15"
                          height="15"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24" />
                          <line x1="1" y1="1" x2="23" y2="23" />
                        </svg>
                      ) : (
                        <svg
                          width="15"
                          height="15"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8z" />
                          <circle cx="12" cy="12" r="3" />
                        </svg>
                      )}
                    </button>
                  </div>
                  <small className="ai-menu-hint">
                    {hasSavedGoogleKey && !googleKey ? t('aiMenuKeySavedHint') : t('aiMenuKeyHint')}
                  </small>
                </div>
              </div>
            )}

            {/* OpenAI-Compatible Panel */}
            {aiChoice === 'custom' && (
              <div className="ai-menu-fields">
                <div className="ai-menu-field-block">
                  <div className="ai-field-label-row">
                    <label htmlFor="ai-menu-endpoint" className="ai-field-label">
                      {t('aiMenuEndpoint')}
                    </label>
                    <div className="ai-menu-preset-chips">
                      <span className="ai-presets-label">{t('aiMenuPresets')}:</span>
                      {CUSTOM_ENDPOINT_PRESETS.map((preset) => (
                        <button
                          key={preset.id}
                          type="button"
                          className="ai-menu-preset-chip"
                          onClick={() => setAiEndpoint(preset.baseUrl)}
                        >
                          {preset.label}
                        </button>
                      ))}
                    </div>
                  </div>
                  <input
                    id="ai-menu-endpoint"
                    value={aiEndpoint}
                    placeholder="https://api.openai.com/v1"
                    onChange={(event) => setAiEndpoint(event.target.value)}
                  />
                </div>

                <div className="ai-menu-field-block">
                  <div className="ai-field-label-row">
                    <label htmlFor="ai-menu-model" className="ai-field-label">
                      {t('aiMenuModel')}
                    </label>
                    {availableModels.length > 0 && (
                      <span className="ai-models-count-badge">
                        {availableModels.length} {availableModels.length === 1 ? 'model' : 'models'}
                      </span>
                    )}
                  </div>
                  <div className="ai-combobox-wrapper" ref={modelComboboxRef}>
                    <div className="ai-combobox-input-wrap">
                      <input
                        id="ai-menu-model"
                        className="ai-combobox-input"
                        value={aiModelId}
                        placeholder="gpt-4o-mini"
                        autoComplete="off"
                        onClick={() => {
                          if (availableModels.length > 0) setModelDropdownOpen(true)
                        }}
                        onFocus={() => {
                          if (availableModels.length > 0) setModelDropdownOpen(true)
                        }}
                        onChange={(event) => {
                          setAiModelId(event.target.value)
                        }}
                      />
                      <div className="ai-combobox-actions">
                        <button
                          type="button"
                          className="ai-fetch-models-btn"
                          onClick={() => void handleFetchModels()}
                          disabled={isFetchingModels || !aiEndpoint.trim()}
                          title={t('aiMenuFetchModels')}
                        >
                          {isFetchingModels ? (
                            <>
                              <span className="ai-fetch-spinner" aria-hidden="true" />
                              <span>...</span>
                            </>
                          ) : (
                            <span>{t('aiMenuFetchModels')}</span>
                          )}
                        </button>
                        {availableModels.length > 0 && (
                          <button
                            type="button"
                            className="ai-dropdown-arrow-btn"
                            onClick={() => setModelDropdownOpen(!modelDropdownOpen)}
                            aria-label="Toggle model list"
                          >
                            <svg
                              width="12"
                              height="12"
                              viewBox="0 0 24 24"
                              fill="none"
                              stroke="currentColor"
                              strokeWidth="2.5"
                              strokeLinecap="round"
                              strokeLinejoin="round"
                              style={{
                                transform: modelDropdownOpen ? 'rotate(180deg)' : 'rotate(0deg)',
                                transition: 'transform 0.15s ease',
                              }}
                            >
                              <path d="M6 9l6 6 6-6" />
                            </svg>
                          </button>
                        )}
                      </div>
                    </div>

                    {modelDropdownOpen && availableModels.length > 0 && (
                      <div className="ai-custom-dropdown-menu" role="listbox">
                        {availableModels.length > 6 && (
                          <div className="ai-dropdown-search-wrap">
                            <svg
                              className="ai-dropdown-search-icon"
                              width="12"
                              height="12"
                              viewBox="0 0 24 24"
                              fill="none"
                              stroke="currentColor"
                              strokeWidth="2.2"
                              strokeLinecap="round"
                              strokeLinejoin="round"
                            >
                              <circle cx="11" cy="11" r="8" />
                              <line x1="21" y1="21" x2="16.65" y2="16.65" />
                            </svg>
                            <input
                              type="text"
                              className="ai-dropdown-search"
                              placeholder="Filter models..."
                              value={modelFilter}
                              onChange={(e) => setModelFilter(e.target.value)}
                              autoFocus
                            />
                          </div>
                        )}
                        <div className="ai-dropdown-list">
                          {filteredModels.length === 0 ? (
                            <div className="ai-dropdown-empty">No matching models</div>
                          ) : (
                            filteredModels.map((m) => {
                              const isSelected = m === aiModelId
                              return (
                                <div
                                  key={m}
                                  role="option"
                                  aria-selected={isSelected}
                                  className={`ai-dropdown-item${isSelected ? ' selected' : ''}`}
                                  onClick={() => {
                                    setAiModelId(m)
                                    setModelDropdownOpen(false)
                                    setModelFilter('')
                                  }}
                                >
                                  <span className="ai-dropdown-item-name">{m}</span>
                                  {isSelected && (
                                    <svg
                                      className="ai-dropdown-item-check"
                                      width="13"
                                      height="13"
                                      viewBox="0 0 24 24"
                                      fill="none"
                                      stroke="currentColor"
                                      strokeWidth="2.5"
                                      strokeLinecap="round"
                                      strokeLinejoin="round"
                                    >
                                      <polyline points="20 6 9 17 4 12" />
                                    </svg>
                                  )}
                                </div>
                              )
                            })
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                </div>

                <div className="ai-menu-field-block">
                  <div className="ai-field-label-row">
                    <label htmlFor="ai-menu-key" className="ai-field-label">
                      {t('aiMenuKey')}
                    </label>
                    {hasSavedKey && (
                      <span className={`ai-key-status-badge ${aiKey ? 'replacing' : 'saved'}`}>
                        <span className="ai-key-status-dot" aria-hidden="true" />
                        {aiKey ? t('aiMenuKeyReplacing') : t('aiMenuKeySaved')}
                      </span>
                    )}
                  </div>
                  <div className="ai-key-input-wrap">
                    <input
                      id="ai-menu-key"
                      type={showAiKey ? 'text' : 'password'}
                      autoComplete="off"
                      value={aiKey}
                      placeholder={
                        hasSavedKey ? '•••••••••••••••••••• (leave blank to keep)' : 'sk-...'
                      }
                      onChange={(event) => setAiKey(event.target.value)}
                    />
                    <button
                      type="button"
                      className="ai-pwd-toggle"
                      onClick={() => setShowAiKey(!showAiKey)}
                      tabIndex={-1}
                      aria-label={showAiKey ? 'Hide key' : 'Show key'}
                    >
                      {showAiKey ? (
                        <svg
                          width="15"
                          height="15"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24" />
                          <line x1="1" y1="1" x2="23" y2="23" />
                        </svg>
                      ) : (
                        <svg
                          width="15"
                          height="15"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8z" />
                          <circle cx="12" cy="12" r="3" />
                        </svg>
                      )}
                    </button>
                  </div>
                  <small className="ai-menu-hint">
                    {hasSavedKey && !aiKey
                      ? t('aiMenuKeySavedHint')
                      : isLocalhostEndpoint(aiEndpoint)
                        ? t('aiMenuKeyOptional')
                        : t('aiMenuKeyHint')}
                  </small>
                </div>
              </div>
            )}

            {/* Error Message */}
            {aiMenuError && (
              <div className="ai-menu-error-banner" role="alert">
                <svg
                  width="14"
                  height="14"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  aria-hidden="true"
                >
                  <circle cx="12" cy="12" r="10" />
                  <line x1="12" y1="8" x2="12" y2="12" />
                  <line x1="12" y1="16" x2="12.01" y2="16" />
                </svg>
                <span>{aiMenuError}</span>
              </div>
            )}

            {/* Modal Actions */}
            <div className="ai-menu-actions">
              <button
                type="button"
                className="ai-menu-cancel-btn"
                onClick={() => setAiMenuOpen(false)}
              >
                {t('cancel')}
              </button>
              <button
                type="button"
                className="btn btn-primary ai-menu-save-btn"
                onClick={() => void saveAiMenu()}
              >
                <span>{t('aiMenuSave')}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {confirmDelete && (
        <div className="modal-overlay" onClick={() => setConfirmDelete(null)}>
          <div
            className="modal"
            role="dialog"
            aria-modal="true"
            aria-label={t('deleteModalTitle')}
            onClick={(event) => event.stopPropagation()}
          >
            <h3>{t('deleteModalTitle')}</h3>
            <p>
              {confirmDelete.length === 1
                ? t('deleteConfirmOne', { name: fileName(confirmDelete[0]) })
                : t('deleteConfirmMany', { n: confirmDelete.length })}
            </p>
            {confirmDelete.length > 1 && (
              <ul className="modal-file-list">
                {confirmDelete.slice(0, 6).map((p) => (
                  <li key={p}>{fileName(p)}</li>
                ))}
                {confirmDelete.length > 6 && (
                  <li>{t('deleteMoreCount', { n: confirmDelete.length })}</li>
                )}
              </ul>
            )}
            <div className="modal-buttons">
              <button
                className="btn btn-secondary"
                autoFocus
                onClick={() => setConfirmDelete(null)}
              >
                {t('cancel')}
              </button>
              <button className="btn btn-danger" onClick={confirmDeleteNow}>
                {t('delete')}
              </button>
            </div>
          </div>
        </div>
      )}
      {showChangelog && (
        <ChangelogModal
          onClose={() => setShowChangelog(false)}
          updateState={updateState}
          onRestart={() => void window.aiOffice.restartAndInstall()}
        />
      )}

      {showUpdateRestart && (
        <UpdateRestartModal
          onClose={() => setShowUpdateRestart(false)}
          updateState={updateState}
          onRestart={() => void window.aiOffice.restartAndInstall()}
          onOpenChangelog={() => setShowChangelog(true)}
        />
      )}
      {showDonation && <DonationModal onClose={() => setShowDonation(false)} />}
      {showFeedback && <FeedbackModal onClose={() => setShowFeedback(false)} />}
      {showShortcuts && <ShortcutsModal onClose={() => setShowShortcuts(false)} />}
    </div>
  )
}
