/**
 * AI IPC for the slides main process, extracted from slides-main.ts:
 * settings persistence, the streaming proxy (main process does the networking
 * to avoid renderer CORS), search tools, and the slides-only ai:* channels
 * (image generation, media analysis, style templates).
 */
import { app, ipcMain, webContents } from 'electron'
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  AiTimeoutError,
  activeProviderConfig,
  defaultAiSettings,
  fetchCustomModels,
  mergeAiSettings,
  redactAiSettings,
  resolveAiSettings,
  streamForProvider,
  type AiSettings,
  type AiStreamChunk,
  type AiStreamRequest,
} from '@araoffice/ai-provider'
import { webExtract, webSearch } from '@araoffice/ai-search'
import { tm } from './i18n-main'
import { sessions } from './session-state'

// ---- AI settings + streaming proxy (the main process does the networking to avoid renderer CORS; implementation shared via @araoffice/ai-provider) ----

const AI_SETTINGS_PATH = () => join(app.getPath('userData'), 'ai-settings.json')

function readJson<T>(path: string, fallback: T): T {
  try {
    if (existsSync(path)) return JSON.parse(readFileSync(path, 'utf-8')) as T
  } catch {
    /* Corrupted state file: fall back to defaults */
  }
  return fallback
}

function writeJson(path: string, value: unknown): void {
  mkdirSync(join(path, '..'), { recursive: true })
  writeFileSync(path, JSON.stringify(value, null, 2))
}

function storedSettings(): Partial<AiSettings> {
  return readJson<Partial<AiSettings>>(AI_SETTINGS_PATH(), {})
}

const activeAiStreams = new Map<string, AbortController>()
const WEB_FETCH_MAX_CHARS = 24_000
const WEB_FETCH_MODE = 'full' as const

let aiIpcRegistered = false

export function registerAiIpc(): void {
  if (aiIpcRegistered) return
  aiIpcRegistered = true

  ipcMain.handle('ai:get-settings', (): AiSettings => {
    const stored = readJson<Partial<AiSettings>>(AI_SETTINGS_PATH(), {})
    const settings = resolveAiSettings(stored, defaultAiSettings())
    return redactAiSettings(settings)
  })

  ipcMain.handle('ai:set-settings', (_event, settings: AiSettings) => {
    const stored = readJson<Partial<AiSettings>>(AI_SETTINGS_PATH(), {})
    // The renderer only ever sees redacted settings, so an untouched key arrives
    // as ''. mergeAiSettings keeps the stored key per provider in that case.
    const normalized = mergeAiSettings(
      resolveAiSettings(stored, defaultAiSettings()),
      settings as Partial<AiSettings>,
    )
    writeJson(AI_SETTINGS_PATH(), normalized)
    const redacted = redactAiSettings(normalized)
    for (const wc of webContents.getAllWebContents()) {
      if (!wc.isDestroyed()) {
        try {
          wc.send('ai:settings-changed', redacted)
        } catch {
          /* frame disposed */
        }
      }
    }
  })

  ipcMain.handle(
    'ai:fetch-custom-models',
    async (
      _event,
      override?: { baseUrl?: string; apiKey?: string },
    ): Promise<{ ok: boolean; models: string[]; error?: string }> => {
      const stored = readJson<Partial<AiSettings>>(AI_SETTINGS_PATH(), {})
      const settings = resolveAiSettings(stored, defaultAiSettings())
      const custom = settings.providers.custom
      const baseUrl = override?.baseUrl?.trim() || custom.baseUrl
      const apiKey = override?.apiKey?.trim() || custom.apiKey
      if (!baseUrl) return { ok: true, models: [] }
      try {
        const models = await fetchCustomModels({ model: custom.model, baseUrl, apiKey })
        return { ok: true, models }
      } catch (err) {
        return { ok: false, models: [], error: String(err) }
      }
    },
  )

  const handleCancelStream = (_event: Electron.IpcMainInvokeEvent, requestId: unknown) => {
    if (typeof requestId === 'string') {
      const controller = activeAiStreams.get(requestId)
      if (controller) {
        controller.abort()
        activeAiStreams.delete(requestId)
      }
    }
    return { ok: true }
  }
  ipcMain.handle('ai:cancel', handleCancelStream)
  ipcMain.handle('ai:stream-cancel', handleCancelStream)

  ipcMain.handle('ai:stream', async (event, request: AiStreamRequest) => {
    const { requestId, system, messages } = request
    const tools = request.tools ?? []
    const maxTokens = request.maxTokens ?? 8192
    // The renderer holds redacted settings only (no apiKey), so the stored copy
    // is the single source of truth for provider + credentials.
    const normalized = resolveAiSettings(storedSettings(), defaultAiSettings())
    const { provider, config } = activeProviderConfig(normalized)
    const send = (chunk: AiStreamChunk) => {
      if (!event.sender.isDestroyed()) event.sender.send('ai:stream-chunk', chunk)
    }
    if (!config.model) {
      send({ requestId, type: 'error', error: tm('errNoModel') })
      return
    }
    const controller = new AbortController()
    activeAiStreams.set(requestId, controller)
    // wire-activity keepalive: lets the renderer's silence watchdog tell a slow turn from a dead one
    let lastPing = 0
    const ping = () => {
      const now = Date.now()
      if (now - lastPing < 5_000) return
      lastPing = now
      send({ requestId, type: 'ping' })
    }
    try {
      let stopReason: string | undefined
      await streamForProvider(provider, config, system, messages, tools, maxTokens, {
        signal: controller.signal,
        onDelta: (text) => send({ requestId, type: 'delta', text }),
        onThinking: (text) => send({ requestId, type: 'thinking', text }),
        onToolCall: (toolCall) => send({ requestId, type: 'tool-call', toolCall }),
        onActivity: ping,
        onStopReason: (reason) => {
          stopReason = reason
        },
      })
      send({ requestId, type: 'done', stopReason })
    } catch (err) {
      if (controller.signal.aborted) {
        send({ requestId, type: 'done' })
      } else {
        const msg = err instanceof Error ? err.message : String(err)
        console.error(`[ai-stream] ${requestId} (${provider}/${config.model}) failed:`, msg)
        send({
          requestId,
          type: 'error',
          error: msg,
          ...(err instanceof AiTimeoutError ? { errorCode: 'timeout' as const } : {}),
        })
      }
    } finally {
      activeAiStreams.delete(requestId)
    }
  })

  // Shared textual web search.
  ipcMain.handle('ai:web-search', async (_event, query: string, maxResults?: number) => {
    try {
      return await webSearch(String(query), typeof maxResults === 'number' ? maxResults : 6)
    } catch (err) {
      return { results: [], method: 'error', error: String(err) }
    }
  })
  ipcMain.handle('ai:web-fetch', async (_event, url: unknown) => {
    const input = typeof url === 'string' ? url : ''
    try {
      const result = await webExtract(input, WEB_FETCH_MAX_CHARS, WEB_FETCH_MODE)
      return {
        title: result.title ?? '',
        url: result.url,
        content: result.content ?? '',
      }
    } catch {
      return { title: '', url: input, content: '', isError: true }
    }
  })
}

// ── ai:* handlers unique to slides ──────────────────────────────────────
// Must be registered inside registerSlidesIpc (not registerAiIpc): in shell aggregate mode the
// generic ai:* channels are registered by docs-main.registerAiIpc, and slides' registerAiIpc is
// never called; docs does not have these channels, so putting them in the wrong place raises
// "No handler registered".
export function registerSlidesOnlyAiIpc(): void {
  // ── Style Skill sidecar persistence: write a same-named .styleskill.json next to the draft (fail-open)
  ipcMain.handle(
    'ai:save-sidecar',
    async (
      event,
      data: { topic: string; styleSkill: string; createdAt: string },
    ): Promise<{ ok: boolean }> => {
      try {
        const session = sessions.get(event.sender.id)
        const draftPath = session?.path
        if (!draftPath || !draftPath.endsWith('.pptx')) return { ok: false }
        const sidecarPath = draftPath.replace(/\.pptx$/i, '.styleskill.json')
        writeFileSync(sidecarPath, JSON.stringify(data, null, 2))
        return { ok: true }
      } catch {
        return { ok: false }
      }
    },
  )

  // ── Style template save: stored in userData/style-templates/<name>.json
  const STYLE_TEMPLATES_DIR = () => join(app.getPath('userData'), 'style-templates')

  ipcMain.handle(
    'ai:save-style-template',
    (
      _event,
      name: string,
      data: { topic: string; styleSkill: string; createdAt: string },
    ): { ok: boolean; error?: string } => {
      try {
        const dir = STYLE_TEMPLATES_DIR()
        mkdirSync(dir, { recursive: true })
        // Filename: replace illegal characters in the name with _ then truncate to 64 chars
        const safeName = name.replace(/[/\\:*?"<>|]/g, '_').slice(0, 64)
        if (!safeName) return { ok: false, error: tm('errTplNameInvalid') }
        writeJson(join(dir, `${safeName}.json`), { ...data, name: safeName })
        return { ok: true }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  // ── Style template list
  ipcMain.handle(
    'ai:list-style-templates',
    (): Array<{ name: string; topic: string; createdAt: string }> => {
      try {
        const dir = STYLE_TEMPLATES_DIR()
        if (!existsSync(dir)) return []
        const files = readdirSync(dir).filter((f) => f.endsWith('.json'))
        return files
          .map((f) => {
            try {
              const raw = readJson<{
                name?: string
                topic?: string
                createdAt?: string
                styleSkill?: string
              }>(join(dir, f), {})
              return {
                name: raw.name ?? f.replace(/\.json$/, ''),
                topic: raw.topic ?? '',
                createdAt: raw.createdAt ?? '',
              }
            } catch {
              return null
            }
          })
          .filter(Boolean) as Array<{ name: string; topic: string; createdAt: string }>
      } catch {
        return []
      }
    },
  )

  // ── Style template load
  ipcMain.handle(
    'ai:load-style-template',
    (
      _event,
      name: string,
    ): { ok: boolean; styleSkill?: string; topic?: string; error?: string } => {
      try {
        const dir = STYLE_TEMPLATES_DIR()
        const safeName = name.replace(/[/\\:*?"<>|]/g, '_').slice(0, 64)
        const filePath = join(dir, `${safeName}.json`)
        if (!existsSync(filePath)) return { ok: false, error: tm('errTplMissing', { name }) }
        const raw = readJson<{ styleSkill?: string; topic?: string }>(filePath, {})
        if (!raw.styleSkill) return { ok: false, error: tm('errTplNoSkill', { name }) }
        return { ok: true, styleSkill: raw.styleSkill, topic: raw.topic ?? '' }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )
}
