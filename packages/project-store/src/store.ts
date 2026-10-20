/**
 * project-store core implementation
 *
 * Storage layout (baseDir = userData/projects/):
 *   index.json
 *   <project-id>/
 *     project.json
 *     chats/
 *       <chat-id>.jsonl
 *
 * Design principles:
 * - No Electron dependency; the userData path is injected by the caller
 * - All write failures return false, never throw (append path)
 * - JSONL parsing is line-by-line tolerant: bad lines are skipped, no crash
 * - seq is maintained by the store layer: auto-incremented on each appendChatMessage
 */

import { createHash } from 'node:crypto'
import {
  appendFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  statSync,
  unlinkSync,
  writeFileSync,
  readdirSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, sep } from 'node:path'
import type {
  ChatMeta,
  ChatMessage,
  ProjectData,
  ProjectIndex,
  ProjectInfo,
  ProjectSummary,
} from './types.js'

// ────────────────────────────────────────────────────────────
// Internal helpers
// ────────────────────────────────────────────────────────────

/** Max stored characters for a single tool input/output field */
const TOOL_FIELD_MAX_CHARS = 16_000

// Sheets opens unsaved work in this private temp directory; it is not a project file.
const TRANSIENT_SHEETS_DIR_PREFIX = `${join(tmpdir(), 'araoffice-sheets')}${sep}`.toLowerCase()

function nowIso(): string {
  return new Date().toISOString()
}

function ensureDir(dir: string): void {
  if (!existsSync(dir)) mkdirSync(dir, { recursive: true })
}

function readJson<T>(filePath: string): T | null {
  try {
    if (!existsSync(filePath)) return null
    return JSON.parse(readFileSync(filePath, 'utf8')) as T
  } catch {
    return null
  }
}

const RETRYABLE_RENAME_CODES: Record<string, true> = {
  EPERM: true,
  EACCES: true,
  EBUSY: true,
}

/** Sync sleep that yields the CPU. This store's public API is sync, so the main thread
 *  still blocks across a retry window — it just no longer spins.
 *  ponytail: sync fs on the Electron main thread; go async + electron-utils renameWithRetry
 *  if project persists ever measurably jank the UI. */
const sleepBuf = new Int32Array(new SharedArrayBuffer(4))
function sleepSync(ms: number): void {
  Atomics.wait(sleepBuf, 0, 0, ms)
}

/** Atomic write: write to .tmp then rename, with retry on Windows file locks */
function writeJson(filePath: string, data: unknown): void {
  ensureDir(dirname(filePath))
  const tmpPath = `${filePath}.tmp`
  const content = JSON.stringify(data, null, 2)
  writeFileSync(tmpPath, content, 'utf8')
  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      renameSync(tmpPath, filePath)
      return
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code ?? ''
      if (!RETRYABLE_RENAME_CODES[code] || attempt === 4) {
        try {
          unlinkSync(tmpPath)
        } catch {
          /* ignore unlink failure */
        }
        if (RETRYABLE_RENAME_CODES[code]) {
          writeFileSync(filePath, content, 'utf8')
          return
        }
        throw err
      }
      sleepSync(25 * 2 ** attempt)
    }
  }
}

/** projectId/chatId must be a single safe path segment: no separators, no traversal, ≤128 chars.
 *  The mandatory leading alnum inherently rejects '.', '..' and any path separator. */
const ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/

function validateId(kind: 'projectId' | 'chatId', id: string): void {
  if (!ID_PATTERN.test(id)) throw new Error(`Invalid ${kind}: ${JSON.stringify(id)}`)
}

/**
 * baseDir:projectId:chatId → current max seq. Module-level on purpose: several ProjectStore
 * instances are constructed over the same on-disk store in one Electron process, and per-instance
 * counters would mint duplicate seq values for the same chat.
 */
const seqCounters = new Map<string, number>()

// ────────────────────────────────────────────────────────────
// ProjectStore class
// ────────────────────────────────────────────────────────────
export class ProjectStore {
  private baseDir: string

  constructor(userDataPath: string) {
    this.baseDir = join(userDataPath, 'projects')
  }

  // ── Path helpers ──────────────────────────────────────────

  private indexPath(): string {
    return join(this.baseDir, 'index.json')
  }

  private projectDir(projectId: string): string {
    validateId('projectId', projectId)
    return join(this.baseDir, projectId)
  }

  private projectJsonPath(projectId: string): string {
    return join(this.projectDir(projectId), 'project.json')
  }

  private chatsDir(projectId: string): string {
    return join(this.projectDir(projectId), 'chats')
  }

  private chatPath(projectId: string, chatId: string): string {
    validateId('chatId', chatId)
    return join(this.chatsDir(projectId), `${chatId}.jsonl`)
  }

  // ── seq counters (in-memory cache, initialized from JSONL line count on first read) ──

  /** baseDir:projectId:chatId → current max seq (shared across instances; see module-level seqCounters) */
  private seqKey(projectId: string, chatId: string): string {
    return `${this.baseDir}:${projectId}:${chatId}`
  }

  private nextSeq(projectId: string, chatId: string): number {
    const key = this.seqKey(projectId, chatId)
    const cur = seqCounters.get(key)
    if (cur !== undefined) {
      const next = cur + 1
      seqCounters.set(key, next)
      return next
    }
    // Initialization: scan the existing file for the max seq. Uncapped: a chat
    // over the default read limit would otherwise re-mint duplicate seq values
    // after every restart.
    const existing = this.loadChat(projectId, chatId, Infinity)
    const maxSeq = existing.reduce((m, msg) => Math.max(m, msg.seq), -1)
    const next = maxSeq + 1
    seqCounters.set(key, next)
    return next
  }

  // ── Index read/write ──────────────────────────────────────

  private readIndex(): ProjectIndex {
    return readJson<ProjectIndex>(this.indexPath()) ?? { projects: [], fileMap: {} }
  }

  private writeIndex(index: ProjectIndex): void {
    ensureDir(this.baseDir)
    writeJson(this.indexPath(), index)
  }

  // ── Project read/write ────────────────────────────────────

  private readProject(projectId: string): ProjectData | null {
    return readJson<ProjectData>(this.projectJsonPath(projectId))
  }

  private writeProject(data: ProjectData): void {
    writeJson(this.projectJsonPath(data.id), data)
  }

  // ────────────────────────────────────────────────────────────
  // Public API
  // ────────────────────────────────────────────────────────────

  /**
   * Ensures the default project exists (id: "default", name: the Chinese "default project" label).
   * Idempotent: returns directly if it already exists.
   */
  ensureDefaultProject(): ProjectData {
    const existing = this.readProject('default')
    if (existing) return existing

    const now = nowIso()
    const data: ProjectData = {
      id: 'default',
      name: 'Default Project',
      createdAt: now,
      updatedAt: now,
      files: [],
    }
    ensureDir(this.projectDir('default'))
    this.writeProject(data)

    const index = this.readIndex()
    if (!index.projects.find((p) => p.id === 'default')) {
      index.projects.unshift({ id: data.id, name: data.name, createdAt: now, updatedAt: now })
      this.writeIndex(index)
    }
    return data
  }

  /**
   * Looks up the projectId by absolute file path.
   * If not found, assign to default and register in fileMap.
   */
  resolveProjectForFile(filePath: string): string {
    this.ensureDefaultProject()
    const index = this.readIndex()
    const existing = index.fileMap[filePath]
    if (existing) return existing

    // Assign to default
    index.fileMap[filePath] = 'default'
    this.writeIndex(index)

    // Update the files list in project.json
    const proj = this.readProject('default')
    if (proj && !proj.files.includes(filePath)) {
      proj.files.push(filePath)
      proj.updatedAt = nowIso()
      this.writeProject(proj)
    }
    return 'default'
  }

  /**
   * Derives a chatId from a file path (first 16 hex chars of sha256).
   * Unsaved files use an externally provided temp id (e.g. "unsaved-<timestamp>").
   * Only a fallback derivation for old data without a mapping; new code uses
   * resolveChatForFile (stable mapping).
   */
  static chatIdForFile(filePath: string): string {
    return createHash('sha256').update(filePath).digest('hex').slice(0, 16)
  }

  /** Gets the chatId from the mapping; falls back to the path hash without registering. */
  chatIdForPath(filePath: string): string {
    const index = this.readIndex()
    return index.chatIdByPath?.[filePath] ?? ProjectStore.chatIdForFile(filePath)
  }

  /**
   * Resolves { projectId, chatId } by file path (the core of the resolveChat IPC).
   * The chatId is registered into chatIdByPath on first resolve; from then on,
   * renaming/moving the file only changes the mapping key — the chatId stays
   * stable and history always follows the file.
   */
  resolveChatForFile(filePath: string): { projectId: string; chatId: string } {
    const projectId = this.resolveProjectForFile(filePath)
    const index = this.readIndex()
    const mapped = index.chatIdByPath?.[filePath]
    if (mapped) return { projectId, chatId: mapped }
    const chatId = ProjectStore.chatIdForFile(filePath)
    index.chatIdByPath = { ...(index.chatIdByPath ?? {}), [filePath]: chatId }
    this.writeIndex(index)
    return { projectId, chatId }
  }

  /**
   * Called after a file is renamed/moved on disk: the keys in fileMap,
   * project.files and chatIdByPath are updated accordingly, while the chatId
   * stays the same (history needs no relocation).
   */
  fileRenamed(oldPath: string, newPath: string): void {
    if (oldPath === newPath) return
    const index = this.readIndex()
    const pid = index.fileMap[oldPath]
    if (pid !== undefined) {
      delete index.fileMap[oldPath]
      index.fileMap[newPath] = pid
      const proj = this.readProject(pid)
      if (proj) {
        proj.files = proj.files.map((f) => (f === oldPath ? newPath : f))
        proj.updatedAt = nowIso()
        this.writeProject(proj)
      }
    }
    // Old data without a mapping: the chatId was derived from the old path hash; register the mapping under that hash on rename so history keeps up
    const chatId = index.chatIdByPath?.[oldPath] ?? ProjectStore.chatIdForFile(oldPath)
    if (index.chatIdByPath?.[oldPath] !== undefined) delete index.chatIdByPath[oldPath]
    index.chatIdByPath = { ...(index.chatIdByPath ?? {}), [newPath]: chatId }
    this.writeIndex(index)
  }

  /**
   * Buffer for the opening messages of a chat that has no file yet: the file
   * is not created until the first assistant reply arrives, so aborted or
   * failed requests never leave behind an empty record with a lone user
   * message.
   */
  private readonly pendingFirstWrite = new Map<string, ChatMessage[]>()

  /** Flushes buffered opening messages to disk (materialized before rebind: once the file is saved, the opening messages should be kept). */
  private flushPending(projectId: string, chatId: string): boolean {
    const key = this.seqKey(projectId, chatId)
    const buf = this.pendingFirstWrite.get(key)
    if (!buf || buf.length === 0) return true
    try {
      ensureDir(this.chatsDir(projectId))
      const lines = buf.map((r) => JSON.stringify(r) + '\n').join('')
      appendFileSync(this.chatPath(projectId, chatId), lines, 'utf8')
      this.pendingFirstWrite.delete(key)
      return true
    } catch (err) {
      console.warn('[project-store] flushPending failed:', err)
      return false
    }
  }

  /**
   * Appends one message to the JSONL. Returns true when the message was
   * persisted or intentionally buffered; false when persistence failed.
   * seq is auto-assigned by the store layer (monotonically increasing).
   * When the record file doesn't exist yet, non-assistant messages are buffered;
   * the file is created and flushed only when the first assistant message arrives.
   */
  appendChatMessage(
    projectId: string,
    chatId: string,
    msg: Omit<ChatMessage, 'seq' | 'ts'> & { ts?: string },
  ): boolean {
    try {
      const seq = this.nextSeq(projectId, chatId)
      const ts = msg.ts ?? nowIso()
      const record: ChatMessage = { seq, ts, role: msg.role, text: msg.text }
      if (msg.fileRef !== undefined) record.fileRef = msg.fileRef
      if (msg.tools && msg.tools.length > 0) {
        // Truncate tool inputs/outputs so one JSONL line can't blow up on a huge payload
        record.tools = msg.tools.map((t) => ({
          ...t,
          ...(t.input !== undefined ? { input: t.input.slice(0, TOOL_FIELD_MAX_CHARS) } : {}),
          ...(t.output !== undefined ? { output: t.output.slice(0, TOOL_FIELD_MAX_CHARS) } : {}),
        }))
      }
      if (msg.attachments !== undefined) record.attachments = msg.attachments

      const key = this.seqKey(projectId, chatId)
      if (msg.role !== 'assistant' && !existsSync(this.chatPath(projectId, chatId))) {
        const buf = this.pendingFirstWrite.get(key) ?? []
        buf.push(record)
        this.pendingFirstWrite.set(key, buf)
        return true
      }
      ensureDir(this.chatsDir(projectId))
      const buf = this.pendingFirstWrite.get(key) ?? []
      const lines = [...buf, record].map((r) => JSON.stringify(r) + '\n').join('')
      appendFileSync(this.chatPath(projectId, chatId), lines, 'utf8')
      this.pendingFirstWrite.delete(key)
      return true
    } catch (err) {
      console.warn('[project-store] appendChatMessage failed:', err)
      return false
    }
  }

  /**
   * Reads the most recent `limit` messages (in ascending seq order).
   * A bad JSONL line is skipped without crashing.
   */
  loadChat(projectId: string, chatId: string, limit = 200): ChatMessage[] {
    const pending = this.pendingFirstWrite.get(this.seqKey(projectId, chatId)) ?? []
    const filePath = this.chatPath(projectId, chatId)
    const messages: ChatMessage[] = [...pending]
    try {
      if (existsSync(filePath)) {
        const raw = readFileSync(filePath, 'utf8')
        const lines = raw.split('\n').filter((l) => l.trim())
        for (const line of lines) {
          try {
            const msg = JSON.parse(line) as ChatMessage
            if (
              typeof msg.seq === 'number' &&
              typeof msg.role === 'string' &&
              typeof msg.text === 'string'
            ) {
              messages.push(msg)
            }
          } catch {
            // skip bad lines
          }
        }
      }
      // Sort by seq and take the most recent `limit` entries
      messages.sort((a, b) => a.seq - b.seq)
      return messages.slice(-limit)
    } catch {
      return messages
    }
  }

  /**
   * Lists metadata of all chats in a project.
   */
  listChats(projectId: string): ChatMeta[] {
    const dir = this.chatsDir(projectId)
    if (!existsSync(dir)) return []
    try {
      const files = readdirSync(dir).filter((f) => f.endsWith('.jsonl'))
      return files.map((f) => {
        const chatId = f.replace(/\.jsonl$/, '')
        const fullPath = join(dir, f)
        let updatedAt = nowIso()
        let approxCount = 0
        try {
          const st = statSync(fullPath)
          updatedAt = st.mtime.toISOString()
          // Estimate: assume an average of 120 bytes per line
          approxCount = Math.round(st.size / 120)
        } catch {
          // use defaults if stat fails
        }
        return { chatId, updatedAt, approxCount }
      })
    } catch {
      return []
    }
  }

  /**
   * Moves chats/<fromId>.jsonl to chats/<toId>.jsonl (possibly across projects).
   * If the target exists, don't overwrite: renumber the source messages' seq and
   * append them at the target's end (old conversations of a same-named file are kept).
   */
  private renameOrMergeChat(
    fromProjectId: string,
    fromId: string,
    toProjectId: string,
    toId: string,
  ): void {
    if (fromProjectId === toProjectId && fromId === toId) return
    // The source may still have buffered opening messages: materialize them first (once the file is saved, they should be kept)
    if (!this.flushPending(fromProjectId, fromId)) return
    const oldPath = this.chatPath(fromProjectId, fromId)
    const newPath = this.chatPath(toProjectId, toId)
    let mergedMaxSeq: number | undefined
    try {
      if (existsSync(oldPath)) {
        ensureDir(dirname(newPath))
        if (!existsSync(newPath)) {
          renameSync(oldPath, newPath)
        } else {
          const existing = this.loadChat(toProjectId, toId, 10_000)
          let seq = existing.reduce((m, msg) => Math.max(m, msg.seq), -1) + 1
          const moved = this.loadChat(fromProjectId, fromId, 10_000)
          const lines = moved.map((m) => JSON.stringify({ ...m, seq: seq++ }) + '\n').join('')
          if (lines) appendFileSync(newPath, lines, 'utf8')
          unlinkSync(oldPath)
          mergedMaxSeq = seq - 1
        }
      }
    } catch (err) {
      console.warn('[project-store] rebindChat rename failed:', err)
    }

    // Migrate the seq counter (when merged, the renumbered max seq wins)
    const oldKey = this.seqKey(fromProjectId, fromId)
    const curSeq = seqCounters.get(oldKey)
    seqCounters.delete(oldKey)
    seqCounters.delete(this.seqKey(toProjectId, toId))
    const next = mergedMaxSeq ?? curSeq
    if (next !== undefined) {
      seqCounters.set(this.seqKey(toProjectId, toId), next)
    }
  }

  /**
   * Renames the JSONL file (called after an unsaved file is first written to disk):
   * chats/<tempId>.jsonl → chats/<newChatId>.jsonl (if the target exists, merge by continuing seq).
   */
  rebindChat(projectId: string, tempId: string, newChatId: string): void {
    this.renameOrMergeChat(projectId, tempId, projectId, newChatId)
  }

  /**
   * After an unsaved session first gets a real file path: register in fileMap,
   * compute the chatId from the path, and move the temp JSONL into the target project.
   * Returns the new { projectId, chatId } for the renderer to update its references.
   */
  rebindChatToFile(
    tempProjectId: string,
    tempChatId: string,
    filePath: string,
  ): { projectId: string; chatId: string } {
    const { projectId, chatId } = this.resolveChatForFile(filePath)
    this.renameOrMergeChat(tempProjectId, tempChatId, projectId, chatId)
    return { projectId, chatId }
  }

  /**
   * Gets project info.
   */
  getProject(projectId: string): ProjectData | null {
    return this.readProject(projectId)
  }

  /**
   * Lists all projects.
   */
  listProjects(): ProjectInfo[] {
    return this.readIndex().projects
  }

  // ── P1 extended API ────────────────────────────────────────

  /**
   * Lists all projects (with file count + last active time).
   */
  listProjectsSummary(): ProjectSummary[] {
    this.ensureDefaultProject()
    const index = this.readIndex()
    return index.projects.map((info) => {
      const fileCount = this.listProjectFiles(info.id).length
      // Take the max of project.json updatedAt and all chat mtimes
      let lastActiveAt = info.updatedAt
      const chats = this.listChats(info.id)
      for (const c of chats) {
        if (c.updatedAt > lastActiveAt) lastActiveAt = c.updatedAt
      }
      return {
        ...info,
        fileCount,
        lastActiveAt,
        isDefault: info.id === 'default',
      }
    })
  }

  /**
   * Lists files that currently exist for a project. Stored paths are historical
   * records and may outlive files deleted or moved outside ARAOffice.
   */
  listProjectFiles(projectId: string): string[] {
    const proj = this.readProject(projectId)
    if (!proj) return []
    return [...new Set(proj.files)].filter(
      (filePath) =>
        existsSync(filePath) && !filePath.toLowerCase().startsWith(TRANSIENT_SHEETS_DIR_PREFIX),
    )
  }

  /**
   * Creates a project (name must be non-empty; id is the first 12 hex chars of
   * sha256(name) plus a timestamp suffix to avoid collisions).
   * Returns the newly created ProjectData.
   */
  createProject(name: string): ProjectData {
    const trimmed = name.trim()
    if (!trimmed) throw new Error('Project name cannot be empty')
    const now = nowIso()
    // Generate a stable yet unique id
    const hash = createHash('sha256')
      .update(trimmed + now)
      .digest('hex')
      .slice(0, 12)
    const id = `proj-${hash}`
    const data: ProjectData = {
      id,
      name: trimmed,
      createdAt: now,
      updatedAt: now,
      files: [],
    }
    ensureDir(this.projectDir(id))
    this.writeProject(data)
    const index = this.readIndex()
    // Append at the end (default always stays first)
    index.projects.push({ id, name: trimmed, createdAt: now, updatedAt: now })
    this.writeIndex(index)
    return data
  }

  /**
   * Renames a project (the default project cannot be renamed).
   */
  renameProject(id: string, name: string): void {
    if (id === 'default') throw new Error('The default project cannot be renamed')
    const trimmed = name.trim()
    if (!trimmed) throw new Error('Project name cannot be empty')
    const now = nowIso()
    const proj = this.readProject(id)
    if (!proj) throw new Error(`Project does not exist: ${id}`)
    proj.name = trimmed
    proj.updatedAt = now
    this.writeProject(proj)
    const index = this.readIndex()
    const entry = index.projects.find((p) => p.id === id)
    if (entry) {
      entry.name = trimmed
      entry.updatedAt = now
    }
    this.writeIndex(index)
  }

  /**
   * Soft-deletes a project:
   * 1. Move the directory into projects/.trash/<id>-<ts>/
   * 2. Reassign all of its files in fileMap back to default
   * 3. Remove the project from index.projects
   * The default project cannot be deleted.
   */
  deleteProject(id: string): void {
    if (id === 'default') throw new Error('The default project cannot be deleted')
    const proj = this.readProject(id)
    if (!proj) throw new Error(`Project does not exist: ${id}`)

    // 1. Soft-delete the directory
    const src = this.projectDir(id)
    const ts = Date.now()
    const trashDir = join(this.baseDir, '.trash')
    ensureDir(trashDir)
    const dst = join(trashDir, `${id}-${ts}`)
    try {
      if (existsSync(src)) renameSync(src, dst)
    } catch (err) {
      console.warn('[project-store] deleteProject rename to trash failed:', err)
      return
    }

    // 2. Reassign this project's files in fileMap back to default
    this.ensureDefaultProject()
    const index = this.readIndex()
    const movedFiles: string[] = []
    for (const [filePath, pid] of Object.entries(index.fileMap)) {
      if (pid === id) {
        index.fileMap[filePath] = 'default'
        movedFiles.push(filePath)
      }
    }
    // Update the default project.json
    if (movedFiles.length > 0) {
      const defaultProj = this.readProject('default')
      if (defaultProj) {
        for (const f of movedFiles) {
          if (!defaultProj.files.includes(f)) defaultProj.files.push(f)
        }
        defaultProj.updatedAt = nowIso()
        this.writeProject(defaultProj)
      }
    }

    // 3. Remove the index.projects entry
    index.projects = index.projects.filter((p) => p.id !== id)
    this.writeIndex(index)
  }

  /**
   * Moves a file and its corresponding chat into a target project:
   * 1. Materialize and move the chat's JSONL file
   * 2. Update fileMap and the files lists in both project.json files
   */
  moveFileToProject(filePath: string, targetProjectId: string): void {
    this.ensureDefaultProject()
    const index = this.readIndex()
    const fromProjectId = index.fileMap[filePath] ?? 'default'

    if (fromProjectId === targetProjectId) return // nothing to move

    // The target project must exist
    const targetProj = this.readProject(targetProjectId)
    if (!targetProj) throw new Error(`Target project does not exist: ${targetProjectId}`)

    // 1. Materialize and move the corresponding chat before changing metadata.
    const chatId = this.chatIdForPath(filePath)
    if (!this.flushPending(fromProjectId, chatId)) return
    const srcChatPath = this.chatPath(fromProjectId, chatId)
    const dstChatPath = this.chatPath(targetProjectId, chatId)
    try {
      if (existsSync(srcChatPath)) {
        ensureDir(this.chatsDir(targetProjectId))
        renameSync(srcChatPath, dstChatPath)
      }
    } catch (err) {
      console.warn('[project-store] moveFileToProject chat rename failed:', err)
      return
    }

    // 2. Update fileMap
    index.fileMap[filePath] = targetProjectId
    this.writeIndex(index)

    // 3. Update fromProject.files
    const fromProj = this.readProject(fromProjectId)
    if (fromProj) {
      fromProj.files = fromProj.files.filter((f) => f !== filePath)
      fromProj.updatedAt = nowIso()
      this.writeProject(fromProj)
    }

    // 4. Update targetProject.files
    if (!targetProj.files.includes(filePath)) targetProj.files.push(filePath)
    targetProj.updatedAt = nowIso()
    this.writeProject(targetProj)

    // 5. Migrate the seq counter cache after the chat move succeeds.
    const oldKey = this.seqKey(fromProjectId, chatId)
    const newKey = this.seqKey(targetProjectId, chatId)
    const cur = seqCounters.get(oldKey)
    seqCounters.delete(oldKey)
    if (cur !== undefined) seqCounters.set(newKey, cur)
  }
}
