import { mkdir, readFile, rm, stat } from 'node:fs/promises'
import { setTimeout as sleep } from 'node:timers/promises'
import { atomicWriteFile } from '@araoffice/electron-utils'
/**
 * Helpers for userData/app-settings.json — a flat JSON object holding
 * cross-module app preferences (UI language, first-run onboarding flag).
 * Editor modules read the same file, so the shape must stay a plain object.
 */

export async function readAppSettings(path: string): Promise<Record<string, unknown>> {
  try {
    const raw: unknown = JSON.parse(await readFile(path, 'utf8'))
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) {
      return raw as Record<string, unknown>
    }
  } catch {
    // missing or corrupt file: treat as empty settings
  }
  return {}
}

const LOCK_RETRY_MS = 10
const LOCK_STALE_MS = 30_000

async function withAppSettingsLock<T>(path: string, fn: () => Promise<T>): Promise<T> {
  const lockPath = `${path}.lock`
  let locked = false
  while (!locked) {
    try {
      await mkdir(lockPath)
      locked = true
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'EEXIST') return fn()
      try {
        if (Date.now() - (await stat(lockPath)).mtimeMs > LOCK_STALE_MS) {
          await rm(lockPath, { recursive: true, force: true })
        }
      } catch {
        // lock vanished or is unreadable; retry the acquire
      }
      await sleep(LOCK_RETRY_MS)
    }
  }
  try {
    return await fn()
  } finally {
    try {
      await rm(lockPath, { recursive: true, force: true })
    } catch {
      /* ignore lock cleanup errors */
    }
  }
}

async function writeSettingsUnlocked(
  path: string,
  settings: Record<string, unknown>,
): Promise<void> {
  await atomicWriteFile(path, Buffer.from(JSON.stringify(settings, null, 2), 'utf8'))
}

async function writeAppSettingUnlocked(path: string, key: string, value: unknown): Promise<void> {
  const settings = { ...(await readAppSettings(path)), [key]: value }
  await writeSettingsUnlocked(path, settings)
}

export async function writeAppSetting(path: string, key: string, value: unknown): Promise<void> {
  await withAppSettingsLock(path, () => writeAppSettingUnlocked(path, key, value))
}

export async function updateAppSetting(
  path: string,
  key: string,
  update: (current: unknown) => unknown,
): Promise<void> {
  return withAppSettingsLock(path, async () => {
    const settings = await readAppSettings(path)
    settings[key] = update(settings[key])
    await writeSettingsUnlocked(path, settings)
  })
}
