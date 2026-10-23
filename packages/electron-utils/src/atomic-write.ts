import { randomBytes } from 'node:crypto'
import { rename, unlink, writeFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

/** Transient Windows error codes: antivirus/indexer briefly locks the rename target. */
export const RETRYABLE_RENAME_CODES: Record<string, true> = {
  EPERM: true,
  EACCES: true,
  EBUSY: true,
}
const RENAME_RETRIES = 4

/**
 * Rename with exponential backoff on transient Windows file lock errors (Defender/indexer).
 */
export async function renameWithRetry(
  sourcePath: string,
  targetPath: string,
  retries = RENAME_RETRIES,
): Promise<void> {
  for (let attempt = 0; ; attempt++) {
    try {
      await rename(sourcePath, targetPath)
      return
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code ?? ''
      if (!RETRYABLE_RENAME_CODES[code] || attempt >= retries) throw err
      await new Promise((resolve) => setTimeout(resolve, 50 * 2 ** attempt))
    }
  }
}

/**
 * Same-dir temp file + rename, so a crash mid-write can't truncate the target.
 * Rename-over-existing can fail under Windows file locks; after retries the
 * error is returned and the original target remains untouched.
 */
export async function atomicWriteFile(filePath: string, data: Buffer | Uint8Array): Promise<void> {
  const tmp = join(
    dirname(filePath),
    `.${basename(filePath)}.${randomBytes(6).toString('hex')}.tmp`,
  )
  try {
    await writeFile(tmp, data)
    await renameWithRetry(tmp, filePath)
  } catch (err) {
    try {
      await unlink(tmp)
    } catch {
      /* tmp never created */
    }
    throw err
  }
}
