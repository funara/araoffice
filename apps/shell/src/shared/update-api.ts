// Contracts between the shell main-process auto-updater and the renderer
// surfaces: AppUpdateState / ChangelogEntry for Home, the changelog modal and
// the download-and-restart prompt; UpdateChannel for settings.
type UpdateStatus = 'idle' | 'checking' | 'available' | 'downloading' | 'downloaded' | 'error'

export interface AppUpdateState {
  status: UpdateStatus
  currentVersion: string
  availableVersion?: string
  percent?: number
  error?: string
  releaseNotes?: string
}

export interface ChangelogEntry {
  version: string
  title: string
  body: string
  publishedAt: string
  htmlUrl?: string
}

/** user-selectable update channel; 'stable' maps to the latest.yml feed, 'beta' to beta.yml */
export type UpdateChannel = 'stable' | 'beta'

export function isUpdateChannel(v: unknown): v is UpdateChannel {
  return v === 'stable' || v === 'beta'
}
