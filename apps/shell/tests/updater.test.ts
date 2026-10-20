import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

const appState = { isPackaged: true }
const mockWebContents = {
  isDestroyed: () => false,
  send: vi.fn(),
}

vi.mock('electron', () => ({
  app: {
    get isPackaged() {
      return appState.isPackaged
    },
    getVersion: () => '0.1.0',
    getPath: () => 'C:/Users/test/AppData/Roaming/ARAOffice',
  },
  webContents: {
    getAllWebContents: () => [mockWebContents],
  },
}))

vi.mock('../src/main/app-settings', () => ({
  readAppSettings: () => ({ lastSeenAppVersion: '0.0.9' }),
  writeAppSetting: vi.fn(),
}))

type Listener = (...args: unknown[]) => unknown

const updaterState = {
  listeners: new Map<string, Listener>(),
  autoDownload: true,
  autoInstallOnAppQuit: false,
  disableDifferentialDownload: false,
  channel: null as string | null,
  allowDowngrade: false,
}
const checkForUpdatesMock = vi.fn(() => Promise.resolve(null))
const downloadUpdateMock = vi.fn<() => Promise<unknown>>(() => Promise.resolve([]))
const quitAndInstallMock = vi.fn()

vi.mock('electron-updater', () => ({
  autoUpdater: {
    on: (event: string, listener: Listener) => {
      updaterState.listeners.set(event, listener)
    },
    checkForUpdates: () => checkForUpdatesMock(),
    downloadUpdate: () => downloadUpdateMock(),
    quitAndInstall: (...args: unknown[]) => quitAndInstallMock(...(args as [boolean, boolean])),
    set autoDownload(v: boolean) {
      updaterState.autoDownload = v
    },
    get autoDownload() {
      return updaterState.autoDownload
    },
    set autoInstallOnAppQuit(v: boolean) {
      updaterState.autoInstallOnAppQuit = v
    },
    get autoInstallOnAppQuit() {
      return updaterState.autoInstallOnAppQuit
    },
    set disableDifferentialDownload(v: boolean) {
      updaterState.disableDifferentialDownload = v
    },
    get disableDifferentialDownload() {
      return updaterState.disableDifferentialDownload
    },
    set channel(v: string | null) {
      updaterState.channel = v
      updaterState.allowDowngrade = true
    },
    get channel() {
      return updaterState.channel
    },
    set allowDowngrade(v: boolean) {
      updaterState.allowDowngrade = v
    },
    get allowDowngrade() {
      return updaterState.allowDowngrade
    },
  },
}))

const FIRST_CHECK_DELAY_MS = 5_000
const RECHECK_INTERVAL_MS = 4 * 60 * 60 * 1000

let platformSpy: { restore: () => void } | null = null

function setPlatform(platform: string): void {
  const original = Object.getOwnPropertyDescriptor(process, 'platform')!
  Object.defineProperty(process, 'platform', { value: platform })
  platformSpy = {
    restore: () => Object.defineProperty(process, 'platform', original),
  }
}

async function loadUpdater() {
  return import('../src/main/updater')
}

beforeEach(() => {
  vi.resetModules()
  vi.useFakeTimers()
  appState.isPackaged = true
  delete process.env.ARAOFFICE_FAKE_UPDATE
  delete process.env.ARAOFFICE_FAKE_JUST_UPDATED
  delete process.env.ARA_GITHUB_OWNER
  delete process.env.ARA_GITHUB_REPO
  updaterState.listeners.clear()
  updaterState.autoDownload = true
  updaterState.autoInstallOnAppQuit = false
  updaterState.disableDifferentialDownload = false
  updaterState.channel = null
  updaterState.allowDowngrade = false
  checkForUpdatesMock.mockClear()
  downloadUpdateMock.mockReset()
  quitAndInstallMock.mockClear()
  mockWebContents.send.mockClear()
  setPlatform('darwin')
})

afterEach(() => {
  vi.useRealTimers()
  platformSpy?.restore()
  platformSpy = null
  delete process.env.ARAOFFICE_FAKE_UPDATE
  delete process.env.ARAOFFICE_FAKE_JUST_UPDATED
  delete process.env.ARA_GITHUB_OWNER
  delete process.env.ARA_GITHUB_REPO
})

describe('initAutoUpdater', () => {
  it('does nothing in unpacked (dev) runs without the fake-update env', async () => {
    appState.isPackaged = false
    const { initAutoUpdater } = await loadUpdater()
    await initAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(updaterState.listeners.size).toBe(0)
    expect(checkForUpdatesMock).not.toHaveBeenCalled()
  })

  it('surfaces a manual download state on unsupported platforms', async () => {
    platformSpy?.restore()
    setPlatform('linux')
    process.env.ARA_GITHUB_OWNER = 'example-owner'
    process.env.ARA_GITHUB_REPO = 'example-repo'
    const { checkForUpdates, getAppUpdateState, initAutoUpdater } = await loadUpdater()
    await initAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(updaterState.listeners.size).toBe(0)
    expect(checkForUpdatesMock).not.toHaveBeenCalled()
    // auto-update cannot run here; the state must be explicit and point at the releases page
    expect(getAppUpdateState().status).toBe('error')
    expect(getAppUpdateState().error).toContain(
      'https://github.com/example-owner/example-repo/releases',
    )
    // a manual check re-asserts the notice instead of faking a check or reporting up-to-date
    checkForUpdates()
    expect(getAppUpdateState().status).toBe('error')
    expect(checkForUpdatesMock).not.toHaveBeenCalled()
    expect(updaterState.listeners.size).toBe(0)
  })

  it('omits the releases link when no release repo is configured', async () => {
    platformSpy?.restore()
    setPlatform('linux')
    process.env.ARA_GITHUB_OWNER = ''
    process.env.ARA_GITHUB_REPO = ''
    const { getAppUpdateState, initAutoUpdater } = await loadUpdater()
    await initAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(getAppUpdateState().status).toBe('error')
    expect(getAppUpdateState().error).toContain('Auto-update is not available')
    expect(getAppUpdateState().error).not.toContain('github.com')
  })

  it('returns no changelog entries when no release repo is configured', async () => {
    process.env.ARA_GITHUB_OWNER = ''
    process.env.ARA_GITHUB_REPO = ''
    const { fetchChangelog } = await loadUpdater()
    await expect(fetchChangelog()).resolves.toEqual([])
  })

  it('configures auto-download and checks after the initial delay', async () => {
    const { initAutoUpdater } = await loadUpdater()
    await initAutoUpdater(() => null)
    expect(updaterState.autoDownload).toBe(true)
    expect(updaterState.autoInstallOnAppQuit).toBe(true)
    expect(updaterState.disableDifferentialDownload).toBe(false)
    expect(checkForUpdatesMock).not.toHaveBeenCalled()
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(checkForUpdatesMock).toHaveBeenCalledTimes(1)
  })

  it('re-checks on the periodic interval', async () => {
    const { initAutoUpdater } = await loadUpdater()
    await initAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(checkForUpdatesMock).toHaveBeenCalledTimes(1)
    vi.advanceTimersByTime(RECHECK_INTERVAL_MS)
    expect(checkForUpdatesMock).toHaveBeenCalledTimes(2)
  })

  it('is idempotent across repeated init calls', async () => {
    const { initAutoUpdater } = await loadUpdater()
    await initAutoUpdater(() => null)
    await initAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(checkForUpdatesMock).toHaveBeenCalledTimes(1)
  })

  it('handles update-available, download-progress, and update-downloaded events', async () => {
    const { initAutoUpdater, getAppUpdateState } = await loadUpdater()
    await initAutoUpdater(() => null)

    // Trigger update-available
    updaterState.listeners.get('update-available')!({
      version: '0.2.0',
      releaseNotes: 'New features',
    })
    expect(getAppUpdateState().status).toBe('downloading')
    expect(getAppUpdateState().availableVersion).toBe('0.2.0')
    expect(mockWebContents.send).toHaveBeenCalled()

    // Trigger download progress
    updaterState.listeners.get('download-progress')!({ percent: 55 })
    expect(getAppUpdateState().percent).toBe(55)

    // Trigger update downloaded
    updaterState.listeners.get('update-downloaded')!({ version: '0.2.0' })
    expect(getAppUpdateState().status).toBe('downloaded')
    expect(getAppUpdateState().percent).toBe(100)
  })

  it('handles update error', async () => {
    const { initAutoUpdater, getAppUpdateState } = await loadUpdater()
    await initAutoUpdater(() => null)

    updaterState.listeners.get('error')!(new Error('Network disconnected'))
    expect(getAppUpdateState().status).toBe('error')
    expect(getAppUpdateState().error).toContain('Network disconnected')
  })

  it('calls quitAndInstall on restartAndInstall', async () => {
    const { restartAndInstall } = await loadUpdater()
    restartAndInstall()
    vi.runAllTimers()
    expect(quitAndInstallMock).toHaveBeenCalledWith(true, true)
  })

  it('detects justUpdated when lastSeenAppVersion is older', async () => {
    const { initAutoUpdater, getJustUpdated, clearJustUpdated } = await loadUpdater()
    await initAutoUpdater(() => null)
    expect(getJustUpdated()).toBe(true)
    clearJustUpdated()
    expect(getJustUpdated()).toBe(false)
  })

  it('runs fake update simulation in unpacked runs', async () => {
    appState.isPackaged = false
    process.env.ARAOFFICE_FAKE_UPDATE = '1.2.0'
    const { initAutoUpdater, getAppUpdateState } = await loadUpdater()
    await initAutoUpdater(() => null)

    // Initial trigger
    vi.advanceTimersByTime(2000)
    expect(getAppUpdateState().status).toBe('checking')

    // Available -> downloading
    vi.advanceTimersByTime(1500)
    expect(getAppUpdateState().status).toBe('downloading')
    expect(getAppUpdateState().availableVersion).toBe('1.2.0')

    // Ticks until completed
    vi.advanceTimersByTime(5000)
    expect(getAppUpdateState().status).toBe('downloaded')
    expect(getAppUpdateState().percent).toBe(100)
  })
})
