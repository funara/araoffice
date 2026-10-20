export interface WindowApi {
  readonly isWindows: boolean
  minimize(): Promise<void>
  toggleMaximize(): Promise<void>
  isMaximized(): Promise<boolean>
  close(): Promise<void>
  openExternal(url: string): Promise<void>
  openThirdPartyNotices(): Promise<void>
  onMaximizedChanged(handler: (maximized: boolean) => void): () => void
}

export const WINDOW_CHANNELS = {
  minimize: 'window:minimize',
  toggleMaximize: 'window:toggle-maximize',
  isMaximized: 'window:is-maximized',
  close: 'window:close',
  stateChanged: 'window:state-changed',
  openExternal: 'window:open-external',
  openThirdPartyNotices: 'window:open-third-party-notices',
} as const
