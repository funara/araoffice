/// <reference types="vite/client" />

import type { ProjectApi } from '@araoffice/project-store'
import type { DiagramApi } from '../shared/ipc'

declare global {
  interface Window {
    diagramApi: DiagramApi
    projectApi?: Pick<ProjectApi, 'resolveChat' | 'appendChat' | 'loadChat' | 'rebindChat'>
  }
}
