import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

const here = dirname(fileURLToPath(import.meta.url))

// Pin resolution to this repo's workspace sources (matches tsconfig paths;
// avoids bundling stale implementations when node_modules links point elsewhere)
const workspaceAlias = {
  // Subpath before the bare name: string aliases are prefix replacements
  '@araoffice/pptx-engine/table-grid': resolve(
    here,
    '../../packages/pptx-engine/src/table-grid.ts',
  ),
  '@araoffice/pptx-engine/smartart-layout': resolve(
    here,
    '../../packages/pptx-engine/src/smartart-layout.ts',
  ),
  '@araoffice/pptx-engine/background-promote': resolve(
    here,
    '../../packages/pptx-engine/src/background-promote.ts',
  ),
  '@araoffice/pptx-engine': resolve(here, '../../packages/pptx-engine/src/index.ts'),
  '@araoffice/pptx-render/preset-geometry': resolve(
    here,
    '../../packages/pptx-render/src/preset-geometry.ts',
  ),
  '@araoffice/pptx-render': resolve(here, '../../packages/pptx-render/src/index.ts'),
  '@araoffice/ai-provider/providers': resolve(here, '../../packages/ai-provider/src/providers.ts'),
  '@araoffice/ui/ai-panel.css': resolve(here, '../../packages/ui/src/ai-panel.css'),
  '@araoffice/agent-core': resolve(here, '../../packages/agent-core/src/index.ts'),
  '@araoffice/docx-engine': resolve(here, '../../packages/docx-engine/src/index.ts'),
  '@araoffice/ai-provider': resolve(here, '../../packages/ai-provider/src/index.ts'),
  '@araoffice/ai-search': resolve(here, '../../packages/ai-search/src/index.ts'),
  '@araoffice/electron-utils': resolve(here, '../../packages/electron-utils/src/index.ts'),
  '@araoffice/file-parse': resolve(here, '../../packages/file-parse/src/index.ts'),
  '@araoffice/i18n': resolve(here, '../../packages/i18n/src/index.ts'),
  '@araoffice/project-store': resolve(here, '../../packages/project-store/src/index.ts'),
  '@araoffice/ui': resolve(here, '../../packages/ui/src/index.ts'),
}

export default defineConfig({
  // Main process/preload must bundle @araoffice/* sources (they are pulled in as TS
  // source with extensionless relative imports; externalizing them under Node
  // yields ERR_MODULE_NOT_FOUND).
  main: {
    resolve: { alias: workspaceAlias },
    // Bundle opentype.js too (the packaged app ships only out/**, so external deps are unresolvable at runtime)
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@araoffice/agent-core',
          '@araoffice/ai-provider',
          '@araoffice/ai-search',
          '@araoffice/docx-engine',
          '@araoffice/electron-utils',
          '@araoffice/file-parse',
          '@araoffice/i18n',
          '@araoffice/project-store',
          '@araoffice/pptx-engine',
          '@araoffice/pptx-render',
          'opentype.js',
        ],
      }),
    ],
  },
  preload: {
    plugins: [externalizeDepsPlugin()],
  },
  renderer: {
    resolve: { alias: workspaceAlias },
    plugins: [react()],
    publicDir: resolve(__dirname, '../../public'),
    server: {
      port: Number(process.env.SLIDES_DEV_PORT) || 5175,
      strictPort: Boolean(process.env.SLIDES_DEV_PORT),
    },
  },
})
