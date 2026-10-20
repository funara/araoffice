import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig } from 'electron-vite'
const workspaceAlias = {
  '@araoffice/ai-provider/providers': resolve(
    __dirname,
    '../../packages/ai-provider/src/providers.ts',
  ),
  '@araoffice/pptx-engine/table-grid': resolve(
    __dirname,
    '../../packages/pptx-engine/src/table-grid.ts',
  ),
  '@araoffice/pptx-engine/background-promote': resolve(
    __dirname,
    '../../packages/pptx-engine/src/background-promote.ts',
  ),
  '@araoffice/pptx-render/preset-geometry': resolve(
    __dirname,
    '../../packages/pptx-render/src/preset-geometry.ts',
  ),
  '@araoffice/ui/ai-panel.css': resolve(__dirname, '../../packages/ui/src/ai-panel.css'),
  '@araoffice/agent-core': resolve(__dirname, '../../packages/agent-core/src/index.ts'),
  '@araoffice/ai-provider': resolve(__dirname, '../../packages/ai-provider/src/index.ts'),
  '@araoffice/ai-search': resolve(__dirname, '../../packages/ai-search/src/index.ts'),
  '@araoffice/docx-engine': resolve(__dirname, '../../packages/docx-engine/src/index.ts'),
  '@araoffice/electron-utils': resolve(__dirname, '../../packages/electron-utils/src/index.ts'),
  '@araoffice/file-parse': resolve(__dirname, '../../packages/file-parse/src/index.ts'),
  '@araoffice/i18n': resolve(__dirname, '../../packages/i18n/src/index.ts'),
  '@araoffice/project-store': resolve(__dirname, '../../packages/project-store/src/index.ts'),
  '@araoffice/pptx-engine': resolve(__dirname, '../../packages/pptx-engine/src/index.ts'),
  '@araoffice/pptx-render': resolve(__dirname, '../../packages/pptx-render/src/index.ts'),
  '@araoffice/ui': resolve(__dirname, '../../packages/ui/src/index.ts'),
}

export default defineConfig({
  // Bundle everything into the shell main (same policy as apps/docs): the
  // imported docs/sheets main modules are TS source with no build artifacts,
  main: { resolve: { alias: workspaceAlias } },
  preload: {
    build: {
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'src/preload/index.ts'),
        },
      },
    },
  },
  renderer: {
    resolve: { alias: workspaceAlias },
    plugins: [react()],
    publicDir: resolve(__dirname, '../../public'),
    build: {
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'src/renderer/index.html'),
        },
      },
    },
    server: {
      port: Number(process.env.SHELL_DEV_PORT) || 5199,
      strictPort: Boolean(process.env.SHELL_DEV_PORT),
    },
  },
})
