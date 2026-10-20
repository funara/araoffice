import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
const workspaceAlias = {
  '@araoffice/ai-provider/providers': resolve(
    __dirname,
    '../../packages/ai-provider/src/providers.ts',
  ),
  '@araoffice/ai-provider': resolve(__dirname, '../../packages/ai-provider/src/index.ts'),
  '@araoffice/docx-engine': resolve(__dirname, '../../packages/docx-engine/src/index.ts'),
  '@araoffice/pptx-render/preset-geometry': resolve(
    __dirname,
    '../../packages/pptx-render/src/preset-geometry.ts',
  ),
  '@araoffice/agent-core': resolve(__dirname, '../../packages/agent-core/src/index.ts'),
  '@araoffice/ui/ai-panel.css': resolve(__dirname, '../../packages/ui/src/ai-panel.css'),
  '@araoffice/ui': resolve(__dirname, '../../packages/ui/src/index.ts'),
  '@araoffice/ai-search': resolve(__dirname, '../../packages/ai-search/src/index.ts'),
  '@araoffice/electron-utils': resolve(__dirname, '../../packages/electron-utils/src/index.ts'),
  '@araoffice/file-parse': resolve(__dirname, '../../packages/file-parse/src/index.ts'),
  '@araoffice/i18n': resolve(__dirname, '../../packages/i18n/src/index.ts'),
  '@araoffice/project-store': resolve(__dirname, '../../packages/project-store/src/index.ts'),
}

export default defineConfig({
  // Main and preload use only electron + node builtins; bundle everything so
  // the packaged app doesn't rely on node_modules at runtime.
  // @araoffice/* deps ship as raw TS source with extensionless imports, so they
  // must be bundled — externalizing them yields ERR_MODULE_NOT_FOUND under Node
  // (same setup as apps/slides).
  main: {
    resolve: { alias: workspaceAlias },
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@araoffice/ai-provider',
          '@araoffice/ai-search',
          '@araoffice/docx-engine',
          '@araoffice/electron-utils',
          '@araoffice/file-parse',
          '@araoffice/i18n',
          '@araoffice/project-store',
        ],
      }),
    ],
  },
  preload: {},
  renderer: {
    resolve: { alias: workspaceAlias },
    plugins: [react()],
    publicDir: resolve(__dirname, '../../public'),
    server: {
      // Overridable so multiple araoffice dev instances can coexist (default 5173).
      port: Number(process.env.DOCS_DEV_PORT) || 5173,
      strictPort: Boolean(process.env.DOCS_DEV_PORT),
    },
  },
})
