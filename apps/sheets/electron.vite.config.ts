import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
const workspaceAlias = {
  '@araoffice/ai-provider/providers': resolve(
    __dirname,
    '../../packages/ai-provider/src/providers.ts',
  ),
  '@araoffice/docx-engine/math': resolve(__dirname, '../../packages/docx-engine/src/math.ts'),
  '@araoffice/agent-core': resolve(__dirname, '../../packages/agent-core/src/index.ts'),
  '@araoffice/ai-provider': resolve(__dirname, '../../packages/ai-provider/src/index.ts'),
  '@araoffice/ai-search': resolve(__dirname, '../../packages/ai-search/src/index.ts'),
  '@araoffice/docx-engine': resolve(__dirname, '../../packages/docx-engine/src/index.ts'),
  '@araoffice/pptx-render/preset-geometry': resolve(
    __dirname,
    '../../packages/pptx-render/src/preset-geometry.ts',
  ),
  '@araoffice/electron-utils': resolve(__dirname, '../../packages/electron-utils/src/index.ts'),
  '@araoffice/file-parse': resolve(__dirname, '../../packages/file-parse/src/index.ts'),
  '@araoffice/i18n': resolve(__dirname, '../../packages/i18n/src/index.ts'),
  '@araoffice/project-store': resolve(__dirname, '../../packages/project-store/src/index.ts'),
  '@araoffice/ui/ai-panel.css': resolve(__dirname, '../../packages/ui/src/ai-panel.css'),
  '@araoffice/ui': resolve(__dirname, '../../packages/ui/src/index.ts'),
}

export default defineConfig({
  main: {
    resolve: { alias: workspaceAlias },
    // compiled entry point) — externalizing them makes Node's ESM loader try
    // to resolve their relative imports at runtime and fail. Bundle those;
    // externalize everything else (Electron, zod, node builtins).
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@araoffice/ai-provider',
          '@araoffice/agent-core',
          '@araoffice/ai-search',
          '@araoffice/docx-engine',
          '@araoffice/file-parse',
          '@araoffice/electron-utils',
          '@araoffice/i18n',
        ],
      }),
    ],
  },
  preload: {
    // Sandboxed preload scripts cannot require arbitrary npm packages at runtime.
    plugins: [],
  },
  renderer: {
    resolve: { alias: workspaceAlias },
    plugins: [react()],
    publicDir: resolve(__dirname, '../../public'),
  },
})
