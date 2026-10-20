import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

// npm hoists some @tiptap packages to the repo root (shared with docs at a
// different version) and nests others under this app — dedupe forces every
// import onto this app's single copy so the bundle never carries two cores.
const TIPTAP_DEDUPE = [
  '@tiptap/core',
  '@tiptap/pm',
  '@tiptap/react',
  '@tiptap/extensions',
  '@tiptap/extension-list',
  '@tiptap/extension-table',
  '@tiptap/extension-image',
  '@tiptap/suggestion',
  '@tiptap/markdown',
  '@tiptap/extension-highlight',
  '@tiptap/extension-code-block',
]
const workspaceAlias = {
  '@araoffice/ai-provider/providers': resolve(
    __dirname,
    '../../packages/ai-provider/src/providers.ts',
  ),
  '@araoffice/ui/ai-panel.css': resolve(__dirname, '../../packages/ui/src/ai-panel.css'),
  '@araoffice/pptx-render/preset-geometry': resolve(
    __dirname,
    '../../packages/pptx-render/src/preset-geometry.ts',
  ),
  '@araoffice/agent-core': resolve(__dirname, '../../packages/agent-core/src/index.ts'),
  '@araoffice/ai-provider': resolve(__dirname, '../../packages/ai-provider/src/index.ts'),
  '@araoffice/electron-utils': resolve(__dirname, '../../packages/electron-utils/src/index.ts'),
  '@araoffice/i18n': resolve(__dirname, '../../packages/i18n/src/index.ts'),
  '@araoffice/project-store': resolve(__dirname, '../../packages/project-store/src/index.ts'),
  '@araoffice/ui': resolve(__dirname, '../../packages/ui/src/index.ts'),
}

export default defineConfig({
  // @araoffice/i18n and @araoffice/electron-utils ship as TS source — must be bundled
  main: {
    resolve: { alias: workspaceAlias },
    plugins: [externalizeDepsPlugin({ exclude: ['@araoffice/i18n', '@araoffice/electron-utils'] })],
  },
  preload: {
    plugins: [externalizeDepsPlugin({ exclude: ['@araoffice/i18n'] })],
  },
  renderer: {
    plugins: [react()],
    publicDir: resolve(__dirname, '../../public'),
    resolve: { alias: workspaceAlias, dedupe: TIPTAP_DEDUPE },
    server: {
      port: Number(process.env.MARKDOWN_DEV_PORT) || 5177,
      strictPort: Boolean(process.env.MARKDOWN_DEV_PORT),
    },
  },
})
