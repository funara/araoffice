import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

const workspaceAlias = {
  '@araoffice/ai-provider/providers': resolve(
    __dirname,
    '../../packages/ai-provider/src/providers.ts',
  ),
  '@araoffice/ui/ai-panel.css': resolve(__dirname, '../../packages/ui/src/ai-panel.css'),
  '@araoffice/agent-core': resolve(__dirname, '../../packages/agent-core/src/index.ts'),
  '@araoffice/ai-provider': resolve(__dirname, '../../packages/ai-provider/src/index.ts'),
  '@araoffice/electron-utils': resolve(__dirname, '../../packages/electron-utils/src/index.ts'),
  '@araoffice/i18n': resolve(__dirname, '../../packages/i18n/src/index.ts'),
  '@araoffice/ui': resolve(__dirname, '../../packages/ui/src/index.ts'),
}

export default defineConfig({
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
    resolve: { alias: workspaceAlias },
    server: {
      port: Number(process.env.DIAGRAM_DEV_PORT) || 5178,
      strictPort: Boolean(process.env.DIAGRAM_DEV_PORT),
    },
    build: {
      rollupOptions: {
        onwarn(warning, warn) {
          if (warning.code === 'EVAL' && warning.id?.includes('@maxgraph/core')) {
            return
          }
          warn(warning)
        },
      },
    },
  },
})
