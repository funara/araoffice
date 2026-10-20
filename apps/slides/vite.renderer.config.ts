import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

const here = dirname(fileURLToPath(import.meta.url))

// renderer-only dev server (embedded by the shell via SLIDES_RENDERER_URL for HMR; no standalone Electron)
export default defineConfig({
  root: 'src/renderer',
  publicDir: resolve(here, '../../public'),
  plugins: [react()],
  server: {
    port: Number(process.env.SLIDES_DEV_PORT) || 5175,
    strictPort: true,
  },
})
