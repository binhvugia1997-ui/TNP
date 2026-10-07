import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  // pywebview opens the production bundle from a local file URI; relative asset URLs keep CSS/JS beside index.html.
  base: './',
  plugins: [react()],
  // Prototype sample imagery remains in source but is never copied into the real desktop build.
  publicDir: false,
  server: {
    host: true,
    port: 5173,
    strictPort: true,
    allowedHosts: true,
  },
  preview: {
    host: true,
    port: 5173,
    allowedHosts: true,
  },
})
