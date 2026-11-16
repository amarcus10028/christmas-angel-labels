import react from '@vitejs/plugin-react'
import { readFileSync } from 'node:fs'
import { defineConfig, type Plugin } from 'vite'

const { version } = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8')) as { version: string }

// Privacy guarantee: the built page may not make ANY network request (connect-src 'none'),
// so spreadsheet contents can never leave the browser. Added only to production builds
// because the dev server needs inline scripts and a websocket.
const CSP =
  "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; " +
  "font-src 'self'; connect-src 'none'; form-action 'none'; base-uri 'self'; object-src 'none'"

const csp = (): Plugin => ({
  name: 'csp',
  apply: 'build',
  transformIndexHtml: () => [{ tag: 'meta', attrs: { 'http-equiv': 'Content-Security-Policy', content: CSP }, injectTo: 'head-prepend' }],
})

// Relative base so the built site works at any GitHub Pages path (e.g. /christmas-angel-labels/).
export default defineConfig({
  base: './',
  plugins: [react(), csp()],
  // Shown in the footer as "v1.1". Bump "version" in package.json and add a CHANGELOG.md entry for each release.
  define: { __APP_VERSION__: JSON.stringify(version) },
  build: { chunkSizeWarningLimit: 1200 },
})
