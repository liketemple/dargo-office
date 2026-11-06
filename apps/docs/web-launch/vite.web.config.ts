/**
 * 浏览器形态启动配置：renderer 原样 + 注入 WebDesktopApi 垫片
 * （pnpm --filter @genoffice/docs dev:web，端口 5199）
 */
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

const shimPath = fileURLToPath(new URL('./web-desktop-shim.js', import.meta.url))

export default defineConfig({
  root: 'src/renderer',
  plugins: [
    react(),
    {
      name: 'inject-web-desktop-shim',
      transformIndexHtml: {
        order: 'pre',
        handler(html) {
          // 1) 注入 WebDesktopApi 垫片 2) 移除 Electron 专用 CSP meta（浏览器形态按同源策略即可）
          html = html.replace(
            /<meta[^>]+http-equiv=["']?Content-Security-Policy["']?[^>]*>/i,
            '<!-- CSP removed for web launch -->'
          )
          return html.replace('<head>', `<head>\n    <script src="${'/@fs' + shimPath}"></script>`)
        }
      }
    }
  ],
  server: {
    port: 5199,
    strictPort: true,
    fs: { allow: [fileURLToPath(new URL('../..', import.meta.url)), fileURLToPath(new URL('../../..', import.meta.url))] }
  }
})
