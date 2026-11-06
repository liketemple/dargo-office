/**
 * 浏览器形态启动配置：renderer 原样 + 注入 WebDesktopApi 垫片
 * （pnpm --filter @genoffice/docs dev:web，端口 5199）
 */
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

const shimPath = fileURLToPath(new URL('./web-desktop-shim.js', import.meta.url))
let isBuild = false

export default defineConfig({
  root: 'src/renderer',
  // 前缀托管（/file-editor/…）：build 资产相对化 + 垫片作为独立 entry（shim-entry.js）
  base: './',
  plugins: [
    react(),
    {
      name: 'inject-web-desktop-shim',
      configResolved(cfg) {
        isBuild = cfg.command === 'build'
      },
      transformIndexHtml: {
        order: 'pre',
        handler(html: string) {
          // 1) 注入 WebDesktopApi 垫片 2) 移除 Electron 专用 CSP meta（浏览器形态按同源策略即可）
          html = html.replace(
            /<meta[^>]+http-equiv=["']?Content-Security-Policy["']?[^>]*>/i,
            '<!-- CSP removed for web launch -->'
          )
          // dev：/@fs 直读源文件；build：垫片是 rollup 独立 entry（shim-entry.js，同目录）
          const src = isBuild ? './shim-entry.js' : '/@fs' + shimPath
          return html.replace('<head>', `<head>\n    <script src="${src}"></script>`)
        }
      }
    }
  ],
  build: {
    rollupOptions: {
      input: { index: fileURLToPath(new URL('../src/renderer/index.html', import.meta.url)), 'shim-entry': shimPath },
      output: {
        entryFileNames: (chunk) => (chunk.name === 'shim-entry' ? 'shim-entry.js' : 'assets/[name]-[hash].js')
      }
    }
  },
  server: {
    port: 5199,
    strictPort: true,
    fs: { allow: [fileURLToPath(new URL('../..', import.meta.url)), fileURLToPath(new URL('../../..', import.meta.url))] }
  }
})
