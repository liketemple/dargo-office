/**
 * Slides 浏览器形态启动（:5200）：renderer 原样 + 注入模块版垫片（slides-shim.ts，
 * 内嵌 pptx-engine/pptx-render 打开链路）
 */
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

const rootDir = fileURLToPath(new URL('..', import.meta.url))
// renderer index.html 位置与 docs 一致（src/renderer/index.html）
const shimPath = fileURLToPath(new URL('./slides-shim.ts', import.meta.url))

export default defineConfig({
  root: 'src/renderer',
  resolve: {
    alias: {
      'node:crypto': fileURLToPath(new URL('./shims/node-crypto.ts', import.meta.url)),
      'node:zlib': fileURLToPath(new URL('./shims/node-zlib.ts', import.meta.url)),
      'node:stream/promises': fileURLToPath(new URL('./shims/node-stream.ts', import.meta.url)),
      stream: fileURLToPath(new URL('./shims/node-stream.ts', import.meta.url)),
      'readable-stream': fileURLToPath(new URL('./shims/node-stream.ts', import.meta.url)),
      // jszip 预打包默认解析进了 node 分支（内联 readable-stream→require 链）；
      // 钉死浏览器 dist
      jszip: fileURLToPath(new URL('./shims/jszip.ts', import.meta.url)),
      process: fileURLToPath(new URL('./shims/node-process.ts', import.meta.url)),
      'node:fs/promises': fileURLToPath(new URL('./shims/node-fs.ts', import.meta.url))
    }
  },
  plugins: [
    react(),
    {
      name: 'inject-slides-web-shim',
      transformIndexHtml: {
        order: 'pre',
        handler(html: string) {
          html = html.replace(
            /<meta[^>]+http-equiv=["']?Content-Security-Policy["']?[^>]*>/i,
            '<!-- CSP removed for web launch -->'
          )
          // 经典 script 先行：jszip UMD → window.JSZip（CDN 同款浏览器路径，内部
          // webpack 图在真实 window 上下文初始化，规避 new Function 沙箱的 node 分支坑）
          const jszipDist = fileURLToPath(new URL('../../../node_modules/jszip/dist/jszip.min.js', import.meta.url))
          return html.replace(
            '<head>',
            `<head>\n    <script src="${'/@fs' + jszipDist}"></script>\n    <script type="module" src="${'/@fs' + shimPath}"></script>`
          )
        }
      }
    }
  ],
  server: {
    port: 5200,
    strictPort: true,
    fs: { allow: [fileURLToPath(new URL('../../..', import.meta.url))] }
  }
})
