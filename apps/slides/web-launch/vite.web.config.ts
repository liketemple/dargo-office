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
let isBuild = false

import { copyFileSync, readFileSync, writeFileSync } from 'node:fs'
const jszipDistPath = () =>
  fileURLToPath(new URL('../../../node_modules/jszip/dist/jszip.min.js', import.meta.url))
const webLaunchDist = () =>
  fileURLToPath(new URL('../src/renderer/web-launch/dist/', import.meta.url))

export default defineConfig({
  root: 'src/renderer',
  // 前缀托管（/file-editor-{slides,sheets}/）：build 资产相对化 + 垫片独立 entry
  base: './',
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
      configResolved(cfg) {
        isBuild = cfg.command === 'build'
      },
      transformIndexHtml: {
        order: 'pre',
        handler(html: string) {
          html = html.replace(
            /<meta[^>]+http-equiv=["']?Content-Security-Policy["']?[^>]*>/i,
            '<!-- CSP removed for web launch -->'
          )
          // 经典 script 先行：jszip UMD → window.JSZip（CDN 同款浏览器路径，内部
          // webpack 图在真实 window 上下文初始化，规避 new Function 沙箱的 node 分支坑）
          // dev：/@fs 直读；build：jszip 由 closeBundle 拷到产物根、垫片是 rollup 独立 entry
          const jszipSrc = isBuild ? './jszip.min.js' : '/@fs' + jszipDistPath()
          const shimSrc = isBuild ? '<!--__SHIM_SCRIPT__-->' : `<script type="module" src="${'/@fs' + shimPath}"></script>`
          return html.replace(
            '<head>',
            `<head>\n    <script src="${jszipSrc}"></script>\n    ${shimSrc}`
          )
        }
      },
      closeBundle() {
        if (!isBuild) return
        // ① jszip UMD 拷产物根 ② 垫片 entry 回填 index.html（build-html 会解析
        // module script 相对 src，故 transformIndexHtml 只放占位符、此处文件级回填）
        copyFileSync(jszipDistPath(), webLaunchDist() + 'jszip.min.js')
        const htmlPath = webLaunchDist() + 'index.html'
        const html = readFileSync(htmlPath, 'utf8')
        writeFileSync(
          htmlPath,
          html.replace('<!--__SHIM_SCRIPT__-->', '<script type="module" src="./shim-entry.js"></script>')
        )
      }
    }
  ],
  // M9b/c：pptx-ops/pptx-engine 在浏览器跑，node: 系命名导入映射到 web-launch/shims
  resolve: {
    alias: [
      { find: /^node:fs\/promises$/, replacement: fileURLToPath(new URL('./shims/node-fs-promises.ts', import.meta.url)) },
      { find: /^node:fs$/, replacement: fileURLToPath(new URL('./shims/node-fs.ts', import.meta.url)) },
      { find: /^node:path$/, replacement: fileURLToPath(new URL('./shims/node-path.ts', import.meta.url)) },
      { find: /^node:crypto$/, replacement: fileURLToPath(new URL('./shims/node-crypto.ts', import.meta.url)) },
      { find: /^node:zlib$/, replacement: fileURLToPath(new URL('./shims/node-zlib.ts', import.meta.url)) },
      { find: /^node:stream$/, replacement: fileURLToPath(new URL('./shims/node-stream.ts', import.meta.url)) },
      { find: /^buffer$/, replacement: fileURLToPath(new URL('./shims/buffer.ts', import.meta.url)) }
    ]
  },
  build: {
    rollupOptions: {
      input: { index: fileURLToPath(new URL('../src/renderer/index.html', import.meta.url)), 'shim-entry': shimPath },
      output: {
        entryFileNames: (chunk) => (chunk.name === 'shim-entry' ? 'shim-entry.js' : 'assets/[name]-[hash].js')
      }
    }
  },
  server: {
    port: 5200,
    strictPort: true,
    fs: { allow: [fileURLToPath(new URL('../../..', import.meta.url))] }
  }
})
