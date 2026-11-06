/**
 * Sheets 浏览器形态启动（:5201）：renderer 原样 + 注入 SheetsApi 垫片（模块版，
 * 引擎调用走 xlsxEngineWorker postMessage 协议）
 */
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { fileURLToPath } from 'node:url'
import { copyFileSync, readFileSync, writeFileSync } from 'node:fs'

const shimPath = fileURLToPath(new URL('./sheets-shim.ts', import.meta.url))
let isBuild = false
// build 产物根（--outDir 相对 root:src/renderer 解析后的落点）
const webLaunchDist = () =>
  fileURLToPath(new URL('../src/renderer/web-launch/dist/', import.meta.url))
const jszipDist = fileURLToPath(new URL('../../../node_modules/jszip/dist/jszip.min.js', import.meta.url))

export default defineConfig({
  root: 'src/renderer',
  // 前缀托管（/file-editor-sheets/）：build 资产相对化 + 垫片独立 entry
  base: './',
  plugins: [
    react(),
    {
      name: 'force-dev-inline',
      enforce: 'post' as const,
      transform(code: string, id: string) {
        if (code.includes('import.meta.env.DEV')) {
          console.log('[force-dev] replacing in', id.slice(-50))
          let out = code.replaceAll('import.meta.env.DEV', 'true')
          out = out.replace(
            'if (!(await openLazyWorkbook(selected,',
            'console.warn("[AFTER-SELECT] about to call openLazyWorkbook, selected=", typeof selected, selected?.name); if (!(await openLazyWorkbook(selected,'
          )
          out = out.replace('if (!(await commitActiveEditor())) return false',
            'console.warn("[STEP] 1-commitActiveEditor"); if (!(await commitActiveEditor())) { console.warn("[STEP] FAIL-commit"); return false }')
          out = out.replace('await untilCellFontsLoaded(',
            'console.warn("[STEP] 2-fonts"); await untilCellFontsLoaded(')
          out = out.replace('lazyWorkbookRef.current = state',
            'console.warn("[STEP] 3-lazyRef"); lazyWorkbookRef.current = state')
          out = out.replace('loadWorkbookSkeleton(univerRef.current, selected)',
            'console.warn("[STEP] 4-skeleton"); loadWorkbookSkeleton(univerRef.current, selected)')
          // 在 openLazyWorkbook 的 debug hooks 前注入到达日志
          out = out.replace(
            /if \(true\) \{\n(\s+);\n(\s+)window\.__journal/,
            'if (true) { console.warn(\"[OPEN-REACHED] openLazyWorkbook 到达 debug hooks\"); window.__journal'
          )
          return out
        }
      },
    },
    {
      name: 'inject-sheets-web-shim',
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
          // dev：/@fs 直读；build：jszip 由 closeBundle 拷到产物根、垫片是 rollup 独立 entry
          const jszipSrc = isBuild ? './jszip.min.js' : '/@fs' + jszipDist
          const shimSrc = isBuild ? '<!--__SHIM_SCRIPT__-->' : `<script type="module" src="${'/@fs' + shimPath}"></script>`
          return html.replace(
            '<head>',
            `<head>\n    <script src="${jszipSrc}"></script>\n    ${shimSrc}`
          )
        }
      },
      closeBundle() {
        if (!isBuild) return
        copyFileSync(jszipDist, webLaunchDist() + 'jszip.min.js')
        // 垫片 entry 回填 index.html（占位符 → 相对引用；理由见 slides 配置同款注释）
        const htmlPath = webLaunchDist() + 'index.html'
        const html = readFileSync(htmlPath, 'utf8')
        writeFileSync(
          htmlPath,
          html.replace('<!--__SHIM_SCRIPT__-->', '<script type="module" src="./shim-entry.js"></script>')
        )
      }
    }
  ],
  // M9 保存链：xlsx-gateway（规划器+JSZip 组装器）在浏览器跑，其顶层 node: 导入
  // 映射到 M8 已备的浏览器 shim（fs/path 走 vite 浏览器外置桩，保存路径不会触达）
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

  define: { 'import.meta.env.DEV': 'true', 'import.meta.env.MODE': '"development"' },
  publicDir: fileURLToPath(new URL('./public', import.meta.url)),
  build: {
    rollupOptions: {
      input: { index: fileURLToPath(new URL('../src/renderer/index.html', import.meta.url)), 'shim-entry': shimPath },
      output: {
        entryFileNames: (chunk) => (chunk.name === 'shim-entry' ? 'shim-entry.js' : 'assets/[name]-[hash].js')
      }
    }
  },
  server: {
    port: 5201,
    strictPort: true,
    fs: { allow: [fileURLToPath(new URL('../../..', import.meta.url))] }
  }
})
