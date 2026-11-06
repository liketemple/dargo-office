/**
 * Sheets 浏览器形态启动（:5201）：renderer 原样 + 注入 SheetsApi 垫片（模块版，
 * 引擎调用走 xlsxEngineWorker postMessage 协议）
 */
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { fileURLToPath } from 'node:url'

const shimPath = fileURLToPath(new URL('./sheets-shim.ts', import.meta.url))
const jszipDist = fileURLToPath(new URL('../../../node_modules/jszip/dist/jszip.min.js', import.meta.url))

export default defineConfig({
  root: 'src/renderer',
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
      transformIndexHtml: {
        order: 'pre',
        handler(html: string) {
          html = html.replace(
            /<meta[^>]+http-equiv=["']?Content-Security-Policy["']?[^>]*>/i,
            '<!-- CSP removed for web launch -->'
          )
          // 经典 script 先行 jszip UMD → window.JSZip（CDN 同款路径）
          return html.replace(
            '<head>',
            `<head>\n    <script src="${'/@fs' + jszipDist}"></script>\n    <script type="module" src="${'/@fs' + shimPath}"></script>`
          )
        }
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
  server: {
    port: 5201,
    strictPort: true,
    fs: { allow: [fileURLToPath(new URL('../../..', import.meta.url))] }
  }
})
