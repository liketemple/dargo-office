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
  define: { 'import.meta.env.DEV': 'true', 'import.meta.env.MODE': '"development"' },
  publicDir: fileURLToPath(new URL('./public', import.meta.url)),
  server: {
    port: 5201,
    strictPort: true,
    fs: { allow: [fileURLToPath(new URL('../../..', import.meta.url))] }
  }
})
