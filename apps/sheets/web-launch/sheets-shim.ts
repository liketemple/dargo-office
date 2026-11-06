/**
 * GenOffice Sheets renderer 的浏览器垫片（模块版）：
 * desktopApi 桥的引擎调用走 xlsxEngineWorker postMessage 协议（M8 核心），
 * 其余偏走万能代理（docs/slides 三连修最终形态）。
 *
 * 核心实装方法（打开→渲染数据链路）：
 * - selectWorkbook / reopenWorkbook → ?src= 参数 fetch 字节 → worker open
 * - readWorkbookRange → worker read_range（映射 sessionId/sheetId/range）
 * - getLanguage / getTheme / 偏好族 → localStorage
 */
import { Buffer } from 'buffer'
// 与 Electron 主进程同构的归一化：workbookFileSchema.parse 会给引擎 serde
// 省略的空数组字段补 .default([])（sparklines/cellImages/pivotTables…）。
// 缺了这层，sheet.sparklines 是 undefined，编辑提交监听器读 .length 即抛
// TypeError（App.tsx 编辑流水线每次提交必炸）。
import { workbookFileSchema } from '../src/shared/desktop-api'

const g = globalThis as Record<string, unknown>
if (!g.Buffer) g.Buffer = Buffer
if (!g.process) {
  g.process = { env: { NODE_ENV: 'production' }, browser: true, version: 'v20' }
}
if (!g.setImmediate) {
  g.setImmediate = (fn: (...a: unknown[]) => void, ...args: unknown[]) =>
    setTimeout(() => fn(...args), 0)
}
if (!g.global) g.global = globalThis

;(() => {
  if ((window as any).desktopApi) return
  const q = new URLSearchParams(location.search)
  const LS = 'genoffice-sheets-shim:'
  const prefGet = (k: string, f: unknown) => {
    try {
      const v = localStorage.getItem(LS + k)
      return v === null ? f : JSON.parse(v)
    } catch {
      return f
    }
  }

  // ---- xlsx 引擎 Worker ----
  let worker: Worker | null = null
  let wSeq = 0
  const wPending = new Map<number, { resolve: (v: any) => void; reject: (e: any) => void }>()

  function ensureWorker(): Promise<void> {
    if (worker) return Promise.resolve()
    return new Promise((resolve, reject) => {
      // Worker 脚本经父页（编辑器路由）同源 /src/workers/ 加载；独立 :5201 形态
      // 跨源时 fallback 到 File Viewer 只读（后续接产物化路径）
      worker = new Worker('./xlsxEngineWorker.js', { type: 'module' })
      worker.onmessage = (ev) => {
        const d = ev.data
        if (d.type === 'ready') {
          resolve()
          return
        }
        const p = wPending.get(d.id)
        if (!p) return
        wPending.delete(d.id)
        if (d.type === 'resp') p.resolve(d.json ? JSON.parse(d.json) : null)
        else p.reject(new Error(d.message || 'worker error'))
      }
      worker.onerror = (e) => reject(new Error('worker fatal: ' + e.message))
      worker.postMessage({ type: 'init' })
    })
  }

  async function wOpen(bytes: ArrayBuffer, name: string) {
    await ensureWorker()
    return new Promise<any>((resolve, reject) => {
      const id = ++wSeq
      wPending.set(id, { resolve, reject })
      worker!.postMessage({ type: 'open', id, bytes, name })
    })
  }

  async function wCmd(line: string) {
    await ensureWorker()
    return new Promise<any>((resolve, reject) => {
      const id = ++wSeq
      wPending.set(id, { resolve, reject })
      worker!.postMessage({ type: 'cmd', id, line })
    })
  }

  // ---- 万能代理（docs/slides 同款）----
  const callable: any = new Proxy(function () {} as any, {
    get(_t, p) {
      if (p === Symbol.toPrimitive || typeof p === 'symbol') return undefined
      // 真异步：回调抛错变成 rejected promise（调用方的 .catch 能接住），
      // 而不是同步穿透把调用函数本身打死。
      if (p === 'then')
        return (onFulfilled: (v: unknown) => unknown) =>
          Promise.resolve().then(() => (onFulfilled ? onFulfilled(undefined) : undefined))
      if (p === 'catch') return () => Promise.resolve(undefined)
      if (p === 'finally') return (cb: any) => Promise.resolve(cb ? cb() : undefined)
      return callable
    },
    apply() {
      return callable
    }
  })

  let pendingSrc: string | null = q.get('src')
  let pendingName: string = q.get('name') || 'workbook.xlsx'
  let sessionMeta: any = null // open 应答回存的 sheets 元数据

  const api: Record<string, unknown> = {
    // ---- 打开 ----
    hasQueuedWorkbook: async () => !!pendingSrc,
    listWorkbookConnections: async () => ({ connections: [] }),
    readWorkbookExternalLinks: async () => ({ links: [], usages: [], suggestions: [] }),
    onMenuAction: () => () => {},
    selectWorkbook: async () => {
      if (!pendingSrc) return null
      try {
        const res = await fetch(pendingSrc)
        if (!res.ok) throw new Error(`加载失败 ${res.status}`)
        const bytes = await res.arrayBuffer()
        const open = await wOpen(bytes, pendingName)
        console.info('[sheets-shim] open ✅', open?.ok, open?.result?.sheets?.map((s: any) => s.name).join(','))
        if (!open?.ok) throw new Error(open?.error?.message || 'open failed')
        const r = open.result
        sessionMeta = r
        pendingSrc = null
        // sha256（WorkbookFile 必填）：WebCrypto 计算
        const digest = await crypto.subtle.digest('SHA-256', bytes)
        const sha256 = Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, '0')).join('')
        const result = {
          sessionId: r.sessionId,
          name: pendingName,
          path: pendingName,
          sha256,
          fileBytes: bytes.byteLength,
          entryCount: r.entryCount ?? 0,
          sheets: r.sheets ?? [],
          activeTab: r.activeTab ?? 0,
          styles: r.styles ?? [],
          dxfStyles: r.dxfStyles ?? [],
          visuals: r.visuals ?? [],
          definedNames: r.definedNames ?? [],
          normalFontName: r.normalFontName,
          workbookProtection: r.workbookProtection,
          readOnly: false,
          needsSaveAs: false,
          unsavedNew: true,
          date1904: r.date1904 ?? false,
          shortDateFormat: r.shortDateFormat,
          themeColors: r.themeColors,
          themeFonts: r.themeFonts
        }
        // Electron 主进程在这一步做 workbookFileSchema.parse（sheets-main.ts
        // openWorkbooksFromPaths），省略字段在此补默认——垫片同构照做。
        return workbookFileSchema.parse(result)
      } catch (e) {
        console.error('[sheets-shim] open 失败', e)
        return null
      }
    },
    reopenWorkbook: async () => null,
    openEncryptedWorkbook: async () => null,
    openSourcePath: async () => {},
    openWorkbooksForMerge: async () => null,
    selectWorkbooksForMerge: async () => null,

    // ---- 引擎读 ----
    // 引擎不提取 redaction 部件：按「文件无脱敏标记」作答（status:'absent'）。
    // 缺了这个，callable 兜底的 .then(undefined) 会让 redactionSessionFor 读
    // undefined.status 抛 TypeError，把 openLazyWorkbook 同步打死——
    // loadWorkbookSkeleton 不执行，Univer 停在 initialSnapshot（Sheet1/单sheet/默认行高列宽）。
    readWorkbookRedactions: async () => ({ status: 'absent' }),
    readWorkbookRange: async (request: any) => {
      const cmd = {
        version: 1,
        requestId: `rr-${++wSeq}`,
        command: 'read_range',
        sessionId: request.sessionId,
        sheetId: request.sheetId,
        range: {
          startRow: request.range.startRow,
          endRow: request.range.endRow,
          startColumn: request.range.startColumn,
          endColumn: request.range.endColumn
        }
      }
      const resp = await wCmd(JSON.stringify(cmd))
      if (!resp?.ok) {
        console.warn('[sheets-shim] read_range 失败', resp?.error?.message)
        return { cells: [], rows: [] }
      }
      // zod .strict() 裁剪：只保留 renderer 契约的字段（多余字段致校验静默失败）
      const r = resp.result
      console.info('[sheets-shim] read_range ✅', r.cells?.length, 'cells,', r.rows?.length, 'rows')
      return {
        cells: (r.cells || []).map((c: any) => ({
          row: c.row, column: c.column, value: c.value,
          ...(c.formula ? { formula: c.formula } : {}),
          ...(c.styleIndex !== undefined ? { styleIndex: c.styleIndex } : {})
        })),
        rows: (r.rows || []).map((rw: any) => ({
          row: rw.row, hidden: rw.hidden || false,
          ...(rw.height !== undefined ? { height: rw.height } : {}),
          ...(rw.customHeight !== undefined ? { customHeight: rw.customHeight } : {})
        })),
        merges: r.merges || [],
        hyperlinks: r.hyperlinks || [],
        conditionalRules: r.conditionalRules || [],
        autoFilter: r.autoFilter,
        autoFilterColumns: r.autoFilterColumns || [],
        dataValidations: r.dataValidations || [],
        sheetProtection: r.sheetProtection,
        rowBreaks: r.rowBreaks || [],
        colBreaks: r.colBreaks || [],
        protectedRanges: r.protectedRanges || [],
        pageSetup: r.pageSetup,
        // 关键：renderer 靠这两个字段判断流式加载是否完成（停止「正在流式加载」提示）。
        // 引擎行协议序列化用 camelCase（serde rename_all），两个都有值。
        indexedThroughRow: r.indexedThroughRow ?? null,
        indexingComplete: r.indexingComplete ?? false
      }
    },
    readWorkbookFormulaCells: async (request: any) => {
      const cmd = {
        version: 1, requestId: `fc-${++wSeq}`, command: 'read_formula_cells',
        sessionId: request.sessionId, sheetId: request.sheetId
      }
      const resp = await wCmd(JSON.stringify(cmd))
      return resp?.ok ? resp.result : { cells: [] }
    },
    readWorkbookMedia: async () => null,
    readWorkbookComments: async () => [],
    findWorkbookCells: async () => ({ results: [], totalMatches: 0 }),
    readWorkbookRowOutline: async () => [],

    // ---- 偏好 ----
    getLanguage: async () => prefGet('lang', 'zh'),
    getSystemLocale: async () => navigator.language || 'zh-CN',
    onLanguageChanged: () => () => {},
    getTheme: async () => prefGet('theme', 'light'),
    onThemeChanged: () => () => {},
    getDocumentTheme: async () => 'follow',
    onDocumentThemeChanged: () => () => {},
    // shortcut: 自动保存开着自己会每 30s 触发一次注定失败的保存，M9 编辑落盘后恢复 'on'
    getAutoSaveDefault: async () => 'off',
    onAutoSaveDefaultChanged: () => () => {},
    getAiPanelPrefs: async () => ({ textSize: 'md', spellcheck: true }),
    setAiPanelPrefs: async (p: unknown) => p,
    onAiPanelPrefsChanged: () => () => {},
    onChromePressed: () => () => {},
    getAuthorName: async () => prefGet('author', '轻舟用户'),
    getUserDisplayName: async () => prefGet('author', '轻舟用户'),
    onAuthorNameChanged: () => () => {},
    getRecentFiles: async () => [],
    onSidecarCrashed: () => () => {},
    updateWorkbookPassword: async () => ({ ok: false }),

    // ---- 保存 ----
    // 真实保存路径是 saveWorkbookEdits（renderer 只调它，save-actions.ts:783/871）：
    // 增量 journal 需引擎侧合并产新字节，web Worker 宿主尚无此层（M9）。
    // 返回 { canceled: true } 是 workbookSaveResultSchema 的合法分支，
    // 走 renderer 现成的「已取消」UX；缺字段对象会让 result.file.sha256 抛 TypeError。
    saveWorkbookEdits: async () => ({ canceled: true }),

    // ---- 杂项 ----
    fontMetrics: async () => null,
    getAiSettings: async () => ({
      provider: 'genspark', providers: { genspark: { apiKey: '', model: '' } },
      gskToolsEnabled: false, media: {}, search: {}
    }),
    setAiSettings: async () => {},
    onAiSettingsChanged: () => () => {},
    clipboardProbe: async () => false,
    onFontsChanged: () => () => {},
    onHistoryChanged: () => () => {},
    onDeckChanged: () => () => {},
  }

  ;(window as any).desktopApi = new Proxy(api, {
    get(t, p) {
      if (p in t) return t[p as string]
      if (p === 'then') return undefined
      console.warn('[sheets-fallback]', String(p))
      return callable
    }
  })

  // projectApi / aiNamingApi
  ;(window as any).projectApi = {
    resolveChat: async (args: { tempChatId?: string }) => ({ projectId: 'web', chatId: args?.tempChatId || 't' }),
    appendChat: async () => {}, loadChat: async () => [],
    rebindChat: async (args: { chatId?: string }) => ({ projectId: 'web', chatId: args?.chatId || 't' }),
    listChats: async () => [], deleteChat: async () => {}
  }
  ;(window as any).aiNamingApi = { suggest: async () => null }
  ;(window as any).nativeSpell = { available: false }

  console.info('[sheets-shim] DesktopApi 浏览器垫片已激活', pendingSrc ? '带文档:' + pendingName : '空白')

  // ---- 品牌覆盖（eos 嵌入形态 ?embed=1）：Genspark → Dargo（DOM 文本层，源仓不动）----
  if (q.get('embed') === '1') {
    const rebrand = (root: Node) => {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      const hits: Text[] = []
      while (walker.nextNode()) {
        const n = walker.currentNode as Text
        if (n.nodeValue.includes('Genspark') && !/^(script|style)$/i.test((n.parentElement as HTMLElement)?.tagName || ''))
          hits.push(n)
      }
      for (const n of hits) n.nodeValue = n.nodeValue.replace(/Genspark/g, 'Dargo')
    }
    const startRebrand = () => {
      rebrand(document.body)
      new MutationObserver((muts) => {
        for (const m of muts) {
          if (m.type === 'childList') for (const n of m.addedNodes) if (n.nodeType === 1) rebrand(n)
          else if (m.type === 'characterData' && m.target.nodeValue?.includes('Genspark'))
            ;(m.target as Text).nodeValue = m.target.nodeValue.replace(/Genspark/g, 'Dargo')
        }
      }).observe(document.body, { childList: true, subtree: true, characterData: true })
      console.info('[shim] 品牌覆盖激活（Genspark → Dargo）')
    }
    if (document.body) startRebrand()
    else document.addEventListener('DOMContentLoaded', startRebrand)
  }
})()
