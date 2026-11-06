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
      if (p === 'then')
        return (onFulfilled: (v: unknown) => unknown) =>
          Promise.resolve(onFulfilled ? onFulfilled(undefined) : undefined)
      if (p === 'catch') return () => Promise.resolve(undefined)
      if (p === 'finally') return (cb: any) => Promise.resolve(cb ? cb() : undefined)
      return callable
    },
    apply() {
      return callable
    }
  })

  let pendingSrc: string | null = q.get('src')
  let selectInFlight: Promise<unknown> | null = null
  let pendingName: string = q.get('name') || 'workbook.xlsx'
  let sessionMeta: any = null // open 应答回存的 sheets 元数据

  const api: Record<string, unknown> = {
    // ---- 打开 ----
    hasQueuedWorkbook: async () => {
      // React StrictMode 双挂载防竞争：等 in-flight 的 selectWorkbook 完成再答
      // （否则第二次 hasQueuedWorkbook=true → 第二次 selectWorkbook=null → 取消第一次）
      if (selectInFlight) await selectInFlight
      return !!pendingSrc
    },
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
          name: pendingName.replace(/\.[^.]+$/, ''),
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
          workbookProtection: r.workbookProtection
        }
        console.warn('[TRACE] selectWorkbook returning:', typeof result, result?.name, 'sheets:', result?.sheets?.length)
        return result
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
        indexedThroughRow: r.indexedThroughRow,
        indexingComplete: r.indexingComplete
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
    getAutoSaveDefault: async () => 'on',
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
    saveWorkbook: async () => ({ ok: false, error: 'web-shim: 保存链未实装' }),
    saveWorkbookAs: async () => ({ ok: false, error: 'web-shim' }),

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

  // 包装 selectWorkbook 为 in-flight 追踪版
  const rawSelect = api.selectWorkbook
  api.selectWorkbook = (...args: unknown[]) => {
    const p = Promise.resolve(rawSelect(...(args as []))).finally(() => {
      if (selectInFlight === p) selectInFlight = null
    })
    selectInFlight = p
    return p
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
})()
