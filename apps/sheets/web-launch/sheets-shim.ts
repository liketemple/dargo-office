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
// M9 保存链（全浏览器，无 Rust）：规划器 + JSZip 组装器均来自上游 xlsx-gateway
// 的内存实现（assembleWithJsZip 注释明说 sidecar 流式才是大文件路径）。
// 懒加载：该链顶层带 node: 命名导入（经 vite 别名到 shims）且模块图庞大，
// 静态导入会拖死首屏——只在首次保存时按需拉取
type Gateway = typeof import('@genoffice/xlsx-gateway/gateway/xlsx-gateway')
let gatewayPromise: Promise<Gateway> | null = null
const loadGateway = (): Promise<Gateway> =>
  (gatewayPromise ??= import('@genoffice/xlsx-gateway/gateway/xlsx-gateway'))

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
  let currentName = 'workbook.xlsx'
  let originalBytes: Uint8Array | null = null // M9 保存的基准字节（保存后换新）
  let sheetNames = new Map<string, string>() // sheetId → 文件内表名（保存映射用）

  /** open 应答 → workbookFile（与 Electron 主进程 openWorkbooksFromPaths 同构：
   *  schema.parse 补齐引擎 serde 省略的默认字段） */
  async function buildWorkbookFile(r: any, bytes: Uint8Array | ArrayBuffer, name: string) {
    const digest = await crypto.subtle.digest('SHA-256', bytes)
    const sha256 = Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, '0')).join('')
    return workbookFileSchema.parse({
      sessionId: r.sessionId,
      name,
      path: name,
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
      unsavedNew: false,
      date1904: r.date1904 ?? false,
      shortDateFormat: r.shortDateFormat,
      themeColors: r.themeColors,
      themeFonts: r.themeFonts
    })
  }

  /** M9 保存：宿主 postMessage（带回执，与 docs 垫片同协议） */
  function saveViaHost(bytes: Uint8Array, name: string): Promise<{ ok: boolean; error?: string }> {
    return new Promise((resolve) => {
      if (window.parent === window) return resolve({ ok: false, error: '无宿主（独立调试形态）' })
      const tag = 'save-' + Math.random().toString(36).slice(2)
      const onReply = (ev: MessageEvent) => {
        if (ev?.data?.type === 'file-editor:save:result' && ev.data.tag === tag) {
          window.removeEventListener('message', onReply)
          resolve({ ok: !!ev.data.ok, error: ev.data.error })
        }
      }
      window.addEventListener('message', onReply)
      window.parent.postMessage({ type: 'file-editor:save', tag, name, bytes }, '*')
      setTimeout(() => {
        window.removeEventListener('message', onReply)
        resolve({ ok: false, error: '宿主保存超时（编辑器宿主页未响应）' })
      }, 30000)
    })
  }

  /** renderer 保存请求 → planner 入参映射（sheets-main writeWorkbookTo 的浏览器同构，
   *  去掉对话框/磁盘守卫/加密——那些是 Electron 专属语义） */
  function mapSaveRequest(request: any) {
    const addedSheetNames = new Map<string, string>()
    const duplicateSources = new Map<string, string>()
    const renames: { sheetName: string; newName: string }[] = []
    const removals: string[] = []
    const hiddenChanges: { sheetName: string; hidden: boolean }[] = []
    const tabColorStates: { sheetName: string; color: string | null }[] = []
    let orderChanged = false
    for (const op of request.sheetOps ?? []) {
      if (op.kind === 'add-sheet') { addedSheetNames.set(op.sheetId, op.name); continue }
      if (op.kind === 'duplicate-sheet') {
        const sourceName = sheetNames.get(op.sourceSheetId)
        if (!sourceName) throw new Error('Unknown duplicate source ' + op.sourceSheetId + '.')
        addedSheetNames.set(op.sheetId, op.name)
        duplicateSources.set(op.sheetId, sourceName)
        continue
      }
      if (op.kind === 'reorder-sheets') { orderChanged = true; continue }
      const sheetName = addedSheetNames.get(op.sheetId) ?? sheetNames.get(op.sheetId)
      if (!sheetName) throw new Error('Unknown worksheet ' + op.sheetId + '.')
      if (op.kind === 'rename-sheet') renames.push({ sheetName, newName: op.newName })
      else if (op.kind === 'set-sheet-hidden') hiddenChanges.push({ sheetName, hidden: op.hidden })
      else if (op.kind === 'set-sheet-tab-color') tabColorStates.push({ sheetName, color: op.color })
      else removals.push(sheetName)
    }
    const renameByOriginal = new Map(renames.map((rn) => [rn.sheetName, rn.newName]))
    const resolveSheetName = (sheetId: string): string => {
      const n = addedSheetNames.get(sheetId) ?? sheetNames.get(sheetId)
      if (!n) throw new Error('Unknown worksheet ' + sheetId + '.')
      return n
    }
    const sheetPlan =
      (request.sheetOps ?? []).some((op: any) => op.kind !== 'set-sheet-tab-color')
        ? {
            renames,
            additions: [...addedSheetNames].map(([sheetId, name]) => ({
              name,
              sourceSheetName: duplicateSources.get(sheetId)
            })),
            removals,
            hiddenChanges,
            orderChanged,
            order: (request.sheetOrder ?? []).map((sheetId: string) => {
              const original = resolveSheetName(sheetId)
              return addedSheetNames.has(sheetId)
                ? original
                : renameByOriginal.get(original) ?? original
            })
          }
        : undefined
    const edits = (request.edits ?? []).map((edit: any) => ({
      sheetName: resolveSheetName(edit.sheetId),
      row: edit.row,
      column: edit.column,
      writeValue: edit.writeValue,
      cell: { value: edit.value, formula: edit.formula },
      style: edit.style,
      rich: edit.rich,
      styleReset: edit.styleReset
    }))
    const bulkConstantFills = (request.bulkConstantFills ?? []).map(({ sheetId, ...fill }: any) => ({
      sheetName: resolveSheetName(sheetId),
      ...fill
    }))
    const opsBySheet = new Map<string, any[]>()
    for (const op of request.structuralOps ?? []) {
      const sheetName = resolveSheetName(op.sheetId)
      const sheetOps = opsBySheet.get(sheetName) ?? []
      if ('range' in op) sheetOps.push({ kind: op.kind, range: op.range })
      else if ('size' in op) sheetOps.push({ kind: op.kind, start: op.start, end: op.end, size: op.size })
      else if ('level' in op)
        sheetOps.push({
          kind: op.kind, start: op.start, end: op.end, level: op.level,
          ...(op.collapsed === undefined ? {} : { collapsed: op.collapsed })
        })
      else if ('summaryBelow' in op)
        sheetOps.push({ kind: op.kind, summaryBelow: op.summaryBelow, summaryRight: op.summaryRight })
      else if ('hidden' in op) sheetOps.push({ kind: op.kind, start: op.start, end: op.end, hidden: op.hidden })
      else if ('style' in op) sheetOps.push({ kind: op.kind, start: op.start, end: op.end, style: op.style })
      else if ('before' in op) sheetOps.push({ kind: op.kind, index: op.index, count: op.count, before: op.before })
      else sheetOps.push({ kind: op.kind, index: op.index, count: op.count })
      opsBySheet.set(sheetName, sheetOps)
    }
    const structuralOps = [...opsBySheet].map(([sheetName, ops]) => ({ sheetName, ops }))
    const linksBySheet = new Map<string, { row: number; column: number; target: string | null }[]>()
    for (const link of request.hyperlinkEdits ?? []) {
      const sheetName = resolveSheetName(link.sheetId)
      const list = linksBySheet.get(sheetName) ?? []
      list.push({ row: link.row, column: link.column, target: link.target })
      linksBySheet.set(sheetName, list)
    }
    const bySheet = (items: any[] = []) =>
      items.map(({ sheetId, ...rest }: any) => ({ sheetName: resolveSheetName(sheetId), ...rest }))
    const formulaValuesBySheet = new Map<string, any[]>()
    for (const cell of request.formulaValues ?? []) {
      const sheetName = resolveSheetName(cell.sheetId)
      const list = formulaValuesBySheet.get(sheetName) ?? []
      list.push({ row: cell.row, column: cell.column, value: cell.value })
      formulaValuesBySheet.set(sheetName, list)
    }
    const pivotAdditions = (request.pivotAdditions ?? []).map((pv: any) => ({
      sheetName: resolveSheetName(pv.sheetId),
      sourceSheetName: resolveSheetName(pv.sourceSheetId),
      ...pv,
      sheetId: undefined,
      sourceSheetId: undefined
    }))
    return {
      edits,
      structuralOps,
      chartEdits: request.chartEdits ?? [],
      sheetPlan,
      filterStates: (request.filterStates ?? []).map((st: any) => ({
        sheetName: resolveSheetName(st.sheetId), filter: st.filter, hiddenRows: st.hiddenRows, visibilityRange: st.visibilityRange
      })),
      hyperlinkEdits: [...linksBySheet].map(([sheetName, links]) => ({ sheetName, edits: links })),
      cfStates: (request.cfStates ?? []).map((st: any) => ({ sheetName: resolveSheetName(st.sheetId), rules: st.rules })),
      dvStates: (request.dvStates ?? []).map((st: any) => ({ sheetName: resolveSheetName(st.sheetId), rules: st.rules })),
      sheetProtections: bySheet(request.sheetProtections),
      definedNamesState: request.definedNamesState ?? null,
      visualAdditions: (request.visualAdditions ?? []).map((v: any) => ({
        sheetName: resolveSheetName(v.sheetId), anchor: v.anchor, chart: v.chart, shape: v.shape, image: v.image
      })),
      pageSetupStates: bySheet(request.pageSetupStates),
      noteStates: (request.noteStates ?? []).map(({ sheetId, notes }: any) => ({ sheetName: resolveSheetName(sheetId), notes })),
      tableAdditions: (request.tableAdditions ?? []).map((tb: any) => ({
        sheetName: resolveSheetName(tb.sheetId), area: tb.area, name: tb.name, columnNames: tb.columnNames,
        style: tb.style, bandedRows: tb.bandedRows,
        options: {
          headerRow: tb.headerRow, totalsRow: tb.totalsRow, firstColumn: tb.firstColumn,
          lastColumn: tb.lastColumn, bandedColumns: tb.bandedColumns, filterButton: tb.filterButton
        }
      })),
      pivotAdditions,
      pivotCacheRefreshPaths: request.pivotCacheRefreshPaths ?? [],
      pivotRefreshUpdates: request.pivotRefreshUpdates ?? [],
      visualEdits: request.visualEdits ?? [],
      sparklineAdditions: bySheet(request.sparklineAdditions),
      formulaValues: [...formulaValuesBySheet].map(([sheetName, cells]) => ({ sheetName, cells })),
      themeState: request.themeState ?? null,
      workbookProtectionState: request.workbookProtectionState ?? null,
      protectedRangeStates: (request.protectedRangeStates ?? []).map((st: any) => ({
        sheetName: resolveSheetName(st.sheetId), ranges: st.ranges
      })),
      bulkConstantFills,
      tabColorStates,
      tableEdits: (request.tableEdits ?? []).map(({ sheetId, ...edit }: any) => ({ ...edit, sheetName: resolveSheetName(sheetId) })),
      cellImageAdditions: bySheet(request.cellImageAdditions),
      externalLinkOps: request.externalLinkOps ?? null,
      redactionStates:
        (request.redactionStates ?? null)?.map((state: any) => ({
          sheetName: state.sheetName,
          marks: state.marks.map((mark: any) => ({
            startRow: mark.startRow, endRow: mark.endRow, startColumn: mark.startColumn,
            endColumn: mark.endColumn, label: mark.label,
            ...('previousFill' in mark ? { previousFill: mark.previousFill ?? null } : {})
          }))
        })) ?? null,
      slicerState:
        request.slicerState === null || request.slicerState === undefined
          ? null
          : {
              slicers: request.slicerState.slicers.map(({ sheetId, ...control }: any) => ({
                ...control, sheetName: resolveSheetName(sheetId)
              })),
              removals: [...request.slicerState.removals]
            }
    }
  }

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
        currentName = pendingName
        originalBytes = new Uint8Array(bytes)
        sheetNames = new Map((r.sheets ?? []).map((sh: any) => [sh.id, sh.name]))
        pendingSrc = null
        return buildWorkbookFile(r, bytes, currentName)
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
    // M9 已落盘：自动保存恢复（30s + 失焦），与桌面语义一致
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

    // ---- 保存（M9：全浏览器链，规划器+JSZip 组装器来自上游 xlsx-gateway 内存实现）----
    saveWorkbookEdits: async (request: any) => {
      if (!originalBytes || !sessionMeta) return { canceled: true }
      if (request.mode !== 'save') return { canceled: true } // save-as/恢复等 Electron 专属语义
      const args = mapSaveRequest(request)
      console.info('[sheets-shim] 保存规划开始', (request.edits ?? []).length, '处编辑')
      const { planCellEditsToXlsx, createBufferEntrySource, assembleWithJsZip } = await loadGateway()
      const source = await createBufferEntrySource(originalBytes)
      const plan = await planCellEditsToXlsx(
        source,
        args.edits,
        args.structuralOps,
        args.chartEdits,
        args.sheetPlan,
        args.filterStates,
        args.hyperlinkEdits,
        args.cfStates,
        args.dvStates,
        args.sheetProtections,
        args.definedNamesState,
        args.visualAdditions,
        args.pageSetupStates,
        args.noteStates,
        args.tableAdditions,
        args.pivotAdditions,
        args.pivotCacheRefreshPaths,
        args.pivotRefreshUpdates,
        args.visualEdits,
        args.sparklineAdditions,
        args.formulaValues,
        args.themeState,
        args.workbookProtectionState,
        args.protectedRangeStates,
        args.bulkConstantFills,
        args.tabColorStates,
        args.tableEdits,
        args.cellImageAdditions,
        args.externalLinkOps,
        args.redactionStates,
        args.slicerState
      )
      const mutation = await assembleWithJsZip(originalBytes, plan)
      const bytes = new Uint8Array(mutation.buffer)
      console.info('[sheets-shim] 组装完成', bytes.byteLength, 'B，touched', plan.touchedEntries.length, '项')
      // 经宿主落盘（带回执）；失败抛错走 renderer 现成的保存失败 UX
      const host = await saveViaHost(bytes, currentName)
      if (!host.ok) throw new Error(host.error || '宿主保存失败')
      // 会话换新：旧 worker 会话关掉，按保存后字节重开（renderer 拿到新 workbookFile）
      wCmd(JSON.stringify({ version: 1, requestId: 'close-' + Date.now(), command: 'close', sessionId: sessionMeta.sessionId })).catch(() => {})
      const reopen = await wOpen(bytes.buffer as ArrayBuffer, currentName)
      if (!reopen?.ok) throw new Error(reopen?.error?.message || '保存后重开会话失败')
      originalBytes = bytes
      sessionMeta = reopen.result
      sheetNames = new Map((reopen.result.sheets ?? []).map((sh: any) => [sh.id, sh.name]))
      const file = await buildWorkbookFile(reopen.result, bytes, currentName)
      return { canceled: false, file, touchedEntries: [...plan.touchedEntries] }
    },

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
