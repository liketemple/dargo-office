/**
 * Dargo Docs renderer（原 GenOffice）的浏览器垫片：实现 DesktopApi 的最小可运行子集。
 * - 文档注入：?src=<编码URL>&name=<文件名> → consumePendingOpenDocx 返回该文件
 * - 保存：saveDocx/saveDocxAs → postMessage 给宿主（父窗口）+ 可选 Blob 下载
 * - 锁定模式：?readonly=1 → 只读打开（禁输入/禁保存/禁选区复制）
 * - 偏好：localStorage 持久化（语言/主题/作者名/自动保存）
 * - 未实现的方法：安全 no-op / reject，renderer 已有降级路径
 */
(function () {
  if (window.desktop) return
  const q = new URLSearchParams(location.search)
  const LS = 'genoffice-webshim:'
  const prefGet = (key, fallback) => {
    try { const v = localStorage.getItem(LS + key); return v === null ? fallback : JSON.parse(v) } catch { return fallback }
  }
  const prefSet = (key, v) => { try { localStorage.setItem(LS + key, JSON.stringify(v)) } catch {} }
  const listeners = { language: [], theme: [], docTheme: [] }
  let pendingDoc = null
  let consumed = false
  if (q.get('src')) {
    pendingDoc = { src: q.get('src'), name: q.get('name') || 'document.docx' }
  }
  // 锁定模式（宿主 &readonly=1）：只读打开。
  // shortcut: 工具栏个别直改文档模型的动作拦不净，renderer 原生只读模式待后续里程碑
  const locked = q.get('readonly') === '1'
  if (locked) {
    const style = document.createElement('style')
    style.textContent =
      '[contenteditable="true"]{pointer-events:none!important;user-select:none!important;}'
    ;(document.head || document.documentElement).appendChild(style)
    const blockEditable = (e) => {
      const t = e.target
      if (t && (t.isContentEditable || t.closest?.('[contenteditable="true"]'))) {
        e.preventDefault()
        e.stopPropagation()
      }
    }
    document.addEventListener('beforeinput', blockEditable, true)
    document.addEventListener('keydown', blockEditable, true)
    console.info('[web-shim] 只读锁定模式（readonly=1）')
  }
  const notify = (kind, payload) => listeners[kind]?.forEach((h) => { try { h(payload) } catch {} })

  async function buildOpenResult() {
    if (!pendingDoc) return null
    console.info('[shim-fetch] start', pendingDoc.src.slice(0, 90))
    const res = await fetch(pendingDoc.src).catch((e) => {
      console.warn('[shim-fetch] failed', String(e))
      throw e
    })
    console.info('[shim-fetch] status', res.status)
    if (!res.ok) throw new Error('load failed: ' + res.status)
    const blob = await res.blob()
    console.info('[shim-fetch] blob', blob.size)
    const dataUrl = URL.createObjectURL(blob)
    const buf = await blob.arrayBuffer()
    const digest = await crypto.subtle.digest('SHA-256', buf)
    const hash = Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, '0')).join('')
    return { path: 'web://' + pendingDoc.name, name: pendingDoc.name, dataUrl, hash }
  }

  const saveViaHost = async (data, suggestedName) => {
    // 保存语义（2026-11-06 拍板）：仅经宿主回传服务器（MinIO），**禁止浏览器下载**。
    // 带回执协议：宿主处理完（上传成功/失败）回复 file-editor:save:result，
    // 编辑器据此显示真实保存状态；无宿主（独立 :5199 调试形态）才退化为下载。
    if (locked) return { ok: false, error: '只读模式：文档已锁定，不可保存' }
    if (window.parent !== window) {
      return await new Promise((resolve) => {
        const tag = 'save-' + Math.random().toString(36).slice(2)
        const onReply = (ev) => {
          if (ev?.data?.type === 'file-editor:save:result' && ev.data.tag === tag) {
            window.removeEventListener('message', onReply)
            resolve({ ok: !!ev.data.ok, error: ev.data.error })
          }
        }
        window.addEventListener('message', onReply)
        window.parent.postMessage({ type: 'file-editor:save', tag, name: suggestedName, bytes: data }, '*')
        setTimeout(() => {
          window.removeEventListener('message', onReply)
          resolve({ ok: false, error: '宿主保存超时（编辑器宿主页未响应）' })
        }, 30000)
      })
    }
    const blob = new Blob([data], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    })
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = suggestedName || 'document.docx'
    a.click()
    return { ok: true }
  }

  const api = {
    // ---- 偏好 ----
    getLanguage: async () => prefGet('lang', 'zh'),
    getSystemLocale: async () => navigator.language || 'zh-CN',
    onLanguageChanged: (h) => (listeners.language.push(h), () => {}),
    getTheme: async () => prefGet('theme', 'light'), // 2026-11-05 拍板：编辑器主题默认固定浅色
    onThemeChanged: (h) => (listeners.theme.push(h), () => {}),
    getDocumentTheme: async () => prefGet('docTheme', 'follow'), // DocTheme 是字符串枚举 'follow'|'light'|'dark'
    onDocumentThemeChanged: (h) => (listeners.docTheme.push(h), () => {}),
    getAutoSaveDefault: async () => prefGet('autoSave', locked ? 'off' : 'on'), // 锁定模式禁自动保存（免每 30s 保存失败提示）
    onAutoSaveDefaultChanged: () => () => {},
    getAuthorName: async () => prefGet('author', '轻舟用户'),
    onAuthorNameChanged: () => () => {},
    getAiPanelPrefs: async () => prefGet('aiPanel', { textSize: 'md', spellcheck: true }),
    setAiPanelPrefs: async (p) => prefSet('aiPanel', p) || p,
    onAiPanelPrefsChanged: () => () => {},
    onChromePressed: () => () => {},
    // 字体垂直度量（Word 行排版用）：浏览器形态返回 null → renderer 走 UNKNOWN 降级
    fontMetrics: async () => null,
    // AI 设置：最小禁用态（provider=genspark 无 key → AiPanel 校验为不可用，UI 降级）
    getAiSettings: async () => ({
      provider: 'genspark',
      providers: { genspark: { apiKey: '', model: '' } },
      gskToolsEnabled: false,
      media: {},
      search: {}
    }),
    setAiSettings: async () => {},
    kbList: async () => [],
    onAiSettingsChanged: () => () => {},
    getRecentFiles: async () => [],
    getPrinterList: async () => [],
    onOpenPathChanged: () => () => {},
    armContextMenu: () => {},
    // ---- 打开 ----
    openDocx: async () => null, // 浏览器形态：文件选择由宿主负责
    openDocxPath: async () => null,
    confirmDocumentReplace: async () => true,
    openDocxDecrypt: async () => ({ ok: false, reason: 'unsupported' }),
    consumePendingOpenDocx: async () => {
      if (consumed) return null
      consumed = true
      return pendingDoc ? await buildOpenResult() : null
    },
    consumeNewBlankDoc: async () => !pendingDoc,
    consumeAiDocContent: async () => null,
    // ---- 保存 ----
    saveDocx: async (path, data, name) => saveViaHost(data, name || (pendingDoc && pendingDoc.name) || 'document.docx'),
    saveDocxAs: async (data, name) => saveViaHost(data, name),
    saveDocxNew: async (data, name) => saveViaHost(data, name),
    saveDocxTo: async (path, data) => saveViaHost(data, path.split('/').pop()),
    // ---- 通用事件/杂项（安全 no-op） ----
    onOpenDocx: (h) => (listeners.open = listeners.open || [], listeners.open.push(h), () => {}),
    onContextMenuRequest: () => () => {},
    onViewImage: () => () => {},
    saveImageAs: () => {},
    fetchImage: async (src) => { const r = await fetch(src); return r.ok ? src : null },
    exportPdf: async () => ({ ok: false, error: 'web-shim: unsupported' }),
    exportHtml: async () => ({ ok: false, error: 'web-shim: unsupported' }),
    exportMarkdown: async () => ({ ok: false, error: 'web-shim: unsupported' }),
  }
  // 万可调用链式代理：调用返回自身、属性访问返回自身（订阅函数/链式 API 全兼容）；
  // then/toPrimitive 显式 undefined，避免被当 thenable 死循环
  const callable = new Proxy(function () {}, {
    get(_t, p) {
      if (p === Symbol.toPrimitive || typeof p === 'symbol') return undefined
      if (p === 'then') {
        // thenable 协议：then() 必须返回**真 Promise**（调用方会链式 .then().catch()），
        // 且 resolve 到非 thenable 终值（undefined）——resolve(callable) 会造成
        // 「resolve→then→resolve」微任务死循环冻结主线程（实测教训×2）
        return (onFulfilled) => Promise.resolve(onFulfilled ? onFulfilled(undefined) : undefined)
      }
      return callable
    },
    apply() {
      return callable
    }
  })
  // TRACE：每次调用打点，定位启动卡点（验收后移除）
  const traced = {}
  for (const k of Object.keys(api)) {
    traced[k] = function (...args) {
      const t0 = performance.now()
      try {
        const r = api[k].apply(this, args)
        if (r && typeof r.then === 'function') {
          return r.then((v) => { console.info('[shim]', k, (performance.now()-t0).toFixed(0)+'ms'); return v },
                        (e) => { console.warn('[shim-err]', k, String(e)); throw e })
        }
        console.info('[shim]', k, 'sync')
        return r
      } catch (e) { console.warn('[shim-throw]', k, String(e)); throw e }
    }
  }
  window.__shimTrace__ = true
  window.desktop = new Proxy(traced, {
    get(target, prop) {
      if (prop in target) return target[prop]
      if (prop === 'then') return undefined
      // 未实现方法打点（once）：定位未知调用方
      const logged = { v: false }
      const wrapped = new Proxy(function () {}, {
        get(_t, p2) {
          if (p2 === 'then') return (onFulfilled) => Promise.resolve(onFulfilled ? onFulfilled(undefined) : undefined)
          if (p2 === Symbol.toPrimitive || typeof p2 === 'symbol') return undefined
          return callable
        },
        apply() {
          if (!logged.v) { logged.v = true; console.warn('[shim-fallback]', String(prop)) }
          return callable
        }
      })
      return wrapped
    }
  })
  window.aiNamingApi = { suggest: async () => null }
  // ProjectApi（chat 历史持久化，packages/project-store/src/ipc.ts 契约）：
  // 浏览器形态给空实现——loadChat 返回 [] → AiPanel 跳过历史回放
  window.projectApi = {
    resolveChat: async (args) => ({ projectId: 'web', chatId: args?.tempChatId || 'temp' }),
    appendChat: async () => {},
    loadChat: async () => [],
    rebindChat: async (args) => ({ projectId: 'web', chatId: args?.chatId || 'temp' }),
    listChats: async () => [],
    deleteChat: async () => {}
  }
  console.info('[web-shim] GenOffice DesktopApi 浏览器垫片已激活', pendingDoc ? '带文档:' + pendingDoc.name : '空白文档')
})()
