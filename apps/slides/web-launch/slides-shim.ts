/**
 * GenOffice Slides renderer 的浏览器垫片（模块版）：
 * 与 Docs 垫片同套路（web-desktop-shim 的三连修经验全部内置），差异点：
 * - 桥名是 slidesApi（编辑操作大 API）+ desktop（文件服务）+ projectApi + aiNamingApi
 * - openPptx/openPptxPath 直接用 **vendor 的 pptx-engine + pptx-render** 在页面内实现
 *   （Electron main 只是薄转发，见 apps/slides/src/main/slides-main.ts:935-975）
 * - 其余编辑操作先走「万能可调用代理」（真 thenable 语义），UI 可渲染；
 *   落盘编辑（savePptx 等）按需逐个实装
 */
import { Buffer } from 'buffer'
import { openPptx } from '@genoffice/pptx-engine'
import { buildRenderSlide, type RenderSlide, type OpenedPptx } from '@genoffice/pptx-render'

;(() => {
  if ((window as any).slidesApi) return
  // pptx-engine 打开链使用 Node Buffer：注入 polyfill（与 eos-product 的
  // vite-plugin-node-polyfills 同款，此 dev server 无该插件）
  ;(globalThis as any).Buffer = Buffer
  const q = new URLSearchParams(location.search)
  const LS = 'genoffice-slides-shim:'
  const prefGet = (k: string, f: unknown) => {
    try {
      const v = localStorage.getItem(LS + k)
      return v === null ? f : JSON.parse(v)
    } catch {
      return f
    }
  }
  let pendingDoc: { src: string; name: string } | null = q.get('src')
    ? { src: q.get('src')!, name: q.get('name') || 'presentation.pptx' }
    : null
  let consumed = false
  let originalBytes: Uint8Array | null = null // save v1：回存原始字节（编辑落盘=M9）

  // ---- 万能可调用代理（Docs 垫片三连修后的最终形态）----
  // 调用返回自身、属性访问返回自身；then() 返回真 Promise 且 resolve 到非 thenable 终值
  const callable: any = new Proxy(function () {} as any, {
    get(_t, p) {
      if (p === Symbol.toPrimitive || typeof p === 'symbol') return undefined
      if (p === 'then') {
        return (onFulfilled: (v: unknown) => unknown) =>
          Promise.resolve(onFulfilled ? onFulfilled(undefined) : undefined)
      }
      if (p === 'catch') return (onRejected: any) => Promise.resolve(undefined)
      if (p === 'finally') return (cb: any) => Promise.resolve(cb ? cb() : undefined)
      return callable
    },
    apply() {
      return callable
    }
  })

  /** main 侧 buildAllRenderSlides 的页面内等价（session-state.ts:427）：
   *  buildRenderSlide(slide, size, { fitWidthPx, media, slideNo })；
   *  metrics 省略 → HeuristicMetrics；media 简化为 base64 data-url（main 版含 SVG 主题
   *  着色上下文，浏览器版先不追平，SVG 图片可能缺主题色） */
  function mediaResolver(opened: OpenedPptx, slidePath?: string) {
    return (ref: string): string | undefined => {
      const bytes = (opened as any).archive.readBytes(ref)
      if (!bytes) return undefined
      const ext = (ref.split('.').pop() || '').toLowerCase()
      const mime =
        ext === 'png' ? 'image/png'
        : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg'
        : ext === 'gif' ? 'image/gif'
        : ext === 'webp' ? 'image/webp'
        : ext === 'svg' ? 'image/svg+xml'
        : 'application/octet-stream'
      let bin = ''
      const CH = 0x8000
      for (let i = 0; i < bytes.length; i += CH) {
        bin += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + CH)) as number[])
      }
      return `data:${mime};base64,${btoa(bin)}`
    }
  }
  function buildAllRenderSlides(opened: OpenedPptx, fitWidthPx: number): RenderSlide[] {
    return (opened.deck.slides as any[]).map((slide, i) =>
      buildRenderSlide(slide, opened.deck.size, {
        fitWidthPx,
        media: mediaResolver(opened, slide.path),
        slideNo: i + 1
      })
    )
  }

  async function openFromPending(fitWidthPx: number) {
    if (!pendingDoc) return null
    console.info('[slides-open] fetch 开始', pendingDoc.src.slice(0, 80))
    const res = await fetch(pendingDoc.src)
    if (!res.ok) throw new Error(`slides shim: 加载失败 ${res.status}`)
    const bytes = new Uint8Array(await res.arrayBuffer())
    console.info('[slides-open] 字节就绪', bytes.length, 'openPptx 调用…')
    const opened = await openPptx(bytes)
    console.info('[slides-open] openPptx 返回, slides=', opened.deck.slides.length)
    const slides = buildAllRenderSlides(opened, fitWidthPx || 960)
    console.info('[slides-shim] openPptx ✅', pendingDoc.name, `${slides.length} 页`)
    return {
      path: 'web://' + pendingDoc.name,
      slides,
      size: { cx: opened.deck.size.cx, cy: opened.deck.size.cy }
    }
  }

  const slidesApi: Record<string, unknown> = {
    consumePendingOpen: async (fitWidthPx: number) => {
      if (consumed) return null
      consumed = true
      if (!pendingDoc) return null
      try {
        const r = await openFromPending(fitWidthPx)
        pendingDoc = null
        console.info('[slides-open] consumePendingOpen 成功')
        // 首帧居中修正：文档挂载后触发 resize 让画布重新 fit（否则初见左上角，
        // 用户操作一下才居中——fit 计算晚于首帧渲染）
        setTimeout(() => window.dispatchEvent(new Event('resize')), 120)
        setTimeout(() => window.dispatchEvent(new Event('resize')), 400)
        return r
      } catch (e) {
        console.error('[slides-open] consumePendingOpen 失败', e, (e as Error)?.stack)
        throw e
      }
    },
    newBlank: async (_fitWidthPx: number) => {
      consumed = true
      return null as unknown as { path: string; slides: RenderSlide[]; size: { cx: number; cy: number } }
    },
    openPptx: async (fitWidthPx: number) => {
      const r = await openFromPending(fitWidthPx)
      pendingDoc = null
      return r
    },
    openPptxPath: async (_path: string, fitWidthPx: number) => {
      const r = await openFromPending(fitWidthPx)
      pendingDoc = null
      return r
    },
    getLanguage: async () => prefGet('lang', 'zh'),
    getSystemLocale: async () => navigator.language || 'zh-CN',
    getTheme: async () => prefGet('theme', 'light'),
    onThemeChanged: () => () => {},
    onLanguageChanged: () => () => {},
    onChromePressed: () => () => {},
    setShowFullScreen: async () => {},
    getRecentFiles: async () => [],
    getAutoSaveDefault: async () => 'on' as const,
    setAutoSavePref: async () => {},
    getAiPanelPrefs: async () => ({ textSize: 'md' as const, spellcheck: true }),
    setAiPanelPrefs: async (p: unknown) => p,
    onAiPanelPrefsChanged: () => () => {},
    getAiSettings: async () => ({
      provider: 'genspark',
      providers: { genspark: { apiKey: '', model: '' } },
      gskToolsEnabled: false,
      media: {},
      search: {}
    }),
    setAiSettings: async () => {},
    onAiSettingsChanged: () => () => {},
    clipboardProbe: async () => false,
    onFontsChanged: () => () => {},
    onHistoryChanged: () => () => {},
    onDeckChanged: () => () => {},
    onOpened: () => () => {},
    onRenamed: () => () => {},
    onCloseSaveRequest: () => () => {},
    onMenuCommand: () => () => {},
    getSlideSize: async () => null,
    // 幻灯片属性族（渲染期被 .length/迭代消费，必须真形状，兜底 undefined 会崩）
    privateFontFaces: async () => [],
    privateFontData: async () => null,
    fontCatalog: async () => [],
    fontMissing: async () => [],
    getLayouts: async () => null,
    getNotes: async () => '',
    getComments: async () => [],
    getSections: async () => [],
    getTransition: async () => null,
    getAnimations: async () => [],
    aiGskStatus: async () => ({ loggedIn: false }) as unknown,
    // save v1：回传宿主（MinIO）保存原始字节；带回执。编辑操作经引擎落盘 = M9
    save: async () => {
      if (window.parent === window) return { ok: false, error: 'web-shim: 无宿主' }
      const bytes = originalBytes
      if (!bytes) return { ok: false, error: '无已加载文档' }
      return await new Promise((resolve) => {
        const tag = 'save-' + Math.random().toString(36).slice(2)
        const onReply = (ev: MessageEvent) => {
          if (ev.data?.type === 'genoffice:save:result' && ev.data.tag === tag) {
            window.removeEventListener('message', onReply)
            resolve({ ok: !!ev.data.ok, error: ev.data.error })
          }
        }
        window.addEventListener('message', onReply)
        window.parent.postMessage(
          { type: 'genoffice:save', tag, name: pendingDoc?.name || 'presentation.pptx', bytes },
          '*'
        )
        setTimeout(() => {
          window.removeEventListener('message', onReply)
          resolve({ ok: false, error: '宿主保存超时' })
        }, 30000)
      })
    },
    fontDownload: async () => ({ ok: false, error: 'web-shim' }),
    fontInstallLocal: async () => ({ families: [] as string[] })
  }
  ;(window as any).slidesApi = new Proxy(slidesApi, {
    get(t, p) {
      if (p in t) return t[p as string]
      if (p === 'then') return undefined
      console.warn('[slides-shim] fallback 调用:', String(p))
      return callable
    }
  })

  // desktop(filesApi)/projectApi/aiNamingApi：与 Docs 同款最小实现
  ;(window as any).desktop = new Proxy(
    {
      getLanguage: async () => prefGet('lang', 'zh'),
      getTheme: async () => 'light',
      onThemeChanged: () => () => {},
      onLanguageChanged: () => () => {},
      getAuthorName: async () => prefGet('author', '轻舟用户'),
      onAuthorNameChanged: () => () => {},
      getAiPanelPrefs: async () => ({ textSize: 'md', spellcheck: true }),
      setAiPanelPrefs: async (p: unknown) => p,
      onAiPanelPrefsChanged: () => () => {},
      armContextMenu: () => {},
      onContextMenuRequest: () => () => {},
      addPastedImage: async () => null,
      readAttachment: async () => null,
      readAttachmentImage: async () => null,
      getPathForFile: () => ''
    },
    {
      get(t, p) {
        if (p in t) return t[p as string]
        if (p === 'then') return undefined
        return callable
      }
    }
  )
  ;(window as any).projectApi = {
    resolveChat: async (args: { tempChatId?: string }) => ({ projectId: 'web', chatId: args?.tempChatId || 't' }),
    appendChat: async () => {},
    loadChat: async () => [],
    rebindChat: async (args: { chatId?: string }) => ({ projectId: 'web', chatId: args?.chatId || 't' }),
    listChats: async () => [],
    deleteChat: async () => {}
  }
  ;(window as any).aiNamingApi = { suggest: async () => null }
  console.info('[slides-shim] SlidesApi 浏览器垫片已激活', pendingDoc ? '带文档:' + pendingDoc.name : '空白')
})()
