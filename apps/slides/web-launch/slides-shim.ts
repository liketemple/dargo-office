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
// M9b/c：编辑桥直跑 op 事务引擎（@genoffice/pptx-ops，58 op 纯 TS）+ 引擎保存
import { openPptx, savePptx, commitSaved, findGroupChild } from '@genoffice/pptx-engine'
import { runTxn, resolveGroupChildId } from '@genoffice/pptx-ops'
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
  let originalBytes: Uint8Array | null = null // 打开时的原始字节（保存链基准）
  // M9b 编辑会话：deck 模型 + 渲染宽度 + 脏标记（isDirty/保存判定）
  let openedRef: OpenedPptx | null = null
  let fitWidth = 960
  let metaDirty = false
  let currentName = 'presentation.pptx'

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

  // ---- M9b 编辑桥：op 事务引擎胶水（slides-main 同构，去掉 Electron 会话/历史簿记）----
  const EMU_PER_PX_96 = 9525
  const EMU_PER_PT = 12700

  /** 单页重建（事务后返回给 renderer 的 RenderSlide） */
  function rebuildOne(idx: number): RenderSlide | null {
    if (!openedRef) return null
    const slide = (openedRef.deck.slides as any[])[idx]
    if (!slide) return null
    return buildRenderSlide(slide, openedRef.deck.size, {
      fitWidthPx: fitWidth,
      media: mediaResolver(openedRef, slide.path),
      slideNo: idx + 1
    })
  }

  /** 事务执行（数组或 {ops, isolation}；脏标记；undo/redo 簿记 v1 从简——见 shortcut 注释） */
  function sessionTxn(req: any): any | null {
    if (!openedRef) return null
    const r = runTxn(openedRef as any, Array.isArray(req) ? { ops: req } : req)
    if (!r.applied) return null
    metaDirty = true
    return r
  }

  /** px→EMU 变换载荷（slides-main transformPayload 同构，含组子元素坐标系换算） */
  function transformPayload(slideIndex: number, fitWidthPx: number, item: any) {
    if (!openedRef) return null
    const slide = (openedRef.deck.slides as any[])[slideIndex]
    if (!slide) return null
    const childId = item.groupId ? resolveGroupChildId(slide, item.groupId, item.sourceId) : item.sourceId
    const grpChild = item.groupId ? findGroupChild(slide, item.groupId, childId as any) : null
    if (item.groupId && !grpChild) return null
    const baseWidthPx = openedRef.deck.size.cx / EMU_PER_PX_96
    const scale = fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    let box: { x: number; y: number; cx: number; cy: number }
    if (grpChild) {
      const g = (grpChild as any).grp
      const ch = g.childOffset
      const chX = ch?.x ?? g.transform.offset.x
      const chY = ch?.y ?? g.transform.offset.y
      const gExt = g.transform.offset
      const gsx = ch?.cx ? gExt.cx / ch.cx : 1
      const gsy = ch?.cy ? gExt.cy / ch.cy : 1
      box = {
        x: toEmu(item.xPx / gsx) + chX,
        y: toEmu(item.yPx / gsy) + chY,
        cx: toEmu(item.wPx / gsx),
        cy: toEmu(item.hPx / gsy)
      }
    } else {
      box = { x: toEmu(item.xPx), y: toEmu(item.yPx), cx: toEmu(item.wPx), cy: toEmu(item.hPx) }
    }
    return {
      op: 'setTransform' as const,
      target: { slide: slideIndex, el: item.sourceId },
      box,
      rotDeg: item.rotationDeg,
      ...(item.groupId ? { group: item.groupId } : { resizeTableGrid: true })
    }
  }

  /** 全量 SlideUpdate（v1 简化：不做增量 diff，全部重建——shortcut: 大 deck 可按 buildSlideUpdate 增量化） */
  function slideUpdate(): { order: string[]; slides: RenderSlide[] } | null {
    if (!openedRef) return null
    return {
      order: (openedRef.deck.slides as any[]).map((sl) => sl.path),
      slides: buildAllRenderSlides(openedRef, fitWidth)
    }
  }

  /** M9c 保存：引擎序列化 → 宿主 postMessage（带回执）→ commitSaved 清脏 */
  async function saveDeck(): Promise<{ ok: boolean; error?: string; slides?: RenderSlide[] }> {
    if (!openedRef) return { ok: false, error: '无打开的文档' }
    const bytes = await savePptx(openedRef)
    const host = await saveViaHost(bytes)
    if (!host.ok) return { ok: false, error: host.error }
    commitSaved(openedRef)
    metaDirty = false
    originalBytes = bytes
    console.info('[slides-shim] 保存完成', bytes.byteLength, 'B')
    return { ok: true, slides: buildAllRenderSlides(openedRef, fitWidth) }
  }

  function saveViaHost(bytes: Uint8Array): Promise<{ ok: boolean; error?: string }> {
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
      window.parent.postMessage({ type: 'file-editor:save', tag, name: currentName, bytes }, '*')
      setTimeout(() => {
        window.removeEventListener('message', onReply)
        resolve({ ok: false, error: '宿主保存超时（编辑器宿主页未响应）' })
      }, 30000)
    })
  }

  async function openFromPending(fitWidthPx: number) {
    if (!pendingDoc) return null
    console.info('[slides-open] fetch 开始', pendingDoc.src.slice(0, 80))
    const res = await fetch(pendingDoc.src)
    if (!res.ok) throw new Error(`slides shim: 加载失败 ${res.status}`)
    const bytes = new Uint8Array(await res.arrayBuffer())
    console.info('[slides-open] 字节就绪', bytes.length, 'openPptx 调用…')
    currentName = pendingDoc.name
    const opened = await openPptx(bytes)
    console.info('[slides-open] openPptx 返回, slides=', opened.deck.slides.length)
    openedRef = opened
    fitWidth = fitWidthPx || 960
    const slides = buildAllRenderSlides(opened, fitWidth)
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
    // M9 已落盘：自动保存恢复（30s + 失焦），与桌面语义一致
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
    // ---- M9b 编辑桥（op 事务引擎直跑；v1 覆盖高频编辑面，其余仍走万能代理）----
    isDirty: () =>
      !!metaDirty ||
      !!openedRef?.deck.slides.some(
        (sl: any) => sl.structureDirty || sl.elements.some((el: any) => el.dirty || el.dirtyTransform)
      ),
    // shortcut: 撤销/重做 v1 未接（需移植 session-state 历史快照）；批次钩子接受并忽略
    beginHistoryBatch: () => true,
    endHistoryBatch: () => true,
    editText: (op: any) => {
      const r = sessionTxn([
        {
          op: 'setText',
          target: { slide: op.slideIndex, el: op.sourceId },
          paragraphs: op.paragraphs,
          ...(op.groupId ? { group: op.groupId } : {})
        }
      ])
      return r ? rebuildOne(op.slideIndex) : null
    },
    editTransform: (op: any) => {
      const payload = transformPayload(op.slideIndex, op.fitWidthPx ?? fitWidth, op)
      if (!payload) return null
      const r = sessionTxn([payload])
      return r ? rebuildOne(op.slideIndex) : null
    },
    editTransformMulti: (op: any) => {
      const ops: any[] = []
      for (const item of op.items ?? []) {
        const payload = transformPayload(op.slideIndex, op.fitWidthPx ?? fitWidth, item)
        if (!payload) return null
        ops.push(payload)
      }
      if (!ops.length) return null
      const r = sessionTxn(ops)
      return r ? rebuildOne(op.slideIndex) : null
    },
    batchEditTransform: (op: any) => {
      if (!openedRef) return null
      const baseWidthPx = openedRef.deck.size.cx / EMU_PER_PX_96
      const scale = (op.fitWidthPx ?? fitWidth) / baseWidthPx
      const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
      const r = sessionTxn(
        (op.items ?? []).map((item: any) => ({
          op: 'setTransform',
          target: { slide: op.slideIndex, el: item.sourceId },
          box: { x: toEmu(item.xPx), y: toEmu(item.yPx), cx: toEmu(item.wPx), cy: toEmu(item.hPx) },
          rotDeg: item.rotationDeg,
          resizeTableGrid: true
        }))
      )
      return r ? rebuildOne(op.slideIndex) : null
    },
    addElement: (op: any) => {
      if (!openedRef) return null
      const baseWidthPx = openedRef.deck.size.cx / EMU_PER_PX_96
      const scale = (op.fitWidthPx ?? fitWidth) / baseWidthPx
      const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
      const paragraphs = op.paragraphs?.length
        ? op.paragraphs
        : op.text
          ? String(op.text).split('\n').map((line: string) => ({ runs: [{ text: line }] }))
          : undefined
      const r = sessionTxn([
        {
          op: 'addElement',
          target: { slide: op.slideIndex },
          kind: op.kind,
          offset: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
          ...(paragraphs ? { paragraphs } : {}),
          ...(op.fillColor ? { fill: op.fillColor } : {}),
          ...(op.stroke
            ? { stroke: { color: op.stroke.color, widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT) } }
            : {}),
          ...(op.bodyPr ? { bodyPr: op.bodyPr } : {})
        }
      ])
      if (!r) return null
      const rebuilt = rebuildOne(op.slideIndex)
      return rebuilt ? { slide: rebuilt, sourceId: r.records?.[0]?.created?.[0] } : null
    },
    deleteElement: (op: any) => {
      const r = sessionTxn([{ op: 'deleteElement', target: { slide: op.slideIndex, el: op.sourceId } }])
      return r ? rebuildOne(op.slideIndex) : null
    },
    deleteElements: (op: any) => {
      const r = sessionTxn({
        ops: (op.sourceIds ?? []).map((id: string) => ({
          op: 'deleteElement',
          target: { slide: op.slideIndex, el: id }
        })),
        isolation: 'per_op'
      })
      return r ? rebuildOne(op.slideIndex) : null
    },
    addBlankSlide: (op: any) => {
      const r = sessionTxn([
        { op: 'addBlankSlide', target: { slide: op.sourceIndex } },
        ...(op.before ? [{ op: 'moveSlide', target: { slide: op.sourceIndex }, to: op.sourceIndex + 1 }] : [])
      ])
      if (!r) return null
      fitWidth = op.fitWidthPx ?? fitWidth
      return { ...slideUpdate(), index: op.before ? op.sourceIndex : op.sourceIndex + 1 }
    },
    deleteSlide: (op: any) => {
      const r = sessionTxn([{ op: 'deleteSlide', target: { slide: op.slideIndex } }])
      if (!r) return null
      return slideUpdate()
    },
    duplicateSlides: (op: any) => {
      const r = sessionTxn(
        (op.indices ?? []).map((i: number) => ({ op: 'duplicateSlide', target: { slide: i } })),
      )
      if (!r) return null
      return slideUpdate()
    },
    getRenderSlides: () => (openedRef ? buildAllRenderSlides(openedRef, fitWidth) : null),

    // ---- M9c 保存：引擎序列化 → 宿主（协作服务版本留底）→ commitSaved ----
    save: () => saveDeck(),
    saveAs: () => saveDeck(),
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
  // 调试出口（E2E/联调）：deck 概览与首页文本元素定位
  ;(window as any).__shimDebug = {
    opened: () => !!openedRef,
    name: () => currentName,
    slideCount: () => openedRef?.deck.slides.length ?? 0,
    // 首个文本元素的定位（slideIndex/sourceId/当前文本）
    firstText: () => {
      for (let i = 0; i < (openedRef?.deck.slides.length ?? 0); i++) {
        const sl = (openedRef!.deck.slides as any[])[i]
        for (const el of sl.elements ?? []) {
          const anyEl = el as any
          if (anyEl.text?.paragraphs) {
            return {
              slideIndex: i,
              sourceId: anyEl.id,
              text: anyEl.text.paragraphs
                .flatMap((pa: any) => pa.runs.map((r: any) => r.text))
                .join('')
            }
          }
        }
      }
      return null
    }
  }
  console.info('[slides-shim] SlidesApi 浏览器垫片已激活', pendingDoc ? '带文档:' + pendingDoc.name : '空白')

  // ---- 品牌覆盖（eos 嵌入形态 ?embed=1）：Genspark → Dargo（DOM 文本层，源仓不动）----
  if (q.get('embed') === '1') {
    // 空值防护 + try/catch：文本节点可能在遍历中被 React 置空（实测 null nodeValue
    // 会在 observer 里循环抛错拖垮渲染——品牌覆盖绝不能成为页面故障源）
    const rebrand = (root: Node) => {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      while (walker.nextNode()) {
        try {
          const n = walker.currentNode as Text | null
          const v = n?.nodeValue
          if (!v || !v.includes('Genspark')) continue
          if (/^(script|style)$/i.test((n.parentElement as HTMLElement)?.tagName || '')) continue
          n.nodeValue = v.replace(/Genspark/g, 'Dargo')
        } catch { /* 单节点失败继续 */ }
      }
    }
    const startRebrand = () => {
      rebrand(document.body)
      new MutationObserver((muts) => {
        try {
          for (const m of muts) {
            if (m.type === 'childList') for (const n of m.addedNodes) if (n.nodeType === 1) rebrand(n)
            else if (m.type === 'characterData') {
              const t = m.target as Text | null
              const v = t?.nodeValue
              if (v && v.includes('Genspark')) t.nodeValue = v.replace(/Genspark/g, 'Dargo')
            }
          }
        } catch { /* 兜底：绝不抛出 observer */ }
      }).observe(document.body, { childList: true, subtree: true, characterData: true })
      console.info('[shim] 品牌覆盖激活（Genspark → Dargo）')
    }
    if (document.body) startRebrand()
    else document.addEventListener('DOMContentLoaded', startRebrand)
  }
})()
