/** 'buffer' 包浏览器 shim：兜回垫片已注入的 globalThis.Buffer（sheets-shim 顶部 polyfill） */
export const Buffer: typeof globalThis.Buffer = (globalThis as any).Buffer
export default { Buffer }
