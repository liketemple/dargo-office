/** jszip 浏览器出口：经典 script 注入的 window.JSZip */
const JSZip = (globalThis as unknown as { JSZip?: unknown }).JSZip
if (!JSZip) throw new Error('[jszip-shim] window.JSZip 未就绪')
export default JSZip
