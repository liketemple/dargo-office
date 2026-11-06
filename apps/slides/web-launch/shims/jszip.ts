/** jszip 浏览器出口：值由经典 <script>（vite 注入，先于本模块）写入 window.JSZip */
const JSZip = (globalThis as unknown as { JSZip?: unknown }).JSZip
if (!JSZip) throw new Error('[jszip-shim] window.JSZip 未就绪（经典 script 注入失败）')
export default JSZip
