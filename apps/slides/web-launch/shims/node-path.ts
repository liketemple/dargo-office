/** node:path 浏览器 shim：简单字符串实现（保存链只用 join/dirname/basename 拼路径名，
 *  不触盘） */
export const join = (...parts: string[]): string => parts.filter(Boolean).join('/').replace(/\/+/g, '/')
export const dirname = (p: string): string => (p.includes('/') ? p.slice(0, p.lastIndexOf('/')) : '.')
export const basename = (p: string): string => (p.includes('/') ? p.slice(p.lastIndexOf('/') + 1) : p)
export default { join, dirname, basename }
