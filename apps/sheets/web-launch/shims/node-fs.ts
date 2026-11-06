/** node:fs 浏览器 shim：xlsx-gateway 顶层命名导入需要绑定存在；
 *  这些函数只在其 Electron 文件写路径（writeXlsxAtomically 等）被调用，
 *  浏览器保存链走 JSZip 内存路径，永不触达——真被调到即抛错暴露误用 */
const untouchable = (name: string) => () => {
  throw new Error(`web-shim: node:fs.${name} 在浏览器不可用`)
}
export const closeSync = untouchable('closeSync')
export const fsyncSync = untouchable('fsyncSync')
export const openSync = untouchable('openSync')
export const readFileSync = untouchable('readFileSync')
export const writeFileSync = untouchable('writeFileSync')
export const existsSync = untouchable('existsSync')
export const mkdirSync = untouchable('mkdirSync')
export default { closeSync, fsyncSync, openSync, readFileSync, writeFileSync, existsSync, mkdirSync }
