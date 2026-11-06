/** node:fs/promises 浏览器 shim（同 node-fs：仅满足命名绑定，调用即抛） */
const untouchable = (name: string) => () => {
  throw new Error(`web-shim: node:fs/promises.${name} 在浏览器不可用`)
}
export const rename = untouchable('rename')
export const rm = untouchable('rm')
export const mkdtemp = untouchable('mkdtemp')
export default { rename, rm, mkdtemp }
