/** node:stream/promises 与 node:fs/promises 的浏览器空垫片
 *  pptx-engine 的 zip.ts 等仅在「文件路径形态」使用它们；页面内以字节形态打开不触达，
 *  提供空导出防止惰性 import 抛 externalized 错误 */
export const pipeline = async () => {}
export const finished = async () => {}
export default { pipeline, finished }
