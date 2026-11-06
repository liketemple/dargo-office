/** node:zlib 浏览器垫片（pako 3.x：无 default/crc32 导出，crc32 用内联实现） */
import { inflate, inflateRaw, deflate, deflateRaw } from 'pako'
export const inflateSync = (b: Uint8Array) => inflate(b)
export const inflateRawSync = (b: Uint8Array) => inflateRaw(b)
export const deflateSync = (b: Uint8Array) => deflate(b)
export const deflateRawSync = (b: Uint8Array) => deflateRaw(b)
export function crc32(buf: Uint8Array): number {
  let c = ~0
  for (let i = 0; i < buf.length; i++) {
    c ^= buf[i]
    for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1))
  }
  return ~c >>> 0
}
