/** node:crypto 浏览器垫片（与 eos-product src/lib/genoffice-shims 同款） */
export const randomUUID = (): string =>
  (globalThis.crypto?.randomUUID?.() ??
    'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      const r = (Math.random() * 16) | 0
      return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16)
    })) as string
export function createHash(_algo?: string) {
  let acc = 0x811c9dc5
  return {
    update(data: Uint8Array | string) {
      const bytes = typeof data === 'string' ? new TextEncoder().encode(data) : data
      for (let i = 0; i < bytes.length; i++) {
        acc ^= bytes[i]
        acc = Math.imul(acc, 0x01000193) >>> 0
      }
      return this
    },
    digest() {
      const out = new Uint8Array(4)
      new DataView(out.buffer).setUint32(0, acc)
      return Array.from(out, (b) => b.toString(16).padStart(2, '0')).join('')
    }
  }
}
