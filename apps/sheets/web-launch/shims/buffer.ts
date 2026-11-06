/** 'buffer' 包浏览器 polyfill：Uint8Array 子类 + 引擎用到的静态面
 *  （from/alloc/concat/isView + utf8 编解码）。自安装到 globalThis——
 *  模块求值顺序无关（pptx-engine 可能在垫片主体之前求值）。 */
class BufferPolyfill extends Uint8Array {
  static from(input: any, encoding?: string): BufferPolyfill {
    if (typeof input === 'string') {
      const bytes = new TextEncoder().encode(encoding === 'base64' ? atob(input) : input)
      return new BufferPolyfill(bytes)
    }
    if (ArrayBuffer.isView(input) || input instanceof ArrayBuffer) return new BufferPolyfill(input as any)
    if (Array.isArray(input)) return new BufferPolyfill(new Uint8Array(input))
    throw new TypeError('Buffer.from: 不支持的输入')
  }
  static alloc(size: number, fill = 0): BufferPolyfill {
    const b = new BufferPolyfill(size)
    if (fill) b.fill(fill)
    return b
  }
  static concat(list: Uint8Array[]): BufferPolyfill {
    const total = list.reduce((n, x) => n + x.length, 0)
    const out = new BufferPolyfill(total)
    let off = 0
    for (const x of list) { out.set(x, off); off += x.length }
    return out
  }
  static isView(x: unknown): boolean { return ArrayBuffer.isView(x) }
  static isBuffer(x: unknown): boolean { return x instanceof BufferPolyfill }
  toString(encoding?: string): string {
    if (encoding === 'base64') {
      let bin = ''
      const CH = 0x8000
      for (let i = 0; i < this.length; i += CH)
        bin += String.fromCharCode.apply(null, Array.from(this.subarray(i, i + CH)))
      return btoa(bin)
    }
    return new TextDecoder().decode(this)
  }
}
const g = globalThis as any
if (!g.Buffer) g.Buffer = BufferPolyfill
export const Buffer: any = g.Buffer
export default { Buffer }
