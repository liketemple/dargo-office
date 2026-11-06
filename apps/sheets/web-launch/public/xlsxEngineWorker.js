/**
 * xlsx 引擎 Worker 宿主（M8）：Rust WASM FFI 直调 + 手写最小 WASI 导入垫片。
 *
 * 架构升级（放弃 @wasmer/wasi）：引擎的 open 已改为内存字节直开
 * （sessions.open_bytes_with_locale，ZipArchive<Cursor<Vec<u8>>> 全链路），
 * 不再需要真实文件系统——WASI 导入只需 22 个符号的空/桩实现。
 *
 * 消息协议（主线程 ↔ worker）：
 *   → {type:'init'}                      装配 wasm
 *   → {type:'open', id, bytes, name}     打开工作簿
 *   → {type:'cmd', id, line}             行协议（read_range/close 等）
 *   ← {type:'ready'} | {type:'resp', id, json} | {type:'error', id, message}
 */

let ex = null
let engine = 0

const enc = new TextEncoder()
const dec = new TextDecoder()

/** 最小 WASI preview1 导入：所有 fd/path 调用回 errno（BADF=8/NOTSUP=52），
 *  clock/random 提供真实值（引擎解析时间戳/UUID 用），stdout 写丢弃 */
function wasiImports() {
  const mem = () => new DataView(ex.memory.buffer)
  const errno = (n) => n >>> 0
  const imp = {
    clock_time_get: (clockId, precision, out) => {
      mem().setBigUint64(out, BigInt(Date.now()) * 1000000n, true)
      return errno(0)
    },
    random_get: (buf, len) => {
      const u8 = new Uint8Array(ex.memory.buffer, buf, len)
      crypto.getRandomValues(u8)
      return errno(0)
    },
    environ_sizes_get: (count, bufSize) => {
      mem().setUint32(count, 0, true)
      mem().setUint32(bufSize, 0, true)
      return errno(0)
    },
    environ_get: () => errno(0),
    fd_write: (fd, iovs, iovsLen, nwritten) => {
      // stdout/stderr 丢弃（FFI 模式不用管道）
      mem().setUint32(nwritten, 0, true)
      return errno(0)
    },
    fd_read: (fd, iovs, iovsLen, nread) => {
      mem().setUint32(nread, 0, true)
      return errno(8) // BADF：无 stdin
    },
    fd_close: () => errno(0),
    fd_fdstat_get: () => errno(52),
    fd_filestat_get: () => errno(52),
    fd_prestat_get: () => errno(8),
    fd_prestat_dir_name: () => errno(8),
    fd_readdir: () => errno(52),
    fd_seek: () => errno(52),
    fd_sync: () => errno(0),
    path_create_directory: () => errno(52),
    path_filestat_get: () => errno(52),
    path_open: () => errno(52),
    path_remove_directory: () => errno(52),
    path_unlink_file: () => errno(52),
    poll_oneoff: () => errno(52),
    proc_exit: () => {}
  }
  return { wasi_snapshot_preview1: imp }
}

async function init() {
  const res = await fetch('./xlsx/xlsx-sidecar.wasm')
  const bytes = await res.arrayBuffer()
  const mod = await WebAssembly.compile(bytes)
  const inst = await WebAssembly.instantiate(mod, wasiImports())
  ex = inst.exports
  engine = ex.engine_new()
  if (!engine) throw new Error('engine_new 失败')
  self.postMessage({ type: 'ready' })
}

function callFFI(invoke, writeChunks) {
  const total = writeChunks.reduce((n, b) => n + b.length, 0)
  const ptr = ex.engine_alloc(engine, total + 8)
  new Uint8Array(ex.memory.buffer).set(writeChunks[0], ptr)
  let off = ptr + writeChunks[0].length
  for (const b of writeChunks.slice(1)) {
    new Uint8Array(ex.memory.buffer).set(b, off)
    off += b.length
  }
  const outLenOff = (off + 3) & ~3
  new DataView(ex.memory.buffer).setUint32(outLenOff, 0, true)
  const retPtr = invoke(ptr, off, outLenOff)
  const len = new DataView(ex.memory.buffer).getUint32(outLenOff, true)
  if (!len) return ''
  return dec.decode(new Uint8Array(ex.memory.buffer).subarray(retPtr, retPtr + len))
}

self.onmessage = async (ev) => {
  const d = ev.data
  try {
    if (d.type === 'init') {
      await init()
      return
    }
    if (!ex) throw new Error('worker 未初始化')
    if (d.type === 'open') {
      const bytes = new Uint8Array(d.bytes)
      const name = enc.encode(d.name || 'workbook.xlsx')
      const json = callFFI(
        (base, _end, outOff) =>
          ex.engine_open_bytes(
            engine,
            base,
            bytes.length,
            base + bytes.length,
            name.length,
            outOff
          ),
        [bytes, name]
      )
      self.postMessage({ type: 'resp', id: d.id, json })
      return
    }
    if (d.type === 'cmd') {
      const line = enc.encode(d.line)
      const json = callFFI(
        (base, _end, outOff) => ex.engine_handle_line(engine, base, line.length, outOff),
        [line]
      )
      self.postMessage({ type: 'resp', id: d.id, json })
      return
    }
  } catch (e) {
    self.postMessage({ type: 'error', id: d?.id, message: String(e) })
  }
}
