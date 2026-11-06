/**
 * xlsx-sidecar.wasm · Node 原生 WASI 验证 v2（M8 地基）
 * 技巧：node:wasi 的 stdin/stdout 收 fd 数字 → 用普通文件 fd（指令文件进、应答文件出）
 */
import { WASI } from 'node:wasi'
import fs from 'node:fs'

const WASM =
  '/Users/like/gitsrc/genoffice/apps/sheets/native/xlsx-engine/target/wasm32-wasip1/release/xlsx-sidecar.wasm'
const XLSX = '/Users/like/orca/workspaces/eos-product/dev-lk/public/fv-samples/sample.xlsx'
const CMD_FILE = '/tmp/xlsx-cmds.txt'
const OUT_FILE = '/tmp/xlsx-out.txt'

const cmds = [
  JSON.stringify({ version: 1, requestId: 'r1', command: 'open', path: XLSX, locale: 'zh' })
].join('\n') + '\n'
fs.writeFileSync(CMD_FILE, cmds)
fs.writeFileSync(OUT_FILE, '')

const stdinFd = fs.openSync(CMD_FILE, 'r')
const stdoutFd = fs.openSync(OUT_FILE, 'w')

const wasi = new WASI({
  version: 'preview1',
  args: ['xlsx-sidecar'],
  preopens: { '/': '/' },
  stdin: stdinFd,
  stdout: stdoutFd,
  stderr: 2
})

console.log('>> open', XLSX)
const t0 = Date.now()
try {
  const mod = await WebAssembly.compile(fs.readFileSync(WASM))
  const inst = await WebAssembly.instantiate(mod, wasi.getImportObject())
  wasi.start(inst)
} catch (e) {
  console.log('[start 返回]', String(e).slice(0, 140))
}
fs.closeSync(stdoutFd)
const result = fs.readFileSync(OUT_FILE, 'utf8')
console.log(`<< ${Date.now() - t0}ms, ${result.length}B:`)
console.log(result.slice(0, 600) || '(空)')
