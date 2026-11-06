//! WebAssembly FFI 层：给 JS 宿主的直接调用口（绕过 stdin/stdout 管道）。
//!
//! 设计：引擎代码零改动——open 的文件读写经 WASI 宿主提供的内存文件系统
//! （@wasmer/wasi + WasmFs preopen '/'），本层只做「字节 → 临时文件 → 行协议」
//! 的桥接。会话常驻实例内存（sessionId 跨调用有效）。
//!
//! 缓冲契约：应答写入引擎自有复用缓冲并返回 (ptr, len)；JS 必须立即拷贝，
//! 下一次调用使缓冲失效。

use crate::protocol::Engine;

const WORKBOOK_DIR: &str = "/__wasm__";

pub struct WasmEngine {
    engine: Engine,
    buf: Vec<u8>,
    scratch: Vec<u8>,
}

fn make_path(name: &str) -> String {
    let safe: String = name
        .chars()
        .map(|c| if c.is_alphanumeric() || c == '.' || c == '-' || c == '_' { c } else { '_' })
        .collect();
    let safe = if safe.is_empty() { "workbook.xlsx".to_string() } else { safe };
    format!("{}/{}", WORKBOOK_DIR, safe)
}

#[unsafe(no_mangle)]
pub extern "C" fn engine_new() -> *mut WasmEngine {
    let _ = std::fs::create_dir_all(WORKBOOK_DIR);
    Box::into_raw(Box::new(WasmEngine {
        engine: Engine::new(),
        buf: Vec::new(),
        scratch: Vec::new(),
    }))
}

#[unsafe(no_mangle)]
pub unsafe extern "C" fn engine_drop(e: *mut WasmEngine) {
    if !e.is_null() {
        drop(Box::from_raw(e));
    }
}

/// 打开工作簿（字节直开，Engine::open_bytes → sessions 内存解析）。
#[unsafe(no_mangle)]
pub unsafe extern "C" fn engine_open_bytes(
    e: *mut WasmEngine,
    bytes_ptr: *const u8,
    bytes_len: usize,
    name_ptr: *const u8,
    name_len: usize,
    out_len: *mut u32,
) -> *const u8 {
    let w = &mut *e;
    let bytes = std::slice::from_raw_parts(bytes_ptr, bytes_len).to_vec();
    let name = if name_len > 0 {
        String::from_utf8_lossy(std::slice::from_raw_parts(name_ptr, name_len)).into_owned()
    } else {
        "workbook.xlsx".to_string()
    };
    w.buf = w.engine.open_bytes(bytes, &name, "zh").into_bytes();
    *out_len = w.buf.len() as u32;
    w.buf.as_ptr()
}

/// 输入暂存分配：JS 把请求字节/文件字节写进返回的指针（下一次调用失效）。
#[unsafe(no_mangle)]
pub unsafe extern "C" fn engine_alloc(e: *mut WasmEngine, len: usize) -> *mut u8 {
    let w = &mut *e;
    w.scratch.resize(len, 0);
    w.scratch.as_mut_ptr()
}

/// 行协议直通（read_range/close/save 等后续命令，camelCase JSON）。
#[unsafe(no_mangle)]
pub unsafe extern "C" fn engine_handle_line(
    e: *mut WasmEngine,
    line_ptr: *const u8,
    line_len: usize,
    out_len: *mut u32,
) -> *const u8 {
    let w = &mut *e;
    let line = String::from_utf8_lossy(std::slice::from_raw_parts(line_ptr, line_len)).into_owned();
    w.buf = w.engine.handle_line(&line).unwrap_or_default().into_bytes();
    *out_len = w.buf.len() as u32;
    w.buf.as_ptr()
}
