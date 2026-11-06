/** node:fs/promises 浏览器空垫片（同 node-stream 理由） */
export const open = async () => { throw new Error('web-shim: fs 不可用（浏览器形态以字节直开）') }
export const readFile = async () => { throw new Error('web-shim: fs 不可用') }
export const writeFile = async () => { throw new Error('web-shim: fs 不可用') }
export const stat = async () => { throw new Error('web-shim: fs 不可用') }
