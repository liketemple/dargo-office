const p: any = {
  env: { NODE_ENV: 'production' }, version: 'v20.0.0', platform: 'browser',
  nextTick: (fn: (...a: unknown[]) => void, ...args: unknown[]) =>
    Promise.resolve().then(() => fn(...args)),
  browser: true
}
export default p
export const env = p.env
export const nextTick = p.nextTick
