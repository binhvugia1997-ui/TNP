export type BridgeError = { code: string; message: string }
type BridgeEnvelope<T> = { ok: true; data: T } | { ok: false; error: BridgeError }
type NativeApi = Record<string, (...args: unknown[]) => Promise<BridgeEnvelope<unknown>>>

export class BridgeCallError extends Error {
  readonly code: string

  constructor(error: BridgeError) {
    super(error.message)
    this.name = 'BridgeCallError'
    this.code = error.code
  }
}

declare global {
  interface Window {
    pywebview?: { api?: NativeApi }
  }
}

export function hasNativeBridge(): boolean {
  return typeof window !== 'undefined' && Boolean(window.pywebview?.api)
}

/** Call only the Python API registered by the native pywebview host. No browser/mock fallback exists. */
export async function callBridge<T>(method: string, ...args: unknown[]): Promise<T> {
  if (typeof window === 'undefined') throw new Error('Native bridge is available only in the desktop window.')
  let api = window.pywebview?.api
  if (!api) {
    api = await new Promise<NativeApi>((resolve, reject) => {
      const timeout = window.setTimeout(() => {
        window.removeEventListener('pywebviewready', ready)
        reject(new Error('Python bridge is not connected. Start the application with `python -m app.desktop`.'))
      }, 2500)
      const ready = () => {
        window.clearTimeout(timeout)
        const connected = window.pywebview?.api
        if (connected) resolve(connected)
        else reject(new Error('pywebviewready fired without a Python API.'))
      }
      window.addEventListener('pywebviewready', ready, { once: true })
    })
  }
  const action = api[method]
  if (typeof action !== 'function') throw new Error(`Python bridge method is not available: ${method}`)
  const result = (await action(...args)) as BridgeEnvelope<T>
  if (!result || typeof result !== 'object' || !('ok' in result)) {
    throw new Error(`Python bridge returned an invalid response for ${method}.`)
  }
  if (!result.ok) throw new BridgeCallError(result.error)
  return result.data
}
