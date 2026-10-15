export type BridgeError = { code: string; message: string }
type BridgeEnvelope<T> = { ok: true; data: T } | { ok: false; error: BridgeError }
type NativeMethod = (...args: unknown[]) => Promise<BridgeEnvelope<unknown>>
type NativeApi = Record<string, NativeMethod>

export const BRIDGE_WAIT_TIMEOUT_MS = 8_000
const BRIDGE_METHOD_POLL_MS = 50
const BRIDGE_UNAVAILABLE_MESSAGE =
  'Chưa kết nối ứng dụng Python. Hãy khởi chạy bằng python -m app.desktop; ứng dụng sẽ tự kết nối lại.'

declare global {
  interface Window {
    pywebview?: { api?: NativeApi }
  }
}

export class BridgeUnavailableError extends Error {
  readonly code = 'BRIDGE_UNAVAILABLE'

  constructor() {
    super(BRIDGE_UNAVAILABLE_MESSAGE)
    this.name = 'BridgeUnavailableError'
  }
}

export class BridgeMethodMissingError extends Error {
  readonly code = 'BRIDGE_METHOD_MISSING'

  constructor(method: string) {
    super(`Hợp đồng bridge Python không cung cấp phương thức bắt buộc “${method}”. Hãy kiểm tra frontend và backend cùng phiên bản.`)
    this.name = 'BridgeMethodMissingError'
  }
}

export class BridgeProtocolError extends Error {
  readonly code = 'BRIDGE_CONTRACT_ERROR'

  constructor() {
    super('Python bridge trả về phản hồi không đúng hợp đồng.')
    this.name = 'BridgeProtocolError'
  }
}

export class BridgeCallError extends Error {
  readonly code: string

  constructor(error: BridgeError) {
    super(error.message)
    this.name = 'BridgeCallError'
    this.code = error.code
  }
}

export type BridgeCallOptions = {
  timeoutMs?: number
  signal?: AbortSignal
}

let handshakeConfirmed = false

function abortError(): Error {
  const error = new Error('Bridge readiness wait was cancelled.')
  error.name = 'AbortError'
  return error
}

function awaitWithAbort<T>(operation: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return operation
  if (signal.aborted) return Promise.reject(abortError())
  return new Promise<T>((resolve, reject) => {
    const cleanup = () => signal.removeEventListener('abort', onAbort)
    const onAbort = () => {
      cleanup()
      reject(abortError())
    }
    signal.addEventListener('abort', onAbort, { once: true })
    operation.then(
      (value) => { cleanup(); resolve(value) },
      (error: unknown) => { cleanup(); reject(error) },
    )
  })
}

export function isAbortError(error: unknown): boolean {
  return error instanceof Error && error.name === 'AbortError'
}

export function isBridgeConnectionError(
  error: unknown,
): error is BridgeUnavailableError | BridgeMethodMissingError | BridgeProtocolError {
  return error instanceof BridgeUnavailableError
    || error instanceof BridgeMethodMissingError
    || error instanceof BridgeProtocolError
}

function nativeMethod(method: string): NativeMethod | undefined {
  if (typeof window === 'undefined') return undefined
  const api = window.pywebview?.api
  const action = api?.[method]
  return typeof action === 'function' ? action.bind(api) as NativeMethod : undefined
}

/** True only after pywebview has exposed its stable, callable health-check method. */
export function hasNativeBridge(): boolean {
  return Boolean(nativeMethod('ping'))
}

/** Start a fresh handshake after a disconnect or before a reconnect attempt. */
export function resetBridgeHandshake(): void {
  handshakeConfirmed = false
}

/**
 * Wait for the requested pywebview method, not merely the API object. The event handles the normal
 * initialization path; bounded polling also covers the API-object/method-registration race and the case
 * where pywebviewready fired before this caller was created.
 */
export function waitForBridgeMethod(
  method: string,
  timeoutMs = BRIDGE_WAIT_TIMEOUT_MS,
  signal?: AbortSignal,
): Promise<NativeMethod> {
  if (signal?.aborted) return Promise.reject(abortError())
  if (typeof window === 'undefined') return Promise.reject(new BridgeUnavailableError())

  const target = window
  const readyMethod = nativeMethod(method)
  if (readyMethod) return Promise.resolve(readyMethod)

  const boundedTimeout = Number.isFinite(timeoutMs) ? Math.max(0, timeoutMs) : BRIDGE_WAIT_TIMEOUT_MS
  if (boundedTimeout === 0) return Promise.reject(new BridgeUnavailableError())

  return new Promise<NativeMethod>((resolve, reject) => {
    let settled = false
    let timeoutTimer: number | undefined
    let pollingTimer: number | undefined

    const cleanup = () => {
      if (timeoutTimer !== undefined) target.clearTimeout(timeoutTimer)
      if (pollingTimer !== undefined) target.clearTimeout(pollingTimer)
      target.removeEventListener('pywebviewready', onPywebviewReady)
      signal?.removeEventListener('abort', onAbort)
    }

    const finish = (action?: NativeMethod, error?: Error) => {
      if (settled) return
      settled = true
      cleanup()
      if (action) resolve(action)
      else reject(error || new BridgeUnavailableError())
    }

    const checkReady = (): boolean => {
      if (signal?.aborted) {
        finish(undefined, abortError())
        return true
      }
      const action = nativeMethod(method)
      if (!action) return false
      finish(action)
      return true
    }

    const onPywebviewReady = () => {
      // The event can precede exposure of an individual callable. Keep polling if it is still absent.
      checkReady()
    }
    const onAbort = () => finish(undefined, abortError())
    const poll = () => {
      if (checkReady() || settled) return
      pollingTimer = target.setTimeout(poll, Math.min(BRIDGE_METHOD_POLL_MS, boundedTimeout))
    }

    target.addEventListener('pywebviewready', onPywebviewReady)
    signal?.addEventListener('abort', onAbort, { once: true })

    // Recheck after registering listeners so every ordering is covered without an event-only race.
    if (checkReady()) return

    timeoutTimer = target.setTimeout(() => {
      if (!checkReady()) finish(undefined, new BridgeUnavailableError())
    }, boundedTimeout)
    pollingTimer = target.setTimeout(poll, Math.min(BRIDGE_METHOD_POLL_MS, boundedTimeout))
  })
}

/** Shared implementation used by both normal calls and cancellable startup/polling calls. */
export async function callBridgeWithOptions<T>(
  method: string,
  args: readonly unknown[] = [],
  options: BridgeCallOptions = {},
): Promise<T> {
  const { signal } = options
  if (signal?.aborted) throw abortError()
  if (typeof window === 'undefined') throw new BridgeUnavailableError()

  let action = nativeMethod(method)
  if (!action) {
    try {
      action = await waitForBridgeMethod(method, options.timeoutMs ?? BRIDGE_WAIT_TIMEOUT_MS, signal)
    } catch (error) {
      // A successful ping confirms the native bridge is alive. If this specific method still never
      // appears, report a contract mismatch rather than disguising it as a startup timeout.
      if (error instanceof BridgeUnavailableError && handshakeConfirmed && nativeMethod('ping')) {
        throw new BridgeMethodMissingError(method)
      }
      throw error
    }
  }

  if (signal?.aborted) throw abortError()
  let result: BridgeEnvelope<unknown>
  try {
    result = await awaitWithAbort(Promise.resolve(action(...args)), signal)
  } catch (error) {
    if (signal?.aborted || isAbortError(error)) throw abortError()
    handshakeConfirmed = false
    // A rejected pywebview proxy call is a transport failure. Python service failures use the
    // explicit { ok: false, error } envelope and are handled below without exposing stack traces.
    throw new BridgeUnavailableError()
  }

  if (!result || typeof result !== 'object' || typeof result.ok !== 'boolean') {
    throw new BridgeProtocolError()
  }
  if (!result.ok) {
    const error = result.error
    if (!error || typeof error.code !== 'string' || typeof error.message !== 'string') {
      throw new BridgeProtocolError()
    }
    throw new BridgeCallError(error)
  }

  if (method === 'ping') handshakeConfirmed = true
  return result.data as T
}

/** Call only the Python API registered by the native pywebview host. No browser/mock fallback exists. */
export function callBridge<T>(method: string, ...args: unknown[]): Promise<T> {
  return callBridgeWithOptions<T>(method, args)
}
