import {
  BRIDGE_WAIT_TIMEOUT_MS,
  isBridgeConnectionError,
  resetBridgeHandshake,
  callBridgeWithOptions,
} from './bridge'

export type BridgeBootstrap<Version, Config, Dashboard> = {
  version: Version
  config: Config
  dashboard: Dashboard
}

/** One native handshake gates all initial state reads and the dashboard polling session. */
export async function startAppBridge<Version, Config, Dashboard>(
  signal?: AbortSignal,
  timeoutMs = BRIDGE_WAIT_TIMEOUT_MS,
): Promise<BridgeBootstrap<Version, Config, Dashboard>> {
  resetBridgeHandshake()
  const startupAbort = new AbortController()
  const relayAbort = () => startupAbort.abort()
  if (signal?.aborted) startupAbort.abort()
  else signal?.addEventListener('abort', relayAbort, { once: true })

  const request = <T,>(method: string) =>
    callBridgeWithOptions<T>(method, [], { signal: startupAbort.signal, timeoutMs })

  try {
    await request<unknown>('ping')
    const [version, config, dashboard] = await Promise.all([
      request<Version>('get_app_version'),
      request<Config>('get_current_config'),
      request<Dashboard>('get_dashboard_state'),
    ])
    return { version, config, dashboard }
  } finally {
    // If one contract call fails, cancel sibling readiness waits instead of leaving their timers/listeners alive.
    startupAbort.abort()
    signal?.removeEventListener('abort', relayAbort)
  }
}

export type BridgeSessionOptions<Initial> = {
  connect: (signal: AbortSignal) => Promise<Initial>
  poll: (signal: AbortSignal) => Promise<void>
  onConnected: (initial: Initial) => void
  onDisconnected: (error: unknown) => void
  onPollError?: (error: unknown) => void
  pollIntervalMs?: number
  reconnectDelaysMs?: readonly number[]
}

export type BridgeSessionController = {
  start: () => void
  stop: () => void
  invalidate: (error: unknown) => void
  readonly ready: boolean
}

const DEFAULT_POLL_INTERVAL_MS = 750
const DEFAULT_RECONNECT_DELAYS_MS = [1_500, 3_000, 5_000, 8_000] as const

/**
 * Owns exactly one post-handshake polling interval and one bounded reconnect timer. Failed startup attempts
 * never poll; disconnects stop polling immediately and retry with a capped backoff.
 */
export function createBridgeSessionController<Initial>(
  options: BridgeSessionOptions<Initial>,
): BridgeSessionController {
  const pollIntervalMs = Math.max(1, options.pollIntervalMs ?? DEFAULT_POLL_INTERVAL_MS)
  const reconnectDelays = options.reconnectDelaysMs?.length
    ? options.reconnectDelaysMs
    : DEFAULT_RECONNECT_DELAYS_MS

  let disposed = false
  let started = false
  let ready = false
  let connecting = false
  let pollInFlight = false
  let generation = 0
  let reconnectAttempt = 0
  let pollTimer: number | undefined
  let reconnectTimer: number | undefined
  let connectAbort: AbortController | undefined
  let pollAbort: AbortController | undefined

  const clearReconnectTimer = () => {
    if (reconnectTimer === undefined) return
    window.clearTimeout(reconnectTimer)
    reconnectTimer = undefined
  }

  const stopPolling = () => {
    if (pollTimer !== undefined) {
      window.clearInterval(pollTimer)
      pollTimer = undefined
    }
    if (pollAbort) {
      pollAbort.abort()
      pollAbort = undefined
    }
  }

  const scheduleReconnect = () => {
    if (disposed || reconnectTimer !== undefined) return
    const index = Math.min(reconnectAttempt, reconnectDelays.length - 1)
    const delayMs = reconnectDelays[index]
    reconnectAttempt = Math.min(reconnectAttempt + 1, reconnectDelays.length - 1)
    reconnectTimer = window.setTimeout(() => {
      reconnectTimer = undefined
      connect()
    }, delayMs)
  }

  const onDisconnected = (error: unknown) => {
    if (disposed) return
    const wasAlreadyWaiting = !ready && reconnectTimer !== undefined
    ready = false
    stopPolling()
    if (!wasAlreadyWaiting) options.onDisconnected(error)
    scheduleReconnect()
  }

  const pollOnce = () => {
    if (disposed || !ready || pollInFlight) return
    const controller = new AbortController()
    pollAbort = controller
    pollInFlight = true
    void options.poll(controller.signal).catch((error: unknown) => {
      if (disposed || controller.signal.aborted) return
      if (isBridgeConnectionError(error)) onDisconnected(error)
      else {
        try {
          options.onPollError?.(error)
        } catch {
          // A UI notification callback must not create an unhandled polling rejection.
        }
      }
    }).finally(() => {
      pollInFlight = false
      if (pollAbort === controller) pollAbort = undefined
    })
  }

  const startPolling = () => {
    if (disposed || !ready || pollTimer !== undefined) return
    pollTimer = window.setInterval(pollOnce, pollIntervalMs)
  }

  function connect() {
    if (disposed || connecting) return
    clearReconnectTimer()
    connecting = true
    const attemptGeneration = ++generation
    const controller = new AbortController()
    connectAbort = controller

    void (async () => {
      try {
        const initial = await options.connect(controller.signal)
        if (disposed || controller.signal.aborted || attemptGeneration !== generation) return
        ready = true
        reconnectAttempt = 0
        options.onConnected(initial)
        startPolling()
      } catch (error) {
        if (disposed || controller.signal.aborted || attemptGeneration !== generation) return
        onDisconnected(error)
      } finally {
        if (attemptGeneration === generation) {
          connecting = false
          if (connectAbort === controller) connectAbort = undefined
        }
      }
    })()
  }

  const invalidate = (error: unknown) => {
    if (disposed) return
    const alreadyRetrying = !ready && reconnectTimer !== undefined
    if (alreadyRetrying) return
    if (connecting) {
      generation += 1
      connectAbort?.abort()
      connectAbort = undefined
      connecting = false
    }
    onDisconnected(error)
  }

  return {
    start: () => {
      if (disposed || started) return
      started = true
      connect()
    },
    stop: () => {
      if (disposed) return
      disposed = true
      ready = false
      generation += 1
      clearReconnectTimer()
      stopPolling()
      connectAbort?.abort()
      connectAbort = undefined
      connecting = false
    },
    invalidate,
    get ready() {
      return ready && !disposed
    },
  }
}
