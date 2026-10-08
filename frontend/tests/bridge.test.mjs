import assert from 'node:assert/strict'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { after, test } from 'node:test'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import ts from 'typescript'

const sourceDirectory = fileURLToPath(new URL('../src/lib/', import.meta.url))
const compiledDirectory = await mkdtemp(join(tmpdir(), 'tnp-bridge-tests-'))
const compile = async (sourceName, outputName) => {
  const source = await readFile(join(sourceDirectory, sourceName), 'utf8')
  const { outputText } = ts.transpileModule(source, {
    compilerOptions: { module: ts.ModuleKind.ES2022, target: ts.ScriptTarget.ES2022 },
  })
  const javascript = outputText.replace(/from (["'])\.\/bridge\1/g, "from './bridge.mjs'")
  await writeFile(join(compiledDirectory, outputName), javascript)
}
await compile('bridge.ts', 'bridge.mjs')
await compile('startup.ts', 'startup.mjs')
const bridgeModules = {
  ...(await import(pathToFileURL(join(compiledDirectory, 'bridge.mjs')).href)),
  ...(await import(pathToFileURL(join(compiledDirectory, 'startup.mjs')).href)),
}
after(() => rm(compiledDirectory, { recursive: true, force: true }))
const delay = (ms) => new Promise((resolve) => globalThis.setTimeout(resolve, ms))

class FakeWindow extends EventTarget {
  timeouts = new Set()
  intervals = new Set()
  readyListeners = new Set()

  addEventListener(type, listener, options) {
    super.addEventListener(type, listener, options)
    if (type === 'pywebviewready' && listener) this.readyListeners.add(listener)
  }

  removeEventListener(type, listener, options) {
    super.removeEventListener(type, listener, options)
    if (type === 'pywebviewready') this.readyListeners.delete(listener)
  }

  setTimeout(callback, ms = 0, ...args) {
    let timer
    timer = globalThis.setTimeout(() => {
      this.timeouts.delete(timer)
      callback(...args)
    }, ms)
    this.timeouts.add(timer)
    return timer
  }

  clearTimeout(timer) {
    globalThis.clearTimeout(timer)
    this.timeouts.delete(timer)
  }

  setInterval(callback, ms = 0, ...args) {
    const timer = globalThis.setInterval(callback, ms, ...args)
    this.intervals.add(timer)
    return timer
  }

  clearInterval(timer) {
    globalThis.clearInterval(timer)
    this.intervals.delete(timer)
  }
}

async function withWindow(run) {
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'window')
  const fakeWindow = new FakeWindow()
  Object.defineProperty(globalThis, 'window', { configurable: true, writable: true, value: fakeWindow })
  bridgeModules.resetBridgeHandshake()
  try {
    return await run(fakeWindow)
  } finally {
    for (const timer of fakeWindow.timeouts) fakeWindow.clearTimeout(timer)
    for (const timer of fakeWindow.intervals) fakeWindow.clearInterval(timer)
    if (previous) Object.defineProperty(globalThis, 'window', previous)
    else delete globalThis.window
  }
}

async function waitUntil(predicate, timeoutMs = 600) {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    if (predicate()) return
    await delay(5)
  }
  assert.fail('condition did not become true before timeout')
}

test('API absent, pywebviewready fires, then requested method appears and call succeeds', async () => {
  await withWindow(async (fakeWindow) => {
    const request = bridgeModules.callBridge('greet', 'TNP')
    await delay(15)
    fakeWindow.pywebview = { api: {} }
    fakeWindow.dispatchEvent(new Event('pywebviewready'))
    await delay(20)
    fakeWindow.pywebview.api.greet = async (name) => ({ ok: true, data: `hello ${name}` })

    assert.equal(await request, 'hello TNP')
    assert.equal(fakeWindow.readyListeners.size, 0)
    assert.equal(fakeWindow.timeouts.size, 0)
  })
})

test('existing API object is not ready until the requested method appears', async () => {
  await withWindow(async (fakeWindow) => {
    fakeWindow.pywebview = { api: {} }
    const request = bridgeModules.callBridge('late_method', 7)
    globalThis.setTimeout(() => {
      fakeWindow.pywebview.api.late_method = async (value) => ({ ok: true, data: value * 2 })
    }, 35)

    assert.equal(await request, 14)
    assert.equal(fakeWindow.readyListeners.size, 0)
    assert.equal(fakeWindow.timeouts.size, 0)
  })
})

test('method is immediately available even when pywebviewready fired before the waiter', async () => {
  await withWindow(async (fakeWindow) => {
    fakeWindow.pywebview = { api: { ready_method: async () => ({ ok: true, data: 'ready' }) } }
    fakeWindow.dispatchEvent(new Event('pywebviewready'))

    const method = await bridgeModules.waitForBridgeMethod('ready_method', 200)
    assert.deepEqual(await method(), { ok: true, data: 'ready' })
    assert.equal(fakeWindow.readyListeners.size, 0)
    assert.equal(fakeWindow.timeouts.size, 0)
  })
})

test('readiness timeout is bounded, classified as BRIDGE_UNAVAILABLE, and cleans timers/listeners', async () => {
  await withWindow(async (fakeWindow) => {
    const started = Date.now()
    await assert.rejects(
      bridgeModules.waitForBridgeMethod('never_ready', 70),
      (error) => error.code === 'BRIDGE_UNAVAILABLE',
    )
    assert.ok(Date.now() - started < 500, 'timeout should remain bounded')
    assert.equal(fakeWindow.readyListeners.size, 0)
    assert.equal(fakeWindow.timeouts.size, 0)
  })
})

test('aborting a readiness wait removes its event listener and both timers', async () => {
  await withWindow(async (fakeWindow) => {
    const controller = new AbortController()
    const request = bridgeModules.waitForBridgeMethod('cancelled_wait', 500, controller.signal)
    assert.equal(fakeWindow.readyListeners.size, 1)
    assert.equal(fakeWindow.timeouts.size, 2)
    controller.abort()

    await assert.rejects(request, (error) => error.name === 'AbortError')
    assert.equal(fakeWindow.readyListeners.size, 0)
    assert.equal(fakeWindow.timeouts.size, 0)
  })
})

test('a confirmed live bridge with a truly missing method reports BRIDGE_METHOD_MISSING', async () => {
  await withWindow(async (fakeWindow) => {
    fakeWindow.pywebview = {
      api: { ping: async () => ({ ok: true, data: 'Python bridge OK' }) },
    }
    assert.equal(await bridgeModules.callBridgeWithOptions('ping', [], { timeoutMs: 80 }), 'Python bridge OK')
    await assert.rejects(
      bridgeModules.callBridgeWithOptions('get_dashboard_state', [], { timeoutMs: 70 }),
      (error) => error.code === 'BRIDGE_METHOD_MISSING',
    )
    assert.equal(fakeWindow.readyListeners.size, 0)
    assert.equal(fakeWindow.timeouts.size, 0)
  })
})

test('transport rejection is BRIDGE_UNAVAILABLE and Python error envelopes retain their own code', async () => {
  await withWindow(async (fakeWindow) => {
    fakeWindow.pywebview = {
      api: {
        disconnected: async () => { throw new Error('D:\\private\\internal-stack') },
        python_failure: async () => ({
          ok: false,
          error: { code: 'INTERNAL_ERROR', message: 'Đã xảy ra lỗi nội bộ.' },
        }),
      },
    }
    await assert.rejects(
      bridgeModules.callBridge('disconnected'),
      (error) => error.code === 'BRIDGE_UNAVAILABLE' && !error.message.includes('internal-stack'),
    )
    await assert.rejects(
      bridgeModules.callBridge('python_failure'),
      (error) => error.code === 'INTERNAL_ERROR' && error.message === 'Đã xảy ra lỗi nội bộ.',
    )
  })
})

test('hasNativeBridge requires the stable callable ping method', async () => {
  await withWindow(async (fakeWindow) => {
    fakeWindow.pywebview = { api: {} }
    assert.equal(bridgeModules.hasNativeBridge(), false)
    fakeWindow.pywebview.api.ping = async () => ({ ok: true, data: 'ok' })
    assert.equal(bridgeModules.hasNativeBridge(), true)
  })
})

test('polling is gated until startup handshake succeeds', async () => {
  await withWindow(async (fakeWindow) => {
    let resolveConnect
    let pollCount = 0
    let connected = false
    const firstConnect = new Promise((resolve) => { resolveConnect = resolve })
    const controller = bridgeModules.createBridgeSessionController({
      connect: () => firstConnect,
      poll: async () => { pollCount += 1 },
      onConnected: () => { connected = true },
      onDisconnected: () => {},
      pollIntervalMs: 12,
      reconnectDelaysMs: [20],
    })

    controller.start()
    await delay(35)
    assert.equal(pollCount, 0)
    assert.equal(connected, false)
    resolveConnect({ ok: true })
    await waitUntil(() => connected)
    await waitUntil(() => pollCount >= 2)
    assert.equal(fakeWindow.intervals.size, 1)

    controller.stop()
    await delay(15)
    assert.equal(fakeWindow.intervals.size, 0)
    assert.equal(fakeWindow.timeouts.size, 0)
  })
})

test('disconnected startup retries with bounded backoff and never starts dashboard polling', async () => {
  await withWindow(async (fakeWindow) => {
    let connectCount = 0
    let pollCount = 0
    const controller = bridgeModules.createBridgeSessionController({
      connect: async () => {
        connectCount += 1
        throw new bridgeModules.BridgeUnavailableError()
      },
      poll: async () => { pollCount += 1 },
      onConnected: () => {},
      onDisconnected: () => {},
      pollIntervalMs: 8,
      reconnectDelaysMs: [20, 35],
    })

    controller.start()
    await delay(105)
    assert.ok(connectCount >= 2 && connectCount <= 5, `unexpected reconnect attempts: ${connectCount}`)
    assert.equal(pollCount, 0)
    assert.equal(fakeWindow.intervals.size, 0)
    controller.stop()
    assert.equal(fakeWindow.timeouts.size, 0)
  })
})

test('bridge-unavailable startup retries do not call get_dashboard_state', async () => {
  await withWindow(async (fakeWindow) => {
    let dashboardCalls = 0
    fakeWindow.pywebview = {
      api: {
        get_dashboard_state: async () => {
          dashboardCalls += 1
          return { ok: true, data: {} }
        },
      },
    }
    const controller = bridgeModules.createBridgeSessionController({
      connect: (signal) => bridgeModules.startAppBridge(signal, 30),
      poll: async () => { dashboardCalls += 1 },
      onConnected: () => {},
      onDisconnected: () => {},
      pollIntervalMs: 8,
      reconnectDelaysMs: [20],
    })

    controller.start()
    await delay(100)
    assert.equal(dashboardCalls, 0)
    assert.equal(fakeWindow.intervals.size, 0)
    controller.stop()
    await delay(10)
    assert.equal(fakeWindow.readyListeners.size, 0)
    assert.equal(fakeWindow.timeouts.size, 0)
  })
})

test('poll failure pauses the interval, reconnects once, and resumes without timer leaks', async () => {
  await withWindow(async (fakeWindow) => {
    let connectCount = 0
    let connectedCount = 0
    let disconnectCount = 0
    let pollCount = 0
    const controller = bridgeModules.createBridgeSessionController({
      connect: async () => ({ session: ++connectCount }),
      poll: async () => {
        pollCount += 1
        if (pollCount === 1) throw new bridgeModules.BridgeUnavailableError()
      },
      onConnected: () => { connectedCount += 1 },
      onDisconnected: () => { disconnectCount += 1 },
      pollIntervalMs: 10,
      reconnectDelaysMs: [20, 35],
    })

    controller.start()
    await waitUntil(() => connectedCount >= 2 && pollCount >= 2)
    assert.equal(connectCount, 2)
    assert.equal(disconnectCount, 1)
    assert.equal(controller.ready, true)
    assert.equal(fakeWindow.intervals.size, 1)
    const before = pollCount
    await delay(35)
    assert.ok(pollCount > before, 'polling should resume after reconnect')
    assert.equal(fakeWindow.intervals.size, 1, 'there must be only one active dashboard interval')

    controller.stop()
    await delay(20)
    const stoppedAt = pollCount
    const attemptsAtStop = connectCount
    await delay(40)
    assert.equal(pollCount, stoppedAt)
    assert.equal(connectCount, attemptsAtStop)
    assert.equal(fakeWindow.intervals.size, 0)
    assert.equal(fakeWindow.timeouts.size, 0)
    assert.equal(fakeWindow.readyListeners.size, 0)
  })
})

test('failed startup cancels sibling method waits and cleans their timers/listeners', async () => {
  await withWindow(async (fakeWindow) => {
    fakeWindow.pywebview = {
      api: {
        ping: async () => ({ ok: true, data: 'Python bridge OK' }),
        get_app_version: async () => ({
          ok: false,
          error: { code: 'STARTUP_FAILED', message: 'Không thể đọc phiên bản.' },
        }),
      },
    }

    await assert.rejects(bridgeModules.startAppBridge(undefined, 500), (error) => error.code === 'STARTUP_FAILED')
    assert.equal(fakeWindow.readyListeners.size, 0)
    assert.equal(fakeWindow.timeouts.size, 0)
  })
})

test('normal ready path pings, loads version/config/dashboard, and adds no startup delay', async () => {
  await withWindow(async (fakeWindow) => {
    const calls = []
    fakeWindow.pywebview = {
      api: {
        ping: async () => { calls.push('ping'); return { ok: true, data: 'Python bridge OK' } },
        get_app_version: async () => { calls.push('get_app_version'); return { ok: true, data: { version: '1.3.4', build: '017' } } },
        get_current_config: async () => { calls.push('get_current_config'); return { ok: true, data: { config: true } } },
        get_dashboard_state: async () => { calls.push('get_dashboard_state'); return { ok: true, data: { dashboard: true } } },
      },
    }
    const started = Date.now()
    const initial = await bridgeModules.startAppBridge(undefined, 250)

    assert.deepEqual(calls, ['ping', 'get_app_version', 'get_current_config', 'get_dashboard_state'])
    assert.deepEqual(initial, {
      version: { version: '1.3.4', build: '017' },
      config: { config: true },
      dashboard: { dashboard: true },
    })
    assert.ok(Date.now() - started < 200, 'a callable ready bridge should not wait for the timeout or poll interval')
    assert.equal(fakeWindow.timeouts.size, 0)
  })
})
