import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import * as React from 'react'
import { fileURLToPath } from 'node:url'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
let server
let shouldAcceptScanSnapshot
let StoreProvider
let useStore

before(async () => {
  server = await createServer({
    configFile: `${frontendRoot}/vite.config.ts`,
    root: frontendRoot,
    appType: 'custom',
    server: { middlewareMode: true, hmr: false },
  })
  ;({ shouldAcceptScanSnapshot, StoreProvider, useStore } = await server.ssrLoadModule('/src/state/store.tsx'))
})

after(async () => {
  await server?.close()
})

test('scan snapshot generations reject delayed rows and accept monotonic leaf revisions', () => {
  const current = { generation: 7, revision: 2 }
  assert.equal(shouldAcceptScanSnapshot(current, {
    scanned: true, message: 'old generation', generation: 6, revision: 99, inProgress: true,
  }), false)
  assert.equal(shouldAcceptScanSnapshot(current, {
    scanned: true, message: 'old revision', generation: 7, revision: 1, inProgress: true,
  }), false)
  assert.equal(shouldAcceptScanSnapshot(current, {
    scanned: true, message: 'next leaf', generation: 7, revision: 3, inProgress: true,
  }), true)
  assert.equal(shouldAcceptScanSnapshot(current, {
    scanned: true, message: 'next scan', generation: 8, revision: 0, inProgress: true,
  }), true)
})

test('a pending Quét ignores the previous final dashboard but accepts incremental and final new generation', () => {
  const current = { generation: 11, revision: 4 }
  const pendingBaseline = 11
  assert.equal(shouldAcceptScanSnapshot(current, {
    scanned: true, message: 'delayed prior final', generation: 11, revision: 4, inProgress: false,
  }, pendingBaseline), false)
  assert.equal(shouldAcceptScanSnapshot(current, {
    scanned: true, message: 'first completed leaf', generation: 12, revision: 1, inProgress: true,
  }, pendingBaseline), true)
  assert.equal(shouldAcceptScanSnapshot({ generation: 12, revision: 2 }, {
    scanned: true, message: 'final aggregate', generation: 12, revision: 3, inProgress: false,
  }, pendingBaseline), true)
})

const delay = (ms) => new Promise((resolve) => globalThis.setTimeout(resolve, ms))
const envelope = (data) => ({ ok: true, data })

function report(id, fileName) {
  return {
    id, stt: Number(id) + 1, managementNumber: `26101010${id}-VOC`, fileName, path: '',
    occurrenceDate: '10/10/2026', vendor: '', manualFields: {}, slides: 0, status: 'waiting',
    shortResult: 'Cần bổ sung dữ liệu', results: {
      causes: [], countermeasures: [], temporaryRemoved: false, afterImages: [],
    },
  }
}

function dashboard(scan, reports = []) {
  const config = {
    paths: { reportFolder: 'reports', template: '', output: '' },
    period: { mode: 'month', month: '10', year: '2026', from: '', to: '' },
    forceReprocess: false,
    ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [] },
    update: { path: '', autoCheck: false },
  }
  return {
    app: { name: 'Report Extractor', title: 'Report Extractor', version: '1.3.4', build: '017', buildNumber: 17 },
    config, scan, reports,
    job: { status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0,
      currentFile: '', elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false },
    logs: [],
    ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [], aiStatus: '',
      discovering: false, discoveryMessage: '', discoveryResults: [], discoveryChecked: 0, discoveryTotal: 0,
      serverApplyRunning: false, serverApplyMessage: '' },
    update: { path: '', autoCheck: false, status: 'idle', message: '', progress: 0, progressStage: '' },
    diagnostics: { running: false, rows: [], error: '' },
  }
}

async function waitFor(predicate, description, timeoutMs = 3_000) {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    if (predicate()) return
    await React.act(async () => { await delay(10) })
  }
  assert.fail(`Timed out waiting for ${description}`)
}

test('the store renders completed leaf snapshots before scan_reports returns and ignores an older revision', async () => {
  const dom = new JSDOM('<!doctype html><html><body><div id="root"></div></body></html>', { url: 'http://localhost/' })
  const globalNames = ['window', 'document', 'navigator', 'HTMLElement', 'Node', 'Event', 'IS_REACT_ACT_ENVIRONMENT']
  const prior = new Map(globalNames.map((name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)]))
  const define = (name, value) => Object.defineProperty(globalThis, name, { configurable: true, writable: true, value })
  define('window', dom.window)
  define('document', dom.window.document)
  define('navigator', dom.window.navigator)
  define('HTMLElement', dom.window.HTMLElement)
  define('Node', dom.window.Node)
  define('Event', dom.window.Event)
  define('IS_REACT_ACT_ENVIRONMENT', true)

  let current = dashboard({ scanned: false, message: 'ready', generation: 0, revision: 0, inProgress: false })
  let resolveFinal
  let scanCalled = false
  const finalPromise = new Promise((resolve) => { resolveFinal = resolve })
  dom.window.pywebview = { api: {
    ping: async () => envelope('ok'),
    get_app_version: async () => envelope(current.app),
    get_current_config: async () => envelope(current.config),
    get_dashboard_state: async () => envelope(current),
    scan_reports: async () => { scanCalled = true; return envelope(await finalPromise) },
  } }

  function Harness() {
    const store = useStore()
    return React.createElement(React.Fragment, null,
      React.createElement('output', { 'data-testid': 'connected' }, store.connected ? 'yes' : 'no'),
      React.createElement('output', { 'data-testid': 'files' }, store.reports.map((item) => item.fileName).join(',')),
      React.createElement('output', { 'data-testid': 'busy' }, store.scanBusy ? 'yes' : 'no'),
      React.createElement('button', { 'data-testid': 'scan', onClick: () => { void store.scan() } }, 'scan'))
  }

  const { createRoot } = await import('react-dom/client')
  const root = createRoot(dom.window.document.getElementById('root'))
  try {
    await React.act(async () => { root.render(React.createElement(StoreProvider, null, React.createElement(Harness))) })
    const select = (id) => dom.window.document.querySelector(`[data-testid="${id}"]`)
    await waitFor(() => select('connected')?.textContent === 'yes', 'bridge startup')

    await React.act(async () => { select('scan').dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true })) })
    await waitFor(() => scanCalled && select('busy')?.textContent === 'yes', 'pending scan call')

    current = dashboard({ scanned: true, message: 'leaf 1', generation: 1, revision: 1, inProgress: true,
      incremental: true, completedLeaves: 1, totalLeaves: 2 }, [report('0', 'leaf-one.pptx')])
    await waitFor(() => select('files').textContent === 'leaf-one.pptx', 'first completed leaf poll', 2_000)
    assert.equal(select('files').textContent, 'leaf-one.pptx')
    assert.equal(select('busy').textContent, 'yes', 'the original batched scan call is still pending')

    current = dashboard({ scanned: true, message: 'leaf 2', generation: 1, revision: 2, inProgress: true,
      incremental: true, completedLeaves: 2, totalLeaves: 2 }, [report('1', 'leaf-two.pptx')])
    await waitFor(() => select('files').textContent === 'leaf-one.pptx,leaf-two.pptx',
      'second completed leaf poll', 2_000)
    assert.equal(select('files').textContent, 'leaf-one.pptx,leaf-two.pptx')

    current = dashboard({ scanned: true, message: 'delayed leaf 1', generation: 1, revision: 1, inProgress: true,
      incremental: true, completedLeaves: 1, totalLeaves: 2 }, [report('0', 'leaf-one.pptx')])
    await React.act(async () => { await delay(850) })
    assert.equal(select('files').textContent, 'leaf-one.pptx,leaf-two.pptx', 'older revision must not remove a row')

    const final = dashboard({ scanned: true, message: 'done', generation: 1, revision: 3, inProgress: false,
      completedLeaves: 2, totalLeaves: 2 }, [report('0', 'leaf-one.pptx'), report('1', 'leaf-two.pptx')])
    await React.act(async () => { resolveFinal(final); await delay(20) })
    await waitFor(() => select('busy')?.textContent === 'no', 'final scan response')
    assert.equal(select('files').textContent, 'leaf-one.pptx,leaf-two.pptx')
  } finally {
    await React.act(async () => { root.unmount() })
    dom.window.close()
    for (const [name, descriptor] of prior) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete globalThis[name]
    }
  }
})
