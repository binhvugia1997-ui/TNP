import assert from 'node:assert/strict'
import { test } from 'node:test'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import * as React from 'react'
import { fileURLToPath } from 'node:url'

// PROMPT-036S: a Learning label click must update local state immediately, persist
// asynchronously through a coalescing per-candidate worker (bounded calls, last choice
// wins), never be reverted by stale/deferred polling, and stay visible as unsaved when
// persistence fails until a Save retry recovers it. Save = ONE atomic bridge call.

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const delay = (ms) => new Promise((resolve) => globalThis.setTimeout(resolve, ms))

function deferred() {
  let resolve
  let reject
  const promise = new Promise((res, rej) => { resolve = res; reject = rej })
  return { promise, resolve, reject }
}

const envelope = (data) => ({ ok: true, data })
const failure = (message) => ({ ok: false, error: { code: 'LABEL_SAVE_FAILED', message } })

function learningState({ imageLabel = 'UNLABELED', imagePending = false, contentLabel = 'UNLABELED' } = {}) {
  return {
    images: [{
      id: 'img-1', sourceFile: 'report.pptx', managementNumber: 'TNP-001', slide: 3, pictureId: '9',
      src: 'data:image/png;base64,AA==', bounds: { x: 1, y: 2, w: 3, h: 4 },
      decision: 'AFTER', confidenceBand: 'high', confidence: 0.9, evidence: [],
      excelEligible: true, userLabel: imageLabel, labelPending: imagePending, note: '',
    }],
    contents: [{
      id: 'content-1', sourceFile: 'report.pptx', managementNumber: 'TNP-001', slide: 3, blockId: '22',
      text: 'Nội dung cải tiến nguyên văn', decision: 'include', confidenceBand: 'high', confidence: 0.9,
      evidence: [], section: 'improvement', nearestTitle: 'Cải tiến', userLabel: contentLabel, note: '',
    }],
    candidateStatus: {
      image: { state: 'ready', reason: 'available', count: 1, message: 'Đã tải 1 ảnh ứng viên từ lượt xử lý hiện tại.' },
      content: { state: 'ready', reason: 'available', count: 1, message: 'Đã tải 1 khối nội dung từ lượt xử lý hiện tại.' },
    },
    counts: {
      image: { total: 1, labeled: imageLabel === 'UNLABELED' ? 0 : 1 },
      content: { total: 1, labeled: contentLabel === 'UNLABELED' ? 0 : 1 },
    },
    modelStatus: { image: 'Ready', content: 'Ready' },
    excelPending: 0,
    excelLastResult: { kind: 'none', message: '' },
    training: false,
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

test('Learning labels: optimistic click, coalesced persist, stale-poll guard, visible failure, atomic save', async () => {
  const dom = new JSDOM('<!doctype html><html><body><div id="root"></div></body></html>', { url: 'http://localhost/' })
  const globalNames = ['window', 'document', 'navigator', 'HTMLElement', 'Node', 'Event', 'MouseEvent', 'ResizeObserver', 'IS_REACT_ACT_ENVIRONMENT']
  const priorGlobals = new Map(globalNames.map((name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)]))
  const defineGlobal = (name, value) => Object.defineProperty(globalThis, name, { configurable: true, writable: true, value })
  defineGlobal('window', dom.window)
  defineGlobal('document', dom.window.document)
  defineGlobal('navigator', dom.window.navigator)
  defineGlobal('HTMLElement', dom.window.HTMLElement)
  defineGlobal('Node', dom.window.Node)
  defineGlobal('Event', dom.window.Event)
  defineGlobal('MouseEvent', dom.window.MouseEvent)
  defineGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  class TestResizeObserver { observe() {} unobserve() {} disconnect() {} }
  defineGlobal('ResizeObserver', TestResizeObserver)
  dom.window.ResizeObserver = TestResizeObserver

  let server
  let root
  try {
    server = await createServer({
      configFile: `${frontendRoot}/vite.config.ts`, root: frontendRoot, appType: 'custom',
      server: { middlewareMode: true, hmr: false },
    })
    const { StoreProvider, useStore } = await server.ssrLoadModule('/src/state/store.tsx')
    const { createRoot } = await import('react-dom/client')

    const config = {
      paths: { reportFolder: '', template: '', output: '' },
      period: { mode: 'auto', month: '', year: '', from: '', to: '' },
      forceReprocess: false,
      ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [] },
      update: { path: '', autoCheck: true },
    }
    const dashboard = {
      app: { name: 'Report Extractor', title: 'Report Extractor', version: '1.3.7', build: '020', buildNumber: 20 },
      config,
      scan: { scanned: false, message: 'No reports selected.' },
      reports: [],
      job: { status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0, currentFile: '', elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false },
      logs: [],
      ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [], aiStatus: '', discovering: false, discoveryMessage: '', discoveryResults: [], discoveryChecked: 0, discoveryTotal: 0, serverApplyRunning: false, serverApplyMessage: '' },
      update: { path: '', autoCheck: true, status: 'idle', message: '', progress: 0, progressStage: '' },
      diagnostics: { running: false, rows: [], error: '' },
    }

    // Controllable fake backend.
    const backend = { imageLabel: 'UNLABELED', imagePending: false, contentLabel: 'UNLABELED' }
    let lsMode = 'backend' // 'backend' | 'stale' | 'manual'
    let lsCalls = 0
    const lsDeferreds = []
    const fastCalls = [] // { kind, id, label, note, deferred|null, fail }
    let fastFailNext = false
    const saveCalls = []
    const legacyCalls = []

    dom.window.pywebview = {
      api: {
        ping: async () => envelope('Python bridge OK'),
        get_app_version: async () => envelope(dashboard.app),
        get_current_config: async () => envelope(config),
        get_dashboard_state: async () => envelope(dashboard),
        get_learning_state: async () => {
          lsCalls += 1
          if (lsMode === 'manual') {
            const d = deferred()
            lsDeferreds.push(d)
            return d.promise
          }
          if (lsMode === 'stale') return envelope(learningState())
          return envelope(learningState({
            imageLabel: backend.imageLabel, imagePending: backend.imagePending, contentLabel: backend.contentLabel,
          }))
        },
        set_learning_label_fast: async (kind, id, label, note) => {
          const call = { kind, id, label, note, deferred: deferred() }
          fastCalls.push(call)
          if (fastFailNext) {
            fastFailNext = false
            return failure('Không ghi được nhãn học.')
          }
          return call.deferred.promise
        },
        save_learning_candidate: async (kind, id, label, note) => {
          saveCalls.push({ kind, id, label, note })
          backend.imageLabel = label
          backend.imagePending = false
          return envelope({
            kind, candidateId: id, label, labelPending: false, message: 'Đã lưu 1 nhãn xác nhận.',
            excelPending: 0, excelLastResult: { kind: 'ok', message: 'Đã lưu 1 nhãn xác nhận.' },
          })
        },
        set_learning_label: async (...args) => { legacyCalls.push(['set_learning_label', ...args]); return failure('legacy path must not be used') },
        save_learning_labels: async (...args) => { legacyCalls.push(['save_learning_labels', ...args]); return failure('legacy path must not be used') },
      },
    }

    const ackFast = (call) => {
      if (call.kind === 'image') { backend.imageLabel = call.label; backend.imagePending = true }
      else backend.contentLabel = call.label
      call.deferred.resolve(envelope({ kind: call.kind, candidateId: call.id, label: call.label, labelPending: true }))
    }

    const probe = { current: null }
    function Harness() {
      const store = useStore()
      probe.current = store
      const img = store.learning.images[0]
      const content = store.learning.contents[0]
      return React.createElement(
        React.Fragment, null,
        React.createElement('output', { 'data-testid': 'img' }, img ? `${img.userLabel}|${img.labelPending ? 'pending' : 'saved'}` : 'none'),
        React.createElement('output', { 'data-testid': 'content' }, content ? `${content.userLabel}|${content.labelPending ? 'pending' : 'saved'}` : 'none'),
        React.createElement('output', { 'data-testid': 'error' }, store.errorMessage),
        React.createElement('output', { 'data-testid': 'connected' }, store.connected ? 'ready' : 'disconnected'),
      )
    }
    const text = (id) => dom.window.document.querySelector(`[data-testid="${id}"]`)?.textContent
    const click = async (fn) => { await React.act(async () => { fn(probe.current); await delay(0) }) }

    root = createRoot(dom.window.document.getElementById('root'))
    await React.act(async () => {
      root.render(React.createElement(React.StrictMode, null, React.createElement(StoreProvider, null, React.createElement(Harness))))
    })
    await React.act(async () => { dom.window.dispatchEvent(new dom.window.Event('pywebviewready')) })
    await waitFor(() => text('connected') === 'ready', 'bridge handshake')
    await click((s) => s.setTab('learning'))
    await waitFor(() => text('img') === 'UNLABELED|saved', 'initial learning state')

    // 1) Click updates UI immediately, before the (hung) backend responds.
    await click((s) => s.labelImage('img-1', 'BEFORE'))
    assert.equal(text('img'), 'BEFORE|pending', 'click must render locally without waiting for the bridge')
    assert.equal(fastCalls.length, 1)
    assert.equal(fastCalls[0].label, 'BEFORE')

    // 2) Rapid clicks during the flight coalesce: still 1 call; latest wins on drain.
    await click((s) => s.labelImage('img-1', 'AFTER'))
    await click((s) => s.labelImage('img-1', 'CONTROL'))
    assert.equal(text('img'), 'CONTROL|pending')
    assert.equal(fastCalls.length, 1, 'superseded clicks must not create a request backlog')
    await React.act(async () => { ackFast(fastCalls[0]); await delay(20) })
    assert.equal(fastCalls.length, 2, 'worker persists exactly the latest superseded value')
    assert.equal(fastCalls[1].label, 'CONTROL')
    await React.act(async () => { ackFast(fastCalls[1]); await delay(20) })
    assert.equal(backend.imageLabel, 'CONTROL', 'final persisted value must be the last click')

    // 3) A refresh requested during a write is deferred; a stale response cannot revert.
    await click((s) => s.labelImage('img-1', 'AFTER'))
    assert.equal(fastCalls.length, 3)
    const lsBefore = lsCalls
    await click((s) => { void s.refreshLearning() })
    assert.equal(lsCalls, lsBefore, 'polling must not issue get_learning_state while a label write is in flight')
    lsMode = 'stale'
    await React.act(async () => { ackFast(fastCalls[2]); await delay(30) })
    await waitFor(() => lsCalls > lsBefore, 'deferred refresh after writes drain')
    assert.equal(text('img'), 'AFTER|pending', 'stale get_learning_state must not revert the optimistic label')
    lsMode = 'backend'

    // 4) A response that overlapped a click is discarded by the epoch guard.
    lsMode = 'manual'
    let refreshPromise
    await click((s) => { refreshPromise = s.refreshLearning() })
    await waitFor(() => lsDeferreds.length > 0, 'manual refresh in flight')
    await click((s) => s.labelImage('img-1', 'BEFORE'))
    assert.equal(text('img'), 'BEFORE|pending')
    const staleFlight = lsDeferreds.shift()
    await React.act(async () => {
      staleFlight.resolve(envelope(learningState({ imageLabel: 'AFTER', imagePending: true })))
      await delay(20)
    })
    assert.equal(text('img'), 'BEFORE|pending', 'overlapping stale response must be discarded')
    // Let the fast-set land; the queued re-refresh then reconciles from the backend.
    const lastFast = fastCalls[fastCalls.length - 1]
    lsMode = 'backend'
    await React.act(async () => { ackFast(lastFast); await delay(30) })
    await waitFor(() => text('img') === 'BEFORE|pending' && backend.imageLabel === 'BEFORE', 'reconciled backend state')
    await refreshPromise

    // 5) Persistence failure stays visible as unsaved; Save retries and recovers atomically.
    fastFailNext = true
    await click((s) => s.labelImage('img-1', 'CONTROL'))
    assert.equal(text('img'), 'CONTROL|pending')
    await waitFor(() => text('error').includes('Không ghi được nhãn'), 'visible persistence failure')
    assert.equal(text('img'), 'CONTROL|pending', 'failed label must stay visible and marked unsaved, never dropped')
    await click((s) => { void s.saveImageCandidate('img-1', 'CONTROL', '') })
    await waitFor(() => text('img') === 'CONTROL|saved', 'save retry recovers the label')
    assert.equal(saveCalls.length, 1, 'Save candidate must be exactly one atomic bridge call')
    assert.deepEqual(legacyCalls, [], 'old set_learning_label/save_learning_labels sequence must not be used')

    // 6) Content labels share the same optimistic path.
    const contentFastBefore = fastCalls.filter((c) => c.kind === 'content').length
    await click((s) => s.labelContent('content-1', 'IMPROVEMENT_CONTENT'))
    assert.equal(text('content'), 'IMPROVEMENT_CONTENT|pending', 'content click must also render immediately')
    const contentCall = fastCalls.filter((c) => c.kind === 'content')[contentFastBefore]
    await React.act(async () => { ackFast(contentCall); await delay(20) })
    assert.equal(backend.contentLabel, 'IMPROVEMENT_CONTENT')
  } finally {
    if (root) await React.act(async () => { root.unmount() })
    if (server) await server.close()
    dom.window.close()
    for (const [name, descriptor] of priorGlobals) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete globalThis[name]
    }
  }
})
