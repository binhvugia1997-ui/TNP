import assert from 'node:assert/strict'
import { before, after, test } from 'node:test'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import * as React from 'react'
import { fileURLToPath } from 'node:url'

// Explicit-save flow (Bug B): selecting/classifying a candidate — single click, multi-select
// label click, or double-click — must only STAGE the label in memory (one fast pending write
// per candidate). NOTHING persists until the Save button is pressed, and staging must not
// trigger a full get_learning_state reload. Real LearningTab in StrictMode, fake bridge.

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
const envelope = (data) => ({ ok: true, data })
let server, StoreProvider, useStore, LearningTab
before(async () => {
  server = await createServer({
    configFile: `${frontendRoot}/vite.config.ts`, root: frontendRoot,
    appType: 'custom', server: { middlewareMode: true, hmr: false },
  })
  ;({ StoreProvider, useStore } = await server.ssrLoadModule('/src/state/store.tsx'))
  ;({ LearningTab } = await server.ssrLoadModule('/src/tabs/LearningTab.tsx'))
})
after(async () => { await server?.close() })

function candidate(id, scope = 'report-a', slide = 5) {
  return {
    id, reviewId: `review:${scope}:${id}`, reportScope: scope,
    sourceFile: 'same.pptx', managementNumber: 'M', slide, pictureId: id,
    src: 'data:image/jpeg;base64,AA==', bounds: { x: 10, y: 20, w: 30, h: 40 },
    slideWidth: 1600, slideHeight: 900, targetBboxPct: { x: 10, y: 20, w: 30, h: 40 },
    slidePreview: 'data:image/jpeg;base64,AA==', slidePreviewWidth: 1600, slidePreviewHeight: 900,
    slidePreviewBackend: 'builtin', slidePreviewFaithful: false,
    evidenceRegionBboxPct: { x: 10, y: 20, w: 35, h: 45 }, evidenceRegionPictureCount: 1,
    itemIndex: Number(id.slice(-1)) || 0, itemHeading: 'Mục', evidence: [],
    decision: 'review', confidence: 0.6, confidenceBand: 'medium', excelEligible: false,
    userLabel: 'UNLABELED', labelPending: false, note: `note-${id}`,
  }
}
function learning(images) {
  return {
    images, contents: [],
    candidateStatus: {
      image: { state: 'ready', reason: images.length ? 'available' : 'empty', count: images.length, message: '' },
      content: { state: 'ready', reason: 'empty', count: 0, message: '' },
    },
    counts: {
      image: { total: images.length, labeled: images.filter((c) => c.userLabel !== 'UNLABELED').length },
      content: { total: 0, labeled: 0 },
    },
    modelStatus: { image: 'Ready', content: 'Ready' },
    excelPending: 0, excelLastResult: { kind: 'none', message: '' }, training: false,
  }
}
async function waitFor(predicate, description) {
  const end = Date.now() + 4000
  while (Date.now() < end) {
    if (predicate()) return
    await React.act(async () => { await delay(10) })
  }
  assert.fail(`timeout: ${description}`)
}

async function harness(images) {
  const dom = new JSDOM('<!doctype html><html><body><div id="root"></div></body></html>', { url: 'http://localhost/' })
  const names = ['window', 'document', 'navigator', 'HTMLElement', 'Node', 'Event', 'MouseEvent', 'ResizeObserver', 'IS_REACT_ACT_ENVIRONMENT']
  const prior = new Map(names.map((name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)]))
  const define = (name, value) => Object.defineProperty(globalThis, name, { configurable: true, writable: true, value })
  for (const name of names.slice(0, 7)) define(name, dom.window[name])
  define('IS_REACT_ACT_ENVIRONMENT', true)
  class Observer { observe() {} disconnect() {} }
  define('ResizeObserver', Observer)
  dom.window.ResizeObserver = Observer
  dom.window.HTMLElement.prototype.scrollIntoView = () => {}
  dom.window.matchMedia = (media) => ({ matches: true, media, addEventListener() {}, removeEventListener() {} })
  Object.defineProperty(dom.window.HTMLElement.prototype, 'clientWidth', { configurable: true, get: () => 1000 })
  Object.defineProperty(dom.window.HTMLElement.prototype, 'clientHeight', { configurable: true, get: () => 700 })
  const state = { store: null, images, fullCalls: 0, fast: [], single: [], bulk: [], legacy: [] }
  const config = {
    paths: { reportFolder: '', template: '', output: '' },
    period: { mode: 'all', month: '', year: '', from: '', to: '' }, forceReprocess: false,
    ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [] },
    update: { path: '', autoCheck: false },
  }
  const dashboard = {
    app: { name: 'Report Extractor', title: 'Report Extractor', version: '1.4.0', build: '022', buildNumber: 22 },
    config, scan: { scanned: false, message: '' }, reports: [],
    job: { status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0, currentFile: '', elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false },
    logs: [], ollama: { ...config.ollama, aiStatus: '', discovering: false, discoveryMessage: '', discoveryResults: [], discoveryChecked: 0, discoveryTotal: 0, serverApplyRunning: false, serverApplyMessage: '' },
    update: { path: '', autoCheck: false, status: 'idle', message: '', progress: 0, progressStage: '' },
    diagnostics: { running: false, rows: [], error: '' },
  }
  const update = (ids, label, notes, pending) => {
    state.images = state.images.map((c) => ids.includes(c.reviewId || c.id)
      ? { ...c, userLabel: label, note: notes[c.reviewId || c.id] ?? c.note, labelPending: pending } : c)
  }
  const save = async (ids, label, notes, kind) => {
    state[kind].push({ ids, label, notes })
    update(ids, label, notes, false)
    return envelope({
      candidateIds: ids, candidateId: ids[0], kind: 'image', label, labelPending: false,
      message: 'Saved', excelPending: 0, excelLastResult: { kind: 'ok', message: 'Saved' },
    })
  }
  dom.window.pywebview = {
    api: {
      ping: async () => envelope('ok'),
      get_app_version: async () => envelope(dashboard.app),
      get_current_config: async () => envelope(config),
      get_dashboard_state: async () => envelope(dashboard),
      get_learning_state: async () => { state.fullCalls++; return envelope(learning(state.images)) },
      set_learning_label_fast: async (kind, id, label, note) => {
        state.fast.push({ kind, id, label, note })
        update([id], label, { [id]: note }, true)
        return envelope({ kind, candidateId: id, label, labelPending: true })
      },
      save_learning_candidate: async (kind, id, label, note) => save([id], label, { [id]: note }, 'single'),
      save_learning_images: async (ids, label, notes) => save(ids, label, notes, 'bulk'),
      set_learning_label: async (...args) => { state.legacy.push(args); assert.fail('old full-state set API used') },
      save_learning_labels: async (...args) => { state.legacy.push(args); assert.fail('old full-state save API used') },
    },
  }
  function Harness() { state.store = useStore(); return React.createElement(LearningTab) }
  const { createRoot } = await import('react-dom/client')
  const root = createRoot(dom.window.document.getElementById('root'))
  await React.act(async () => {
    root.render(React.createElement(React.StrictMode, null,
      React.createElement(StoreProvider, null, React.createElement(Harness))))
  })
  await waitFor(() => state.store?.connected, 'connected')
  const container = dom.window.document.getElementById('root')
  const event = async (element, type = 'click', options = {}) => {
    assert.ok(element, `target for ${type}`)
    await React.act(async () => {
      element.dispatchEvent(new dom.window.MouseEvent(type, { bubbles: true, detail: 1, ...options }))
      await delay(0)
    })
  }
  await event([...container.querySelectorAll('aside nav button')].find((b) => b.textContent.trim() === 'Kiểm tra ảnh cải tiến'))
  await React.act(async () => { await state.store.refreshLearning() })
  return {
    container, state, event,
    item: (idx) => container.querySelector(`[data-testid="candidate-${idx}"]`),
    label: (text) => [...container.querySelectorAll('[data-testid="learning-label-controls"] button')].find((b) => b.textContent.trim() === text),
    async cleanup() {
      await React.act(async () => { root.unmount() })
      dom.window.close()
      for (const [name, descriptor] of prior) {
        if (descriptor) Object.defineProperty(globalThis, name, descriptor)
        else delete globalThis[name]
      }
    },
  }
}

test('single label click stages only; Save persists the staged label', async () => {
  const h = await harness([candidate('a1')])
  try {
    const fullBefore = h.state.fullCalls
    await h.event(h.label('Sau cải tiến'))
    assert.equal(h.state.fast.length, 1, 'one staged pending write')
    assert.equal(h.state.fast[0].label, 'AFTER')
    assert.equal(h.state.single.length, 0, 'no save on selection')
    assert.equal(h.state.bulk.length, 0)
    assert.equal(h.state.store.learning.images[0].labelPending, true, 'staged label shows as unsaved')
    await React.act(async () => { await delay(50) })
    assert.equal(h.state.fullCalls, fullBefore, 'staging must not trigger a full-state reload')
    await h.event(h.label('Lưu nhãn & ghi chú'))
    await waitFor(() => h.state.single.length === 1, 'Save persists')
    assert.equal(h.state.single[0].label, 'AFTER')
    assert.equal(h.state.store.learning.images[0].labelPending, false)
    assert.equal(h.state.legacy.length, 0)
  } finally { await h.cleanup() }
})

test('multi-select label click stages every selected candidate; Save makes one bulk save', async () => {
  const h = await harness([candidate('a1'), candidate('a2')])
  try {
    await h.event(h.item(0))
    await h.event(h.item(1), 'click', { ctrlKey: true })
    assert.equal(h.container.querySelectorAll('[data-selected="true"]').length, 2)
    await h.event(h.label('Trước cải tiến'))
    assert.equal(h.state.fast.length, 2, 'one staged pending write per selected candidate')
    assert.equal(h.state.single.length, 0, 'no save on selection')
    assert.equal(h.state.bulk.length, 0)
    assert.deepEqual(h.state.store.learning.images.map((c) => [c.userLabel, c.labelPending]),
      [['BEFORE', true], ['BEFORE', true]])
    await h.event(h.label('Lưu nhãn & ghi chú'))
    await waitFor(() => h.state.bulk.length === 1, 'Save persists all staged labels at once')
    assert.equal(h.state.bulk[0].ids.length, 2)
    assert.equal(h.state.bulk[0].label, 'BEFORE')
    assert.equal(h.state.legacy.length, 0)
  } finally { await h.cleanup() }
})

test('double-click stages only; Save persists', async () => {
  const h = await harness([candidate('a1')])
  try {
    await h.event(h.label('Sau cải tiến'))
    assert.equal(h.state.fast.length, 1)
    await h.event(h.item(0), 'click', { detail: 1 })
    await h.event(h.item(0), 'click', { detail: 2 })
    await h.event(h.item(0), 'dblclick', { detail: 2 })
    assert.equal(h.state.single.length, 0, 'double-click must stage only, never save')
    assert.equal(h.state.bulk.length, 0)
    assert.equal(h.state.store.learning.images[0].userLabel, 'AFTER')
    await h.event(h.label('Lưu nhãn & ghi chú'))
    await waitFor(() => h.state.single.length === 1, 'Save persists the staged label')
    assert.equal(h.state.single[0].label, 'AFTER')
  } finally { await h.cleanup() }
})
