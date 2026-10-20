import assert from 'node:assert/strict'
import { before, after, test } from 'node:test'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import * as React from 'react'
import { fileURLToPath } from 'node:url'

// Current 036S APIs + real LearningTab in StrictMode. In particular, emit the
// browser click/click/dblclick sequence, not just a synthetic dblclick event.
const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
const envelope = (data) => ({ ok: true, data })
const failure = () => ({ ok: false, error: { code: 'LABEL_SAVE_FAILED', message: 'Disk full; retry.' } })
let server, StoreProvider, useStore, LearningTab
before(async () => {
  server = await createServer({ configFile: `${frontendRoot}/vite.config.ts`, root: frontendRoot,
    appType: 'custom', server: { middlewareMode: true, hmr: false } })
  ;({ StoreProvider, useStore } = await server.ssrLoadModule('/src/state/store.tsx'))
  ;({ LearningTab } = await server.ssrLoadModule('/src/tabs/LearningTab.tsx'))
})
after(async () => { await server?.close() })

function candidate(id, scope = 'report-a', slide = 5) {
  return { id, reviewId: `review:${scope}:${id}`, reportScope: scope,
    sourceFile: 'same.pptx', managementNumber: 'M', slide, pictureId: id,
    src: 'data:image/jpeg;base64,AA==', bounds: { x: 10, y: 20, w: 30, h: 40 },
    slideWidth: 1600, slideHeight: 900, targetBboxPct: { x: 10, y: 20, w: 30, h: 40 },
    slidePreview: 'data:image/jpeg;base64,AA==', slidePreviewWidth: 1600, slidePreviewHeight: 900,
    slidePreviewBackend: 'builtin', slidePreviewFaithful: false,
    evidenceRegionBboxPct: { x: 10, y: 20, w: 35, h: 45 }, evidenceRegionPictureCount: 1,
    itemIndex: Number(id.slice(-1)) || 0, itemHeading: 'Mục', evidence: [],
    decision: 'review', confidence: 0.6, confidenceBand: 'medium', excelEligible: false,
    userLabel: 'UNLABELED', labelPending: false, note: `note-${id}` }
}
function learning(images) {
  return { images, contents: [], candidateStatus: {
    image: { state: 'ready', reason: images.length ? 'available' : 'empty', count: images.length, message: '' },
    content: { state: 'ready', reason: 'empty', count: 0, message: '' },
  }, counts: { image: { total: images.length, labeled: images.filter((c) => c.userLabel !== 'UNLABELED').length },
    content: { total: 0, labeled: 0 } }, modelStatus: { image: 'Ready', content: 'Ready' },
  excelPending: 0, excelLastResult: { kind: 'none', message: '' }, training: false }
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
  const state = { store: null, images, fullCalls: 0, fast: [], single: [], bulk: [], legacy: [], defer: false, resolve: null }
  const config = { paths: { reportFolder: '', template: '', output: '' },
    period: { mode: 'all', month: '', year: '', from: '', to: '' }, forceReprocess: false,
    ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [] },
    update: { path: '', autoCheck: false } }
  const dashboard = { app: { name: 'Report Extractor', title: 'Report Extractor', version: '1.3.7', build: '020', buildNumber: 20 },
    config, scan: { scanned: false, message: '' }, reports: [],
    job: { status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0, currentFile: '', elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false },
    logs: [], ollama: { ...config.ollama, aiStatus: '', discovering: false, discoveryMessage: '', discoveryResults: [], discoveryChecked: 0, discoveryTotal: 0, serverApplyRunning: false, serverApplyMessage: '' },
    update: { path: '', autoCheck: false, status: 'idle', message: '', progress: 0, progressStage: '' },
    diagnostics: { running: false, rows: [], error: '' } }
  const update = (ids, label, notes, pending) => { state.images = state.images.map((c) => ids.includes(c.reviewId || c.id)
    ? { ...c, userLabel: label, note: notes[c.reviewId || c.id] ?? c.note, labelPending: pending } : c) }
  const save = async (ids, label, notes, kind) => {
    state[kind].push({ ids, label, notes })
    const ok = state.defer ? await new Promise((resolve) => { state.resolve = resolve }) : true
    if (!ok) return failure()
    update(ids, label, notes, false)
    return envelope({ candidateIds: ids, candidateId: ids[0], kind: 'image', label, labelPending: false,
      message: 'Saved', excelPending: 0, excelLastResult: { kind: 'ok', message: 'Saved' } })
  }
  dom.window.pywebview = { api: {
    ping: async () => envelope('ok'), get_app_version: async () => envelope(dashboard.app),
    get_current_config: async () => envelope(config), get_dashboard_state: async () => envelope(dashboard),
    get_learning_state: async () => { state.fullCalls++; return envelope(learning(state.images)) },
    set_learning_label_fast: async (kind, id, label, note) => {
      state.fast.push({ kind, id, label, note }); update([id], label, { [id]: note }, true)
      return envelope({ kind, candidateId: id, label, labelPending: true })
    },
    save_learning_candidate: async (kind, id, label, note) => save([id], label, { [id]: note }, 'single'),
    save_learning_images: async (ids, label, notes) => save(ids, label, notes, 'bulk'),
    set_learning_label: async (...args) => { state.legacy.push(args); assert.fail('old full-state set API used') },
    save_learning_labels: async (...args) => { state.legacy.push(args); assert.fail('old full-state save API used') },
  } }
  function Harness() { state.store = useStore(); return React.createElement(LearningTab) }
  const { createRoot } = await import('react-dom/client')
  const root = createRoot(dom.window.document.getElementById('root'))
  await React.act(async () => { root.render(React.createElement(React.StrictMode, null,
    React.createElement(StoreProvider, null, React.createElement(Harness)))) })
  await waitFor(() => state.store?.connected, 'connected')
  const container = dom.window.document.getElementById('root')
  const event = async (element, type = 'click', options = {}) => {
    assert.ok(element, `target for ${type}`)
    await React.act(async () => { element.dispatchEvent(new dom.window.MouseEvent(type, { bubbles: true, detail: 1, ...options })); await delay(0) })
  }
  await event([...container.querySelectorAll('aside nav button')].find((b) => b.textContent.trim() === 'Kiểm tra ảnh cải tiến'))
  await React.act(async () => { await state.store.refreshLearning() })
  return { container, state, event,
    item: (idx) => container.querySelector(`[data-testid="candidate-${idx}"]`),
    label: (text) => [...container.querySelectorAll('[data-testid="learning-label-controls"] button')].find((b) => b.textContent.trim() === text),
    async deliver(next) { state.images = next; await React.act(async () => { await state.store.refreshLearning() }) },
    async settle(ok) { await React.act(async () => { state.resolve(ok); await delay(20) }) },
    async cleanup() {
      await React.act(async () => { root.unmount() })
      dom.window.close()
      for (const [name, descriptor] of prior) {
        if (descriptor) Object.defineProperty(globalThis, name, descriptor); else delete globalThis[name]
      }
    },
  }
}

test('labels are first in right panel; single click only focuses; no-label double-click never guesses', async () => {
  const h = await harness([candidate('a1'), candidate('a2')])
  try {
    const panel = h.container.querySelector('section[aria-label="Thông tin và nhãn xác nhận"]')
    assert.equal(panel.querySelector('.overflow-y-auto').firstElementChild.dataset.testid, 'learning-label-controls')
    await h.event(h.item(1))
    assert.equal(h.state.single.length + h.state.bulk.length, 0)
    await h.event(h.item(1), 'dblclick', { detail: 2 })
    assert.equal(h.state.single.length + h.state.bulk.length, 0)
    assert.match(h.container.textContent, /Vui lòng chọn nhãn/)
    assert.ok(h.container.querySelector('[data-testid="overlay-picture-candidate"]'))
    assert.match(h.container.textContent, /Trình vẽ tích hợp/)
  } finally { await h.cleanup() }
})

test('native double-click sequence confirms once in StrictMode via the 036S atomic API', async () => {
  const h = await harness([candidate('a1')])
  try {
    await h.event(h.label('Sau cải tiến'))
    assert.equal(h.state.fast.length, 1)
    h.state.defer = true
    await h.event(h.item(0), 'click', { detail: 1 })
    await h.event(h.item(0), 'click', { detail: 2 })
    await h.event(h.item(0), 'dblclick', { detail: 2 })
    await h.event(h.item(0), 'dblclick', { detail: 2 })
    assert.equal(h.state.single.length, 1)
    assert.equal(h.state.bulk.length, 0)
    await h.settle(true)
    assert.match(h.container.textContent, /Đã xác nhận 1 ảnh/)
    assert.equal(h.state.single[0].ids[0], 'review:report-a:a1')
    assert.equal(h.state.legacy.length, 0)
  } finally { await h.cleanup() }
})

test('Ctrl/Shift selection is limited to report scope plus slide, even identical filenames/IDs', async () => {
  const h = await harness([candidate('a1'), candidate('a2'), candidate('a3'),
    candidate('a1', 'report-b'), candidate('b1', 'report-b', 6)])
  try {
    await h.event(h.item(0))
    await h.event(h.item(1), 'click', { ctrlKey: true })
    assert.equal(h.container.querySelectorAll('[data-selected="true"]').length, 2)
    await h.event(h.item(0))
    await h.event(h.item(2), 'click', { shiftKey: true })
    assert.equal(h.container.querySelectorAll('[data-selected="true"]').length, 3)
    await h.event(h.item(3), 'click', { ctrlKey: true })
    assert.equal(h.container.querySelectorAll('[data-selected="true"]').length, 1)
    assert.equal(h.item(3).dataset.selected, 'true')
    await h.event(h.item(4), 'click', { shiftKey: true })
    assert.equal(h.container.querySelectorAll('[data-selected="true"]').length, 1)
    assert.equal(h.item(4).dataset.selected, 'true')
    assert.equal(h.state.bulk.length + h.state.single.length, 0)
  } finally { await h.cleanup() }
})

test('bulk save is optimistic, bounded to one ack, defers polling and retains unsaved choices on failure/retry', async () => {
  const h = await harness([candidate('a1'), candidate('a2')])
  try {
    await h.event(h.item(0))
    await h.event(h.item(1), 'click', { ctrlKey: true })
    h.state.defer = true
    await h.event(h.label('Sau cải tiến'))
    assert.equal(h.state.bulk.length, 1)
    assert.deepEqual(h.state.store.learning.images.map((c) => [c.userLabel, c.labelPending]), [['AFTER', true], ['AFTER', true]])
    const fullBefore = h.state.fullCalls
    await React.act(async () => { await h.state.store.refreshLearning() })
    assert.equal(h.state.fullCalls, fullBefore, 'poll must defer while saving')
    await h.settle(false)
    assert.match(h.container.textContent, /Chưa lưu hết nhãn/)
    assert.deepEqual(h.state.store.learning.images.map((c) => [c.userLabel, c.labelPending]), [['AFTER', true], ['AFTER', true]])
    assert.equal(h.container.querySelectorAll('[data-selected="true"]').length, 2)
    h.state.defer = false
    await h.event(h.label('Sau cải tiến'))
    await waitFor(() => h.state.store.learning.images.every((c) => c.labelPending === false), 'bulk retry saved')
    assert.equal(h.state.bulk.length, 2)
    assert.equal(h.state.fast.length + h.state.legacy.length + h.state.single.length, 0)
    assert.deepEqual(h.state.bulk[1].notes, { 'review:report-a:a1': 'note-a1', 'review:report-a:a2': 'note-a2' })
    assert.equal(h.state.store.learning.counts.image.labeled, 2)
  } finally { await h.cleanup() }
})

test('double-click preserves existing multi-selection through both native clicks and makes one bulk save', async () => {
  const h = await harness([candidate('a1'), candidate('a2')])
  try {
    await h.event(h.label('Sau cải tiến')) // explicit active label, single optimistic fast set only
    await h.event(h.item(0))
    await h.event(h.item(1), 'click', { ctrlKey: true })
    await h.event(h.item(1), 'click', { detail: 1 })
    await h.event(h.item(1), 'click', { detail: 2 })
    await h.event(h.item(1), 'dblclick', { detail: 2 })
    await waitFor(() => h.state.bulk.length === 1, 'one bulk quick confirmation')
    assert.equal(h.state.bulk[0].ids.length, 2)
    assert.equal(h.state.single.length, 0)
    assert.equal(h.state.legacy.length, 0)
  } finally { await h.cleanup() }
})

test('same-length replacement prunes selection by identity; no stale candidate is bulk-labeled', async () => {
  const h = await harness([candidate('a1'), candidate('a2')])
  try {
    await h.event(h.item(0))
    await h.event(h.item(1), 'click', { ctrlKey: true })
    await h.deliver([candidate('c1'), candidate('c2')])
    assert.equal(h.container.querySelectorAll('[data-selected="true"]').length, 0)
    await h.event(h.label('Trước cải tiến'))
    assert.equal(h.state.bulk.length, 0)
    assert.equal(h.state.fast.at(-1).id, 'review:report-a:c2')
  } finally { await h.cleanup() }
})

test('Save keeps the current candidate label, not a sticky quick-label from the previous candidate', async () => {
  const h = await harness([candidate('a1'), { ...candidate('a2'), userLabel: 'BEFORE' }])
  try {
    await h.event(h.label('Sau cải tiến'))
    await h.event(h.item(1))
    await h.event(h.label('Lưu nhãn & ghi chú'))
    await waitFor(() => h.state.single.length === 1, 'save of already labeled candidate')
    assert.equal(h.state.single[0].label, 'BEFORE')
    assert.equal(h.state.single[0].ids[0], 'review:report-a:a2')
  } finally { await h.cleanup() }
})
