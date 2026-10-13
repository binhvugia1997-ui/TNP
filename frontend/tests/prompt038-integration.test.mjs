/**
 * PROMPT-038 — integration acceptance regressions (B normalisation + C preview lifecycle).
 *
 * Bug B: the ONE authoritative visibility rule must hide every "Bỏ qua" classification even if a
 * future DTO pathway emits a non-canonical status string (casing/whitespace). Canonical statuses
 * are covered by prompt037; here the predicate itself is proven robust.
 *
 * Bug C: ImageReview early-returns while no candidate exists, so a mount-once ResizeObserver
 * effect never sees the viewport element that appears when the async candidate arrives. These
 * tests mount the real LearningTab with NO candidate, then deliver one asynchronously, with a
 * real (stubbed-geometry) viewport: 100% must fit the ACTUAL viewport, 400% must be 4x that fit,
 * resizes and candidate switches must recompute, and overlays stay percentage-aligned.
 */
import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import * as React from 'react'
import { fileURLToPath } from 'node:url'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const delay = (ms) => new Promise((resolve) => globalThis.setTimeout(resolve, ms))
const envelope = (data) => ({ ok: true, data })

let server
let isVisibleReport
let presentReports
let normalizeReportStatus
let StoreProvider
let useStore
let LearningTab

before(async () => {
  server = await createServer({
    configFile: `${frontendRoot}/vite.config.ts`,
    root: frontendRoot,
    appType: 'custom',
    server: { middlewareMode: true, hmr: false },
  })
  ;({ isVisibleReport, presentReports, normalizeReportStatus } = await server.ssrLoadModule('/src/types.ts'))
  ;({ StoreProvider, useStore } = await server.ssrLoadModule('/src/state/store.tsx'))
  ;({ LearningTab } = await server.ssrLoadModule('/src/tabs/LearningTab.tsx'))
})

after(async () => { await server?.close() })

/* ------------------------------------------------------------------ Bug B */

test('PROMPT-038 B: visibility rule hides non-canonical spellings of the four skip statuses', () => {
  for (const raw of ['outside_period', 'OUTSIDE_PERIOD', ' Outside_Period ', 'skipped', 'SKIPPED',
    ' source_duplicate', 'SOURCE_DUPLICATE ', 'fast_skip', 'Fast_Skip']) {
    assert.equal(isVisibleReport({ status: raw }), false, `${JSON.stringify(raw)} must be hidden`)
    assert.equal(normalizeReportStatus(raw), raw.trim().toLowerCase())
  }
  for (const raw of ['waiting', 'completed', 'WAITING', 'needs_review', 'error', 'excluded', '']) {
    assert.equal(isVisibleReport({ status: raw }), true, `${JSON.stringify(raw)} must stay visible`)
  }
  const visible = presentReports([
    { id: '1', stt: 9, status: 'OUTSIDE_PERIOD' },
    { id: '2', stt: 8, status: 'waiting' },
    { id: '3', stt: 7, status: ' fast_skip ' },
  ])
  assert.deepEqual(visible.map((r) => r.id), ['2'])
  assert.equal(visible[0].stt, 1, 'STT renumbered over visible rows only')
})

/* ------------------------------------------------------------------ Bug C */

const PREVIEW = 'data:image/svg+xml;base64,PHN2ZyBpZD0ic2xpZGUiLz4='

function emptyLearning(images = []) {
  return {
    images, contents: [],
    candidateStatus: {
      image: { state: 'ready', reason: images.length ? 'available' : 'empty', count: images.length, message: '' },
      content: { state: 'ready', reason: 'empty', count: 0, message: '' },
    },
    counts: { image: { total: images.length, labeled: 0 }, content: { total: 0, labeled: 0 } },
    modelStatus: { image: 'Ready', content: 'Ready' },
    excelPending: 0, excelLastResult: { kind: 'none', message: '' }, training: false,
  }
}

function candidate({ id, slideW = 9144000, slideH = 5143500 }) {
  return {
    id, sourceFile: 'r.pptx', managementNumber: 'TNP-001', slide: 5, pictureId: '7',
    src: 'data:image/png;base64,AA==',
    bounds: { x: 50, y: 35, w: 12, h: 14 },
    slideWidth: slideW, slideHeight: slideH,
    targetBboxPct: { x: 50, y: 35, w: 12, h: 14 }, targetKind: 'picture',
    slidePreview: PREVIEW, slidePreviewWidth: 1600, slidePreviewHeight: 900,
    slidePreviewBackend: 'powerpoint', slidePreviewFaithful: true,
    evidenceRegionBbox: null, evidenceRegionBboxPct: { x: 47, y: 31, w: 33, h: 24 },
    evidenceRegionKind: 'after_region', evidenceRegionItemIndex: 0, evidenceRegionItemHeading: 'Mục một',
    evidenceRegionPictureCount: 1, itemIndex: 0, itemHeading: 'Mục một',
    decision: 'AFTER', confidenceBand: 'high', confidence: 0.9, evidence: [],
    excelEligible: true, userLabel: 'UNLABELED', note: '',
  }
}

async function mountReview() {
  const dom = new JSDOM('<!doctype html><html><body><div id="root"></div></body></html>', { url: 'http://localhost/' })
  const names = ['window', 'document', 'navigator', 'HTMLElement', 'Node', 'Event', 'MouseEvent', 'ResizeObserver', 'IS_REACT_ACT_ENVIRONMENT']
  const prior = new Map(names.map((n) => [n, Object.getOwnPropertyDescriptor(globalThis, n)]))
  const define = (n, v) => Object.defineProperty(globalThis, n, { configurable: true, writable: true, value: v })
  define('window', dom.window); define('document', dom.window.document); define('navigator', dom.window.navigator)
  define('HTMLElement', dom.window.HTMLElement); define('Node', dom.window.Node); define('Event', dom.window.Event)
  define('MouseEvent', dom.window.MouseEvent); define('IS_REACT_ACT_ENVIRONMENT', true)

  // Controllable real geometry: jsdom reports 0x0, so the viewport size is injected here and
  // ResizeObserver notifications are fired manually — exactly what a real resize does.
  const viewportSize = { w: 1000, h: 700 }
  Object.defineProperty(dom.window.HTMLElement.prototype, 'clientWidth', { configurable: true, get() { return viewportSize.w } })
  Object.defineProperty(dom.window.HTMLElement.prototype, 'clientHeight', { configurable: true, get() { return viewportSize.h } })
  const observers = []
  class TestResizeObserver {
    constructor(cb) { this.cb = cb; this.el = null; observers.push(this) }
    observe(el) { this.el = el }
    unobserve() {}
    disconnect() { this.el = null }
    fire() { if (this.el) this.cb() }
  }
  define('ResizeObserver', TestResizeObserver)
  dom.window.ResizeObserver = TestResizeObserver
  dom.window.HTMLElement.prototype.scrollIntoView = () => {}
  dom.window.matchMedia = (media) => ({ matches: media === '(min-width: 1280px)', media, onchange: null,
    addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {}, dispatchEvent() { return true } })

  const config = { paths: { reportFolder: '', template: '', output: '' },
    period: { mode: 'auto', month: '', year: '', from: '', to: '' }, forceReprocess: false,
    ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [] },
    update: { path: '', autoCheck: true } }
  const dashboard = { app: { name: 'Report Extractor', title: 'Report Extractor', version: '1.3.6', build: '019', buildNumber: 19 },
    config, scan: { scanned: false, message: '' }, reports: [],
    job: { status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0, currentFile: '', elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false },
    logs: [], ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [], aiStatus: '', discovering: false, discoveryMessage: '', discoveryResults: [], discoveryChecked: 0, discoveryTotal: 0, serverApplyRunning: false, serverApplyMessage: '' },
    update: { path: '', autoCheck: false, status: 'idle', message: '', progress: 0, progressStage: '' },
    diagnostics: { running: false, rows: [], error: '' } }
  const state = { learning: emptyLearning(), store: null }
  dom.window.pywebview = { api: {
    ping: async () => envelope('ok'),
    get_app_version: async () => envelope(dashboard.app),
    get_current_config: async () => envelope(config),
    get_dashboard_state: async () => envelope(dashboard),
    get_learning_state: async () => envelope(state.learning),
  } }
  function Harness() {
    state.store = useStore()
    return React.createElement(React.Fragment, null,
      React.createElement('button', { type: 'button', 'data-testid': 'refresh-learning', onClick: () => { void state.store.refreshLearning() } }, 'refresh'),
      React.createElement(LearningTab))
  }
  const { createRoot } = await import('react-dom/client')
  const root = createRoot(dom.window.document.getElementById('root'))
  await React.act(async () => { root.render(React.createElement(StoreProvider, null, React.createElement(Harness))) })
  const container = dom.window.document.getElementById('root')
  const click = async (el) => { assert.ok(el, 'click target'); await React.act(async () => { el.dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true })); await delay(0) }) }
  const waitFor = async (pred, desc, ms = 4000) => {
    const end = Date.now() + ms
    while (Date.now() < end) { if (pred()) return; await React.act(async () => { await delay(10) }) }
    assert.fail(`timeout: ${desc}`)
  }
  await waitFor(() => state.store?.connected === true, 'bridge ready')
  const menu = [...container.querySelectorAll('aside nav button')].find((b) => b.textContent.trim() === 'Kiểm tra ảnh cải tiến')
  await click(menu)

  return {
    container, viewportSize, observers,
    async deliver(images) {
      state.learning = emptyLearning(images)
      await click(container.querySelector('[data-testid="refresh-learning"]'))
    },
    async resize(w, h) {
      viewportSize.w = w; viewportSize.h = h
      await React.act(async () => { for (const o of observers) o.fire(); await delay(0) })
    },
    async clickEl(el) { await click(el) },
    slideBox() {
      const img = container.querySelector('[data-testid="slide-preview-bitmap"]')
      if (!img) return null
      const style = img.parentElement.getAttribute('style') || ''
      const read = (n) => { const m = style.match(new RegExp(`${n}:\\s*([0-9.]+)px`)); return m ? Number(m[1]) : null }
      return { width: read('width'), height: read('height') }
    },
    async cleanup() {
      await React.act(async () => { root.unmount() })
      dom.window.close()
      for (const [n, d] of prior) { if (d) Object.defineProperty(globalThis, n, d); else delete globalThis[n] }
    },
  }
}

test('PROMPT-038 C: async candidate arrival measures the real viewport (not the 80px fallback)', async () => {
  const h = await mountReview()
  try {
    // No candidate yet: ImageReview shows only the empty feedback card (viewport DOM absent).
    assert.equal(h.container.querySelector('[data-testid="slide-preview-bitmap"]'), null)
    await h.deliver([candidate({ id: 'c1' })])
    // 100% fit for a 1000x700 viewport, pad 24: avail 952x652, 16:9 -> 952 x 535.5
    await h.resize(1000, 700)
    const box = h.slideBox()
    assert.ok(box, 'slide preview rendered after async candidate')
    assert.ok(Math.abs(box.width - 952) < 1.5, `100% width fits viewport: got ${box.width}`)
    assert.ok(Math.abs(box.height - 535.5) < 1.5, `aspect preserved: got ${box.height}`)
    assert.ok(box.width > 500, 'must not be the ~80px fallback size')
    // Overlays stay exactly on the DTO percentages regardless of display size.
    const pic = h.container.querySelector('[data-testid="overlay-picture-candidate"]')
    assert.match(pic.getAttribute('style'), /left:\s*50%/)
    assert.match(pic.getAttribute('style'), /width:\s*12%/)
    const ev = h.container.querySelector('[data-testid="overlay-evidence-region"]')
    assert.match(ev.getAttribute('style'), /left:\s*47%/)
    assert.match(ev.getAttribute('style'), /width:\s*33%/)
  } finally { await h.cleanup() }
})

test('PROMPT-038 C: 400% is 4x the real fit, resize recomputes, candidate switch stays measured', async () => {
  const h = await mountReview()
  try {
    await h.deliver([candidate({ id: 'c1' }), candidate({ id: 'c2', slideW: 9144000, slideH: 6858000 })])
    await h.resize(1000, 700)
    const fit = h.slideBox()
    assert.ok(Math.abs(fit.width - 952) < 1.5, `fit ${JSON.stringify(fit)}`)
    const zoomIn = [...h.container.querySelectorAll('button')].find((b) => b.getAttribute('aria-label') === 'Phóng to')
    for (let i = 0; i < 8; i += 1) await h.clickEl(zoomIn)
    const zoomed = h.slideBox()
    assert.ok(Math.abs(zoomed.width - fit.width * 4) < 6, `400% = 4x fit: ${zoomed.width} vs ${fit.width * 4}`)
    assert.ok(Math.abs(zoomed.height - fit.height * 4) < 6, 'zoomed height = 4x fit height')
    const fitBtn = [...h.container.querySelectorAll('button')].find((b) => b.textContent.includes('Vừa khung'))
    await h.clickEl(fitBtn)
    // Resize to 500x400: avail 452x352, 16:9 -> 452 x 254.25
    await h.resize(500, 400)
    const small = h.slideBox()
    assert.ok(Math.abs(small.width - 452) < 1.5, `resize recomputes fit: ${small.width}`)
    assert.ok(Math.abs(small.height - 254.25) < 1.5)
    // Switch to the 4:3 candidate: viewport 500x400 -> avail 452x352; 4:3 height-limited: 352 x 469.33? no:
    // dispW=452, dispH=339 <= 352, so 452 x 339.
    const items = [...h.container.querySelectorAll('section[aria-label="Đối tượng ảnh"] button')]
    await h.clickEl(items[items.length - 1])
    const switched = h.slideBox()
    assert.ok(switched && switched.width > 400, `candidate switch remains measured: ${JSON.stringify(switched)}`)
    assert.ok(Math.abs(switched.width / switched.height - 4 / 3) < 0.02, 'new candidate aspect preserved')
  } finally { await h.cleanup() }
})
