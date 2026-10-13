/**
 * PROMPT-036 — "Học cải tiến" click/input responsiveness.
 *
 * Assessment evidence (`.specify/bugs/prompt036-learning-ui-lag/assessment.md`):
 *   M4  a selection was only visible AFTER the whole bridge round trip (418.6 ms with a 400 ms bridge);
 *   M5  a `get_learning_state` snapshot already in flight when the reviewer clicked was delivered after the
 *       `set_learning_label` reply and visibly REVERTED the accepted selection (the reported "click swallowed").
 *
 * These tests pin the fixed behaviour: the selection is on screen inside the click event, rapid
 * Trước <-> Sau switching loses nothing, a stale snapshot can never revert a newer local edit, a backend
 * rejection is rolled back and surfaced, and the PROMPT-032 refresh/feedback contracts still hold.
 *
 * Linux/Node + jsdom only: this is not Windows pywebview or PowerPoint COM acceptance.
 */
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import * as React from 'react'
import { fileURLToPath } from 'node:url'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const delay = (ms) => new Promise((resolve) => globalThis.setTimeout(resolve, ms))

/* A candidate whose slide bitmap is deliberately huge: the click must stay responsive even when the
   preview payload is heavy (the real workspace ships every candidate's slide preview as a data URI). */
const HEAVY_PREVIEW = `data:image/jpeg;base64,${'A'.repeat(240 * 1024)}`

function imageCandidate(id, slide, { heavy = false } = {}) {
  return {
    id, sourceFile: 'deck.pptx', managementNumber: 'TNP-001', slide, pictureId: `p-${slide}`,
    src: 'data:image/png;base64,AA==', bounds: { x: 20, y: 25, w: 30, h: 35 },
    slideWidth: 9144000, slideHeight: 5143500, targetBboxPct: { x: 20, y: 25, w: 30, h: 35 },
    slidePreview: heavy ? HEAVY_PREVIEW : 'data:image/jpeg;base64,AA==',
    slidePreviewWidth: 1600, slidePreviewHeight: 900,
    slidePreviewBackend: 'builtin', slidePreviewFaithful: false,
    itemId: `item-${slide}`, itemIndex: 0, itemHeading: 'Mục cải tiến',
    decision: 'AFTER', confidenceBand: 'high', confidence: 0.94, evidence: ['evidence'],
    excelEligible: true, eligibilityReason: '', userLabel: 'UNLABELED', labelPending: false, note: '',
  }
}

function contentCandidate(id, slide) {
  return {
    id, sourceFile: 'deck.pptx', managementNumber: 'TNP-001', slide, blockId: `b-${slide}`,
    text: 'Nội dung cải tiến nguyên văn', decision: 'include', confidenceBand: 'high', confidence: 0.9,
    evidence: ['evidence'], section: 'improvement', nearestTitle: 'Cải tiến trong sản xuất',
    userLabel: 'UNLABELED', labelPending: false, note: '',
  }
}

function learningState(images, contents) {
  return {
    images, contents,
    candidateStatus: {
      image: { state: 'ready', reason: images.length ? 'available' : 'empty', count: images.length,
        message: `Đã tải ${images.length} ảnh ứng viên từ lượt xử lý hiện tại.` },
      content: { state: 'ready', reason: contents.length ? 'available' : 'empty', count: contents.length,
        message: `Đã tải ${contents.length} khối nội dung từ lượt xử lý hiện tại.` },
    },
    counts: { image: { total: images.length, labeled: 0 }, content: { total: contents.length, labeled: 0 } },
    modelStatus: { image: 'Chưa huấn luyện.', content: 'Chưa huấn luyện.' },
    excelPending: 0, excelLastResult: { kind: 'none', message: '' }, training: false,
  }
}

const dashboard = {
  app: { name: 'Report Extractor', title: 'Report Extractor', version: '1.3.6', build: '019', buildNumber: 19 },
  config: { paths: { reportFolder: '', template: '', output: '' },
    period: { mode: 'auto', month: '', year: '', from: '', to: '' }, forceReprocess: false,
    ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [] },
    update: { path: '', autoCheck: true } },
  scan: { scanned: false, message: '' }, reports: [],
  job: { status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0, currentFile: '',
    elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false },
  logs: [],
  ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [], aiStatus: '',
    discovering: false, discoveryMessage: '', discoveryResults: [], discoveryChecked: 0, discoveryTotal: 0,
    serverApplyRunning: false, serverApplyMessage: '' },
  update: { path: '', autoCheck: true, status: 'idle', message: '', progress: 0, progressStage: '' },
  diagnostics: { running: false, rows: [], error: '' },
}

function installGlobals(dom) {
  const names = ['window', 'document', 'navigator', 'HTMLElement', 'Node', 'Event', 'MouseEvent',
    'ResizeObserver', 'IS_REACT_ACT_ENVIRONMENT']
  const prior = new Map(names.map((name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)]))
  const define = (name, value) => Object.defineProperty(globalThis, name,
    { configurable: true, writable: true, value })
  define('window', dom.window)
  define('document', dom.window.document)
  define('navigator', dom.window.navigator)
  define('HTMLElement', dom.window.HTMLElement)
  define('Node', dom.window.Node)
  define('Event', dom.window.Event)
  define('MouseEvent', dom.window.MouseEvent)
  define('IS_REACT_ACT_ENVIRONMENT', true)
  class TestResizeObserver { observe() {} unobserve() {} disconnect() {} }
  define('ResizeObserver', TestResizeObserver)
  dom.window.ResizeObserver = TestResizeObserver
  dom.window.HTMLElement.prototype.scrollIntoView = () => {}
  dom.window.matchMedia = (media) => ({ matches: media === '(min-width: 1280px)', media, onchange: null,
    addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {},
    dispatchEvent() { return true } })
  return () => {
    dom.window.close()
    for (const [name, descriptor] of prior) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete globalThis[name]
    }
  }
}

/**
 * Boot the real StoreProvider + LearningTab against an instrumented pywebview mock.
 * `park` holds every get_learning_state reply until the returned release() is called, which is how the
 * stale-snapshot ordering (M5) is reproduced deterministically.
 */
async function bootLearningTab({ nImages = 6, nContents = 3, labelLatencyMs = 300, stateLatencyMs = 300,
  heavy = false, rejectLabels = false } = {}) {
  const dom = new JSDOM('<!doctype html><html><body><div id="root"></div></body></html>',
    { url: 'http://localhost/' })
  const restoreGlobals = installGlobals(dom)
  const server = await createServer({ configFile: `${frontendRoot}vite.config.ts`, root: frontendRoot,
    appType: 'custom', server: { middlewareMode: true, hmr: false } })

  const calls = { getLearning: 0, setLabel: 0, saveLabels: 0, setLabelArgs: [] }
  const persisted = { image: new Map(), content: new Map(), notes: new Map() }
  let parked = null
  const images = Array.from({ length: nImages }, (_, i) => imageCandidate(`img-${i}`, (i % 4) + 1, { heavy }))
  const contents = Array.from({ length: nContents }, (_, i) => contentCandidate(`txt-${i}`, (i % 4) + 1))

  const snapshot = () => learningState(
    images.map((c) => ({ ...c, userLabel: persisted.image.get(c.id) || 'UNLABELED',
      labelPending: persisted.image.has(c.id), note: persisted.notes.get(`image:${c.id}`) ?? c.note })),
    contents.map((c) => ({ ...c, userLabel: persisted.content.get(c.id) || 'UNLABELED',
      labelPending: persisted.content.has(c.id), note: persisted.notes.get(`content:${c.id}`) ?? c.note })),
  )

  const [{ StoreProvider, useStore }, { LearningTab }] = await Promise.all([
    server.ssrLoadModule('/src/state/store.tsx'), server.ssrLoadModule('/src/tabs/LearningTab.tsx')])
  const { createRoot } = await import('react-dom/client')

  function Harness() {
    const store = useStore()
    return React.createElement(React.Fragment, null,
      React.createElement('output', { 'data-testid': 'bridge-state' }, store.connected ? 'ready' : 'down'),
      React.createElement('output', { 'data-testid': 'error-code' }, store.errorCode),
      React.createElement('output', { 'data-testid': 'error-message' }, store.errorMessage),
      React.createElement('output', { 'data-testid': 'load-status' }, store.learningLoad.status),
      React.createElement('button', { type: 'button', 'data-testid': 'enter-learning',
        onClick: () => store.setTab('learning') }, 'enter'),
      React.createElement('button', { type: 'button', 'data-testid': 'refresh-learning',
        onClick: () => { void store.refreshLearning() } }, 'refresh'),
      React.createElement(LearningTab))
  }

  const root = createRoot(dom.window.document.getElementById('root'))
  await React.act(async () => {
    root.render(React.createElement(StoreProvider, null, React.createElement(Harness)))
  })
  const container = dom.window.document.getElementById('root')

  dom.window.pywebview = { api: {
    ping: async () => ({ ok: true, data: 'Python bridge OK' }),
    get_app_version: async () => ({ ok: true, data: dashboard.app }),
    get_current_config: async () => ({ ok: true, data: dashboard.config }),
    get_dashboard_state: async () => ({ ok: true, data: dashboard }),
    get_learning_state: async () => {
      calls.getLearning += 1
      const stale = snapshot()                 // taken now, exactly as Python would
      if (stateLatencyMs) await delay(stateLatencyMs)
      if (parked) await parked                 // held: pywebview serializes a huge reply late
      return { ok: true, data: stale }
    },
    set_learning_label: async (kind, id, label, note) => {
      calls.setLabel += 1
      calls.setLabelArgs.push({ kind, id, label, note })
      if (labelLatencyMs) await delay(labelLatencyMs)
      if (rejectLabels) {
        return { ok: false, error: { code: 'CANDIDATE_NOT_FOUND',
          message: 'Ảnh này không còn trong danh sách học hiện tại.' } }
      }
      persisted.notes.set(`${kind}:${id}`, note)
      ;(kind === 'image' ? persisted.image : persisted.content).set(id, label)
      return { ok: true, data: snapshot() }
    },
    save_learning_labels: async (kind, notes) => {
      calls.saveLabels += 1
      if (labelLatencyMs) await delay(labelLatencyMs)
      for (const [id, note] of Object.entries(notes || {})) persisted.notes.set(`${kind}:${id}`, note)
      return { ok: true, data: { message: 'Đã lưu 1 nhãn xác nhận.', state: snapshot() } }
    },
  } }

  await React.act(async () => { dom.window.dispatchEvent(new dom.window.Event('pywebviewready')) })
  await waitFor(() => container.querySelector('[data-testid="bridge-state"]')?.textContent === 'ready',
    'the pywebview startup handshake')
  await React.act(async () => {
    container.querySelector('[data-testid="enter-learning"]')
      .dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true }))
  })
  // Learning state is only loaded once its tab is entered; that read is the first authoritative snapshot.
  await waitFor(() => calls.getLearning > 0
    && container.querySelector('[data-testid="load-status"]')?.textContent === 'ready',
    'the Learning-entry refresh')

  const click = (element) => React.act(async () => {
    element.dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true }))
  })
  const api = {
    dom, container, calls, persisted, snapshot, click,
    openSection: async (label) => {
      const button = [...container.querySelectorAll('aside nav button')]
        .find((entry) => entry.textContent.trim() === label)
      assert.ok(button, `${label} section button should render`)
      await click(button)
    },
    labelButton: (text, scope = 'Thông tin và nhãn xác nhận') => {
      const panel = container.querySelector(`section[aria-label="${scope}"]`) || container
      return [...panel.querySelectorAll('button')].find((b) => b.textContent.trim() === text)
    },
    isActive: (text, scope) => api.labelButton(text, scope).className.includes('ring-brand-400'),
    candidateButton: (text) => [...container.querySelectorAll('section[aria-label="Đối tượng ảnh"] button')]
      .find((button) => button.textContent.includes(text)),
    testId: (id) => container.querySelector(`[data-testid="${id}"]`)?.textContent ?? '',
    /** Hold every learning-state reply so an in-flight snapshot can be released after a click. */
    parkPolls: () => { parked = new Promise((resolve) => { parkedRelease = resolve }) },
    releasePolls: () => { const resolve = parkedRelease; parked = null; parkedRelease = null; resolve?.() },
    settle: (ms) => React.act(async () => { await delay(ms) }),
    cleanup: async () => {
      api.releasePolls()
      await React.act(async () => { root.unmount() })
      await server.close()
      restoreGlobals()
    },
  }
  return api
}

let parkedRelease = null

async function waitFor(predicate, description, timeoutMs = 8_000) {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    if (predicate()) return
    await React.act(async () => { await delay(10) })
  }
  assert.fail(`Timed out waiting for ${description}`)
}

/* ------------------------------------------------------------------ M4: a click is visible immediately */
test('PROMPT-036: a Trước/Sau selection is on screen before the bridge answers', async () => {
  const h = await bootLearningTab({ labelLatencyMs: 500, stateLatencyMs: 60 })
  try {
    await h.openSection('Kiểm tra ảnh cải tiến')
    assert.ok(h.labelButton('Sau cải tiến'), 'the label buttons should render')
    assert.equal(h.isActive('Sau cải tiến'), false)

    const started = performance.now()
    await h.click(h.labelButton('Sau cải tiến'))
    const elapsed = performance.now() - started

    // The bridge has NOT answered yet (500 ms), so this can only be true with a local update.
    assert.equal(h.calls.setLabel, 1, 'exactly one set_learning_label per click (no duplicates)')
    assert.equal(h.isActive('Sau cải tiến'), true,
      'the selection must be visible inside the click, not after the round trip')
    assert.ok(elapsed < 250, `the click handler blocked for ${elapsed.toFixed(0)} ms`)
    assert.match(h.container.textContent, /chưa lưu/, 'a provisional label is honestly marked unsaved')

    await waitFor(() => h.calls.setLabel === 1 && h.persisted.image.get('img-0') === 'AFTER',
      'the backend to accept the label')
    await h.settle(700)
    assert.equal(h.isActive('Sau cải tiến'), true, 'the authoritative reply must keep the selection')
  } finally { await h.cleanup() }
})

test('PROMPT-036: rapid Trước <-> Sau switching loses no click and converges on the last one', async () => {
  const h = await bootLearningTab({ labelLatencyMs: 250, stateLatencyMs: 250 })
  try {
    await h.openSection('Kiểm tra ảnh cải tiến')
    const sequence = ['Sau cải tiến', 'Trước cải tiến', 'Sau cải tiến', 'Trước cải tiến', 'Sau cải tiến']
    for (const label of sequence) {
      await h.click(h.labelButton(label))
      assert.equal(h.isActive(label), true, `${label} must be visible the moment it is clicked`)
      await h.settle(30)                          // the reviewer clicks faster than the bridge answers
    }
    assert.equal(h.calls.setLabel, sequence.length, 'one bridge call per click, none dropped or duplicated')
    assert.deepEqual(h.calls.setLabelArgs.map((entry) => entry.label),
      ['AFTER', 'BEFORE', 'AFTER', 'BEFORE', 'AFTER'], 'every click reached the backend in order')

    await waitFor(() => h.persisted.image.get('img-0') === 'AFTER', 'the last click to be persisted')
    await h.settle(2_500)                         // let the polling settle on the authoritative state
    assert.equal(h.isActive('Sau cải tiến'), true, 'the UI must end on the last selection')
    assert.equal(h.isActive('Trước cải tiến'), false)
  } finally { await h.cleanup() }
})

/* ------------------------------------------------------------------ M5: a stale snapshot must not win */
test('PROMPT-036: a snapshot in flight before the click cannot revert the accepted selection', async () => {
  const h = await bootLearningTab({ labelLatencyMs: 20, stateLatencyMs: 0 })
  try {
    await h.openSection('Kiểm tra ảnh cải tiến')
    await h.settle(1_800)                         // make sure a poll is due

    h.parkPolls()                                 // a get_learning_state snapshot is now held mid-bridge
    await h.settle(1_800)
    const parkedCalls = h.calls.getLearning

    await h.click(h.labelButton('Trước cải tiến'))
    assert.equal(h.isActive('Trước cải tiến'), true)
    await waitFor(() => h.persisted.image.get('img-0') === 'BEFORE', 'the backend to accept BEFORE')
    await h.settle(150)
    assert.equal(h.isActive('Trước cải tiến'), true, 'the selection is visible while the stale reply is held')

    h.releasePolls()                              // the pre-click snapshot finally lands
    await h.settle(300)
    assert.equal(h.isActive('Trước cải tiến'), true,
      'a stale snapshot reverted an accepted selection (PROMPT-036 M5 regression)')
    assert.equal(h.persisted.image.get('img-0'), 'BEFORE')
    assert.ok(h.calls.getLearning >= parkedCalls)
  } finally { await h.cleanup() }
})

/* ------------------------------------------------------------------ feedback is never lost */
test('PROMPT-036: a rejected selection rolls back and surfaces the real backend error', async () => {
  const h = await bootLearningTab({ labelLatencyMs: 60, stateLatencyMs: 60, rejectLabels: true })
  try {
    await h.openSection('Kiểm tra ảnh cải tiến')
    await h.click(h.labelButton('Sau cải tiến'))
    assert.equal(h.isActive('Sau cải tiến'), true, 'optimistic feedback is shown first')
    await waitFor(() => h.testId('error-message').includes('không còn trong danh sách học hiện tại'),
      'the backend rejection to be surfaced')
    assert.equal(h.testId('error-code'), 'CANDIDATE_NOT_FOUND')
    await h.settle(200)
    assert.equal(h.isActive('Sau cải tiến'), false,
      'a rejected label must not stay on screen pretending to be accepted')
    assert.equal(h.persisted.image.size, 0, 'nothing was persisted')
  } finally { await h.cleanup() }
})

test('PROMPT-036: an unsaved note survives an authoritative snapshot', async () => {
  const h = await bootLearningTab({ labelLatencyMs: 40, stateLatencyMs: 40 })
  try {
    await h.openSection('Kiểm tra ảnh cải tiến')
    const input = h.container.querySelector('section[aria-label="Thông tin và nhãn xác nhận"] input')
    assert.ok(input, 'the note input should render')
    await React.act(async () => {
      const setter = Object.getOwnPropertyDescriptor(h.dom.window.HTMLInputElement.prototype, 'value').set
      setter.call(input, 'ghi chú đang gõ dở')
      input.dispatchEvent(new h.dom.window.Event('input', { bubbles: true }))
    })
    assert.equal(input.value, 'ghi chú đang gõ dở')

    // explicit refresh: a full authoritative snapshot that does not know the local note yet
    const before = h.calls.getLearning
    await h.click(h.container.querySelector('[data-testid="refresh-learning"]'))
    await waitFor(() => h.calls.getLearning > before, 'the manual refresh to reach the bridge')
    await h.settle(300)
    const after = h.container.querySelector('section[aria-label="Thông tin và nhãn xác nhận"] input')
    assert.equal(after.value, 'ghi chú đang gõ dở', 'a poll erased an unsaved note (PROMPT-036 input lag)')

    // and the note is what actually reaches Python when the label is chosen
    await h.click(h.labelButton('Sau cải tiến'))
    await waitFor(() => h.calls.setLabel === 1, 'the label call')
    assert.equal(h.calls.setLabelArgs[0].note, 'ghi chú đang gõ dở')
  } finally { await h.cleanup() }
})

/* ------------------------------------------------------------------ heavy preview must not block */
test('PROMPT-036: a candidate with a heavy slide preview is still click-responsive', async () => {
  const h = await bootLearningTab({ nImages: 8, heavy: true, labelLatencyMs: 400, stateLatencyMs: 400 })
  try {
    await h.openSection('Kiểm tra ảnh cải tiến')
    await waitFor(() => h.container.querySelector('[data-testid="slide-preview-bitmap"]') !== null,
      'the heavy slide preview to render')

    const started = performance.now()
    await h.click(h.labelButton('Trước cải tiến'))
    const elapsed = performance.now() - started
    assert.equal(h.isActive('Trước cải tiến'), true,
      'a heavy preview payload must not delay the visible selection')
    assert.ok(elapsed < 300, `the click handler blocked for ${elapsed.toFixed(0)} ms on a heavy preview`)

    // switching to another heavy candidate is a pure local index change: no bridge call at all
    const callsBefore = h.calls.getLearning + h.calls.setLabel
    await h.click(h.candidateButton('Slide 2'))
    await waitFor(() => h.container.querySelector('section[aria-label="Slide ngữ cảnh"] h2')
      ?.textContent === 'Slide 2', 'switching candidate')
    assert.equal(h.calls.getLearning + h.calls.setLabel, callsBefore,
      'selecting a candidate must not reload Learning state')
  } finally { await h.cleanup() }
})

/* ------------------------------------------------------------------ PROMPT-032 contracts still hold */
test('PROMPT-036: explicit refresh and the content-review path keep their PROMPT-032 contracts', async () => {
  const h = await bootLearningTab({ labelLatencyMs: 30, stateLatencyMs: 30 })
  try {
    // an explicit user refresh always performs a real backend read
    for (let i = 0; i < 2; i += 1) {
      const before = h.calls.getLearning
      await h.click(h.container.querySelector('[data-testid="refresh-learning"]'))
      await waitFor(() => h.calls.getLearning > before, 'a fresh manual Learning bridge call')
    }

    // content review gets the same immediate-feedback treatment, with its own label semantics untouched
    await h.openSection('Kiểm tra nội dung cải tiến')
    await waitFor(() => h.container.textContent.includes('Nội dung cải tiến nguyên văn'),
      'the verbatim content candidate')
    const contentButton = [...h.container.querySelectorAll('button')]
      .find((b) => b.textContent.trim() === 'Nội dung cải tiến')
    assert.ok(contentButton, 'the content label button should render')
    await h.click(contentButton)
    assert.equal(contentButton.className.includes('ring-brand-400'), true,
      'the content selection must be visible immediately')
    await waitFor(() => h.persisted.content.get('txt-0') === 'IMPROVEMENT_CONTENT', 'the content label call')
    assert.equal(h.calls.setLabelArgs.at(-1).kind, 'content')
    assert.equal(h.calls.setLabelArgs.at(-1).label, 'IMPROVEMENT_CONTENT')
    assert.equal(h.testId('load-status'), 'ready')
  } finally { await h.cleanup() }
})
